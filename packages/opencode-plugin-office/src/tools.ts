import { tool, type ToolResult } from "@opencode-ai/plugin"
import {
  OfficeError,
  createDocx,
  createPptx,
  editDocx,
  editPptx,
  ensureVenv,
  formatDocxRead,
  formatPptxRead,
  readDocx,
  readPptx,
  renderOffice,
  type DocxOperation,
  type PptxOperation,
} from "@opencode-office/core"
import { truncateForModel } from "./truncate"

// `tool()` (see node_modules/.../@opencode-ai/plugin/dist/tool.js) is a plain identity
// function — it does not parse/validate `args` against the declared zod shape before
// `execute` runs. So the shape below is metadata for the host/LLM only; every runtime
// check (defaults, refinements, error codes) has to happen inside `execute` itself.
const z = tool.schema

const DOCX_OP_NAMES = ["replace_text", "insert_content", "delete_element", "set_style", "set_table_cell"] as const
const PPTX_OP_NAMES = [
  "set_shape_text",
  "set_notes",
  "insert_slide",
  "duplicate_slide",
  "delete_slide",
  "move_slide",
  "replace_image",
] as const
const DOCX_OPS = new Set<string>(DOCX_OP_NAMES)
const PPTX_OPS = new Set<string>(PPTX_OP_NAMES)

// Ops whose core type requires `anchor` (not just allows it, e.g. set_table_cell's is optional).
// Checked so a MISSING anchor fails fast here too, not only a present-but-empty one.
const REQUIRED_ANCHOR_OPS = new Set<string>(["replace_text", "delete_element", "set_style", "set_shape_text"])

// office-core throws OfficeError with `message` and `hint` as separate fields. Some hosts only
// surface `error.message` to the model, which would silently drop the hint — and the whole
// error-recovery workflow in SKILL.md ("follow every error hint") depends on it being visible.
// Fold hint into message defensively so it survives even a message-only host, while keeping
// `.code`/`.hint` intact for hosts (and tests) that read the structured fields directly.
async function withHintedErrors<R>(run: () => Promise<R>): Promise<R> {
  try {
    return await run()
  } catch (err) {
    if (err instanceof OfficeError && !err.message.includes(err.hint)) {
      throw new OfficeError(err.code, `${err.message} — ${err.hint}`, err.hint)
    }
    throw err
  }
}

function requireOfficeExt(file: string): "docx" | "pptx" {
  const dot = file.lastIndexOf(".")
  const ext = dot >= 0 ? file.slice(dot + 1).toLowerCase() : ""
  if (ext === "docx" || ext === "pptx") return ext
  throw new OfficeError(
    "UNSUPPORTED_FORMAT",
    `Unsupported file extension: .${ext || "(none)"}`,
    "Only .docx and .pptx files are supported by these tools.",
  )
}

function validateOperations(ext: "docx" | "pptx", operations: unknown): Record<string, unknown>[] {
  const parsed = z.array(z.record(z.string(), z.unknown())).min(1).safeParse(operations)
  if (!parsed.success) {
    throw new OfficeError(
      "BAD_ARGS",
      `operations must be a non-empty array of operation objects: ${parsed.error.issues[0]?.message ?? "invalid"}`,
      "Pass an array with at least one operation object, each shaped like { op, target, anchor, ... }.",
    )
  }

  const validOps = ext === "docx" ? DOCX_OPS : PPTX_OPS
  for (const el of parsed.data) {
    const op = el.op
    if (typeof op !== "string" || !validOps.has(op)) {
      throw new OfficeError(
        "WRONG_OPS_FORMAT",
        `Operation "${String(op)}" is not valid for .${ext} files`,
        `Valid ops for .${ext}: ${[...validOps].join(", ")}`,
      )
    }

    if (REQUIRED_ANCHOR_OPS.has(op) && !("anchor" in el)) {
      throw new OfficeError("BAD_ANCHOR", `${op} requires an anchor`, "Copy the exact current text from office_read as the anchor.")
    }
    if ("anchor" in el) {
      const anchor = el.anchor
      if (typeof anchor !== "string" || anchor.length < 1) {
        throw new OfficeError("BAD_ANCHOR", "anchor must be non-empty", "Copy the exact current text from office_read as the anchor.")
      }
    }

    for (const key of ["row", "col", "index"] as const) {
      if (key in el) {
        const value = el[key]
        if (typeof value !== "number" || !Number.isInteger(value) || value < 0) {
          throw new OfficeError("BAD_ARGS", `${key} must be a non-negative integer`, `Provide ${key} as an integer >= 0.`)
        }
      }
    }
  }

  return parsed.data
}

const officeRead = tool({
  description:
    'Read a .docx or .pptx file. Start with mode "outline" (headings/titles only), then "content" for full text — optionally narrowed with a target element ID copied from a prior read.',
  args: {
    file: z.string().describe("Path to the .docx or .pptx file to read"),
    mode: z
      .enum(["outline", "content", "full"])
      .optional()
      .describe(
        'Read granularity. "outline": headings/titles only (default). "content": full text. "full": content plus comments/tracked changes (docx only — pptx coerces "full" to "content").',
      ),
    target: z.string().optional().describe("Element ID from a prior office_read call (e.g. p:3, tbl:1, s:2, s:2/sh:1) to narrow the read to one element"),
  },
  async execute(args): Promise<ToolResult> {
    return withHintedErrors(async () => {
      const mode = args.mode ?? "outline"
      const ext = requireOfficeExt(args.file)
      if (ext === "docx") {
        const result = await readDocx(args.file, mode, args.target)
        return truncateForModel(formatDocxRead(result))
      }
      const pptxMode = mode === "full" ? "content" : mode
      const result = await readPptx(args.file, pptxMode, args.target)
      return truncateForModel(formatPptxRead(result))
    })
  },
})

const officeEdit = tool({
  description:
    "Apply a batch of anchored edit operations to a .docx or .pptx file. Element IDs and anchor text must be copied exactly from a prior office_read call — IDs shift per-op inside a batch, so order structural ops last or re-read between batches.",
  args: {
    file: z.string().describe("Path to the .docx or .pptx file to edit"),
    operations: z
      .array(z.record(z.string(), z.unknown()))
      .min(1)
      .describe("Non-empty array of operation objects, each with an `op` name valid for the file's format plus that op's fields (target/anchor/text/etc.)"),
  },
  async execute(args, ctx): Promise<ToolResult> {
    return withHintedErrors(async () => {
      const ext = requireOfficeExt(args.file)
      const operations = validateOperations(ext, args.operations)

      await ctx.ask({
        permission: "office_edit",
        patterns: [args.file],
        always: [args.file],
        metadata: { operations: operations.length },
      })

      const result =
        ext === "docx"
          ? await editDocx(args.file, operations as unknown as DocxOperation[])
          : await editPptx(args.file, operations as unknown as PptxOperation[])

      const lines = [`Applied ${result.applied} operation(s). Backup: ${result.backup}`]
      result.results.forEach((r, i) => lines.push(`  [${i}] ${JSON.stringify(r)}`))
      return truncateForModel(lines.join("\n"))
    })
  },
})

const officeCreate = tool({
  description:
    "Create a new .docx (from markdown) or .pptx (from a slide spec array) file. Fails with FILE_EXISTS if the target already exists — use office_edit on an existing file instead.",
  args: {
    file: z.string().describe("Path for the new .docx or .pptx file"),
    markdown: z.string().optional().describe("Markdown body for a new .docx (required for .docx)"),
    slides: z
      .array(
        z.object({
          layout: z.string(),
          title: z.string().optional(),
          bullets: z.array(z.string()).optional(),
          notes: z.string().optional(),
        }),
      )
      .optional()
      .describe("Slide specs for a new .pptx (required for .pptx)"),
    reference: z.string().optional().describe("Optional reference .docx to base styles on"),
    template: z.string().optional().describe("Optional template .pptx to base the deck on"),
  },
  async execute(args, ctx): Promise<ToolResult> {
    return withHintedErrors(async () => {
      const ext = requireOfficeExt(args.file)

      if (ext === "docx") {
        if (!args.markdown) {
          throw new OfficeError("BAD_ARGS", "markdown is required to create a .docx file", "Pass markdown describing the document body in the `markdown` argument.")
        }
        await ctx.ask({ permission: "office_create", patterns: [args.file], always: [args.file], metadata: { format: "docx" } })
        const result = await createDocx(args.file, args.markdown, { reference: args.reference })
        return truncateForModel(`Created ${result.file} with ${result.paragraphs} paragraph(s).`)
      }

      if (!args.slides || args.slides.length === 0) {
        throw new OfficeError("BAD_ARGS", "slides is required (non-empty) to create a .pptx file", "Pass a non-empty array of slide specs ({ layout, title?, bullets?, notes? }) in the `slides` argument.")
      }
      await ctx.ask({ permission: "office_create", patterns: [args.file], always: [args.file], metadata: { format: "pptx" } })
      const result = await createPptx(args.file, args.slides, { template: args.template })
      const skipped = result.skipped.length ? ` Skipped: ${result.skipped.map((s) => `slide ${s.slide} field ${s.field}`).join(", ")}.` : ""
      return truncateForModel(`Created ${result.file} with ${result.slides} slide(s).${skipped}`)
    })
  },
})

const officeRender = tool({
  description: "Render a .docx or .pptx file to PNG images (one per page/slide) via LibreOffice, for visual verification after create/edit.",
  args: {
    file: z.string().describe("Path to the .docx or .pptx file to render"),
    pages: z.array(z.number().int().min(1)).optional().describe("1-indexed page/slide numbers to render; omit for all pages, [] for none"),
  },
  async execute(args): Promise<ToolResult> {
    return withHintedErrors(async () => {
      const result = await renderOffice(args.file, { pages: args.pages })
      return {
        output: `Rendered ${result.pages.length} page(s)`,
        attachments: result.pages.map((p) => ({
          type: "file" as const,
          mime: "image/png",
          url: `file://${p.path}`,
          filename: `page-${p.page}.png`,
        })),
      }
    })
  },
})

const PYTHON_TIMEOUT_MS = 120_000

const officePython = tool({
  description:
    "Escape hatch: run arbitrary Python with python-docx/python-pptx/pillow/pymupdf preloaded in the managed venv. Use only as a last resort, when the typed office_read/office_edit/office_create/office_render tools cannot express the needed change.",
  args: {
    code: z.string().describe("Python source to execute via `python -c`"),
    files: z.array(z.string()).optional().describe("File paths this script touches, for the permission prompt; omit/empty to request access to all files"),
  },
  async execute(args, ctx): Promise<ToolResult> {
    return withHintedErrors(async () => {
      const files = args.files ?? []
      await ctx.ask({
        permission: "office_python",
        patterns: files.length ? files : ["*"],
        always: [],
        metadata: {},
      })

      const python = await ensureVenv()
      const proc = Bun.spawn([python, "-c", args.code], {
        cwd: ctx.directory,
        stdout: "pipe",
        stderr: "pipe",
      })
      let timedOut = false
      const timer = setTimeout(() => {
        timedOut = true
        proc.kill()
        setTimeout(() => proc.kill(9), 5_000).unref?.()
      }, PYTHON_TIMEOUT_MS)
      const [stdout, stderr, code] = await Promise.all([new Response(proc.stdout).text(), new Response(proc.stderr).text(), proc.exited])
      clearTimeout(timer)
      if (timedOut) {
        throw new OfficeError("PYTHON_TIMEOUT", `office_python exceeded ${PYTHON_TIMEOUT_MS}ms`, "Simplify the script or operate on a smaller slice of the file, then retry.")
      }

      const lines = [`exit code: ${code}`, "", "stdout:", stdout || "(empty)"]
      if (stderr) lines.push("", "stderr:", stderr)
      return truncateForModel(lines.join("\n"))
    })
  },
})

export const officeTools = {
  office_read: officeRead,
  office_edit: officeEdit,
  office_create: officeCreate,
  office_render: officeRender,
  office_python: officePython,
}
