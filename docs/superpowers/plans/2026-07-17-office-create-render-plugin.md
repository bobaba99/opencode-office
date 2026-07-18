# Office Tools Plan 3: Create, Render, Plugin Binding Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Pay the two hard-deadline debts (venv lock, hyperlink text model), add comments reading, `createDocx`/`createPptx`, `renderOffice` (soffice→PDF→pymupdf→PNGs), and the installable OpenCode plugin (`opencode-plugin-office`) with five zod-validated tools plus a bundled SKILL.md.

**Architecture:** unchanged core patterns. New: pymupdf joins the pinned venv (rasterizes LibreOffice-produced PDFs — no poppler dependency); the plugin package is a thin binding over office-core using `@opencode-ai/plugin`'s `tool()` helper.

**Verified API facts (do not re-litigate):** python-docx 1.2.0 has `Document.add_comment(runs, text="", author="", initials=None) -> Comment`; `doc.comments` iterates Comment objects with `.comment_id/.author/.text`; `Paragraph.text` INCLUDES hyperlink text while `Paragraph.runs` EXCLUDES hyperlink runs; `Paragraph.iter_inner_content()` yields Run and Hyperlink objects (Hyperlink has `.runs`). `soffice` exists at /opt/homebrew/bin/soffice on the dev machine. python-pptx 1.0.2 `get_or_add_image_part(image_file) -> (part, rId)`.

**Spec:** `docs/superpowers/specs/2026-07-17-office-tools-design.md`. Plan 3 of 4.

## Global Constraints

- Everything from Plans 1–2 Global Constraints binds (OfficeError recovery-first hints, ID grammar, atomic saves, anchors, no OpenCode imports in office-core — the PLUGIN package may import `@opencode-ai/plugin`).
- New pins: add `pymupdf==1.26.3` to `PINNED` (fingerprint bump auto-reprovisions).
- New error codes: `RENDER_UNAVAILABLE` (soffice missing; hint: `brew install --cask libreoffice` / apt equivalent), `RENDER_FAILED`, `UNSUPPORTED_FORMAT` (file is neither .docx nor .pptx), `WRONG_OPS_FORMAT` (docx ops sent to a pptx file or vice versa).
- Tool output cap: 24_000 chars via `truncateForModel` with a notice naming `target`/outline as the recovery.
- Plugin tool args validated with zod: anchors `.min(1)`, row/col/index `.int()`, operations arrays `.min(1)`.
- If the published `@opencode-ai/plugin` typings differ from what this plan assumes (`tool()` helper, `ToolContext.ask`, `ToolResult.attachments`), adapt minimally to the installed typings and record every adaptation in the task report.

---

### Task 1: Debt — venv lock, worker robustness, timeout exposure

**Files:**
- Modify: `packages/office-core/src/runtime.ts`
- Modify: `packages/office-core/src/worker.ts`
- Modify: `packages/office-core/src/docx/read.ts`, `src/docx/edit.ts`, `src/pptx/read.ts`, `src/pptx/edit.ts` (opts gain `timeoutMs?: number`, passed through to `runWorker`)
- Modify: `packages/office-core/src/python/docx_edit.py`, `pptx_edit.py` (save via try/finally tmp cleanup)
- Test: `packages/office-core/test/runtime.test.ts` (extend)

**Interfaces:**
- `acquireLock(dir: string, opts?: { staleMs?: number; timeoutMs?: number }): Promise<() => Promise<void>>` exported from runtime.ts — atomic `mkdir(<dir>.lock)`; on EEXIST, poll every 250ms; a lock older than `staleMs` (default 300_000, by the lock dir's mtime) is removed and retaken; give up after `timeoutMs` (default 120_000) with `OfficeError("LOCK_TIMEOUT", ..., hint naming the lock path and how to remove it)`. `ensureVenv` wraps provisioning (not the fast path) in this lock and re-checks `venvIsCurrent` after acquiring.
- `runWorker`: wrap `proc.stdin.write/end` in try/catch (a dead worker → fall through to the exit-code path, no raw EPIPE); timeout kill escalates: `proc.kill()` then after 5s `proc.kill(9)`.
- Every read/edit function signature gains `timeoutMs?: number` in its opts object, forwarded to `runWorker`.

- [ ] **Step 1: Write the failing tests** — append to `packages/office-core/test/runtime.test.ts`:

```ts
import { rm, mkdir, utimes } from "node:fs/promises"

test("acquireLock excludes a second contender until released", async () => {
  const dir = "/tmp/oc-office-lock-test"
  await rm(dir + ".lock", { recursive: true, force: true })
  const release = await acquireLock(dir)
  let second = false
  const contender = acquireLock(dir, { timeoutMs: 5_000 }).then(async (rel) => {
    second = true
    await rel()
  })
  await new Promise((resolve) => setTimeout(resolve, 400))
  expect(second).toBe(false)
  await release()
  await contender
  expect(second).toBe(true)
})

test("stale lock is stolen", async () => {
  const dir = "/tmp/oc-office-stale-test"
  await rm(dir + ".lock", { recursive: true, force: true })
  await mkdir(dir + ".lock", { recursive: true })
  const old = new Date(Date.now() - 600_000)
  await utimes(dir + ".lock", old, old)
  const release = await acquireLock(dir, { staleMs: 300_000, timeoutMs: 5_000 })
  await release()
})
```

Also add `acquireLock` to the runtime import line.

- [ ] **Step 2: `bun test runtime` — FAIL (acquireLock not exported)**

- [ ] **Step 3: Implement**

Append to `packages/office-core/src/runtime.ts`:

```ts
export async function acquireLock(
  dir: string,
  opts?: { staleMs?: number; timeoutMs?: number },
): Promise<() => Promise<void>> {
  const lockDir = dir + ".lock"
  const staleMs = opts?.staleMs ?? 300_000
  const deadline = Date.now() + (opts?.timeoutMs ?? 120_000)
  for (;;) {
    try {
      await mkdir(lockDir)
      return async () => {
        await rm(lockDir, { recursive: true, force: true })
      }
    } catch {
      try {
        const age = Date.now() - (await stat(lockDir)).mtimeMs
        if (age > staleMs) {
          await rm(lockDir, { recursive: true, force: true })
          continue
        }
      } catch {
        continue
      }
      if (Date.now() > deadline)
        throw new OfficeError(
          "LOCK_TIMEOUT",
          `Another process holds the Office provisioning lock at ${lockDir}`,
          `Wait for the other provisioning to finish, or remove the stale directory: rm -rf ${lockDir}`,
        )
      await new Promise((resolve) => setTimeout(resolve, 250))
    }
  }
}
```

(add `rm`, `stat` to the fs/promises import). In `ensureVenv`, after the fast-path `venvIsCurrent` check fails:

```ts
  const release = await acquireLock(path.join(cacheDir, "venv"))
  try {
    if (await venvIsCurrent(venvDir)) return python
    // ...existing provisioning body unchanged...
    return python
  } finally {
    await release()
  }
```

In `worker.ts`, replace the stdin write/end lines with:

```ts
  try {
    proc.stdin.write(JSON.stringify(payload))
    await proc.stdin.end()
  } catch {
    // worker died before consuming stdin; the exit-code path below reports it
  }
```

and the timeout callback becomes:

```ts
  const timer = setTimeout(() => {
    timedOut = true
    proc.kill()
    setTimeout(() => proc.kill(9), 5_000).unref?.()
  }, timeoutMs)
```

In the four read/edit TS files: add `timeoutMs?: number` to the opts type and pass `{ timeoutMs: opts?.timeoutMs, cacheDir: opts?.cacheDir }` to `runWorker`.

In both edit workers, replace the bare save with:

```python
    tmp = path + ".tmp-opencode-office"
    try:
        doc.save(tmp)   # prs.save(tmp) in pptx_edit.py
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
```

- [ ] **Step 4: `bun test runtime` (5 tests), full `bun test` (52), `bun run typecheck`**

- [ ] **Step 5: Commit** — `fix: venv provisioning lock, worker stdin/kill robustness, timeoutMs exposure, tmp cleanup`

---

### Task 2: Debt — hyperlink-aware docx text model + formatting-preservation probe + style-error polish

**Files:**
- Modify: `packages/office-core/src/python/_docx_common.py` (add `flat_runs`, `docx_para_text`)
- Modify: `packages/office-core/src/python/docx_edit.py` (use them; STYLE_NOT_FOUND enumerates styles and also catches wrong-type ValueError)
- Modify: `packages/office-core/src/python/gen_fixtures.py` (hyperlink paragraph in edit-report.docx)
- Create: `packages/office-core/src/python/docx_probe.py` (test-support: runs with formatting for a target paragraph)
- Test: `packages/office-core/test/docx-edit.test.ts` (extend)

**Interfaces:**
- `_docx_common.flat_runs(p)` — ordered run list including runs inside hyperlinks, via `p.iter_inner_content()` (Run appended directly; Hyperlink contributes its `.runs`).
- `_docx_common.docx_para_text(p)` — join of `flat_runs` texts (equals `p.text`).
- `docx_edit.py`: everywhere it previously used `para_text(el)` for a docx paragraph, use `docx_para_text(el)`; everywhere it calls `replace_in_paragraph(el, ...)`, call `replace_in_paragraph(RunSeq(flat_runs(el)), ...)` where `RunSeq` is a 3-line class with a `.runs` attribute (add it to `_textops.py`). `STYLE_NOT_FOUND` hint becomes `f"Available paragraph styles: {', '.join(available[:25])}"` where `available = [s.name for s in doc.styles if s.type == WD_STYLE_TYPE.PARAGRAPH]` (import `from docx.enum.style import WD_STYLE_TYPE`), and the `el.style = ...` assignment also catches `ValueError` (existing style of the wrong type) with the same code and a hint saying the style exists but is not a paragraph style.
- `docx_probe.py` — payload `{file, target}` → `{runs: [{text, bold}]}` for the target paragraph using `flat_runs` (bold via `run.bold is True`).
- Fixture: `make_edit_docx` gains, after the tracked paragraph, a hyperlink paragraph built exactly like this (recipe verified against python-docx 1.2.0):

```python
    link_p = doc.add_paragraph("See ")
    rid = link_p.part.relate_to(
        "https://example.com",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = link_p._p.makeelement(qn("w:hyperlink"), {qn("r:id"): rid})
    link_run = link_p._p.makeelement(qn("w:r"), {})
    link_text = link_p._p.makeelement(qn("w:t"), {})
    link_text.text = "the appendix"
    link_run.append(link_text)
    hyperlink.append(link_run)
    link_p._p.append(hyperlink)
    link_p.add_run(" for details.")
```

- [ ] **Step 1: Failing tests** — append to `docx-edit.test.ts` (remember `rm -rf packages/office-core/test/.fixtures` before running so the new fixture generates):

```ts
test("anchors spanning hyperlink text now match", async () => {
  const before = await readDocx(WORK, "content")
  const link = before.elements.find((e) => e.type === "paragraph" && e.text.includes("appendix"))!
  await editDocx(WORK, [{ op: "replace_text", target: link.id, anchor: "the appendix", text: "appendix B" }])
  const after = await readDocx(WORK, "content", link.id)
  expect((after.elements[0] as { text: string }).text).toBe("See appendix B for details.")
})

test("replace inside a bold run preserves bold", async () => {
  const before = await readDocx(WORK, "content")
  const target = before.elements.find((e) => e.type === "paragraph" && e.text.includes("strong"))!
  await editDocx(WORK, [{ op: "replace_text", target: target.id, anchor: "strong", text: "robust" }])
  const probe = await runWorker<{ runs: Array<{ text: string; bold: boolean }> }>("docx_probe.py", {
    file: WORK,
    target: target.id,
  })
  const robust = probe.runs.find((r) => r.text.includes("robust"))!
  expect(robust.bold).toBe(true)
})

test("STYLE_NOT_FOUND lists available styles", async () => {
  try {
    await editDocx(WORK, [{ op: "set_style", target: "p:0", anchor: "Edit Playground", style: "No Such" }])
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("STYLE_NOT_FOUND")
    expect((e as OfficeError).hint).toContain("Heading 1")
  }
})
```

(add `import { runWorker } from "../src/worker"` to the test file's imports)

- [ ] **Step 2: `rm -rf packages/office-core/test/.fixtures && bun test docx-edit` — new tests FAIL**
- [ ] **Step 3: Implement per Interfaces above** (`RunSeq` in `_textops.py`: `class RunSeq:` / `def __init__(self, runs):` / `self.runs = runs`)
- [ ] **Step 4: `bun test docx-edit` (15), full `bun test`, `bun run typecheck`**
- [ ] **Step 5: Commit** — `fix: hyperlink-aware anchors, bold-survival proof, style-error enumeration`

---

### Task 3: Comments reading (full mode)

**Files:**
- Modify: `packages/office-core/src/python/gen_fixtures.py` (add a comment in make_edit_docx: `doc.add_comment(runs=p.runs[:1], text="Verify this figure", author="Reviewer", initials="RV")` where `p` is the "Growth was strong" paragraph — add it right after that paragraph is built)
- Modify: `packages/office-core/src/python/docx_read.py` (full mode: result gains doc-level `comments: [{id, author, text}]` from `list(doc.comments)` — `{"id": c.comment_id, "author": c.author, "text": c.text}`; only in full mode, omitted when empty)
- Modify: `packages/office-core/src/docx/read.ts` (`DocxRead` gains `comments?: Array<{ id: number; author: string; text: string }>`; `formatDocxRead` appends a `comments:` section, one `  [id] author: text` line each, when present)
- Test: `packages/office-core/test/docx-read.test.ts` (extend)

- [ ] **Step 1: Failing test**:

```ts
test("full mode lists document comments; content mode does not", async () => {
  const EDIT = path.join(FIXTURE_DIR, "edit-report.docx")
  const full = await readDocx(EDIT, "full")
  expect(full.comments).toEqual([{ id: 0, author: "Reviewer", text: "Verify this figure" }])
  expect(formatDocxRead(full)).toContain("Reviewer: Verify this figure")
  const content = await readDocx(EDIT, "content")
  expect(content.comments).toBeUndefined()
})
```

(if `comment_id` turns out not to be 0 for the first comment, assert the actual observed id and note it in the report — author/text assertions must stand as written)

- [ ] **Step 2: `rm -rf packages/office-core/test/.fixtures && bun test docx-read` — FAIL**
- [ ] **Step 3: Implement per Files above**
- [ ] **Step 4: `bun test docx-read` (10), full suite, typecheck**
- [ ] **Step 5: Commit** — `feat: full-mode docx read lists comments`

---

### Task 4: createDocx / createPptx

**Files:**
- Create: `packages/office-core/src/python/_pptx_common.py` (move `find_layout`, `move_entry`, `delete_entry` out of pptx_edit.py; pptx_edit.py imports them)
- Create: `packages/office-core/src/python/docx_create.py`, `packages/office-core/src/python/pptx_create.py`
- Create: `packages/office-core/src/docx/create.ts`, `packages/office-core/src/pptx/create.ts`
- Modify: `packages/office-core/src/index.ts` (export both)
- Test: `packages/office-core/test/create.test.ts`

**Interfaces:**
- `createDocx(file: string, markdown: string, opts?: { reference?: string; cacheDir?: string; timeoutMs?: number }): Promise<{ file: string; paragraphs: number }>` — worker `docx_create.py`: `reference` present → `Document(reference)` then clear body (`for child in list(doc.element.body): if child.tag != qn("w:sectPr"): doc.element.body.remove(child)`) so its styles/theme survive; else `Document()`. Markdown mapping identical to insert_content (`# ## ### -` + plain; blank lines skipped; same STYLE_NOT_FOUND path). Saves atomically (tmp + replace + finally-cleanup). Refuses to overwrite an existing file → `OfficeError("FILE_EXISTS", ..., "office_create makes new files; use office_edit to change existing ones, or pass a different path.")` — checked worker-side with `os.path.exists` before any work.
- `createPptx(file: string, slides: Array<{ layout: string; title?: string; bullets?: string[]; notes?: string }>, opts?: { template?: string; cacheDir?: string; timeoutMs?: number }): Promise<{ file: string; slides: number; skipped: Array<{ slide: number; field: string }> }>` — worker `pptx_create.py`: `template` → `Presentation(template)` then delete ALL existing slides via `delete_entry` loop (theme/layouts survive); else `Presentation()`. Per spec slide: `find_layout` (LAYOUT_NOT_FOUND lists names), `add_slide`, set title if given (no title placeholder → record `{"slide": i, "field": "title"}` in `skipped`), bullets → idx-1 placeholder (absent → skipped entry), notes → notes_slide. Same FILE_EXISTS guard + atomic save.
- TS wrappers validate nothing beyond passing through (zod lives at the plugin boundary); they do NOT take a backup (new files).

- [ ] **Step 1: Failing tests** — `packages/office-core/test/create.test.ts`:

```ts
import { beforeAll, expect, test } from "bun:test"
import { rm } from "node:fs/promises"
import path from "node:path"
import { OfficeError } from "../src/errors"
import { createDocx } from "../src/docx/create"
import { createPptx } from "../src/pptx/create"
import { readDocx } from "../src/docx/read"
import { readPptx } from "../src/pptx/read"
import { FIXTURE_DIR, ensureFixtures } from "./fixtures"

const OUT_DOCX = path.join(FIXTURE_DIR, "created.docx")
const OUT_PPTX = path.join(FIXTURE_DIR, "created.pptx")

beforeAll(async () => {
  await ensureFixtures()
}, 180_000)

test("createDocx from markdown with reference styles", async () => {
  await rm(OUT_DOCX, { force: true })
  const result = await createDocx(OUT_DOCX, "# Report\nIntro paragraph.\n- first\n- second", {
    reference: path.join(FIXTURE_DIR, "report.docx"),
  })
  expect(result.paragraphs).toBe(4)
  const read = await readDocx(OUT_DOCX, "content")
  expect(read.elements.map((e) => (e.type === "paragraph" ? e.text : "table"))).toEqual([
    "Report",
    "Intro paragraph.",
    "first",
    "second",
  ])
  const heading = read.elements[0]
  expect(heading.type === "paragraph" && heading.style).toBe("Heading 1")
})

test("createPptx on a template inherits its layouts", async () => {
  await rm(OUT_PPTX, { force: true })
  const result = await createPptx(
    OUT_PPTX,
    [
      { layout: "Title Slide", title: "Q4 Plan" },
      { layout: "Title and Content", title: "Goals", bullets: ["Ship", "Measure"], notes: "keep it short" },
    ],
    { template: path.join(FIXTURE_DIR, "edit-deck.pptx") },
  )
  expect(result.slides).toBe(2)
  const read = await readPptx(OUT_PPTX, "content")
  expect(read.slides.map((s) => s.title)).toEqual(["Q4 Plan", "Goals"])
  expect(read.slides[1].notes).toBe("keep it short")
})

test("creating over an existing file is refused", async () => {
  try {
    await createDocx(OUT_DOCX, "# Again")
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("FILE_EXISTS")
  }
})

test("unknown layout in a slide spec lists layouts", async () => {
  await rm(OUT_PPTX, { force: true })
  try {
    await createPptx(OUT_PPTX, [{ layout: "Nope" }], { template: path.join(FIXTURE_DIR, "edit-deck.pptx") })
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("LAYOUT_NOT_FOUND")
  }
})
```

- [ ] **Step 2: `bun test create` — FAIL**
- [ ] **Step 3: Implement per Interfaces** (workers follow the exact patterns of docx_edit/pptx_edit: `_worker.run(main)`, FILE_OPEN on unreadable reference/template, atomic save with finally-cleanup; `docx_create.py` reuses `MD_STYLES`, `styled_line`, `add_styled_paragraph` — move those three from docx_edit.py into `_docx_common.py` and import from both workers rather than duplicating)
- [ ] **Step 4: `bun test create` (4), full suite, typecheck**
- [ ] **Step 5: Commit** — `feat: createDocx/createPptx — markdown and slide-spec creation on reference themes`

---

### Task 5: renderOffice

**Files:**
- Modify: `packages/office-core/src/runtime.ts` (PINNED gains `pymupdf: "1.26.3"`; add `findSoffice(): Promise<string | null>` — probe `soffice` on PATH via `which`, then `/opt/homebrew/bin/soffice`, `/Applications/LibreOffice.app/Contents/MacOS/soffice`, `/usr/bin/soffice`)
- Create: `packages/office-core/src/python/render_pdf.py` (payload `{pdf, outDir, pages?}` → pymupdf opens pdf, for each requested 1-based page (default all) writes `page-<n>.png` at 144 dpi → returns `{pages: [{page, path, width, height}]}`; bad page number → `WorkerError("RENDER_FAILED", ..., hint with page count)`)
- Create: `packages/office-core/src/render.ts`
- Modify: `packages/office-core/src/index.ts`
- Test: `packages/office-core/test/render.test.ts`

**Interfaces:**
- `renderOffice(file: string, opts?: { pages?: number[]; outDir?: string; cacheDir?: string; timeoutMs?: number }): Promise<{ pages: Array<{ page: number; path: string; width: number; height: number }> }>`:
  1. `findSoffice()` → null → `OfficeError("RENDER_UNAVAILABLE", "LibreOffice (soffice) is required for rendering", "macOS: brew install --cask libreoffice. Linux: apt install libreoffice. Reads and edits work without it.")`
  2. Spawn `[soffice, "--headless", `-env:UserInstallation=file://${profileDir}`, "--convert-to", "pdf", "--outdir", tmpDir, file]` with a fresh profileDir under cacheDir (`render-profile`) and a 120s default timeout; nonzero exit or missing output pdf → `OfficeError("RENDER_FAILED", ..., stderr-tail hint with retry guidance)`.
  3. `runWorker("render_pdf.py", { pdf, outDir, pages })` where outDir defaults to `<cacheDir>/renders/<basename-sans-ext>`.

- [ ] **Step 1: Failing tests** — `packages/office-core/test/render.test.ts`:

```ts
import { beforeAll, expect, test } from "bun:test"
import { existsSync } from "node:fs"
import path from "node:path"
import { findSoffice } from "../src/runtime"
import { renderOffice } from "../src/render"
import { FIXTURE_DIR, ensureFixtures } from "./fixtures"

const HAS_SOFFICE = (await findSoffice()) !== null

beforeAll(async () => {
  await ensureFixtures()
}, 180_000)

test.skipIf(!HAS_SOFFICE)("renders every slide of a deck to PNGs", async () => {
  const result = await renderOffice(path.join(FIXTURE_DIR, "deck.pptx"))
  expect(result.pages.map((p) => p.page)).toEqual([1, 2])
  for (const page of result.pages) {
    expect(existsSync(page.path)).toBe(true)
    expect(page.width).toBeGreaterThan(100)
  }
}, 300_000)

test.skipIf(!HAS_SOFFICE)("renders a single requested page of a document", async () => {
  const result = await renderOffice(path.join(FIXTURE_DIR, "report.docx"), { pages: [1] })
  expect(result.pages).toHaveLength(1)
  expect(result.pages[0].path.endsWith("page-1.png")).toBe(true)
}, 300_000)
```

- [ ] **Step 2: `bun test render` — FAIL**
- [ ] **Step 3: Implement per Interfaces** (note: bumping PINNED re-provisions the venv on the next worker call — expected, ~15s once)
- [ ] **Step 4: `bun test render` (2 on this machine), full suite, typecheck**
- [ ] **Step 5: Commit** — `feat: renderOffice — soffice to PDF to per-page PNGs via pymupdf`

---

### Task 6: Plugin package + SKILL.md + output truncation

**Files:**
- Create: `packages/opencode-plugin-office/package.json` (`name: "opencode-plugin-office"`, `type: module`, exports `./src/index.ts`, dependencies: `@opencode-ai/plugin` (latest), `@opencode-office/core: "workspace:*"`)
- Create: `packages/opencode-plugin-office/src/index.ts` (plugin entry)
- Create: `packages/opencode-plugin-office/src/tools.ts` (the five tool definitions)
- Create: `packages/opencode-plugin-office/src/truncate.ts` (`truncateForModel(text: string, limit = 24_000): string` — under limit → unchanged; over → first `limit` chars + `\n[truncated: showing ${limit} of ${total} chars — use mode:"outline" or a target ID to narrow]`)
- Create: `packages/opencode-plugin-office/skill/SKILL.md`
- Test: `packages/opencode-plugin-office/test/tools.test.ts`

**Interfaces:**
- `src/index.ts`:

```ts
import type { Plugin } from "@opencode-ai/plugin"
import { officeTools } from "./tools"

export const OfficePlugin: Plugin = async () => ({ tool: officeTools })
export default OfficePlugin
```

- `src/tools.ts` — five `tool()` definitions (`import { tool } from "@opencode-ai/plugin"`, zod via `tool.schema`), each `execute` returning a string (or `{ output, attachments }` for render). Behavior:
  - `office_read({ file, mode = "outline", target? })` — `.docx` → `readDocx` (mode as given), `.pptx` → `readPptx` (mode `"full"` coerced to `"content"`), anything else → throw `UNSUPPORTED_FORMAT` OfficeError (message names the extension; hint: only .docx/.pptx). Output = `truncateForModel(format*(result))`.
  - `office_edit({ file, operations })` — operations schema: `z.array(z.record(z.string(), z.unknown())).min(1)` refined per element: `op` must be a known docx or pptx op name; docx-op names sent to a .pptx file (or vice versa) → `WRONG_OPS_FORMAT` OfficeError listing the valid op names for that extension; every `anchor` field present must be a non-empty string (zod `.min(1)` semantics enforced in the refine — empty anchor → throw `OfficeError("BAD_ANCHOR", "anchor must be non-empty", "Copy the exact current text from office_read as the anchor.")`); `row`/`col`/`index` must be integers ≥ 0. Calls `await ctx.ask({ permission: "office_edit", patterns: [file], always: [file], metadata: { operations: operations.length } })` before editing; then `editDocx`/`editPptx` with the ops cast to their typed unions. Output: `Applied ${applied} operation(s). Backup: ${backup}` plus per-op result lines.
  - `office_create({ file, markdown?, slides?, reference?, template? })` — .docx requires `markdown` (else OfficeError `BAD_ARGS` with hint), .pptx requires `slides`; same ask() gate; maps to `createDocx`/`createPptx`.
  - `office_render({ file, pages? })` — `renderOffice`; result `{ output: \`Rendered ${n} page(s)\`, attachments: pages.map((p) => ({ type: "file", mime: "image/png", url: \`file://${p.path}\`, filename: \`page-${p.page}.png\` })) }`.
  - `office_python({ code, files = [] })` — `await ctx.ask({ permission: "office_python", patterns: files.length ? files : ["*"], always: [], metadata: {} })`; `ensureVenv()` then spawn `[python, "-c", code]` with cwd `ctx.directory`, 120s timeout, capture stdout+stderr; nonzero exit → output still returned with exit code noted (this tool surfaces raw results; it is the escape hatch). Output truncated via `truncateForModel`.
- `skill/SKILL.md` — frontmatter `name: office-tools`, description triggering on docx/pptx read/edit/create/render requests; body teaches: outline first → content with target → edit with IDs+anchors copied exactly from read output (IDs shift per-op inside a batch — order structural ops last or re-read); anchors must be unique in the element (extend if AMBIGUOUS_ANCHOR); every error hint states the recovery — follow it; create for new files (FILE_EXISTS means edit instead); render to visually verify after create/edit when LibreOffice is installed; office_python (python-docx/python-pptx/pillow/pymupdf preloaded) only when the typed ops cannot express the change.
- Tests (`test/tools.test.ts`) use a fake context `const ctx = { sessionID: "t", messageID: "t", agent: "t", directory: process.cwd(), worktree: process.cwd(), abort: new AbortController().signal, metadata: () => {}, ask: async () => {} }` (cast as needed) and the office-core fixtures via a relative import of `../../office-core/test/fixtures`:
  - office_read outline on deck.pptx contains `[s:0]`
  - office_read on a `.txt` path throws UNSUPPORTED_FORMAT
  - office_edit with an empty anchor throws BAD_ANCHOR without calling ask (assert via an ask spy that throws if called before validation — simplest: ask spy increments a counter; expect counter 0 after the failed call)
  - office_edit with pptx ops on a .docx throws WRONG_OPS_FORMAT
  - office_edit happy path on a work-copy docx applies and reports backup
  - office_python runs `print("hi")` and returns `hi`
  - `truncateForModel` unit: 30k-char input → contains `[truncated:` and length < 25_000
  - render test `skipIf(!HAS_SOFFICE)`: attachments array has 2 image/png entries for deck.pptx

- [ ] **Step 1: Write the failing tests; `bun install` after adding the package so the workspace links**
- [ ] **Step 2: `bun test packages/opencode-plugin-office` — FAIL**
- [ ] **Step 3: Implement.** If the installed `@opencode-ai/plugin` typings differ from the assumed `tool()`/`ToolContext`/`ToolResult` shapes, adapt minimally and record every adaptation in the report.
- [ ] **Step 4: package tests pass; full `bun test`; `bun run typecheck`**
- [ ] **Step 5: Commit** — `feat: opencode-plugin-office — five office tools, skill, output truncation`
