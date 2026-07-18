import { beforeAll, expect, test } from "bun:test"
import { copyFile } from "node:fs/promises"
import path from "node:path"
import type { ToolContext } from "@opencode-ai/plugin"
import { OfficeError, findSoffice, readDocx } from "@opencode-office/core"
import { FIXTURE_DIR, ensureFixtures } from "../../office-core/test/fixtures"
import { officeTools } from "../src/tools"
import { truncateForModel } from "../src/truncate"

const HAS_SOFFICE = (await findSoffice()) !== null

function makeCtx(overrides: Partial<ToolContext> = {}): ToolContext {
  return {
    sessionID: "t",
    messageID: "t",
    agent: "t",
    directory: process.cwd(),
    worktree: process.cwd(),
    abort: new AbortController().signal,
    metadata: () => {},
    ask: async () => {},
    ...overrides,
  } as ToolContext
}

const ctx = makeCtx()

beforeAll(async () => {
  await ensureFixtures()
}, 180_000)

test("office_read outline on deck.pptx contains slide ids", async () => {
  const output = await officeTools.office_read.execute({ file: path.join(FIXTURE_DIR, "deck.pptx") } as never, ctx)
  expect(String(output)).toContain("[s:0]")
})

test("office_read on an unsupported extension throws UNSUPPORTED_FORMAT", async () => {
  try {
    await officeTools.office_read.execute({ file: "notes.txt" } as never, ctx)
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("UNSUPPORTED_FORMAT")
  }
})

test("office_edit empty anchor throws BAD_ANCHOR without calling ask", async () => {
  let askCalls = 0
  const spyCtx = makeCtx({
    ask: async () => {
      askCalls++
    },
  })
  try {
    await officeTools.office_edit.execute(
      {
        file: path.join(FIXTURE_DIR, "report.docx"),
        operations: [{ op: "replace_text", target: "p:0", anchor: "", text: "x" }],
      } as never,
      spyCtx,
    )
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("BAD_ANCHOR")
  }
  expect(askCalls).toBe(0)
})

test("office_edit rejects pptx ops on a docx file with WRONG_OPS_FORMAT without calling ask", async () => {
  let askCalls = 0
  const spyCtx = makeCtx({
    ask: async () => {
      askCalls++
    },
  })
  try {
    await officeTools.office_edit.execute(
      {
        file: path.join(FIXTURE_DIR, "report.docx"),
        operations: [{ op: "set_shape_text", target: "s:0/sh:0", anchor: "x", text: "y" }],
      } as never,
      spyCtx,
    )
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("WRONG_OPS_FORMAT")
  }
  expect(askCalls).toBe(0)
})

test("office_edit missing anchor on an anchor-required op throws BAD_ANCHOR without calling ask", async () => {
  let askCalls = 0
  const spyCtx = makeCtx({
    ask: async () => {
      askCalls++
    },
  })
  try {
    await officeTools.office_edit.execute(
      {
        file: path.join(FIXTURE_DIR, "report.docx"),
        operations: [{ op: "replace_text", target: "p:0", text: "x" }],
      } as never,
      spyCtx,
    )
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("BAD_ANCHOR")
  }
  expect(askCalls).toBe(0)
})

test("missing target id fails BAD_ARGS before ask", async () => {
  let askCalls = 0
  const spyCtx = makeCtx({
    ask: async () => {
      askCalls++
    },
  })
  try {
    await officeTools.office_edit.execute(
      {
        file: path.join(FIXTURE_DIR, "report.docx"),
        operations: [{ op: "replace_text", anchor: "x", text: "y" }],
      } as never,
      spyCtx,
    )
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("BAD_ARGS")
  }
  expect(askCalls).toBe(0)
})

test("malformed target id fails BAD_ID before ask", async () => {
  let askCalls = 0
  const spyCtx = makeCtx({
    ask: async () => {
      askCalls++
    },
  })
  try {
    await officeTools.office_edit.execute(
      {
        file: path.join(FIXTURE_DIR, "report.docx"),
        operations: [{ op: "replace_text", target: "banana", anchor: "x", text: "y" }],
      } as never,
      spyCtx,
    )
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("BAD_ID")
  }
  expect(askCalls).toBe(0)
})

test("office_create rejects an empty slides array with BAD_ARGS", async () => {
  try {
    await officeTools.office_create.execute({ file: path.join(FIXTURE_DIR, "plugin-new-deck.pptx"), slides: [] } as never, ctx)
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("BAD_ARGS")
  }
})

test("office_edit happy path on a work-copy docx applies and reports backup", async () => {
  const work = path.join(FIXTURE_DIR, "plugin-work-report.docx")
  await copyFile(path.join(FIXTURE_DIR, "edit-report.docx"), work)
  const before = await readDocx(work, "content")
  const target = before.elements.find((e) => e.type === "paragraph" && e.text.includes("strong"))!

  const output = await officeTools.office_edit.execute(
    {
      file: work,
      operations: [{ op: "replace_text", target: target.id, anchor: "was strong this", text: "was exceptional this" }],
    } as never,
    ctx,
  )

  expect(String(output)).toContain("Applied 1 operation(s)")
  expect(String(output)).toContain("Backup:")
})

test("office_python runs code and captures stdout", async () => {
  const output = await officeTools.office_python.execute({ code: 'print("hi")', files: [] } as never, ctx)
  expect(String(output)).toContain("hi")
})

test("truncateForModel truncates over-limit text", () => {
  const big = "x".repeat(30_000)
  const result = truncateForModel(big)
  expect(result).toContain("[truncated:")
  expect(result.length).toBeLessThan(25_000)
})

test("truncateForModel leaves under-limit text unchanged", () => {
  const small = "hello world"
  expect(truncateForModel(small)).toBe(small)
})

test("office_render on an unsupported extension throws UNSUPPORTED_FORMAT without invoking soffice", async () => {
  try {
    await officeTools.office_render.execute({ file: "notes.txt" } as never, ctx)
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("UNSUPPORTED_FORMAT")
  }
})

test.skipIf(!HAS_SOFFICE)(
  "office_render returns png attachments for each page",
  async () => {
    const result = await officeTools.office_render.execute({ file: path.join(FIXTURE_DIR, "deck.pptx") } as never, ctx)
    if (typeof result === "string") throw new Error("expected object ToolResult with attachments")
    expect(result.attachments).toHaveLength(2)
    for (const attachment of result.attachments!) {
      expect(attachment.type).toBe("file")
      expect(attachment.mime).toBe("image/png")
      expect(attachment.url.startsWith("file://")).toBe(true)
    }
  },
  300_000,
)
