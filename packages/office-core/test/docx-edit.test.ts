import { beforeEach, expect, test } from "bun:test"
import { copyFile } from "node:fs/promises"
import path from "node:path"
import { OfficeError } from "../src/errors"
import { editDocx } from "../src/docx/edit"
import { readDocx } from "../src/docx/read"
import { FIXTURE_DIR, ensureFixtures } from "./fixtures"

const WORK = path.join(FIXTURE_DIR, "work-report.docx")

beforeEach(async () => {
  await ensureFixtures()
  await copyFile(path.join(FIXTURE_DIR, "edit-report.docx"), WORK)
}, 180_000)

test("replace_text across runs preserves surrounding text", async () => {
  const before = await readDocx(WORK, "content")
  const target = before.elements.find((e) => e.type === "paragraph" && e.text.includes("strong"))!
  const result = await editDocx(WORK, [
    { op: "replace_text", target: target.id, anchor: "was strong this", text: "was exceptional this" },
  ])
  expect(result.applied).toBe(1)
  const after = await readDocx(WORK, "content", target.id)
  expect((after.elements[0] as { text: string }).text).toBe("Growth was exceptional this quarter overall.")
})

test("batch applies in order and untouched elements are unchanged", async () => {
  const before = await readDocx(WORK, "content")
  const del = before.elements.find((e) => e.type === "paragraph" && e.text === "Delete me entirely.")!
  const style = before.elements.find((e) => e.type === "paragraph" && e.text === "Style me.")!
  await editDocx(WORK, [
    { op: "delete_element", target: del.id, anchor: "Delete me entirely." },
    // NOTE: after the delete, positional IDs above the deleted element shift down by one.
    { op: "set_style", target: `p:${Number(style.id.split(":")[1]) - 1}`, anchor: "Style me.", style: "Heading 2" },
  ])
  const after = await readDocx(WORK, "content")
  expect(after.elements.some((e) => e.type === "paragraph" && e.text === "Delete me entirely.")).toBe(false)
  const styled = after.elements.find((e) => e.type === "paragraph" && e.text === "Style me.")!
  expect(styled.type === "paragraph" && styled.style).toBe("Heading 2")
  expect(after.elements[0]).toEqual(before.elements[0])
})

test("insert_content inserts styled markdown after an element", async () => {
  const before = await readDocx(WORK, "content")
  const heading = before.elements[0]
  await editDocx(WORK, [
    { op: "insert_content", after: heading.id, markdown: "## New Section\nIntro line.\n- bullet one" },
  ])
  const after = await readDocx(WORK, "content")
  const texts = after.elements.map((e) => (e.type === "paragraph" ? e.text : "table"))
  expect(texts.slice(1, 4)).toEqual(["New Section", "Intro line.", "bullet one"])
  const section = after.elements[1]
  expect(section.type === "paragraph" && section.style).toBe("Heading 2")
})

test("set_table_cell updates one cell", async () => {
  const before = await readDocx(WORK, "content")
  const tbl = before.elements.find((e) => e.type === "table")!
  await editDocx(WORK, [{ op: "set_table_cell", target: tbl.id, row: 1, col: 1, text: "two", anchor: "one" }])
  const after = await readDocx(WORK, "content", tbl.id)
  expect((after.elements[0] as { text?: string }).text).toContain("alpha | two")
})

test("anchor mismatch aborts the whole batch atomically", async () => {
  const before = await readDocx(WORK, "content")
  const first = before.elements.find((e) => e.type === "paragraph" && e.text.includes("Growth"))!
  try {
    await editDocx(WORK, [
      { op: "replace_text", target: first.id, anchor: "Growth was ", text: "Expansion was " },
      { op: "replace_text", target: first.id, anchor: "NOT PRESENT", text: "x" },
    ])
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("ANCHOR_MISMATCH")
    expect((e as OfficeError).hint).toContain("Growth")
  }
  const after = await readDocx(WORK, "content")
  expect(after.elements).toEqual(before.elements)
})

test("ambiguous anchor is rejected with guidance", async () => {
  await editDocx(WORK, [
    { op: "insert_content", after: "p:0", markdown: "same same" },
  ])
  const doc = await readDocx(WORK, "content")
  const dup = doc.elements.find((e) => e.type === "paragraph" && e.text === "same same")!
  try {
    await editDocx(WORK, [{ op: "replace_text", target: dup.id, anchor: "same", text: "x" }])
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("AMBIGUOUS_ANCHOR")
  }
})

test("unknown style yields STYLE_NOT_FOUND", async () => {
  try {
    await editDocx(WORK, [{ op: "set_style", target: "p:0", anchor: "Edit Playground", style: "No Such Style" }])
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("STYLE_NOT_FOUND")
  }
})

test("backup is written before editing and contains the original", async () => {
  const original = await Bun.file(WORK).arrayBuffer()
  const result = await editDocx(WORK, [
    { op: "replace_text", target: "p:1", anchor: "Growth", text: "Expansion" },
  ])
  expect(result.backup).toBeDefined()
  const backup = await Bun.file(result.backup!).arrayBuffer()
  expect(Buffer.from(backup).equals(Buffer.from(original))).toBe(true)
})

test("pptx-style target is rejected client-side", async () => {
  try {
    await editDocx(WORK, [{ op: "replace_text", target: "s:0", anchor: "x", text: "y" }])
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("BAD_ID")
  }
})
