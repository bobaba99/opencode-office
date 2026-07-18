import { beforeEach, expect, test } from "bun:test"
import { mkdir, rm } from "node:fs/promises"
import path from "node:path"
import { editDocx } from "../packages/office-core/src/docx/edit"
import { editPptx } from "../packages/office-core/src/pptx/edit"
import { readDocx } from "../packages/office-core/src/docx/read"
import { FIXTURE_DIR, ensureFixtures } from "../packages/office-core/test/fixtures"
import { CARDS } from "./cards"

// Unit-tests the score checks WITHOUT running any agent: apply edits directly via
// office-core, then assert `check()` scores the resulting file correctly. This is the
// ground-truth layer the harness leans on — it must be right before any agent runs.

const ARENA = path.join("/tmp", "office-eval-score-test")

beforeEach(async () => {
  await ensureFixtures()
  await rm(ARENA, { recursive: true, force: true })
  await mkdir(ARENA, { recursive: true })
}, 180_000)

function card(id: string) {
  const found = CARDS.find((c) => c.id === id)
  if (!found) throw new Error(`missing card ${id}`)
  return found
}

async function copyCardFixtures(id: string): Promise<void> {
  for (const file of card(id).files) {
    await Bun.write(path.join(ARENA, file), Bun.file(path.join(FIXTURE_DIR, file)))
  }
}

test("docx-replace: correct edit scores success:true fidelity:true", async () => {
  await copyCardFixtures("docx-replace")
  const file = path.join(ARENA, "edit-report.docx")
  const before = await readDocx(file, "content")
  const target = before.elements.find((e) => e.type === "paragraph" && e.text.includes("strong"))!
  await editDocx(file, [{ op: "replace_text", target: target.id, anchor: "strong", text: "robust" }])

  const result = await card("docx-replace").check(ARENA)
  expect(result.success).toBe(true)
  expect(result.fidelity).toBe(true)
})

test("docx-replace: wrong replacement word flips success to false", async () => {
  await copyCardFixtures("docx-replace")
  const file = path.join(ARENA, "edit-report.docx")
  const before = await readDocx(file, "content")
  const target = before.elements.find((e) => e.type === "paragraph" && e.text.includes("strong"))!
  await editDocx(file, [{ op: "replace_text", target: target.id, anchor: "strong", text: "mediocre" }])

  const result = await card("docx-replace").check(ARENA)
  expect(result.success).toBe(false)
})

test("docx-replace: correct text edit plus deleting the table flips fidelity to false", async () => {
  await copyCardFixtures("docx-replace")
  const file = path.join(ARENA, "edit-report.docx")
  const before = await readDocx(file, "content")
  const target = before.elements.find((e) => e.type === "paragraph" && e.text.includes("strong"))!
  await editDocx(file, [{ op: "replace_text", target: target.id, anchor: "strong", text: "robust" }])
  const after = await readDocx(file, "content")
  const table = after.elements.find((e) => e.type === "table")!
  await editDocx(file, [{ op: "delete_element", target: table.id, anchor: "K" }])

  const result = await card("docx-replace").check(ARENA)
  expect(result.success).toBe(true)
  expect(result.fidelity).toBe(false)
})

test("pptx-image: correct edit (duplicate then swap the duplicate's picture) scores success:true fidelity:true", async () => {
  await copyCardFixtures("pptx-image")
  const file = path.join(ARENA, "edit-deck.pptx")
  const swap = path.join(ARENA, "swap.png")
  await editPptx(file, [{ op: "duplicate_slide", target: "s:2" }])
  await editPptx(file, [{ op: "replace_image", target: "s:3/sh:0", image: swap }])

  const result = await card("pptx-image").check(ARENA)
  expect(result.success).toBe(true)
  expect(result.fidelity).toBe(true)
})

test("pptx-image: duplicate landing at a different index than the native tool's default placement still scores success:true fidelity:true", async () => {
  // The prompt doesn't pin where the duplicate lands. duplicate_slide's native placement puts
  // it right after the source (s:3 here), but this simulates a different tool path relocating
  // it further down the deck (s:4) before the picture swap — the scoring must locate the
  // duplicate by content, not by the fixed index the native op happens to produce.
  await copyCardFixtures("pptx-image")
  const file = path.join(ARENA, "edit-deck.pptx")
  const swap = path.join(ARENA, "swap.png")
  await editPptx(file, [{ op: "duplicate_slide", target: "s:2" }])
  await editPptx(file, [{ op: "move_slide", target: "s:3", index: 4 }])
  await editPptx(file, [{ op: "replace_image", target: "s:4/sh:0", image: swap }])

  const result = await card("pptx-image").check(ARENA)
  expect(result.success).toBe(true)
  expect(result.fidelity).toBe(true)
})

test("pptx-image: swapping the ORIGINAL's picture instead of the duplicate's flips both flags", async () => {
  await copyCardFixtures("pptx-image")
  const file = path.join(ARENA, "edit-deck.pptx")
  const swap = path.join(ARENA, "swap.png")
  await editPptx(file, [{ op: "duplicate_slide", target: "s:2" }])
  await editPptx(file, [{ op: "replace_image", target: "s:2/sh:0", image: swap }])

  const result = await card("pptx-image").check(ARENA)
  expect(result.success).toBe(false)
  expect(result.fidelity).toBe(false)
})

test("docx-table: correct cell edit with corrupting header edit scores success:true fidelity:false", async () => {
  await copyCardFixtures("docx-table")
  const file = path.join(ARENA, "report.docx")
  const before = await readDocx(file, "content")
  const table = before.elements.find((e) => e.type === "table")!

  // Apply correct edit: cell(1,1) → "$5.1M" and corrupting edit: cell(0,0) → "Territory"
  await editDocx(file, [
    { op: "set_table_cell", target: table.id, row: 1, col: 1, text: "$5.1M" },
    { op: "set_table_cell", target: table.id, row: 0, col: 0, text: "Territory" },
  ])

  const result = await card("docx-table").check(ARENA)
  expect(result.success).toBe(true)
  expect(result.fidelity).toBe(false)
})
