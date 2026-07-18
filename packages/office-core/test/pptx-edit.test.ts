import { beforeEach, expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { copyFile } from "node:fs/promises"
import path from "node:path"
import { OfficeError } from "../src/errors"
import { editPptx } from "../src/pptx/edit"
import { readPptx } from "../src/pptx/read"
import { runWorker } from "../src/worker"
import { FIXTURE_DIR, ensureFixtures } from "./fixtures"

type ProbeResult = { pictures: Array<{ id: string; part: string; sha256: string; content_type: string }> }

const WORK = path.join(FIXTURE_DIR, "work-deck.pptx")

beforeEach(async () => {
  await ensureFixtures()
  await copyFile(path.join(FIXTURE_DIR, "edit-deck.pptx"), WORK)
}, 180_000)

test("set_shape_text replaces anchored text in a shape", async () => {
  const before = await readPptx(WORK, "content")
  const bullets = before.slides[1].shapes!.find((sh) => sh.text.includes("Second point"))!
  await editPptx(WORK, [{ op: "set_shape_text", target: bullets.id, anchor: "Second point", text: "Second point, revised" }])
  const after = await readPptx(WORK, "content", "s:1")
  expect(after.slides[0].shapes!.map((sh) => sh.text).join("\n")).toContain("Second point, revised")
})

test("set_notes overwrites speaker notes", async () => {
  await editPptx(WORK, [{ op: "set_notes", target: "s:0", text: "Open with the numbers." }])
  const after = await readPptx(WORK, "content", "s:0")
  expect(after.slides[0].notes).toBe("Open with the numbers.")
})

test("insert_slide lands directly after the reference slide with theme layout", async () => {
  await editPptx(WORK, [
    { op: "insert_slide", after: "s:0", layout: "Title and Content", title: "Agenda", bullets: ["One", "Two"] },
  ])
  const after = await readPptx(WORK, "outline")
  expect(after.slides.map((s) => s.title)).toEqual(["Edit Deck", "Agenda", "Points", ""])
  expect(after.slides[1].layout).toBe("Title and Content")
})

test("duplicate_slide copies a picture slide without breaking the file", async () => {
  await editPptx(WORK, [{ op: "duplicate_slide", target: "s:2" }])
  const after = await readPptx(WORK, "outline")
  expect(after.slides).toHaveLength(4)
})

test("delete_slide and move_slide restructure the deck", async () => {
  await editPptx(WORK, [
    { op: "delete_slide", target: "s:2" },
    { op: "move_slide", target: "s:1", index: 0 },
  ])
  const after = await readPptx(WORK, "outline")
  expect(after.slides.map((s) => s.title)).toEqual(["Points", "Edit Deck"])
})

test("duplicate + replace_image: bytes swap on the copy, original untouched", async () => {
  const swap = path.join(FIXTURE_DIR, "swap.png")
  const swapHash = createHash("sha256").update(Buffer.from(await Bun.file(swap).arrayBuffer())).digest("hex")
  await editPptx(WORK, [{ op: "duplicate_slide", target: "s:2" }])
  const before = await runWorker<ProbeResult>("pptx_probe.py", { file: WORK })
  const originalHash = before.pictures.find((p) => p.id === "s:2/sh:0")!.sha256
  expect(before.pictures.find((p) => p.id === "s:3/sh:0")!.sha256).toBe(originalHash)
  await editPptx(WORK, [{ op: "replace_image", target: "s:3/sh:0", image: swap }])
  const after = await runWorker<ProbeResult>("pptx_probe.py", { file: WORK })
  expect(after.pictures.find((p) => p.id === "s:3/sh:0")!.sha256).toBe(swapHash)
  expect(after.pictures.find((p) => p.id === "s:3/sh:0")!.content_type).toBe("image/png")
  expect(after.pictures.find((p) => p.id === "s:2/sh:0")!.sha256).toBe(originalHash)
})

test("replace_image on a text shape fails with SHAPE_NOT_PICTURE", async () => {
  const swap = path.join(FIXTURE_DIR, "swap.png")
  try {
    await editPptx(WORK, [{ op: "replace_image", target: "s:0/sh:0", image: swap }])
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("SHAPE_NOT_PICTURE")
  }
})

test("unknown layout lists available layouts", async () => {
  try {
    await editPptx(WORK, [{ op: "insert_slide", after: "s:0", layout: "Nope", title: "x" }])
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("LAYOUT_NOT_FOUND")
    expect((e as OfficeError).hint).toContain("Title and Content")
  }
})

test("failed op aborts the batch atomically", async () => {
  const before = await readPptx(WORK, "content")
  try {
    await editPptx(WORK, [
      { op: "set_notes", target: "s:0", text: "should not survive" },
      { op: "set_shape_text", target: "s:1/sh:1", anchor: "NOT THERE", text: "x" },
    ])
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("ANCHOR_MISMATCH")
  }
  const after = await readPptx(WORK, "content")
  expect(after.slides).toEqual(before.slides)
})

test("docx-style target is rejected client-side", async () => {
  try {
    await editPptx(WORK, [{ op: "set_notes", target: "p:0", text: "x" }])
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("BAD_ID")
  }
})

test("slide op with shape target is rejected with BAD_TARGET_KIND", async () => {
  try {
    await editPptx(WORK, [{ op: "delete_slide", target: "s:0/sh:1" }])
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("BAD_TARGET_KIND")
  }
})

test("shape op with slide target is rejected with BAD_TARGET_KIND", async () => {
  try {
    await editPptx(WORK, [{ op: "set_shape_text", target: "s:1", anchor: "x", text: "y" }])
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("BAD_TARGET_KIND")
  }
})
