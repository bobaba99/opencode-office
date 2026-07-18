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
