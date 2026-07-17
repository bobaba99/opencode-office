import { beforeAll, expect, test } from "bun:test"
import path from "node:path"
import { OfficeError } from "../src/errors"
import { formatPptxRead, readPptx } from "../src/pptx/read"
import { FIXTURE_DIR, ensureFixtures } from "./fixtures"

const DECK = () => path.join(FIXTURE_DIR, "deck.pptx")

beforeAll(async () => {
  await ensureFixtures()
}, 180_000)

test("outline lists slides with titles and layouts", async () => {
  const out = await readPptx(DECK(), "outline")
  expect(out.slides.map((s) => s.title)).toEqual(["Q3 Review", "Highlights"])
  expect(out.slides[0].id).toBe("s:0")
  expect(out.slides[0].shapes).toBeUndefined()
})

test("content includes shape text and speaker notes", async () => {
  const out = await readPptx(DECK(), "content")
  const highlights = out.slides[1]
  expect(highlights.notes).toBe("Pause here for questions.")
  const texts = (highlights.shapes ?? []).map((sh) => sh.text).join("\n")
  expect(texts).toContain("EMEA leads growth")
  expect((highlights.shapes ?? [])[0].id).toMatch(/^s:1\/sh:\d+$/)
  expect(formatPptxRead(out)).toContain("Q3 Review")
})

test("target narrows to one slide", async () => {
  const out = await readPptx(DECK(), "content", "s:1")
  expect(out.slides).toHaveLength(1)
  expect(out.slides[0].title).toBe("Highlights")
})

test("unknown slide raises TARGET_NOT_FOUND", async () => {
  try {
    await readPptx(DECK(), "content", "s:9")
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("TARGET_NOT_FOUND")
  }
})
