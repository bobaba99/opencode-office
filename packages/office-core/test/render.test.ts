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

test.skipIf(!HAS_SOFFICE)("pages: [] renders nothing (distinct from omitted = all)", async () => {
  const result = await renderOffice(path.join(FIXTURE_DIR, "deck.pptx"), { pages: [] })
  expect(result.pages).toEqual([])
}, 300_000)

test.skipIf(!HAS_SOFFICE)("concurrent renders both succeed", async () => {
  const [deckResult, reportResult] = await Promise.all([
    renderOffice(path.join(FIXTURE_DIR, "deck.pptx"), { outDir: path.join(FIXTURE_DIR, "..", ".concurrent-out", "deck") }),
    renderOffice(path.join(FIXTURE_DIR, "report.docx"), { pages: [1], outDir: path.join(FIXTURE_DIR, "..", ".concurrent-out", "report") }),
  ])
  expect(deckResult.pages.map((p) => p.page)).toEqual([1, 2])
  for (const page of deckResult.pages) {
    expect(existsSync(page.path)).toBe(true)
  }
  expect(reportResult.pages).toHaveLength(1)
  expect(existsSync(reportResult.pages[0].path)).toBe(true)
}, 600_000)
