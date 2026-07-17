import { expect, test } from "bun:test"
import { existsSync } from "node:fs"
import path from "node:path"
import { FIXTURE_DIR, ensureFixtures } from "./fixtures"

test("generates docx and pptx fixtures", async () => {
  await ensureFixtures()
  expect(existsSync(path.join(FIXTURE_DIR, "report.docx"))).toBe(true)
  expect(existsSync(path.join(FIXTURE_DIR, "deck.pptx"))).toBe(true)
}, 180_000)

test("generates edit fixtures", async () => {
  await ensureFixtures()
  for (const name of ["edit-report.docx", "edit-deck.pptx", "swap.png"]) {
    expect(existsSync(path.join(FIXTURE_DIR, name))).toBe(true)
  }
}, 180_000)
