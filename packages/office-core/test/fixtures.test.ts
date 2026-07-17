import { expect, test } from "bun:test"
import { existsSync } from "node:fs"
import path from "node:path"
import { FIXTURE_DIR, ensureFixtures } from "./fixtures"

test("generates docx and pptx fixtures", async () => {
  await ensureFixtures()
  expect(existsSync(path.join(FIXTURE_DIR, "report.docx"))).toBe(true)
  expect(existsSync(path.join(FIXTURE_DIR, "deck.pptx"))).toBe(true)
}, 180_000)
