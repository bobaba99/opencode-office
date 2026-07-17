import { beforeAll, expect, test } from "bun:test"
import path from "node:path"
import { OfficeError } from "../src/errors"
import { formatDocxRead, readDocx } from "../src/docx/read"
import { FIXTURE_DIR, ensureFixtures } from "./fixtures"

const REPORT = () => path.join(FIXTURE_DIR, "report.docx")

beforeAll(async () => {
  await ensureFixtures()
}, 180_000)

test("outline returns only headings, with IDs and styles", async () => {
  const out = await readDocx(REPORT(), "outline")
  expect(out.elements.map((e) => (e.type === "paragraph" ? e.text : e.type))).toEqual([
    "Quarterly Report",
    "Regional Breakdown",
    "table",
  ])
  const first = out.elements[0]
  expect(first.type).toBe("paragraph")
  if (first.type === "paragraph") expect(first.style).toBe("Heading 1")
  expect(first.id).toMatch(/^p:\d+$/)
})

test("content returns every block including table text", async () => {
  const out = await readDocx(REPORT(), "content")
  const table = out.elements.find((e) => e.type === "table")
  expect(table).toBeDefined()
  if (table?.type === "table") {
    expect(table.rows).toBe(2)
    expect(table.text).toContain("EMEA | $4.2M")
  }
  expect(formatDocxRead(out)).toContain("Q3 revenue grew 12% year over year.")
})

test("target narrows to a single element", async () => {
  const all = await readDocx(REPORT(), "content")
  const target = all.elements[1].id
  const one = await readDocx(REPORT(), "content", target)
  expect(one.elements).toHaveLength(1)
  expect(one.elements[0].id).toBe(target)
})

test("unknown target raises TARGET_NOT_FOUND", async () => {
  try {
    await readDocx(REPORT(), "content", "p:999")
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("TARGET_NOT_FOUND")
  }
})
