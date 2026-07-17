import { expect, test } from "bun:test"
import { formatId, parseId } from "../src/ids"
import { OfficeError } from "../src/errors"

test("round-trips every ID form", () => {
  for (const id of ["p:12", "tbl:3", "s:4", "s:4/sh:2"]) {
    expect(formatId(parseId(id))).toBe(id)
  }
})

test("parses shape IDs into slide and shape indices", () => {
  expect(parseId("s:4/sh:2")).toEqual({ kind: "shape", slide: 4, shape: 2 })
})

test("rejects malformed IDs with BAD_ID", () => {
  try {
    parseId("slide-4")
    expect.unreachable()
  } catch (e) {
    expect(e).toBeInstanceOf(OfficeError)
    expect((e as OfficeError).code).toBe("BAD_ID")
    expect((e as OfficeError).hint).toContain("p:12")
  }
})
