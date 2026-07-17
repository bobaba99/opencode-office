import { OfficeError } from "./errors"

export type ElementRef =
  | { kind: "paragraph"; index: number }
  | { kind: "table"; index: number }
  | { kind: "slide"; index: number }
  | { kind: "shape"; slide: number; shape: number }

export function formatId(ref: ElementRef): string {
  switch (ref.kind) {
    case "paragraph":
      return `p:${ref.index}`
    case "table":
      return `tbl:${ref.index}`
    case "slide":
      return `s:${ref.index}`
    case "shape":
      return `s:${ref.slide}/sh:${ref.shape}`
  }
}

export function parseId(id: string): ElementRef {
  const shape = id.match(/^s:(\d+)\/sh:(\d+)$/)
  if (shape) return { kind: "shape", slide: Number(shape[1]), shape: Number(shape[2]) }
  const simple = id.match(/^(p|tbl|s):(\d+)$/)
  if (simple) {
    const index = Number(simple[2])
    if (simple[1] === "p") return { kind: "paragraph", index }
    if (simple[1] === "tbl") return { kind: "table", index }
    return { kind: "slide", index }
  }
  throw new OfficeError(
    "BAD_ID",
    `Unrecognized element ID: ${id}`,
    "Valid forms: p:12, tbl:3, s:4, s:4/sh:2 — IDs come from office_read output.",
  )
}
