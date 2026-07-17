import { OfficeError } from "../errors"
import { parseId } from "../ids"
import { runWorker } from "../worker"

export type PptxShape = { id: string; name: string; text: string }
export type PptxSlide = { id: string; title: string; layout: string; shapes?: PptxShape[]; notes?: string }
export type PptxRead = { format: "pptx"; mode: string; slides: PptxSlide[] }

export async function readPptx(
  file: string,
  mode: "outline" | "content",
  target?: string,
  opts?: { cacheDir?: string },
): Promise<PptxRead> {
  if (target !== undefined) {
    const ref = parseId(target)
    if (ref.kind !== "slide" && ref.kind !== "shape")
      throw new OfficeError(
        "BAD_ID",
        `Target ${target} is not a pptx element ID`,
        "pptx targets use s:<n> or s:<n>/sh:<m> — get IDs from office_read output for this file.",
      )
  }
  return runWorker<PptxRead>("pptx_read.py", { file, mode, target }, opts)
}

export function formatPptxRead(result: PptxRead): string {
  return result.slides
    .map((slide) => {
      const lines = [`[${slide.id}] ${slide.title || "(untitled)"} — layout: ${slide.layout}`]
      for (const shape of slide.shapes ?? []) lines.push(`  [${shape.id}] (${shape.name}) ${shape.text}`)
      if (slide.notes) lines.push(`  notes: ${slide.notes}`)
      return lines.join("\n")
    })
    .join("\n")
}
