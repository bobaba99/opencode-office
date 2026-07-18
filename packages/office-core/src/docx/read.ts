import { OfficeError } from "../errors"
import { parseId } from "../ids"
import { runWorker } from "../worker"

export type DocxElement =
  | { id: string; type: "paragraph"; style: string; text: string; tracked_insertions?: string[]; tracked_deletions?: string[] }
  | { id: string; type: "table"; rows: number; cols: number; text?: string }

export type DocxRead = { format: "docx"; mode: string; elements: DocxElement[]; comments?: Array<{ id: number; author: string; text: string }> }

export async function readDocx(
  file: string,
  mode: "outline" | "content" | "full",
  target?: string,
  opts?: { cacheDir?: string; timeoutMs?: number },
): Promise<DocxRead> {
  if (target !== undefined) {
    const ref = parseId(target)
    if (ref.kind !== "paragraph" && ref.kind !== "table")
      throw new OfficeError(
        "BAD_ID",
        `Target ${target} is not a docx element ID`,
        "docx targets use p:<n> or tbl:<n> — get IDs from office_read output for this file.",
      )
  }
  return runWorker<DocxRead>("docx_read.py", { file, mode, target }, { timeoutMs: opts?.timeoutMs, cacheDir: opts?.cacheDir })
}

export function formatDocxRead(result: DocxRead): string {
  const elementLines = result.elements
    .map((el) =>
      el.type === "paragraph"
        ? `[${el.id}] (${el.style}) ${el.text}` +
          (el.tracked_insertions ?? []).map((t) => `\n  tracked insertion: ${JSON.stringify(t)}`).join("") +
          (el.tracked_deletions ?? []).map((t) => `\n  tracked deletion: ${JSON.stringify(t)}`).join("")
        : `[${el.id}] (table ${el.rows}x${el.cols})${el.text ? "\n" + el.text : ""}`,
    )
    .join("\n")

  if (result.comments && result.comments.length > 0) {
    const commentLines = result.comments.map((c) => `  [${c.id}] ${c.author}: ${c.text}`).join("\n")
    return elementLines + "\n\nComments:\n" + commentLines
  }

  return elementLines
}
