import { runWorker } from "../worker"

export type DocxElement =
  | { id: string; type: "paragraph"; style: string; text: string }
  | { id: string; type: "table"; rows: number; cols: number; text?: string }

export type DocxRead = { format: "docx"; mode: string; elements: DocxElement[] }

export async function readDocx(
  file: string,
  mode: "outline" | "content",
  target?: string,
  opts?: { cacheDir?: string },
): Promise<DocxRead> {
  return runWorker<DocxRead>("docx_read.py", { file, mode, target }, opts)
}

export function formatDocxRead(result: DocxRead): string {
  return result.elements
    .map((el) =>
      el.type === "paragraph"
        ? `[${el.id}] (${el.style}) ${el.text}`
        : `[${el.id}] (table ${el.rows}x${el.cols})${el.text ? "\n" + el.text : ""}`,
    )
    .join("\n")
}
