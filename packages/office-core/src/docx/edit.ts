import { OfficeError } from "../errors"
import { parseId } from "../ids"
import { backupFile } from "../backup"
import { runWorker } from "../worker"

export type DocxOperation =
  | { op: "replace_text"; target: string; anchor: string; text: string }
  | { op: "insert_content"; after: string; markdown: string }
  | { op: "delete_element"; target: string; anchor: string }
  | { op: "set_style"; target: string; anchor: string; style: string }
  | { op: "set_table_cell"; target: string; row: number; col: number; text: string; anchor?: string }

export type EditResult = { applied: number; results: Array<Record<string, unknown>>; backup?: string }

function assertDocxId(id: string): void {
  const ref = parseId(id)
  if (ref.kind !== "paragraph" && ref.kind !== "table")
    throw new OfficeError(
      "BAD_ID",
      `Target ${id} is not a docx element ID`,
      "docx targets use p:<n> or tbl:<n> — get IDs from office_read output for this file.",
    )
}

export async function editDocx(
  file: string,
  operations: DocxOperation[],
  opts?: { backup?: boolean; cacheDir?: string; timeoutMs?: number },
): Promise<EditResult> {
  for (const operation of operations) {
    assertDocxId("target" in operation ? operation.target : operation.after)
  }
  const backup = opts?.backup === false ? undefined : await backupFile(file, opts?.cacheDir)
  const data = await runWorker<Omit<EditResult, "backup">>("docx_edit.py", { file, operations }, { timeoutMs: opts?.timeoutMs, cacheDir: opts?.cacheDir })
  return { ...data, backup }
}
