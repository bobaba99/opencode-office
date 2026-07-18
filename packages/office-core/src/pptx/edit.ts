import { OfficeError } from "../errors"
import { parseId } from "../ids"
import { backupFile } from "../backup"
import { runWorker } from "../worker"
import type { EditResult } from "../docx/edit"

export type PptxOperation =
  | { op: "set_shape_text"; target: string; anchor: string; text: string }
  | { op: "set_notes"; target: string; text: string }
  | { op: "insert_slide"; after: string; layout: string; title?: string; bullets?: string[] }
  | { op: "duplicate_slide"; target: string }
  | { op: "delete_slide"; target: string }
  | { op: "move_slide"; target: string; index: number }
  | { op: "replace_image"; target: string; image: string }

const SHAPE_OPS = new Set(["set_shape_text", "replace_image"])

function assertPptxId(id: string, op: string): void {
  const ref = parseId(id)
  if (ref.kind !== "slide" && ref.kind !== "shape")
    throw new OfficeError(
      "BAD_ID",
      `Target ${id} is not a pptx element ID`,
      "pptx targets use s:<n> or s:<n>/sh:<m> — get IDs from office_read output for this file.",
    )
  const needShape = SHAPE_OPS.has(op)
  if (needShape && ref.kind !== "shape")
    throw new OfficeError("BAD_TARGET_KIND", `${op} needs a shape target (s:<n>/sh:<m>), got ${id}`, "office_read content mode lists each slide's shape IDs.")
  if (!needShape && ref.kind !== "slide")
    throw new OfficeError("BAD_TARGET_KIND", `${op} needs a slide target (s:<n>), got ${id}`, "Use the slide ID without the /sh:<m> suffix.")
}

export async function editPptx(
  file: string,
  operations: PptxOperation[],
  opts?: { backup?: boolean; cacheDir?: string; timeoutMs?: number },
): Promise<EditResult> {
  for (const operation of operations) {
    assertPptxId("target" in operation ? operation.target : operation.after, operation.op)
  }
  const backup = opts?.backup === false ? undefined : await backupFile(file, opts?.cacheDir)
  const data = await runWorker<Omit<EditResult, "backup">>("pptx_edit.py", { file, operations }, { timeoutMs: opts?.timeoutMs, cacheDir: opts?.cacheDir })
  return { ...data, backup }
}
