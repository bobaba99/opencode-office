import { runWorker } from "../worker"

export type CreateDocxResult = { file: string; paragraphs: number }

export async function createDocx(
  file: string,
  markdown: string,
  opts?: { reference?: string; cacheDir?: string; timeoutMs?: number },
): Promise<CreateDocxResult> {
  return runWorker<CreateDocxResult>(
    "docx_create.py",
    { file, markdown, reference: opts?.reference },
    { timeoutMs: opts?.timeoutMs, cacheDir: opts?.cacheDir },
  )
}
