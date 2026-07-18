import { runWorker } from "../worker"

export type PptxSlideSpec = { layout: string; title?: string; bullets?: string[]; notes?: string }

export type CreatePptxResult = { file: string; slides: number; skipped: Array<{ slide: number; field: string }> }

export async function createPptx(
  file: string,
  slides: PptxSlideSpec[],
  opts?: { template?: string; cacheDir?: string; timeoutMs?: number },
): Promise<CreatePptxResult> {
  return runWorker<CreatePptxResult>(
    "pptx_create.py",
    { file, slides, template: opts?.template },
    { timeoutMs: opts?.timeoutMs, cacheDir: opts?.cacheDir },
  )
}
