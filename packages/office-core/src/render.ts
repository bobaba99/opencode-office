import { existsSync } from "node:fs"
import { mkdir, mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { OfficeError } from "./errors"
import { acquireLock, defaultCacheDir, findSoffice } from "./runtime"
import { runWorker } from "./worker"

export type RenderResult = { pages: Array<{ page: number; path: string; width: number; height: number }> }

async function convertToPdf(soffice: string, file: string, profileDir: string, outDir: string, timeoutMs: number): Promise<{ code: number; stderr: string }> {
  const proc = Bun.spawn(
    [soffice, "--headless", `-env:UserInstallation=file://${profileDir}`, "--convert-to", "pdf", "--outdir", outDir, file],
    { stdout: "pipe", stderr: "pipe" },
  )
  let timedOut = false
  const timer = setTimeout(() => {
    timedOut = true
    proc.kill()
    setTimeout(() => proc.kill(9), 5_000).unref?.()
  }, timeoutMs)
  const [, stderr, code] = await Promise.all([new Response(proc.stdout).text(), new Response(proc.stderr).text(), proc.exited])
  clearTimeout(timer)
  if (timedOut)
    throw new OfficeError(
      "RENDER_FAILED",
      `soffice exceeded ${timeoutMs}ms converting ${file} to pdf`,
      "LibreOffice conversion can be slow on first run (profile/font setup); retry once, or raise timeoutMs.",
    )
  return { code, stderr }
}

export async function renderOffice(
  file: string,
  opts?: { pages?: number[]; outDir?: string; cacheDir?: string; timeoutMs?: number },
): Promise<RenderResult> {
  const cacheDir = opts?.cacheDir ?? defaultCacheDir()
  const soffice = await findSoffice()
  if (!soffice)
    throw new OfficeError(
      "RENDER_UNAVAILABLE",
      "LibreOffice (soffice) is required for rendering",
      "macOS: brew install --cask libreoffice. Linux: apt install libreoffice. Reads and edits work without it.",
    )

  const timeoutMs = opts?.timeoutMs ?? 120_000
  // Isolated from the user's real LibreOffice profile (avoids first-run wizards and
  // config collisions), but reused across calls like the venv cache — wiping it on
  // every call would both re-pay soffice's profile/font init cost each time and race
  // with any other renderOffice() call still using it concurrently.
  const profileDir = path.join(cacheDir, "render-profile")
  await mkdir(profileDir, { recursive: true })

  const tmpDir = await mkdtemp(path.join(os.tmpdir(), "opencode-office-render-"))
  try {
    // soffice conversions against the same UserInstallation profile dir corrupt each
    // other when run concurrently (the second call exits 0 but produces no output) —
    // serialize just the soffice invocation through the profile-dir lock; the pymupdf
    // rasterization worker below runs outside it since it doesn't touch the profile.
    const release = await acquireLock(profileDir, {
      timeoutMs: Math.max(180_000, timeoutMs * 3 + 30_000),
      // staleMs must scale with the wait budget above it: a legitimately long conversion
      // (opts.timeoutMs > ~270s) would otherwise still be running past the 300s default
      // staleness threshold and have its ACTIVE lock stolen out from under it mid-convert.
      staleMs: Math.max(300_000, timeoutMs + 60_000),
    })
    let code: number, stderr: string
    try {
      ;({ code, stderr } = await convertToPdf(soffice, file, profileDir, tmpDir, timeoutMs))
    } finally {
      await release()
    }
    const base = path.basename(file, path.extname(file))
    const pdfPath = path.join(tmpDir, `${base}.pdf`)
    if (code !== 0 || !existsSync(pdfPath)) {
      const message =
        code === 0
          ? `soffice exited 0 but produced no output converting ${file} to pdf`
          : `soffice failed to convert ${file} to pdf (exit ${code})`
      const hint =
        code === 0
          ? `Retry once — LibreOffice can flake on first run. Renders are serialized through a provisioning lock, so profile contention shouldn't cause this anymore; if it persists, check for a stale lock dir at ${profileDir}.lock or confirm the file opens in LibreOffice directly. stderr: ${(stderr || "none").slice(-1500)}`
          : `Retry once — LibreOffice can flake on first run; if it persists, confirm the file opens in LibreOffice directly. stderr: ${(stderr || "none").slice(-1500)}`
      throw new OfficeError("RENDER_FAILED", message, hint)
    }

    const outDir = opts?.outDir ?? path.join(cacheDir, "renders", base)
    await mkdir(outDir, { recursive: true })

    return await runWorker<RenderResult>("render_pdf.py", { pdf: pdfPath, outDir, pages: opts?.pages }, { cacheDir, timeoutMs })
  } finally {
    await rm(tmpDir, { recursive: true, force: true })
  }
}
