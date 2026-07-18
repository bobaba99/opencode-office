import { existsSync } from "node:fs"
import { mkdir, readFile, writeFile, rm, stat } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { OfficeError } from "./errors"

export const PINNED: Record<string, string> = {
  "python-docx": "1.2.0",
  "python-pptx": "1.0.2",
  pillow: "11.3.0",
  pymupdf: "1.26.3",
}

export function defaultCacheDir(): string {
  return path.join(os.homedir(), ".cache", "opencode-office")
}

function fingerprint(): string {
  return Object.entries(PINNED)
    .map(([name, version]) => `${name}==${version}`)
    .sort()
    .join("\n")
}

export function venvPython(venvDir: string): string {
  return process.platform === "win32"
    ? path.join(venvDir, "Scripts", "python.exe")
    : path.join(venvDir, "bin", "python")
}

export async function venvIsCurrent(venvDir: string): Promise<boolean> {
  const marker = path.join(venvDir, ".fingerprint")
  if (!existsSync(venvPython(venvDir)) || !existsSync(marker)) return false
  return (await readFile(marker, "utf8")) === fingerprint()
}

async function run(cmd: string[]): Promise<{ code: number; stderr: string }> {
  const proc = Bun.spawn(cmd, { stdout: "pipe", stderr: "pipe" })
  const [code, stderr] = await Promise.all([proc.exited, new Response(proc.stderr).text()])
  return { code, stderr }
}

async function which(bin: string): Promise<boolean> {
  const probe = process.platform === "win32" ? ["where", bin] : ["which", bin]
  return (await run(probe)).code === 0
}

const SOFFICE_CANDIDATES = ["/opt/homebrew/bin/soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice", "/usr/bin/soffice"]

export async function findSoffice(): Promise<string | null> {
  const probe = process.platform === "win32" ? ["where", "soffice"] : ["which", "soffice"]
  const proc = Bun.spawn(probe, { stdout: "pipe", stderr: "pipe" })
  const [code, stdout] = await Promise.all([proc.exited, new Response(proc.stdout).text()])
  if (code === 0) {
    const resolved = stdout.split("\n")[0]?.trim()
    if (resolved) return resolved
  }
  for (const candidate of SOFFICE_CANDIDATES) {
    if (existsSync(candidate)) return candidate
  }
  return null
}

async function checkVersion(python: string): Promise<void> {
  const check = await run([python, "-c", "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)"])
  if (check.code !== 0)
    throw new OfficeError(
      "PYTHON_TOO_OLD",
      "Python 3.10+ required for Office tools — install python3 or uv",
      "Your python3 is older than 3.10. Install a newer python3, or install uv (https://docs.astral.sh/uv/) which manages its own.",
    )
}

export async function acquireLock(
  dir: string,
  opts?: { staleMs?: number; timeoutMs?: number },
): Promise<() => Promise<void>> {
  const lockDir = dir + ".lock"
  const staleMs = opts?.staleMs ?? 300_000
  const deadline = Date.now() + (opts?.timeoutMs ?? 120_000)
  await mkdir(path.dirname(lockDir), { recursive: true })
  for (;;) {
    try {
      await mkdir(lockDir)
      return async () => {
        await rm(lockDir, { recursive: true, force: true })
      }
    } catch {
      let lockAge: number | null = null
      try {
        lockAge = Date.now() - (await stat(lockDir)).mtimeMs
        if (lockAge > staleMs) await rm(lockDir, { recursive: true, force: true })
      } catch {
        // lock vanished or is unreadable; fall through to deadline + sleep, then retry
      }
      if (Date.now() > deadline) {
        const ageStr = lockAge !== null ? `${Math.round(lockAge / 1000)}s old` : "unknown age"
        const hint =
          lockAge !== null && lockAge > staleMs
            ? `Lock at ${lockDir} looks stale (${Math.round(lockAge / 1000)}s old); remove it with: rm -rf ${lockDir}`
            : `Another operation holds the lock at ${lockDir} (age ${ageStr}). It appears active — wait and retry; only remove the directory if no opencode/office process is running.`
        throw new OfficeError(
          "LOCK_TIMEOUT",
          `Timed out waiting for the office lock at ${lockDir}`,
          hint,
        )
      }
      await new Promise((resolve) => setTimeout(resolve, 250))
    }
  }
}

export async function ensureVenv(cacheDir = defaultCacheDir()): Promise<string> {
  const venvDir = path.join(cacheDir, "venv")
  const python = venvPython(venvDir)
  if (await venvIsCurrent(venvDir)) return python
  // Must exist before acquireLock so the very first office call on a fresh machine
  // isn't relying on the lock's own mkdir to create the cache directory's parent.
  await mkdir(cacheDir, { recursive: true })
  const release = await acquireLock(path.join(cacheDir, "venv"))
  try {
    if (await venvIsCurrent(venvDir)) return python
    // A stale venv (fingerprint mismatch from a PINNED change) must be cleared first —
    // `uv venv` refuses to create over an existing directory without --clear.
    await rm(venvDir, { recursive: true, force: true })
    const pkgs = Object.entries(PINNED).map(([name, version]) => `${name}==${version}`)
    if (await which("uv")) {
      const created = await run(["uv", "venv", venvDir])
      if (created.code !== 0)
        throw new OfficeError("VENV_CREATE", "uv could not create the Office tools venv", `Retry once; if it persists, update uv (\`uv self update\`) and check disk space. stderr: ${created.stderr.slice(0, 400)}`)
      await checkVersion(python)
      const installed = await run(["uv", "pip", "install", "--python", python, ...pkgs])
      if (installed.code !== 0)
        throw new OfficeError("VENV_INSTALL", "Failed to install pinned Office python packages", `Check network access to PyPI and retry; if it persists, delete ~/.cache/opencode-office/venv and retry. stderr: ${installed.stderr.slice(0, 400)}`)
    } else if (await which("python3")) {
      const created = await run(["python3", "-m", "venv", venvDir])
      if (created.code !== 0)
        throw new OfficeError("VENV_CREATE", "python3 -m venv failed", `Ensure the venv module is available (Debian/Ubuntu: \`apt install python3-venv\`) and retry. stderr: ${created.stderr.slice(0, 400)}`)
      await checkVersion(python)
      const installed = await run([python, "-m", "pip", "install", "--quiet", ...pkgs])
      if (installed.code !== 0)
        throw new OfficeError("VENV_INSTALL", "Failed to install pinned Office python packages", `Check network access to PyPI and retry; if it persists, delete ~/.cache/opencode-office/venv and retry. stderr: ${installed.stderr.slice(0, 400)}`)
    } else {
      throw new OfficeError(
        "PYTHON_MISSING",
        "Python 3.10+ required for Office tools — install python3 or uv",
        "macOS: `brew install uv`. Linux: `apt install python3-venv` or install uv. Then retry.",
      )
    }
    await writeFile(path.join(venvDir, ".fingerprint"), fingerprint())
    return python
  } finally {
    await release()
  }
}
