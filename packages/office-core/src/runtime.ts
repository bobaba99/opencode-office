import { existsSync } from "node:fs"
import { mkdir, readFile, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { OfficeError } from "./errors"

export const PINNED: Record<string, string> = {
  "python-docx": "1.2.0",
  "python-pptx": "1.0.2",
  pillow: "11.3.0",
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

async function checkVersion(python: string): Promise<void> {
  const check = await run([python, "-c", "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)"])
  if (check.code !== 0)
    throw new OfficeError(
      "PYTHON_TOO_OLD",
      "Python 3.10+ required for Office tools — install python3 or uv",
      "Your python3 is older than 3.10. Install a newer python3, or install uv (https://docs.astral.sh/uv/) which manages its own.",
    )
}

export async function ensureVenv(cacheDir = defaultCacheDir()): Promise<string> {
  const venvDir = path.join(cacheDir, "venv")
  const python = venvPython(venvDir)
  if (await venvIsCurrent(venvDir)) return python
  await mkdir(cacheDir, { recursive: true })
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
}
