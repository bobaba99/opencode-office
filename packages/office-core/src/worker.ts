import path from "node:path"
import { OfficeError } from "./errors"
import { ensureVenv } from "./runtime"

const PYTHON_DIR = path.join(import.meta.dir, "python")

type WorkerEnvelope<T> = {
  ok: boolean
  data?: T
  error?: { code: string; message: string; hint: string }
}

export async function runWorker<T>(
  script: string,
  payload: unknown,
  opts?: { timeoutMs?: number; cacheDir?: string },
): Promise<T> {
  const python = await ensureVenv(opts?.cacheDir)
  const timeoutMs = opts?.timeoutMs ?? 60_000
  const proc = Bun.spawn([python, path.join(PYTHON_DIR, script)], {
    stdin: "pipe",
    stdout: "pipe",
    stderr: "pipe",
  })
  let timedOut = false
  const timer = setTimeout(() => {
    timedOut = true
    proc.kill()
  }, timeoutMs)
  proc.stdin.write(JSON.stringify(payload))
  await proc.stdin.end()
  const [stdout, stderr, code] = await Promise.all([
    new Response(proc.stdout).text(),
    new Response(proc.stderr).text(),
    proc.exited,
  ])
  clearTimeout(timer)
  if (timedOut)
    throw new OfficeError(
      "WORKER_TIMEOUT",
      `${script} exceeded ${timeoutMs}ms`,
      "Large documents can be slow; retry with a narrower target, or raise timeoutMs.",
    )
  if (code !== 0)
    throw new OfficeError("WORKER_CRASH", `${script} exited with code ${code}`, (stderr || "no stderr").slice(0, 2000))
  let parsed: WorkerEnvelope<T>
  try {
    parsed = JSON.parse(stdout)
  } catch {
    throw new OfficeError("WORKER_PROTOCOL", `${script} returned non-JSON output`, stdout.slice(0, 500))
  }
  if (!parsed.ok || parsed.error) {
    const err = parsed.error ?? { code: "WORKER_PROTOCOL", message: "worker returned ok=false without error", hint: "Bug in worker script." }
    throw new OfficeError(err.code, err.message, err.hint)
  }
  return parsed.data as T
}
