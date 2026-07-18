import { expect, test } from "bun:test"
import { existsSync } from "node:fs"
import { rm, mkdir, utimes } from "node:fs/promises"
import path from "node:path"
import type { OfficeError } from "../src/errors"
import { PINNED, ensureVenv, venvIsCurrent, acquireLock } from "../src/runtime"

test("pins the required packages", () => {
  expect(Object.keys(PINNED).sort()).toEqual(["pillow", "pymupdf", "python-docx", "python-pptx"])
})

test("venvIsCurrent is false for a directory with no venv", async () => {
  expect(await venvIsCurrent("/tmp/definitely-not-a-venv")).toBe(false)
})

// Integration: provisions the real shared venv (slow on first run, cached after).
test("ensureVenv provisions python with pinned deps", async () => {
  const python = await ensureVenv()
  expect(existsSync(python)).toBe(true)
  const venvDir = path.dirname(path.dirname(python))
  expect(await venvIsCurrent(venvDir)).toBe(true)
  const proc = Bun.spawn([python, "-c", "import docx, pptx, PIL; print('ok')"], { stdout: "pipe" })
  expect(await proc.exited).toBe(0)
  expect(await new Response(proc.stdout).text()).toContain("ok")
}, 180_000)

test("acquireLock excludes a second contender until released", async () => {
  const dir = "/tmp/oc-office-lock-test"
  await rm(dir + ".lock", { recursive: true, force: true })
  const release = await acquireLock(dir)
  let second = false
  const contender = acquireLock(dir, { timeoutMs: 5_000 }).then(async (rel) => {
    second = true
    await rel()
  })
  await new Promise((resolve) => setTimeout(resolve, 400))
  expect(second).toBe(false)
  await release()
  await contender
  expect(second).toBe(true)
})

test("stale lock is stolen", async () => {
  const dir = "/tmp/oc-office-stale-test"
  await rm(dir + ".lock", { recursive: true, force: true })
  await mkdir(dir + ".lock", { recursive: true })
  const old = new Date(Date.now() - 600_000)
  await utimes(dir + ".lock", old, old)
  const release = await acquireLock(dir, { staleMs: 300_000, timeoutMs: 5_000 })
  await release()
})

test("acquireLock with a missing parent resolves quickly instead of spinning", async () => {
  const dir = "/tmp/oc-office-parent-test/nested/venv"
  await rm("/tmp/oc-office-parent-test", { recursive: true, force: true })
  const start = Date.now()
  const release = await acquireLock(dir, { timeoutMs: 5_000 })
  expect(Date.now() - start).toBeLessThan(2_000)
  await release()
})

test("acquireLock honors a caller staleMs larger than the default: a 400s-old lock is not stolen when staleMs is 600s", async () => {
  // Regresses render.ts's bug: acquireLock was called with a derived timeoutMs but a
  // hardcoded default staleMs (300_000), so a long-running conversion's still-ACTIVE lock
  // could be stolen out from under it once it crossed 300s. Here the lock is 400s old —
  // older than the 300s default — but the caller passes staleMs: 600_000, so the lock must
  // still be treated as active and the short timeoutMs must cause a LOCK_TIMEOUT instead of
  // a steal.
  const dir = "/tmp/oc-office-stale-honors-caller-test"
  await rm(dir + ".lock", { recursive: true, force: true })
  await mkdir(dir + ".lock", { recursive: true })
  const old = new Date(Date.now() - 400_000)
  await utimes(dir + ".lock", old, old)
  try {
    await acquireLock(dir, { staleMs: 600_000, timeoutMs: 1_500 })
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("LOCK_TIMEOUT")
  }
  await rm(dir + ".lock", { recursive: true, force: true })
})
