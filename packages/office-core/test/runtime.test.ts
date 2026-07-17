import { expect, test } from "bun:test"
import { existsSync } from "node:fs"
import path from "node:path"
import { PINNED, ensureVenv, venvIsCurrent } from "../src/runtime"

test("pins the three required packages", () => {
  expect(Object.keys(PINNED).sort()).toEqual(["pillow", "python-docx", "python-pptx"])
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
