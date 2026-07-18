import { existsSync } from "node:fs"
import path from "node:path"
import { ensureVenv } from "../src/runtime"

export const FIXTURE_DIR = path.join(import.meta.dir, ".fixtures")

export async function ensureFixtures(): Promise<void> {
  if (existsSync(path.join(FIXTURE_DIR, "swap.png"))) return
  const python = await ensureVenv()
  const script = path.join(import.meta.dir, "..", "src", "python", "gen_fixtures.py")
  const proc = Bun.spawn([python, script, FIXTURE_DIR], { stdout: "pipe", stderr: "pipe" })
  if ((await proc.exited) !== 0) {
    throw new Error(`fixture generation failed: ${await new Response(proc.stderr).text()}`)
  }
}
