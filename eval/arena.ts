import { randomUUID } from "node:crypto"
import { mkdir, cp, writeFile } from "node:fs/promises"
import path from "node:path"
import { ensureFixtures, FIXTURE_DIR } from "../packages/office-core/test/fixtures"
import { CARDS } from "./cards"

const ARENA_ROOT = "/tmp/office-eval"

// Pilot-discovered (Task 1, Step 3 — see docs/superpowers/sdd/task-1-report.md): opencode
// 1.18.3 honors named tool-permission keys in an arena-local opencode.json for plugin
// `ctx.ask` calls under headless `opencode run`. This is escalation-list attempt 1 from the
// plan; it worked, so the wider `"*": "allow"` and OPENCODE_PERMISSION fallbacks were never
// needed.
function arenaConfig(): Record<string, unknown> {
  return {
    $schema: "https://opencode.ai/config.json",
    permission: {
      office_edit: "allow",
      office_create: "allow",
      office_python: "allow",
    },
  }
}

// Fresh arena directory for one card run: copies just the fixtures that card needs and
// writes an arena-local opencode.json granting the office tools' permissions headlessly.
export async function makeArena(cardId: string): Promise<string> {
  const card = CARDS.find((c) => c.id === cardId)
  if (!card) throw new Error(`Unknown card id: ${cardId}`)

  await ensureFixtures()

  // `--parallel-local` fires multiple models' first makeArena() call for the same cardId
  // within the same millisecond (ensureFixtures resolves instantly once fixtures exist), so
  // Date.now() alone can collide and let two opencode processes edit the same arena files
  // concurrently. The random suffix guarantees a unique directory regardless of timing.
  const dir = path.join(ARENA_ROOT, `${cardId}-${Date.now()}-${randomUUID().slice(0, 8)}`)
  await mkdir(dir, { recursive: true })

  for (const file of card.files) {
    await cp(path.join(FIXTURE_DIR, file), path.join(dir, file))
  }

  await writeFile(path.join(dir, "opencode.json"), JSON.stringify(arenaConfig(), null, 2))

  return dir
}
