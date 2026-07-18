import { existsSync } from "node:fs"
import { mkdir, readFile, rename, writeFile } from "node:fs/promises"
import path from "node:path"
import { CARDS } from "./cards"
import { makeArena } from "./arena"

export const OPENCODE_BIN = "/Users/zihaogeng/.opencode/bin/opencode"
const WALL_CLOCK_CAP_MS = 300_000
const RESULTS_DIR = path.join(import.meta.dir, "results")

type CardResult = { success: boolean; fidelity: boolean; notes: string[]; seconds: number }
type ModelResults = { model: string; startedAt: string; cards: Record<string, CardResult> }

type Options = { models: string[]; cardIds: string[]; parallelLocal: boolean; force: boolean }

function parseArgs(argv: string[]): Options {
  let models: string[] | undefined
  let cardIds: string[] | undefined
  let parallelLocal = false
  let force = false

  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i]
    if (arg === "--models") models = (argv[++i] ?? "").split(",").map((s) => s.trim()).filter(Boolean)
    else if (arg === "--cards") cardIds = (argv[++i] ?? "").split(",").map((s) => s.trim()).filter(Boolean)
    else if (arg === "--parallel-local") parallelLocal = true
    else if (arg === "--force") force = true
  }

  if (!models || models.length === 0) throw new Error("--models <csv> is required")
  const knownIds = new Set(CARDS.map((c) => c.id))
  const resolvedCardIds = cardIds ?? CARDS.map((c) => c.id)
  for (const id of resolvedCardIds) {
    if (!knownIds.has(id)) throw new Error(`Unknown card id: ${id}`)
  }

  return { models, cardIds: resolvedCardIds, parallelLocal, force }
}

// `/` and `:` both map to distinct characters (not the same "-") so e.g. "a/b" and "a:b"
// can't collide on disk.
function sanitizeModel(model: string): string {
  return model.replace(/\//g, "_").replace(/:/g, "-")
}

function resultsPath(model: string): string {
  return path.join(RESULTS_DIR, `${sanitizeModel(model)}.json`)
}

async function loadResults(model: string): Promise<ModelResults> {
  const file = resultsPath(model)
  if (existsSync(file)) {
    try {
      return JSON.parse(await readFile(file, "utf8")) as ModelResults
    } catch (e) {
      // A prior run may have been killed mid-write (timeout SIGKILL, crash). Don't let a
      // corrupt results file permanently block resume — start fresh and let the run overwrite it.
      console.warn(`[${model}] results file ${file} is corrupt (${e instanceof Error ? e.message : String(e)}); starting fresh`)
    }
  }
  return { model, startedAt: new Date().toISOString(), cards: {} }
}

async function saveResults(model: string, results: ModelResults): Promise<void> {
  await mkdir(RESULTS_DIR, { recursive: true })
  const file = resultsPath(model)
  const tmp = `${file}.tmp-${process.pid}`
  await writeFile(tmp, JSON.stringify(results, null, 2))
  await rename(tmp, file)
}

async function runCard(model: string, cardId: string): Promise<CardResult> {
  const card = CARDS.find((c) => c.id === cardId)
  if (!card) throw new Error(`Unknown card id: ${cardId}`)

  const arena = await makeArena(cardId)
  const prompt = card.prompt(arena)
  const start = Date.now()

  const proc = Bun.spawn([OPENCODE_BIN, "run", "-m", model, prompt], {
    cwd: arena,
    stdout: "pipe",
    stderr: "pipe",
  })

  let timedOut = false
  const timer = setTimeout(() => {
    timedOut = true
    proc.kill()
    // opencode may have spawned office_python/worker grandchildren; best-effort reap them
    // too so a timeout doesn't leave orphans running across the rest of the battery.
    if (proc.pid) Bun.spawn(["pkill", "-9", "-P", String(proc.pid)], { stdout: "ignore", stderr: "ignore" }).exited.catch(() => {})
    setTimeout(() => proc.kill(9), 5_000).unref?.()
  }, WALL_CLOCK_CAP_MS)

  const [stdout, stderr, exitCode] = await Promise.all([new Response(proc.stdout).text(), new Response(proc.stderr).text(), proc.exited])
  clearTimeout(timer)
  const seconds = (Date.now() - start) / 1000

  if (timedOut) {
    return { success: false, fidelity: false, notes: ["timeout"], seconds }
  }

  // A check() that throws (malformed output the check didn't anticipate) must still record
  // a scored failure, not abort the whole model/batch run out from under sibling cards.
  let score: { success: boolean; fidelity: boolean; notes: string[] }
  try {
    score = await card.check(arena)
  } catch (e) {
    score = { success: false, fidelity: false, notes: [`check() threw: ${e instanceof Error ? e.message : String(e)}`] }
  }
  const notes = exitCode !== 0 ? [...score.notes, `opencode exited with code ${exitCode}`, stderr.slice(0, 500)] : score.notes
  if (process.env.EVAL_VERBOSE) {
    console.log(`--- ${model} / ${cardId} stdout ---\n${stdout.slice(0, 4000)}`)
  }
  return { success: score.success, fidelity: score.fidelity, notes, seconds }
}

async function runModel(model: string, cardIds: string[], force: boolean): Promise<void> {
  const results = await loadResults(model)
  for (const cardId of cardIds) {
    if (!force && results.cards[cardId]) {
      console.log(`[${model}] skipping ${cardId} (already in results; use --force to re-run)`)
      continue
    }
    console.log(`[${model}] running ${cardId}...`)
    const result = await runCard(model, cardId)
    results.cards[cardId] = result
    await saveResults(model, results)
    console.log(`[${model}] ${cardId}: success=${result.success} fidelity=${result.fidelity} (${result.seconds}s)`)
  }
}

async function main(): Promise<void> {
  const opts = parseArgs(process.argv.slice(2))
  if (opts.parallelLocal) {
    await Promise.all(opts.models.map((model) => runModel(model, opts.cardIds, opts.force)))
  } else {
    for (const model of opts.models) {
      await runModel(model, opts.cardIds, opts.force)
    }
  }
}

if (import.meta.main) {
  main().catch((e) => {
    console.error(e instanceof Error ? e.message : String(e))
    process.exit(1)
  })
}
