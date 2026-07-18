import { mkdir, readdir, readFile, writeFile } from "node:fs/promises"
import path from "node:path"
import { CARDS } from "./cards"

// `bun eval/report.ts` regenerates docs/BENCHMARK.md from eval/results/*.json.
// This file is a pure reducer over committed result JSON — it never re-runs the battery
// and never invents a number. If a value isn't in the JSON, it isn't in the report.

type CardResult = { success: boolean; fidelity: boolean; notes: string[]; seconds: number }
type ModelResults = { model: string; startedAt: string; cards: Record<string, CardResult> }

const RESULTS_DIR = path.join(import.meta.dir, "results")
const OUT_FILE = path.join(import.meta.dir, "..", "docs", "BENCHMARK.md")
const CARD_ORDER = CARDS.map((c) => c.id)
const FULL_BATTERY = CARD_ORDER.length // 8

// Models attempted but never produced a scored card: no results file, no data to lie with.
// Hardcoded because "unavailable" is not a fact that lives in eval/results/*.json — there is
// no JSON to auto-classify it from.
const UNAVAILABLE: Array<{ model: string; reason: string }> = [
  { model: "openai/gpt-5.5", reason: "provider credential expired — 401 on token refresh" },
  { model: "openai/gpt-5.4-mini", reason: "provider credential expired — 401 on token refresh" },
  { model: "kimi-k3", reason: "no such upstream model — only k2.x releases exist" },
  {
    model: "ollama-cloud (provider entry)",
    reason: "401 on every request; superseded by declaring cloud models under the working `ollama` provider instead",
  },
]

async function loadResults(): Promise<ModelResults[]> {
  const files = (await readdir(RESULTS_DIR)).filter((f) => f.endsWith(".json"))
  const all = await Promise.all(
    files.map(async (f) => JSON.parse(await readFile(path.join(RESULTS_DIR, f), "utf8")) as ModelResults),
  )
  return all.sort((a, b) => a.model.localeCompare(b.model))
}

// provider/id -> a short, honest class label. Not asserted by hand per model: derived from
// the model string itself so a new results file classifies itself correctly next run.
function classify(model: string): string {
  const slash = model.indexOf("/")
  const provider = slash === -1 ? model : model.slice(0, slash)
  const id = slash === -1 ? "" : model.slice(slash + 1)
  if (/cloud/i.test(id)) return provider === "ollama" ? "ollama-cloud" : `${provider}-cloud`
  if (provider === "opencode") return "opencode-hosted"
  if (provider === "ollama") {
    const sizeMatch = id.match(/(\d+(?:\.\d+)?b)/i)
    return sizeMatch ? `local (${sizeMatch[1].toUpperCase()})` : "local"
  }
  return `api (${provider})`
}

function median(nums: number[]): number {
  if (nums.length === 0) return NaN
  const sorted = [...nums].sort((a, b) => a - b)
  const mid = Math.floor(sorted.length / 2)
  return sorted.length % 2 === 0 ? ((sorted[mid - 1] as number) + (sorted[mid] as number)) / 2 : (sorted[mid] as number)
}

function fmtSeconds(n: number): string {
  return Number.isFinite(n) ? `${n.toFixed(2)}s` : "n/a"
}

function countBy(cards: Record<string, CardResult>, pred: (c: CardResult) => boolean): number {
  return Object.values(cards).filter(pred).length
}

// The two scores are orthogonal (see method note b), so a plain pass/fail matrix would hide
// the case that matters most for this battery: the agent completed the ask but destroyed
// something it wasn't supposed to touch. Four symbols instead of two.
function cellSymbol(c?: CardResult): string {
  if (!c) return "–" // not run
  if (c.success && c.fidelity) return "✓" // full pass
  if (c.success && !c.fidelity) return "⚠" // completed the ask, damaged something else
  if (!c.success && c.fidelity) return "✗" // missed the ask, nothing else touched
  return "✗✗" // missed the ask AND damaged something else
}

function noteFor(cards: Record<string, CardResult>): string {
  const entries = Object.entries(cards)
  const fidelityBreaches = entries.filter(([, c]) => c.success && !c.fidelity)
  const failed = entries.filter(([, c]) => !c.success)
  const parts: string[] = []
  if (fidelityBreaches.length > 0) {
    const [cardId, c] = fidelityBreaches[0] as [string, CardResult]
    const extra = fidelityBreaches.length > 1 ? ` (+${fidelityBreaches.length - 1} more)` : ""
    parts.push(`fidelity breach: ${cardId} — ${c.notes[0] ?? "see matrix"}${extra}`)
  }
  if (failed.length > 0) {
    const label = failed.length === 1 ? "task miss" : "task misses"
    parts.push(`${failed.length}/${FULL_BATTERY} ${label}: ${failed.map(([id]) => id).join(", ")}`)
  }
  return parts.length > 0 ? parts.join("; ") : "clean sweep — full success and fidelity"
}

function mainTableRow(r: ModelResults): string {
  const cards = r.cards
  const success = countBy(cards, (c) => c.success)
  const fidelity = countBy(cards, (c) => c.fidelity)
  const secs = Object.values(cards).map((c) => c.seconds)
  return `| ${r.model} | ${classify(r.model)} | ${success}/${FULL_BATTERY} | ${fidelity}/${FULL_BATTERY} | ${fmtSeconds(median(secs))} | ${noteFor(cards)} |`
}

function matrixRow(r: ModelResults): string {
  const cells = CARD_ORDER.map((id) => cellSymbol(r.cards[id]))
  return `| ${r.model} | ${cells.join(" | ")} |`
}

function partialRow(r: ModelResults): string {
  const cards = r.cards
  const success = countBy(cards, (c) => c.success)
  const fidelity = countBy(cards, (c) => c.fidelity)
  const ran = Object.keys(cards).length
  return `| ${r.model} | ${classify(r.model)} | ${ran}/${FULL_BATTERY} | ${success}/${ran} | ${fidelity}/${ran} | ${noteFor(cards)} |`
}

function renderMethodSection(complete: ModelResults[]): string {
  // Note (f) pulls its evidence straight from the committed JSON rather than asserting the
  // finding from prose — if the underlying result ever changes, this sentence goes stale in
  // an obviously wrong way instead of silently.
  const gptOss = complete.find((r) => r.model === "ollama/gpt-oss:120b-cloud")
  const gptOssReplace = gptOss?.cards["docx-replace"]
  const colorNote =
    gptOss && gptOssReplace && gptOssReplace.success && !gptOssReplace.fidelity
      ? `\`${gptOss.model}\` is the clearest example the battery has produced: it passed \`docx-replace\`'s text change (the sentence reads correctly) but the fidelity probe caught it destroying the paragraph's comment reference marker in the process (\`${gptOssReplace.notes[0]}\`). Task success alone would have scored this a clean pass — exactly the class of damage this battery exists to detect.`
      : `See the per-card matrix above for cases where a model passed the task-success check while the fidelity probe caught collateral damage — that gap is the reason the two scores are kept separate.`

  return `## Method

- **Scoring is programmatic office-core ground truth.** Every check reads the resulting
  \`.docx\`/\`.pptx\` (and, where needed, runs a Python-side probe such as \`docx_probe.py\` /
  \`pptx_probe.py\` for comment references, tracked changes, or image content hashes) and
  compares it against the expected structure. Nothing is scored from the agent's own prose —
  an agent claiming success proves nothing here.
- **Prompts are tool-agnostic.** Cards describe the file, the absolute path, and the concrete
  change; they never name \`office_read\`/\`office_edit\`/\`office_create\`/\`office_python\`.
  Tool selection is part of what the battery measures, so it can't be handed to the model.
- **\`success\` and \`fidelity\` are separate, orthogonal scores.** \`success\` = the asked-for
  change landed. \`fidelity\` = everything the prompt didn't ask about is still byte/value
  identical. For the \`docx-create\` card there is no pre-existing document to damage, so
  \`fidelity\` is always \`true\` by construction — including on a run that never produced the
  output file. That's intended: \`success\` alone captures the miss for a create card, and
  \`fidelity\` is reserved for collateral damage, which a from-scratch file can't cause.
- **Reproduction gotcha.** opencode treats \`provider.<id>.models\` in \`opencode.json\` as an
  **allowlist** for custom OpenAI-compatible providers. A model not declared there fails with
  \`ProviderModelNotFoundError\`, surfaced to the caller as an opaque "Unexpected server error"
  — if you add a model and it fails immediately with that message, check the allowlist before
  suspecting the provider.
- **Hardware.** MacBook Pro M4 Max, 128 GB unified memory. Every model in the roster below runs
  through a cloud provider (opencode-hosted, ollama-cloud, or a direct API), so this hardware
  bounds the harness (arena I/O, office-core scoring, process orchestration) — it does not
  bound model inference.
- **Variance caveat.** A single run per card per model. Cloud model outputs vary run to run;
  treat single-digit success/fidelity differences between models as noise, not signal.
  Local-model numbers (where present) additionally depend on quantization and would not
  reproduce on different hardware.
- **Refresh policy.** Re-run when the tool surface changes in a way that could invalidate a
  score (new office-core operations, changed fixtures, changed card prompts/checks) — not on a
  calendar. A stale-but-still-valid benchmark is preferable to a fresh one measuring something
  else.
- **Mid-battery re-measurement.** An ollama-cloud usage limit was hit partway through the run
  and lifted by a subscription upgrade. Every card that failed with a 300s timeout under the
  limit, plus one fast-fail that looked suspicious, was re-run with \`--force\` after the
  upgrade. All timeouts turned out to be quota stalls, not genuine model failures; the numbers
  below are the post-upgrade values. Any failure that survived the re-run is a genuine result.
- ${colorNote}
`
}

function render(all: ModelResults[]): string {
  const complete = all.filter((r) => Object.keys(r.cards).length === FULL_BATTERY)
  const partial = all.filter((r) => Object.keys(r.cards).length < FULL_BATTERY)
  const rankedComplete = [...complete].sort((a, b) => {
    const sa = countBy(a.cards, (c) => c.success)
    const sb = countBy(b.cards, (c) => c.success)
    if (sb !== sa) return sb - sa
    const fa = countBy(a.cards, (c) => c.fidelity)
    const fb = countBy(b.cards, (c) => c.fidelity)
    if (fb !== fa) return fb - fa
    return a.model.localeCompare(b.model)
  })

  const mainTable = [
    "| Model | Class | Task success | Fidelity | Median s/card | Notes |",
    "|---|---|---|---|---|---|",
    ...rankedComplete.map(mainTableRow),
  ].join("\n")

  const matrixHeader = `| Model | ${CARD_ORDER.join(" | ")} |`
  const matrixDivider = `|---|${CARD_ORDER.map(() => "---").join("|")}|`
  const matrixTable = [matrixHeader, matrixDivider, ...rankedComplete.map(matrixRow)].join("\n")

  const partialSection =
    partial.length > 0
      ? `## Partial / discontinued

Three local models were discontinued mid-run by author decision — the roster shifted from
local inference to cloud partway through the battery, and these were not re-run. Their partial
results (1-3 of 8 cards) are kept for the record but excluded from the ranked table above: a
1/8 or 3/8 sample isn't comparable to a completed 8/8 run.

| Model | Class | Cards run | Success | Fidelity | Notes |
|---|---|---|---|---|---|
${partial.map(partialRow).join("\n")}
`
      : ""

  const unavailableSection = `## Unavailable

Models the roster intended to include but that never produced a scored run:

${UNAVAILABLE.map((u) => `- **${u.model}** — ${u.reason}`).join("\n")}
`

  const generatedAt = new Date().toISOString()

  return `# Office Tools Benchmark

_Generated by \`bun eval/report.ts\` from \`eval/results/*.json\` on ${generatedAt}. Do not hand-edit — re-run the generator instead._

An 8-card agent-level eval battery, run through real OpenCode + the \`opencode-plugin-office\`
plugin, scored programmatically against office-core ground truth. Each card asks for one
concrete change to a \`.docx\`/\`.pptx\` fixture; scoring checks both that the change landed
(**Task success**) and that everything else in the file is untouched (**Fidelity**).

## Results

${mainTable}

Legend: ✓ full pass &nbsp; ⚠ completed the task but damaged something else (fidelity breach) &nbsp; ✗ missed the task, nothing else touched &nbsp; ✗✗ missed the task and damaged something else.

### Per-card matrix

${matrixTable}

${partialSection}
${unavailableSection}
${renderMethodSection(complete)}`
}

async function main(): Promise<void> {
  const all = await loadResults()
  const markdown = render(all)
  await mkdir(path.dirname(OUT_FILE), { recursive: true })
  await writeFile(OUT_FILE, markdown)
  console.log(`wrote ${OUT_FILE}`)
}

if (import.meta.main) {
  main().catch((e) => {
    console.error(e instanceof Error ? e.message : String(e))
    process.exit(1)
  })
}
