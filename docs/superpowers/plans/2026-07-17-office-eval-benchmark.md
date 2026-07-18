# Office Tools Plan 4: Eval Battery + Published Benchmark Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A runnable agent-level eval battery (8 task cards through real OpenCode + the office plugin, scored programmatically on task success AND fidelity), executed across a 6-model roster, published as docs/BENCHMARK.md + a README table.

**Architecture:** `eval/` at repo root (Bun TS, not a workspace package — it drives the product, it isn't the product). Each card: `prepare(arena)` copies fixtures into a fresh arena dir → `opencode run -m <model>` with the card prompt (cwd = arena) → `check(arena)` scores via office-core reads/probes (never via the agent's own claims). Results accumulate in `eval/results/<model>.json`; a generator renders BENCHMARK.md.

**Tech Stack:** Bun TS; opencode binary at `/Users/zihaogeng/.opencode/bin/opencode` (1.18.3 — NEVER the brew 1.15.12); office-core for scoring; plugin already registered globally in ~/.config/opencode/opencode.json.

**Spec:** docs/superpowers/specs/2026-07-17-office-tools-design.md (Published model benchmark section). Plan 4 of 4.

## Global Constraints

- Scoring is programmatic office-core ground truth only — never parse the agent's prose for success.
- Success and fidelity are SEPARATE scores per card: success = the asked-for change landed; fidelity = untouched content byte/value-identical per the card's fidelity checks.
- Per-run wall-clock cap 300s (kill the opencode process, score the card as failed with note "timeout").
- Cards must not require office_render (shared render profile contends under parallel runs — ledger-known); if an agent renders spontaneously the structured error is acceptable noise.
- Local-model batteries follow the standing parallel rule: check combined model footprint vs 80% of RAM (`sysctl hw.memsize`) before parallelizing across models; runs for a single model are always sequential.
- Roster (6): `openai/gpt-5.5` (frontier), `ollama-cloud/glm-5.1` (default-class cloud), `ollama-cloud/gpt-oss:120b` (open cloud), `ollama/qwen3-coder:30b` (local 30B — the spec's success-criterion tier), `ollama/qwen2.5-coder:14b` (local mid), `ollama/llama3.1:8b` (local floor). If a model errors on auth/availability, record `unavailable` in results and continue — never fake a score.
- Benchmark table discloses hardware (MacBook Pro M4 Max, 128 GB unified) and quantization (ollama defaults) and carries the variance caveat + refresh policy from the spec.

---

### Task 1: Harness + cards + scoring + headless-permission pilot

**Files:**
- Create: `eval/cards.ts`, `eval/score.ts`, `eval/run.ts`, `eval/arena.ts`
- Test: `eval/score.test.ts` (bun test discovers it via root tsconfig include — verify; if not, add the path)

**Interfaces:**
- `arena.ts`: `makeArena(cardId: string): Promise<string>` — fresh dir under `/tmp/office-eval/<cardId>-<Date.now()>`; copies the office-core fixtures it needs (ensureFixtures() from `../packages/office-core/test/fixtures` first) and writes an arena `opencode.json` that (a) grants permissions headlessly and (b) pins nothing else. **Pilot obligation:** discover the exact permission config 1.18.3 honors for plugin `ctx.ask` in `opencode run` — try in order: `{"permission": {"office_edit": "allow", "office_create": "allow", "office_python": "allow"}}`; then `{"permission": {"*": "allow"}}`; then env `OPENCODE_PERMISSION`. Record which works in the report; hardcode the winner.
- `cards.ts`: `type Card = { id: string; files: string[]; prompt: (arena: string) => string; check: (arena: string) => Promise<CardScore> }`, `type CardScore = { success: boolean; fidelity: boolean; notes: string[] }`, `export const CARDS: Card[]` — the 8 cards below. Prompts name absolute paths and the concrete change; they do NOT name tools (tool selection is what we're measuring).
  1. **docx-replace** (edit-report.docx): "In <arena>/edit-report.docx, change the word 'strong' to 'robust' in the growth sentence. Change nothing else." — success: paragraph text becomes "Growth was robust this quarter overall."; fidelity: docx_probe comment_refs === 1, tracked paragraph's tracked_insertions unchanged, table cells unchanged, p:0 style Heading 1.
  2. **docx-insert** (report.docx): "In <arena>/report.docx, add a new section right after the 'Regional Breakdown' heading: a level-2 heading 'Next Steps' followed by bullet points 'Hire' and 'Ship'." — success: outline order contains Regional Breakdown → Next Steps (Heading 2), bullets styled List Bullet with texts Hire, Ship; fidelity: pre-existing paragraphs and table unchanged.
  3. **docx-table** (report.docx): "In <arena>/report.docx, update the EMEA revenue figure in the table to $5.1M." — success: cell(1,1) === "$5.1M"; fidelity: cell(0,0..1)+cell(1,0) unchanged, all paragraphs unchanged.
  4. **docx-create**: "Create <arena>/summary.docx: a heading 'Weekly Summary', a paragraph 'All systems nominal.', then bullet points 'Uptime 99.9%' and 'Zero incidents'." — success: 4 elements w/ those texts, heading styled Heading*, bullets List Bullet; fidelity: n/a → true.
  5. **pptx-retitle** (deck.pptx): "In <arena>/deck.pptx, retitle the second slide to 'Q3 Highlights' and set its speaker notes to 'Two minutes max.'" — success: s:1 title + notes; fidelity: s:0 title/subtitle unchanged.
  6. **pptx-insert** (deck.pptx): "In <arena>/deck.pptx, insert a new slide titled 'Agenda' with bullet points 'Numbers' and 'Risks' directly after the title slide, using the 'Title and Content' layout." — success: outline [Q3 Review, Agenda, Highlights], s:1 layout "Title and Content", body contains Numbers/Risks; fidelity: first/last slides' shape texts unchanged.
  7. **pptx-image** (edit-deck.pptx + swap.png): "In <arena>/edit-deck.pptx, duplicate the third slide (the one with the picture) and on the DUPLICATE swap the picture for <arena>/swap.png. Leave the original slide untouched." — success: 4 slides, pptx_probe s:3/sh:0 sha === sha(swap.png); fidelity: s:2/sh:0 sha unchanged.
  8. **pptx-create** (template edit-deck.pptx): "Create <arena>/plan.pptx using <arena>/edit-deck.pptx as the template, with exactly two slides: 'Kickoff' on the 'Title Slide' layout, and 'Timeline' on 'Title and Content' with bullets 'Design', 'Build', 'Ship'." — success: 2 slides, titles+layouts+bullets; fidelity: template file's own sha unchanged.
- `score.ts`: implements every check with office-core imports (`readDocx`, `readPptx`, `runWorker` for docx_probe/pptx_probe, node:crypto sha256). Each check failure appends a human-readable note.
- `run.ts`: CLI `bun eval/run.ts --models <csv> [--cards <csv>] [--parallel-local]` — per model×card: makeArena → spawn `[OPENCODE_BIN, "run", "-m", model, prompt]` cwd arena, 300s kill; then check(arena); accumulate `eval/results/<model-sanitized>.json` `{model, startedAt, cards: {id: {success, fidelity, notes, seconds}}}`; idempotent resume (skip cards already in the file unless `--force`). `OPENCODE_BIN = "/Users/zihaogeng/.opencode/bin/opencode"`.
- `score.test.ts`: unit-tests the checks WITHOUT agents — for 2 cards (docx-replace, pptx-image), programmatically apply the correct edit via office-core `editDocx`/`editPptx` in a fresh arena and assert `check` returns success:true/fidelity:true; then apply a deliberately wrong/destructive edit and assert the relevant flag flips.

- [ ] **Step 1: write score.test.ts, watch it fail**
- [ ] **Step 2: implement arena/cards/score, make tests pass**
- [ ] **Step 3: PILOT (mandatory before Task 2, per the standing preflight rule): run `bun eval/run.ts --models ollama-cloud/glm-5.1 --cards docx-replace` — one real agent run of a costly edit card. It must complete headlessly (permission config proven) and produce a scored results file. Record wall-clock in the report; if the permission escalation list is exhausted without success, STOP and report BLOCKED with the observed behavior.**
- [ ] **Step 4: full `bun test` (81+) and `bun run typecheck`**
- [ ] **Step 5: Commit** — `feat(eval): office battery harness, 8 cards, programmatic scoring, headless pilot`

---

### Task 2: Run the battery across the roster

- [ ] **Step 1: Preflight arithmetic (standing rule): from the pilot's wall-clock, estimate full-battery time = pilot_seconds × 8 cards × 6 models (halve for parallel-local if memory-fit passes: qwen3-coder:30b 17.3GB + qwen2.5-coder:14b 8.4GB + llama3.1:8b 4.6GB = ~30GB + 15% KV overhead « 102GB budget → parallel OK: `OLLAMA_MAX_LOADED_MODELS=3 OLLAMA_NUM_PARALLEL=3`). Write the estimate to the report BEFORE launching; if projected > 90 min, run cloud models first and report partial.**
- [ ] **Step 2: cloud models: `bun eval/run.ts --models openai/gpt-5.5,ollama-cloud/glm-5.1,ollama-cloud/gpt-oss:120b` (sequential).**
- [ ] **Step 3: local models: `OLLAMA_MAX_LOADED_MODELS=3 OLLAMA_NUM_PARALLEL=3 bun eval/run.ts --models ollama/qwen3-coder:30b,ollama/qwen2.5-coder:14b,ollama/llama3.1:8b --parallel-local` (three per-model sequential card streams in parallel).**
- [ ] **Step 4: sanity-pass over results: any model with ≥6 `unavailable`/timeout cards → re-run once; still failing → record as-is with note.**
- [ ] **Step 5: Commit results** — `feat(eval): battery results for 6-model roster`

---

### Task 3: BENCHMARK.md + README table

**Files:**
- Create: `eval/report.ts` (results/*.json → markdown), `docs/BENCHMARK.md`, `README.md` (repo root — currently absent)

**Interfaces:**
- `eval/report.ts`: `bun eval/report.ts` regenerates docs/BENCHMARK.md from eval/results/*.json. Table columns: Model | Class (frontier/cloud/local+size) | Task success (n/8) | Fidelity (n/8) | Median s/card | Notes. Below the table: per-card breakdown matrix (✓/✗/–), method section (battery described, scoring = programmatic office-core ground truth, prompts tool-agnostic), hardware disclosure (MacBook Pro M4 Max, 128 GB unified memory; ollama default quantizations; cloud models via their providers), variance caveat (local results vary with hardware/quantization), and the refresh policy (re-run when the tool surface changes in score-invalidating ways, not on a calendar).
- Repo root `README.md`: what opencode-office is (five tools + skill for OpenCode), install (plugin entry in opencode.json + skill copy, per packages/opencode-plugin-office/README.md), the benchmark table transcluded (copy the table section, link to docs/BENCHMARK.md for the full matrix), dependency notes (Python auto-provisioned; LibreOffice optional for render).

- [ ] **Step 1: implement report.ts; generate docs/BENCHMARK.md from the real results**
- [ ] **Step 2: write README.md with the real table**
- [ ] **Step 3: `bun test` + typecheck still green**
- [ ] **Step 4: Commit** — `docs: published model benchmark and README`
