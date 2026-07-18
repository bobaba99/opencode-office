# Office Tools Plan 5: Release Prep Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make both packages publish-ready (NOT published — publishing needs the user's npm login and explicit go-ahead): fail-loud `duplicate_slide` guard, render serialization, per-op field validation, skill auto-registration, exact pins, LICENSE/metadata, dry-run-verified packaging.

**Tech Stack:** unchanged. Branch: `feat/release-prep`.

**Verified API facts:** `@opencode-ai/plugin@1.18.3` typings include v2 skill hooks (`SkillHooks`/`SkillDraft` with a `source()` mechanism — the P3 final reviewer confirmed their presence in the installed `.d.ts`); python-pptx 1.0.2 has `shapes.add_chart` (charts produce `c:chart` parts related via `r:id`). `acquireLock` exists in office-core runtime (parent-safe since 9d61fba).

## Global Constraints

- Plans 1–4 Global Constraints still bind (OfficeError recovery-first hints, atomic saves, no OpenCode imports in office-core, bun test from root, commit per task).
- New error code: `UNSUPPORTED_SLIDE_CONTENT` — hint must name what was found (e.g. "chart") and state the recovery (edit the original, or delete/replace the unsupported element first).
- NOTHING is published in this plan: `bun publish --dry-run` / `npm pack --dry-run` only. Publishing is a separate, user-triggered action.
- Version 0.1.0 for both packages. Exact pin `"@opencode-ai/plugin": "1.18.3"` (no caret — we depend on verified internals).
- LICENSE: MIT, copyright "Gavin (Zihao Geng)" — flagged for user veto in the final summary.

---

### Task 1: duplicate_slide fail-loud guard + plugin per-op field validation

**Files:**
- Modify: `packages/office-core/src/python/pptx_edit.py` (guard in `copy_slide`)
- Modify: `packages/office-core/src/python/gen_fixtures.py` (chart slide fixture)
- Modify: `packages/opencode-plugin-office/src/tools.ts` (required-field map)
- Test: `packages/office-core/test/pptx-edit.test.ts`, `packages/opencode-plugin-office/test/tools.test.ts` (extend both)

**Interfaces:**
- `copy_slide` guard: after appending copied shapes with the rid_map rewrites, iterate every element in the new slide's `_spTree`; for each attribute in `{qn("r:id"), qn("r:embed"), qn("r:link")}` present on any descendant, verify the value exists in `new_slide.part.rels`; any miss → raise `WorkerError("UNSUPPORTED_SLIDE_CONTENT", f"Slide contains content duplicate_slide cannot copy safely ({tag name of the offending element})", "Charts, SmartArt, and embedded objects are not yet supported by duplicate_slide. Edit the original slide, or delete/replace the unsupported element first.")` — raised BEFORE returning, so the batch aborts and nothing is written (atomicity holds).
- Fixture: `make_edit_pptx` gains a 4th slide with a chart:

```python
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    chart_data = CategoryChartData()
    chart_data.categories = ["A", "B"]
    chart_data.add_series("S1", (1.0, 2.0))
    s4 = prs.slides.add_slide(prs.slide_layouts[6])
    s4.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1), Inches(4), Inches(3), chart_data)
```

  NOTE: existing pptx tests assume the deck has exactly 3 slides. Audit `pptx-edit.test.ts` and `eval/cards.ts`/`score.ts` for slide-count/outline assumptions against edit-deck.pptx and update ONLY assertions that count slides or enumerate outlines (e.g. `insert_slide` test's expected title list gains a trailing `""`; duplicate test's expected count 4→5; eval pptx-image card targets s:2 — unchanged, chart slide is s:3 now so the eval duplicate lands at... AUDIT eval/cards.ts pptx-image: it duplicates s:2 (picture slide) and expects the copy at s:3 — with the chart as new s:3, the duplicate STILL inserts directly after s:2 (position 3) pushing the chart to s:4; probes address by id so assertions hold. Verify rather than assume; adjust where reality disagrees and document each adjustment.)
- Plugin per-op required-field map in `validateOperations` (BEFORE ctx.ask, after the existing id checks): for each op name, verify the presence and primitive type of every required field (docx: replace_text{anchor:str,text:str}, insert_content{markdown:str}, delete_element{anchor:str}, set_style{anchor:str,style:str}, set_table_cell{row:int,col:int,text:str}; pptx: set_shape_text{anchor:str,text:str}, set_notes{text:str}, insert_slide{layout:str}, move_slide{index:int}, replace_image{image:str}) — missing/mistyped → `OfficeError("BAD_ARGS", \`op ${op.op} is missing or has invalid field ${field}\`, "See the operations catalog in the office-tools skill for each op's required fields.")`.

- [ ] **Step 1: failing tests** — pptx-edit: `duplicate_slide` on the chart slide → `UNSUPPORTED_SLIDE_CONTENT`, file unchanged (re-read before/after equality); duplicate of the PICTURE slide still succeeds (regression). tools.test: `office_edit` with `{op:"replace_text", target:"p:0", anchor:"x"}` (missing text) → BAD_ARGS with askCalls===0.
- [ ] **Step 2: RED** (`rm -rf packages/office-core/test/.fixtures` first — fixture changed)
- [ ] **Step 3: implement per Interfaces; run the full suite and fix ONLY assertions invalidated by the 4th slide, documenting each in the report**
- [ ] **Step 4: full `bun test` + `bun run typecheck` green**
- [ ] **Step 5: Commit** — `feat: fail-loud duplicate_slide guard; plugin per-op field validation`

---

### Task 2: serialize renders

**Files:**
- Modify: `packages/office-core/src/render.ts`
- Test: `packages/office-core/test/render.test.ts` (extend)

**Interfaces:** wrap the soffice conversion (profile-dir usage) in `acquireLock(profileDir, { timeoutMs: 180_000 })` from runtime.ts — concurrent `renderOffice` calls queue instead of failing. The lock covers only the soffice invocation, not the pymupdf worker. Update the exit-0-no-output failure hint (profile contention wording can soften to mention the queue).

- [ ] **Step 1: failing test** — `test.skipIf(!HAS_SOFFICE)("concurrent renders both succeed", ...)`: `Promise.all([renderOffice(deck), renderOffice(report, {pages:[1]})])` with distinct outDirs → both resolve with pages, no RENDER_FAILED (600_000 timeout).
- [ ] **Step 2: RED (currently the loser deterministically fails)**
- [ ] **Step 3: implement**
- [ ] **Step 4: `bun test render` (4), full suite, typecheck**
- [ ] **Step 5: Commit** — `fix: serialize soffice renders via provisioning lock`

---

### Task 3: skill auto-registration

**Files:**
- Modify: `packages/opencode-plugin-office/src/index.ts`
- Test: `packages/opencode-plugin-office/test/plugin.test.ts` (new)

**Interfaces:** read the INSTALLED `@opencode-ai/plugin@1.18.3` typings for the skill hook shape (the P3 review confirmed `SkillHooks`/`SkillDraft.source()` exist in the v2 surface — find the exact registration path a v1-shaped `Plugin`'s `Hooks` supports; if the installed `Hooks` type genuinely has NO skill field, implement the closest supported mechanism and, if none exists, keep manual copy, update README to say so definitively, and report the finding — do NOT fake it). On success: the plugin serves `skill/SKILL.md`'s content (read once at load via Bun file read relative to `import.meta.dir`). Test what is testable without a live host: the hook/export exists and the draft/source resolves to content containing "office_read" and the op catalog heading.

- [ ] **Step 1: failing test**
- [ ] **Step 2: RED**
- [ ] **Step 3: implement (or document impossibility with evidence)**
- [ ] **Step 4: full suite + typecheck; ALSO run one live verification: `/Users/zihaogeng/.opencode/bin/opencode run "Is an office-tools skill available to you? Answer with only yes or no."` from a temp dir with NO local skills — record the answer in the report (informative, not gating)**
- [ ] **Step 5: Commit** — `feat: plugin auto-registers the office-tools skill` (or `docs:` variant if impossible)

---

### Task 4: packaging + metadata + dry-run

**Files:**
- Modify: `packages/office-core/package.json`, `packages/opencode-plugin-office/package.json`
- Create: `LICENSE` (root, MIT, "Copyright (c) 2026 Gavin (Zihao Geng)"), copied into each package dir (npm includes per-package LICENSE)
- Modify: `README.md`, `packages/opencode-plugin-office/README.md`

**Interfaces:**
- Both packages: `version: "0.1.0"`, `license: "MIT"`, `repository`/`description`/`keywords` fields, `files` allowlist (src/, skill/ for plugin, README, LICENSE — EXCLUDING test/), remove `"private": true` from the plugin.
- Plugin deps: `"@opencode-ai/plugin": "1.18.3"` exact; `"@opencode-office/core": "workspace:*"` — verify `bun publish --dry-run` rewrites workspace protocol to `0.1.0`; if it does not, hardcode `"0.1.0"`.
- core package: confirm its Python workers ship (src/python/** in files) and `exports` maps stay TS-source-based (opencode/bun consumers execute TS directly — document that node-only consumers are unsupported for now in the README).
- Dry-run both: `cd packages/office-core && bun publish --dry-run` (fallback `npm pack --dry-run`) and same for the plugin; the report must list the tarball file lists and flag anything unexpected (missing python/, test/ leaking, etc.).
- READMEs: npm install path (`bun add opencode-plugin-office` + opencode.json plugin entry by package name) alongside the existing local-path instructions; platform note ("macOS/Linux; Windows untested"); remove the manual-skill-copy instruction IF Task 3 succeeded (else keep).

- [ ] **Step 1: metadata edits**
- [ ] **Step 2: dry-runs; capture file lists in the report; fix `files` until clean**
- [ ] **Step 3: README updates consistent with Task 3's outcome**
- [ ] **Step 4: full `bun test` + typecheck (nothing behavioral changed)**
- [ ] **Step 5: Commit** — `chore: release metadata, LICENSE, dry-run-verified packaging for 0.1.0`
