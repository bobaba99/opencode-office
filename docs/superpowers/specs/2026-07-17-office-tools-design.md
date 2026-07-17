# OpenCode Office Tools — Design

**Date:** 2026-07-17
**Status:** Approved pending user spec review
**Repo:** `~/development/opencode-office` (new); consumes stock OpenCode (`anomalyco/opencode` clone at `~/development/opencode` used as reference)

## Purpose

Give OpenCode agents full-lifecycle Microsoft Word (.docx) and PowerPoint (.pptx) capability: read and extract, create from scratch, and surgically edit existing documents without destroying formatting, styles, comments, or tracked changes.

## Requirements (decided in brainstorming)

1. **Full lifecycle** — read, create, and surgical edit for both docx and pptx.
2. **Installable add-on** — an npm plugin package for stock OpenCode. No fork divergence, no upstream PR required.
3. **Dependencies allowed** — Python (auto-provisioned venv with python-docx/python-pptx) and LibreOffice (`soffice`) for rendering. Both degrade with actionable errors when absent.
4. **Layered model support** — a small set of typed tools as the floor for ~7–30B local models, plus a Python escape hatch for frontier models. Aligns with the capacity-parity objective.

## Approach (chosen: B)

One installable plugin package with typed tools + script escape hatch + bundled skill. Rejected alternatives: (A) Anthropic-style skill pack only — fastest but frontier-only, no permission/attachment integration; (C) standalone MCP server — portable but loses OpenCode-native integration. The core/binding split below keeps C cheap to add later.

## Architecture

```
opencode-office/
├── packages/
│   ├── office-core/              # All document intelligence, no OpenCode imports
│   │   ├── src/
│   │   │   ├── docx/             # read / edit ops / create for Word
│   │   │   ├── pptx/             # read / edit ops / create for PowerPoint
│   │   │   ├── ids.ts            # stable element-ID scheme shared by read & edit
│   │   │   ├── render.ts         # soffice page/slide → PNG
│   │   │   ├── runtime.ts        # venv provisioning (uv → python3 -m venv fallback), soffice detection
│   │   │   └── python/           # Python worker scripts core shells out to
│   │   └── test/                 # golden-file fixtures, round-trip fidelity tests
│   └── opencode-plugin-office/   # Thin binding: tool registrations + bundled skill
│       ├── src/index.ts          # plugin entry (@opencode-ai/plugin)
│       └── skill/SKILL.md        # teaches the layering: typed tools first, python when needed
```

- **TypeScript orchestrates, Python executes.** office-core (TS, Bun-friendly) handles IDs, validation, diffing, atomicity. Every OOXML mutation runs in a short-lived Python worker (python-docx/python-pptx), JSON over stdin/stdout. No daemon.
- **The plugin binding is deliberately thin** — maps tool calls to office-core, wires permissions, converts rendered PNGs to tool attachments, registers the bundled skill via the v2 plugin skill hook. Future MCP or Claude Code bindings are additional thin layers over office-core.

## Tool surface

All tools are file-path based.

### `office_read(file, mode, target?)`
- `mode: "outline"` — skeleton: headings/sections (docx), slide list with titles (pptx); every element tagged with an ID.
- `mode: "content"` — readable text for a target range, IDs inline (`[p:12]`, `[s:4/sh:2]`).
- `mode: "full"` — adds formatting, styles, comments, tracked changes, speaker notes.
- Large documents paginate.

### `office_edit(file, operations[])`
- Each op: target **ID + anchor** (text the model believes is at the target, like Edit's `old_string`). IDs are positional and shift across edits; the anchor makes stale IDs safe. Mismatch → structured error containing the actual text at that ID; nothing written.
- **Atomic batch**: validate all → apply all → single save. Any failure rolls back everything.
- docx ops: `replace_text`, `insert_content` (markdown → paragraphs styled with the doc's own styles), `delete_element`, `set_style` (assign a named paragraph style that already exists in the doc), `set_table_cell` (set the text of the cell at row/col of a table ID).
- pptx ops: `set_shape_text`, `set_notes`, `insert_slide` (from an existing layout, inheriting theme), `duplicate_slide`, `delete_slide`, `reorder_slides`, `replace_image`.

### `office_create(file, spec)`
- docx: markdown in → styled document out; optional reference docx to clone styles from.
- pptx: slide-spec array (layout, title, bullets, notes) applied onto a template deck. Never generates theme XML from scratch.

### `office_render(file, range?)`
- soffice renders pages/slides to PNG, returned as tool attachments. Enables edit → render → inspect → fix loop; primary guardrail for small models.

### `office_python(code, files?)`
- Runs in the managed venv (python-docx, python-pptx, pillow preinstalled). `files` lists document paths the script intends to touch — used for permission prompting and the per-session backup, not sandboxing. Timeout-bounded, permission-gated like shell. Bundled SKILL.md directs models to prefer typed tools.

### Write safety
- Atomic writes (temp file + rename).
- First mutation of a file per session stashes a pristine backup in the plugin cache dir (office files are usually not in git; this is the undo story).

## Runtime provisioning

- **Python:** first use → prefer `uv`, fall back to `python3 -m venv`. Dedicated venv at `~/.cache/opencode-office/venv`, pinned dependency versions, fingerprint recorded to avoid re-resolution. No Python → structured error: "Python 3.10+ required for Office tools — install python3 or uv".
- **LibreOffice:** detected at first `office_render` (`soffice` on PATH + standard macOS/Linux app locations). Absence degrades rendering only; read/edit/create unaffected. Renders use an isolated LibreOffice profile dir (no conflict with a running desktop instance) and a hard timeout.

## Error handling

Every failure returns `{code, message, hint}`. Example: `ANCHOR_MISMATCH` includes the actual text found at the target ID so the model can correct without re-reading the file. Errors must contain the recovery path — this is the small-model floor expressed in the details.

## Permissions

- Reads: no prompt.
- `office_edit` / `office_create`: ask via OpenCode's permission system with the file path as pattern ("always allow for this file" works).
- `office_python`: same weight as shell.

## Testing

1. **Golden-file round-trips** (office-core): fixture corpus of real-world docx/pptx (styled corporate docs, tracked changes, embedded images, non-Latin text). Each edit op asserts (a) the edit landed and (b) nothing else changed — byte-level OOXML diff of untouched parts. Fidelity regression is the primary failure mode; it is tested directly.
2. **Render-based visual checks**: after create/edit, render and compare to reference PNGs via perceptual hash — catches "valid XML, broken layout" bugs invisible to XML diffs.
3. **Agent-level eval battery** (designed now, built later): task cards (e.g. "change Q3 to Q4 across this deck and update the title-slide date") run through OpenCode with a frontier model and a local model, scored on task success + file fidelity. Directly measures the small-model-floor claim; compatible with the existing repair-parity harness style.

## Out of scope (v1)

- Excel (.xlsx) — architecture leaves room (a future `xlsx/` module in office-core) but nothing in v1 depends on it.
- Authoring tracked changes / comments (reading them is in scope via `office_read` full mode).
- MCP server binding (kept cheap by the core/binding split; not built now).
- Rendering-dependent features beyond PNG previews (no PDF export in v1).

## Success criteria

- All five tools work end-to-end on the fixture corpus with zero fidelity regressions.
- A local ~30B model can complete representative edit tasks using only the typed tools.
- One-command install into stock OpenCode; first-run provisioning completes without manual steps on a machine with Python; absence of Python/LibreOffice produces the specified actionable errors.
