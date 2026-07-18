# opencode-office

An [opencode](https://opencode.ai) plugin that gives coding agents five tools for working with
Word (`.docx`) and PowerPoint (`.pptx`) files, plus the skill that teaches an agent how to use
them well:

- **`office_read`** — outline-first or full-content reads of a `.docx`/`.pptx`, so an agent can
  see structure before editing.
- **`office_edit`** — anchored, targeted edits (replace a paragraph's text, retitle a slide, fix
  one table cell) that touch only what's asked and leave everything else byte/value identical.
- **`office_create`** — build a new `.docx`/`.pptx` from scratch, or a new `.pptx` from an
  existing file used as a layout template.
- **`office_render`** — render pages/slides to images via LibreOffice, for visual review.
- **`office_python`** — an escape hatch into the underlying Python worker for operations the
  other four don't cover.
- **skill** (`packages/opencode-plugin-office/skill/SKILL.md`) — the workflow doc the agent
  reads: outline-first reads, anchored edits, and the full operations catalog for the tools
  above.

The tools are built on `@opencode-office/core`, which drives `python-docx` / `python-pptx` /
`pillow` / `pymupdf` in a managed Python venv, plus LibreOffice (`soffice`) for rendering.
`@opencode-office/core` ships as TypeScript source (its `exports` map points straight at
`.ts` files, no build step) — consumers must be able to execute TypeScript directly, as
opencode and Bun do. Plain Node.js consumers are not supported for now.

## Install

**From npm** (once published — this repo currently ships dry-run-verified packaging only, see
[`docs/superpowers/plans`](docs/superpowers/plans) for status):

```sh
bun add opencode-plugin-office
```

**From a local checkout / workspace** (the only path available today): install/link the package
wherever `opencode.json` resolves it from — opencode loads plugins from `node_modules` or a
workspace path.

Either way, reference the package by name in `opencode.json`:

```json
{
  "plugin": ["opencode-plugin-office"]
}
```

The skill cannot be auto-registered by this plugin (see below), so copy it into an opencode
skills directory yourself:

```sh
cp packages/opencode-plugin-office/skill/SKILL.md ~/.config/opencode/skill/office-tools/SKILL.md
```

Full install details, including why skill auto-registration isn't possible with the installed
`@opencode-ai/plugin` version, live in
[`packages/opencode-plugin-office/README.md`](packages/opencode-plugin-office/README.md).

**Platform:** developed and tested on macOS/Linux. Windows is untested — the Python venv
provisioning and LibreOffice discovery paths assume a POSIX shell.

## Dependencies

- **Python** — auto-provisioned. The core package creates and caches a venv (pinned
  `python-docx`, `python-pptx`, `pillow`, `pymupdf` versions) under
  `~/.cache/opencode-office` the first time it's needed; there's nothing to install by hand.
- **LibreOffice** — optional, only required for `office_render`. The plugin looks for `soffice`
  on `PATH` and in the usual macOS install locations; if it isn't found, every tool except
  `office_render` still works normally.

## Benchmark

An 8-card agent-level eval battery — real OpenCode + this plugin, scored programmatically
against office-core ground truth (never against the agent's own claims) — run across a roster
of cloud models. Full methodology, the per-card matrix, partial/discontinued runs, and models
that were unavailable live in [`docs/BENCHMARK.md`](docs/BENCHMARK.md).

| Model | Class | Task success | Fidelity | Median s/card | Notes |
|---|---|---|---|---|---|
| ollama/qwen3.5:397b-cloud | ollama-cloud | 8/8 | 8/8 | 22.26s | clean sweep — full success and fidelity |
| opencode/big-pickle | opencode-hosted | 8/8 | 8/8 | 32.64s | clean sweep — full success and fidelity |
| zai/glm-4.5-flash | api (zai) | 7/8 | 7/8 | 58.61s | 1/8 task miss: pptx-insert |
| ollama/gpt-oss:120b-cloud | ollama-cloud | 6/8 | 7/8 | 21.23s | fidelity breach: docx-replace — comment_refs expected 1, got 0; 2/8 task misses: pptx-image, pptx-create |
| ollama/minimax-m3:cloud | ollama-cloud | 5/8 | 8/8 | 19.18s | 3/8 task misses: docx-create, pptx-insert, pptx-image |
| opencode/nemotron-3-ultra-free | opencode-hosted | 5/8 | 7/8 | 28.54s | 3/8 task misses: docx-replace, docx-create, pptx-create |

Hardware: MacBook Pro M4 Max, 128 GB unified memory (all models run through a cloud provider,
so this bounds the harness, not model inference). See `docs/BENCHMARK.md` for the variance
caveat and refresh policy before treating close scores as a ranking.

This table is copied by hand from `docs/BENCHMARK.md`, which is regenerated with
`bun eval/report.ts`; re-copy it here after a regeneration if the numbers change.
