# opencode-plugin-office

An [opencode](https://opencode.ai) plugin that gives the agent five tools for working with
Word (`.docx`) and PowerPoint (`.pptx`) files: `office_read`, `office_edit`, `office_create`,
`office_render`, and `office_python`. It's built on `opencode-office-core`, which drives
`python-docx` / `python-pptx` / `pillow` / `pymupdf` in a managed venv, plus LibreOffice
(`soffice`) for rendering.

See `skill/SKILL.md` for the workflow the agent follows (outline-first reads, anchored edits,
the full operations catalog) and the underlying tool contracts.

## Install

**From npm** (once published):

```sh
bun add opencode-plugin-office
```

**From a local checkout / workspace** (the only path available today): install/link the package
wherever `opencode.json` resolves it from — opencode loads plugins from `node_modules` or a
workspace path.

Either way, reference it by name in `opencode.json`:

```json
{
  "plugin": ["opencode-plugin-office"]
}
```

**Platform:** developed and tested on macOS/Linux. Windows is untested — the Python venv
provisioning and LibreOffice discovery paths in `opencode-office-core` assume a POSIX shell.

## Skill registration (manual — not possible to automate on `@opencode-ai/plugin@1.18.3`)

`skill/SKILL.md` is not auto-registered by this plugin, and cannot be with the currently
pinned `@opencode-ai/plugin@1.18.3`. Copy it into an opencode skills directory yourself:

```sh
mkdir -p ~/.config/opencode/skills/office-tools && cp skill/SKILL.md ~/.config/opencode/skills/office-tools/SKILL.md
```

**Why this can't be automated:** this plugin exports the v1 `Plugin` shape
(`(input, options?) => Promise<Hooks>`), because that's the only shape that can register the
`office_read`/`office_edit`/`office_create`/`office_render`/`office_python` tools (via
`Hooks.tool`) — tool registration has no equivalent in the v2 surface. The installed v1 `Hooks`
interface (`@opencode-ai/plugin`'s root export, `dist/index.d.ts`) has no skill-registration
field of any kind. Skill registration (`SkillHooks` / `SkillDraft.source()`) exists only behind
`PluginContext` (`dist/v2/promise/context.d.ts`), the parameter to an entirely different plugin
shape — `{ id, setup(context) }`, created via `define()` — exported from the separate subpaths
`@opencode-ai/plugin/v2/promise` / `@opencode-ai/plugin/v2/effect`. Nothing on the v1 side
bridges to a `PluginContext`, and a single plugin module's default export can only be one shape
or the other, not both. A future `@opencode-ai/plugin` release that adds skill registration to
the v1 `Hooks` interface would remove this constraint; `test/plugin.test.ts` carries a
type-level assertion (checked by `bun run typecheck`, not by `bun test` alone) that fails to
compile if that happens, as a prompt to revisit this. It does not detect the other route out
(a tool-registration mechanism landing on the v2 surface) — that would need to be checked
by hand against `dist/v2/promise/context.d.ts`.
