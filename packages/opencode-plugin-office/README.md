# opencode-plugin-office

An [opencode](https://opencode.ai) plugin that gives the agent five tools for working with
Word (`.docx`) and PowerPoint (`.pptx`) files: `office_read`, `office_edit`, `office_create`,
`office_render`, and `office_python`. It's built on `@opencode-office/core`, which drives
`python-docx` / `python-pptx` / `pillow` / `pymupdf` in a managed venv, plus LibreOffice
(`soffice`) for rendering.

See `skill/SKILL.md` for the workflow the agent follows (outline-first reads, anchored edits,
the full operations catalog) and the underlying tool contracts.

## Install

Add the plugin to your `opencode.json`:

```json
{
  "plugin": ["opencode-plugin-office"]
}
```

opencode loads plugins from `node_modules` (or a workspace path during local development), so
the package must be installed/linked wherever `opencode.json` resolves it from.

## Skill registration (manual for now)

`skill/SKILL.md` is not auto-registered by this plugin. Until the v2 skill hook lands (see
below), copy it into an opencode skills directory yourself, e.g.:

```sh
cp skill/SKILL.md ~/.config/opencode/skill/office-tools/SKILL.md
```

**Tracked future work:** automatic skill registration via opencode's v2 skill hook, so
installing this plugin is enough on its own and this manual copy step goes away.
