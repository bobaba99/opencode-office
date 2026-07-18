import { expect, test } from "bun:test"
import type { Hooks, PluginInput } from "@opencode-ai/plugin"
import { OfficePlugin } from "../src/index"

const SKILL_PATH = new URL("../skill/SKILL.md", import.meta.url)

// --- Finding (Task 3, skill auto-registration): -----------------------------
//
// @opencode-ai/plugin@1.18.3's v1 `Plugin` type -- the shape `OfficePlugin`
// implements -- is `(input: PluginInput, options?) => Promise<Hooks>`. The
// full v1 `Hooks` interface (node_modules/.bun/@opencode-ai+plugin@1.18.3/
// node_modules/@opencode-ai/plugin/dist/index.d.ts) has NO skill-registration
// field. Skill registration (`SkillHooks` / `SkillDraft.source()`) lives
// exclusively behind `PluginContext` (.../dist/v2/promise/context.d.ts:20,
// `readonly skill: SkillHooks & Reload`), which is the parameter to a wholly
// different plugin shape -- `interface Plugin { id; setup(context) }` via
// `define()` -- exported only from the separate subpaths
// `@opencode-ai/plugin/v2/promise` and `@opencode-ai/plugin/v2/effect`, per
// that package's package.json `exports` map. Nothing on v1 `PluginInput` or
// `Hooks` bridges to a `PluginContext`.
//
// v2 also has no tool-registration surface (no `tool.d.ts` under `dist/v2/`),
// so this plugin cannot migrate its default export to the v2 `{id, setup}`
// shape either -- doing so would forfeit `Hooks.tool`, the only mechanism
// that registers office_read/office_edit/office_create/office_render/
// office_python, which is this plugin's entire purpose.
//
// Conclusion: no supported (or "closest available") registration path exists
// for auto-registering skill/SKILL.md from this v1-shaped plugin entry.
// Manual copy remains the only mechanism; see the READMEs.
//
// The type-level assertion below fails to compile if a future
// @opencode-ai/plugin version adds a `skill` field to v1 `Hooks` -- that
// would mean this finding no longer holds and auto-registration should be
// revisited. It's checked by `bun run typecheck`, not by `bun test` alone --
// bun test does not type-check, so a broken assertion here would run green
// under `bun test` and only be caught by the typecheck step.
type NoV1SkillHook = "skill" extends keyof Hooks ? never : true
const assertNoV1SkillHook: NoV1SkillHook = true
void assertNoV1SkillHook

function makeInput(): PluginInput {
  return {} as PluginInput
}

test("OfficePlugin returns v1 Hooks with a tool field registering all five tools", async () => {
  const hooks = await OfficePlugin(makeInput())
  expect(hooks.tool).toBeDefined()
  expect(Object.keys(hooks.tool ?? {}).sort()).toEqual(
    ["office_create", "office_edit", "office_python", "office_read", "office_render"].sort(),
  )
})

test("OfficePlugin's returned Hooks has no skill field -- no v1 registration path exists", async () => {
  const hooks = (await OfficePlugin(makeInput())) as Record<string, unknown>
  expect(hooks.skill).toBeUndefined()
})

test("skill/SKILL.md (served manually, not auto-registered) documents office_read", async () => {
  const content = await Bun.file(SKILL_PATH).text()
  expect(content).toContain("office_read")
})

test("skill/SKILL.md has the operations catalog heading", async () => {
  const content = await Bun.file(SKILL_PATH).text()
  expect(content).toContain("## Operations catalog")
})
