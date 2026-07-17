### Task 1: Monorepo scaffold

**Files:**
- Create: `package.json`, `tsconfig.json`, `.gitignore`
- Create: `packages/office-core/package.json`
- Test: `packages/office-core/test/smoke.test.ts`

**Interfaces:**
- Consumes: nothing.
- Produces: a workspace where `bun test` discovers `packages/*/test/*.test.ts`.

- [ ] **Step 1: Write the failing test**

`packages/office-core/test/smoke.test.ts`:

```ts
import { expect, test } from "bun:test"

test("workspace resolves office-core", async () => {
  const pkg = await Bun.file(new URL("../package.json", import.meta.url)).json()
  expect(pkg.name).toBe("@opencode-office/core")
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun test` (from `/Users/zihaogeng/development/opencode-office`)
Expected: FAIL — `packages/office-core/package.json` does not exist.

- [ ] **Step 3: Write the scaffold**

`package.json` (repo root):

```json
{
  "name": "opencode-office",
  "private": true,
  "workspaces": ["packages/*"]
}
```

`tsconfig.json` (repo root):

```json
{
  "compilerOptions": {
    "target": "ESNext",
    "module": "ESNext",
    "moduleResolution": "bundler",
    "strict": true,
    "types": ["bun-types"],
    "noEmit": true
  },
  "include": ["packages"]
}
```

`.gitignore` (repo root):

```
node_modules/
packages/office-core/test/.fixtures/
.cache/
```

`packages/office-core/package.json`:

```json
{
  "name": "@opencode-office/core",
  "version": "0.0.1",
  "type": "module",
  "exports": { ".": "./src/index.ts" }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `bun install && bun test`
Expected: PASS (1 test).

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "chore: monorepo scaffold with office-core package"
```

---

