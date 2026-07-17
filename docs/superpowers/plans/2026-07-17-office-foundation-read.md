# Office Tools Plan 1: Foundation + Read Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Stand up the `office-core` package with Python runtime provisioning, the TS↔Python worker protocol, the element-ID scheme, and working `readDocx`/`readPptx` (outline + content modes) against generated fixtures.

**Architecture:** TypeScript (Bun) orchestrates; every OOXML operation runs in a short-lived Python worker (python-docx/python-pptx) speaking JSON over stdin/stdout. A dedicated venv is auto-provisioned at `~/.cache/opencode-office/venv` (uv preferred, `python3 -m venv` fallback) with pinned versions and a fingerprint file to skip re-resolution.

**Tech Stack:** Bun + TypeScript (`bun test`), Python 3.10+, python-docx 1.2.0, python-pptx 1.0.2, pillow 11.3.0, uv (optional).

**Spec:** `docs/superpowers/specs/2026-07-17-office-tools-design.md`. This is Plan 1 of 4 — deferred to later plans: `office_edit` + full-mode read (Plan 2), `office_create`/`office_render`/plugin binding/SKILL.md/read-output pagination-truncation (Plan 3, where model-facing output limits live), eval battery + published benchmark (Plan 4).

## Global Constraints

- Pinned Python deps, exact: `python-docx==1.2.0`, `python-pptx==1.0.2`, `pillow==11.3.0`.
- Every error surfaced to callers is an `OfficeError` carrying `{code, message, hint}`; hints must state the recovery path.
- Python-missing error message, verbatim: `Python 3.10+ required for Office tools — install python3 or uv`.
- Element IDs: `p:<n>` / `tbl:<n>` (docx body-child index), `s:<n>` (slide index), `s:<n>/sh:<m>` (shape index within slide). Indices are positional and refreshed on every read.
- No OpenCode imports anywhere in `office-core`.
- Run all tests from repo root with `bun test`; commit after every task.

---

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

### Task 2: Errors and element IDs

**Files:**
- Create: `packages/office-core/src/errors.ts`
- Create: `packages/office-core/src/ids.ts`
- Create: `packages/office-core/src/index.ts`
- Test: `packages/office-core/test/ids.test.ts`

**Interfaces:**
- Consumes: nothing.
- Produces: `class OfficeError extends Error { code: string; hint: string }` (constructor `(code, message, hint)`); `toToolError(err: unknown): {code, message, hint}`; `type ElementRef`; `formatId(ref: ElementRef): string`; `parseId(id: string): ElementRef` (throws `OfficeError` code `BAD_ID`).

- [ ] **Step 1: Write the failing test**

`packages/office-core/test/ids.test.ts`:

```ts
import { expect, test } from "bun:test"
import { formatId, parseId } from "../src/ids"
import { OfficeError } from "../src/errors"

test("round-trips every ID form", () => {
  for (const id of ["p:12", "tbl:3", "s:4", "s:4/sh:2"]) {
    expect(formatId(parseId(id))).toBe(id)
  }
})

test("parses shape IDs into slide and shape indices", () => {
  expect(parseId("s:4/sh:2")).toEqual({ kind: "shape", slide: 4, shape: 2 })
})

test("rejects malformed IDs with BAD_ID", () => {
  try {
    parseId("slide-4")
    expect.unreachable()
  } catch (e) {
    expect(e).toBeInstanceOf(OfficeError)
    expect((e as OfficeError).code).toBe("BAD_ID")
    expect((e as OfficeError).hint).toContain("p:12")
  }
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun test ids`
Expected: FAIL — cannot resolve `../src/ids`.

- [ ] **Step 3: Write minimal implementation**

`packages/office-core/src/errors.ts`:

```ts
export class OfficeError extends Error {
  constructor(
    readonly code: string,
    message: string,
    readonly hint: string,
  ) {
    super(message)
    this.name = "OfficeError"
  }
}

export function toToolError(err: unknown): { code: string; message: string; hint: string } {
  if (err instanceof OfficeError) return { code: err.code, message: err.message, hint: err.hint }
  const message = err instanceof Error ? err.message : String(err)
  return {
    code: "INTERNAL",
    message,
    hint: "Likely a bug in opencode-office. Re-run once to confirm, then report it.",
  }
}
```

`packages/office-core/src/ids.ts`:

```ts
import { OfficeError } from "./errors"

export type ElementRef =
  | { kind: "paragraph"; index: number }
  | { kind: "table"; index: number }
  | { kind: "slide"; index: number }
  | { kind: "shape"; slide: number; shape: number }

export function formatId(ref: ElementRef): string {
  switch (ref.kind) {
    case "paragraph":
      return `p:${ref.index}`
    case "table":
      return `tbl:${ref.index}`
    case "slide":
      return `s:${ref.index}`
    case "shape":
      return `s:${ref.slide}/sh:${ref.shape}`
  }
}

export function parseId(id: string): ElementRef {
  const shape = id.match(/^s:(\d+)\/sh:(\d+)$/)
  if (shape) return { kind: "shape", slide: Number(shape[1]), shape: Number(shape[2]) }
  const simple = id.match(/^(p|tbl|s):(\d+)$/)
  if (simple) {
    const index = Number(simple[2])
    if (simple[1] === "p") return { kind: "paragraph", index }
    if (simple[1] === "tbl") return { kind: "table", index }
    return { kind: "slide", index }
  }
  throw new OfficeError(
    "BAD_ID",
    `Unrecognized element ID: ${id}`,
    "Valid forms: p:12, tbl:3, s:4, s:4/sh:2 — IDs come from office_read output.",
  )
}
```

`packages/office-core/src/index.ts`:

```ts
export * from "./errors"
export * from "./ids"
```

- [ ] **Step 4: Run test to verify it passes**

Run: `bun test ids`
Expected: PASS (3 tests).

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: OfficeError shape and element ID scheme"
```

---

### Task 3: Python venv provisioning

**Files:**
- Create: `packages/office-core/src/runtime.ts`
- Modify: `packages/office-core/src/index.ts` (add `export * from "./runtime"`)
- Test: `packages/office-core/test/runtime.test.ts`

**Interfaces:**
- Consumes: `OfficeError` from Task 2.
- Produces: `PINNED: Record<string, string>`; `defaultCacheDir(): string`; `venvPython(venvDir: string): string`; `venvIsCurrent(venvDir: string): Promise<boolean>`; `ensureVenv(cacheDir?: string): Promise<string>` → absolute path of the venv's python. Error codes: `PYTHON_MISSING`, `PYTHON_TOO_OLD`, `VENV_CREATE`, `VENV_INSTALL`.

- [ ] **Step 1: Write the failing test**

`packages/office-core/test/runtime.test.ts`:

```ts
import { expect, test } from "bun:test"
import { existsSync } from "node:fs"
import path from "node:path"
import { PINNED, ensureVenv, venvIsCurrent } from "../src/runtime"

test("pins the three required packages", () => {
  expect(Object.keys(PINNED).sort()).toEqual(["pillow", "python-docx", "python-pptx"])
})

test("venvIsCurrent is false for a directory with no venv", async () => {
  expect(await venvIsCurrent("/tmp/definitely-not-a-venv")).toBe(false)
})

// Integration: provisions the real shared venv (slow on first run, cached after).
test("ensureVenv provisions python with pinned deps", async () => {
  const python = await ensureVenv()
  expect(existsSync(python)).toBe(true)
  const venvDir = path.dirname(path.dirname(python))
  expect(await venvIsCurrent(venvDir)).toBe(true)
  const proc = Bun.spawn([python, "-c", "import docx, pptx, PIL; print('ok')"], { stdout: "pipe" })
  expect(await proc.exited).toBe(0)
  expect(await new Response(proc.stdout).text()).toContain("ok")
}, 180_000)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun test runtime`
Expected: FAIL — cannot resolve `../src/runtime`.

- [ ] **Step 3: Write minimal implementation**

`packages/office-core/src/runtime.ts`:

```ts
import { existsSync } from "node:fs"
import { mkdir, readFile, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { OfficeError } from "./errors"

export const PINNED: Record<string, string> = {
  "python-docx": "1.2.0",
  "python-pptx": "1.0.2",
  pillow: "11.3.0",
}

export function defaultCacheDir(): string {
  return path.join(os.homedir(), ".cache", "opencode-office")
}

function fingerprint(): string {
  return Object.entries(PINNED)
    .map(([name, version]) => `${name}==${version}`)
    .sort()
    .join("\n")
}

export function venvPython(venvDir: string): string {
  return process.platform === "win32"
    ? path.join(venvDir, "Scripts", "python.exe")
    : path.join(venvDir, "bin", "python")
}

export async function venvIsCurrent(venvDir: string): Promise<boolean> {
  const marker = path.join(venvDir, ".fingerprint")
  if (!existsSync(venvPython(venvDir)) || !existsSync(marker)) return false
  return (await readFile(marker, "utf8")) === fingerprint()
}

async function run(cmd: string[]): Promise<{ code: number; stderr: string }> {
  const proc = Bun.spawn(cmd, { stdout: "pipe", stderr: "pipe" })
  const [code, stderr] = await Promise.all([proc.exited, new Response(proc.stderr).text()])
  return { code, stderr }
}

async function which(bin: string): Promise<boolean> {
  const probe = process.platform === "win32" ? ["where", bin] : ["which", bin]
  return (await run(probe)).code === 0
}

async function checkVersion(python: string): Promise<void> {
  const check = await run([python, "-c", "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)"])
  if (check.code !== 0)
    throw new OfficeError(
      "PYTHON_TOO_OLD",
      "Python 3.10+ required for Office tools — install python3 or uv",
      "Your python3 is older than 3.10. Install a newer python3, or install uv (https://docs.astral.sh/uv/) which manages its own.",
    )
}

export async function ensureVenv(cacheDir = defaultCacheDir()): Promise<string> {
  const venvDir = path.join(cacheDir, "venv")
  const python = venvPython(venvDir)
  if (await venvIsCurrent(venvDir)) return python
  await mkdir(cacheDir, { recursive: true })
  const pkgs = Object.entries(PINNED).map(([name, version]) => `${name}==${version}`)
  if (await which("uv")) {
    const created = await run(["uv", "venv", venvDir])
    if (created.code !== 0)
      throw new OfficeError("VENV_CREATE", "uv could not create the Office tools venv", created.stderr.slice(0, 500))
    await checkVersion(python)
    const installed = await run(["uv", "pip", "install", "--python", python, ...pkgs])
    if (installed.code !== 0)
      throw new OfficeError("VENV_INSTALL", "Failed to install pinned Office python packages", installed.stderr.slice(0, 500))
  } else if (await which("python3")) {
    const created = await run(["python3", "-m", "venv", venvDir])
    if (created.code !== 0)
      throw new OfficeError("VENV_CREATE", "python3 -m venv failed", created.stderr.slice(0, 500))
    await checkVersion(python)
    const installed = await run([python, "-m", "pip", "install", "--quiet", ...pkgs])
    if (installed.code !== 0)
      throw new OfficeError("VENV_INSTALL", "Failed to install pinned Office python packages", installed.stderr.slice(0, 500))
  } else {
    throw new OfficeError(
      "PYTHON_MISSING",
      "Python 3.10+ required for Office tools — install python3 or uv",
      "macOS: `brew install uv`. Linux: `apt install python3-venv` or install uv. Then retry.",
    )
  }
  await writeFile(path.join(venvDir, ".fingerprint"), fingerprint())
  return python
}
```

Add to `packages/office-core/src/index.ts`:

```ts
export * from "./runtime"
```

- [ ] **Step 4: Run test to verify it passes**

Run: `bun test runtime`
Expected: PASS (3 tests; the integration test takes 10–60s on first run, <1s after).

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: python venv provisioning with uv/python3 fallback and fingerprint cache"
```

---

### Task 4: Worker protocol

**Files:**
- Create: `packages/office-core/src/worker.ts`
- Create: `packages/office-core/src/python/_worker.py`
- Create: `packages/office-core/src/python/echo.py`
- Modify: `packages/office-core/src/index.ts` (add `export * from "./worker"`)
- Test: `packages/office-core/test/worker.test.ts`

**Interfaces:**
- Consumes: `ensureVenv` (Task 3), `OfficeError` (Task 2).
- Produces: `runWorker<T>(script: string, payload: unknown, opts?: { timeoutMs?: number; cacheDir?: string }): Promise<T>`. Python side: `_worker.run(main)` wrapper and `_worker.WorkerError(code, message, hint)`; stdout protocol `{"ok": true, "data": ...}` | `{"ok": false, "error": {code, message, hint}}`. Error codes: `WORKER_TIMEOUT`, `WORKER_CRASH`, `WORKER_PROTOCOL`, plus whatever the worker raises.

- [ ] **Step 1: Write the failing test**

`packages/office-core/test/worker.test.ts`:

```ts
import { expect, test } from "bun:test"
import { OfficeError } from "../src/errors"
import { runWorker } from "../src/worker"

test("echoes a payload through the venv python", async () => {
  const out = await runWorker<{ echo: { n: number } }>("echo.py", { n: 42 })
  expect(out.echo.n).toBe(42)
}, 180_000)

test("propagates WorkerError as OfficeError with code and hint", async () => {
  try {
    await runWorker("echo.py", { boom: true })
    expect.unreachable()
  } catch (e) {
    expect(e).toBeInstanceOf(OfficeError)
    expect((e as OfficeError).code).toBe("BOOM")
    expect((e as OfficeError).hint).toBe("test hint")
  }
}, 30_000)

test("kills a hung worker and reports WORKER_TIMEOUT", async () => {
  try {
    await runWorker("echo.py", { sleep_ms: 5_000 }, { timeoutMs: 500 })
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("WORKER_TIMEOUT")
  }
}, 30_000)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun test worker`
Expected: FAIL — cannot resolve `../src/worker`.

- [ ] **Step 3: Write minimal implementation**

`packages/office-core/src/python/_worker.py`:

```python
import json
import sys
import traceback


class WorkerError(Exception):
    def __init__(self, code, message, hint):
        super().__init__(message)
        self.code = code
        self.message = message
        self.hint = hint


def run(main):
    try:
        payload = json.load(sys.stdin)
        print(json.dumps({"ok": True, "data": main(payload)}))
    except WorkerError as e:
        print(json.dumps({"ok": False, "error": {"code": e.code, "message": e.message, "hint": e.hint}}))
    except Exception:
        print(
            json.dumps(
                {
                    "ok": False,
                    "error": {
                        "code": "PYTHON_EXCEPTION",
                        "message": traceback.format_exc(limit=5),
                        "hint": "Unexpected worker failure; likely a bug in opencode-office.",
                    },
                }
            )
        )
```

`packages/office-core/src/python/echo.py`:

```python
import time

from _worker import run, WorkerError


def main(payload):
    if payload.get("boom"):
        raise WorkerError("BOOM", "requested failure", "test hint")
    if payload.get("sleep_ms"):
        time.sleep(payload["sleep_ms"] / 1000)
    return {"echo": payload}


run(main)
```

`packages/office-core/src/worker.ts`:

```ts
import path from "node:path"
import { OfficeError } from "./errors"
import { ensureVenv } from "./runtime"

const PYTHON_DIR = path.join(import.meta.dir, "python")

type WorkerEnvelope<T> = {
  ok: boolean
  data?: T
  error?: { code: string; message: string; hint: string }
}

export async function runWorker<T>(
  script: string,
  payload: unknown,
  opts?: { timeoutMs?: number; cacheDir?: string },
): Promise<T> {
  const python = await ensureVenv(opts?.cacheDir)
  const timeoutMs = opts?.timeoutMs ?? 60_000
  const proc = Bun.spawn([python, path.join(PYTHON_DIR, script)], {
    stdin: "pipe",
    stdout: "pipe",
    stderr: "pipe",
  })
  let timedOut = false
  const timer = setTimeout(() => {
    timedOut = true
    proc.kill()
  }, timeoutMs)
  proc.stdin.write(JSON.stringify(payload))
  await proc.stdin.end()
  const [stdout, stderr, code] = await Promise.all([
    new Response(proc.stdout).text(),
    new Response(proc.stderr).text(),
    proc.exited,
  ])
  clearTimeout(timer)
  if (timedOut)
    throw new OfficeError(
      "WORKER_TIMEOUT",
      `${script} exceeded ${timeoutMs}ms`,
      "Large documents can be slow; retry with a narrower target, or raise timeoutMs.",
    )
  if (code !== 0)
    throw new OfficeError("WORKER_CRASH", `${script} exited with code ${code}`, (stderr || "no stderr").slice(0, 2000))
  let parsed: WorkerEnvelope<T>
  try {
    parsed = JSON.parse(stdout)
  } catch {
    throw new OfficeError("WORKER_PROTOCOL", `${script} returned non-JSON output`, stdout.slice(0, 500))
  }
  if (!parsed.ok || parsed.error) {
    const err = parsed.error ?? { code: "WORKER_PROTOCOL", message: "worker returned ok=false without error", hint: "Bug in worker script." }
    throw new OfficeError(err.code, err.message, err.hint)
  }
  return parsed.data as T
}
```

Add to `packages/office-core/src/index.ts`:

```ts
export * from "./worker"
```

- [ ] **Step 4: Run test to verify it passes**

Run: `bun test worker`
Expected: PASS (3 tests).

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: JSON stdin/stdout worker protocol with timeout and error propagation"
```

---

### Task 5: Fixture generation

**Files:**
- Create: `packages/office-core/src/python/gen_fixtures.py`
- Create: `packages/office-core/test/fixtures.ts`
- Test: `packages/office-core/test/fixtures.test.ts`

**Interfaces:**
- Consumes: `ensureVenv` (Task 3).
- Produces: `ensureFixtures(): Promise<void>` and `FIXTURE_DIR: string` (in `test/fixtures.ts`, test-only helper). Generates `report.docx` (Heading 1 "Quarterly Report", body paragraphs, Heading 2, a 2x2 table, closing paragraph) and `deck.pptx` (title slide "Q3 Review", bullets slide "Highlights" with speaker notes) into `test/.fixtures/`.

- [ ] **Step 1: Write the failing test**

`packages/office-core/test/fixtures.test.ts`:

```ts
import { expect, test } from "bun:test"
import { existsSync } from "node:fs"
import path from "node:path"
import { FIXTURE_DIR, ensureFixtures } from "./fixtures"

test("generates docx and pptx fixtures", async () => {
  await ensureFixtures()
  expect(existsSync(path.join(FIXTURE_DIR, "report.docx"))).toBe(true)
  expect(existsSync(path.join(FIXTURE_DIR, "deck.pptx"))).toBe(true)
}, 180_000)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun test fixtures`
Expected: FAIL — cannot resolve `./fixtures`.

- [ ] **Step 3: Write minimal implementation**

`packages/office-core/src/python/gen_fixtures.py`:

```python
import os
import sys

from docx import Document
from pptx import Presentation


def make_docx(path):
    doc = Document()
    doc.add_heading("Quarterly Report", level=1)
    doc.add_paragraph("Q3 revenue grew 12% year over year.")
    doc.add_heading("Regional Breakdown", level=2)
    doc.add_paragraph("EMEA led growth for the third consecutive quarter.")
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Region"
    table.cell(0, 1).text = "Revenue"
    table.cell(1, 0).text = "EMEA"
    table.cell(1, 1).text = "$4.2M"
    doc.add_paragraph("Prepared by the finance team.")
    doc.save(path)


def make_pptx(path):
    prs = Presentation()
    s1 = prs.slides.add_slide(prs.slide_layouts[0])
    s1.shapes.title.text = "Q3 Review"
    s1.placeholders[1].text = "Finance Team"
    s2 = prs.slides.add_slide(prs.slide_layouts[1])
    s2.shapes.title.text = "Highlights"
    s2.placeholders[1].text = "Revenue up 12%\nEMEA leads growth"
    s2.notes_slide.notes_text_frame.text = "Pause here for questions."
    prs.save(path)


if __name__ == "__main__":
    out = sys.argv[1]
    os.makedirs(out, exist_ok=True)
    make_docx(os.path.join(out, "report.docx"))
    make_pptx(os.path.join(out, "deck.pptx"))
    print("ok")
```

`packages/office-core/test/fixtures.ts`:

```ts
import { existsSync } from "node:fs"
import path from "node:path"
import { ensureVenv } from "../src/runtime"

export const FIXTURE_DIR = path.join(import.meta.dir, ".fixtures")

export async function ensureFixtures(): Promise<void> {
  if (existsSync(path.join(FIXTURE_DIR, "deck.pptx"))) return
  const python = await ensureVenv()
  const script = path.join(import.meta.dir, "..", "src", "python", "gen_fixtures.py")
  const proc = Bun.spawn([python, script, FIXTURE_DIR], { stdout: "pipe", stderr: "pipe" })
  if ((await proc.exited) !== 0) {
    throw new Error(`fixture generation failed: ${await new Response(proc.stderr).text()}`)
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `bun test fixtures`
Expected: PASS (1 test); `test/.fixtures/` now contains both files (gitignored).

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "test: deterministic docx/pptx fixture generation"
```

---

### Task 6: readDocx (outline + content)

**Files:**
- Create: `packages/office-core/src/python/docx_read.py`
- Create: `packages/office-core/src/docx/read.ts`
- Modify: `packages/office-core/src/index.ts` (add `export * from "./docx/read"`)
- Test: `packages/office-core/test/docx-read.test.ts`

**Interfaces:**
- Consumes: `runWorker` (Task 4), fixtures (Task 5).
- Produces:
  - `type DocxElement = { id: string; type: "paragraph"; style: string; text: string } | { id: string; type: "table"; rows: number; cols: number; text?: string }`
  - `type DocxRead = { format: "docx"; mode: string; elements: DocxElement[] }`
  - `readDocx(file: string, mode: "outline" | "content", target?: string, opts?: { cacheDir?: string }): Promise<DocxRead>`
  - `formatDocxRead(result: DocxRead): string` — model-facing text, one element per line, `[id]` prefixed.
  - Worker error codes: `FILE_OPEN`, `TARGET_NOT_FOUND`.

- [ ] **Step 1: Write the failing test**

`packages/office-core/test/docx-read.test.ts`:

```ts
import { beforeAll, expect, test } from "bun:test"
import path from "node:path"
import { OfficeError } from "../src/errors"
import { formatDocxRead, readDocx } from "../src/docx/read"
import { FIXTURE_DIR, ensureFixtures } from "./fixtures"

const REPORT = () => path.join(FIXTURE_DIR, "report.docx")

beforeAll(async () => {
  await ensureFixtures()
}, 180_000)

test("outline returns only headings, with IDs and styles", async () => {
  const out = await readDocx(REPORT(), "outline")
  expect(out.elements.map((e) => (e.type === "paragraph" ? e.text : e.type))).toEqual([
    "Quarterly Report",
    "Regional Breakdown",
    "table",
  ])
  const first = out.elements[0]
  expect(first.type).toBe("paragraph")
  if (first.type === "paragraph") expect(first.style).toBe("Heading 1")
  expect(first.id).toMatch(/^p:\d+$/)
})

test("content returns every block including table text", async () => {
  const out = await readDocx(REPORT(), "content")
  const table = out.elements.find((e) => e.type === "table")
  expect(table).toBeDefined()
  if (table?.type === "table") {
    expect(table.rows).toBe(2)
    expect(table.text).toContain("EMEA | $4.2M")
  }
  expect(formatDocxRead(out)).toContain("Q3 revenue grew 12% year over year.")
})

test("target narrows to a single element", async () => {
  const all = await readDocx(REPORT(), "content")
  const target = all.elements[1].id
  const one = await readDocx(REPORT(), "content", target)
  expect(one.elements).toHaveLength(1)
  expect(one.elements[0].id).toBe(target)
})

test("unknown target raises TARGET_NOT_FOUND", async () => {
  try {
    await readDocx(REPORT(), "content", "p:999")
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("TARGET_NOT_FOUND")
  }
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun test docx-read`
Expected: FAIL — cannot resolve `../src/docx/read`.

- [ ] **Step 3: Write minimal implementation**

`packages/office-core/src/python/docx_read.py`:

```python
from _worker import run, WorkerError
from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph


def iter_blocks(doc):
    for i, child in enumerate(doc.element.body.iterchildren()):
        if child.tag == qn("w:p"):
            yield "p", i, Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield "tbl", i, Table(child, doc)


def render_table(table):
    return "\n".join(" | ".join(cell.text.strip() for cell in row.cells) for row in table.rows)


def main(payload):
    path = payload["file"]
    mode = payload["mode"]
    target = payload.get("target")
    try:
        doc = Document(path)
    except Exception as e:
        raise WorkerError("FILE_OPEN", f"Could not open {path} as .docx: {e}", "Check the path; the file must be a .docx (not legacy .doc).")
    elements = []
    for prefix, index, el in iter_blocks(doc):
        el_id = f"{prefix}:{index}"
        if target and el_id != target:
            continue
        if prefix == "p":
            style = el.style.name if el.style is not None else "Normal"
            if mode == "outline" and not style.startswith("Heading"):
                continue
            elements.append({"id": el_id, "type": "paragraph", "style": style, "text": el.text})
        else:
            entry = {"id": el_id, "type": "table", "rows": len(el.rows), "cols": len(el.columns)}
            if mode != "outline":
                entry["text"] = render_table(el)
            elements.append(entry)
    if target and not elements:
        raise WorkerError("TARGET_NOT_FOUND", f"No element {target} in {path}", "IDs come from office_read output and shift after edits; re-read the file to refresh them.")
    return {"format": "docx", "mode": mode, "elements": elements}


run(main)
```

`packages/office-core/src/docx/read.ts`:

```ts
import { runWorker } from "../worker"

export type DocxElement =
  | { id: string; type: "paragraph"; style: string; text: string }
  | { id: string; type: "table"; rows: number; cols: number; text?: string }

export type DocxRead = { format: "docx"; mode: string; elements: DocxElement[] }

export async function readDocx(
  file: string,
  mode: "outline" | "content",
  target?: string,
  opts?: { cacheDir?: string },
): Promise<DocxRead> {
  return runWorker<DocxRead>("docx_read.py", { file, mode, target }, opts)
}

export function formatDocxRead(result: DocxRead): string {
  return result.elements
    .map((el) =>
      el.type === "paragraph"
        ? `[${el.id}] (${el.style}) ${el.text}`
        : `[${el.id}] (table ${el.rows}x${el.cols})${el.text ? "\n" + el.text : ""}`,
    )
    .join("\n")
}
```

Add to `packages/office-core/src/index.ts`:

```ts
export * from "./docx/read"
```

- [ ] **Step 4: Run test to verify it passes**

Run: `bun test docx-read`
Expected: PASS (4 tests).

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: readDocx outline/content modes with positional IDs"
```

---

### Task 7: readPptx (outline + content)

**Files:**
- Create: `packages/office-core/src/python/pptx_read.py`
- Create: `packages/office-core/src/pptx/read.ts`
- Modify: `packages/office-core/src/index.ts` (add `export * from "./pptx/read"`)
- Test: `packages/office-core/test/pptx-read.test.ts`

**Interfaces:**
- Consumes: `runWorker` (Task 4), fixtures (Task 5).
- Produces:
  - `type PptxShape = { id: string; name: string; text: string }`
  - `type PptxSlide = { id: string; title: string; layout: string; shapes?: PptxShape[]; notes?: string }`
  - `type PptxRead = { format: "pptx"; mode: string; slides: PptxSlide[] }`
  - `readPptx(file: string, mode: "outline" | "content", target?: string, opts?: { cacheDir?: string }): Promise<PptxRead>`
  - `formatPptxRead(result: PptxRead): string`
  - Worker error codes: `FILE_OPEN`, `TARGET_NOT_FOUND`. Shape IDs index ALL shapes on the slide (including non-text ones) so indices stay aligned with python-pptx shape order for Plan 2's edit ops; only shapes with a text frame are listed.

- [ ] **Step 1: Write the failing test**

`packages/office-core/test/pptx-read.test.ts`:

```ts
import { beforeAll, expect, test } from "bun:test"
import path from "node:path"
import { OfficeError } from "../src/errors"
import { formatPptxRead, readPptx } from "../src/pptx/read"
import { FIXTURE_DIR, ensureFixtures } from "./fixtures"

const DECK = () => path.join(FIXTURE_DIR, "deck.pptx")

beforeAll(async () => {
  await ensureFixtures()
}, 180_000)

test("outline lists slides with titles and layouts", async () => {
  const out = await readPptx(DECK(), "outline")
  expect(out.slides.map((s) => s.title)).toEqual(["Q3 Review", "Highlights"])
  expect(out.slides[0].id).toBe("s:0")
  expect(out.slides[0].shapes).toBeUndefined()
})

test("content includes shape text and speaker notes", async () => {
  const out = await readPptx(DECK(), "content")
  const highlights = out.slides[1]
  expect(highlights.notes).toBe("Pause here for questions.")
  const texts = (highlights.shapes ?? []).map((sh) => sh.text).join("\n")
  expect(texts).toContain("EMEA leads growth")
  expect((highlights.shapes ?? [])[0].id).toMatch(/^s:1\/sh:\d+$/)
  expect(formatPptxRead(out)).toContain("Q3 Review")
})

test("target narrows to one slide", async () => {
  const out = await readPptx(DECK(), "content", "s:1")
  expect(out.slides).toHaveLength(1)
  expect(out.slides[0].title).toBe("Highlights")
})

test("unknown slide raises TARGET_NOT_FOUND", async () => {
  try {
    await readPptx(DECK(), "content", "s:9")
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("TARGET_NOT_FOUND")
  }
})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun test pptx-read`
Expected: FAIL — cannot resolve `../src/pptx/read`.

- [ ] **Step 3: Write minimal implementation**

`packages/office-core/src/python/pptx_read.py`:

```python
from _worker import run, WorkerError
from pptx import Presentation


def main(payload):
    path = payload["file"]
    mode = payload["mode"]
    target = payload.get("target")
    try:
        prs = Presentation(path)
    except Exception as e:
        raise WorkerError("FILE_OPEN", f"Could not open {path} as .pptx: {e}", "Check the path; the file must be a .pptx (not legacy .ppt).")
    slides = []
    for si, slide in enumerate(prs.slides):
        sid = f"s:{si}"
        if target and target != sid and not target.startswith(sid + "/"):
            continue
        title = slide.shapes.title.text if slide.shapes.title is not None else ""
        entry = {"id": sid, "title": title, "layout": slide.slide_layout.name}
        if mode == "content":
            shapes = []
            for shi, shape in enumerate(slide.shapes):
                if not shape.has_text_frame:
                    continue
                text = "\n".join(p.text for p in shape.text_frame.paragraphs)
                shapes.append({"id": f"s:{si}/sh:{shi}", "name": shape.name, "text": text})
            entry["shapes"] = shapes
            if slide.has_notes_slide:
                entry["notes"] = slide.notes_slide.notes_text_frame.text
        slides.append(entry)
    if target and not slides:
        raise WorkerError("TARGET_NOT_FOUND", f"No slide {target} in {path}", "Slide IDs come from office_read output and shift after edits; re-read to refresh them.")
    return {"format": "pptx", "mode": mode, "slides": slides}


run(main)
```

`packages/office-core/src/pptx/read.ts`:

```ts
import { runWorker } from "../worker"

export type PptxShape = { id: string; name: string; text: string }
export type PptxSlide = { id: string; title: string; layout: string; shapes?: PptxShape[]; notes?: string }
export type PptxRead = { format: "pptx"; mode: string; slides: PptxSlide[] }

export async function readPptx(
  file: string,
  mode: "outline" | "content",
  target?: string,
  opts?: { cacheDir?: string },
): Promise<PptxRead> {
  return runWorker<PptxRead>("pptx_read.py", { file, mode, target }, opts)
}

export function formatPptxRead(result: PptxRead): string {
  return result.slides
    .map((slide) => {
      const lines = [`[${slide.id}] ${slide.title || "(untitled)"} — layout: ${slide.layout}`]
      for (const shape of slide.shapes ?? []) lines.push(`  [${shape.id}] (${shape.name}) ${shape.text}`)
      if (slide.notes) lines.push(`  notes: ${slide.notes}`)
      return lines.join("\n")
    })
    .join("\n")
}
```

Add to `packages/office-core/src/index.ts`:

```ts
export * from "./pptx/read"
```

- [ ] **Step 4: Run test to verify it passes**

Run: `bun test pptx-read`
Expected: PASS (4 tests). Then run the full suite: `bun test` — everything passes.

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: readPptx outline/content modes with slide/shape IDs"
```
