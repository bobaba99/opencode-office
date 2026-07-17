# Office Tools Plan 2: Edit Ops Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** `editDocx`/`editPptx` with anchor-validated targeted operations, atomic single-save writes, content-addressed backups, and full-mode docx read (tracked changes).

**Architecture:** Same TS-orchestrates/Python-executes split as Plan 1. Edit workers apply operations **sequentially in memory** (each op's IDs refer to the document state after all preceding ops in the batch), then save **once** via temp-file + `os.replace` — any op failure means nothing is written. TS wraps each edit with a content-addressed backup of the original file.

**Tech Stack:** unchanged from Plan 1 (Bun/TS, python-docx 1.2.0, python-pptx 1.0.2, pillow 11.3.0).

**Spec:** `docs/superpowers/specs/2026-07-17-office-tools-design.md`. Plan 2 of 4. Spec deviations, deliberate: comments reading moves to Plan 3 (needs OOXML part authoring for fixtures); `reorder_slides` is realized as `move_slide` (target + 0-based destination index — simpler for models than a full permutation).

## Global Constraints

- Everything from Plan 1's Global Constraints still binds (pins, OfficeError `{code, message, hint}` with recovery-path hints, ID grammar, no OpenCode imports, `bun test` from repo root, commit per task).
- **Anchor rule:** anchors are exact substrings of the target element's current text. `ANCHOR_MISMATCH` errors must include the element's actual current text (truncated to 300 chars). An anchor matching more than once → `AMBIGUOUS_ANCHOR` telling the model to extend the anchor.
- **Batch semantics:** ops apply in order; each op's IDs refer to the state after preceding ops. Failure of any op → no file write at all.
- **Atomic write:** save to `<file>.tmp-opencode-office` then `os.replace`.
- **Formatting preservation:** text replacement must keep run formatting — replacement text goes into the first overlapping run; other overlapped runs keep only their uncovered head/tail.
- Error codes introduced here: `AMBIGUOUS_ANCHOR`, `ANCHOR_MISMATCH`, `BAD_TARGET_KIND`, `CELL_OUT_OF_RANGE`, `STYLE_NOT_FOUND`, `LAYOUT_NOT_FOUND`, `SHAPE_NOT_PICTURE`, `UNKNOWN_OP`.

---

### Task 1: Edit fixtures

**Files:**
- Modify: `packages/office-core/src/python/gen_fixtures.py`
- Modify: `packages/office-core/test/fixtures.ts` (sentinel: last-written file changes)
- Test: `packages/office-core/test/fixtures.test.ts` (extend)

**Interfaces:**
- Consumes: Plan 1 fixture generator.
- Produces: NEW fixture files (existing `report.docx`/`deck.pptx` are untouched — Plan 1 tests depend on their exact shape): `edit-report.docx` (Heading "Edit Playground"; multi-run paragraph "Growth was **strong** this quarter overall."; "Delete me entirely."; "Style me."; 2x2 table K/V/alpha/one; paragraph "Reviewed text " carrying a tracked `w:ins` insertion "with tracked insertion"), `edit-deck.pptx` (3 slides: title "Edit Deck"/"v1", bullets "Points" with "First point\nSecond point", blank slide with one 64x64 red PNG picture), and `swap.png` (32x32 blue PNG for replace_image tests).

- [ ] **Step 1: Extend the failing test**

Append to `packages/office-core/test/fixtures.test.ts`:

```ts
test("generates edit fixtures", async () => {
  await ensureFixtures()
  for (const name of ["edit-report.docx", "edit-deck.pptx", "swap.png"]) {
    expect(existsSync(path.join(FIXTURE_DIR, name))).toBe(true)
  }
}, 180_000)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `rm -rf packages/office-core/test/.fixtures && bun test fixtures`
Expected: new test FAILS (files missing).

- [ ] **Step 3: Implement**

Append to `packages/office-core/src/python/gen_fixtures.py` (before the `__main__` block):

```python
import io

from docx.oxml.ns import qn
from pptx.util import Inches
from PIL import Image as PILImage


def make_edit_docx(path):
    doc = Document()
    doc.add_heading("Edit Playground", level=1)
    p = doc.add_paragraph("Growth was ")
    strong = p.add_run("strong")
    strong.bold = True
    p.add_run(" this quarter overall.")
    doc.add_paragraph("Delete me entirely.")
    doc.add_paragraph("Style me.")
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "K"
    table.cell(0, 1).text = "V"
    table.cell(1, 0).text = "alpha"
    table.cell(1, 1).text = "one"
    tracked = doc.add_paragraph("Reviewed text ")
    ins = tracked._p.makeelement(
        qn("w:ins"), {qn("w:id"): "1", qn("w:author"): "Fixture", qn("w:date"): "2026-01-01T00:00:00Z"}
    )
    run_el = tracked._p.makeelement(qn("w:r"), {})
    text_el = tracked._p.makeelement(qn("w:t"), {})
    text_el.text = "with tracked insertion"
    run_el.append(text_el)
    ins.append(run_el)
    tracked._p.append(ins)
    doc.save(path)


def make_edit_pptx(path):
    prs = Presentation()
    s1 = prs.slides.add_slide(prs.slide_layouts[0])
    s1.shapes.title.text = "Edit Deck"
    s1.placeholders[1].text = "v1"
    s2 = prs.slides.add_slide(prs.slide_layouts[1])
    s2.shapes.title.text = "Points"
    s2.placeholders[1].text = "First point\nSecond point"
    buf = io.BytesIO()
    PILImage.new("RGB", (64, 64), (200, 30, 30)).save(buf, format="PNG")
    buf.seek(0)
    s3 = prs.slides.add_slide(prs.slide_layouts[6])
    s3.shapes.add_picture(buf, Inches(1), Inches(1))
    prs.save(path)


def make_png(path, color, size):
    PILImage.new("RGB", size, color).save(path)
```

And extend the `__main__` block to also call:

```python
    make_edit_docx(os.path.join(out, "edit-report.docx"))
    make_edit_pptx(os.path.join(out, "edit-deck.pptx"))
    make_png(os.path.join(out, "swap.png"), (30, 30, 200), (32, 32))
```

In `packages/office-core/test/fixtures.ts`, change the sentinel check to the last-written file:

```ts
  if (existsSync(path.join(FIXTURE_DIR, "swap.png"))) return
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `bun test fixtures` then full `bun test`
Expected: all pass (Plan 1 read tests still green — old fixtures unchanged).

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "test: edit-scenario fixtures (multi-run, tracked change, picture slide)"
```

---

### Task 2: Shared docx helpers + run-preserving text replacement

**Files:**
- Create: `packages/office-core/src/python/_docx_common.py`
- Create: `packages/office-core/src/python/_textops.py`
- Modify: `packages/office-core/src/python/docx_read.py` (import from `_docx_common` instead of defining `iter_blocks`/`render_table` locally)

**Interfaces:**
- Produces (Python): `_docx_common.iter_blocks(doc)` yielding `(prefix, index, element)`; `_docx_common.render_table(table)`; `_textops.para_text(p)` (joins `run.text` — works for both python-docx and python-pptx paragraph proxies); `_textops.replace_in_paragraph(p, anchor, replacement) -> bool` (first occurrence; replacement inherits the first overlapped run's formatting; other overlapped runs keep only uncovered head/tail).

- [ ] **Step 1: This is a refactor + new pure helpers — the regression gate is the existing suite**

`_textops.replace_in_paragraph` is exercised directly by Task 3's cross-run test; no standalone probe worker (YAGNI).

- [ ] **Step 2: Run `bun test` to establish green baseline (24+ tests pass)**

- [ ] **Step 3: Implement**

`packages/office-core/src/python/_docx_common.py`:

```python
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
```

`packages/office-core/src/python/_textops.py`:

```python
def para_text(p):
    return "".join(run.text or "" for run in p.runs)


def replace_in_paragraph(p, anchor, replacement):
    """Replace the first occurrence of anchor, preserving run formatting.

    The replacement text lands in the first run that overlaps the anchor
    (inheriting its formatting); every other overlapped run keeps only the
    parts of its text outside the anchor. Returns True if a replacement
    happened.
    """
    text = para_text(p)
    start = text.find(anchor)
    if start < 0:
        return False
    end = start + len(anchor)
    pos = 0
    replaced = False
    for run in p.runs:
        run_text = run.text or ""
        run_start, run_end = pos, pos + len(run_text)
        pos = run_end
        if run_end <= start or run_start >= end:
            continue
        head = run_text[: max(0, start - run_start)]
        tail = run_text[max(0, min(len(run_text), end - run_start)):]
        if not replaced:
            run.text = head + replacement + tail
            replaced = True
        else:
            run.text = head + tail
    return replaced
```

In `packages/office-core/src/python/docx_read.py`: delete the local `iter_blocks` and `render_table` definitions and their `Table`/`Paragraph`/`qn` imports, replacing with:

```python
from _docx_common import iter_blocks, render_table
```

- [ ] **Step 4: Run `bun test` — all tests still pass (refactor is behavior-neutral)**

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "refactor: shared docx block iteration; run-preserving text replacement helper"
```

---

### Task 3: editDocx

**Files:**
- Create: `packages/office-core/src/python/docx_edit.py`
- Create: `packages/office-core/src/backup.ts`
- Create: `packages/office-core/src/docx/edit.ts`
- Modify: `packages/office-core/src/index.ts` (add `export * from "./backup"` and `export * from "./docx/edit"`)
- Test: `packages/office-core/test/docx-edit.test.ts`

**Interfaces:**
- Consumes: `runWorker`, `parseId`, `OfficeError`, `defaultCacheDir`, fixtures.
- Produces:
  - `backupFile(file: string, cacheDir?: string): Promise<string>` — copies the file to `<cacheDir>/backups/<sha256-12>-<basename>` (content-addressed, idempotent), returns the backup path. Unreadable file → `OfficeError("FILE_OPEN", ...)`.
  - `type DocxOperation` — union of `{op:"replace_text"; target; anchor; text}`, `{op:"insert_content"; after; markdown}`, `{op:"delete_element"; target; anchor}`, `{op:"set_style"; target; anchor; style}`, `{op:"set_table_cell"; target; row; col; text; anchor?}` (all ids strings; row/col numbers).
  - `type EditResult = { applied: number; results: Array<Record<string, unknown>>; backup?: string }`
  - `editDocx(file: string, operations: DocxOperation[], opts?: { backup?: boolean; cacheDir?: string }): Promise<EditResult>` — validates every op's `target`/`after` with `parseId` (must be paragraph/table kind, else `BAD_ID`), takes the backup (unless `opts.backup === false`), then runs the worker.
  - Worker semantics: sequential per-op validate+apply against a freshly rebuilt index, single `doc.save(tmp)` + `os.replace` at the end. Markdown mapping for `insert_content`: `# `→Heading 1, `## `→Heading 2, `### `→Heading 3, `- `→List Bullet, anything else→default paragraph style; blank lines skipped.

- [ ] **Step 1: Write the failing test**

`packages/office-core/test/docx-edit.test.ts`:

```ts
import { beforeEach, expect, test } from "bun:test"
import { copyFile } from "node:fs/promises"
import path from "node:path"
import { OfficeError } from "../src/errors"
import { editDocx } from "../src/docx/edit"
import { readDocx } from "../src/docx/read"
import { FIXTURE_DIR, ensureFixtures } from "./fixtures"

const WORK = path.join(FIXTURE_DIR, "work-report.docx")

beforeEach(async () => {
  await ensureFixtures()
  await copyFile(path.join(FIXTURE_DIR, "edit-report.docx"), WORK)
}, 180_000)

test("replace_text across runs preserves surrounding text", async () => {
  const before = await readDocx(WORK, "content")
  const target = before.elements.find((e) => e.type === "paragraph" && e.text.includes("strong"))!
  const result = await editDocx(WORK, [
    { op: "replace_text", target: target.id, anchor: "was strong this", text: "was exceptional this" },
  ])
  expect(result.applied).toBe(1)
  const after = await readDocx(WORK, "content", target.id)
  expect((after.elements[0] as { text: string }).text).toBe("Growth was exceptional this quarter overall.")
})

test("batch applies in order and untouched elements are unchanged", async () => {
  const before = await readDocx(WORK, "content")
  const del = before.elements.find((e) => e.type === "paragraph" && e.text === "Delete me entirely.")!
  const style = before.elements.find((e) => e.type === "paragraph" && e.text === "Style me.")!
  await editDocx(WORK, [
    { op: "delete_element", target: del.id, anchor: "Delete me entirely." },
    // NOTE: after the delete, positional IDs above the deleted element shift down by one.
    { op: "set_style", target: `p:${Number(style.id.split(":")[1]) - 1}`, anchor: "Style me.", style: "Heading 2" },
  ])
  const after = await readDocx(WORK, "content")
  expect(after.elements.some((e) => e.type === "paragraph" && e.text === "Delete me entirely.")).toBe(false)
  const styled = after.elements.find((e) => e.type === "paragraph" && e.text === "Style me.")!
  expect(styled.type === "paragraph" && styled.style).toBe("Heading 2")
  expect(after.elements[0]).toEqual(before.elements[0])
})

test("insert_content inserts styled markdown after an element", async () => {
  const before = await readDocx(WORK, "content")
  const heading = before.elements[0]
  await editDocx(WORK, [
    { op: "insert_content", after: heading.id, markdown: "## New Section\nIntro line.\n- bullet one" },
  ])
  const after = await readDocx(WORK, "content")
  const texts = after.elements.map((e) => (e.type === "paragraph" ? e.text : "table"))
  expect(texts.slice(1, 4)).toEqual(["New Section", "Intro line.", "bullet one"])
  const section = after.elements[1]
  expect(section.type === "paragraph" && section.style).toBe("Heading 2")
})

test("set_table_cell updates one cell", async () => {
  const before = await readDocx(WORK, "content")
  const tbl = before.elements.find((e) => e.type === "table")!
  await editDocx(WORK, [{ op: "set_table_cell", target: tbl.id, row: 1, col: 1, text: "two", anchor: "one" }])
  const after = await readDocx(WORK, "content", tbl.id)
  expect((after.elements[0] as { text?: string }).text).toContain("alpha | two")
})

test("anchor mismatch aborts the whole batch atomically", async () => {
  const before = await readDocx(WORK, "content")
  const first = before.elements.find((e) => e.type === "paragraph" && e.text.includes("Growth"))!
  try {
    await editDocx(WORK, [
      { op: "replace_text", target: first.id, anchor: "Growth was ", text: "Expansion was " },
      { op: "replace_text", target: first.id, anchor: "NOT PRESENT", text: "x" },
    ])
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("ANCHOR_MISMATCH")
    expect((e as OfficeError).hint).toContain("Growth")
  }
  const after = await readDocx(WORK, "content")
  expect(after.elements).toEqual(before.elements)
})

test("ambiguous anchor is rejected with guidance", async () => {
  await editDocx(WORK, [
    { op: "insert_content", after: "p:0", markdown: "same same" },
  ])
  const doc = await readDocx(WORK, "content")
  const dup = doc.elements.find((e) => e.type === "paragraph" && e.text === "same same")!
  try {
    await editDocx(WORK, [{ op: "replace_text", target: dup.id, anchor: "same", text: "x" }])
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("AMBIGUOUS_ANCHOR")
  }
})

test("unknown style yields STYLE_NOT_FOUND", async () => {
  try {
    await editDocx(WORK, [{ op: "set_style", target: "p:0", anchor: "Edit Playground", style: "No Such Style" }])
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("STYLE_NOT_FOUND")
  }
})

test("backup is written before editing and contains the original", async () => {
  const original = await Bun.file(WORK).arrayBuffer()
  const result = await editDocx(WORK, [
    { op: "replace_text", target: "p:1", anchor: "Growth", text: "Expansion" },
  ])
  expect(result.backup).toBeDefined()
  const backup = await Bun.file(result.backup!).arrayBuffer()
  expect(Buffer.from(backup).equals(Buffer.from(original))).toBe(true)
})

test("pptx-style target is rejected client-side", async () => {
  try {
    await editDocx(WORK, [{ op: "replace_text", target: "s:0", anchor: "x", text: "y" }])
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("BAD_ID")
  }
})
```

- [ ] **Step 2: Run `bun test docx-edit` — FAIL (cannot resolve `../src/docx/edit`)**

- [ ] **Step 3: Implement**

`packages/office-core/src/backup.ts`:

```ts
import { createHash } from "node:crypto"
import { existsSync } from "node:fs"
import { mkdir, readFile, writeFile } from "node:fs/promises"
import path from "node:path"
import { OfficeError } from "./errors"
import { defaultCacheDir } from "./runtime"

export async function backupFile(file: string, cacheDir = defaultCacheDir()): Promise<string> {
  let data: Buffer
  try {
    data = await readFile(file)
  } catch (e) {
    throw new OfficeError(
      "FILE_OPEN",
      `Cannot read ${file} for backup: ${e instanceof Error ? e.message : String(e)}`,
      "Check that the path exists and is readable before editing.",
    )
  }
  const hash = createHash("sha256").update(data).digest("hex").slice(0, 12)
  const dir = path.join(cacheDir, "backups")
  await mkdir(dir, { recursive: true })
  const dest = path.join(dir, `${hash}-${path.basename(file)}`)
  if (!existsSync(dest)) await writeFile(dest, data)
  return dest
}
```

`packages/office-core/src/python/docx_edit.py`:

```python
import os

from _worker import run, WorkerError
from _docx_common import iter_blocks, render_table
from _textops import para_text, replace_in_paragraph
from docx import Document

MD_STYLES = [("### ", "Heading 3"), ("## ", "Heading 2"), ("# ", "Heading 1"), ("- ", "List Bullet")]


def build_index(doc):
    return {f"{prefix}:{index}": (prefix, element) for prefix, index, element in iter_blocks(doc)}


def require(index, op_name, target, kinds):
    if target not in index:
        raise WorkerError(
            "TARGET_NOT_FOUND",
            f"No element {target}",
            "IDs come from office_read and shift as the batch applies; re-read the file, or order ops bottom-up.",
        )
    prefix, element = index[target]
    if prefix not in kinds:
        raise WorkerError(
            "BAD_TARGET_KIND",
            f"{op_name} cannot target {target} (a {prefix} element)",
            f"{op_name} targets {' or '.join(kinds)} elements.",
        )
    return prefix, element


def check_anchor(current, anchor, target):
    if anchor not in current:
        raise WorkerError(
            "ANCHOR_MISMATCH",
            f"Anchor not found at {target}",
            f"Element currently reads: {current[:300]!r} — update the anchor to match, or re-read the file.",
        )


def styled_line(line):
    for marker, style in MD_STYLES:
        if line.startswith(marker):
            return line[len(marker):], style
    return line, None


def add_styled_paragraph(doc, text, style):
    try:
        return doc.add_paragraph(text, style=style) if style else doc.add_paragraph(text)
    except KeyError:
        raise WorkerError(
            "STYLE_NOT_FOUND",
            f"Style {style!r} does not exist in this document",
            "Only styles the document defines can be used; # / ## / ### / - map to Heading 1-3 / List Bullet.",
        )


def apply_one(doc, op):
    kind = op["op"]
    index = build_index(doc)
    if kind == "replace_text":
        _, el = require(index, kind, op["target"], ["p"])
        current = para_text(el)
        check_anchor(current, op["anchor"], op["target"])
        if current.count(op["anchor"]) > 1:
            raise WorkerError(
                "AMBIGUOUS_ANCHOR",
                f"Anchor occurs more than once in {op['target']}",
                "Extend the anchor with surrounding text until it is unique within the element.",
            )
        replace_in_paragraph(el, op["anchor"], op["text"])
        return {"op": kind, "target": op["target"], "text_after": para_text(el)}
    if kind == "insert_content":
        _, el = require(index, kind, op["after"], ["p", "tbl"])
        anchor_el = el._p if hasattr(el, "_p") else el._tbl
        for line in [line for line in op["markdown"].splitlines() if line.strip()]:
            text, style = styled_line(line)
            new_p = add_styled_paragraph(doc, text, style)
            anchor_el.addnext(new_p._p)
            anchor_el = new_p._p
        return {"op": kind, "after": op["after"]}
    if kind == "delete_element":
        prefix, el = require(index, kind, op["target"], ["p", "tbl"])
        current = para_text(el) if prefix == "p" else render_table(el).split("\n")[0]
        check_anchor(current, op["anchor"], op["target"])
        xml_el = el._p if prefix == "p" else el._tbl
        xml_el.getparent().remove(xml_el)
        return {"op": kind, "target": op["target"]}
    if kind == "set_style":
        _, el = require(index, kind, op["target"], ["p"])
        check_anchor(para_text(el), op["anchor"], op["target"])
        try:
            el.style = doc.styles[op["style"]]
        except KeyError:
            raise WorkerError(
                "STYLE_NOT_FOUND",
                f"Style {op['style']!r} does not exist in this document",
                "office_read shows each paragraph's style; only styles the document defines can be applied.",
            )
        return {"op": kind, "target": op["target"]}
    if kind == "set_table_cell":
        _, el = require(index, kind, op["target"], ["tbl"])
        rows, cols = len(el.rows), len(el.columns)
        if not (0 <= op["row"] < rows and 0 <= op["col"] < cols):
            raise WorkerError(
                "CELL_OUT_OF_RANGE",
                f"{op['target']} is {rows}x{cols}; cell ({op['row']},{op['col']}) does not exist",
                "Row and col are 0-based and must be inside the dimensions office_read reports.",
            )
        if op.get("anchor") is not None and op["anchor"] != el.cell(op["row"], op["col"]).text:
            raise WorkerError(
                "ANCHOR_MISMATCH",
                f"Cell ({op['row']},{op['col']}) of {op['target']} does not match anchor",
                f"Cell currently reads: {el.cell(op['row'], op['col']).text[:300]!r}",
            )
        el.cell(op["row"], op["col"]).text = op["text"]
        return {"op": kind, "target": op["target"]}
    raise WorkerError(
        "UNKNOWN_OP",
        f"Unknown docx op: {kind}",
        "Valid ops: replace_text, insert_content, delete_element, set_style, set_table_cell.",
    )


def main(payload):
    path = payload["file"]
    try:
        doc = Document(path)
    except Exception as e:
        raise WorkerError("FILE_OPEN", f"Could not open {path} as .docx: {e}", "Check the path; the file must be a .docx (not legacy .doc).")
    results = [apply_one(doc, op) for op in payload["operations"]]
    tmp = path + ".tmp-opencode-office"
    doc.save(tmp)
    os.replace(tmp, path)
    return {"applied": len(results), "results": results}


run(main)
```

`packages/office-core/src/docx/edit.ts`:

```ts
import { OfficeError } from "../errors"
import { parseId } from "../ids"
import { backupFile } from "../backup"
import { runWorker } from "../worker"

export type DocxOperation =
  | { op: "replace_text"; target: string; anchor: string; text: string }
  | { op: "insert_content"; after: string; markdown: string }
  | { op: "delete_element"; target: string; anchor: string }
  | { op: "set_style"; target: string; anchor: string; style: string }
  | { op: "set_table_cell"; target: string; row: number; col: number; text: string; anchor?: string }

export type EditResult = { applied: number; results: Array<Record<string, unknown>>; backup?: string }

function assertDocxId(id: string): void {
  const ref = parseId(id)
  if (ref.kind !== "paragraph" && ref.kind !== "table")
    throw new OfficeError(
      "BAD_ID",
      `Target ${id} is not a docx element ID`,
      "docx targets use p:<n> or tbl:<n> — get IDs from office_read output for this file.",
    )
}

export async function editDocx(
  file: string,
  operations: DocxOperation[],
  opts?: { backup?: boolean; cacheDir?: string },
): Promise<EditResult> {
  for (const operation of operations) {
    assertDocxId("target" in operation ? operation.target : operation.after)
  }
  const backup = opts?.backup === false ? undefined : await backupFile(file, opts?.cacheDir)
  const data = await runWorker<Omit<EditResult, "backup">>("docx_edit.py", { file, operations }, opts)
  return { ...data, backup }
}
```

Add to `packages/office-core/src/index.ts`:

```ts
export * from "./backup"
export * from "./docx/edit"
```

- [ ] **Step 4: Run `bun test docx-edit` (9 tests pass), then full `bun test` and `bun run typecheck`**

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: editDocx — anchored ops, atomic save, content-addressed backups"
```

---

### Task 4: editPptx

**Files:**
- Create: `packages/office-core/src/python/pptx_edit.py`
- Create: `packages/office-core/src/pptx/edit.ts`
- Modify: `packages/office-core/src/index.ts` (add `export * from "./pptx/edit"`)
- Test: `packages/office-core/test/pptx-edit.test.ts`

**Interfaces:**
- Consumes: `runWorker`, `parseId`, `OfficeError`, `backupFile`, `_textops`, fixtures.
- Produces:
  - `type PptxOperation` — union of `{op:"set_shape_text"; target; anchor; text}` (target is a shape id), `{op:"set_notes"; target; text}` (slide id; overwrites notes), `{op:"insert_slide"; after; layout; title?; bullets?: string[]}`, `{op:"duplicate_slide"; target}`, `{op:"delete_slide"; target}`, `{op:"move_slide"; target; index}` (0-based destination), `{op:"replace_image"; target; image}` (shape id + image file path).
  - `editPptx(file: string, operations: PptxOperation[], opts?: { backup?: boolean; cacheDir?: string }): Promise<EditResult>` (reuses `EditResult` from Task 3) — validates ids client-side (slide/shape kinds only, else `BAD_ID`), backup, worker.
  - Worker: sequential apply, single save via tmp + `os.replace`. `insert_slide`/`duplicate_slide` append then move the new slide to directly after the reference slide. `duplicate_slide` deep-copies shape XML, re-relates image/media parts and rewrites `r:embed` ids, and copies notes text. Anchors for `set_shape_text` must not span paragraph breaks (the mismatch hint says so).

- [ ] **Step 1: Write the failing test**

`packages/office-core/test/pptx-edit.test.ts`:

```ts
import { beforeEach, expect, test } from "bun:test"
import { copyFile } from "node:fs/promises"
import path from "node:path"
import { OfficeError } from "../src/errors"
import { editPptx } from "../src/pptx/edit"
import { readPptx } from "../src/pptx/read"
import { FIXTURE_DIR, ensureFixtures } from "./fixtures"

const WORK = path.join(FIXTURE_DIR, "work-deck.pptx")

beforeEach(async () => {
  await ensureFixtures()
  await copyFile(path.join(FIXTURE_DIR, "edit-deck.pptx"), WORK)
}, 180_000)

test("set_shape_text replaces anchored text in a shape", async () => {
  const before = await readPptx(WORK, "content")
  const bullets = before.slides[1].shapes!.find((sh) => sh.text.includes("Second point"))!
  await editPptx(WORK, [{ op: "set_shape_text", target: bullets.id, anchor: "Second point", text: "Second point, revised" }])
  const after = await readPptx(WORK, "content", "s:1")
  expect(after.slides[0].shapes!.map((sh) => sh.text).join("\n")).toContain("Second point, revised")
})

test("set_notes overwrites speaker notes", async () => {
  await editPptx(WORK, [{ op: "set_notes", target: "s:0", text: "Open with the numbers." }])
  const after = await readPptx(WORK, "content", "s:0")
  expect(after.slides[0].notes).toBe("Open with the numbers.")
})

test("insert_slide lands directly after the reference slide with theme layout", async () => {
  await editPptx(WORK, [
    { op: "insert_slide", after: "s:0", layout: "Title and Content", title: "Agenda", bullets: ["One", "Two"] },
  ])
  const after = await readPptx(WORK, "outline")
  expect(after.slides.map((s) => s.title)).toEqual(["Edit Deck", "Agenda", "Points", ""])
  expect(after.slides[1].layout).toBe("Title and Content")
})

test("duplicate_slide copies a picture slide without breaking the file", async () => {
  await editPptx(WORK, [{ op: "duplicate_slide", target: "s:2" }])
  const after = await readPptx(WORK, "outline")
  expect(after.slides).toHaveLength(4)
})

test("delete_slide and move_slide restructure the deck", async () => {
  await editPptx(WORK, [
    { op: "delete_slide", target: "s:2" },
    { op: "move_slide", target: "s:1", index: 0 },
  ])
  const after = await readPptx(WORK, "outline")
  expect(after.slides.map((s) => s.title)).toEqual(["Points", "Edit Deck"])
})

test("replace_image swaps picture bytes", async () => {
  const swap = path.join(FIXTURE_DIR, "swap.png")
  const before = await readPptx(WORK, "content", "s:2")
  await editPptx(WORK, [{ op: "replace_image", target: "s:2/sh:0", image: swap }])
  const after = await readPptx(WORK, "content", "s:2")
  expect(after.slides).toHaveLength(1)
  expect(before.slides).toHaveLength(1)
})

test("replace_image on a text shape fails with SHAPE_NOT_PICTURE", async () => {
  const swap = path.join(FIXTURE_DIR, "swap.png")
  try {
    await editPptx(WORK, [{ op: "replace_image", target: "s:0/sh:0", image: swap }])
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("SHAPE_NOT_PICTURE")
  }
})

test("unknown layout lists available layouts", async () => {
  try {
    await editPptx(WORK, [{ op: "insert_slide", after: "s:0", layout: "Nope", title: "x" }])
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("LAYOUT_NOT_FOUND")
    expect((e as OfficeError).hint).toContain("Title and Content")
  }
})

test("failed op aborts the batch atomically", async () => {
  const before = await readPptx(WORK, "content")
  try {
    await editPptx(WORK, [
      { op: "set_notes", target: "s:0", text: "should not survive" },
      { op: "set_shape_text", target: "s:1/sh:1", anchor: "NOT THERE", text: "x" },
    ])
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("ANCHOR_MISMATCH")
  }
  const after = await readPptx(WORK, "content")
  expect(after.slides).toEqual(before.slides)
})

test("docx-style target is rejected client-side", async () => {
  try {
    await editPptx(WORK, [{ op: "set_notes", target: "p:0", text: "x" }])
    expect.unreachable()
  } catch (e) {
    expect((e as OfficeError).code).toBe("BAD_ID")
  }
})
```

- [ ] **Step 2: Run `bun test pptx-edit` — FAIL (cannot resolve `../src/pptx/edit`)**

- [ ] **Step 3: Implement**

`packages/office-core/src/python/pptx_edit.py`:

```python
import os
from copy import deepcopy

from _worker import run, WorkerError
from _textops import para_text, replace_in_paragraph
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn


def slide_at(prs, sid):
    n = int(sid.split(":")[1].split("/")[0])
    slides = list(prs.slides)
    if n >= len(slides):
        raise WorkerError(
            "TARGET_NOT_FOUND",
            f"No slide s:{n}",
            "Slide IDs shift as the batch applies; re-read the file, or order structural ops last.",
        )
    return n, slides[n]


def shape_at(slide, sid):
    m = int(sid.split("/sh:")[1])
    shapes = list(slide.shapes)
    if m >= len(shapes):
        raise WorkerError(
            "TARGET_NOT_FOUND",
            f"No shape {sid}",
            "Shape IDs come from office_read content mode; re-read the file to refresh them.",
        )
    return shapes[m]


def find_layout(prs, name):
    names = []
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            names.append(layout.name)
            if layout.name == name:
                return layout
    raise WorkerError("LAYOUT_NOT_FOUND", f"No slide layout named {name!r}", f"Available layouts: {', '.join(names)}")


def replace_in_frame(shape, anchor, text, target):
    if not shape.has_text_frame:
        raise WorkerError(
            "BAD_TARGET_KIND",
            f"{target} has no text frame",
            "set_shape_text targets text shapes; office_read content mode lists them.",
        )
    frame = shape.text_frame
    total = sum(para_text(p).count(anchor) for p in frame.paragraphs)
    if total == 0:
        current = "\n".join(para_text(p) for p in frame.paragraphs)
        raise WorkerError(
            "ANCHOR_MISMATCH",
            f"Anchor not found in {target}",
            f"Shape currently reads: {current[:300]!r} — anchors cannot span paragraph breaks.",
        )
    if total > 1:
        raise WorkerError(
            "AMBIGUOUS_ANCHOR",
            f"Anchor occurs {total} times in {target}",
            "Extend the anchor with surrounding text until it is unique.",
        )
    for p in frame.paragraphs:
        if replace_in_paragraph(p, anchor, text):
            return


def move_entry(prs, from_index, to_index):
    lst = prs.slides._sldIdLst
    entry = list(lst)[from_index]
    lst.remove(entry)
    remaining = list(lst)
    to_index = max(0, min(to_index, len(remaining)))
    if to_index == len(remaining):
        lst.append(entry)
    else:
        remaining[to_index].addprevious(entry)


def delete_entry(prs, index):
    lst = prs.slides._sldIdLst
    entry = list(lst)[index]
    prs.part.drop_rel(entry.get(qn("r:id")))
    lst.remove(entry)


def copy_slide(prs, source):
    new_slide = prs.slides.add_slide(source.slide_layout)
    for shape in list(new_slide.shapes):
        shape._element.getparent().remove(shape._element)
    rid_map = {}
    for rid, rel in source.part.rels.items():
        if rel.is_external:
            continue
        if "image" in rel.reltype or "media" in rel.reltype:
            rid_map[rid] = new_slide.part.relate_to(rel.target_part, rel.reltype)
    for shape in source.shapes:
        el = deepcopy(shape._element)
        for blip in el.iter(qn("a:blip")):
            embed = blip.get(qn("r:embed"))
            if embed in rid_map:
                blip.set(qn("r:embed"), rid_map[embed])
        new_slide.shapes._spTree.append(el)
    if source.has_notes_slide:
        new_slide.notes_slide.notes_text_frame.text = source.notes_slide.notes_text_frame.text
    return new_slide


def apply_one(prs, op):
    kind = op["op"]
    if kind == "set_shape_text":
        _, slide = slide_at(prs, op["target"])
        shape = shape_at(slide, op["target"])
        replace_in_frame(shape, op["anchor"], op["text"], op["target"])
        return {"op": kind, "target": op["target"]}
    if kind == "set_notes":
        _, slide = slide_at(prs, op["target"])
        slide.notes_slide.notes_text_frame.text = op["text"]
        return {"op": kind, "target": op["target"]}
    if kind == "insert_slide":
        n, _ = slide_at(prs, op["after"])
        layout = find_layout(prs, op["layout"])
        new_slide = prs.slides.add_slide(layout)
        if op.get("title") is not None and new_slide.shapes.title is not None:
            new_slide.shapes.title.text = op["title"]
        if op.get("bullets"):
            for placeholder in new_slide.placeholders:
                if placeholder.placeholder_format.idx == 1:
                    placeholder.text = "\n".join(op["bullets"])
                    break
        move_entry(prs, len(list(prs.slides)) - 1, n + 1)
        return {"op": kind, "after": op["after"]}
    if kind == "duplicate_slide":
        n, slide = slide_at(prs, op["target"])
        copy_slide(prs, slide)
        move_entry(prs, len(list(prs.slides)) - 1, n + 1)
        return {"op": kind, "target": op["target"]}
    if kind == "delete_slide":
        n, _ = slide_at(prs, op["target"])
        delete_entry(prs, n)
        return {"op": kind, "target": op["target"]}
    if kind == "move_slide":
        n, _ = slide_at(prs, op["target"])
        move_entry(prs, n, op["index"])
        return {"op": kind, "target": op["target"], "index": op["index"]}
    if kind == "replace_image":
        _, slide = slide_at(prs, op["target"])
        shape = shape_at(slide, op["target"])
        if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            raise WorkerError(
                "SHAPE_NOT_PICTURE",
                f"{op['target']} is not a picture shape",
                "replace_image targets picture shapes; text shapes keep their id but have no image to swap.",
            )
        image_part = shape.part.related_part(shape._element.blip_rId)
        try:
            with open(op["image"], "rb") as f:
                image_part._blob = f.read()
        except OSError as e:
            raise WorkerError("FILE_OPEN", f"Cannot read image {op['image']}: {e}", "Check the image path exists and is readable.")
        return {"op": kind, "target": op["target"]}
    raise WorkerError(
        "UNKNOWN_OP",
        f"Unknown pptx op: {kind}",
        "Valid ops: set_shape_text, set_notes, insert_slide, duplicate_slide, delete_slide, move_slide, replace_image.",
    )


def main(payload):
    path = payload["file"]
    try:
        prs = Presentation(path)
    except Exception as e:
        raise WorkerError("FILE_OPEN", f"Could not open {path} as .pptx: {e}", "Check the path; the file must be a .pptx (not legacy .ppt).")
    results = [apply_one(prs, op) for op in payload["operations"]]
    tmp = path + ".tmp-opencode-office"
    prs.save(tmp)
    os.replace(tmp, path)
    return {"applied": len(results), "results": results}


run(main)
```

`packages/office-core/src/pptx/edit.ts`:

```ts
import { OfficeError } from "../errors"
import { parseId } from "../ids"
import { backupFile } from "../backup"
import { runWorker } from "../worker"
import type { EditResult } from "../docx/edit"

export type PptxOperation =
  | { op: "set_shape_text"; target: string; anchor: string; text: string }
  | { op: "set_notes"; target: string; text: string }
  | { op: "insert_slide"; after: string; layout: string; title?: string; bullets?: string[] }
  | { op: "duplicate_slide"; target: string }
  | { op: "delete_slide"; target: string }
  | { op: "move_slide"; target: string; index: number }
  | { op: "replace_image"; target: string; image: string }

function assertPptxId(id: string): void {
  const ref = parseId(id)
  if (ref.kind !== "slide" && ref.kind !== "shape")
    throw new OfficeError(
      "BAD_ID",
      `Target ${id} is not a pptx element ID`,
      "pptx targets use s:<n> or s:<n>/sh:<m> — get IDs from office_read output for this file.",
    )
}

export async function editPptx(
  file: string,
  operations: PptxOperation[],
  opts?: { backup?: boolean; cacheDir?: string },
): Promise<EditResult> {
  for (const operation of operations) {
    assertPptxId("target" in operation ? operation.target : operation.after)
  }
  const backup = opts?.backup === false ? undefined : await backupFile(file, opts?.cacheDir)
  const data = await runWorker<Omit<EditResult, "backup">>("pptx_edit.py", { file, operations }, opts)
  return { ...data, backup }
}
```

Add to `packages/office-core/src/index.ts`:

```ts
export * from "./pptx/edit"
```

- [ ] **Step 4: Run `bun test pptx-edit` (10 tests), full `bun test`, `bun run typecheck`**

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: editPptx — anchored shape text, slide structure ops, image swap"
```

---

### Task 5: Full-mode docx read (tracked changes)

**Files:**
- Modify: `packages/office-core/src/python/docx_read.py`
- Modify: `packages/office-core/src/docx/read.ts`
- Test: `packages/office-core/test/docx-read.test.ts` (extend)

**Interfaces:**
- Produces: `readDocx` mode widens to `"outline" | "content" | "full"`. Full behaves like content plus, per paragraph: `tracked_insertions?: string[]` (text inside `w:ins` descendants) and `tracked_deletions?: string[]` (text inside `w:del`/`w:delText`), present only when non-empty. `DocxElement` paragraph variant gains both optional fields. `formatDocxRead` appends indented `tracked insertion: "..."` / `tracked deletion: "..."` lines.

- [ ] **Step 1: Write the failing test**

Append to `packages/office-core/test/docx-read.test.ts` (uses the Task 1 fixture):

```ts
test("full mode surfaces tracked insertions; content mode does not", async () => {
  const EDIT = path.join(FIXTURE_DIR, "edit-report.docx")
  const full = await readDocx(EDIT, "full")
  const tracked = full.elements.find(
    (e) => e.type === "paragraph" && (e.tracked_insertions?.length ?? 0) > 0,
  )!
  expect(tracked.type === "paragraph" && tracked.tracked_insertions).toEqual(["with tracked insertion"])
  expect(formatDocxRead(full)).toContain("tracked insertion:")
  const content = await readDocx(EDIT, "content")
  for (const el of content.elements) {
    expect(el.type === "paragraph" ? (el as { tracked_insertions?: string[] }).tracked_insertions : undefined).toBeUndefined()
  }
})
```

- [ ] **Step 2: Run `bun test docx-read` — the new test FAILS (mode "full" behaves like unknown/typecheck rejects)**

- [ ] **Step 3: Implement**

In `packages/office-core/src/python/docx_read.py`: add `from docx.oxml.ns import qn` and, in the paragraph branch, after building `entry` and before appending, insert:

```python
        if mode == "full":
            ins_texts = [
                "".join(t.text or "" for t in ins.iter(qn("w:t")))
                for ins in el._p.iter(qn("w:ins"))
            ]
            del_texts = [
                "".join(t.text or "" for t in d.iter(qn("w:delText")))
                for d in el._p.iter(qn("w:del"))
            ]
            if ins_texts:
                entry["tracked_insertions"] = ins_texts
            if del_texts:
                entry["tracked_deletions"] = del_texts
```

(The outline filter's condition already treats every non-outline mode like content; verify `mode != "outline"` gates — table text and the heading filter — behave for `"full"`, adjusting comparisons from `== "content"` to `!= "outline"` where present.)

In `packages/office-core/src/docx/read.ts`:
- paragraph variant of `DocxElement` gains `tracked_insertions?: string[]; tracked_deletions?: string[]`
- `readDocx` mode parameter becomes `"outline" | "content" | "full"`
- `formatDocxRead` paragraph branch appends, when present:

```ts
        (el.tracked_insertions ?? []).map((t) => `\n  tracked insertion: ${JSON.stringify(t)}`).join("") +
        (el.tracked_deletions ?? []).map((t) => `\n  tracked deletion: ${JSON.stringify(t)}`).join("")
```

- [ ] **Step 4: Run `bun test docx-read`, full `bun test`, `bun run typecheck`**

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: full-mode docx read surfaces tracked changes"
```
