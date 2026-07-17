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

