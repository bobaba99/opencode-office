---
name: office-tools
description: Use when reading, editing, creating, or rendering Word (.docx) or PowerPoint (.pptx) files with the office_read/office_edit/office_create/office_render/office_python tools — covers outline-first navigation, anchored edits, and visual verification.
---

# Office Tools

Five tools operate on `.docx` and `.pptx` files: `office_read`, `office_edit`, `office_create`, `office_render`, `office_python`. They share one contract: **element IDs and anchor text always come from a real `office_read` call on the current state of the file** — never invent, remember from an earlier turn, or guess them.

## Workflow

1. **Outline first.** Call `office_read` with `mode: "outline"` to get the document/deck skeleton (headings for docx, slide titles for pptx) and its element IDs. Don't dump full content up front — it wastes context and you usually don't need most of it.
2. **Content on demand.** Once you know which element(s) matter, call `office_read` again with `mode: "content"` (optionally `target: <id>`) to see full text and the exact wording you'll need as an anchor.
3. **Edit with copied IDs and anchors.** Every `office_edit` operation that takes a `target`/`after` ID or an `anchor` string must use the literal text/ID from that most recent `office_read` output — not paraphrased, not re-typed from memory.
4. **Batches shift IDs.** Positional IDs (`p:3`, `s:2`, etc.) are recomputed after each structural operation (insert, delete, duplicate, move) *within the same batch*. If a batch mixes structural ops with ops targeting elements after the structural change, either order the structural ops **last** in the batch, or split into multiple `office_edit` calls and re-read between them to get fresh IDs.
5. **Anchors must be unique in the element.** If `office_edit` returns `AMBIGUOUS_ANCHOR`, the anchor text matched more than once inside that element — extend it with a bit more surrounding text (from the `office_read` content) until it's unique, then retry.
6. **Follow every error hint.** Each `OfficeError` carries a `hint` field with the specific recovery step (e.g. re-read for a fresh ID, widen the anchor, check the extension). Do what the hint says before trying something else — it was written for exactly this failure.
7. **Creating vs. editing.** `office_create` is for files that don't exist yet. If it fails with `FILE_EXISTS`, the file is already there — switch to `office_read` + `office_edit` on it instead of retrying create.
8. **Render to verify.** After a `office_create` or a non-trivial `office_edit` batch, call `office_render` (requires LibreOffice) and look at the returned PNG attachments to visually confirm the change landed the way you intended, especially for layout-sensitive changes (tables, slide structure, images). If LibreOffice isn't installed, `office_render` fails with `RENDER_UNAVAILABLE` — reads and edits still work without it, just skip the visual check.
9. **`office_python` is the last resort.** Only reach for it when the typed tools genuinely cannot express the change (e.g. a python-docx/python-pptx feature with no equivalent op). It runs arbitrary Python with `python-docx`, `python-pptx`, `pillow`, and `pymupdf` preloaded in the managed venv, but it bypasses the anchor-safety and atomic-batch guarantees the typed ops give you — prefer `office_edit` whenever an op exists for what you need.

## Operations catalog

Every `office_edit` operation object needs `op` plus the fields below. `anchor` means: copy the exact current text for that element from `office_read` output. `target`/`after` means: copy the exact ID for that element from `office_read` output (`after` places the new/inserted content immediately following that ID; `target` names the element being acted on).

### docx ops

| op | required fields | semantics |
|---|---|---|
| `replace_text` | `target`, `anchor`, `text` | Replace the `anchor` text inside paragraph/table element `target` with `text`. |
| `insert_content` | `after`, `markdown` | Insert new paragraph(s)/table(s) parsed from `markdown` immediately after element `after`. |
| `delete_element` | `target`, `anchor` | Delete the paragraph/table `target`, confirming it still contains `anchor` first. |
| `set_style` | `target`, `anchor`, `style` | Apply paragraph style `style` (e.g. `Heading 1`) to `target`, confirming it still contains `anchor` first. |
| `set_table_cell` | `target`, `row`, `col`, `text`, `anchor?` | Set cell `(row, col)` of table `target` to `text`; optional `anchor` confirms the cell's current text first. |

### pptx ops

| op | required fields | semantics |
|---|---|---|
| `set_shape_text` | `target`, `anchor`, `text` | Replace the `anchor` text inside shape `target` (`s:<n>/sh:<n>`) with `text`. |
| `set_notes` | `target`, `text` | Replace the speaker notes of slide `target` (`s:<n>`) with `text`. |
| `insert_slide` | `after`, `layout`, `title?`, `bullets?` | Insert a new slide using `layout` immediately after slide `after`, with optional `title` and `bullets`. |
| `duplicate_slide` | `target` | Duplicate slide `target`, inserting the copy immediately after it. |
| `delete_slide` | `target` | Delete slide `target`. |
| `move_slide` | `target`, `index` | Move slide `target` to 0-indexed position `index` in the deck. |
| `replace_image` | `target`, `image` | Replace the image in picture shape `target` (`s:<n>/sh:<n>`) with the file at path `image`. |

## Tool reference

- **office_read** `{ file, mode?, target? }` — `mode` defaults to `"outline"`. `"full"` is docx-only (comments/tracked changes); on `.pptx` it's coerced to `"content"`.
- **office_edit** `{ file, operations }` — `operations` is a non-empty array; each op's name must match the file's format (docx ops on a `.pptx`, or vice versa, fails with `WRONG_OPS_FORMAT`). Empty or missing anchors fail with `BAD_ANCHOR` before anything is touched.
- **office_create** `{ file, markdown? | slides?, reference?, template? }` — `.docx` needs `markdown`; `.pptx` needs `slides`.
- **office_render** `{ file, pages? }` — omit `pages` for all pages, pass `[]` for none. Returns PNG attachments, one per rendered page/slide.
- **office_python** `{ code, files? }` — runs `code` via `python -c`; list the paths it touches in `files` so the permission prompt is specific (omit for a blanket prompt).
