from _worker import WorkerError
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.hyperlink import Hyperlink
from docx.text.paragraph import Paragraph
from docx.text.run import Run

MD_STYLES = [("### ", "Heading 3"), ("## ", "Heading 2"), ("# ", "Heading 1"), ("- ", "List Bullet")]


def iter_blocks(doc):
    for i, child in enumerate(doc.element.body.iterchildren()):
        if child.tag == qn("w:p"):
            yield "p", i, Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield "tbl", i, Table(child, doc)


def render_table(table):
    return "\n".join(" | ".join(cell.text.strip() for cell in row.cells) for row in table.rows)


def flat_runs(p):
    """Ordered run list for a paragraph, including runs inside hyperlinks.

    Paragraph.runs excludes hyperlink runs, but Paragraph.text includes
    their text. This flattens p.iter_inner_content() so anchor-matching and
    run-preserving edits see the same text python-docx reports via .text.
    """
    runs = []
    for item in p.iter_inner_content():
        if isinstance(item, Run):
            runs.append(item)
        elif isinstance(item, Hyperlink):
            runs.extend(item.runs)
    return runs


def docx_para_text(p):
    return "".join(run.text or "" for run in flat_runs(p))


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
