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
