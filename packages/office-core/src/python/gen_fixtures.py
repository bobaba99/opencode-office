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
