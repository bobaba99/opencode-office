import os

from _worker import run, WorkerError
from _docx_common import styled_line, add_styled_paragraph
from docx import Document
from docx.oxml.ns import qn


def main(payload):
    path = payload["file"]
    markdown = payload["markdown"]
    reference = payload.get("reference")

    if os.path.exists(path):
        raise WorkerError(
            "FILE_EXISTS",
            f"{path} already exists",
            "office_create makes new files; use office_edit to change existing ones, or pass a different path.",
        )

    if reference:
        try:
            doc = Document(reference)
        except Exception as e:
            raise WorkerError(
                "FILE_OPEN",
                f"Could not open {reference} as .docx: {e}",
                "Check the path; the file must be a .docx (not legacy .doc).",
            )
        for child in list(doc.element.body):
            if child.tag != qn("w:sectPr"):
                doc.element.body.remove(child)
    else:
        doc = Document()

    count = 0
    for line in [line for line in markdown.splitlines() if line.strip()]:
        text, style = styled_line(line)
        add_styled_paragraph(doc, text, style)
        count += 1

    tmp = path + ".tmp-opencode-office"
    try:
        doc.save(tmp)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return {"file": path, "paragraphs": count}


run(main)
