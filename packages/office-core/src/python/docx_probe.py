"""Test-support probe: report run text and formatting for a target docx paragraph.

Uses flat_runs (hyperlink-aware) so tests can verify formatting survives an
edit even when the paragraph contains hyperlink runs.
"""
from _worker import run, WorkerError
from _docx_common import flat_runs, iter_blocks
from docx import Document
from docx.oxml.ns import qn


def main(payload):
    path = payload["file"]
    target = payload["target"]
    try:
        doc = Document(path)
    except Exception as e:
        raise WorkerError("FILE_OPEN", f"Could not open {path} as .docx: {e}", "Check the path; the file must be a .docx (not legacy .doc).")

    for prefix, index, el in iter_blocks(doc):
        if f"{prefix}:{index}" != target:
            continue
        if prefix != "p":
            raise WorkerError(
                "BAD_TARGET_KIND",
                f"docx_probe cannot target {target} (a {prefix} element)",
                "docx_probe only supports paragraph targets.",
            )
        runs = [{"text": r.text or "", "bold": r.bold is True} for r in flat_runs(el)]
        comment_refs = len(list(el._p.iter(qn("w:commentReference"))))
        return {"runs": runs, "comment_refs": comment_refs}

    raise WorkerError("TARGET_NOT_FOUND", f"No element {target} in {path}", "IDs come from office_read output and shift after edits; re-read the file to refresh them.")


run(main)
