from _worker import run, WorkerError
from docx import Document
from docx.oxml.ns import qn
from _docx_common import iter_blocks, render_table


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
            if mode == "outline" and not target and not style.startswith("Heading"):
                continue
            entry = {"id": el_id, "type": "paragraph", "style": style, "text": el.text}
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
            elements.append(entry)
        else:
            entry = {"id": el_id, "type": "table", "rows": len(el.rows), "cols": len(el.columns)}
            if mode != "outline":
                entry["text"] = render_table(el)
            elements.append(entry)
    if target and not elements:
        raise WorkerError("TARGET_NOT_FOUND", f"No element {target} in {path}", "IDs come from office_read output and shift after edits; re-read the file to refresh them.")
    result = {"format": "docx", "mode": mode, "elements": elements}
    if mode == "full":
        comments = [{"id": c.comment_id, "author": c.author, "text": c.text} for c in list(doc.comments)]
        if comments:
            result["comments"] = comments
    return result


run(main)
