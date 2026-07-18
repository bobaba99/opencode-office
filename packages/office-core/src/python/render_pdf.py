"""Rasterize PDF pages to PNGs via pymupdf.

Payload: {pdf, outDir, pages?} — pages is a list of 1-based page numbers,
defaulting to every page in the document. Writes page-<n>.png at 144 dpi
into outDir and returns {pages: [{page, path, width, height}]}.
"""
import os

import fitz
from _worker import run, WorkerError

DPI = 144


def main(payload):
    pdf_path = payload["pdf"]
    out_dir = payload["outDir"]
    pages = payload.get("pages")

    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        raise WorkerError("FILE_OPEN", f"Could not open {pdf_path} as a PDF: {e}", "Check the path; the file must be a PDF produced by the render pipeline.")

    page_count = doc.page_count
    targets = pages if pages is not None else list(range(1, page_count + 1))

    for p in targets:
        if p < 1 or p > page_count:
            raise WorkerError(
                "RENDER_FAILED",
                f"Requested page {p} is out of range for {pdf_path}",
                f"The document has {page_count} page(s); request a page number between 1 and {page_count}.",
            )

    os.makedirs(out_dir, exist_ok=True)

    results = []
    for p in targets:
        page = doc.load_page(p - 1)
        pix = page.get_pixmap(dpi=DPI)
        out_path = os.path.join(out_dir, f"page-{p}.png")
        pix.save(out_path)
        results.append({"page": p, "path": out_path, "width": pix.width, "height": pix.height})

    return {"pages": results}


run(main)
