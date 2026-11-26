"""PDF -> one JPEG per page (panel PDFs are scanned slides with no text layer)."""
from __future__ import annotations

import pymupdf

MAX_EDGE = 1600  # px; enough to read handwriting, small enough to keep requests cheap


def pdf_pages(path_or_bytes) -> list[bytes]:
    if isinstance(path_or_bytes, (bytes, bytearray)):
        doc = pymupdf.open(stream=bytes(path_or_bytes), filetype="pdf")
    else:
        doc = pymupdf.open(path_or_bytes)
    out = []
    with doc:
        for page in doc:
            r = page.rect
            zoom = MAX_EDGE / max(r.width, r.height)
            pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
            out.append(pix.tobytes("jpeg", jpg_quality=85))
    return out
