"""PDF -> one JPEG per page (panel PDFs are scanned slides with no text layer)."""
from __future__ import annotations

import io

import pymupdf
from PIL import Image, ImageChops, ImageStat

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


def drop_repeats(images: list[bytes], threshold: float = 3.5) -> list[bytes]:
    """Panel PDFs often repeat a slide while the teacher writes on it.  Of consecutive
    near-identical pages keep only the last (most complete) one."""
    if len(images) < 2:
        return list(images)
    thumbs = [Image.open(io.BytesIO(b)).convert("L").resize((96, 54)) for b in images]
    keep = []
    for i, img in enumerate(images):
        nxt = thumbs[i + 1] if i + 1 < len(images) else None
        same_as_next = nxt is not None and ImageStat.Stat(ImageChops.difference(thumbs[i], nxt)).mean[0] < threshold
        if not same_as_next:
            keep.append(img)
    return keep
