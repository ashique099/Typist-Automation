"""PDF -> one JPEG per page (panel PDFs are scanned slides with no text layer).

PyMuPDF renders any PDF.  If its DLL cannot load (Windows Smart App Control blocked it
on the office PC), a pure-Python fallback (pypdf) takes the slide picture embedded in
each page, which is all a panel PDF contains.
"""
from __future__ import annotations

import io
import re

from PIL import Image, ImageChops, ImageStat

try:
    import pymupdf
    PDF_ENGINE = "pymupdf"
except Exception as e:  # ImportError, or OSError when Windows blocks the DLL
    pymupdf = None
    PDF_ENGINE = f"pypdf (PyMuPDF unavailable: {str(e)[:120]})"

MAX_EDGE = 1600  # px; enough to read handwriting, small enough to keep requests cheap


def pdf_pages(path_or_bytes) -> list[bytes]:
    if pymupdf is None:
        return _pdf_pages_pypdf(path_or_bytes)
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


def _pdf_pages_pypdf(path_or_bytes) -> list[bytes]:
    from pypdf import PdfReader
    src = io.BytesIO(bytes(path_or_bytes)) if isinstance(path_or_bytes, (bytes, bytearray)) else path_or_bytes
    out = []
    for n, page in enumerate(PdfReader(src).pages, 1):
        best = None
        # pages can share one resource list, so only consider pictures this page draws
        try:
            drawn = set(re.findall(rb"/([^\s/<>\[\]()]+)\s+Do\b", page.get_contents().get_data()))
        except Exception:
            drawn = set()
        for img in page.images:  # the slide is the biggest picture on the page
            if drawn and img.name.rsplit(".", 1)[0].encode() not in drawn:
                continue
            try:
                im = img.image
            except Exception:
                continue
            if best is None or im.width * im.height > best.width * best.height:
                best = im
        if best is None:
            raise ValueError(f"Page {n} has no slide image. This PDF needs PyMuPDF, which Windows "
                             "is blocking on this PC.")
        im = best.convert("RGB")
        im.thumbnail((MAX_EDGE, MAX_EDGE))
        buf = io.BytesIO()
        im.save(buf, "JPEG", quality=85)
        out.append(buf.getvalue())
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
