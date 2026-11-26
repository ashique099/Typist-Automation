"""Panel PDFs -> one combined, house-style .docx.

    python -m typist.convert part1.pdf part2.pdf -o Notes.docx [--title "..."] [--save-json blocks.json]
    python -m typist.convert --from-json blocks.json -o Notes.docx      # re-render without the API

PDFs are combined in the order given.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from .blocks import render
from .docxkit import NoteDoc
from .pages import pdf_pages


def transcribe_all(pdfs, progress=None):
    """pdfs: list of (name, path_or_bytes). Returns list of {"name","chapter_title","blocks"}."""
    from .transcribe import transcribe_pdf_pages  # needs the anthropic package + API key

    parts, ctx = [], []
    for n, (name, src) in enumerate(pdfs, 1):
        if progress:
            progress(f"PDF {n}/{len(pdfs)} ({name}): pages nikal rahe hain")
        images = pdf_pages(src)

        def page_progress(i, total, n=n, name=name):
            if progress:
                progress(f"PDF {n}/{len(pdfs)} ({name}): slides {i + 1}-{min(total, i + 10)} of {total} "
                         f"type ho rahi hain")

        title, blocks = transcribe_pdf_pages(images, ctx, page_progress, label=f"(from {name})")
        parts.append({"name": name, "chapter_title": title, "blocks": blocks})
        ctx += blocks
    return parts


def build_docx(parts, out_path, title=None):
    d = NoteDoc()
    current = None
    for p in parts:
        t = (title or p.get("chapter_title") or "").strip()
        if t and t != current:
            d.title(t)
            current = t
        render(p["blocks"], d)
    d.save(out_path)
    return out_path


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdfs", nargs="*")
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("--title")
    ap.add_argument("--save-json")
    ap.add_argument("--from-json")
    a = ap.parse_args(argv)
    if a.from_json:
        parts = json.loads(Path(a.from_json).read_text(encoding="utf-8"))
    else:
        if not a.pdfs:
            ap.error("give PDFs or --from-json")
        parts = transcribe_all([(Path(p).name, p) for p in a.pdfs], progress=print)
        if a.save_json:
            Path(a.save_json).write_text(json.dumps(parts, ensure_ascii=False, indent=1), encoding="utf-8")
    build_docx(parts, a.out, a.title)
    print("wrote", a.out)


if __name__ == "__main__":
    main()
