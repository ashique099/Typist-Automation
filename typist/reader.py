"""Read a notes .docx (or a .pptx) back into blocks, with no API call.

Used by "DOCX to PPT" and "PPT/DOCX to MCQ Test".  Works best on notes made by this
tool (house colours: red topics, blue sub-topics, green examples), but any Word file
gives usable blocks: colour/heading style decides the block type, equations come back
as $..$ markup and our vector figures are restored from their alt text.
"""
from __future__ import annotations

import io
import json
import re

from docx import Document
from docx.oxml.ns import qn

from .docxkit import FIG_TAG, Style

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_MARKER = re.compile(r"^\s*(\(?(?:[ivxlc]+|\d{1,2}|[a-zA-Z]|[क-ह])[.)]|[•●▪◆➤→\-–])\s*")


def _m(tag):
    return f"{{{M}}}{tag}"


def _omml(el) -> str:
    """OMML -> the $..$ markup docxkit/pptkit understand (sub/sup/frac)."""
    tag = el.tag
    kids = lambda name: next((c for c in el if c.tag == _m(name)), None)
    txt = lambda c: "".join(_omml(x) for x in c) if c is not None else ""
    if tag == _m("t"):
        return el.text or ""
    if tag == _m("sSub"):
        return txt(kids("e")) + "_{" + txt(kids("sub")) + "}"
    if tag == _m("sSup"):
        return txt(kids("e")) + "^{" + txt(kids("sup")) + "}"
    if tag == _m("sSubSup"):
        return txt(kids("e")) + "_{" + txt(kids("sub")) + "}^{" + txt(kids("sup")) + "}"
    if tag == _m("f"):
        return "\\frac{" + txt(kids("num")) + "}{" + txt(kids("den")) + "}"
    if tag in (_m("rPr"), _m("ctrlPr"), _m("fPr"), _m("sSubPr"), _m("sSupPr"), _m("oMathParaPr")):
        return ""
    return "".join(_omml(c) for c in el)


def _color(run_el):
    c = run_el.find(f"{{{W}}}rPr/{{{W}}}color")
    return (c.get(qn("w:val")) or "").upper() if c is not None else ""


def _bold(run_el):
    b = run_el.find(f"{{{W}}}rPr/{{{W}}}b")
    return b is not None and b.get(qn("w:val")) not in ("0", "false")


def _in_textbox(el, stop=None):
    for a in el.iterancestors():
        if a is stop:
            return False
        if a.tag == f"{{{W}}}txbxContent":
            return True
    return False


def _para(p_el):
    """Return (text with markup, first colour, figures) for one w:p."""
    out, colors, figs = [], [], []
    for el in p_el.iter():
        if _in_textbox(el, p_el):
            continue
        if el.tag == f"{{{W}}}r" and el.getparent().tag != _m("r"):
            t = "".join((x.text or "") if x.tag == f"{{{W}}}t" else "\t" if x.tag == f"{{{W}}}tab" else ""
                        for x in el)
            if t.strip():
                colors.append(_color(el))
                out.append(f"**{t}**" if _bold(el) and len(t.strip()) > 1 else t)
            elif t:
                out.append(t)
            for d in el.iter(f"{{{WP}}}docPr"):
                descr = d.get("descr") or ""
                if descr.startswith(FIG_TAG):
                    try:
                        figs.append(json.loads(descr[len(FIG_TAG):]))
                    except ValueError:
                        pass
        elif el.tag == _m("oMath") and not any(a.tag == _m("oMath") for a in el.iterancestors()):
            out.append("$" + _omml(el) + "$")
    text = re.sub(r"\*\*\s*\*\*", "", "".join(out))
    text = re.sub(r"\*\*([^*]*?)(\s*)\*\*", lambda m: f"**{m.group(1)}**{m.group(2)}" if m.group(1) else m.group(2), text)
    return text.strip(), (colors[0] if colors else ""), figs


def _plain(s):
    return re.sub(r"\*\*|\$", "", s).strip()


def docx_blocks(src) -> tuple[str, list[dict]]:
    """src: path or bytes.  Returns (chapter_title, blocks)."""
    doc = Document(io.BytesIO(src) if isinstance(src, (bytes, bytearray)) else src)
    S = Style
    title, blocks, last_caption_fig = "", [], None
    for p_el in doc.element.body.iter(f"{{{W}}}p"):
        if _in_textbox(p_el):
            continue  # text inside a figure's shapes
        text, col, figs = _para(p_el)
        for f in figs:
            blocks.append(f)
            last_caption_fig = f
        if not text:
            continue
        style = ""
        ps = p_el.find(f"{{{W}}}pPr/{{{W}}}pStyle")
        if ps is not None:
            style = (ps.get(qn("w:val")) or "").lower()
        plain = _plain(text)
        if style == "caption" and last_caption_fig is not None:
            continue  # caption already inside the figure block
        last_caption_fig = None
        if style == "title":
            if not title:
                title = plain
            blocks.append({"type": "title", "text": plain})
        elif style in ("heading1", "heading 1") or (col == S.red and len(plain) < 120
                                                    and not plain.startswith(("[", "प्रश्न"))):
            blocks.append({"type": "h1", "text": plain})
        elif style in ("heading2", "heading 2") or (col == S.blue and len(plain) < 120 and "\t" not in text
                                                     and not plain.startswith(("उदाहरण", "उत्तर"))):
            blocks.append({"type": "h2", "text": plain})
        elif plain.startswith("उदाहरण"):
            blocks.append({"type": "example", "text": re.sub(r"^\**उदाहरण\s*[–:-]?\s*\**\s*", "", text)})
        elif re.match(r"^\**\[([^\]]+)\]", text):
            m = re.match(r"^\**\[([^\]]+)\]\**\s*", text)
            blocks.append({"type": "note", "label": m.group(1), "text": text[m.end():]})
        elif plain.startswith("प्रश्न"):
            blocks.append({"type": "qa", "q": re.sub(r"^\**प्रश्न\s*[–:-]?\s*\**\s*", "", text), "a": ""})
        elif plain.startswith("उत्तर") and blocks and blocks[-1].get("type") == "qa" and not blocks[-1]["a"]:
            blocks[-1]["a"] = re.sub(r"^\**उत्तर\s*[–:-]?\s*\**\s*", "", text)
        elif "\t" in text and _MARKER.match(_plain(text.split("\t", 1)[0]) + " "):
            marker, rest = text.split("\t", 1)
            blocks.append({"type": "item", "marker": _plain(marker), "text": rest.strip()})
        elif (m := _MARKER.match(plain)) and len(m.group(1)) <= 5:
            mk = m.group(1).strip()
            rest = re.sub(r"^(\*\*)?\s*" + re.escape(mk) + r"\s*(\*\*)?\s*", "", text, count=1)
            blocks.append({"type": "item", "marker": mk, "text": rest})
        else:
            b = {"type": "para", "text": text.replace("\t", " ")}
            if col == S.green:
                b["color"] = "green"
            blocks.append(b)
    if not title:
        title = next((b["text"] for b in blocks if b["type"] in ("title", "h1")), "")
    return title, blocks


def pptx_blocks(src) -> tuple[str, list[dict]]:
    """Text of every slide, in order: slide title -> h1, other text -> paras."""
    from pptx import Presentation
    prs = Presentation(io.BytesIO(src) if isinstance(src, (bytes, bytearray)) else src)
    blocks, last_head = [], None
    for slide in prs.slides:
        shapes = sorted((s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()),
                        key=lambda s: (s.top or 0, s.left or 0))
        for i, sh in enumerate(shapes):
            for para in sh.text_frame.paragraphs:
                t = "".join(r.text for r in para.runs).strip()
                if not t:
                    continue
                head = t.endswith(":-") or (len(shapes) > 1 and len(t) < 80 and t == sh.text_frame.text.strip())
                if i == 0 and head:
                    t = t.rstrip(":- ")
                    if t != last_head:
                        blocks.append({"type": "h1", "text": t})
                        last_head = t
                else:
                    blocks.append({"type": "para", "text": t})
    title = next((b["text"] for b in blocks if b["type"] == "h1"), "")
    return title, blocks


def read_parts(files) -> list[dict]:
    """files: list of (name, path).  .docx / .pptx -> parts like transcribe_all returns."""
    parts = []
    for name, path in files:
        low = str(name).lower()
        title, blocks = (pptx_blocks(path) if low.endswith(".pptx") else docx_blocks(path))
        parts.append({"name": name, "chapter_title": title, "blocks": blocks})
    return parts
