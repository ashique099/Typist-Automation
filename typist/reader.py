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


def _lines(p_el):
    """Return ([(text with markup, first colour)] one per visual line, figures) for one w:p.
    A manual line break (w:br) starts a new line: teams often type all four options of an
    MCQ in one paragraph, separated by Shift+Enter."""
    lines, cur, colors, figs = [], [], [], []

    def newline():
        nonlocal cur, colors
        lines.append((cur, colors))
        cur, colors = [], []

    for el in p_el.iter():
        if _in_textbox(el, p_el):
            continue
        if el.tag == f"{{{W}}}r" and el.getparent().tag != _m("r"):
            bold = _bold(el)
            for x in el:
                if x.tag == f"{{{W}}}t" or x.tag == f"{{{W}}}tab":
                    t = (x.text or "") if x.tag == f"{{{W}}}t" else "\t"
                    if t.strip():
                        colors.append(_color(el))
                        cur.append(f"**{t}**" if bold and len(t.strip()) > 1 else t)
                    elif t:
                        cur.append(t)
                elif x.tag == f"{{{W}}}br" and x.get(qn("w:type")) in (None, "textWrapping"):
                    newline()
            for d in el.iter(f"{{{WP}}}docPr"):
                descr = d.get("descr") or ""
                if descr.startswith(FIG_TAG):
                    try:
                        figs.append(json.loads(descr[len(FIG_TAG):]))
                    except ValueError:
                        pass
        elif el.tag == _m("oMath") and not any(a.tag == _m("oMath") for a in el.iterancestors()):
            cur.append("$" + _omml(el) + "$")
    newline()
    out = []
    for parts, cols in lines:
        text = re.sub(r"\*\*\s*\*\*", "", "".join(parts))
        text = re.sub(r"\*\*([^*]*?)(\s*)\*\*", lambda m: f"**{m.group(1)}**{m.group(2)}" if m.group(1) else m.group(2), text)
        if text.count("**") % 2:  # unbalanced bold marker -> drop the last one
            i = text.rfind("**")
            text = text[:i] + text[i + 2:]
        out.append((text.strip(), cols[0] if cols else ""))
    return out, figs


def _para(p_el):
    """Return (text with markup, first colour, figures) for one w:p (lines joined)."""
    lines, figs = _lines(p_el)
    text = " ".join(t for t, _ in lines if t)
    col = next((c for t, c in lines if t and c), "")
    return text, col, figs


_OPT = re.compile(r"\(\s*([A-Da-d])\s*\)")


def _split_options(text):
    """'(A) x(B) y(C) z' typed on one line -> ['(A) x', '(B) y', '(C) z'].  Only splits at
    markers that come in order (A, B, C, D), so '(C) दोनों (A) तथा (B)' stays whole."""
    plain = text.replace("**", "")
    found = [(m.start(), m.group(1)) for m in _OPT.finditer(plain)]
    if len(found) < 2:
        return [text]
    seq, want = [], None
    for pos, letter in found:
        if want is None:
            if pos <= 2:  # line must start with the first marker
                seq.append(pos)
                want = chr(ord(letter) + 1)
            else:
                return [text]
        elif letter == want:
            seq.append(pos)
            want = chr(ord(letter) + 1)
    if len(seq) < 2:
        return [text]
    seq.append(len(plain))
    return [plain[a:b].strip() for a, b in zip(seq, seq[1:]) if plain[a:b].strip()]


def _plain(s):
    return re.sub(r"\*\*|\$", "", s).strip()


def _mcq_table(tbl):
    """The team's test format: one 2-column table per question
    (Question / Type / Option x4 / Answer / Solution / marks).  -> {"type": "mcq", ...}"""
    rows = []
    for tr in tbl.iter(f"{{{W}}}tr"):
        cells = [c for c in tr if c.tag == f"{{{W}}}tc"]
        if len(cells) < 2:
            continue
        key = _plain(" ".join(_para(p)[0] for p in cells[0].iter(f"{{{W}}}p"))).lower()
        val = " ".join(t for t in (_para(p)[0] for p in cells[1].iter(f"{{{W}}}p")) if t)
        rows.append((key, val))
    if not rows or rows[0][0] != "question":
        return None
    get = lambda k: next((v for kk, v in rows if kk == k), "")
    return {"type": "mcq", "q": get("question"), "options": [v for k, v in rows if k == "option"],
            "answer": get("answer"), "solution": get("solution")}


def docx_blocks(src) -> tuple[str, list[dict]]:
    """src: path or bytes.  Returns (chapter_title, blocks)."""
    doc = Document(io.BytesIO(src) if isinstance(src, (bytes, bytearray)) else src)
    S = Style
    title, blocks, last_caption_fig = "", [], None
    mcq_tables = {}
    for tbl in doc.element.body.iter(f"{{{W}}}tbl"):
        m = _mcq_table(tbl)
        if m:
            mcq_tables[tbl] = m
    for p_el in doc.element.body.iter(f"{{{W}}}p"):
        if _in_textbox(p_el):
            continue  # text inside a figure's shapes
        tbl = next((a for a in p_el.iterancestors(f"{{{W}}}tbl")), None)
        if tbl is not None and tbl in mcq_tables:
            if mcq_tables[tbl] is not None:  # first paragraph of an MCQ table -> one block
                blocks.append(mcq_tables[tbl])
                mcq_tables[tbl] = None
            continue
        lines, figs = _lines(p_el)
        for f in figs:
            blocks.append(f)
            last_caption_fig = f
        todo = []
        for t, c in lines:
            for piece in _split_options(t):
                if piece:
                    todo.append((piece, c))
        for text, col in todo:
            last_caption_fig = _classify(p_el, text, col, blocks, last_caption_fig)
            if blocks and blocks[-1].get("type") == "title" and not title:
                title = blocks[-1]["text"]
    blocks = _group_mcqs(blocks)
    if not title:
        cover = " ".join(dict.fromkeys(_plain(_para(p)[0]) for p in doc.element.body.iter(f"{{{W}}}p")
                                       if _in_textbox(p) and _plain(_para(p)[0])))
        title = next((b["text"] for b in blocks if b["type"] == "h1" and not _QLINE.match(b["text"])), "") or cover
    return title, blocks


_QLINE = re.compile(r"^\s*(?:Q\.?\s*)?\d{1,3}\s*[.)]\s*")


def _clean_option(letter, text):
    t = text.replace("**", "").strip()
    t = re.sub(r"^\(?\s*" + re.escape(letter) + r"\s*\)\s*", "", t)  # marker typed twice
    return t


def _group_mcqs(blocks):
    """Numbered question + options (A)-(D) -> {"type": "mcq", "q", "options", "letters"}."""
    out, i = [], 0
    is_opt = lambda b: b.get("type") == "item" and bool(re.fullmatch(r"\(?\s*[A-Da-d]\s*\)", str(b.get("marker") or "")))
    while i < len(blocks):
        b = blocks[i]
        q = None
        if b.get("type") in ("h1", "h2", "para") and _QLINE.match(_plain(b.get("text", ""))):
            q = b.get("text", "")
        elif b.get("type") == "item" and re.fullmatch(r"(?:Q\.?\s*)?\d{1,3}[.)]", str(b.get("marker") or "")):
            q = f"{b['marker']} {b.get('text', '')}"
        j = i + 1
        extra = []  # "[BB2022]" notes or a "या, ..." second wording between question and options
        while q is not None and j < len(blocks) and not is_opt(blocks[j]) and len(extra) < 2:
            nb = blocks[j]
            if nb.get("type") == "note":
                extra.append(f"[{nb.get('label', '')}] {nb.get('text', '')}".strip())
            elif nb.get("type") in ("h1", "h2", "para") and _plain(nb.get("text", "")).startswith("या"):
                extra.append(nb.get("text", ""))
            else:
                break
            j += 1
        k = j
        while j < len(blocks) and is_opt(blocks[j]):
            j += 1
        if q is not None and j - k >= 2:
            q = "\n".join([q] + extra)
            opts = blocks[k:j]
            letters = [str(o["marker"]).strip("() ") for o in opts]
            out.append({"type": "mcq", "q": q.replace("**", "").strip(), "letters": letters,
                        "options": [_clean_option(l, str(o.get("text") or "")) for l, o in zip(letters, opts)]})
            i = j
            continue
        out.append(b)
        i += 1
    return out


def _classify(p_el, text, col, blocks, last_caption_fig):
    """One line of a w:p -> a block appended to blocks.  Returns the new last_caption_fig."""
    S = Style
    style = ""
    ps = p_el.find(f"{{{W}}}pPr/{{{W}}}pStyle")
    if ps is not None:
        style = (ps.get(qn("w:val")) or "").lower()
    plain = _plain(text)
    if style == "caption" and last_caption_fig is not None:
        return last_caption_fig  # caption already inside the figure block
    is_option = bool(re.match(r"^\(\s*[A-Da-d]\s*\)", plain))
    if style == "title":
        blocks.append({"type": "title", "text": plain})
    elif not is_option and (style in ("heading1", "heading 1") or (
            col == S.red and len(plain) < 120 and not plain.startswith(("[", "प्रश्न")))):
        blocks.append({"type": "h1", "text": text.replace("**", "").strip()})
    elif not is_option and (style in ("heading2", "heading 2") or (
            col == S.blue and len(plain) < 120 and "\t" not in text and not plain.startswith(("उदाहरण", "उत्तर")))):
        blocks.append({"type": "h2", "text": text.replace("**", "").strip()})
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
        if _plain(rest).startswith(mk) or _plain(rest) == _plain(text):  # marker split by bold runs
            rest = re.sub(r"^\s*" + re.escape(mk) + r"\s*", "", rest.replace("**", ""), count=1)
        blocks.append({"type": "item", "marker": mk, "text": rest})
    else:
        b = {"type": "para", "text": text.replace("\t", " ")}
        if col == S.green:
            b["color"] = "green"
        blocks.append(b)
    return None


def _sp_paragraphs(sp):
    """Paragraph texts of one p:sp, with PowerPoint equations (a14:m) as $..$ markup."""
    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
    out = []
    for p in sp.iter(f"{{{A}}}p"):
        bits = []
        for el in p:
            if el.tag == f"{{{A}}}r":
                bits.append("".join(t.text or "" for t in el.iter(f"{{{A}}}t")))
            elif el.tag == f"{{{A14}}}m":
                om = next(el.iter(_m("oMath")), None)
                if om is not None:
                    bits.append("$" + _omml(om) + "$")
        out.append("".join(bits).strip())
    return out


def pptx_blocks(src) -> tuple[str, list[dict]]:
    """Text of every slide, in order: slide title -> h1, other text -> paras.  Shapes inside
    mc:AlternateContent (equations) are read from their Choice part."""
    from pptx import Presentation
    P = "http://schemas.openxmlformats.org/presentationml/2006/main"
    MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    prs = Presentation(io.BytesIO(src) if isinstance(src, (bytes, bytearray)) else src)
    blocks, last_head = [], None
    for slide in prs.slides:
        found = []
        for sp in slide.shapes._spTree.iter(f"{{{P}}}sp"):
            if any(a.tag == f"{{{MC}}}Fallback" for a in sp.iterancestors()):
                continue
            paras = [t for t in _sp_paragraphs(sp) if t]
            if not paras:
                continue
            off = next(sp.iter(f"{{{A}}}off"), None)
            pos = (int(off.get("y", 0)), int(off.get("x", 0))) if off is not None else (0, 0)
            found.append((pos, paras))
        found.sort(key=lambda f: f[0])
        for i, (_, paras) in enumerate(found):
            whole = "\n".join(paras)
            for t in paras:
                head = t.endswith(":-") or (len(found) > 1 and len(t) < 80 and t == whole)
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
