"""Notes blocks (same JSON as the Word notes) -> a .pptx deck.  No AI call: pure layout.

Each h1 starts a slide (title = heading); following blocks become coloured points.  A slide
that fills up continues on the next one with the same title + "(जारी)".  Figures are redrawn
as native PowerPoint shapes on their own slide.  $chem$ markup becomes runs with real
subscript / superscript (baseline offset), so formulas stay editable text.
"""
from __future__ import annotations

import re

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

from .blocks import colour
from .docxkit import Style

FONT = Style.font
W, H = Emu(12192000), Emu(6858000)  # 16:9
MARGIN = Emu(457200)
TITLE_H = Emu(900000)
BODY_PT = 22
MAX_UNITS = 13  # rough "lines" per slide


def _rgb(hexv):
    return RGBColor.from_string(hexv)


# ---------------------------------------------------------------- markup ---
_TOK = re.compile(r"(\$[^$]+\$|\*\*[^*]+\*\*)")


def _math_runs(s):
    """'H_2O^{2-}' -> [(text, baseline)] with baseline -25000 (sub) / 30000 (sup)."""
    out, i = [], 0
    s = re.sub(r"\\frac\{([^{}]*)\}\{([^{}]*)\}", r"(\1)/(\2)", s)
    s = re.sub(r"\((\w+)\)/\((\w+)\)", r"\1/\2", s)

    def group(j):
        if j < len(s) and s[j] == "{":
            depth, k = 1, j + 1
            while k < len(s) and depth:
                depth += {"{": 1, "}": -1}.get(s[k], 0)
                k += 1
            return s[j + 1:k - 1], k
        return (s[j], j + 1) if j < len(s) else ("", j)

    while i < len(s):
        c = s[i]
        if c in "_^":
            g, i = group(i + 1)
            out.append((g.replace("{", "").replace("}", ""), -25000 if c == "_" else 30000))
        elif c in "{}":
            i += 1
        else:
            out.append((c, 0))
            i += 1
    merged = []
    for t, b in out:  # join neighbours with the same baseline
        if merged and merged[-1][1] == b:
            merged[-1] = (merged[-1][0] + t, b)
        else:
            merged.append((t, b))
    return merged


def add_text(p, text, color=None, bold=False, size=BODY_PT):
    for tok in _TOK.split(text or ""):
        if not tok:
            continue
        if tok.startswith("$") and tok.endswith("$") and len(tok) > 1:
            pieces, b = [(t, base, False) for t, base in _math_runs(tok[1:-1])], bold
        elif tok.startswith("**"):
            pieces, b = [(tok[2:-2], 0, True)], True
        else:
            pieces, b = [(tok, 0, bold)], bold
        for t, base, bb in pieces:
            r = p.add_run()
            r.text = t
            f = r.font
            f.size, f.bold, f.name = Pt(size), bb or bold, FONT
            rpr = r._r.get_or_add_rPr()
            for tag in ("a:cs", "a:ea"):
                el = rpr.find(qn(tag))
                if el is None:
                    el = rpr.makeelement(qn(tag), {})
                    rpr.append(el)
                el.set("typeface", FONT)
            if base:
                rpr.set("baseline", str(base))
            if color:
                f.color.rgb = _rgb(color)
    return p


# ------------------------------------------------------------------ deck ---
class Deck:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = W, H
        self.blank = self.prs.slide_layouts[6]
        self.title = ""
        self.body = None
        self.units = 0

    def _slide(self):
        s = self.prs.slides.add_slide(self.blank)
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, Emu(90000))
        bar.fill.solid(); bar.fill.fore_color.rgb = _rgb(Style.red); bar.line.fill.background()
        return s

    def title_slide(self, text):
        s = self._slide()
        tb = s.shapes.add_textbox(MARGIN, Emu(2300000), W - 2 * MARGIN, Emu(1600000))
        tf = tb.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        add_text(p, text, Style.red, True, 44)
        self.body = None

    def new_slide(self, title, cont=False):
        self.title = title
        s = self._slide()
        tb = s.shapes.add_textbox(MARGIN, Emu(200000), W - 2 * MARGIN, TITLE_H)
        tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        add_text(tf.paragraphs[0], title + (" (जारी)" if cont else ""), Style.red, True, 32)
        body = s.shapes.add_textbox(MARGIN, Emu(200000) + TITLE_H, W - 2 * MARGIN,
                                    H - TITLE_H - Emu(500000))
        self.body = body.text_frame
        self.body.word_wrap = True
        self.units = 0
        self._first = True

    def _para(self, cost):
        if self.body is None or self.units + cost > MAX_UNITS:
            self.new_slide(self.title or " ", cont=self.body is not None)
        self.units += cost
        if self._first:
            self._first = False
            p = self.body.paragraphs[0]
        else:
            p = self.body.add_paragraph()
        p.space_after = Pt(6)
        return p

    @staticmethod
    def _cost(*texts):
        n = sum(len(re.sub(r"[$*{}_^]", "", t or "")) for t in texts)
        return max(1, -(-n // 70))  # ~70 chars per line at 22pt

    def point(self, text, marker=None, color=None, level=0, label=None, label_color=None):
        p = self._para(self._cost(text, label or ""))
        p.level = min(level, 4)
        if marker:
            add_text(p, marker + "  ", Style.blue, True)
        if label:
            add_text(p, label + " ", label_color or Style.blue, True)
        add_text(p, text, color)

    def figure(self, b):
        s = self._slide()
        cap = b.get("caption") or self.title
        tb = s.shapes.add_textbox(MARGIN, Emu(200000), W - 2 * MARGIN, TITLE_H)
        add_text(tb.text_frame.paragraphs[0], cap, Style.blue, True, 26)
        fw = max(20.0, float(b.get("width") or 250)); fh = max(20.0, float(b.get("height") or 150))
        avail_w, avail_h = W - 2 * MARGIN, H - TITLE_H - Emu(500000)
        k = min(avail_w / fw, avail_h / fh)  # EMU per figure point
        ox = MARGIN + (avail_w - fw * k) / 2
        oy = Emu(200000) + TITLE_H + (avail_h - fh * k) / 2
        E = lambda v: Emu(int(v))
        for sh in b.get("shapes") or []:
            if not isinstance(sh, dict):
                continue
            kind = sh.get("shape") or sh.get("type")
            col = colour(sh.get("color"), "000000")
            try:
                num = lambda key, d=0.0: float(sh.get(key, d) or d)
                if kind in ("rect", "text", "circle"):
                    if kind == "circle":
                        r = max(1.0, num("r", 5))
                        x, y, w, h = num("cx") - r, num("cy") - r, 2 * r, 2 * r
                        shp = s.shapes.add_shape(MSO_SHAPE.OVAL, E(ox + x * k), E(oy + y * k), E(w * k), E(h * k))
                    else:
                        x, y, w, h = num("x"), num("y"), num("w", 40), num("h", 20)
                        shp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(ox + x * k), E(oy + y * k), E(w * k), E(h * k))
                    fill = colour(sh.get("fill"))
                    if fill:
                        shp.fill.solid(); shp.fill.fore_color.rgb = _rgb(fill)
                    else:
                        shp.fill.background()
                    if kind == "text":
                        shp.line.fill.background()
                    else:
                        shp.line.color.rgb = _rgb(col); shp.line.width = Pt(1.5)
                    if sh.get("text"):
                        tf = shp.text_frame; tf.word_wrap = True
                        for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
                            setattr(tf, m, Emu(20000))
                        lines = str(sh["text"]).split("\n")
                        size = max(10, min(40, num("size", 10) * k / 12700 * 0.9))
                        tcol = col if kind == "text" else "000000"
                        for n, line in enumerate(lines):
                            p = tf.paragraphs[0] if n == 0 else tf.add_paragraph()
                            p.alignment = PP_ALIGN.CENTER
                            add_text(p, line, tcol, bool(sh.get("bold")), size)
                elif kind in ("line", "arrow"):
                    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(ox + num("x1") * k), E(oy + num("y1") * k),
                                               E(ox + num("x2") * k), E(oy + num("y2") * k))
                    c.line.color.rgb = _rgb(col); c.line.width = Pt(1.5)
                    if kind == "arrow":
                        ln = c.line._get_or_add_ln()
                        ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle"}))
            except Exception:
                continue
        self.body = None  # next text goes on a fresh slide

    def save(self, path):
        self.prs.save(path)


def render_deck(parts, out_path, title=None):
    d = Deck()
    current = None
    for part in parts:
        t = (title or part.get("chapter_title") or "").strip()
        if t and t != current:
            d.title_slide(t)
            d.title = t
            current = t
        for b in part.get("blocks") or []:
            if not isinstance(b, dict):
                continue
            ty = b.get("type")
            try:
                if ty == "h1":
                    d.new_slide(str(b.get("text") or ""))
                elif ty == "h2":
                    d.point(f"**{b.get('text') or ''}**", color=Style.blue)
                elif ty == "item":
                    d.point(str(b.get("text") or ""), marker=str(b.get("marker") or "•"),
                            level=int(b.get("level") or 0), color=colour(b.get("color")))
                elif ty == "example":
                    d.point(str(b.get("text") or ""), label="उदाहरण –", color=Style.green)
                elif ty == "note":
                    d.point(str(b.get("text") or ""), label=f"[{b.get('label') or 'NOTE'}]",
                            label_color=Style.red, color=Style.green)
                elif ty == "qa":
                    d.point(str(b.get("q") or ""), label="प्रश्न –", label_color=Style.red, color=Style.red)
                    d.point(str(b.get("a") or ""), label="उत्तर –")
                elif ty == "compare":
                    for k, pr in enumerate(b.get("pairs") or [], 1):
                        if isinstance(pr, (list, tuple)) and len(pr) >= 2:
                            d.point(str(pr[0]), marker=f"({k})", label=f"{b.get('a_title', '')} :", color=Style.blue)
                            d.point(str(pr[1]), level=1, label=f"{b.get('b_title', '')} :",
                                    label_color=Style.green, color=Style.green)
                elif ty == "figure":
                    d.figure(b)
                elif ty == "title":
                    pass
                else:
                    if str(b.get("text") or "").strip():
                        d.point(str(b["text"]), color=colour(b.get("color")))
            except Exception:
                continue
    d.save(out_path)
    return out_path
