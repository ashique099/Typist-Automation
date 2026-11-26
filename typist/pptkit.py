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


def render_deck(parts, out_path, title=None):
    """Default deck = Plain PPT (black, Utsaah 120, yellow/green/white).  See plainppt.py."""
    from .plainppt import render_plain_deck
    return render_plain_deck(parts, out_path, title)


# ------------------------------------------------------- template decks ---
# The team's own class PPT (dark background image in the slide master, white text,
# green headings, magenta labels).  Any .pptx can be used as a template: its slides are
# dropped and only the master/layouts (background, header bar, fonts) are kept.
SUB = str.maketrans("0123456789+-=()n", "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₙ")
SUP = str.maketrans("0123456789+-=()n", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿ")


def unicode_formula(s):
    """'H_2O', 'SO_4^{2-}' -> 'H₂O', 'SO₄²⁻' when every script char has a Unicode form, else None."""
    out = []
    for t, base in _math_runs(s):
        if base:
            table = SUB if base < 0 else SUP
            conv = t.translate(table)
            if any(c == o for c, o in zip(conv, t) if not c.isspace()):
                return None
            out.append(conv)
        else:
            out.append(t)
    return "".join(out).replace("−", "-") if out else s


class Palette:
    heading = "00FF00"   # green, bold, ends with " :-"
    label = "FF00FF"     # magenta: उदाहरण:, उत्तर –, ध्यान दें
    accent = "00FFFF"    # cyan: sub-topics, comparison titles
    alert = "FFFF00"     # yellow: notes
    button = "FF0000"    # divider button
    button_shadow = "A00000"
    figure_map = {"000000": "FFFFFF", Style.black: "FFFFFF", Style.blue: "00FFFF", Style.red: "FF5050",
                  Style.green: "00FF00", "404040": "D0D0D0", "1F4E9F": "00FFFF"}


class TemplateDeck:
    def __init__(self, template, logo=None, pal=Palette):
        self.prs = Presentation(template)
        lst = self.prs.slides._sldIdLst
        for el in list(lst):
            self.prs.part.drop_rel(el.rId)
            lst.remove(el)
        self.pal, self.logo = pal, logo
        self.W, self.H = self.prs.slide_width, self.prs.slide_height
        self.k = self.W / 36576000  # sizes below are tuned for the team's 40-inch slides
        self.layout, self.blank = None, None
        for lay in self.prs.slide_layouts:
            phs = [ph for ph in lay.placeholders if ph.placeholder_format.type in (2, 7)]  # BODY / OBJECT
            if phs and len(list(lay.placeholders)) == 1:
                self.layout = lay
            if lay.name == "Blank":
                self.blank = lay
        self.layout = self.layout or self.blank or self.prs.slide_layouts[-1]
        self.blank = self.blank or self.layout
        self.body = None
        self.units = 0
        self.max_units, self.chars = 8, 52
        self.size = 80 * self.k

    # -- slides
    def _content_slide(self):
        s = self.prs.slides.add_slide(self.layout)
        body = next((ph for ph in s.placeholders if ph.placeholder_format.type in (2, 7)), None)
        if body is None:
            body = s.shapes.add_textbox(int(self.W * 0.034), int(self.H * 0.096), int(self.W * 0.94),
                                        int(self.H * 0.8))
        tf = body.text_frame
        tf.word_wrap = True
        self.body, self.units, self._first = tf, 0, True
        return s

    def divider(self, text, big=False):
        s = self.prs.slides.add_slide(self.blank)
        for ph in list(s.placeholders):
            ph._element.getparent().remove(ph._element)
        bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.W, self.H)
        bg.fill.solid(); bg.fill.fore_color.rgb = _rgb("000000"); bg.line.fill.background()
        y = int(self.H * 0.08)
        if self.logo:
            d = int(self.H * 0.38)
            s.shapes.add_picture(self.logo, int((self.W - d) / 2), y, d, d)
            y += d + int(self.H * 0.06)
        else:
            y = int(self.H * 0.38)
        bw, bh = int(self.W * 0.8), int(self.H * (0.3 if big else 0.28))
        x = int((self.W - bw) / 2)
        sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y + int(self.H * 0.022), bw, bh)
        sh.fill.solid(); sh.fill.fore_color.rgb = _rgb(self.pal.button_shadow); sh.line.fill.background()
        b = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, bw, bh)
        b.fill.solid(); b.fill.fore_color.rgb = _rgb(self.pal.button); b.line.fill.background()
        tf = b.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        n = len(text)
        size = (180 if n < 25 else 150 if n < 40 else 125 if n < 60 else 100) * self.k
        self._runs(p, text, "FFFFFF", True, size)
        self.body = None

    # -- text
    def _runs(self, p, text, color=None, bold=False, size=None):
        for tok in _TOK.split(text or ""):
            if not tok:
                continue
            if tok.startswith("$") and tok.endswith("$") and len(tok) > 1:
                uni = unicode_formula(tok[1:-1])
                pieces = [(uni, 0)] if uni is not None else _math_runs(tok[1:-1])
                b = bold
            elif tok.startswith("**"):
                pieces, b = [(tok[2:-2], 0)], True
            else:
                pieces, b = [(tok, 0)], bold
            for t, base in pieces:
                r = p.add_run()
                r.text = t
                if b:
                    r.font.bold = True
                if size:
                    r.font.size = Pt(size)
                if color:
                    r.font.color.rgb = _rgb(color)
                if base:
                    r._r.get_or_add_rPr().set("baseline", str(base))

    def _cost(self, *texts):
        n = sum(len(re.sub(r"[$*{}_^]", "", t or "")) for t in texts)
        return max(1, -(-n // self.chars))

    def _para(self, cost):
        if self.body is None or self.units + cost > self.max_units:
            self._content_slide()
        self.units += cost
        if self._first:
            self._first = False
            p = self.body.paragraphs[0]
        else:
            p = self.body.add_paragraph()
        pPr = p._p.get_or_add_pPr()  # no bullets, like the team's slides
        pPr.set("marL", "0"); pPr.set("indent", "0")
        if pPr.find(qn("a:buNone")) is None:
            pPr.append(pPr.makeelement(qn("a:buNone"), {}))
        return p

    def heading(self, text, color=None):
        text = re.sub(r"^[#★\s]+", "", text or "").rstrip(" :-–")
        p = self._para(self._cost(text) + (1 if self.units else 0))
        self._runs(p, text + " :-", color or self.pal.heading, True)

    def line(self, text, label=None, label_color=None, color=None, marker=None, indent=0, bold=False):
        p = self._para(self._cost(text, label, marker))
        if indent:
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", str(int(1371600 * self.k * indent)))
        if marker:
            self._runs(p, marker + " ", color, True)
        if label:
            self._runs(p, label + " ", label_color or self.pal.label, True)
        self._runs(p, text, color, bold)

    def figure(self, b):
        s = self._content_slide()
        self.body = None
        p = s.shapes  # caption on top in green
        cap = (b.get("caption") or "").strip()
        tb = s.shapes.add_textbox(int(self.W * 0.034), int(self.H * 0.09), int(self.W * 0.93), int(self.H * 0.1))
        if cap:
            self._runs(tb.text_frame.paragraphs[0], cap, self.pal.heading, True, 80 * self.k)
        for ph in list(s.placeholders):
            ph._element.getparent().remove(ph._element)
        fw = max(20.0, float(b.get("width") or 250)); fh = max(20.0, float(b.get("height") or 150))
        top = int(self.H * 0.2); avail_w, avail_h = int(self.W * 0.9), int(self.H * 0.76)
        k = min(avail_w / fw, avail_h / fh)
        ox = (self.W - fw * k) / 2; oy = top + (avail_h - fh * k) / 2
        E = lambda v: Emu(int(v))
        fmap = self.pal.figure_map
        for sh in b.get("shapes") or []:
            if not isinstance(sh, dict):
                continue
            kind = sh.get("shape") or sh.get("type")
            col = colour(sh.get("color"), "000000")
            col = fmap.get(col, col)
            num = lambda key, d=0.0: float(sh.get(key, d) or d)
            try:
                if kind in ("rect", "text", "circle"):
                    if kind == "circle":
                        r = max(1.0, num("r", 5))
                        shp = p.add_shape(MSO_SHAPE.OVAL, E(ox + (num("cx") - r) * k), E(oy + (num("cy") - r) * k),
                                          E(2 * r * k), E(2 * r * k))
                    else:
                        shp = p.add_shape(MSO_SHAPE.RECTANGLE, E(ox + num("x") * k), E(oy + num("y") * k),
                                          E(num("w", 40) * k), E(num("h", 20) * k))
                    shp.fill.background()  # dark slide: no white fills
                    if kind == "text":
                        shp.line.fill.background()
                    else:
                        shp.line.color.rgb = _rgb(col); shp.line.width = Pt(3 * self.k * 2.5)
                    if sh.get("text"):
                        tf = shp.text_frame; tf.word_wrap = True
                        for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
                            setattr(tf, m, Emu(20000))
                        size = max(20 * self.k, num("size", 10) * k / 12700 * 0.9)
                        for n, ln in enumerate(str(sh["text"]).split("\n")):
                            para = tf.paragraphs[0] if n == 0 else tf.add_paragraph()
                            para.alignment = PP_ALIGN.CENTER
                            self._runs(para, ln, "FFFFFF", bool(sh.get("bold")), size)
                elif kind in ("line", "arrow"):
                    c = p.add_connector(MSO_CONNECTOR.STRAIGHT, E(ox + num("x1") * k), E(oy + num("y1") * k),
                                        E(ox + num("x2") * k), E(oy + num("y2") * k))
                    c.line.color.rgb = _rgb(col); c.line.width = Pt(3 * self.k * 2.5)
                    if kind == "arrow":
                        ln = c.line._get_or_add_ln()
                        ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle"}))
            except Exception:
                continue

    def save(self, path):
        self.prs.save(path)


def render_template_deck(parts, out_path, template, logo=None, title=None):
    d = TemplateDeck(template, logo)
    pal = d.pal
    current = None
    for part in parts:
        t = (title or part.get("chapter_title") or "").strip()
        if t and t != current:
            d.divider(t, big=True)
            current = t
        for b in part.get("blocks") or []:
            if not isinstance(b, dict):
                continue
            ty, text = b.get("type"), str(b.get("text") or "")
            try:
                if ty == "h1":
                    if text.lstrip().startswith(("#", "★")):  # teacher's main-topic marker -> divider slide
                        d.divider(re.sub(r"^[#★\s]+", "", text))
                    d.heading(text)
                elif ty == "h2":
                    d.heading(text, pal.accent)
                elif ty == "item":
                    d.line(text, marker=str(b.get("marker") or "•"), indent=int(b.get("level") or 0))
                elif ty == "example":
                    d.line(text, label="उदाहरण:")
                elif ty == "note":
                    lab = str(b.get("label") or "")
                    d.line(text, label=("ध्यान दें:" if lab.upper() in ("", "NOTE") else lab + ":"),
                           label_color=pal.alert)
                elif ty == "qa":
                    d.line(str(b.get("q") or ""), label="प्रश्न –", label_color=pal.heading)
                    d.line(str(b.get("a") or ""), label="उत्तर –")
                elif ty == "compare":
                    for k, pr in enumerate(b.get("pairs") or [], 1):
                        if isinstance(pr, (list, tuple)) and len(pr) >= 2:
                            d.line(str(pr[0]), marker=f"({_roman(k)})", label=f"{b.get('a_title', '')} :",
                                   label_color=pal.accent)
                            d.line(str(pr[1]), label=f"{b.get('b_title', '')} :", indent=1)
                elif ty == "figure":
                    d.figure(b)
                elif ty == "title":
                    pass
                elif text.strip():
                    d.line(text)
            except Exception:
                continue
    d.save(out_path)
    return out_path


def _roman(n):
    from .docxkit import _roman as r
    return r(n)


def render_mcq_deck(mcqs, out_path, template=None, logo=None, title=None):
    """Objective deck like the team's 'CHEMISTRY OBJECTIVE' PPT: one question per slide,
    question green + bold, options (A)-(D) white.  Without a template a plain deck is made."""
    if template:
        d = TemplateDeck(template, logo)
        if title:
            d.divider(title, big=True)
        d.divider("वस्तुनिष्ठ प्रश्न", big=True)
        for n, m in enumerate(mcqs, 1):
            d.body = None
            d.line(f"{n}. {m.get('q', '')}", color=d.pal.heading, bold=True)
            for letter, o in zip("ABCD", m.get("options") or []):
                d.line(f"({letter}) {o}")
        d.save(out_path)
        return out_path
    from .plainppt import render_plain_mcq_deck
    return render_plain_mcq_deck(mcqs, out_path, title)
