"""Plain PPT: the team's default deck (first priority, before any design template).

House rules (user, 2026-11-27):
  * black slide background, font Utsaah, size 120, text colours yellow / green / white
  * slide size 40 in x 22.504 in (same as the team's class PPTs)
  * topic-wise: every topic starts a new slide; a topic that does not fit continues on the
    next slide (long paragraphs are split between words, never mid-word)
  * MCQ questions are green, options white
  * maths in Unicode (H₂O, x², SO₄²⁻); anything Unicode cannot show (fractions) becomes a
    native PowerPoint equation
  * content is never changed: only layout, colour and spacing are added

How much fits is measured with the real font when it is installed (Windows: utsaah.ttf),
otherwise with a wider stand-in font, so the estimate errs on the side of more slides.
"""
from __future__ import annotations

import copy
import os
import re
from functools import lru_cache
from pathlib import Path
from xml.sax.saxutils import escape

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from .blocks import colour
from .docxkit import Style, omml_inner

FONT = "Utsaah"            # Windows' Hindi font (the user calls it "Utsah")
SIZE = 120                 # pt, every text
YELLOW, GREEN, WHITE, BLACK = "FFFF00", "00FF00", "FFFFFF", "000000"
W, H = Emu(36576000), Emu(20577175)   # 40 in x 22.504 in, as in the team's decks
# margins as on the user's ruler screenshot (2026-11-27): text starts ~4.2 in from the left, ~2 in from the top
ML, MR, MT, MB = Inches(4.2), Inches(2.0), Inches(2.0), Inches(1.2)
MX, MY = ML, MT
BOX_W = Emu(W - ML - MR)
LINE_IN = SIZE * 1.2 / 72             # one line of 120 pt text, in inches (refined from the font below)
GAP_PT = 18                           # space after each paragraph
HANG = Inches(1.6)                    # hanging indent for "(i)", "•", "(A)" ...

NS_A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

SUB = str.maketrans("0123456789+-=()naeoxhklmnpst", "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₙₐₑₒₓₕₖₗₘₙₚₛₜ")
SUP = str.maketrans("0123456789+-=()ni", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱ")
SUB.update({ord("−"): "₋", ord("–"): "₋", ord("¯"): "₋"})
SUP.update({ord("−"): "⁻", ord("–"): "⁻", ord("¯"): "⁻", ord("∘"): "°", ord("°"): "°", ord("o"): "°"})
Q_COLOR = "00FF00"   # MCQ question colour (user rule: green)
_TOK = re.compile(r"(\$[^$]+\$|\*\*[^*]+\*\*)")
_OPTION = re.compile(r"^\(?([A-Da-d]|[क-घ]|i{1,3}|iv)[.)]$")
_QNUM = re.compile(r"^(?:Q\.?\s*)?\(?\d{1,3}[.)]$", re.I)


# ------------------------------------------------------------- measuring --
def _font_files():
    win = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"
    cands = [(win / "utsaah.ttf", win / "utsaahb.ttf", 1.0),
             (Path("/usr/share/fonts/truetype/freefont/FreeSerif.ttf"),
              Path("/usr/share/fonts/truetype/freefont/FreeSerifBold.ttf"), 1.0)]
    for reg, bold, k in cands:
        if reg.exists():
            return reg, (bold if bold.exists() else reg), k
    return None, None, 1.0


@lru_cache(maxsize=2)
def _pil_font(bold):
    try:
        from PIL import ImageFont
        reg, b, _ = _font_files()
        path = b if bold else reg
        return ImageFont.truetype(str(path), SIZE) if path else None
    except Exception:
        return None


def _line_in():
    f = _pil_font(False)
    try:
        a, d = f.getmetrics()
        return max(1.1 * SIZE, a + d) / 72
    except Exception:
        return SIZE * 1.2 / 72


def text_width_in(text, bold=False):
    """Width of text at 120 pt, in inches (1 px at size 120 = 1 pt)."""
    f = _pil_font(bool(bold))
    if f is None:
        return len(text) * SIZE * 0.5 / 72 * (1.06 if bold else 1.0)
    try:
        return f.getlength(text) / 72 * 1.04   # small safety margin for kerning / shaping
    except Exception:
        return len(text) * SIZE * 0.5 / 72


# ---------------------------------------------------------------- maths ---
_TEX = {"\\times": "×", "\\cdot": "·", "\\rightarrow": "→", "\\to": "→", "\\leftarrow": "←",
        "\\rightleftharpoons": "⇌", "\\Delta": "Δ", "\\alpha": "α", "\\beta": "β", "\\gamma": "γ",
        "\\lambda": "λ", "\\mu": "μ", "\\pi": "π", "\\theta": "θ", "\\approx": "≈", "\\neq": "≠",
        "\\leq": "≤", "\\geq": "≥", "\\pm": "±", "\\degree": "°", "\\circ": "°", "\\infty": "∞"}


def tex_symbols(s):
    for k, v in sorted(_TEX.items(), key=lambda kv: -len(kv[0])):
        s = re.sub(re.escape(k) + r"(?![A-Za-z])", v, s)
    return s


def _script_runs(s):
    """Linear '$..$' markup -> [(text, 'sub'|'sup'|'')], or None for fractions."""
    s = tex_symbols(s)
    if "\\frac" in s or "\\sqrt" in s:
        return None
    out, i = [], 0

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
            out.append((g.replace("{", "").replace("}", ""), "sub" if c == "_" else "sup"))
        elif c in "{}":
            i += 1
        else:
            out.append((c, ""))
            i += 1
    return out


def unicode_math(s):
    """'SO_4^{2-}' -> 'SO₄²⁻'.  None when it needs a real equation (fraction, or a script
    character without a Unicode form)."""
    runs = _script_runs(s)
    if runs is None:
        return None
    out = []
    for t, kind in runs:
        if kind:
            conv = t.translate(SUB if kind == "sub" else SUP)
            if any(a == b and not a.isspace() and a != "°" for a, b in zip(conv, t)):
                return None
            out.append(conv)
        else:
            out.append(t)
    return "".join(out)


def linear_math(s):
    """Readable plain form for measuring and for the fallback copy: \\frac{a}{b} -> a/b."""
    s = tex_symbols(s)
    s = re.sub(r"\\frac\{([^{}]*)\}\{([^{}]*)\}", r"(\1)/(\2)", s)
    s = re.sub(r"\((\w+)\)/\((\w+)\)", r"\1/\2", s)
    uni = unicode_math(s)
    if uni is not None:
        return uni
    out = []
    for t, kind in _script_runs(s) or [(s, "")]:
        if kind:
            conv = t.translate(SUB if kind == "sub" else SUP)
            ok = all(a != b or a.isspace() for a, b in zip(conv, t))
            out.append(conv if ok else ("_" if kind == "sub" else "^") + t)
        else:
            out.append(t)
    return "".join(out).replace("\\", "")


def _omml_xml(s, color, bold):
    """PowerPoint equation (a14:m) for linear markup; runs use Cambria Math in our colour."""
    rpr = (f'<a:rPr lang="en-US" sz="{SIZE * 100}" b="{1 if bold else 0}" i="0">'
           f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
           f'<a:latin typeface="Cambria Math"/><a:cs typeface="Cambria Math"/></a:rPr>')
    inner = re.sub(r"<w:rPr>.*?</w:rPr>", rpr, omml_inner(tex_symbols(s), upright=True))
    return (f'<a14:m xmlns:a14="{NS_A14}" xmlns:a="{NS_A}">'
            f'<m:oMath xmlns:m="{NS_M}">{inner}</m:oMath></a14:m>')


# --------------------------------------------------------------- atoms ----
# A paragraph is a list of atoms: (text, color, bold, kind) with kind "" (text) or "math".
def atoms(markup, color=WHITE, bold=False):
    out = []
    for tok in _TOK.split(markup or ""):
        if not tok:
            continue
        if tok.startswith("$") and tok.endswith("$") and len(tok) > 1:
            out.append((tok[1:-1], color, bold, "math"))
        elif tok.startswith("**") and tok.endswith("**"):
            out.append((tok[2:-2], color, True, ""))
        else:
            out.append((tok, color, bold, ""))
    return out


def _words(ats):
    """Split text atoms into word pieces (each keeps its trailing space)."""
    out = []
    for t, c, b, k in ats:
        if k == "math":
            out.append((t, c, b, k))
            continue
        for w in re.findall(r"\S+\s*|\s+", t):
            out.append((w, c, b, ""))
    return out


def _wlen(w):
    t, _, b, k = w
    return text_width_in(linear_math(t) if k == "math" else t, b)


def wrap_lines(ats, width_in):
    """Greedy word wrap -> list of lines (each a list of word atoms)."""
    lines, cur, x = [], [], 0.0
    for w in _words(ats):
        wl = _wlen(w)
        bare = text_width_in(w[0].rstrip(), w[2]) if w[3] == "" else wl
        if cur and x + bare > width_in:
            lines.append(cur)
            cur, x = [], 0.0
            if w[3] == "" and not w[0].strip():
                continue
        cur.append(w)
        x += wl
    if cur:
        lines.append(cur)
    return lines or [[]]


# ---------------------------------------------------------------- deck ----
class PlainDeck:
    def __init__(self):
        global LINE_IN
        LINE_IN = _line_in()
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = W, H
        self.layout = self.prs.slide_layouts[6]   # blank
        self.slide = None
        self.body = None
        self.used = 0.0        # inches used in the body box
        self.avail = 0.0
        self.math_shapes = []
        self.math_text = {}   # a14:m element -> plain text for the fallback copy

    # -- slides
    def _new(self):
        s = self.prs.slides.add_slide(self.layout)
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = RGBColor.from_string(BLACK)
        self.slide = s
        self.body = None
        return s

    def _box(self, top, height):
        tb = self.slide.shapes.add_textbox(MX, top, BOX_W, height)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.auto_size = None
        for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
            setattr(tf, m, Emu(0))
        return tb

    def topic_slide(self, heading_markup):
        """New slide with the topic heading (yellow, bold) on top and a body box below."""
        self._new()
        ats = atoms(heading_markup, YELLOW, True)
        n = len(wrap_lines(ats, BOX_W.inches))
        h = Inches(n * LINE_IN + 0.15)
        tb = self._box(MY, h)
        p = tb.text_frame.paragraphs[0]
        self._fill(p, ats, tb)
        self._open_body(MY + h + Inches(0.35))

    def _open_body(self, top=None):
        if self.slide is None or top is None:
            self._new()
            top = MY
        height = H - top - MB
        self.body = self._box(top, height)
        self.body.text_frame.vertical_anchor = MSO_ANCHOR.TOP
        self.used, self.avail, self._first = 0.0, Emu(height).inches, True

    def _cont(self):
        self._open_body()   # always a fresh slide

    # -- paragraphs
    def _para_height(self, n_lines):
        return n_lines * LINE_IN + GAP_PT / 72

    def add(self, ats, indent=0, hang=False, keep_with_next=0):
        """Add one paragraph; splits it between slides when it does not fit."""
        if self.body is None:
            self._open_body()
        marl = HANG * (indent + (1 if hang else 0))
        width = BOX_W.inches - Emu(marl).inches
        lines = wrap_lines(ats, width)
        while lines:
            room = int((self.avail - self.used - GAP_PT / 72) // LINE_IN)
            need = len(lines) + keep_with_next
            if need <= room:
                take = len(lines)
            elif room >= 2 and len(lines) > room and not keep_with_next:
                take = room          # split a long paragraph across slides
            elif self.used == 0:     # empty slide: put as much as fits
                take = max(1, min(len(lines), room))
            else:
                self._cont()
                continue
            chunk = [w for ln in lines[:take] for w in ln]
            self._emit(chunk, marl, hang)
            self.used += self._para_height(take)
            lines = lines[take:]
            if lines:
                self._cont()
                hang = False         # the continued part has no marker

    def _emit(self, chunk, marl, hang):
        tf = self.body.text_frame
        p = tf.paragraphs[0] if self._first else tf.add_paragraph()
        self._first = False
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(int(marl)))
        pPr.set("indent", str(-int(HANG) if hang else 0))
        if pPr.find(qn("a:buNone")) is None:
            pPr.append(pPr.makeelement(qn("a:buNone"), {}))
        p.space_after = Pt(GAP_PT)
        p.line_spacing = 1.0
        self._fill(p, chunk, self.body)

    def _fill(self, p, ats, shape):
        merged = []
        for a in ats:  # join neighbouring text pieces with the same look
            if merged and a[3] == "" and merged[-1][3] == "" and merged[-1][1:3] == a[1:3]:
                merged[-1] = (merged[-1][0] + a[0],) + merged[-1][1:]
            else:
                merged.append(a)
        if merged and merged[-1][3] == "":
            merged[-1] = (merged[-1][0].rstrip(),) + merged[-1][1:]
        for t, c, b, k in merged:
            if k == "math":
                uni = unicode_math(t)
                if uni is not None:
                    self._run(p, uni, c, b)
                else:
                    el = etree.fromstring(_omml_xml(t, c, b))
                    self.math_text[el] = linear_math(t)
                    end = p._p.find(qn("a:endParaRPr"))
                    (end.addprevious(el) if end is not None else p._p.append(el))
                    if shape not in self.math_shapes:
                        self.math_shapes.append(shape)
            elif t:
                self._run(p, t, c, b)

    @staticmethod
    def _run(p, text, color, bold):
        r = p.add_run()
        r.text = text
        f = r.font
        f.size, f.bold, f.name = Pt(SIZE), bool(bold), FONT
        f.color.rgb = RGBColor.from_string(color)
        rpr = r._r.get_or_add_rPr()
        for tag in ("a:ea", "a:cs"):
            el = rpr.find(qn(tag))
            if el is None:
                el = rpr.makeelement(qn(tag), {})
                rpr.append(el)
            el.set("typeface", FONT)

    # -- special slides
    def title_slide(self, text):
        self._new()
        tb = self.slide.shapes.add_textbox(MX, Inches(2), BOX_W, H - Inches(4))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        self._fill(p, atoms(text, YELLOW, True), tb)
        self.body = None

    FIG_MAP = {"000000": WHITE, Style.black: WHITE, Style.blue: GREEN, Style.red: YELLOW,
               Style.green: GREEN, "404040": WHITE, "1F4E9F": GREEN, "FFFFFF": WHITE}

    def figure(self, b):
        """Vector figure on its own slide: caption on top (yellow), shapes redrawn white/green/yellow."""
        self._new()
        top = MY
        cap = (b.get("caption") or "").strip()
        if cap:
            ats = atoms(cap, YELLOW, True)
            n = len(wrap_lines(ats, BOX_W.inches))
            tb = self._box(MY, Inches(n * LINE_IN))
            self._fill(tb.text_frame.paragraphs[0], ats, tb)
            top = MY + Inches(n * LINE_IN + 0.4)
        fw = max(20.0, float(b.get("width") or 250))
        fh = max(20.0, float(b.get("height") or 150))
        avail_w, avail_h = BOX_W, H - top - MB
        k = min(avail_w / fw, avail_h / fh)
        ox, oy = MX + (avail_w - fw * k) / 2, top + (avail_h - fh * k) / 2
        E = lambda v: Emu(int(v))
        shapes = self.slide.shapes
        lw = Pt(6)
        for sh in b.get("shapes") or []:
            if not isinstance(sh, dict):
                continue
            kind = sh.get("shape") or sh.get("type")
            col = colour(sh.get("color"), "000000")
            col = self.FIG_MAP.get(col.upper(), col)
            num = lambda key, d=0.0: float(sh.get(key, d) or d)
            try:
                if kind in ("rect", "text", "circle"):
                    if kind == "circle":
                        r = max(1.0, num("r", 5))
                        shp = shapes.add_shape(MSO_SHAPE.OVAL, E(ox + (num("cx") - r) * k),
                                               E(oy + (num("cy") - r) * k), E(2 * r * k), E(2 * r * k))
                    else:
                        shp = shapes.add_shape(MSO_SHAPE.RECTANGLE, E(ox + num("x") * k), E(oy + num("y") * k),
                                               E(num("w", 40) * k), E(num("h", 20) * k))
                    shp.fill.background()
                    if kind == "text":
                        shp.line.fill.background()
                    else:
                        shp.line.color.rgb = RGBColor.from_string(col)
                        shp.line.width = lw
                    if sh.get("text"):
                        tf = shp.text_frame
                        tf.word_wrap = True
                        for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
                            setattr(tf, m, Emu(20000))
                        size = max(40.0, min(SIZE, num("size", 10) * k / 12700 * 0.9))
                        tcol = col if kind == "text" else WHITE
                        for n, ln in enumerate(str(sh["text"]).split("\n")):
                            para = tf.paragraphs[0] if n == 0 else tf.add_paragraph()
                            para.alignment = PP_ALIGN.CENTER
                            for t, c, bb, kk in atoms(ln, tcol, bool(sh.get("bold"))):
                                t = (unicode_math(t) or linear_math(t)) if kk == "math" else t
                                r = para.add_run()
                                r.text = t
                                r.font.size, r.font.bold, r.font.name = Pt(size), bb, FONT
                                r.font.color.rgb = RGBColor.from_string(c)
                elif kind in ("line", "arrow"):
                    c = shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(ox + num("x1") * k), E(oy + num("y1") * k),
                                             E(ox + num("x2") * k), E(oy + num("y2") * k))
                    c.line.color.rgb = RGBColor.from_string(col)
                    c.line.width = lw
                    if kind == "arrow":
                        ln = c.line._get_or_add_ln()
                        ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle"}))
            except Exception:
                continue
        self.body = None

    # -- save
    def _wrap_math(self):
        """PowerPoint keeps equation shapes inside mc:AlternateContent; the fallback copy
        (for old viewers) shows the same maths as plain text."""
        for shape in self.math_shapes:
            sp = shape._element
            parent = sp.getparent()
            if parent is None:
                continue
            fb = copy.deepcopy(sp)
            originals = list(sp.iter(f"{{{NS_A14}}}m"))
            for orig, m in zip(originals, list(fb.iter(f"{{{NS_A14}}}m"))):
                txt = self.math_text.get(orig) or "".join(m.itertext())
                rpr = m.find(f".//{{{NS_A}}}rPr")
                r = etree.SubElement(m.getparent(), qn("a:r"))
                if rpr is not None:
                    r.append(copy.deepcopy(rpr))
                etree.SubElement(r, qn("a:t")).text = txt
                m.addprevious(r)
                m.getparent().remove(m)
            ac = etree.Element(f"{{{NS_MC}}}AlternateContent", nsmap={"mc": NS_MC})
            choice = etree.SubElement(ac, f"{{{NS_MC}}}Choice", nsmap={"a14": NS_A14})
            choice.set("Requires", "a14")
            fallback = etree.SubElement(ac, f"{{{NS_MC}}}Fallback")
            parent.replace(sp, ac)
            choice.append(sp)
            fallback.append(fb)
        self.math_shapes = []

    def save(self, path):
        self._wrap_math()
        self.prs.save(path)


# ----------------------------------------------------------- rendering ----
def _is_option(b):
    return b.get("type") == "item" and bool(_OPTION.match(str(b.get("marker") or "").strip()))


def _is_question(blocks, i):
    """A numbered item / paragraph followed by at least two options (A)-(D) is an MCQ."""
    b = blocks[i]
    if b.get("type") not in ("item", "para"):
        return False
    mk = str(b.get("marker") or "").strip()
    starts_num = bool(_QNUM.match(mk)) if b.get("type") == "item" else bool(
        re.match(r"^\s*(?:Q\.?\s*)?\d{1,3}[.)]\s", str(b.get("text") or "")))
    if not starts_num:
        return False
    nxt = blocks[i + 1:i + 3]
    return len(nxt) == 2 and all(isinstance(x, dict) and _is_option(x) for x in nxt)


def _mcq(d, q, options, answer="", solution="", number=None):
    """One question per slide: question green, options white."""
    d._cont() if d.slide is not None else d._open_body()
    qa = atoms(q.replace("**", ""), Q_COLOR, True)
    if number is not None:
        qa = [(f"{number}. ", Q_COLOR, True, "")] + qa
    d.add(qa)
    for letter, o in options:
        d.add([(f"({letter})", WHITE, False, ""), (" ", WHITE, False, "")] + atoms(o, WHITE), hang=True)
    if str(answer).strip():
        d.add([("उत्तर – ", YELLOW, True, "")] + atoms(str(answer), WHITE))
    if str(solution).strip():
        d.add([("हल – ", YELLOW, True, "")] + atoms(str(solution), WHITE))
    d.body = None


def render_plain_deck(parts, out_path, title=None):
    d = PlainDeck()
    current = None
    for part in parts:
        t = (title or part.get("chapter_title") or "").strip()
        if t and t != current:
            d.title_slide(t)
            current = t
        blocks = [b for b in (part.get("blocks") or []) if isinstance(b, dict)]
        i = 0
        while i < len(blocks):
            b = blocks[i]
            ty, text = b.get("type"), str(b.get("text") or "")
            try:
                if ty == "title":
                    pass
                elif ty == "h1":
                    d.topic_slide(text)
                elif ty == "h2":
                    d.add(atoms(text, GREEN, True), keep_with_next=1)
                elif ty == "mcq":
                    letters = b.get("letters") or list("ABCD")
                    opts = list(zip(letters, [str(o).replace("**", "") for o in b.get("options") or []]))
                    _mcq(d, str(b.get("q") or ""), opts, b.get("answer", ""), b.get("solution", ""))
                elif _is_question(blocks, i):
                    mk = str(b.get("marker") or "").strip()
                    q = (mk + " " if ty == "item" else "") + text
                    j, opts = i + 1, []
                    while j < len(blocks) and _is_option(blocks[j]):
                        o = blocks[j]
                        opts.append((str(o.get("marker")).strip("().").strip(), str(o.get("text") or "")))
                        j += 1
                    d.body = None
                    d._cont() if d.slide is not None else d._open_body()
                    d.add(atoms(q.replace("**", ""), Q_COLOR, True))
                    for letter, o in opts:
                        d.add([(f"({letter})", WHITE, False, ""), (" ", WHITE, False, "")] + atoms(o, WHITE),
                              hang=True)
                    d.body = None
                    i = j
                    continue
                elif ty == "item":
                    mk = str(b.get("marker") or "•")
                    col = GREEN if colour(b.get("color")) == Style.green else WHITE
                    d.add([(mk, YELLOW, True, ""), (" ", WHITE, False, "")] + atoms(text, col),
                          indent=int(b.get("level") or 0), hang=True)
                elif ty == "example":
                    d.add([("उदाहरण – ", GREEN, True, "")] + atoms(text, WHITE))
                elif ty == "note":
                    lab = str(b.get("label") or "NOTE")
                    d.add([(f"[{lab}] ", YELLOW, True, "")] + atoms(text, WHITE))
                elif ty == "qa":
                    d.add([("प्रश्न – ", GREEN, True, "")] + atoms(str(b.get("q") or ""), GREEN, True),
                          keep_with_next=1 if str(b.get("a") or "").strip() else 0)
                    if str(b.get("a") or "").strip():
                        d.add([("उत्तर – ", YELLOW, True, "")] + atoms(str(b.get("a") or ""), WHITE))
                elif ty == "compare":
                    for k, pr in enumerate(b.get("pairs") or [], 1):
                        if isinstance(pr, (list, tuple)) and len(pr) >= 2:
                            d.add([(f"({k})", YELLOW, True, ""), (" ", WHITE, False, ""),
                                   (f"{b.get('a_title', '')} : ", YELLOW, True, "")] + atoms(str(pr[0]), WHITE),
                                  hang=True)
                            d.add([(f"{b.get('b_title', '')} : ", GREEN, True, "")] + atoms(str(pr[1]), WHITE),
                                  indent=1)
                elif ty == "figure":
                    d.figure(b)
                elif text.strip():
                    col = GREEN if colour(b.get("color")) == Style.green else WHITE
                    lead_in = bool(re.search(r"[–:-]\**\s*$", text)) and i + 1 < len(blocks)
                    d.add(atoms(text, col, bool(b.get("bold"))), keep_with_next=1 if lead_in else 0)
            except Exception:
                import traceback
                traceback.print_exc()
            i += 1
    if not len(d.prs.slides):
        d.title_slide(title or " ")
    d.save(out_path)
    return out_path


def render_plain_mcq_deck(mcqs, out_path, title=None):
    d = PlainDeck()
    if title:
        d.title_slide(title)
    for n, m in enumerate(mcqs, 1):
        _mcq(d, str(m.get("q", "")), list(zip("ABCD", m.get("options") or [])), number=n)
    if not len(d.prs.slides):
        d.title_slide(title or " ")
    d.save(out_path)
    return out_path
