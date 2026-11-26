"""docxkit: small helpers on top of python-docx for typist-style notes.

Reusable pieces:
  * NoteDoc        - A4 document with Hindi-ready styles (complex-script fonts, bold, size).
  * inline markup  - "text **bold** $NaNO_3$" -> runs + native Word equations (OMML).
  * omml()         - tiny linear syntax -> OMML  (x_2, x^{2+}, {..} groups, \\frac{a}{b}).
  * Figure         - vector drawings as a native DrawingML group (rect/ellipse/line/arrow/text),
                     so figures stay editable shapes in Word instead of pictures.
"""
from __future__ import annotations

import re
from xml.sax.saxutils import escape

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.shared import Cm, Pt, RGBColor

HINDI_FONT = "Sanskrit Text"   # set from Style.font by NoteDoc
MATH_FONT = "Cambria Math"
MATH_SIZE = 13

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
}


def _ns(*prefixes):
    return " ".join(f'xmlns:{p}="{NS[p]}"' for p in prefixes)


def _rpr_xml(bold=False, italic=False, size=None, color=None, font=None):
    font = font or HINDI_FONT
    parts = [f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}" w:eastAsia="{font}"/>']
    if bold:
        parts.append("<w:b/><w:bCs/>")
    if italic:
        parts.append("<w:i/><w:iCs/>")
    if color:
        parts.append(f'<w:color w:val="{color}"/>')
    if size:
        parts.append(f'<w:sz w:val="{int(size*2)}"/><w:szCs w:val="{int(size*2)}"/>')
    parts.append('<w:lang w:val="en-IN" w:bidi="hi-IN"/>')
    return "<w:rPr>" + "".join(parts) + "</w:rPr>"


# ---------------------------------------------------------------- OMML ----
def _m_run(text, upright=True):
    sty = '<m:rPr><m:sty m:val="p"/></m:rPr>' if upright else ""
    return (f'<m:r>{sty}<w:rPr><w:rFonts w:ascii="{MATH_FONT}" w:hAnsi="{MATH_FONT}"/>'
            f'<w:sz w:val="{MATH_SIZE*2}"/><w:szCs w:val="{MATH_SIZE*2}"/></w:rPr>'
            f'<m:t xml:space="preserve">{escape(text)}</m:t></m:r>')


def _parse_group(s, i):
    """Return (content, next_index) for a {group} or a single char at s[i]."""
    if s[i] == "{":
        depth, j = 1, i + 1
        while depth:
            depth += {"{": 1, "}": -1}.get(s[j], 0)
            j += 1
        return s[i + 1:j - 1], j
    return s[i], i + 1


def omml_inner(s, upright=True):
    """Linear syntax -> OMML children.  A_2, A^{2+}, A_{2}^{-}, \\frac{a}{b}, {grouped base}."""
    out, buf, i = [], "", 0

    def flush(keep_last=False):
        nonlocal buf
        base = ""
        if keep_last and buf:
            buf, base = buf[:-1], buf[-1]
        if buf:
            out.append(_m_run(buf, upright))
        buf = ""
        return base

    while i < len(s):
        c = s[i]
        if s.startswith("\\frac", i):
            flush()
            num, i = _parse_group(s, i + 5)
            den, i = _parse_group(s, i)
            out.append(f"<m:f><m:num>{omml_inner(num, upright)}</m:num>"
                       f"<m:den>{omml_inner(den, upright)}</m:den></m:f>")
        elif c == "{" :
            flush()
            grp, i = _parse_group(s, i)
            base_xml = omml_inner(grp, upright)
            i = _script(s, i, base_xml, out, upright)
        elif c in "_^" :
            base = flush(keep_last=True)
            i = _script(s, i, _m_run(base, upright) if base else "", out, upright)
        else:
            buf += c
            i += 1
    flush()
    return "".join(out)


def _script(s, i, base_xml, out, upright):
    sub = sup = None
    while i < len(s) and s[i] in "_^":
        kind = s[i]
        val, i = _parse_group(s, i + 1)
        if kind == "_":
            sub = val
        else:
            sup = val
    if sub is None and sup is None:
        out.append(base_xml)
    elif sup is None:
        out.append(f"<m:sSub><m:e>{base_xml}</m:e><m:sub>{omml_inner(sub, upright)}</m:sub></m:sSub>")
    elif sub is None:
        out.append(f"<m:sSup><m:e>{base_xml}</m:e><m:sup>{omml_inner(sup, upright)}</m:sup></m:sSup>")
    else:
        out.append(f"<m:sSubSup><m:e>{base_xml}</m:e><m:sub>{omml_inner(sub, upright)}</m:sub>"
                   f"<m:sup>{omml_inner(sup, upright)}</m:sup></m:sSubSup>")
    return i


def omml(s, upright=True):
    return parse_xml(f"<m:oMath {_ns('m', 'w')}>{omml_inner(s, upright)}</m:oMath>")


# ------------------------------------------------------- inline markup ----
_TOKEN = re.compile(r"(\$[^$]+\$|\*\*[^*]+\*\*)")


def add_inline(p, text, bold=False, size=None, color=None):
    """Append marked-up text to paragraph p.  **bold**, $chem/math$ -> OMML."""
    for tok in _TOKEN.split(text):
        if not tok:
            continue
        if tok.startswith("$"):
            m = omml(tok[1:-1])
            if color:  # equation text follows the surrounding colour
                for rpr in m.iter(f"{{{NS['w']}}}rPr"):
                    rpr.append(parse_xml(f'<w:color {_ns("w")} w:val="{color}"/>'))
            p._p.append(m)
        else:
            b = bold
            if tok.startswith("**"):
                tok, b = tok[2:-2], True
            r = parse_xml(f'<w:r {_ns("w")}>{_rpr_xml(b, size=size, color=color)}'
                          f'<w:t xml:space="preserve">{escape(tok)}</w:t></w:r>')
            p._p.append(r)
    return p


# ------------------------------------------------------------ document ----
class Style:
    """House style for notes (user's standing requirement, 2026-11-26).

    Font Sanskrit Text 13 pt everywhere; 4 colours: red = topics/headings,
    blue = sub-topics & labels, black = body, green = examples/notes.
    A4 portrait, narrow margins (1.27 cm), 2 columns.  No tables: comparisons are
    written as coloured points.  Class notices / thumbnails / greetings are not typed.
    """
    font = "Sanskrit Text"
    size = 13
    red, blue, black, green = "C00000", "1F4E9F", "000000", "00703C"
    margin_cm = 1.27
    columns = 2
    column_gap_cm = 0.75


class NoteDoc:
    def __init__(self, style=Style):
        global HINDI_FONT
        self.st = style
        HINDI_FONT = style.font
        self.doc = d = Document()
        sec = d.sections[0]
        sec.page_width, sec.page_height = Cm(21), Cm(29.7)
        for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
            setattr(sec, side, Cm(style.margin_cm))
        sec.header_distance = sec.footer_distance = Cm(0.6)
        if style.columns > 1:
            sec._sectPr.append(parse_xml(
                f'<w:cols {_ns("w")} w:num="{style.columns}" w:space="{int(style.column_gap_cm*567)}"/>'))
        usable = 21 - 2 * style.margin_cm
        col_cm = (usable - (style.columns - 1) * style.column_gap_cm) / style.columns
        Figure.max_width = col_cm / 2.54 * 72 - 4  # points
        self._fig_id = 100
        self._style_fonts()

    def _style_fonts(self):
        st, S = self.doc.styles, self.st
        for name, color, bold in (("Normal", S.black, False), ("Title", S.red, True),
                                  ("Heading 1", S.red, True), ("Heading 2", S.blue, True),
                                  ("Caption", S.blue, False)):
            s = st[name]
            rpr = s.element.get_or_add_rPr()
            new = parse_xml(_rpr_xml(bold, size=S.size, color=color, font=S.font)
                            .replace("<w:rPr>", f"<w:rPr {_ns('w')}>"))
            for child in list(rpr):
                rpr.remove(child)
            for child in list(new):
                rpr.append(child)
            pf = s.paragraph_format
            pf.space_before = Pt(6 if name.startswith("Heading") else 0)
            pf.space_after = Pt(3)
            # drop theme borders/shading Word's built-in Title style carries
            ppr = s.element.get_or_add_pPr()
            for el in ppr.findall(f"{{{NS['w']}}}pBdr"):
                ppr.remove(el)
        st["Normal"].paragraph_format.line_spacing = 1.1

    # --- text blocks
    def title(self, text, sub=None):
        p = self.doc.add_paragraph(style="Title")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_inline(p, text)
        if sub:
            self.para(sub, align="center", color=self.st.blue, bold=True)

    def h1(self, text, **_):
        p = self.doc.add_paragraph(style="Heading 1")
        add_inline(p, text)
        return p

    def h2(self, text, **_):
        p = self.doc.add_paragraph(style="Heading 2")
        add_inline(p, text)
        return p

    def para(self, text, bold=False, indent=0.0, align=None, color=None, **_):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if align == "center" else WD_ALIGN_PARAGRAPH.JUSTIFY
        if indent:
            p.paragraph_format.left_indent = Cm(indent)
        add_inline(p, text, bold=bold, color=color)
        return p

    def example(self, text):
        """'उदाहरण – ...' line: label blue, examples green."""
        p = self.doc.add_paragraph()
        add_inline(p, "**उदाहरण –** ", color=self.st.blue)
        add_inline(p, text, color=self.st.green)
        return p

    def item(self, marker, text, level=0, color=None):
        """Numbered point with a hanging indent, marker typed as text e.g. '(i)'."""
        p = self.doc.add_paragraph()
        pf = p.paragraph_format
        pf.left_indent = Cm(1.1 + 0.6 * level)
        pf.first_line_indent = Cm(-1.1)
        pf.tab_stops.add_tab_stop(Cm(1.1 + 0.6 * level))
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        add_inline(p, marker, bold=True, color=self.st.blue)
        add_inline(p, "\t" + text, color=color)
        return p

    def compare(self, title_a, title_b, pairs):
        """Comparison as coloured points instead of a table."""
        for k, (a, b) in enumerate(pairs, 1):
            m = f"({_roman(k)})"
            p = self.item(m, f"**{title_a} :** {a}", color=self.st.blue)
            q = self.doc.add_paragraph()
            q.paragraph_format.left_indent = Cm(1.1)
            q.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            add_inline(q, f"**{title_b} :** {b}", color=self.st.green)

    def qa(self, q, a):
        self.para(f"**प्रश्न –** {q}", color=self.st.red)
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        add_inline(p, "**उत्तर –** ", color=self.st.blue)
        add_inline(p, a)

    def note(self, text, label="NOTE"):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        add_inline(p, f"**[{label}]** ", color=self.st.red)
        add_inline(p, text, color=self.st.green)
        return p

    def figure(self, fig: "Figure", caption=None):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.keep_with_next = True
        self._fig_id += 1
        p._p.append(fig.to_run(self._fig_id))
        if caption:
            c = self.doc.add_paragraph(style="Caption")
            c.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_inline(c, caption)

    def page_break(self):
        self.doc.add_page_break()

    def save(self, path):
        self.doc.save(path)


def _roman(n):
    vals = ((10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i"))
    out = ""
    for v, r in vals:
        while n >= v:
            out, n = out + r, n - v
    return out


# ------------------------------------------------------------ figures -----
EMU = 12700  # per point


class Figure:
    """Vector figure in points; emitted as one inline DrawingML group (editable in Word)."""

    max_width = None  # points; set by NoteDoc to the column width

    def __init__(self, w, h):
        self.s = min(1.0, self.max_width / w) if self.max_width else 1.0
        self.w, self.h, self.shapes = w * self.s, h * self.s, []

    def _add(self, xml):
        self.shapes.append(xml)

    @staticmethod
    def _ln(color="000000", width=1.0, dash=None, head=None, tail=None):
        d = f'<a:prstDash val="{dash}"/>' if dash else ""
        he = f'<a:headEnd type="{head}" w="med" len="med"/>' if head else ""
        te = f'<a:tailEnd type="{tail}" w="med" len="med"/>' if tail else ""
        return (f'<a:ln w="{int(width*EMU)}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
                f'{d}{he}{te}</a:ln>')

    @staticmethod
    def _fill(fill):
        return f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else "<a:noFill/>"

    def _shape(self, prst, x, y, w, h, line, fill=None, text=None, flip="", size=11,
               bold=False, color="000000", tx_inset=1.0):
        k = self.s
        x, y, w, h, size = x * k, y * k, w * k, h * k, max(7.5, size * k)
        xfrm = (f'<a:xfrm{flip}><a:off x="{int(x*EMU)}" y="{int(y*EMU)}"/>'
                f'<a:ext cx="{int(w*EMU)}" cy="{int(h*EMU)}"/></a:xfrm>')
        sp = (f"<wps:spPr>{xfrm}<a:prstGeom prst=\"{prst}\"><a:avLst/></a:prstGeom>"
              f"{self._fill(fill)}{line}</wps:spPr>")
        txb, nv = "", "<wps:cNvSpPr/>"
        ins = int(tx_inset * EMU)
        body = (f'<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="{ins}" tIns="{ins}" '
                f'rIns="{ins}" bIns="{ins}" anchor="ctr" anchorCtr="0"><a:noAutofit/></wps:bodyPr>')
        if text is not None:
            nv = '<wps:cNvSpPr txBox="1"/>' if prst == "rect" and "noFill" in line else nv
            paras = "".join(
                f'<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0" w:line="240" '
                f'w:lineRule="auto"/></w:pPr><w:r>{_rpr_xml(bold, size=size, color=color)}'
                f'<w:t xml:space="preserve">{escape(t)}</w:t></w:r></w:p>'
                for t in text.split("\n"))
            txb = f"<wps:txbx><w:txbxContent>{paras}</w:txbxContent></wps:txbx>"
        n = len(self.shapes) + 1
        self._add(f'<wps:wsp><wps:cNvPr id="{n}" name="s{n}"/>{nv}{sp}{txb}{body}</wps:wsp>')

    # public primitives (all in points)
    def rect(self, x, y, w, h, color="000000", width=1.0, fill=None, text=None, **kw):
        self._shape("rect", x, y, w, h, self._ln(color, width * self.s), fill, text, **kw)

    def text(self, x, y, w, h, text, **kw):
        self._shape("rect", x, y, w, h, "<a:ln><a:noFill/></a:ln>", None, text, **kw)

    def circle(self, cx, cy, r, color="000000", width=1.0, fill=None, text=None, **kw):
        self._shape("ellipse", cx - r, cy - r, 2 * r, 2 * r, self._ln(color, width * self.s), fill, text,
                    tx_inset=0, **kw)

    def line(self, x1, y1, x2, y2, color="000000", width=1.0, dash=None, arrow=False):
        flip = (' flipH="1"' if x2 < x1 else "") + (' flipV="1"' if y2 < y1 else "")
        self._shape("line", min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1),
                    self._ln(color, width * self.s, dash, tail="triangle" if arrow else None), flip=flip)

    def arrow(self, x1, y1, x2, y2, **kw):
        self.line(x1, y1, x2, y2, arrow=True, **kw)

    def to_run(self, fid):
        W, H = int(self.w * EMU), int(self.h * EMU)
        grp = (f'<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/>'
               f'<a:ext cx="{W}" cy="{H}"/><a:chOff x="0" y="0"/><a:chExt cx="{W}" cy="{H}"/>'
               f'</a:xfrm></wpg:grpSpPr>{"".join(self.shapes)}</wpg:wgp>')
        xml = (f'<w:r {_ns("w", "wp", "a", "wps", "wpg")}><w:drawing>'
               f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
               f'<wp:extent cx="{W}" cy="{H}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
               f'<wp:docPr id="{fid}" name="Figure {fid}"/><wp:cNvGraphicFramePr/>'
               f'<a:graphic><a:graphicData uri="{NS["wpg"]}">{grp}</a:graphicData></a:graphic>'
               f'</wp:inline></w:drawing></w:r>')
        return parse_xml(xml)
