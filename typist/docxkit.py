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
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.shared import Cm, Pt, RGBColor

HINDI_FONT = "Nirmala UI"      # ships with Windows; LibreOffice substitutes a Devanagari font
MATH_FONT = "Cambria Math"

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


def _rpr_xml(bold=False, italic=False, size=None, color=None, font=HINDI_FONT):
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
    return (f'<m:r>{sty}<w:rPr><w:rFonts w:ascii="{MATH_FONT}" w:hAnsi="{MATH_FONT}"/></w:rPr>'
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
            p._p.append(omml(tok[1:-1]))
        else:
            b = bold
            if tok.startswith("**"):
                tok, b = tok[2:-2], True
            r = parse_xml(f'<w:r {_ns("w")}>{_rpr_xml(b, size=size, color=color)}'
                          f'<w:t xml:space="preserve">{escape(tok)}</w:t></w:r>')
            p._p.append(r)
    return p


# ------------------------------------------------------------ document ----
class NoteDoc:
    def __init__(self):
        self.doc = d = Document()
        sec = d.sections[0]
        sec.page_width, sec.page_height = Cm(21), Cm(29.7)
        for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
            setattr(sec, side, Cm(2))
        self._fig_id = 100
        self._style_fonts()

    def _style_fonts(self):
        st = self.doc.styles
        for name, size, color in (("Normal", 12, None), ("Title", 22, "1F3864"),
                                  ("Subtitle", 14, "404040"), ("Heading 1", 16, "1F3864"),
                                  ("Heading 2", 13.5, "8B1A1A"), ("Caption", 10.5, "404040")):
            s = st[name]
            rpr = s.element.get_or_add_rPr()
            new = parse_xml(_rpr_xml(size=size, color=color).replace("<w:rPr>", f"<w:rPr {_ns('w')}>"))
            for child in list(rpr):
                rpr.remove(child)
            for child in list(new):
                rpr.append(child)
            if name.startswith("Heading") or name == "Title":
                rpr.append(parse_xml(f'<w:b {_ns("w")}/>'))
                rpr.append(parse_xml(f'<w:bCs {_ns("w")}/>'))
            s.paragraph_format.space_after = Pt(4)
        st["Normal"].paragraph_format.line_spacing = 1.2

    # --- text blocks
    def title(self, text, sub=None):
        p = self.doc.add_paragraph(style="Title")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_inline(p, text)
        if sub:
            q = self.doc.add_paragraph(style="Subtitle")
            q.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_inline(q, sub)

    def h1(self, text, boxed=True):
        p = self.doc.add_paragraph(style="Heading 1")
        add_inline(p, text)
        if boxed:
            self._border(p, "1F3864")
        return p

    def h2(self, text, boxed=False):
        p = self.doc.add_paragraph(style="Heading 2")
        add_inline(p, text)
        if boxed:
            self._border(p, "8B1A1A")
        return p

    def para(self, text, bold=False, indent=0.0, align=None, size=None, color=None):
        p = self.doc.add_paragraph()
        if indent:
            p.paragraph_format.left_indent = Cm(indent)
        if align == "center":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_inline(p, text, bold=bold, size=size, color=color)
        return p

    def item(self, marker, text, level=0):
        """Numbered point with a hanging indent, marker typed as text e.g. '(i)'."""
        p = self.doc.add_paragraph()
        pf = p.paragraph_format
        pf.left_indent = Cm(1.4 + 0.8 * level)
        pf.first_line_indent = Cm(-1.2)
        pf.tab_stops.add_tab_stop(Cm(1.4 + 0.8 * level))
        add_inline(p, f"{marker}\t{text}")
        return p

    def qa(self, q, a):
        self.para(f"**प्रश्न –** {q}", color="8B1A1A")
        self.para(f"**उत्तर –** {a}", indent=0.0)

    def note(self, text, label="NOTE"):
        p = self.para(f"**[{label}]** {text}")
        self._border(p, "C55A11", shade="FBE4D5")
        return p

    def notice(self, lines, label="कक्षा सूचना / विज्ञापन (Class notice)"):
        p = self.para(f"**{label}**", size=10.5, color="595959")
        self._border(p, "A6A6A6", shade="F2F2F2")
        for ln in lines:
            q = self.para(ln, size=10.5, color="595959")
            self._border(q, "A6A6A6", shade="F2F2F2")

    def table(self, header, rows, widths_cm=None):
        t = self.doc.add_table(rows=1, cols=len(header))
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for c, h in zip(t.rows[0].cells, header):
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_inline(c.paragraphs[0], h, bold=True)
            c._tc.get_or_add_tcPr().append(parse_xml(
                f'<w:shd {_ns("w")} w:val="clear" w:color="auto" w:fill="D9E2F3"/>'))
        for row in rows:
            cells = t.add_row().cells
            for c, v in zip(cells, row):
                add_inline(c.paragraphs[0], v)
        if widths_cm:
            for row in t.rows:
                for c, w in zip(row.cells, widths_cm):
                    c.width = Cm(w)
        # repeat header row on each page
        t.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {_ns("w")}/>'))
        self.doc.add_paragraph()
        return t

    def figure(self, fig: "Figure", caption=None):
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._fig_id += 1
        p._p.append(fig.to_run(self._fig_id))
        if caption:
            c = self.doc.add_paragraph(style="Caption")
            c.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_inline(c, caption)

    def page_break(self):
        self.doc.add_page_break()

    @staticmethod
    def _border(p, color, shade=None):
        ppr = p._p.get_or_add_pPr()
        edges = "".join(f'<w:{e} w:val="single" w:sz="8" w:space="4" w:color="{color}"/>'
                        for e in ("top", "left", "bottom", "right"))
        ppr.append(parse_xml(f'<w:pBdr {_ns("w")}>{edges}</w:pBdr>'))
        if shade:
            ppr.append(parse_xml(f'<w:shd {_ns("w")} w:val="clear" w:color="auto" w:fill="{shade}"/>'))

    def save(self, path):
        self.doc.save(path)


# ------------------------------------------------------------ figures -----
EMU = 12700  # per point


class Figure:
    """Vector figure in points; emitted as one inline DrawingML group (editable in Word)."""

    def __init__(self, w, h):
        self.w, self.h, self.shapes = w, h, []

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
        self._shape("rect", x, y, w, h, self._ln(color, width), fill, text, **kw)

    def text(self, x, y, w, h, text, **kw):
        self._shape("rect", x, y, w, h, "<a:ln><a:noFill/></a:ln>", None, text, **kw)

    def circle(self, cx, cy, r, color="000000", width=1.0, fill=None, text=None, **kw):
        self._shape("ellipse", cx - r, cy - r, 2 * r, 2 * r, self._ln(color, width), fill, text,
                    tx_inset=0, **kw)

    def line(self, x1, y1, x2, y2, color="000000", width=1.0, dash=None, arrow=False):
        flip = (' flipH="1"' if x2 < x1 else "") + (' flipV="1"' if y2 < y1 else "")
        self._shape("line", min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1),
                    self._ln(color, width, dash, tail="triangle" if arrow else None), flip=flip)

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
