"""Render a list of JSON "blocks" (what the vision model returns) into a NoteDoc.

Block types (all text fields accept docxkit markup: **bold**, $NaNO_3$ -> Word equation):

  {"type": "title",   "text": "..."}                      chapter title (red, centred)
  {"type": "h1",      "text": "..."}                      topic (red)
  {"type": "h2",      "text": "..."}                      sub-topic (blue)
  {"type": "para",    "text": "...", "color": "black|red|blue|green", "bold": false}
  {"type": "item",    "marker": "(i)", "text": "...", "level": 0, "color": "..."}
  {"type": "example", "text": "..."}                      "उदाहरण –" + green text
  {"type": "note",    "text": "...", "label": "NOTE"}
  {"type": "qa",      "q": "...", "a": "..."}
  {"type": "compare", "a_title": "...", "b_title": "...", "pairs": [["a", "b"], ...]}
  {"type": "figure",  "width": 250, "height": 120, "caption": "...",
                      "shapes": [{"shape": "rect|text|circle|line|arrow", ...}]}

Shapes use points inside the figure box (origin top-left):
  rect   x y w h [color line_width fill text size bold]
  text   x y w h text [size bold color]
  circle cx cy r [color width fill text size bold]
  line / arrow  x1 y1 x2 y2 [color width dash]
Colours are names (red/blue/green/black/grey/white) or 6-digit hex.

Unknown or malformed blocks degrade to a plain paragraph instead of failing the job.
"""
from __future__ import annotations

import json
import re

from .docxkit import Figure, NoteDoc, Style

_NAMED = {"red": Style.red, "blue": Style.blue, "green": Style.green, "black": Style.black,
          "grey": "404040", "gray": "404040", "white": "FFFFFF"}
_HEX = re.compile(r"^#?[0-9A-Fa-f]{6}$")


def colour(v, default=None):
    if not v:
        return default
    v = str(v).strip()
    if v.lower() in _NAMED:
        return _NAMED[v.lower()]
    if _HEX.match(v):
        return v.lstrip("#").upper()
    return default


def _num(v, default=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _s(v):
    return "" if v is None else str(v)


def build_figure(b) -> Figure:
    w = max(20.0, _num(b.get("width"), 250))
    h = max(20.0, _num(b.get("height"), 150))
    f = Figure(w, h)
    f.source = json.dumps(b, ensure_ascii=False, separators=(",", ":"))
    for sh in b.get("shapes") or []:
        if not isinstance(sh, dict):
            continue
        kind = sh.get("shape") or sh.get("type")
        col = colour(sh.get("color"), "000000")
        lw = _num(sh.get("width"), 1.0) if kind in ("line", "arrow", "circle") else _num(sh.get("line_width"), 1.0)
        size = _num(sh.get("size"), 10)
        bold = bool(sh.get("bold"))
        txt = sh.get("text")
        txt = None if txt in (None, "") else str(txt)
        try:
            if kind == "rect":
                f.rect(_num(sh.get("x")), _num(sh.get("y")), _num(sh.get("w"), 40), _num(sh.get("h"), 20),
                       color=col, width=lw, fill=colour(sh.get("fill")), text=txt, size=size, bold=bold)
            elif kind == "text":
                f.text(_num(sh.get("x")), _num(sh.get("y")), _num(sh.get("w"), 60), _num(sh.get("h"), 16),
                       txt or "", size=size, bold=bold, color=col)
            elif kind == "circle":
                f.circle(_num(sh.get("cx")), _num(sh.get("cy")), max(1.0, _num(sh.get("r"), 5)), color=col,
                         width=lw, fill=colour(sh.get("fill")), text=txt, size=size, bold=bold)
            elif kind in ("line", "arrow"):
                dash = "dash" if sh.get("dash") else None
                fn = f.arrow if kind == "arrow" else f.line
                fn(_num(sh.get("x1")), _num(sh.get("y1")), _num(sh.get("x2")), _num(sh.get("y2")),
                   color=col, width=lw, dash=dash)
        except Exception:  # one bad shape must not kill the figure
            continue
    return f


def render(blocks, d: NoteDoc):
    st = d.st
    for b in blocks:
        if not isinstance(b, dict):
            continue
        t = b.get("type")
        c = colour(b.get("color"))
        try:
            if t == "title":
                d.title(_s(b.get("text")))
            elif t == "h1":
                d.h1(_s(b.get("text")))
            elif t == "h2":
                d.h2(_s(b.get("text")))
            elif t == "item":
                d.item(_s(b.get("marker")) or "•", _s(b.get("text")), level=int(_num(b.get("level"), 0)),
                       color=c)
            elif t == "example":
                d.example(_s(b.get("text")))
            elif t == "note":
                d.note(_s(b.get("text")), label=_s(b.get("label")) or "NOTE")
            elif t == "qa":
                d.qa(_s(b.get("q")), _s(b.get("a")))
            elif t == "compare":
                pairs = [(_s(p[0]), _s(p[1])) for p in b.get("pairs") or []
                         if isinstance(p, (list, tuple)) and len(p) >= 2]
                d.compare(_s(b.get("a_title")), _s(b.get("b_title")), pairs)
            elif t == "figure":
                d.figure(build_figure(b), _s(b.get("caption")) or None)
            else:  # "para" and anything unknown
                text = _s(b.get("text"))
                if text.strip():
                    d.para(text, bold=bool(b.get("bold")), color=c,
                           align="center" if b.get("align") == "center" else None)
        except Exception:
            text = _s(b.get("text") or b.get("q") or "")
            if text.strip():
                d.para(re.sub(r"[$*]", "", text))
    return d
