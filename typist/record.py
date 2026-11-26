"""Turn a hand-written content module (like solid_state.py) into the JSON blocks the
website uses, so old chapters can be re-rendered as PPT / tests without the API.

    python -m typist.record solid_state out.json
"""
from __future__ import annotations

import importlib
import json
import sys


class RecFigure:
    """Stands in for docxkit.Figure and records shapes as JSON."""

    def __init__(self, w, h):
        self.w, self.h, self.shapes = w, h, []

    def rect(self, x, y, w, h, color="000000", width=1.0, fill=None, text=None, size=11, bold=False, **_):
        self.shapes.append(dict(shape="rect", x=x, y=y, w=w, h=h, color=color, line_width=width, fill=fill,
                                text=text, size=size, bold=bold))

    def text(self, x, y, w, h, text, size=11, bold=False, color="000000", **_):
        self.shapes.append(dict(shape="text", x=x, y=y, w=w, h=h, text=text, size=size, bold=bold, color=color))

    def circle(self, cx, cy, r, color="000000", width=1.0, fill=None, text=None, size=11, bold=False, **_):
        self.shapes.append(dict(shape="circle", cx=cx, cy=cy, r=r, color=color, width=width, fill=fill,
                                text=text, size=size, bold=bold))

    def line(self, x1, y1, x2, y2, color="000000", width=1.0, dash=None, arrow=False):
        self.shapes.append(dict(shape="arrow" if arrow else "line", x1=x1, y1=y1, x2=x2, y2=y2, color=color,
                                width=width, dash=bool(dash)))

    def arrow(self, x1, y1, x2, y2, **kw):
        self.line(x1, y1, x2, y2, arrow=True, **kw)


class Recorder:
    """Same methods as docxkit.NoteDoc, but collects blocks."""

    def __init__(self):
        self.blocks = []

    def _b(self, **kw):
        self.blocks.append({k: v for k, v in kw.items() if v not in (None, "")})

    def title(self, text, sub=None):
        self._b(type="title", text=text)

    def h1(self, text, **_):
        self._b(type="h1", text=text)

    def h2(self, text, **_):
        self._b(type="h2", text=text)

    def para(self, text, bold=False, indent=0.0, align=None, color=None, **_):
        self._b(type="para", text=text, bold=bold or None, color=color, align=align)

    def example(self, text):
        self._b(type="example", text=text)

    def item(self, marker, text, level=0, color=None):
        self._b(type="item", marker=marker, text=text, level=level or None, color=color)

    def compare(self, a, b, pairs):
        self._b(type="compare", a_title=a, b_title=b, pairs=[list(p) for p in pairs])

    def qa(self, q, a):
        self._b(type="qa", q=q, a=a)

    def note(self, text, label="NOTE"):
        self._b(type="note", text=text, label=label)

    def figure(self, fig, caption=None):
        self._b(type="figure", width=fig.w, height=fig.h, caption=caption, shapes=fig.shapes)

    def page_break(self):
        pass


def record(module_name):
    mod = importlib.import_module(f"typist.{module_name}")
    figs = importlib.import_module(f"typist.{module_name}_figures")
    figs.Figure = RecFigure
    parts = []
    for n, (name, fn) in sorted(mod.PARTS.items()):
        r = Recorder()
        fn(r)
        parts.append({"name": name, "chapter_title": getattr(mod, "TITLE", ""), "blocks": r.blocks})
    return parts


if __name__ == "__main__":
    parts = record(sys.argv[1])
    with open(sys.argv[2], "w", encoding="utf-8") as f:
        json.dump(parts, f, ensure_ascii=False, indent=1)
    print("wrote", sys.argv[2], sum(len(p["blocks"]) for p in parts), "blocks")
