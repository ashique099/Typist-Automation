"""MCQ tests from the notes, in the team's Word table format (one 2-column table per question):

    Question | Type (Multiple Choice) | Option x4 | Answer (1-4) | Solution | Positive Marks 1 | Negative Marks 0

Format copied from the team's "SENTUP EXAM" file: A4, margins 1.27 cm sides / 0.88 cm top-bottom,
Sanskrit Text 15 pt, black, table 30/70, single borders, one empty line between tables.

MCQ JSON item: {"q": "...", "options": ["..","..","..",".."], "answer": 1-4, "solution": ""}
Text may use the same $H_2O$ markup as the notes (becomes a Word equation).
"""
from __future__ import annotations

import json
import os

from docx import Document
from docx.oxml import parse_xml
from docx.shared import Cm, Pt

from . import docxkit
from .docxkit import _ns, add_inline

FONT, SIZE = "Sanskrit Text", 15


def _cell_text(cell, text):
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    add_inline(p, str(text), size=SIZE, color="000000")


def write_mcq_docx(mcqs, path):
    old = docxkit.HINDI_FONT, docxkit.MATH_SIZE
    docxkit.HINDI_FONT, docxkit.MATH_SIZE = FONT, SIZE
    try:
        d = Document()
        s = d.sections[0]
        s.page_width, s.page_height = Cm(21), Cm(29.7)
        s.left_margin = s.right_margin = Cm(1.27)
        s.top_margin = s.bottom_margin = Cm(0.88)
        st = d.styles["Normal"]
        st.font.name, st.font.size = FONT, Pt(SIZE)
        rpr = st.element.get_or_add_rPr()
        rpr.get_or_add_rFonts().set(f"{{{docxkit.NS['w']}}}cs", FONT)
        st.paragraph_format.space_after = Pt(0)
        for m in mcqs:
            opts = list(m.get("options") or [])[:4] + [""] * (4 - len(m.get("options") or []))
            rows = [("Question", m.get("q", "")), ("Type", "Multiple Choice")] + \
                   [("Option", o) for o in opts] + \
                   [("Answer", m.get("answer", "")), ("Solution", m.get("solution", "") or ""),
                    ("Positive Marks", m.get("positive", 1)), ("Negative Marks", m.get("negative", 0))]
            t = d.add_table(rows=len(rows), cols=2)
            tbl = t._tbl
            tblPr = tbl.tblPr
            for el in list(tblPr):
                if el.tag.endswith(("}tblStyle", "}tblW")):
                    tblPr.remove(el)
            tblPr.append(parse_xml(f'<w:tblW {_ns("w")} w:w="5000" w:type="pct"/>'))
            tblPr.append(parse_xml(
                f'<w:tblBorders {_ns("w")}>' + "".join(
                    f'<w:{e} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                    for e in ("top", "left", "bottom", "right", "insideH", "insideV")) + "</w:tblBorders>"))
            for r, (k, v) in zip(t.rows, rows):
                r._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {_ns("w")}/>'))
                for c, pct in zip(r.cells, (1500, 3500)):
                    tcPr = c._tc.get_or_add_tcPr()
                    tcPr.append(parse_xml(f'<w:tcW {_ns("w")} w:w="{pct}" w:type="pct"/>'))
                    tcPr.append(parse_xml(
                        f'<w:tcMar {_ns("w")}><w:top w:w="50" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>'
                        f'<w:bottom w:w="50" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>'))
                _cell_text(r.cells[0], k)
                _cell_text(r.cells[1], v)
            d.add_paragraph()
        d.save(path)
    finally:
        docxkit.HINDI_FONT, docxkit.MATH_SIZE = old
    return path


# ------------------------------------------------------------ generation ---
MCQ_SYSTEM = """You write MCQ tests for Bihar Board (BSEB) Hindi-medium students from a teacher's class notes.
Rules:
- Every question must be answerable from the given notes; cover all topics evenly, important points first.
- Hindi (Devanagari) like the notes; keep English terms in brackets where the notes do.
- Exactly 4 short options, one clearly correct, distractors plausible and from the same topic.
  Spread the correct option across positions 1-4. No "all of these"/"none of these" unless natural.
- Chemical formulas and maths go inside $...$ with _ for subscript and ^ for superscript: $H_2O$, $SO_4^{2-}$.
- "answer" is the number (1-4) of the correct option.
Return ONLY JSON: {"mcqs":[{"q":"...","options":["...","...","...","..."],"answer":2,"solution":""}]}"""


def balance_answers(mcqs, seed=7):
    """Move each correct option so answers are spread evenly over 1-4 (order of the rest kept)."""
    import random
    rnd = random.Random(seed)
    slots = ([1, 2, 3, 4] * (len(mcqs) // 4 + 1))[:len(mcqs)]
    rnd.shuffle(slots)
    out = []
    for m, target in zip(mcqs, slots):
        opts = list(m["options"])
        if any(w in " ".join(opts) for w in ("इनमें से", "उपर्युक्त", "सभी", "None", "All of")):
            out.append(m)  # keep 'none/all of these' in its place
            continue
        right = opts.pop(int(m["answer"]) - 1)
        opts.insert(target - 1, right)
        out.append({**m, "options": opts, "answer": target})
    return out


def _slim(parts):
    out = []
    for p in parts:
        for b in p.get("blocks") or []:
            if b.get("type") == "figure":
                if b.get("caption"):
                    out.append({"type": "figure", "caption": b["caption"]})
            else:
                out.append(b)
    return out


def generate_mcqs(parts, count=50, solution=False, model=None):
    """Text-only call (no images), so it is cheap. Needs ANTHROPIC_API_KEY."""
    from .transcribe import _client, _parse, TranscribeError
    client = _client()
    model = model or os.environ.get("TYPIST_MCQ_MODEL") or os.environ.get("TYPIST_MODEL", "claude-opus-5-5")
    notes = json.dumps(_slim(parts), ensure_ascii=False)
    ask = (f"NOTES (JSON blocks):\n{notes}\n\nWrite exactly {count} MCQs. "
           + ("Give a one-line Hindi explanation in \"solution\"." if solution else 'Leave "solution" empty.'))
    with client.beta.messages.stream(
        model=model, max_tokens=64000, system=MCQ_SYSTEM,
        messages=[{"role": "user", "content": ask}],
        output_config={"effort": os.environ.get("TYPIST_EFFORT", "high")},
        betas=["server-side-fallback-2026-07-01"], fallbacks="default",
    ) as stream:
        msg = stream.get_final_message()
    if msg.stop_reason == "refusal":
        raise TranscribeError("Model ne MCQ banane se mana kar diya.")
    data = _parse("".join(b.text for b in msg.content if b.type == "text").replace('"mcqs"', '"blocks"', 1))
    good = []
    for m in data["blocks"]:
        try:
            opts = [str(o) for o in m["options"]][:4]
            ans = int(m["answer"])
            if len(opts) == 4 and 1 <= ans <= 4 and str(m["q"]).strip():
                good.append({"q": str(m["q"]), "options": opts, "answer": ans,
                             "solution": str(m.get("solution") or "")})
        except (KeyError, TypeError, ValueError):
            continue
    return balance_answers(good)
