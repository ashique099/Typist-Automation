"""Slide images -> JSON blocks, using Claude's vision.

The API key is read from the ANTHROPIC_API_KEY environment variable (never stored in the repo).
Model can be changed with TYPIST_MODEL (default claude-opus-5-5).
"""
from __future__ import annotations

import base64
import json
import os
import re

import anthropic

MODEL = os.environ.get("TYPIST_MODEL", "claude-opus-5-5")
EFFORT = os.environ.get("TYPIST_EFFORT", "high")
PAGES_PER_CALL = 10

SYSTEM = r"""You are an expert Hindi-medium chemistry typist. You receive slide images from a live-class
"panel" PDF (handwritten Hindi + English chemistry notes, Bihar Board Inter level). Type them into clean,
complete study notes and return them as JSON blocks that a program turns into a Word file.

WHAT TO TYPE
- Type every piece of teaching content: definitions, points, examples, reactions, formulas, Q&A,
  notes, figures. Do not summarise and do not skip lines. Fix obvious spelling slips; keep the
  teacher's wording. Use correct Devanagari (Unicode) and proper Hindi punctuation (। –).
- Do NOT type: thumbnails / cover slides, channel or course names, batch / fee / admission / app
  details, class timings, "next class tomorrow", greetings (good night, thank you, like/share/subscribe),
  teacher or institute promotions, slide numbers, watermarks.
- The same slide may repeat while the teacher writes on it. Type the content once, in its final form.
- If a later slide continues a list or topic started earlier (see PREVIOUS CONTEXT), continue it:
  don't repeat its heading and keep the numbering going.

BLOCK TYPES  (return them in reading order)
  {"type":"h1","text":"..."}            main topic heading (printed red). Keep teacher's "#" or "★" if written.
  {"type":"h2","text":"..."}            sub-topic / label heading (printed blue)
  {"type":"para","text":"..."}          normal paragraph (black). Optional "color":"red|blue|green", "bold":true
  {"type":"item","marker":"(i)","text":"...","level":0}   numbered / bulleted point. Use the teacher's
                                        marker style: (i) (ii)…, 1. 2.…, (a) (b)…, •. level 1 = sub-point.
  {"type":"example","text":"..."}       an "उदाहरण / Ex." line: give only the examples after the label
  {"type":"note","text":"...","label":"NOTE"}   Note / ध्यान दें / याद रखें boxes (label as written)
  {"type":"qa","q":"...","a":"..."}     question with its answer (प्रश्न / उत्तर)
  {"type":"compare","a_title":"...","b_title":"...","pairs":[["...","..."],...]}
                                        any table or side-by-side difference (अंतर). NEVER output a table any
                                        other way: each row becomes one pair.
  {"type":"figure","width":W,"height":H,"caption":"चित्र : ...","shapes":[...]}
                                        a redraw of a hand-drawn figure / diagram / flow chart / structure.

TEXT MARKUP (inside any text field)
- **bold** for key words the teacher underlined or stressed (the defined term in a definition).
- Every chemical formula, ion, equation and maths expression goes inside $...$ so it becomes a Word
  equation:  $NaCl$, $H_2O_{(s)}$, $SO_4^{2-}$, $Fe^{3+}$, $\frac{a}{b}$, $x^2$, ${CuSO_4}.{5H_2O}$.
  Syntax: _ subscript, ^ superscript, {...} groups, \frac{num}{den}. Write arrows and symbols as
  Unicode inside the equation: $2H_2 + O_2 → 2H_2O$, $⇌$, $Δ$, $°C$, $≈$, $×$, $δ^+$.
  Words (Hindi or English) never go inside $...$. Keep $ balanced.

FIGURES  (vector redraw; coordinates in points, origin top-left, width ≤ 260)
  shapes: {"shape":"rect","x":..,"y":..,"w":..,"h":..,"color":"blue","text":"label","size":10,"bold":false}
          {"shape":"text","x":..,"y":..,"w":..,"h":..,"text":"...","size":10,"color":"black"}
          {"shape":"circle","cx":..,"cy":..,"r":..,"color":"red","fill":"white","text":"+"}
          {"shape":"line","x1":..,"y1":..,"x2":..,"y2":..,"color":"black","width":1,"dash":false}
          {"shape":"arrow", same as line, arrow head at x2,y2}
  Colours: red, blue, green, black. Keep it clean and readable (text size 9-12, boxes big enough for
  their text, use "\n" for line breaks inside a box). Reproduce structure faithfully: lattices, flow
  charts (box + arrows), unit cells, molecule chains, graphs (axes as arrows + curve as short lines).
  If something is a simple classification tree, a figure is better than text.

OUTPUT
Return ONLY one JSON object, no commentary, no code fence:
{"chapter_title": "<chapter name as Hindi (English), e.g. ठोस अवस्था (Solid State), or empty>",
 "blocks": [ ... ]}

EXAMPLE of the expected style (from a finished chapter):
{"chapter_title":"ठोस अवस्था (Solid State)","blocks":[
 {"type":"h1","text":"# ठोस (Solid)"},
 {"type":"para","text":"ठोस पदार्थ की वह अवस्था है, जिसका आयतन, आकार एवं द्रव्यमान निश्चित होता है।"},
 {"type":"example","text":"ईंट, कलम, मकान, पत्थर, लकड़ी आदि।"},
 {"type":"h1","text":"(i) क्रिस्टलीय ठोस / रवादार ठोस"},
 {"type":"para","text":"वैसा ठोस जिसका एक निश्चित ज्यामितीय आकार होता है तथा जिसमें कणों की दीर्घ परासी व्यवस्था होती है, उसे **क्रिस्टलीय ठोस** कहते हैं।"},
 {"type":"example","text":"$NaCl$ (साधारण नमक), $CsCl$, लोहा, हीरा, आयोडीन ($I_2$), $ZnS$ etc."},
 {"type":"note","text":"पारा ($Hg$) धातु है लेकिन यह द्रव अवस्था में पाया जाता है।"},
 {"type":"item","marker":"(i)","text":"इसका एक निश्चित ज्यामितीय आकार होता है।"},
 {"type":"qa","q":"काँच को अतिशीतित द्रव कहा जाता है, क्यों?","a":"..."},
 {"type":"figure","width":250,"height":120,"caption":"चित्र : पदार्थ का वर्गीकरण","shapes":[
   {"shape":"rect","x":80,"y":4,"w":90,"h":28,"color":"red","text":"पदार्थ","size":12,"bold":true},
   {"shape":"line","x1":125,"y1":32,"x2":125,"y2":48},{"shape":"line","x1":62,"y1":48,"x2":188,"y2":48},
   {"shape":"arrow","x1":62,"y1":48,"x2":62,"y2":66},{"shape":"arrow","x1":188,"y1":48,"x2":188,"y2":66},
   {"shape":"rect","x":2,"y":68,"w":120,"h":48,"color":"blue","text":"भौतिक वर्गीकरण\n(Physical\nClassification)","size":9.5},
   {"shape":"rect","x":128,"y":68,"w":120,"h":48,"color":"blue","text":"रासायनिक वर्गीकरण\n(Chemical\nClassification)","size":9.5}]}
]}"""


class TranscribeError(RuntimeError):
    pass


def _client():
    if not os.environ.get("ANTHROPIC_API_KEY"):
        raise TranscribeError("ANTHROPIC_API_KEY set nahi hai (server ki settings me API key daalein).")
    return anthropic.Anthropic(max_retries=4, timeout=900)


def _parse(text: str) -> dict:
    text = text.strip()
    m = re.search(r"```(?:json)?\s*(.*?)```", text, re.S)
    if m:
        text = m.group(1)
    start, end = text.find("{"), text.rfind("}")
    if start < 0 or end < 0:
        raise ValueError("no JSON object in reply")
    data = json.loads(text[start:end + 1])
    if not isinstance(data.get("blocks"), list):
        raise ValueError("reply has no 'blocks' list")
    return data


def transcribe_pages(images: list[bytes], context_blocks=None, label="") -> dict:
    """One API call for up to PAGES_PER_CALL slide images. Returns {"chapter_title", "blocks"}."""
    client = _client()
    content = []
    for k, img in enumerate(images, 1):
        content.append({"type": "text", "text": f"Slide {k}:"})
        content.append({"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                                                    "data": base64.standard_b64encode(img).decode()}})
    slim = [{"type": "figure", "caption": b.get("caption", "")} if b.get("type") == "figure" else b
            for b in (context_blocks or [])]
    ctx = json.dumps(slim, ensure_ascii=False)
    content.append({"type": "text", "text":
                    f"PREVIOUS CONTEXT (last blocks already typed, do not repeat them): {ctx}\n\n"
                    f"Type the {len(images)} slides above {label}. Return only the JSON object."})
    last_err = None
    for attempt in range(2):
        msgs = [{"role": "user", "content": content}]
        if attempt:
            msgs += [{"role": "assistant", "content": last_text},
                     {"role": "user", "content": f"That reply was not valid JSON ({last_err}). "
                                                 "Send the complete JSON object again, nothing else."}]
        with client.beta.messages.stream(
            model=MODEL,
            max_tokens=64000,
            system=SYSTEM,
            messages=msgs,
            output_config={"effort": EFFORT},
            betas=["server-side-fallback-2026-07-01"],
            fallbacks="default",
        ) as stream:
            msg = stream.get_final_message()
        if msg.stop_reason == "refusal":
            raise TranscribeError("Model ne ye slides process karne se mana kar diya.")
        last_text = "".join(b.text for b in msg.content if b.type == "text")
        try:
            return _parse(last_text)
        except (ValueError, json.JSONDecodeError) as e:
            last_err = str(e)[:200]
    raise TranscribeError(f"Model ka jawab JSON me nahi aaya: {last_err}")


def transcribe_pdf_pages(images: list[bytes], context_blocks=None, progress=None, label=""):
    """All pages of one PDF, in chunks, carrying context from chunk to chunk."""
    title, blocks = "", []
    ctx = list(context_blocks or [])
    for i in range(0, len(images), PAGES_PER_CALL):
        chunk = images[i:i + PAGES_PER_CALL]
        if progress:
            progress(i, len(images))
        res = transcribe_pages(chunk, ctx[-15:], label=label)
        title = title or (res.get("chapter_title") or "").strip()
        new = [b for b in res["blocks"] if isinstance(b, dict)]
        blocks += new
        ctx += new
    return title, blocks
