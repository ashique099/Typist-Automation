# Typist Automation

Turns live-class "panel" PDFs (handwritten slides exported as images) into clean, editable Word files.

## Website (for the whole typist team)

Typists log in with a shared team password, drop one or more panel PDFs (order can be changed with
↑ ↓), and download one combined `.docx` in the house style. Every slide is read by Claude
(vision), returned as structured blocks (`typist/transcribe.py`), and laid out by `typist/blocks.py`
on top of `docxkit` (Word equations, vector figures, coloured points, 2 columns).

Settings are environment variables, never committed:

| Variable | Meaning |
|---|---|
| `ANTHROPIC_API_KEY` | Claude API key (required) |
| `APP_PASSWORD` | shared team password |
| `TYPIST_MODEL` | default `claude-opus-5-5` |
| `DATA_DIR` | uploads + outputs (default `./data`), jobs auto-deleted after `KEEP_DAYS` (30) |
| `JOB_WORKERS` | conversions at the same time (default 2) |

Run it:

* **Office Windows PC** (typists on the same Wi-Fi/LAN): double-click `start-windows.bat`, fill the
  key and password in the `settings.bat` it opens, run again, open `http://<PC-IP>:8000`.
* **Any server / cloud (Render, Railway, a VPS)**: `docker build -t typist . && docker run -p 8000:8000
  -e ANTHROPIC_API_KEY=... -e APP_PASSWORD=... -v typist-data:/data typist`
* **Dev**: `pip install -r requirements.txt && uvicorn web.app:app --reload`

Command line, same pipeline: `python -m typist.convert p1.pdf p2.pdf -o Notes.docx --save-json b.json`
(`--from-json b.json` re-renders without calling the API). Tests: `pytest tests` (offline).

## Pipeline (first job: Solid State, Parts 1-4)

1. **Inspect** – `pdfinfo` / `pdfimages -list`: panel PDFs are one JPEG per page, no text layer.
2. **Extract** – `pdfimages -j in.pdf img/p` gives one image per slide.
3. **Transcribe** – read each slide (handwritten Hindi + English) into structured content
   (`typist/solid_state.py`): headings, definitions, numbered points, Q&A, tables, notes.
   Ads/announcements are kept in a grey "class notice" box so nothing is lost.
4. **Build** – `typist/docxkit.py`:
   * Hindi-ready styles (Nirmala UI for Devanagari, complex-script bold/size set).
   * Inline markup: `**bold**`, `$NaNO_3$` → native Word equation (OMML), Unicode text.
     Syntax: `x_2`, `x^{2+}`, `x_{2(s)}`, `\frac{a}{b}`, `{group}`.
   * `Figure` → hand drawings redrawn as an editable DrawingML shape group (vector).
5. **House style** (`docxkit.Style`, the user's standing rules): Sanskrit Text 13 pt;
   colours red = topics, blue = sub-topics/labels, black = body, green = examples/notes;
   A4 portrait, narrow margins, 2 columns; no tables (comparisons as coloured points);
   no thumbnails, course ads, class timings or greetings.
6. **Check** – convert with LibreOffice to PDF and look at every page.

```
pip install python-docx
python -m typist.solid_state out/
```
