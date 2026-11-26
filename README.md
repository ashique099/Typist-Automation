# Disha Online Solution (Typist Automation)

Turns live-class "panel" PDFs (handwritten slides exported as images) into clean, editable Word files.

## Website (for the whole typist team)

After signing in, each typist sees a dashboard with three tools:

1. **Panel PDF to Notes DOCX**: drop a chapter's panel PDFs (order can be changed with ↑ ↓) and get
   one combined `.docx` in the house style. Repeated slides are skipped, every slide is read by
   Claude (vision, `typist/transcribe.py`) and laid out by `typist/blocks.py` on top of `docxkit`
   (Word equations, vector figures, coloured points, 2 columns).
2. **DOCX to PPT**: drop the (corrected) notes `.docx`; `typist/reader.py` reads it back into blocks
   (colours, equations, and our figures from their alt text) and `typist/pptkit.py` builds slides in
   the chosen PPT design. No API call, so it is free.
3. **PPT or DOCX to MCQ Test**: drop notes or slides; Claude writes the MCQs (`typist/mcq.py`), answers
   are spread evenly over options 1-4, output is the Word table test plus an Objective PPT.

Accounts, history and the **Admin panel** at `/admin` (`web/templates/admin.html`, data in `web/db.py`, SQLite in `DATA_DIR/typist.db`):

* The first visit creates the **admin** account. The admin adds a login for each typist, resets
  passwords, and disables accounts (Typists page).
* **Work history** keeps every job: who, when, which files, pages, MCQs, real API cost (from token
  counts), and the downloads. Admins see everyone's work; typists see their own.
* **Overview** (admin): spend vs budget, today's and this month's jobs, a 30-day spending chart, totals per tool and per typist, running and failed jobs (with Run again), system status. Work history can be filtered by typist/tool/status/month and exported as CSV.
* **Settings** (admin): monthly budget in ₹ (AI jobs are blocked once it is used up), $→₹ rate,
  model for notes and for MCQs, effort. **PPT designs** (admin): upload any class PPT; its master
  background is kept and its slides are dropped.

Forgot the admin password? On the server PC run `python -m web.manage users` (lists usernames) and
`python -m web.manage reset USERNAME` (sets and prints a temporary password). Passwords are stored
hashed, so they can only be replaced, never read back.

Environment variables (never committed):

| Variable | Meaning |
|---|---|
| `ANTHROPIC_API_KEY` | Claude API key (needed for tools 1 and 3) |
| `DATA_DIR` | database, uploads, outputs, PPT designs (default `./data`) |
| `KEEP_DAYS` | job files older than this are deleted; history rows stay (default 90) |
| `JOB_WORKERS` | jobs running at the same time (default 2) |

Run it:

* **Office Windows PC** (typists on the same Wi-Fi/LAN): double-click `start-windows.bat`, put the API
  key in the `settings.bat` it opens, run again, open `http://localhost:8000` and create the admin.
  Other PCs open `http://<PC-IP>:8000`.
* **Any server / cloud**: `docker build -t typist . && docker run -p 8000:8000
  -e ANTHROPIC_API_KEY=... -v typist-data:/data typist`
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
