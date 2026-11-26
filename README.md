# Typist Automation

Turns live-class "panel" PDFs (handwritten slides exported as images) into clean, editable Word files.

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
5. **Check** – convert with LibreOffice to PDF and look at every page.

```
pip install python-docx
python -m typist.solid_state out/
```
