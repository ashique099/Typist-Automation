"""Offline tests: no API key needed (the vision call is replaced by a fixture)."""
import json
import time
import zipfile
from pathlib import Path

import pytest

HERE = Path(__file__).parent
PARTS = json.loads((HERE / "sample_parts.json").read_text(encoding="utf-8"))


def _doc_xml(path):
    return zipfile.ZipFile(path).read("word/document.xml").decode()


def test_render_house_style(tmp_path):
    from typist.convert import build_docx
    out = build_docx(PARTS, tmp_path / "n.docx")
    xml = _doc_xml(out)
    assert xml.count("ठोस अवस्था (Solid State)") == 1          # title once for same chapter
    assert "<m:oMath" in xml and "<wpg:wgp>" in xml              # equations + vector figure
    assert "<w:tbl>" not in xml                                  # no tables
    assert 'w:num="2"' in xml                                    # two columns
    assert "अज्ञात block" in xml                                 # unknown block kept as para


def test_pdf_pages(tmp_path):
    import pymupdf
    from typist.pages import pdf_pages
    doc = pymupdf.open()
    doc.new_page(width=960, height=540).insert_text((50, 50), "slide")
    doc.save(tmp_path / "a.pdf")
    imgs = pdf_pages(str(tmp_path / "a.pdf"))
    assert len(imgs) == 1 and imgs[0][:2] == b"\xff\xd8"


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setenv("DATA_DIR", str(tmp_path))
    monkeypatch.setenv("ANTHROPIC_API_KEY", "test-key")
    import importlib
    import typist.transcribe as tr
    calls = iter(PARTS)
    monkeypatch.setattr(tr, "transcribe_pdf_pages",
                        lambda images, ctx, progress=None, label="": (lambda p: (p["chapter_title"], p["blocks"]))(next(calls)))
    import typist.mcq as mq
    monkeypatch.setattr(mq, "generate_mcqs", lambda parts, n, **k: [
        {"q": f"प्रश्न {i} $H_2O$?", "options": ["a", "b", "c", "d"], "answer": 1 + i % 4, "solution": ""}
        for i in range(n)])
    import web.app as appmod
    importlib.reload(appmod)
    from fastapi.testclient import TestClient
    with TestClient(appmod.app) as c:
        yield c


def _pdf_bytes():
    import pymupdf
    doc = pymupdf.open()
    doc.new_page()
    return doc.tobytes()


def _wait(c, jid):
    for _ in range(150):
        m = c.get(f"/api/jobs/{jid}").json()
        if m["status"] in ("done", "failed"):
            return m
        time.sleep(0.1)
    return m


def test_web_flow(client, tmp_path):
    c = client
    # first visit -> create the admin
    assert c.get("/", follow_redirects=False).headers["location"] == "/setup"
    r = c.post("/setup", data={"name": "Ashique", "username": "admin", "password": "secret1", "password2": "secret1"},
               follow_redirects=False)
    assert r.status_code == 303 and "ta_session" in r.cookies
    assert "Panel PDF to Notes DOCX" in c.get("/").text
    assert c.get("/setup", follow_redirects=False).headers["location"] == "/login"
    # admin adds a typist
    assert c.post("/api/admin/users", data={"username": "rahul", "name": "Rahul", "password": "abc123"}).status_code == 200
    assert c.post("/api/admin/users", data={"username": "rahul", "name": "X", "password": "abc123"}).status_code == 400
    c.get("/logout")
    assert c.post("/api/jobs", data={"task": "ppt"}, files=[("files", ("a.docx", b"PK"))]).status_code == 401
    assert "error" in c.post("/login", data={"username": "rahul", "password": "nope"},
                             follow_redirects=False).headers["location"]
    c.post("/login", data={"username": "Rahul", "password": "abc123"})
    me = c.get("/api/me").json()
    assert me["user"]["name"] == "Rahul" and me["user"]["role"] == "typist" and me["api_key"]
    assert c.get("/api/admin/users").status_code == 403
    assert c.get("/admin", follow_redirects=False).headers["location"] == "/"

    # 1. PDFs -> notes
    r = c.post("/api/jobs", data={"task": "notes"},
               files=[("files", ("Part-1.pdf", _pdf_bytes())), ("files", ("Part-2.pdf", _pdf_bytes()))])
    m = _wait(c, r.json()["id"])
    assert m["status"] == "done", m
    assert [d["kind"] for d in m["downloads"]] == ["notes"] and m["by_name"] == "Rahul" and m["pages"] == 2
    d = c.get(f"/api/jobs/{m['id']}/download?kind=notes")
    assert d.status_code == 200 and d.content[:2] == b"PK" and "Solid" in d.headers["content-disposition"]
    assert c.get(f"/api/jobs/{m['id']}/download?kind=ppt").status_code == 404
    notes = d.content
    assert c.post("/api/jobs", data={"task": "notes"}, files=[("files", ("x.pdf", b"not a pdf"))]).status_code == 400

    # 2. DOCX -> PPT (no API)
    r = c.post("/api/jobs", data={"task": "ppt"}, files=[("files", ("Notes.docx", notes))])
    m = _wait(c, r.json()["id"])
    assert m["status"] == "done", m
    p = c.get(f"/api/jobs/{m['id']}/download?kind=ppt")
    assert p.status_code == 200 and p.content[:2] == b"PK"
    assert c.post("/api/jobs", data={"task": "ppt"}, files=[("files", ("a.pdf", _pdf_bytes()))]).status_code == 400

    # 3. PPT / DOCX -> MCQ test
    r = c.post("/api/jobs", data={"task": "test", "mcq_count": "8"}, files=[("files", ("Deck.pptx", p.content))])
    m = _wait(c, r.json()["id"])
    assert m["status"] == "done" and m["mcqs"] == 8, m
    t = c.get(f"/api/jobs/{m['id']}/download?kind=test")
    assert t.status_code == 200 and "MCQ%20Test" in t.headers["content-disposition"]
    assert c.get(f"/api/jobs/{m['id']}/download?kind=objective").status_code == 200

    # history filters
    h = c.get("/api/jobs?task=ppt").json()
    assert h["total"] == 1 and h["jobs"][0]["task"] == "ppt"
    assert c.get("/api/jobs?q=Deck").json()["total"] == 1

    # admin sees everyone's work and the month's spend; budget blocks AI jobs
    c.get("/logout")
    c.post("/login", data={"username": "admin", "password": "secret1"})
    assert c.get("/api/jobs?scope=all").json()["total"] == 3
    assert c.get("/api/jobs?scope=mine").json()["total"] == 0
    assert c.get("/api/admin/settings").json()["month"]["jobs"] == 3
    assert next(u for u in c.get("/api/admin/users").json() if u["username"] == "rahul")["jobs"] == 3
    # admin panel
    assert c.get("/admin").status_code == 200
    o = c.get("/api/admin/overview").json()
    assert o["month"]["jobs"] == 3 and o["today"]["jobs"] == 3 and len(o["series"]) == 30
    assert o["series"][-1]["jobs"] == 3 and {t["task"] for t in o["by_tool"]} == {"notes", "ppt", "test"}
    rid = next(u for u in o["users"] if u["username"] == "rahul")["id"]
    assert c.get(f"/api/jobs?scope=all&user={rid}").json()["total"] == 3
    import time as _t
    assert c.get(f"/api/jobs?scope=all&month={_t.strftime('%Y-%m')}").json()["total"] == 3
    assert c.get("/api/jobs?scope=all&month=2001-01").json()["total"] == 0
    csv = c.get("/api/admin/jobs.csv?task=notes")
    assert csv.status_code == 200 and csv.text.count("\n") == 2 and "Rahul" in csv.text
    c.post(f"/api/admin/users/{rid}", data={"name": "Rahul K"})
    assert any(u["name"] == "Rahul K" and u["all_jobs"] == 3 for u in c.get("/api/admin/users").json())
    me_id = c.get("/api/me").json()["user"]["id"]
    assert c.post(f"/api/admin/users/{me_id}", data={"active": "0"}).status_code == 400
    jid = h["jobs"][0]["id"]
    assert c.post(f"/api/admin/jobs/{jid}/retry").status_code == 400  # not failed
    assert c.post(f"/api/admin/jobs/{jid}/delete").status_code == 200
    assert c.get(f"/api/jobs/{jid}").json()["downloads"] == []
    c.post("/api/admin/settings", data={"monthly_budget_inr": "0"})
    r = c.post("/api/jobs", data={"task": "notes"}, files=[("files", ("a.pdf", _pdf_bytes()))])
    assert r.status_code == 400 and "budget" in r.json()["detail"]
    r = c.post("/api/jobs", data={"task": "ppt"}, files=[("files", ("Notes.docx", notes))])
    assert r.status_code == 200  # PPT costs nothing, so it still runs


def test_docx_round_trip(tmp_path):
    from typist.convert import build_docx
    from typist.reader import docx_blocks
    build_docx(PARTS, tmp_path / "n.docx")
    title, blocks = docx_blocks(tmp_path / "n.docx")
    assert title == "ठोस अवस्था (Solid State)"
    orig = [b for p in PARTS for b in p["blocks"] if b.get("type") == "figure"]
    figs = [b for b in blocks if b["type"] == "figure"]
    assert len(figs) == len(orig) and figs[0]["shapes"] == orig[0]["shapes"]
    assert any(b["type"] == "h1" for b in blocks) and any("$" in (b.get("text") or "") for b in blocks)


def test_ppt_deck(tmp_path):
    from pptx import Presentation
    from typist.pptkit import render_deck
    out = render_deck(PARTS, tmp_path / "d.pptx")
    prs = Presentation(out)
    assert len(prs.slides) >= 3
    text = " ".join(sh.text_frame.text for s in prs.slides for sh in s.shapes if sh.has_text_frame)
    assert "ठोस अवस्था (Solid State)" in text and "उदाहरण" in text


def test_template_deck(tmp_path):
    from pptx import Presentation
    from typist.pptkit import render_template_deck
    tpl = tmp_path / "tpl.pptx"
    p = Presentation()
    p.slides.add_slide(p.slide_layouts[1])  # existing slides must be dropped
    p.save(tpl)
    out = render_template_deck(PARTS, tmp_path / "t.pptx", str(tpl))
    prs = Presentation(out)
    text = " ".join(sh.text_frame.text for s in prs.slides for sh in s.shapes if sh.has_text_frame)
    assert "ठोस (Solid) :-" in text and "Hg" in text


def test_mcq_docx(tmp_path):
    from docx import Document
    from typist.mcq import balance_answers, write_mcq_docx
    qs = balance_answers([{"q": f"Q{i} $NaCl$", "options": ["w", "x", "y", "z"], "answer": 1} for i in range(8)])
    assert sorted(q["answer"] for q in qs) == [1, 1, 2, 2, 3, 3, 4, 4]
    assert all(q["options"][q["answer"] - 1] == "w" for q in qs)
    d = Document(write_mcq_docx(qs, tmp_path / "t.docx"))
    assert len(d.tables) == 8
    assert [r.cells[0].text for r in d.tables[0].rows] == [
        "Question", "Type", "Option", "Option", "Option", "Option", "Answer", "Solution", "Positive Marks",
        "Negative Marks"]
    assert d.tables[0].rows[1].cells[1].text == "Multiple Choice"
