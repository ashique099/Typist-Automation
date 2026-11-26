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
    monkeypatch.setenv("APP_PASSWORD", "pw")
    import importlib
    import typist.transcribe as tr
    calls = iter(PARTS)
    monkeypatch.setattr(tr, "transcribe_pdf_pages",
                        lambda images, ctx, progress=None, label="": (lambda p: (p["chapter_title"], p["blocks"]))(next(calls)))
    import web.app as appmod
    importlib.reload(appmod)
    from fastapi.testclient import TestClient
    return TestClient(appmod.app)


def _pdf_bytes():
    import pymupdf
    doc = pymupdf.open()
    doc.new_page()
    return doc.tobytes()


def test_web_flow(client):
    assert client.get("/", follow_redirects=False).status_code == 303
    assert client.post("/api/jobs", files=[("files", ("a.pdf", _pdf_bytes()))]).status_code == 401
    r = client.post("/login", data={"password": "bad"}, follow_redirects=False)
    assert "error" in r.headers["location"]
    client.post("/login", data={"password": "pw", "name": "Rahul"})
    assert "Namaste Rahul" in client.get("/").text
    r = client.post("/api/jobs", files=[("files", ("Part-1.pdf", _pdf_bytes())),
                                        ("files", ("Part-2.pdf", _pdf_bytes()))])
    jid = r.json()["id"]
    for _ in range(100):
        m = client.get(f"/api/jobs/{jid}").json()
        if m["status"] in ("done", "failed"):
            break
        time.sleep(0.1)
    assert m["status"] == "done", m
    assert m["output"] == "ठोस अवस्था (Solid State).docx"
    d = client.get(f"/api/jobs/{jid}/download")
    assert d.status_code == 200 and d.content[:2] == b"PK"
    p = client.get(f"/api/jobs/{jid}/download?kind=pptx")
    assert p.status_code == 200 and p.content[:2] == b"PK" and m["has_ppt"]
    assert client.get(f"/api/jobs/{jid}/download?kind=exe").status_code == 404
    assert client.post("/api/jobs", files=[("files", ("x.pdf", b"not a pdf"))]).status_code == 400


def test_ppt_deck(tmp_path):
    from pptx import Presentation
    from typist.pptkit import render_deck
    out = render_deck(PARTS, tmp_path / "d.pptx")
    prs = Presentation(out)
    assert len(prs.slides) >= 3
    text = " ".join(sh.text_frame.text for s in prs.slides for sh in s.shapes if sh.has_text_frame)
    assert "ठोस अवस्था (Solid State)" in text and "उदाहरण" in text
