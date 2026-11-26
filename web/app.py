"""Typist Automation website: upload panel PDFs -> download the finished Word notes.

Run:   APP_PASSWORD=... ANTHROPIC_API_KEY=... uvicorn web.app:app --host 0.0.0.0 --port 8000

Settings (environment variables):
  ANTHROPIC_API_KEY   Claude API key (required for conversion)
  APP_PASSWORD        shared password for the team (strongly recommended)
  SECRET_KEY          signs login cookies (defaults to a value derived from APP_PASSWORD)
  DATA_DIR            where uploads/outputs are kept (default ./data)
  JOB_WORKERS         conversions that run at the same time (default 2)
  KEEP_DAYS           delete jobs older than this (default 30)
"""
from __future__ import annotations

import hashlib
from html import escape
import hmac
import json
import os
import re
import shutil
import threading
import time
import traceback
import uuid
from concurrent.futures import ThreadPoolExecutor
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse

from typist.convert import build_docx, transcribe_all

DATA = Path(os.environ.get("DATA_DIR", "data")).resolve()
JOBS = DATA / "jobs"
JOBS.mkdir(parents=True, exist_ok=True)
PASSWORD = os.environ.get("APP_PASSWORD", "").strip()
SECRET = (os.environ.get("SECRET_KEY") or hashlib.sha256(("typist:" + PASSWORD).encode()).hexdigest()).encode()
WORKERS = int(os.environ.get("JOB_WORKERS", "2"))
KEEP_DAYS = float(os.environ.get("KEEP_DAYS", "30"))
MAX_FILES, MAX_MB = 20, 60


@asynccontextmanager
async def lifespan(_app):
    _startup()
    yield


app = FastAPI(title="Typist Automation", lifespan=lifespan)
pool = ThreadPoolExecutor(max_workers=WORKERS)
lock = threading.Lock()
HERE = Path(__file__).parent


# ------------------------------------------------------------------ auth --
def _token():
    return hmac.new(SECRET, b"logged-in", hashlib.sha256).hexdigest()


def authed(req: Request) -> bool:
    return not PASSWORD or hmac.compare_digest(req.cookies.get("ta_auth", ""), _token())


def need_auth(req: Request):
    if not authed(req):
        raise HTTPException(401, "Login karein")


@app.get("/login", response_class=HTMLResponse)
def login_page(error: str = ""):
    html = (HERE / "templates" / "login.html").read_text(encoding="utf-8")
    return html.replace("{{error}}", "Password galat hai." if error else "")


@app.post("/login")
def login(password: str = Form(""), name: str = Form("")):
    if PASSWORD and not hmac.compare_digest(password, PASSWORD):
        return RedirectResponse("/login?error=1", status_code=303)
    r = RedirectResponse("/", status_code=303)
    age = 60 * 60 * 24 * 60
    r.set_cookie("ta_auth", _token(), max_age=age, httponly=True, samesite="lax")
    r.set_cookie("ta_name", re.sub(r"[^\w .-]", "", name, flags=re.U)[:40] or "Typist", max_age=age,
                 samesite="lax")
    return r


@app.get("/logout")
def logout():
    r = RedirectResponse("/login", status_code=303)
    r.delete_cookie("ta_auth")
    return r


# ------------------------------------------------------------------ jobs --
def _meta_path(jid):
    return JOBS / jid / "job.json"


def load(jid):
    try:
        return json.loads(_meta_path(jid).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def save(meta):
    p = _meta_path(meta["id"])
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(meta, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(p)


def update(jid, **kw):
    with lock:
        m = load(jid)
        m.update(kw, updated=time.time())
        save(m)
        return m


def _safe_name(s, fallback="Notes"):
    s = re.sub(r"[\\/:*?\"<>|\x00-\x1f]+", " ", s or "").strip(" .")
    return (s or fallback)[:80]


def run_job(jid):
    m = load(jid)
    d = JOBS / jid
    try:
        update(jid, status="running", step="Shuru ho raha hai")
        pdfs = [(f, d / "input" / f"{i:02d}.pdf") for i, f in enumerate(m["files"])]
        parts = transcribe_all(pdfs, progress=lambda s: update(jid, step=s))
        (d / "blocks.json").write_text(json.dumps(parts, ensure_ascii=False, indent=1), encoding="utf-8")
        update(jid, step="Word file ban rahi hai")
        title = m.get("title") or next((p["chapter_title"] for p in parts if p["chapter_title"]), "")
        out_name = _safe_name(m.get("title") or title or Path(m["files"][0]).stem) + ".docx"
        build_docx(parts, d / "output.docx", m.get("title") or None)
        update(jid, status="done", step="Taiyar hai", output=out_name, chapter=title)
    except Exception as e:  # report every failure to the page
        traceback.print_exc()
        update(jid, status="failed", step="Fail", error=str(e)[:500])


def _cleanup():
    cutoff = time.time() - KEEP_DAYS * 86400
    for p in JOBS.iterdir():
        m = load(p.name)
        if m and m.get("created", 0) < cutoff:
            shutil.rmtree(p, ignore_errors=True)


def _startup():
    for p in JOBS.iterdir():
        m = load(p.name)
        if m and m.get("status") in ("queued", "running"):  # server restarted mid-job
            update(p.name, status="failed", step="Fail", error="Server restart ho gaya. Dobara upload karein.")
    _cleanup()


@app.post("/api/jobs")
async def create_job(request: Request, files: list[UploadFile] = File(...), title: str = Form("")):
    need_auth(request)
    if not files or len(files) > MAX_FILES:
        raise HTTPException(400, f"1 se {MAX_FILES} PDF tak upload karein.")
    jid = time.strftime("%Y%m%d-%H%M%S-") + uuid.uuid4().hex[:6]
    inp = JOBS / jid / "input"
    inp.mkdir(parents=True)
    names = []
    try:
        for i, f in enumerate(files):  # order = order chosen on the page
            data = await f.read()
            if not data.startswith(b"%PDF"):
                raise HTTPException(400, f"{f.filename} PDF nahi hai.")
            if len(data) > MAX_MB * 1024 * 1024:
                raise HTTPException(400, f"{f.filename} {MAX_MB} MB se bada hai.")
            (inp / f"{i:02d}.pdf").write_bytes(data)
            names.append(Path(f.filename or f"part{i + 1}.pdf").name)
    except HTTPException:
        shutil.rmtree(JOBS / jid, ignore_errors=True)
        raise
    now = time.time()
    meta = {"id": jid, "files": names, "title": title.strip()[:120], "status": "queued",
            "step": "Line me hai", "by": request.cookies.get("ta_name", "Typist"),
            "created": now, "updated": now}
    save(meta)
    pool.submit(run_job, jid)
    return {"id": jid}


@app.get("/api/jobs")
def list_jobs(request: Request):
    need_auth(request)
    metas = [m for p in JOBS.iterdir() if (m := load(p.name))]
    metas.sort(key=lambda m: m["created"], reverse=True)
    return metas[:60]


@app.get("/api/jobs/{jid}")
def get_job(jid: str, request: Request):
    need_auth(request)
    m = load(jid) if re.fullmatch(r"[\w-]+", jid) else None
    if not m:
        raise HTTPException(404)
    return m


@app.get("/api/jobs/{jid}/download")
def download(jid: str, request: Request):
    need_auth(request)
    m = get_job(jid, request)
    out = JOBS / jid / "output.docx"
    if m.get("status") != "done" or not out.exists():
        raise HTTPException(404, "File abhi taiyar nahi hai")
    return FileResponse(out, filename=m.get("output") or "Notes.docx",
                        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.get("/healthz")
def health():
    return {"ok": True, "api_key": bool(os.environ.get("ANTHROPIC_API_KEY", "").strip()), "password": bool(PASSWORD)}


@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    if not authed(request):
        return RedirectResponse("/login", status_code=303)
    html = (HERE / "templates" / "index.html").read_text(encoding="utf-8")
    return html.replace("{{name}}", escape(request.cookies.get("ta_name", "Typist"))) \
               .replace("{{nokey}}", "" if os.environ.get("ANTHROPIC_API_KEY", "").strip() else
                        '<div class="warn">Server par ANTHROPIC_API_KEY set nahi hai, conversion nahi chalega.</div>')


@app.exception_handler(HTTPException)
def http_err(request: Request, exc: HTTPException):
    if exc.status_code == 401 and not request.url.path.startswith("/api/"):
        return RedirectResponse("/login", status_code=303)
    return JSONResponse({"detail": exc.detail}, status_code=exc.status_code)
