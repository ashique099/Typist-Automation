"""Typist Automation website.

Typists log in with their own account and pick one of three tools:
  1. Panel PDF to Notes DOCX   (Claude reads the handwritten slides)
  2. DOCX to PPT               (no API cost: built from the notes text)
  3. PPT / DOCX to MCQ Test    (Claude writes the MCQs; Word test + Objective PPT)  Every job is kept in the work
history with who uploaded it, its files and its real API cost.  The admin adds typists,
picks the model, uploads PPT designs and sets the monthly budget.

Run:   ANTHROPIC_API_KEY=... uvicorn web.app:app --host 0.0.0.0 --port 8000
The first visit asks to create the admin account.

Environment:
  ANTHROPIC_API_KEY   Claude API key (required for conversion; never stored in the repo)
  DATA_DIR            database, uploads and outputs (default ./data)
  JOB_WORKERS         conversions that run at the same time (default 2)
  KEEP_DAYS           delete job files older than this; history rows stay (default 90)
  SETUP_KEY           if set, first-time setup asks for it (use on a public host such as Render)
"""
from __future__ import annotations

import json
import os
import re
import secrets
import shutil
import time
import traceback
import uuid
from concurrent.futures import ThreadPoolExecutor
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse, Response

from typist import mcq as mcqmod
from typist import usage
from typist.convert import build_docx, transcribe_all
from typist.reader import read_parts
from typist.pptkit import render_deck, render_mcq_deck, render_template_deck

from . import db

DATA = Path(os.environ.get("DATA_DIR", "data")).resolve()
JOBS = DATA / "jobs"
TEMPLATES = Path(os.environ.get("TEMPLATE_DIR", DATA / "templates")).resolve()
WORKERS = int(os.environ.get("JOB_WORKERS", "2"))
KEEP_DAYS = float(os.environ.get("KEEP_DAYS", "90"))
MAX_FILES, MAX_MB = 20, 60
COOKIE = "ta_session"        # typist website login
ADMIN_COOKIE = "ta_admin"    # admin panel login (separate page, separate session)
HERE = Path(__file__).parent

DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
# output kind -> (file in job dir, suffix for the download name, media type, label)
KINDS = {"notes": ("notes.docx", "", DOCX, "Notes DOCX"),
         "ppt": ("notes.pptx", "", PPTX, "PPT"),
         "test": ("test.docx", " - MCQ Test", DOCX, "MCQ Test DOCX"),
         "objective": ("objective.pptx", " - Objective", PPTX, "Objective PPT")}


KEYFILE = DATA / "api_key.txt"   # saved from Admin > Settings; lives only on the server, never in git
BAT_KEY = os.environ.get("ANTHROPIC_API_KEY", "").strip()  # key that came from settings.bat (if any)


def _load_key_file():
    try:
        k = KEYFILE.read_text(encoding="utf-8").strip()
    except OSError:
        return
    if k:
        os.environ["ANTHROPIC_API_KEY"] = k


def _key_info():
    k = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    src = "admin panel" if KEYFILE.exists() and k else "settings.bat" if k else ""
    return {"set": bool(k), "source": src, "masked": (k[:7] + "…" + k[-4:]) if len(k) > 14 else ("set" if k else "")}


def _startup():
    JOBS.mkdir(parents=True, exist_ok=True)
    _load_key_file()
    TEMPLATES.mkdir(parents=True, exist_ok=True)
    db.init(DATA / "typist.db")
    for j in db.q("select id from jobs where status in ('queued','running')"):  # restarted mid-job
        db.update_job(j["id"], status="failed", step="Failed", error="The server restarted during this task. Please upload again.")
    cutoff = time.time() - KEEP_DAYS * 86400
    for j in db.q("select id from jobs where created < ?", (cutoff,)):
        shutil.rmtree(JOBS / j["id"], ignore_errors=True)
        db.update_job(j["id"], outputs={})


@asynccontextmanager
async def lifespan(_app):
    _startup()
    yield


app = FastAPI(title="Disha Online Solution", lifespan=lifespan)
pool = ThreadPoolExecutor(max_workers=WORKERS)


# ------------------------------------------------------------------ auth --
def _panel(req: Request) -> str:
    """Which login a request uses: the admin panel sends X-Panel: admin (or ?panel=admin on links)."""
    return "admin" if (req.headers.get("x-panel") or req.query_params.get("panel")) == "admin" else "site"


def admin_user(req: Request):
    u = db.session_user(req.cookies.get(ADMIN_COOKIE))
    return u if u and u["role"] == "admin" else None


def current_user(req: Request):
    """The typist-website user.  An admin account signed in on the typist login works there
    as a normal typist (own jobs only); admin powers come only from the admin login."""
    u = db.session_user(req.cookies.get(COOKIE))
    if u and u["role"] == "admin":
        u = dict(u, role="typist", is_admin=True)
    return u


def need_user(req: Request):
    u = admin_user(req) if _panel(req) == "admin" else current_user(req)
    if not u:
        raise HTTPException(401, "Please log in")
    return u


def need_admin(req: Request):
    u = admin_user(req)
    if not u:
        raise HTTPException(403, "Admins only")
    return u


def _page(name, **subs):
    html = (HERE / "templates" / name).read_text(encoding="utf-8")
    for k, v in subs.items():
        html = html.replace("{{" + k + "}}", v)
    return HTMLResponse(html)


def _login_response(tok, request: Request | None = None, admin=False):
    r = RedirectResponse("/admin" if admin else "/", status_code=303)
    secure = bool(request and request.url.scheme == "https")  # behind nginx + HTTPS on the VPS
    r.set_cookie(ADMIN_COOKIE if admin else COOKIE, tok, max_age=(7 if admin else 60) * 86400,
                 httponly=True, samesite="strict" if admin else "lax", secure=secure)
    return r


# simple brute-force guard for the public website: 8 wrong passwords per IP+user in 15 minutes
_FAILS: dict[str, list[float]] = {}


def _client_ip(request: Request) -> str:
    return (request.client.host if request.client else "") or "?"


def _too_many(key: str) -> bool:
    now = time.time()
    _FAILS[key] = [t for t in _FAILS.get(key, []) if now - t < 900]
    return len(_FAILS[key]) >= 8


@app.get("/login")
def login_page(request: Request, error: str = ""):
    if db.user_count() == 0:
        return RedirectResponse("/setup", status_code=303)
    if current_user(request):
        return RedirectResponse("/", status_code=303)
    msg = {"1": "Wrong username or password.", "2": "Too many wrong tries. Wait 15 minutes and try again."}
    return _page("login.html", error=msg.get(error, "Wrong username or password.") if error else "")


@app.post("/login")
def login(request: Request, username: str = Form(""), password: str = Form("")):
    key = _client_ip(request) + "|" + username.strip().lower()
    if _too_many(key):
        return RedirectResponse("/login?error=2", status_code=303)
    tok = db.login(username, password)
    if not tok:
        _FAILS.setdefault(key, []).append(time.time())
        return RedirectResponse("/login?error=1", status_code=303)
    _FAILS.pop(key, None)
    return _login_response(tok, request)


SETUP_KEY = os.environ.get("SETUP_KEY", "").strip()  # on a public host: only someone with this key can create the first admin
_KEY_FIELD = ('<label for="k">Setup Key (From the Hosting Dashboard)</label>'
              '<input id="k" name="setup_key" type="password" autocomplete="off" required>')


@app.get("/setup")
def setup_page(error: str = ""):
    if db.user_count():
        return RedirectResponse("/login", status_code=303)
    msg = {"2": "Wrong setup key. Copy SETUP_KEY from the hosting dashboard."}.get(
        error, "Use a username of 3-30 letters/numbers and a password of at least 6 characters, typed the same twice.")
    return _page("setup.html", keyfield=_KEY_FIELD if SETUP_KEY else "", error=msg if error else "")


@app.post("/setup")
def setup(request: Request, name: str = Form(""), username: str = Form(""), password: str = Form(""), password2: str = Form(""),
          setup_key: str = Form("")):
    if db.user_count():
        return RedirectResponse("/login", status_code=303)
    if SETUP_KEY and not secrets.compare_digest(setup_key.strip(), SETUP_KEY):
        return RedirectResponse("/setup?error=2", status_code=303)
    if len(password) < 6 or password != password2 or not re.fullmatch(r"[a-zA-Z0-9_.-]{3,30}", username.strip()):
        return RedirectResponse("/setup?error=1", status_code=303)
    db.create_user(username, name, password, role="admin")
    return _login_response(db.login(username, password), request, admin=True)


@app.get("/logout")
def logout(request: Request):
    db.logout(request.cookies.get(COOKIE))
    r = RedirectResponse("/login", status_code=303)
    r.delete_cookie(COOKIE)
    return r


@app.get("/admin/login")
def admin_login_page(request: Request, error: str = ""):
    if db.user_count() == 0:
        return RedirectResponse("/setup", status_code=303)
    if admin_user(request):
        return RedirectResponse("/admin", status_code=303)
    msg = {"1": "Wrong admin username or password.", "2": "Too many wrong tries. Wait 15 minutes and try again."}
    return _page("admin-login.html", error=msg.get(error, msg["1"]) if error else "")


@app.post("/admin/login")
def admin_login(request: Request, username: str = Form(""), password: str = Form("")):
    key = "admin|" + _client_ip(request) + "|" + username.strip().lower()
    if _too_many(key):
        return RedirectResponse("/admin/login?error=2", status_code=303)
    u = db.q("select role, active from users where username=?", (username.strip().lower(),), one=True)
    tok = db.login(username, password) if u and u["role"] == "admin" and u["active"] else None
    if not tok:  # typist accounts never get in here, and we don't say which part was wrong
        _FAILS.setdefault(key, []).append(time.time())
        return RedirectResponse("/admin/login?error=1", status_code=303)
    _FAILS.pop(key, None)
    return _login_response(tok, request, admin=True)


@app.get("/admin/logout")
def admin_logout(request: Request):
    db.logout(request.cookies.get(ADMIN_COOKIE))
    r = RedirectResponse("/admin/login", status_code=303)
    r.delete_cookie(ADMIN_COOKIE)
    return r


@app.get("/api/me")
def me(request: Request):
    u = need_user(request)
    s = db.all_settings()
    spend = db.month_spend_usd()
    rate = float(s["usd_to_inr"])
    return {"user": u, "api_key": bool(os.environ.get("ANTHROPIC_API_KEY", "").strip()),
            "templates": ppt_templates(), "default_mcq_count": int(s["default_mcq_count"]),
            "budget": {"limit_inr": float(s["monthly_budget_inr"]), "spent_inr": round(spend * rate),
                       "spent_usd": round(spend, 2)}}


@app.post("/api/password")
def change_password(request: Request, old: str = Form(""), new: str = Form("")):
    u = need_user(request)
    row = db.q("select pw from users where id=?", (u["id"],), one=True)
    if not db.check_pw(old, row["pw"]):
        raise HTTPException(400, "Your current password is wrong.")
    if len(new) < 6:
        raise HTTPException(400, "The new password must be at least 6 characters.")
    db.x("update users set pw=? where id=?", (db.hash_pw(new), u["id"]))
    return {"ok": True}


# ------------------------------------------------------------- templates --
def ppt_templates():
    return sorted(p.stem for p in TEMPLATES.glob("*.pptx"))


def _template(name):
    tpl = TEMPLATES / f"{name}.pptx"
    if not name or not tpl.exists():
        return None, None
    logo = tpl.with_suffix(".png")
    return str(tpl), (str(logo) if logo.exists() else None)


# ------------------------------------------------------------------ jobs --
def _budget_left_inr():
    s = db.all_settings()
    return float(s["monthly_budget_inr"]) - db.month_spend_usd() * float(s["usd_to_inr"])


def _safe_name(s, fallback="Notes"):
    s = re.sub(r"[\\/:*?\"<>|\x00-\x1f]+", " ", s or "").strip(" .")
    return (s or fallback)[:80]


TASKS = {  # dashboard card -> what it accepts
    "notes": {"label": "Panel PDF to Notes DOCX", "ext": (".pdf",), "ai": True},
    "ppt": {"label": "DOCX to PPT", "ext": (".docx",), "ai": False},
    "test": {"label": "PPT / DOCX to MCQ Test", "ext": (".docx", ".pptx"), "ai": True},
}


def _make_ppt(parts, out, tpl, logo, title):
    if tpl:
        render_template_deck(parts, out, tpl, logo, title)
    else:
        render_deck(parts, out, title)


def run_job(jid, keep_cost=False):
    m = db.get_job(jid)
    opts = m["options"]
    task = opts.get("task", "notes")
    d = JOBS / jid
    s = db.all_settings()
    meter = usage.start(model=s["mcq_model"] if task == "test" else s["model"], effort=s["effort"])
    if keep_cost:  # a retry adds to what the failed attempt already spent
        meter.calls, meter.input, meter.output = m["calls"] or 0, m["input_tokens"] or 0, m["output_tokens"] or 0
        meter.cost = m["cost_usd"] or 0
    outputs = {}

    def save_usage(**kw):
        db.update_job(jid, outputs=outputs, calls=meter.calls, input_tokens=meter.input,
                      output_tokens=meter.output, cost_usd=round(meter.cost, 5), **kw)

    try:
        db.update_job(jid, status="running", step="Starting", started=time.time(),
                      model=meter.model if TASKS[task]["ai"] else "")
        files = [(f, d / "input" / f"{i:02d}{Path(f).suffix.lower()}") for i, f in enumerate(m["files"])]
        if task == "notes":
            stats = {}
            parts = transcribe_all(files, progress=lambda t: save_usage(step=t), stats=stats)
            save_usage(pages=stats.get("pages", 0), pages_sent=stats.get("sent", 0))
        else:
            save_usage(step="Reading files")
            parts = read_parts(files)
            if not any(p["blocks"] for p in parts):
                raise ValueError("No text found in the uploaded file(s).")
        (d / "blocks.json").write_text(json.dumps(parts, ensure_ascii=False, indent=1), encoding="utf-8")
        title = m["title"] or next((p["chapter_title"] for p in parts if p["chapter_title"]), "")
        base = _safe_name(title or Path(m["files"][0]).stem)
        save_usage(chapter=title)
        tpl, logo = _template(opts.get("ppt_template"))
        if task == "notes":
            save_usage(step="Building the Word notes")
            build_docx(parts, d / "notes.docx", m["title"] or None)
            outputs["notes"] = base
        elif task == "ppt":
            save_usage(step="Building the PPT")
            _make_ppt(parts, d / "notes.pptx", tpl, logo, m["title"] or None)
            outputs["ppt"] = base
        else:
            n = int(opts.get("mcq_count") or 50)
            save_usage(step=f"Writing {n} MCQs")
            qs = mcqmod.generate_mcqs(parts, n)
            (d / "mcq.json").write_text(json.dumps(qs, ensure_ascii=False, indent=1), encoding="utf-8")
            mcqmod.write_mcq_docx(qs, d / "test.docx")
            outputs["test"] = base
            save_usage(mcqs=len(qs))
            try:
                render_mcq_deck(qs, d / "objective.pptx", tpl, logo, title or None)
                outputs["objective"] = base
            except Exception:
                traceback.print_exc()
        save_usage(status="done", step="Ready", finished=time.time())
    except Exception as e:
        traceback.print_exc()
        save_usage(status="failed", step="Failed", error=str(e)[:500], finished=time.time())


@app.post("/api/jobs")
async def create_job(request: Request, files: list[UploadFile] = File(...), task: str = Form("notes"),
                     title: str = Form(""), ppt_template: str = Form(""), mcq_count: int = Form(50)):
    u = need_user(request)
    if task not in TASKS:
        raise HTTPException(400, "Unknown task.")
    spec = TASKS[task]
    if spec["ai"]:
        if not os.environ.get("ANTHROPIC_API_KEY", "").strip():
            raise HTTPException(400, f"Notes and MCQ generation are temporarily unavailable. Please contact {CONTACT}.")
        if _budget_left_inr() <= 0:
            raise HTTPException(400, f"This month's limit for Notes and MCQ is reached. Please contact {CONTACT}.")
    if not files or len(files) > MAX_FILES:
        raise HTTPException(400, f"Upload 1 to {MAX_FILES} files.")
    jid = time.strftime("%Y%m%d-%H%M%S-") + uuid.uuid4().hex[:6]
    inp = JOBS / jid / "input"
    inp.mkdir(parents=True)
    names = []
    try:
        for i, f in enumerate(files):  # order = order chosen on the page
            name = Path(f.filename or f"part{i + 1}").name
            ext = Path(name).suffix.lower()
            data = await f.read()
            magic = b"%PDF" if ext == ".pdf" else b"PK"
            if ext not in spec["ext"] or not data.startswith(magic):
                raise HTTPException(400, f"{name}: please upload {' or '.join(e[1:].upper() for e in spec['ext'])} files only.")
            if len(data) > MAX_MB * 1024 * 1024:
                raise HTTPException(400, f"{name} is larger than {MAX_MB} MB.")
            (inp / f"{i:02d}{ext}").write_bytes(data)
            names.append(name)
    except HTTPException:
        shutil.rmtree(JOBS / jid, ignore_errors=True)
        raise
    opts = {"task": task, "ppt_template": ppt_template if ppt_template in ppt_templates() else ""}
    if task == "test":
        opts["mcq_count"] = max(5, min(int(mcq_count or 50), 150))
    db.add_job(jid, u["id"], title.strip()[:120], names, opts)
    pool.submit(run_job, jid)
    return {"id": jid}


CONTACT = "Ashique Sir"   # who typists should contact when something needs the admin
_TECH = re.compile(r"claude|anthropic|api[ _-]?key|ANTHROPIC|model|token|rate.?limit|overloaded|credit|billing|"
                   r"authentication|permission|status code|\b40[0-9]\b|\b5[0-9]{2}\b|budget", re.I)


def typist_error(msg: str) -> str:
    """Typists never see technical wording (API, Claude, model, budget...)."""
    if not msg or not _TECH.search(msg):
        return msg
    if re.search(r"budget|credit|billing", msg, re.I):
        return f"This month's limit for Notes and MCQ is reached. Please contact {CONTACT}."
    if re.search(r"api[ _-]?key|ANTHROPIC_API_KEY|authentication", msg, re.I):
        return f"Notes and MCQ generation are temporarily unavailable. Please contact {CONTACT}."
    return f"The server could not finish this task. Please try again, or contact {CONTACT} if it keeps happening."


def _public(j, rate, admin=False):
    j = dict(j)
    if not admin and j.get("error"):
        j["error"] = typist_error(j["error"])
    j["cost_inr"] = round((j.get("cost_usd") or 0) * rate, 1)
    j["task"] = (j.get("options") or {}).get("task", "notes")
    j["task_label"] = TASKS.get(j["task"], TASKS["notes"])["label"]
    j["downloads"] = [{"kind": k, "label": KINDS[k][3]} for k in KINDS if k in (j.get("outputs") or {})]
    return j


def _job_filters(u, scope, status, q, days, task, user, month):
    """Shared by the history list and the CSV export. month = 'YYYY-MM'."""
    uid = u["id"]
    if u["role"] == "admin" and scope == "all":
        uid = user or None
    since = until = None
    if month and re.fullmatch(r"\d{4}-\d{2}", month):
        y, mo = map(int, month.split("-"))
        since = time.mktime((y, mo, 1, 0, 0, 0, 0, 0, -1))
        until = time.mktime((y + mo // 12, mo % 12 + 1, 1, 0, 0, 0, 0, 0, -1))
    elif days:
        since = time.time() - days * 86400
    return dict(user_id=uid, status=status or None, since=since, until=until, search=q.strip() or None,
                task=task if task in TASKS else None)


@app.get("/api/jobs")
def list_jobs(request: Request, scope: str = "mine", status: str = "", q: str = "", days: int = 0,
              page: int = 1, task: str = "", user: int = 0, month: str = "", per: int = 25):
    u = need_user(request)
    per = max(5, min(per, 100))
    f = _job_filters(u, scope, status, q, days, task, user, month)
    rows, total = db.list_jobs(limit=per, offset=(max(page, 1) - 1) * per, **f)
    rate = float(db.setting("usd_to_inr"))
    adm = _panel(request) == "admin"
    return {"jobs": [_public(r, rate, adm) for r in rows], "total": total, "page": page, "per": per}


def _own_job(request, jid):
    u = need_user(request)
    j = db.get_job(jid) if re.fullmatch(r"[\w-]+", jid) else None
    if not j or (j["user_id"] != u["id"] and u["role"] != "admin"):
        raise HTTPException(404)
    return j


@app.get("/api/jobs/{jid}")
def get_job(jid: str, request: Request):
    return _public(_own_job(request, jid), float(db.setting("usd_to_inr")), _panel(request) == "admin")


@app.get("/api/jobs/{jid}/download")
def download(jid: str, request: Request, kind: str = "notes"):
    j = _own_job(request, jid)
    if kind not in KINDS or kind not in (j["outputs"] or {}):
        raise HTTPException(404, "This file is not ready")
    fname, suffix, media, _ = KINDS[kind]
    out = JOBS / jid / fname
    if not out.exists():
        raise HTTPException(404, "This file was deleted because it is older than the keep period")
    return FileResponse(out, filename=j["outputs"][kind] + suffix + Path(fname).suffix, media_type=media)


# ----------------------------------------------------------------- admin --
@app.get("/api/admin/overview")
def admin_overview(request: Request):
    need_admin(request)
    s = db.all_settings()
    rate = float(s["usd_to_inr"])
    o = db.overview(rate)
    o["budget_inr"] = float(s["monthly_budget_inr"])
    o["api_key"] = bool(os.environ.get("ANTHROPIC_API_KEY", "").strip())
    o["models"] = {"notes": usage.MODELS.get(s["model"], s["model"]),
                   "mcq": usage.MODELS.get(s["mcq_model"], s["mcq_model"]), "effort": s["effort"]}
    o["labels"] = {k: v["label"] for k, v in TASKS.items()}
    du = shutil.disk_usage(DATA)
    used = sum(f.stat().st_size for f in JOBS.rglob("*") if f.is_file())
    from typist.pages import PDF_ENGINE
    o["pdf_engine"] = PDF_ENGINE
    o["storage"] = {"jobs_mb": round(used / 2 ** 20, 1), "free_gb": round(du.free / 2 ** 30, 1),
                    "keep_days": KEEP_DAYS}
    return o


@app.get("/api/admin/jobs.csv")
def admin_jobs_csv(request: Request, status: str = "", q: str = "", days: int = 0, task: str = "",
                   user: int = 0, month: str = ""):
    import csv
    import io
    u = need_admin(request)
    f = _job_filters(u, "all", status, q, days, task, user, month)
    rows, _ = db.list_jobs(limit=100000, offset=0, **f)
    rate = float(db.setting("usd_to_inr"))
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["Date", "Time", "Typist", "Tool", "Chapter", "Files", "Status", "Pages", "Pages sent",
                "MCQs", "Minutes", "Cost (INR)", "Cost (USD)", "Model", "Error"])
    for r in rows:
        t = time.localtime(r["created"])
        mins = round((r["finished"] - r["started"]) / 60, 1) if r.get("finished") and r.get("started") else ""
        w.writerow([time.strftime("%Y-%m-%d", t), time.strftime("%H:%M", t), r.get("by_name") or "",
                    TASKS.get((r["options"] or {}).get("task", "notes"), TASKS["notes"])["label"],
                    r.get("title") or r.get("chapter") or "", " | ".join(r["files"]), r["status"],
                    r.get("pages") or 0, r.get("pages_sent") or 0, r.get("mcqs") or 0, mins,
                    round((r.get("cost_usd") or 0) * rate, 2), round(r.get("cost_usd") or 0, 4),
                    r.get("model") or "", r.get("error") or ""])
    name = f"work-history-{month or time.strftime('%Y-%m-%d')}.csv"
    return Response("\ufeff" + buf.getvalue(), media_type="text/csv; charset=utf-8",
                    headers={"Content-Disposition": f'attachment; filename="{name}"'})


@app.post("/api/admin/jobs/{jid}/retry")
def admin_retry(jid: str, request: Request):
    need_admin(request)
    j = _own_job(request, jid)
    if j["status"] != "failed":
        raise HTTPException(400, "Only failed jobs can be run again.")
    if not (JOBS / jid / "input").exists():
        raise HTTPException(400, "The uploaded files of this job were already deleted.")
    if TASKS.get(j["options"].get("task", "notes"), TASKS["notes"])["ai"]:
        if not os.environ.get("ANTHROPIC_API_KEY", "").strip():
            raise HTTPException(400, "The Claude API key is not set on the server.")
        if _budget_left_inr() <= 0:
            raise HTTPException(400, "This month's budget is used up.")
    # the earlier attempt's cost stays counted: it was really spent
    db.update_job(jid, status="queued", step="Queued", error=None, outputs={}, finished=None)
    pool.submit(run_job, jid, keep_cost=True)
    return {"ok": True}


@app.post("/api/admin/jobs/{jid}/delete")
def admin_delete_job(jid: str, request: Request):
    need_admin(request)
    j = _own_job(request, jid)
    if j["status"] in ("queued", "running"):
        raise HTTPException(400, "This task is still running.")
    shutil.rmtree(JOBS / jid, ignore_errors=True)
    db.update_job(jid, outputs={}, error="Files deleted by admin", step=j["step"])
    return {"ok": True}


@app.get("/api/admin/users")
def users(request: Request):
    need_admin(request)
    rate = float(db.setting("usd_to_inr"))
    use = {r["id"]: r for r in db.usage_by_user(db.month_start())}
    out = []
    for u in db.q("select id,username,name,role,active,created, "
                  "(select max(last_seen) from sessions s where s.user_id=users.id) last_seen, "
                  "(select count(*) from jobs j where j.user_id=users.id) all_jobs "
                  "from users order by role, name"):
        r = use.get(u["id"], {})
        out.append({**u, "jobs": r.get("jobs", 0), "pages": r.get("pages", 0), "mcqs": r.get("mcqs", 0),
                    "cost_inr": round(r.get("cost_usd", 0) * rate)})
    return out


@app.post("/api/admin/users")
def add_user(request: Request, username: str = Form(""), name: str = Form(""), password: str = Form(""),
             role: str = Form("typist")):
    need_admin(request)
    if not re.fullmatch(r"[a-zA-Z0-9_.-]{3,30}", username.strip()):
        raise HTTPException(400, "Username must be 3-30 characters (a-z, 0-9, _ . -).")
    if len(password) < 6:
        raise HTTPException(400, "Password must be at least 6 characters.")
    if db.q("select 1 from users where username=?", (username.strip().lower(),), one=True):
        raise HTTPException(400, "That username is already taken.")
    u = db.create_user(username, name, password, "admin" if role == "admin" else "typist")
    return {"id": u["id"]}


@app.post("/api/admin/users/{uid}")
def edit_user(uid: int, request: Request, active: str = Form(""), password: str = Form(""), role: str = Form(""),
              name: str = Form("")):
    me_ = need_admin(request)
    u = db.get_user(uid)
    if not u:
        raise HTTPException(404)
    if active in ("0", "1"):
        if uid == me_["id"] and active == "0":
            raise HTTPException(400, "You cannot disable your own account.")
        if active == "0" and u["role"] == "admin" and db.q(
                "select count(*) n from users where role='admin' and active=1", one=True)["n"] <= 1:
            raise HTTPException(400, "At least one admin must stay active.")
        db.x("update users set active=? where id=?", (int(active), uid))
        if active == "0":
            db.x("delete from sessions where user_id=?", (uid,))
    if password:
        if len(password) < 6:
            raise HTTPException(400, "Password must be at least 6 characters.")
        db.x("update users set pw=? where id=?", (db.hash_pw(password), uid))
    if name.strip():
        db.x("update users set name=? where id=?", (name.strip()[:60], uid))
    if role in ("admin", "typist") and uid != me_["id"]:
        db.x("update users set role=? where id=?", (role, uid))
    return {"ok": True}


@app.get("/api/admin/settings")
def get_settings(request: Request):
    need_admin(request)
    s = db.all_settings()
    since = db.month_start()
    rate = float(s["usd_to_inr"])
    month = db.q("select count(*) jobs, coalesce(sum(pages),0) pages, coalesce(sum(pages_sent),0) sent, "
                 "coalesce(sum(mcqs),0) mcqs, coalesce(sum(cost_usd),0) usd from jobs where created>=?",
                 (since,), one=True)
    month["inr"] = round(month["usd"] * rate)
    notes = db.q("select coalesce(sum(cost_usd),0) usd from jobs where created>=? and "
                 "coalesce(json_extract(options,'$.task'),'notes')='notes'", (since,), one=True)
    month["notes_inr"] = notes["usd"] * rate
    return {"settings": s, "models": usage.MODELS, "month": month, "templates": ppt_templates(), "api": _key_info()}


# --------------------------------------------------------------- notices --
@app.get("/api/notices")
def notices(request: Request):
    """Notice Board on the typist dashboard (only the notices the admin has switched on)."""
    need_user(request)
    return [{k: n[k] for k in ("id", "title", "body", "color", "pinned", "updated")} for n in db.list_notices(True)]


@app.get("/api/admin/notices")
def admin_notices(request: Request):
    need_admin(request)
    return db.list_notices()


@app.post("/api/admin/notices")
def admin_save_notice(request: Request, id: int = Form(0), title: str = Form(""), body: str = Form(""),
                      color: str = Form("blue"), visible: str = Form("1"), pinned: str = Form("0")):
    u = need_admin(request)
    title, body = title.strip()[:120], body.strip()[:4000]
    if not title:
        raise HTTPException(400, "Please write a title for the notice.")
    if id and not db.q("select 1 from notices where id=?", (id,), one=True):
        raise HTTPException(404, "Notice not found.")
    nid = db.save_notice(id, title, body, color, visible == "1", pinned == "1", u["name"])
    return {"ok": True, "id": nid}


@app.post("/api/admin/notices/{nid}/toggle")
def admin_toggle_notice(nid: int, request: Request):
    need_admin(request)
    db.x("update notices set visible=1-visible, updated=? where id=?", (time.time(), nid))
    return {"ok": True}


@app.post("/api/admin/notices/{nid}/delete")
def admin_delete_notice(nid: int, request: Request):
    need_admin(request)
    db.x("delete from notices where id=?", (nid,))
    return {"ok": True}


@app.post("/api/admin/apikey")
def save_api_key(request: Request, key: str = Form("")):
    need_admin(request)
    key = key.strip()
    if not key.startswith("sk-ant-") or len(key) < 30 or any(c.isspace() for c in key):
        raise HTTPException(400, "That does not look like a Claude API key (it starts with sk-ant-).")
    DATA.mkdir(parents=True, exist_ok=True)
    KEYFILE.write_text(key, encoding="utf-8")
    try:
        os.chmod(KEYFILE, 0o600)
    except OSError:
        pass
    os.environ["ANTHROPIC_API_KEY"] = key
    return {"ok": True, "api": _key_info()}


@app.post("/api/admin/apikey/delete")
def delete_api_key(request: Request):
    need_admin(request)
    if KEYFILE.exists():
        KEYFILE.unlink()
        if BAT_KEY:
            os.environ["ANTHROPIC_API_KEY"] = BAT_KEY
        else:
            os.environ.pop("ANTHROPIC_API_KEY", None)
    return {"ok": True, "api": _key_info(),
            "note": "Removed. The key from settings.bat is in use again." if BAT_KEY else ""}


@app.post("/api/admin/apikey/test")
def test_api_key(request: Request):
    """Checks the key and each chosen model with a free models lookup (no tokens are spent)."""
    need_admin(request)
    k = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not k:
        raise HTTPException(400, "No API key is set yet.")
    import anthropic
    s = db.all_settings()
    try:
        c = anthropic.Anthropic(api_key=k, max_retries=1, timeout=20)
        for m in {s["model"], s["mcq_model"]}:
            c.models.retrieve(m)
    except anthropic.AuthenticationError:
        raise HTTPException(400, "The key was rejected by Anthropic (wrong or revoked key).")
    except anthropic.PermissionDeniedError:
        raise HTTPException(400, "The key works but has no access to the chosen model.")
    except anthropic.NotFoundError:
        raise HTTPException(400, "The key works but a chosen model was not found for this account.")
    except anthropic.APIConnectionError:
        raise HTTPException(400, "Could not reach api.anthropic.com. Check the PC's internet connection.")
    except anthropic.APIError as e:
        raise HTTPException(400, f"Anthropic error: {str(e)[:200]}")
    return {"ok": True, "models": [usage.MODELS.get(m, m) for m in {s["model"], s["mcq_model"]}]}


@app.post("/api/admin/settings")
def save_settings(request: Request, monthly_budget_inr: str = Form(""), usd_to_inr: str = Form(""),
                  model: str = Form(""), mcq_model: str = Form(""), effort: str = Form(""),
                  default_mcq_count: str = Form("")):
    need_admin(request)
    for k, v in (("monthly_budget_inr", monthly_budget_inr), ("usd_to_inr", usd_to_inr),
                 ("default_mcq_count", default_mcq_count)):
        if v:
            try:
                if float(v) < 0:
                    raise ValueError
            except ValueError:
                raise HTTPException(400, f"{k} is not a valid number.")
            db.set_setting(k, v)
    for k, v in (("model", model), ("mcq_model", mcq_model)):
        if v in usage.MODELS:
            db.set_setting(k, v)
    if effort in ("low", "medium", "high", "xhigh"):
        db.set_setting("effort", effort)
    return {"ok": True}


@app.post("/api/admin/templates")
async def upload_template(request: Request, name: str = Form(""), pptx: UploadFile = File(...),
                          logo: UploadFile | None = File(None)):
    need_admin(request)
    name = re.sub(r"[^\w .()-]+", "", name, flags=re.U).strip()[:60] or Path(pptx.filename or "Design").stem[:60]
    data = await pptx.read()
    if not data.startswith(b"PK"):
        raise HTTPException(400, "That is not a .pptx file.")
    tmp = TEMPLATES / f".{uuid.uuid4().hex}.pptx"
    tmp.write_bytes(data)
    try:  # keep only the master/layouts so the template stays small
        from pptx import Presentation
        p = Presentation(tmp)
        lst = p.slides._sldIdLst
        for el in list(lst):
            p.part.drop_rel(el.rId)
            lst.remove(el)
        p.save(TEMPLATES / f"{name}.pptx")
    except Exception:
        raise HTTPException(400, "Could not open that PPT file.")
    finally:
        tmp.unlink(missing_ok=True)
    if logo is not None and logo.filename:
        img = await logo.read()
        if img.startswith(b"\x89PNG") or img[:2] == b"\xff\xd8":
            from PIL import Image
            import io
            Image.open(io.BytesIO(img)).save(TEMPLATES / f"{name}.png")
    return {"name": name}


@app.post("/api/admin/templates/{name}/delete")
def delete_template(name: str, request: Request):
    need_admin(request)
    if name not in ppt_templates():
        raise HTTPException(404)
    (TEMPLATES / f"{name}.pptx").unlink(missing_ok=True)
    (TEMPLATES / f"{name}.png").unlink(missing_ok=True)
    return {"ok": True}


# ----------------------------------------------------------------- pages --
STATIC = {"disha-logo.png": "image/png", "favicon.png": "image/png", "island.css": "text/css",
          "island.js": "text/javascript"}


@app.get("/static/{name}")
def static(name: str):
    if name not in STATIC:
        raise HTTPException(404)
    return FileResponse(HERE / "static" / name, media_type=STATIC[name],
                        headers={"Cache-Control": "public, max-age=86400"})


@app.get("/favicon.ico")
def favicon():
    return FileResponse(HERE / "static" / "favicon.png", media_type="image/png")


@app.get("/healthz")
def health():
    return {"ok": True, "api_key": bool(os.environ.get("ANTHROPIC_API_KEY", "").strip())}


@app.get("/admin")
def admin_page(request: Request):
    if db.user_count() == 0:
        return RedirectResponse("/setup", status_code=303)
    if not admin_user(request):
        return RedirectResponse("/admin/login", status_code=303)
    return _page("admin.html")


@app.get("/")
def home(request: Request):
    if db.user_count() == 0:
        return RedirectResponse("/setup", status_code=303)
    if not current_user(request):
        return RedirectResponse("/login", status_code=303)
    return _page("app.html")


@app.exception_handler(HTTPException)
def http_err(request: Request, exc: HTTPException):
    if exc.status_code in (401, 403) and not request.url.path.startswith("/api/") or (
            exc.status_code in (401, 403) and request.method == "GET" and "download" in request.url.path):
        return RedirectResponse("/admin/login" if _panel(request) == "admin" else "/login", status_code=303)
    return JSONResponse({"detail": exc.detail}, status_code=exc.status_code)
