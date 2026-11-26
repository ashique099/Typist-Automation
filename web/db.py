"""SQLite storage: users, login sessions, jobs (work history) and settings."""
from __future__ import annotations

import hashlib
import hmac
import json
import os
import secrets
import sqlite3
import threading
import time

_lock = threading.Lock()
_conn: sqlite3.Connection | None = None

SCHEMA = """
create table if not exists users (
  id integer primary key, username text unique not null, name text not null,
  pw text not null, role text not null default 'typist', active integer not null default 1,
  created real not null);
create table if not exists sessions (
  token text primary key, user_id integer not null, created real not null, last_seen real not null);
create table if not exists jobs (
  id text primary key, user_id integer not null, title text, files text not null,
  status text not null, step text, error text, outputs text not null default '{}',
  options text not null default '{}', chapter text, pages integer default 0, pages_sent integer default 0,
  mcqs integer default 0, model text, calls integer default 0, input_tokens integer default 0,
  output_tokens integer default 0, cost_usd real default 0, created real not null, started real,
  finished real, updated real);
create index if not exists jobs_created on jobs(created);
create table if not exists settings (key text primary key, value text not null);
"""

DEFAULTS = {
    "monthly_budget_inr": "10000",
    "usd_to_inr": "90",            # includes a little for bank/forex charges
    "model": "claude-opus-5-5",
    "mcq_model": "claude-sonnet-5-5",
    "effort": "high",
    "default_mcq_count": "50",
}


def init(path):
    global _conn
    _conn = sqlite3.connect(path, check_same_thread=False, isolation_level=None)
    _conn.row_factory = sqlite3.Row
    _conn.execute("pragma journal_mode=wal")
    _conn.executescript(SCHEMA)
    return _conn


def q(sql, args=(), one=False):
    with _lock:
        cur = _conn.execute(sql, args)
        rows = cur.fetchall()
    rows = [dict(r) for r in rows]
    return (rows[0] if rows else None) if one else rows


def x(sql, args=()):
    with _lock:
        return _conn.execute(sql, args)


# ------------------------------------------------------------- passwords --
def hash_pw(pw: str) -> str:
    salt = secrets.token_bytes(16)
    h = hashlib.scrypt(pw.encode(), salt=salt, n=2 ** 14, r=8, p=1)
    return "scrypt$" + salt.hex() + "$" + h.hex()


def check_pw(pw: str, stored: str) -> bool:
    try:
        _, salt, h = stored.split("$")
        got = hashlib.scrypt(pw.encode(), salt=bytes.fromhex(salt), n=2 ** 14, r=8, p=1)
        return hmac.compare_digest(got.hex(), h)
    except (ValueError, TypeError):
        return False


# ----------------------------------------------------------------- users --
def user_count():
    return q("select count(*) n from users", one=True)["n"]


def create_user(username, name, pw, role="typist"):
    x("insert into users(username,name,pw,role,active,created) values(?,?,?,?,1,?)",
      (username.strip().lower(), name.strip() or username, hash_pw(pw), role, time.time()))
    return q("select * from users where username=?", (username.strip().lower(),), one=True)


def get_user(uid):
    return q("select id,username,name,role,active,created from users where id=?", (uid,), one=True)


def login(username, pw):
    u = q("select * from users where username=?", ((username or "").strip().lower(),), one=True)
    if not u or not u["active"] or not check_pw(pw or "", u["pw"]):
        return None
    tok = secrets.token_urlsafe(32)
    now = time.time()
    x("insert into sessions values(?,?,?,?)", (tok, u["id"], now, now))
    return tok


def session_user(tok, max_age_days=60):
    if not tok:
        return None
    s = q("select * from sessions where token=?", (tok,), one=True)
    if not s or s["last_seen"] < time.time() - max_age_days * 86400:
        return None
    u = get_user(s["user_id"])
    if not u or not u["active"]:
        return None
    if time.time() - s["last_seen"] > 300:
        x("update sessions set last_seen=? where token=?", (time.time(), tok))
    return u


def logout(tok):
    x("delete from sessions where token=?", (tok,))


# -------------------------------------------------------------- settings --
def setting(key):
    r = q("select value from settings where key=?", (key,), one=True)
    return r["value"] if r else DEFAULTS.get(key, "")


def all_settings():
    out = dict(DEFAULTS)
    out.update({r["key"]: r["value"] for r in q("select * from settings")})
    return out


def set_setting(key, value):
    x("insert into settings values(?,?) on conflict(key) do update set value=excluded.value", (key, str(value)))


# ------------------------------------------------------------------ jobs --
JSON_COLS = ("files", "outputs", "options")


def _job(row):
    if row:
        for c in JSON_COLS:
            row[c] = json.loads(row[c] or ("[]" if c == "files" else "{}"))
    return row


def add_job(jid, user_id, title, files, options):
    now = time.time()
    x("insert into jobs(id,user_id,title,files,status,step,options,created,updated) values(?,?,?,?,?,?,?,?,?)",
      (jid, user_id, title, json.dumps(files, ensure_ascii=False), "queued", "Queued",
       json.dumps(options, ensure_ascii=False), now, now))


def update_job(jid, **kw):
    kw["updated"] = time.time()
    for c in JSON_COLS:
        if c in kw:
            kw[c] = json.dumps(kw[c], ensure_ascii=False)
    cols = ",".join(f"{k}=?" for k in kw)
    x(f"update jobs set {cols} where id=?", (*kw.values(), jid))


def get_job(jid):
    return _job(q("select j.*, u.name as by_name from jobs j left join users u on u.id=j.user_id where j.id=?",
                  (jid,), one=True))


def list_jobs(user_id=None, limit=50, offset=0, status=None, since=None, search=None, task=None):
    where, args = [], []
    if task:
        where.append("coalesce(json_extract(j.options,'$.task'),'notes')=?"); args.append(task)
    if user_id:
        where.append("j.user_id=?"); args.append(user_id)
    if status:
        where.append("j.status=?"); args.append(status)
    if since:
        where.append("j.created>=?"); args.append(since)
    if search:
        where.append("(j.title like ? or j.chapter like ? or j.files like ?)"); args += [f"%{search}%"] * 3
    w = ("where " + " and ".join(where)) if where else ""
    rows = q(f"select j.*, u.name as by_name from jobs j left join users u on u.id=j.user_id {w} "
             f"order by j.created desc limit ? offset ?", (*args, limit, offset))
    total = q(f"select count(*) n from jobs j {w}", args, one=True)["n"]
    return [_job(r) for r in rows], total


def month_start(ts=None):
    t = time.localtime(ts or time.time())
    return time.mktime((t.tm_year, t.tm_mon, 1, 0, 0, 0, 0, 0, -1))


def month_spend_usd():
    return q("select coalesce(sum(cost_usd),0) s from jobs where created>=?", (month_start(),), one=True)["s"]


def usage_by_user(since):
    return q("select u.id, u.name, u.username, count(j.id) jobs, coalesce(sum(j.pages),0) pages, "
             "coalesce(sum(j.mcqs),0) mcqs, coalesce(sum(j.cost_usd),0) cost_usd "
             "from users u left join jobs j on j.user_id=u.id and j.created>=? group by u.id order by cost_usd desc",
             (since,))
