"""Account rescue from the server PC (no browser needed).

    python -m web.manage users                      list accounts (no passwords: they are stored hashed)
    python -m web.manage reset USERNAME             set a new random temporary password and print it
    python -m web.manage reset USERNAME NEWPASS     set this password
    python -m web.manage add-admin USERNAME "Name"  create an extra admin with a temporary password

Passwords cannot be read back, only replaced. Change a temporary password after signing in
(My account / Settings > My password). Uses DATA_DIR like the website (default ./data).
"""
from __future__ import annotations

import os
import secrets
import sys
import time
from pathlib import Path

from . import db


def _temp():
    return "Disha-" + secrets.token_hex(3)


def main(argv=None):
    a = list(sys.argv[1:] if argv is None else argv)
    data = Path(os.environ.get("DATA_DIR", "data")).resolve()
    if not (data / "typist.db").exists():
        print(f"No database at {data / 'typist.db'}. Open the website once and create the admin on /setup.")
        return 1
    db.init(data / "typist.db")
    if not a or a[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    if a[0] == "users":
        rows = db.q("select id, username, name, role, active, created from users order by role, name")
        if not rows:
            print("No accounts yet. Open http://localhost:8000/setup to create the admin.")
        for r in rows:
            print(f"{r['username']:<20} {r['name']:<24} {r['role']:<7} {'active' if r['active'] else 'DISABLED'}"
                  f"  created {time.strftime('%Y-%m-%d', time.localtime(r['created']))}")
        return 0
    if a[0] == "reset" and len(a) >= 2:
        u = db.q("select id from users where username=?", (a[1].strip().lower(),), one=True)
        if not u:
            print(f"No user '{a[1]}'. Run: python -m web.manage users")
            return 1
        pw = a[2] if len(a) > 2 else _temp()
        if len(pw) < 6:
            print("Password must be at least 6 characters.")
            return 1
        db.x("update users set pw=?, active=1 where id=?", (db.hash_pw(pw), u["id"]))
        db.x("delete from sessions where user_id=?", (u["id"],))
        print(f"Password for '{a[1].lower()}' is now: {pw}")
        print("Sign in and change it from My account.")
        return 0
    if a[0] == "add-admin" and len(a) >= 2:
        if db.q("select 1 from users where username=?", (a[1].strip().lower(),), one=True):
            print("That username already exists; use reset instead.")
            return 1
        pw = _temp()
        db.create_user(a[1], a[2] if len(a) > 2 else a[1], pw, role="admin")
        print(f"Admin '{a[1].lower()}' created with temporary password: {pw}")
        return 0
    print(__doc__)
    return 1


if __name__ == "__main__":
    sys.exit(main())
