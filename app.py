"""
MailMerge Studio — multi-user web app.
Google Sign-In, per-user SMTP/lists/templates/logs, admin panel.
"""
import json
import os
import re
import secrets
import smtplib
import ssl
import threading
import time
import urllib.parse
import urllib.request
import uuid
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr, parseaddr

from flask import (Flask, jsonify, redirect, render_template, request,
                   send_from_directory, session, url_for)
from openpyxl import load_workbook

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.environ.get("DATA_DIR") or os.path.join(BASE, "data")
USERS_DIR = os.path.join(DATA, "users")
UPLOADS = os.path.join(DATA, "uploads")          # shared temp imports
os.makedirs(USERS_DIR, exist_ok=True)
os.makedirs(UPLOADS, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024
app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(32)

GOOGLE_CLIENT_ID = os.environ.get("GOOGLE_CLIENT_ID", "")
GOOGLE_CLIENT_SECRET = os.environ.get("GOOGLE_CLIENT_SECRET", "")
ADMIN_EMAILS = [e.strip().lower() for e in os.environ.get("ADMIN_EMAILS", "").split(",") if e.strip()]
BASE_URL = os.environ.get("BASE_URL", "http://127.0.0.1:5000").rstrip("/")


def udir(uid, sub=""):
    d = os.path.join(USERS_DIR, uid, sub)
    os.makedirs(d, exist_ok=True)
    return d


# ---------------- user accounts ----------------

def load_user(uid):
    p = os.path.join(USERS_DIR, uid, "profile.json")
    if os.path.exists(p):
        try:
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return None


def save_user(u):
    os.makedirs(os.path.join(USERS_DIR, u["id"]), exist_ok=True)
    with open(os.path.join(USERS_DIR, u["id"], "profile.json"), "w", encoding="utf-8") as f:
        json.dump(u, f, indent=2)


def current_user():
    uid = session.get("uid")
    if not uid:
        return None
    u = load_user(uid)
    if u and u.get("blocked"):
        return None
    return u


def is_admin(u):
    return bool(u) and (u["email"].lower() in ADMIN_EMAILS or u.get("admin"))


class SafeJSONEncoder(json.JSONEncoder):
    def default(self, o):
        try:
            import datetime as _dt
            if isinstance(o, (_dt.datetime, _dt.date)):
                return o.strftime("%d/%m/%Y")
        except Exception:
            pass
        return str(o)


# ---------------- username / password auth ----------------

def hash_pwd(pwd, salt=None):
    salt = salt or secrets.token_hex(16)
    import hashlib
    h = hashlib.pbkdf2_hmac("sha256", pwd.encode(), bytes.fromhex(salt), 200_000).hex()
    return salt, h


def verify_pwd(pwd, salt, expect):
    return secrets.compare_digest(hash_pwd(pwd, salt)[1], expect)


@app.route("/api/auth/register", methods=["POST"])
def api_register():
    """Open registration only until an admin exists; then admin-only."""
    d = request.json or {}
    username = (d.get("username") or "").strip().lower()
    name = (d.get("name") or username).strip()
    pwd = d.get("password") or ""
    if not re.fullmatch(r"[a-z0-9_.-]{3,30}", username):
        return jsonify(error="Username: 3-30 chars, letters/numbers/._- only"), 400
    if len(pwd) < 6:
        return jsonify(error="Password must be at least 6 characters"), 400
    if os.path.exists(os.path.join(USERS_DIR, username, "profile.json")):
        return jsonify(error="That username is taken"), 400
    any_admin = any((load_user(x) or {}).get("admin") for x in os.listdir(USERS_DIR))
    # First ever account becomes admin; afterwards only admins may register others.
    u = current_user()
    if any_admin and not is_admin(u):
        return jsonify(error="Registration is closed. Ask an admin to create your account."), 403
    salt, h = hash_pwd(pwd)
    user = {"id": username, "username": username, "email": name, "name": name,
            "pwd_salt": salt, "pwd_hash": h,
            "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "admin": not any_admin, "blocked": False}
    save_user(user)
    session["uid"] = username
    return jsonify(ok=True)


@app.route("/api/auth/login", methods=["POST"])
def api_auth_login():
    d = request.json or {}
    username = (d.get("username") or "").strip().lower()
    u = load_user(username)
    if not u or not u.get("pwd_hash") or not verify_pwd(d.get("password") or "", u["pwd_salt"], u["pwd_hash"]):
        return jsonify(error="Wrong username or password"), 401
    if u.get("blocked"):
        return jsonify(error="Your account is blocked. Contact the administrator."), 403
    session["uid"] = username
    return jsonify(ok=True)


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")


@app.route("/api/me")
def api_me():
    u = current_user()
    if not u:
        return jsonify(loggedIn=False)
    return jsonify(loggedIn=True, email=u.get("email") or u["id"], name=u["name"], admin=is_admin(u))


# ---------------- per-user workspace ----------------

STATE = {}  # uid -> {rows, columns, source_file, config, campaign, email_col}


def ws(u):
    uid = u["id"]
    if uid not in STATE:
        st = {"rows": [], "columns": [], "source_file": None, "config": {}, "campaign": None, "email_col": None}
        cfg_path = os.path.join(udir(uid), "config.json")
        if not os.path.exists(cfg_path):
            # Migrate pre-multi-user shared config to this user's workspace (once)
            legacy = os.path.join(DATA, "config.json")
            if os.path.exists(legacy):
                try:
                    with open(legacy, "r", encoding="utf-8") as f:
                        st["config"] = json.load(f)
                    save_config(u, st)
                except Exception:
                    pass
        if os.path.exists(cfg_path):
            try:
                with open(cfg_path, "r", encoding="utf-8") as f:
                    st["config"] = json.load(f)
            except Exception:
                pass
        # auto-restore last used list
        last = st["config"].get("last_list")
        if last:
            lp = os.path.join(udir(uid, "lists"), re.sub(r"[^A-Za-z0-9_-]+", "_", str(last).strip())[:60] + ".json")
            if os.path.exists(lp):
                try:
                    with open(lp, "r", encoding="utf-8") as f:
                        d = json.load(f)
                    st["rows"], st["columns"] = d.get("rows", []), d.get("columns", [])
                    st["source_file"] = d.get("source", last)
                    st["email_col"] = find_email_column(st["columns"])
                except Exception:
                    pass
        STATE[uid] = st
    return STATE[uid]


def slugify(s):
    return re.sub(r"[^A-Za-z0-9_-]+", "_", s.strip())[:60] or "list"


def find_email_column(cols):
    for exact in ("email", "email id", "email_id", "emailid", "e-mail", "mail id", "mail"):
        for c in cols:
            if c.strip().lower() == exact:
                return c
    for c in cols:
        if "email" in c.strip().lower():
            return c
    return None


TAG_RE = re.compile(r"\{\{\s*([^}]+?)\s*\}\}")


def render(text, row):
    def sub(m):
        val = row.get(m.group(1), "")
        return "" if val is None else str(val)
    return TAG_RE.sub(sub, text or "")


def valid_email(a):
    try:
        addr = parseaddr(str(a))[1]
        return "@" in addr and "." in addr.split("@")[-1]
    except Exception:
        return False


# ---------------- data import ----------------

def parse_csv(path):
    import csv
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        sample = f.read(4096)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
        except Exception:
            dialect = csv.excel
        all_rows = list(csv.reader(f, dialect))
    if not all_rows:
        return [], []
    cols = [c.strip() for c in all_rows[0]]
    rows = []
    for r in all_rows[1:]:
        if not any(str(x).strip() for x in r):
            continue
        rows.append({cols[i] if i < len(cols) else f"Column{i+1}": v for i, v in enumerate(r)})
    return rows, cols


def parse_xlsx(path):
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    all_rows = [[("" if c is None else c) for c in row] for row in ws.iter_rows(values_only=True)]
    while all_rows and not any(str(x).strip() for x in all_rows[0]):
        all_rows.pop(0)
    if not all_rows:
        return [], []
    cols = [str(c).strip() if str(c).strip() else f"Column{i+1}" for i, c in enumerate(all_rows[0])]
    rows = []
    for r in all_rows[1:]:
        if not any(str(x).strip() for x in r):
            continue
        rows.append({c: ("" if i >= len(r) else r[i]) for i, c in enumerate(cols)})
    return rows, cols


@app.route("/api/import", methods=["POST"])
def api_import():
    u = current_user()
    if not u:
        return jsonify(error="Not logged in"), 401
    file = request.files.get("file")
    if not file:
        return jsonify(error="No file provided"), 400
    ext = os.path.splitext(file.filename or "")[1].lower()
    path = os.path.join(UPLOADS, f"{uuid.uuid4().hex}{ext}")
    file.save(path)
    try:
        if ext in (".csv", ".tsv", ".txt"):
            rows, cols = parse_csv(path)
        elif ext in (".xlsx", ".xlsm"):
            rows, cols = parse_xlsx(path)
        else:
            return jsonify(error="Unsupported file type. Use .xlsx, .csv or .tsv"), 400
    except Exception as e:
        return jsonify(error=f"Could not read file: {e}"), 400
    if not rows or not cols:
        return jsonify(error="The file appears to be empty"), 400
    st = ws(u)
    st["rows"], st["columns"] = rows, cols
    st["source_file"] = file.filename
    st["email_col"] = find_email_column(cols)
    return jsonify(columns=cols, rows=rows, total=len(rows), source=file.filename,
                   emailColumn=st["email_col"])


# ---------------- attachments ----------------

@app.route("/api/attachments", methods=["GET", "POST"])
def api_attachments():
    u = current_user()
    if not u:
        return jsonify(error="Not logged in"), 401
    adir = udir(u["id"], "attachments")
    if request.method == "POST":
        saved = []
        for f in request.files.getlist("files"):
            name = f.filename or f"file_{uuid.uuid4().hex}"
            f.save(os.path.join(adir, os.path.basename(name)))
            saved.append(os.path.basename(name))
        return jsonify(saved=saved)
    return jsonify(files=sorted(os.listdir(adir)))


@app.route("/api/attachments/delete", methods=["POST"])
def api_attachment_delete():
    u = current_user()
    if not u:
        return jsonify(error="Not logged in"), 401
    p = os.path.join(udir(u["id"], "attachments"), os.path.basename(request.json.get("name", "")))
    if os.path.exists(p):
        os.remove(p)
    return jsonify(ok=True)


# ---------------- lists & templates ----------------

@app.route("/api/lists")
def api_lists():
    u = current_user()
    if not u:
        return jsonify(error="Not logged in"), 401
    out = []
    d = udir(u["id"], "lists")
    for f in sorted(os.listdir(d)):
        try:
            with open(os.path.join(d, f), "r", encoding="utf-8") as fh:
                x = json.load(fh)
            out.append({"name": x.get("name", f), "created": x.get("created", ""),
                        "count": len(x.get("rows", [])), "columns": len(x.get("columns", []))})
        except Exception:
            pass
    return jsonify(lists=out)


@app.route("/api/lists/save", methods=["POST"])
def api_lists_save():
    u = current_user()
    if not u:
        return jsonify(error="Not logged in"), 401
    st = ws(u)
    name = (request.json.get("name") or "").strip() or "My List"
    data = {"name": name, "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "columns": st["columns"], "rows": st["rows"], "source": st["source_file"]}
    with open(os.path.join(udir(u["id"], "lists"), slugify(name) + ".json"), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, cls=SafeJSONEncoder)
    st["config"]["last_list"] = name
    save_config(u, st)
    return jsonify(ok=True)


@app.route("/api/lists/load", methods=["POST"])
def api_lists_load():
    u = current_user()
    if not u:
        return jsonify(error="Not logged in"), 401
    name = request.json.get("name", "")
    p = os.path.join(udir(u["id"], "lists"), slugify(name) + ".json")
    if not os.path.exists(p):
        return jsonify(error="List not found"), 404
    with open(p, "r", encoding="utf-8") as f:
        d = json.load(f)
    st = ws(u)
    st["rows"], st["columns"] = d["rows"], d["columns"]
    st["source_file"] = d.get("source", name)
    st["email_col"] = find_email_column(st["columns"])
    st["config"]["last_list"] = name
    save_config(u, st)
    return jsonify(columns=st["columns"], rows=st["rows"], total=len(st["rows"]),
                   source=st["source_file"], emailColumn=st["email_col"])


@app.route("/api/lists/delete", methods=["POST"])
def api_lists_delete():
    u = current_user()
    if not u:
        return jsonify(error="Not logged in"), 401
    p = os.path.join(udir(u["id"], "lists"), slugify(request.json.get("name", "")) + ".json")
    if os.path.exists(p):
        os.remove(p)
    return jsonify(ok=True)


@app.route("/api/templates")
def api_templates():
    u = current_user()
    if not u:
        return jsonify(error="Not logged in"), 401
    out = []
    d = udir(u["id"], "templates")
    for f in sorted(os.listdir(d)):
        try:
            with open(os.path.join(d, f), "r", encoding="utf-8") as fh:
                x = json.load(fh)
            out.append({"name": x.get("name", f), "list": x.get("list", ""), "updated": x.get("updated", "")})
        except Exception:
            pass
    return jsonify(templates=out)


@app.route("/api/templates/save", methods=["POST"])
def api_templates_save():
    u = current_user()
    if not u:
        return jsonify(error="Not logged in"), 401
    d = request.json or {}
    name = (d.get("name") or "").strip() or "Untitled template"
    data = {"name": name, "subject": d.get("subject", ""), "html": d.get("html", ""),
            "list": d.get("list", ""), "updated": datetime.now().strftime("%Y-%m-%d %H:%M")}
    with open(os.path.join(udir(u["id"], "templates"), slugify(name) + ".json"), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, cls=SafeJSONEncoder)
    return jsonify(ok=True)


@app.route("/api/templates/load", methods=["POST"])
def api_templates_load():
    u = current_user()
    if not u:
        return jsonify(error="Not logged in"), 401
    p = os.path.join(udir(u["id"], "templates"), slugify(request.json.get("name", "")) + ".json")
    if not os.path.exists(p):
        return jsonify(error="Template not found"), 404
    with open(p, "r", encoding="utf-8") as f:
        return jsonify(json.load(f))


@app.route("/api/templates/delete", methods=["POST"])
def api_templates_delete():
    u = current_user()
    if not u:
        return jsonify(error="Not logged in"), 401
    p = os.path.join(udir(u["id"], "templates"), slugify(request.json.get("name", "")) + ".json")
    if os.path.exists(p):
        os.remove(p)
    return jsonify(ok=True)


# ---------------- config ----------------

def save_config(u, st):
    with open(os.path.join(udir(u["id"]), "config.json"), "w", encoding="utf-8") as f:
        json.dump(st["config"], f, indent=2, ensure_ascii=False)


@app.route("/api/config", methods=["GET", "POST"])
def api_config():
    u = current_user()
    if not u:
        return jsonify(error="Not logged in"), 401
    st = ws(u)
    if request.method == "POST":
        c = request.json or {}
        old_pwd = st["config"].get("smtp", {}).get("password", "")
        new_smtp = c.get("smtp")
        if isinstance(new_smtp, dict) and not new_smtp.get("password"):
            new_smtp["password"] = old_pwd
        st["config"].update(c)
        save_config(u, st)
        return jsonify(ok=True)
    cfg = json.loads(json.dumps(st["config"]))
    if "smtp" in cfg:
        cfg["smtp"]["password"] = "********" if cfg["smtp"].get("password") else ""
    return jsonify(cfg)


@app.route("/api/config/save-password", methods=["POST"])
def api_save_password():
    u = current_user()
    if not u:
        return jsonify(error="Not logged in"), 401
    st = ws(u)
    st["config"].setdefault("smtp", {})["password"] = request.json.get("password", "")
    save_config(u, st)
    return jsonify(ok=True)


PRESETS = {
    "gmail": {"host": "smtp.gmail.com", "port": 587, "security": "starttls"},
    "outlook": {"host": "smtp.office365.com", "port": 587, "security": "starttls"},
    "yahoo": {"host": "smtp.mail.yahoo.com", "port": 587, "security": "starttls"},
    "zoho": {"host": "smtp.zoho.com", "port": 587, "security": "starttls"},
    "icloud": {"host": "smtp.mail.me.com", "port": 587, "security": "starttls"},
    "custom": {"host": "", "port": 587, "security": "starttls"},
}


@app.route("/api/presets")
def api_presets():
    return jsonify(PRESETS)


@app.route("/api/test-connection", methods=["POST"])
def api_test_connection():
    u = current_user()
    if not u:
        return jsonify(error="Not logged in"), 401
    cfg = ws(u)["config"].get("smtp", {})
    host, port = cfg.get("host", ""), int(cfg.get("port", 587))
    user, pwd = cfg.get("user", ""), cfg.get("password", "")
    sec = cfg.get("security", "starttls")
    if not host or not user or not pwd:
        return jsonify(ok=False, error="Save your connection (host, email and password) first")
    try:
        if sec == "ssl":
            with smtplib.SMTP_SSL(host, port, timeout=15, context=ssl.create_default_context()) as s:
                s.login(user, pwd)
        else:
            with smtplib.SMTP(host, port, timeout=15) as s:
                s.ehlo()
                if sec == "starttls":
                    s.starttls(context=ssl.create_default_context())
                s.login(user, pwd)
        return jsonify(ok=True)
    except smtplib.SMTPAuthenticationError:
        return jsonify(ok=False, error="Authentication rejected — check the app password (535 error)")
    except Exception as e:
        return jsonify(ok=False, error=f"Could not reach server: {e}")


# ---------------- sending engine ----------------

class Campaign(threading.Thread):
    def __init__(self, user, payload):
        super().__init__(daemon=True)
        self.user = user
        self.payload = payload
        self.stop_flag = threading.Event()
        self.log = []
        self.status = "running"
        self.sent = self.failed = self.total = 0

    def run(self):
        try:
            self._run()
        except Exception as e:
            self.status = "error"
            self.log.append({"to": "-", "status": "failed", "error": f"Campaign crashed: {e}",
                             "at": datetime.now().isoformat(timespec="seconds")})
        finally:
            if self.log:
                with open(os.path.join(udir(self.user["id"]), "send_log.json"), "w", encoding="utf-8") as f:
                    json.dump(self.log, f, indent=2, ensure_ascii=False, cls=SafeJSONEncoder)
            if not self.payload.get("testMode"):
                bump_daily(self.user["id"], self.log)
            if self.status == "running":
                self.status = "done"

    def _run(self):
        p = self.payload
        cfg = ws(self.user)["config"].get("smtp", {})
        host, port = cfg.get("host", ""), int(cfg.get("port", 587))
        user, pwd = cfg.get("user", ""), cfg.get("password", "")
        sec = cfg.get("security", "starttls")
        from_name, from_addr = cfg.get("fromName", ""), cfg.get("fromEmail") or user
        reply_to = cfg.get("replyTo", "")

        st = ws(self.user)
        rows = st["rows"]
        selected = p.get("selectedRows")
        if p.get("testMode"):
            rows = [{}]
        elif selected is not None:
            rows = [st["rows"][i] for i in selected if 0 <= i < len(st["rows"])]

        self.total = len(rows)
        throttle = max(0.5, float(p.get("delaySeconds", 3)))
        batch, pause = max(1, int(p.get("batchSize", 40))), max(30, int(p.get("batchPauseSeconds", 120)))
        count = 0
        for row in rows:
            if self.stop_flag.is_set():
                self.status = "stopped"
                break
            if p.get("testMode"):
                to = p.get("testRecipient") or "test@example.com"
            else:
                ecol = st["email_col"]
                to = str(row.get(ecol) or row.get("Email") or row.get("email") or "").strip()
            if not valid_email(to):
                self.log.append({"to": to, "status": "failed", "error": "Invalid/missing email",
                                 "at": datetime.now().isoformat(timespec="seconds")})
                self.failed += 1
                count += 1
                continue
            subject, body = render(p.get("subject", ""), row), render(p.get("html", ""), row)
            files = list(p.get("attachments", []))
            if p.get("attachmentColumn"):
                extra = str(row.get(p["attachmentColumn"], "") or "").strip()
                if extra:
                    files += [x.strip() for x in extra.split(";") if x.strip()]
            ok, err = True, ""
            if not p.get("testMode"):
                try:
                    msg = EmailMessage()
                    msg["From"] = formataddr((from_name or from_addr, from_addr))
                    msg["To"] = to
                    if reply_to:
                        msg["Reply-To"] = reply_to
                    msg["Subject"] = subject
                    msg.set_content("Please view this email in an HTML-capable client.")
                    msg.add_alternative(body, subtype="html")
                    adir = udir(self.user["id"], "attachments")
                    for fn in files:
                        fp = os.path.join(adir, os.path.basename(fn))
                        if os.path.exists(fp):
                            with open(fp, "rb") as fh:
                                msg.add_attachment(fh.read(), maintype="application",
                                                   subtype="octet-stream", filename=os.path.basename(fn))
                    if sec == "ssl":
                        with smtplib.SMTP_SSL(host, port, timeout=30, context=ssl.create_default_context()) as s:
                            s.login(user, pwd)
                            s.send_message(msg)
                    else:
                        with smtplib.SMTP(host, port, timeout=30) as s:
                            s.ehlo()
                            if sec == "starttls":
                                s.starttls(context=ssl.create_default_context())
                            s.login(user, pwd)
                            s.send_message(msg)
                except Exception as e:
                    ok, err = False, str(e)
            self.log.append({"to": to, "status": "sent" if ok else "failed", "error": err,
                             "at": datetime.now().isoformat(timespec="seconds")})
            self.sent += int(ok)
            self.failed += int(not ok)
            count += 1
            time.sleep(throttle)
            if not p.get("testMode") and count % batch == 0 and count < self.total:
                for _ in range(pause):
                    if self.stop_flag.is_set():
                        break
                    time.sleep(1)
        if self.status != "stopped":
            self.status = "done"


def bump_daily(uid, entries):
    p = os.path.join(udir(uid), "daily_stats.json")
    stats = {}
    if os.path.exists(p):
        try:
            with open(p, "r", encoding="utf-8") as f:
                stats = json.load(f)
        except Exception:
            stats = {}
    for e in entries:
        day = str(e.get("at", ""))[:10]
        if not day:
            continue
        d = stats.setdefault(day, {"sent": 0, "failed": 0})
        d["sent" if e.get("status") == "sent" else "failed"] += 1
    with open(p, "w", encoding="utf-8") as f:
        json.dump(stats, f, indent=2)


@app.route("/api/send", methods=["POST"])
def api_send():
    u = current_user()
    if not u:
        return jsonify(error="Not logged in"), 401
    st = ws(u)
    if st["campaign"] and st["campaign"].is_alive():
        return jsonify(error="A campaign is already running"), 409
    camp = Campaign(u, request.json)
    st["campaign"] = camp
    camp.start()
    return jsonify(id="ok")


@app.route("/api/stop", methods=["POST"])
def api_stop():
    u = current_user()
    if not u:
        return jsonify(error="Not logged in"), 401
    c = ws(u).get("campaign")
    if c:
        c.stop_flag.set()
    return jsonify(ok=True)


@app.route("/api/progress")
def api_progress():
    u = current_user()
    if not u:
        return jsonify(error="Not logged in"), 401
    c = ws(u).get("campaign")
    if not c:
        prev = []
        lp = os.path.join(udir(u["id"]), "send_log.json")
        if os.path.exists(lp):
            try:
                with open(lp, "r", encoding="utf-8") as f:
                    prev = json.load(f)
            except Exception:
                pass
        return jsonify(running=False, status="idle", log=prev[-200:],
                       sent=sum(1 for x in prev if x["status"] == "sent"),
                       failed=sum(1 for x in prev if x["status"] == "failed"), total=len(prev))
    return jsonify(running=c.is_alive(), status=c.status, sent=c.sent, failed=c.failed,
                   total=c.total, log=c.log[-200:])


DAILY_LIMIT = 500


@app.route("/api/daily")
def api_daily():
    u = current_user()
    if not u:
        return jsonify(error="Not logged in"), 401
    p = os.path.join(udir(u["id"]), "daily_stats.json")
    stats = {}
    if os.path.exists(p):
        try:
            with open(p, "r", encoding="utf-8") as f:
                stats = json.load(f)
        except Exception:
            pass
    today = datetime.now().strftime("%Y-%m-%d")
    return jsonify(limit=DAILY_LIMIT, today=stats.get(today, {"sent": 0, "failed": 0}),
                   days=sorted(stats.items(), reverse=True)[:30])


@app.route("/api/log/export")
def api_log_export():
    u = current_user()
    if not u:
        return jsonify(error="Not logged in"), 401
    return send_from_directory(udir(u["id"]), "send_log.json", as_attachment=True, download_name="send_report.json")


@app.route("/api/state")
def api_state():
    u = current_user()
    if not u:
        return jsonify(error="Not logged in"), 401
    st = ws(u)
    return jsonify(columns=st["columns"], rows=st["rows"], total=len(st["rows"]),
                   source=st["source_file"], emailColumn=st["email_col"],
                   lastList=st["config"].get("last_list", ""))


# ---------------- admin panel ----------------

@app.route("/api/admin/users")
def api_admin_users():
    u = current_user()
    if not is_admin(u):
        return jsonify(error="Admin only"), 403
    out = []
    for d in os.listdir(USERS_DIR):
        usr = load_user(d)
        if usr:
            dp = os.path.join(USERS_DIR, d, "daily_stats.json")
            sent = 0
            if os.path.exists(dp):
                try:
                    with open(dp, "r", encoding="utf-8") as f:
                        sent = sum(v["sent"] for v in json.load(f).values())
                except Exception:
                    pass
            out.append({"name": usr.get("name"), "email": usr["email"], "created": usr.get("created"),
                        "admin": usr.get("admin", False), "blocked": usr.get("blocked", False), "sent": sent})
    return jsonify(users=out)


@app.route("/api/admin/create-user", methods=["POST"])
def api_admin_create_user():
    u = current_user()
    if not is_admin(u):
        return jsonify(error="Admin only"), 403
    d = request.json or {}
    username = (d.get("username") or "").strip().lower()
    pwd = d.get("password") or ""
    if not re.fullmatch(r"[a-z0-9_.-]{3,30}", username):
        return jsonify(error="Username: 3-30 chars, letters/numbers/._- only"), 400
    if len(pwd) < 6:
        return jsonify(error="Password must be at least 6 characters"), 400
    if os.path.exists(os.path.join(USERS_DIR, username, "profile.json")):
        return jsonify(error="That username is taken"), 400
    salt, h = hash_pwd(pwd)
    save_user({"id": username, "username": username, "email": (d.get("name") or username).strip(),
               "name": (d.get("name") or username).strip(), "pwd_salt": salt, "pwd_hash": h,
               "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
               "admin": bool(d.get("admin")), "blocked": False})
    return jsonify(ok=True)


@app.route("/api/admin/toggle-block", methods=["POST"])
def api_admin_toggle():
    u = current_user()
    if not is_admin(u):
        return jsonify(error="Admin only"), 403
    target = load_user(request.json.get("username", ""))
    if target:
        if target.get("admin"):
            return jsonify(error="Cannot block an admin account"), 400
        target["blocked"] = not target.get("blocked", False)
        save_user(target)
        return jsonify(ok=True, blocked=target["blocked"])
    return jsonify(error="User not found"), 404


# ---------------- front page ----------------

@app.route("/")
def index():
    if not current_user():
        return send_from_directory(os.path.join(BASE, "templates"), "login.html")
    return send_from_directory(os.path.join(BASE, "templates"), "index.html")


if __name__ == "__main__":
    print("\n  MailMerge Studio is running →  http://127.0.0.1:5000\n")
    app.run(host="127.0.0.1", port=5000, debug=False)