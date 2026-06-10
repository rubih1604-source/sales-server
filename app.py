import os, json, re, sqlite3
import requests as req
from flask import Flask, request, session, jsonify, redirect, render_template
from urllib.parse import urlencode
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from datetime import datetime, date, timedelta

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "dev-secret")

GOOGLE_CLIENT_ID     = os.environ.get("GOOGLE_CLIENT_ID")
GOOGLE_CLIENT_SECRET = os.environ.get("GOOGLE_CLIENT_SECRET")
REDIRECT_URI         = os.environ.get("REDIRECT_URI", "http://localhost:5000/oauth/callback")
DB_PATH              = os.environ.get("DB_PATH", "sales.db")

SCOPES = "openid email https://www.googleapis.com/auth/gmail.readonly"
OPS_SENDERS = ["oshrityes2901@gmail.com", "oritapiro22@gmail.com", "avielv014@gmail.com"]

HOURS_RE = r'(\d{1,2}(?::\d{2})?[-–]\d{1,2}(?::\d{2})?)'

def get_db():
    db = sqlite3.connect(DB_PATH)
    db.row_factory = sqlite3.Row
    return db

def init_db():
    db = get_db()
    db.executescript("""
    CREATE TABLE IF NOT EXISTS sales (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        thread_id TEXT UNIQUE NOT NULL,
        name TEXT, customer_id TEXT, sale_date TEXT,
        install_date TEXT, install_hours TEXT, converters INTEGER,
        package TEXT, status TEXT, contract_ok INTEGER DEFAULT 0,
        notes TEXT, last_scanned TEXT, updated_at TEXT
    );
    CREATE TABLE IF NOT EXISTS scan_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        scanned_at TEXT, total INTEGER, new_rows INTEGER, updated INTEGER, month_ctx TEXT
    );
    CREATE TABLE IF NOT EXISTS install_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        thread_id TEXT, old_date TEXT, new_date TEXT, changed_at TEXT, reason TEXT
    );
    """)
    db.commit()
    db.close()

@app.route("/login")
def login():
    params = {
        "client_id": GOOGLE_CLIENT_ID,
        "redirect_uri": REDIRECT_URI,
        "response_type": "code",
        "scope": SCOPES,
        "access_type": "offline",
        "prompt": "consent",
    }
    return redirect("https://accounts.google.com/o/oauth2/v2/auth?" + urlencode(params))

@app.route("/oauth/callback")
def oauth_callback():
    code = request.args.get("code")
    if not code:
        return "Login failed", 400
    resp = req.post("https://oauth2.googleapis.com/token", data={
        "code": code,
        "client_id": GOOGLE_CLIENT_ID,
        "client_secret": GOOGLE_CLIENT_SECRET,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code",
    })
    tokens = resp.json()
    if "access_token" not in tokens:
        return f"Token error: {tokens}", 400
    session["access_token"]  = tokens["access_token"]
    session["refresh_token"] = tokens.get("refresh_token", "")
    return redirect("/")

def get_gmail():
    if "access_token" not in session:
        return None
    creds = Credentials(
        token=session["access_token"],
        refresh_token=session.get("refresh_token"),
        token_uri="https://oauth2.googleapis.com/token",
        client_id=GOOGLE_CLIENT_ID,
        client_secret=GOOGLE_CLIENT_SECRET,
    )
    return build("gmail", "v1", credentials=creds, cache_discovery=False)

def normalize_hours(h):
    h = str(h).replace("–", "-").strip()
    h = re.sub(r":00", "", h)
    parts = h.split("-")
    if len(parts) == 2:
        try:
            return f"{int(parts[0])}-{int(parts[1])}"
        except:
            pass
    return h

def normalize_date(d):
    parts = str(d).strip().split("/")
    if len(parts) == 2:
        return f"{parts[0].zfill(2)}/{parts[1].zfill(2)}"
    return str(d)

def parse_install_from_text(text):
    """חילוץ תאריך+שעות מטקסט עברי."""
    n = text
    n = re.sub(r"ב(\d{1,2}/\d{1,2})", r"\1", n)
    n = re.sub(r"ל(\d{1,2}/\d{1,2})", r"\1", n)
    n = re.sub(r"בין\s+", "", n)
    n = re.sub(r"מתואם", "תואם", n)
    n = re.sub(r"שיבוץ", "שובץ", n)

    best_date = ""
    best_hours = ""

    # תואם/שובץ DD/MM HH-HH
    for m in re.finditer(r"(?:תואם|שובץ|מוקלד)\s+(\d{1,2}/\d{1,2})\s+" + HOURS_RE, n):
        best_date = normalize_date(m.group(1))
        best_hours = normalize_hours(m.group(2))

    # לקוח XXXXXXX תואם DD/MM HH-HH
    for m in re.finditer(r"לקוח\s+\d+\s+(?:תואם|שובץ)\s+(\d{1,2}/\d{1,2})\s+" + HOURS_RE, n):
        best_date = normalize_date(m.group(1))
        best_hours = normalize_hours(m.group(2))

    # DD/MM HH-HH
    if not best_date:
        for m in re.finditer(r"(\d{1,2}/\d{1,2})\s+" + HOURS_RE, n):
            best_date = normalize_date(m.group(1))
            best_hours = normalize_hours(m.group(2))

    # HH-HH DD/MM
    if not best_date:
        for m in re.finditer(HOURS_RE + r"\s+(\d{1,2}/\d{1,2})", n):
            best_date = normalize_date(m.group(2))
            best_hours = normalize_hours(m.group(1))

    return best_date, best_hours

def ddmm_to_iso(ddmm):
    """14/06 → 2026-06-14"""
    try:
        parts = ddmm.split("/")
        return f"2026-{parts[1].zfill(2)}-{parts[0].zfill(2)}"
    except:
        return ""

def parse_thread(thread_id, messages):
    """ניתוח thread עם regex – בלי AI."""
    full_text = ""
    first_subject = ""
    sale_date_raw = ""

    for i, msg in enumerate(messages):
        hdrs = {h["name"]: h["value"] for h in msg.get("payload", {}).get("headers", [])}
        subj = hdrs.get("Subject", "")
        if i == 0:
            first_subject = subj
            sale_date_raw = hdrs.get("Date", "")
        full_text += subj + "\n" + msg.get("snippet", "") + "\n"

    # בדוק שזו מכירה
    sale_kw = ["תואם", "לאשר חוזה", "ממירים", "שובץ", "להקים", "מוקלד", "אישר"]
    if not any(k in full_text for k in sale_kw):
        return None

    # שם לקוח מנושא
    name = re.sub(r"^(Re|Fwd|FW|RE):\s*", "", first_subject, flags=re.IGNORECASE).strip()
    name = re.sub(r"[-–].*", "", name).strip()
    if not name or len(name) < 2:
        return None

    # מספר לקוח
    customer_id = ""
    m = re.search(r"לקוח\s+(\d{7})", full_text)
    if m:
        customer_id = m.group(1)

    # ממירים
    converters = 0
    m = re.search(r"(\d+)\s*ממירים?", full_text)
    if m:
        converters = int(m.group(1))

    # חבילה
    package = "other"
    for p in ["דאבל יס פלוס", "דאבל יס", "רק יס", "דאבל סטינג", "דאבל סיבים", "מסך בלבד"]:
        if p in full_text:
            package = p
            break

    # תאריך התקנה – מהודעת תפעול האחרונה
    install_date_ddmm = ""
    install_hours = ""
    for msg in reversed(messages):
        sender = ""
        for h in msg.get("payload", {}).get("headers", []):
            if h["name"] == "From":
                sender = h["value"]
        if any(op in sender for op in OPS_SENDERS):
            snippet = msg.get("snippet", "")
            d, h = parse_install_from_text(snippet)
            if d:
                install_date_ddmm = d
                install_hours = h
                break

    # אם לא מצאנו בהודעת תפעול – חפש בכל הטקסט
    if not install_date_ddmm:
        install_date_ddmm, install_hours = parse_install_from_text(full_text)

    install_date_iso = ddmm_to_iso(install_date_ddmm) if install_date_ddmm else ""

    # סטטוס
    cancel_kw = ["ביטול", "נשלח לביטול", "לא ניתן להקים", "בוטל", "ביטל", "ביטלה"]
    if any(k in full_text for k in cancel_kw):
        status = "cancelled"
    elif "מוקלד" in full_text:
        status = "recorded"
    elif any(k in full_text for k in ["אישר", "אישרה"]):
        status = "approved"
    elif any(k in full_text for k in ["דוור", "שיבוץ נשלח", "שולחת שיבוץ"]):
        status = "sent"
    elif "לאשר חוזה" in full_text:
        status = "waiting"
    else:
        status = "waiting"

    # חוזה
    contract_ok = 1 if any(k in full_text for k in ["אישר", "אישרה", "מוקלד"]) else 0

    # תאריך מכירה
    sale_date = ""
    try:
        from email.utils import parsedate
        p = parsedate(sale_date_raw)
        if p:
            sale_date = f"{p[0]}-{p[1]:02d}-{p[2]:02d}"
    except:
        pass

    return {
        "thread_id": thread_id,
        "name": name,
        "customer_id": customer_id,
        "sale_date": sale_date,
        "install_date": install_date_iso,
        "install_hours": install_hours,
        "converters": converters,
        "package": package,
        "status": status,
        "contract_ok": contract_ok,
        "notes": "",
    }

def sync_to_db(parsed, current_month):
    db = get_db()
    now = datetime.utcnow().isoformat()
    stats = {"new": 0, "updated": 0, "unchanged": 0}

    for row in parsed:
        tid = row.get("thread_id")
        if not tid:
            continue
        existing = db.execute("SELECT * FROM sales WHERE thread_id=?", (tid,)).fetchone()
        install_date = row.get("install_date") or ""

        if existing is None:
            db.execute("""INSERT INTO sales
                (thread_id,name,customer_id,sale_date,install_date,install_hours,
                 converters,package,status,contract_ok,notes,last_scanned,updated_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (tid, row.get("name",""), row.get("customer_id",""), row.get("sale_date",""),
                 install_date, row.get("install_hours",""), row.get("converters"),
                 row.get("package",""), row.get("status","waiting"),
                 row.get("contract_ok", 0), row.get("notes",""), now, now))
            stats["new"] += 1
        else:
            changes = {}
            old_inst = existing["install_date"] or ""
            if install_date and install_date != old_inst:
                changes["install_date"] = install_date
                db.execute(
                    "INSERT INTO install_history (thread_id,old_date,new_date,changed_at,reason) VALUES (?,?,?,?,?)",
                    (tid, old_inst, install_date, now, "סריקה אוטומטית")
                )
            new_status = row.get("status", "")
            if new_status and new_status != existing["status"]:
                changes["status"] = new_status
            new_c = row.get("contract_ok", 0)
            if new_c != existing["contract_ok"]:
                changes["contract_ok"] = new_c
            if changes:
                changes["updated_at"] = now
                changes["last_scanned"] = now
                set_clause = ", ".join(f"{k}=?" for k in changes)
                db.execute(f"UPDATE sales SET {set_clause} WHERE thread_id=?", (*changes.values(), tid))
                stats["updated"] += 1
            else:
                db.execute("UPDATE sales SET last_scanned=? WHERE thread_id=?", (now, tid))
                stats["unchanged"] += 1

    db.execute(
        "INSERT INTO scan_log (scanned_at,total,new_rows,updated,month_ctx) VALUES (?,?,?,?,?)",
        (now, len(parsed), stats["new"], stats["updated"], current_month)
    )
    db.commit()
    db.close()
    return stats

@app.route("/")
def index():
    logged_in = "access_token" in session
    if logged_in:
        init_db()
    return render_template("index.html", logged_in=logged_in)

@app.route("/api/scan", methods=["POST"])
def api_scan():
    gmail = get_gmail()
    if not gmail:
        return jsonify({"error": "not_logged_in"}), 401

    today = date.today()
    current_month = today.strftime("%Y-%m")
    body = request.get_json() or {}
    days_back = 180 if body.get("full") else 70
    after = (today - timedelta(days=days_back)).strftime("%Y/%m/%d")

    query = (
        f"after:{after} "
        "(from:oshrityes2901@gmail.com OR from:oritapiro22@gmail.com OR "
        "from:avielv014@gmail.com OR to:oshrityes2901@gmail.com)"
    )

    results = gmail.users().messages().list(
        userId="me", q=query, maxResults=200
    ).execute()

    threads_seen = set()
    all_parsed = []

    for msg in results.get("messages", []):
        tid = msg["threadId"]
        if tid in threads_seen:
            continue
        threads_seen.add(tid)
        try:
            thread = gmail.users().threads().get(
                userId="me", id=tid, format="metadata",
                metadataHeaders=["Subject", "From", "Date"]
            ).execute()
            sale = parse_thread(tid, thread.get("messages", []))
            if sale:
                all_parsed.append(sale)
        except:
            continue

    stats = sync_to_db(all_parsed, current_month)
    db = get_db()
    total = db.execute("SELECT COUNT(*) FROM sales").fetchone()[0]
    db.close()
    return jsonify({
        "ok": True,
        "scanned": len(all_parsed),
        "new": stats["new"],
        "updated": stats["updated"],
        "unchanged": stats["unchanged"],
        "total_in_db": total,
        "scan_month": current_month
    })

@app.route("/api/sales")
def api_sales():
    db = get_db()
    month  = request.args.get("month", "")
    status = request.args.get("status", "")
    search = request.args.get("q", "")
    q = "SELECT * FROM sales WHERE 1=1"
    params = []
    if month:
        q += " AND install_date LIKE ?"
        params.append(f"{month}%")
    if status:
        q += " AND status=?"
        params.append(status)
    if search:
        q += " AND (name LIKE ? OR customer_id LIKE ?)"
        params += [f"%{search}%", f"%{search}%"]
    q += " ORDER BY install_date DESC, sale_date DESC"
    rows = db.execute(q, params).fetchall()
    db.close()
    return jsonify([dict(r) for r in rows])

@app.route("/api/stats")
def api_stats():
    db = get_db()
    today = date.today()
    this_month = today.strftime("%Y-%m")
    last_month = (today.replace(day=1) - timedelta(days=1)).strftime("%Y-%m")

    def count(where, params=[]):
        return db.execute(f"SELECT COUNT(*) FROM sales WHERE {where}", params).fetchone()[0]

    stats = {
        "total": count("1=1"),
        "this_month": count("install_date LIKE ?", [f"{this_month}%"]),
        "last_month": count("install_date LIKE ?", [f"{last_month}%"]),
        "waiting": count("status='waiting'"),
        "approved": count("status='approved'"),
        "recorded": count("status='recorded'"),
        "cancelled": count("status='cancelled'"),
        "no_contract": count("contract_ok=0 AND status NOT IN ('cancelled')"),
        "current_month": this_month,
        "last_month": last_month,
    }
    months = db.execute("""
        SELECT substr(install_date,1,7) as m, COUNT(*) as cnt
        FROM sales WHERE install_date != ''
        GROUP BY m ORDER BY m DESC LIMIT 6
    """).fetchall()
    stats["by_month"] = [dict(r) for r in months]
    last_scan = db.execute(
        "SELECT scanned_at,total,new_rows,updated FROM scan_log ORDER BY id DESC LIMIT 1"
    ).fetchone()
    stats["last_scan"] = dict(last_scan) if last_scan else None
    changes = db.execute("""
        SELECT ih.*, s.name FROM install_history ih
        LEFT JOIN sales s ON s.thread_id=ih.thread_id
        ORDER BY ih.changed_at DESC LIMIT 10
    """).fetchall()
    stats["recent_changes"] = [dict(r) for r in changes]
    db.close()
    return jsonify(stats)

@app.route("/api/months")
def api_months():
    db = get_db()
    rows = db.execute("""
        SELECT DISTINCT substr(install_date,1,7) as m
        FROM sales WHERE install_date != ''
        ORDER BY m DESC
    """).fetchall()
    db.close()
    return jsonify([r["m"] for r in rows])

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

if __name__ == "__main__":
    init_db()
    app.run(debug=True, port=5000)
