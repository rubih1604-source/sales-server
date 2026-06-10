import os
import json
import re
import sqlite3
import anthropic
from datetime import datetime, date, timedelta
from flask import Flask, redirect, request, session, jsonify, render_template
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", os.environ.get("SECRET_KEY", "dev-secret"))

GOOGLE_CLIENT_ID     = os.environ.get("GOOGLE_CLIENT_ID")
GOOGLE_CLIENT_SECRET = os.environ.get("GOOGLE_CLIENT_SECRET")
ANTHROPIC_API_KEY    = os.environ.get("ANTHROPIC_API_KEY")
REDIRECT_URI         = os.environ.get("REDIRECT_URI", "http://localhost:5000/oauth/callback")
DB_PATH              = os.environ.get("DB_PATH", "sales.db")

SCOPES = [
    "openid",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/gmail.readonly",
]

OPS_SENDERS = [
    "oshrityes2901@gmail.com",
    "oritapiro22@gmail.com",
    "avielv014@gmail.com",
]

def get_db():
    db = sqlite3.connect(DB_PATH)
    db.row_factory = sqlite3.Row
    return db

def init_db():
    db = get_db()
    db.executescript("""
    CREATE TABLE IF NOT EXISTS sales (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        thread_id     TEXT UNIQUE NOT NULL,
        name          TEXT,
        customer_id   TEXT,
        sale_date     TEXT,
        install_date  TEXT,
        install_hours TEXT,
        converters    INTEGER,
        package       TEXT,
        status        TEXT,
        contract_ok   INTEGER DEFAULT 0,
        notes         TEXT,
        raw_snippet   TEXT,
        last_scanned  TEXT,
        updated_at    TEXT
    );
    CREATE TABLE IF NOT EXISTS scan_log (
        id         INTEGER PRIMARY KEY AUTOINCREMENT,
        scanned_at TEXT,
        total      INTEGER,
        new_rows   INTEGER,
        updated    INTEGER,
        month_ctx  TEXT
    );
    CREATE TABLE IF NOT EXISTS install_history (
        id           INTEGER PRIMARY KEY AUTOINCREMENT,
        thread_id    TEXT,
        old_date     TEXT,
        new_date     TEXT,
        changed_at   TEXT,
        reason       TEXT
    );
    """)
    db.commit()
    db.close()

def make_flow():
    return Flow.from_client_config(
        {
            "web": {
                "client_id":     GOOGLE_CLIENT_ID,
                "client_secret": GOOGLE_CLIENT_SECRET,
                "auth_uri":      "https://accounts.google.com/o/oauth2/auth",
                "token_uri":     "https://oauth2.googleapis.com/token",
            }
        },
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )

@app.route("/login")
def login():
    flow = make_flow()
    url, state = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    session["oauth_state"] = state
    return redirect(url)

@app.route("/oauth/callback")
def oauth_callback():
    flow = make_flow()
    flow.fetch_token(authorization_response=request.url)
    creds = flow.credentials
    session["token"]         = creds.token
    session["refresh_token"] = creds.refresh_token
    session["token_uri"]     = creds.token_uri
    session["client_id"]     = creds.client_id
    session["client_secret"] = creds.client_secret
    return redirect("/")

def get_gmail():
    if "token" not in session:
        return None
    creds = Credentials(
        token         = session["token"],
        refresh_token = session.get("refresh_token"),
        token_uri     = session.get("token_uri"),
        client_id     = session.get("client_id"),
        client_secret = session.get("client_secret"),
    )
    return build("gmail", "v1", credentials=creds, cache_discovery=False)

def fetch_threads(gmail, days_back=90, page_token=None):
    after = (date.today() - timedelta(days=days_back)).strftime("%Y/%m/%d")
    query = (
        f"after:{after} "
        "({from:oshrityes2901@gmail.com OR from:oritapiro22@gmail.com OR from:avielv014@gmail.com} "
        "לקוח) OR "
        f"(after:{after} from:rubih1604@gmail.com to:oshrityes2901@gmail.com)"
    )
    params = {"userId": "me", "q": query, "maxResults": 50}
    if page_token:
        params["pageToken"] = page_token

    resp = gmail.users().threads().list(**params).execute()
    threads = resp.get("threads", [])
    next_page = resp.get("nextPageToken")

    results = []
    for t in threads:
        full = gmail.users().threads().get(
            userId="me", id=t["id"], format="metadata",
            metadataHeaders=["Subject", "From", "Date"]
        ).execute()
        msgs = full.get("messages", [])
        if not msgs:
            continue

        subject = ""
        sale_date_raw = ""
        for h in msgs[0].get("payload", {}).get("headers", []):
            if h["name"] == "Subject":
                subject = h["value"].replace("Re: ", "").replace("Fwd: ", "").strip()
            if h["name"] == "Date":
                sale_date_raw = h["value"]

        all_snippets = []
        latest_ops_snippet = ""
        for msg in msgs:
            snip = msg.get("snippet", "")
            all_snippets.append(snip)
            sender = ""
            for h in msg.get("payload", {}).get("headers", []):
                if h["name"] == "From":
                    sender = h["value"]
            if any(op in sender for op in OPS_SENDERS):
                latest_ops_snippet = snip

        results.append({
            "thread_id":          t["id"],
            "subject":            subject,
            "sale_date_raw":      sale_date_raw,
            "latest_ops_snippet": latest_ops_snippet,
            "all_snippets":       " ||| ".join(all_snippets[-6:]),
        })

    return results, next_page

def analyze_batch(threads_data, current_month):
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    batch_text = "\n\n".join([
        f"THREAD_ID: {t['thread_id']}\n"
        f"SUBJECT: {t['subject']}\n"
        f"SALE_DATE_RAW: {t['sale_date_raw']}\n"
        f"LATEST_OPS_MSG: {t['latest_ops_snippet']}\n"
        f"ALL_SNIPPETS: {t['all_snippets']}"
        for t in threads_data
    ])
    today_str = date.today().isoformat()
    prompt = f"""You are analyzing Israeli YES telecom sales email threads for agent Rubi.
Today: {today_str}. Current month: {current_month}.

RULES:
1. Only extract real sales threads (customer name as subject + installation).
2. LATEST_OPS_MSG is ground truth.
3. installDate: from "תואם DD/MM" or "שובץ DD/MM" - use MOST RECENT if rescheduled. Year 2026.
4. contractApproved: true if "אישר"/"אישרה"/"מוקלד" anywhere.
5. status: "cancelled"→ביטול/נשלח לביטול/לא ניתן להקים | "recorded"→מוקלד | "approved"→אישר/אישרה | "sent"→דוור/שיבוץ | "waiting"→לאשר חוזה | null→not a sale
6. converters: integer before ממירים
7. package: דאבל יס פלוס/דאבל יס/רק יס/דאבל סטינג/דאבל סיבים/מסך בלבד/other
8. notes: short Hebrew note if unusual

THREADS:
{batch_text}

Return ONLY JSON array, null for non-sales:
[{{"thread_id":"...","name":"...","customer_id":"...","sale_date":"YYYY-MM-DD","install_date":"YYYY-MM-DD","install_hours":"8-10","converters":2,"package":"...","status":"waiting","contract_ok":false,"notes":"..."}}]"""

    resp = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1000,
        messages=[{"role": "user", "content": prompt}]
    )
    raw = resp.content[0].text.strip()
    raw = re.sub(r"^```json\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    try:
        return [r for r in json.loads(raw) if r]
    except:
        m = re.search(r"\[[\s\S]*\]", raw)
        if m:
            return [r for r in json.loads(m.group()) if r]
        return []

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
            db.execute("""
                INSERT INTO sales
                (thread_id,name,customer_id,sale_date,install_date,install_hours,
                 converters,package,status,contract_ok,notes,raw_snippet,last_scanned,updated_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                tid, row.get("name",""), row.get("customer_id",""),
                row.get("sale_date",""), install_date, row.get("install_hours",""),
                row.get("converters"), row.get("package",""),
                row.get("status","waiting"), 1 if row.get("contract_ok") else 0,
                row.get("notes",""), "", now, now,
            ))
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
            new_status = row.get("status", existing["status"])
            if new_status != existing["status"]:
                changes["status"] = new_status
            new_contract = 1 if row.get("contract_ok") else 0
            if new_contract != existing["contract_ok"]:
                changes["contract_ok"] = new_contract
            new_notes = row.get("notes","")
            if new_notes and new_notes != existing["notes"]:
                changes["notes"] = new_notes
            if changes:
                changes["updated_at"] = now
                changes["last_scanned"] = now
                set_clause = ", ".join(f"{k}=?" for k in changes)
                db.execute(f"UPDATE sales SET {set_clause} WHERE thread_id=?", (*changes.values(), tid))
                stats["updated"] += 1
            else:
                db.execute("UPDATE sales SET last_scanned=? WHERE thread_id=?", (now, tid))
                stats["unchanged"] += 1

    db.execute("INSERT INTO scan_log (scanned_at,total,new_rows,updated,month_ctx) VALUES (?,?,?,?,?)",
               (now, len(parsed), stats["new"], stats["updated"], current_month))
    db.commit()
    db.close()
    return stats

@app.route("/")
def index():
    if "token" not in session:
        return render_template("index.html", logged_in=False)
    return render_template("index.html", logged_in=True)

@app.route("/api/scan", methods=["POST"])
def api_scan():
    gmail = get_gmail()
    if not gmail:
        return jsonify({"error": "not_logged_in"}), 401
    today = date.today()
    current_month = today.strftime("%Y-%m")
    body = request.get_json() or {}
    days_back = 180 if body.get("full") else 70
    all_parsed = []
    page_token = None
    page = 0
    while True:
        threads, next_page = fetch_threads(gmail, days_back=days_back, page_token=page_token)
        page += 1
        if not threads:
            break
        for i in range(0, len(threads), 15):
            parsed = analyze_batch(threads[i:i+15], current_month)
            all_parsed.extend(parsed)
        page_token = next_page
        if not next_page or page >= 4:
            break
    stats = sync_to_db(all_parsed, current_month)
    db = get_db()
    total = db.execute("SELECT COUNT(*) FROM sales").fetchone()[0]
    db.close()
    return jsonify({"ok":True,"scanned":len(all_parsed),"new":stats["new"],"updated":stats["updated"],"unchanged":stats["unchanged"],"total_in_db":total,"scan_month":current_month})

@app.route("/api/sales")
def api_sales():
    db = get_db()
    month  = request.args.get("month","")
    status = request.args.get("status","")
    search = request.args.get("q","")
    q = "SELECT * FROM sales WHERE 1=1"
    params = []
    if month:  q += " AND install_date LIKE ?";  params.append(f"{month}%")
    if status: q += " AND status=?";              params.append(status)
    if search: q += " AND (name LIKE ? OR customer_id LIKE ?)"; params += [f"%{search}%",f"%{search}%"]
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
        "total": count("1=1"), "this_month": count("install_date LIKE ?", [f"{this_month}%"]),
        "last_month": count("install_date LIKE ?", [f"{last_month}%"]),
        "waiting": count("status='waiting'"), "approved": count("status='approved'"),
        "recorded": count("status='recorded'"), "cancelled": count("status='cancelled'"),
        "no_contract": count("contract_ok=0 AND status NOT IN ('cancelled','unknown')"),
        "current_month": this_month, "last_month": last_month,
    }
    months = db.execute("SELECT substr(install_date,1,7) as m, COUNT(*) as cnt FROM sales WHERE install_date != '' GROUP BY m ORDER BY m DESC LIMIT 6").fetchall()
    stats["by_month"] = [dict(r) for r in months]
    last_scan = db.execute("SELECT scanned_at,total,new_rows,updated FROM scan_log ORDER BY id DESC LIMIT 1").fetchone()
    stats["last_scan"] = dict(last_scan) if last_scan else None
    changes = db.execute("SELECT ih.*, s.name FROM install_history ih LEFT JOIN sales s ON s.thread_id=ih.thread_id ORDER BY ih.changed_at DESC LIMIT 10").fetchall()
    stats["recent_changes"] = [dict(r) for r in changes]
    db.close()
    return jsonify(stats)

@app.route("/api/months")
def api_months():
    db = get_db()
    rows = db.execute("SELECT DISTINCT substr(install_date,1,7) as m FROM sales WHERE install_date != '' ORDER BY m DESC").fetchall()
    db.close()
    return jsonify([r["m"] for r in rows])

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

if __name__ == "__main__":
    os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"
    init_db()
    app.run(debug=True, port=5000)
