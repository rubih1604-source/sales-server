import os, json, base64, re
from flask import Flask, request, jsonify, redirect, Response
from flask_cors import CORS
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET', 'sales-rubi-2026')
CORS(app, origins="*")

CLIENT_ID     = os.environ.get('GOOGLE_CLIENT_ID')
CLIENT_SECRET = os.environ.get('GOOGLE_CLIENT_SECRET')
REDIRECT_URI  = 'https://sales-server-egdf.onrender.com/oauth/callback'
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly',
          'https://www.googleapis.com/auth/gmail.send',
          'https://www.googleapis.com/auth/gmail.compose']

# ── שמירת טוקן — ENV VAR ראשון, /tmp גיבוי ──────────────
# הטוקן נשמר ב-GMAIL_TOKEN env var כדי שלא יימחק עם /tmp
TOKEN_FILE = '/tmp/gmail_token.json'
TOKEN_ENV  = 'GMAIL_TOKEN'

def load_token_data():
    # נסה env var קודם (עמיד לאיפוס)
    env_val = os.environ.get(TOKEN_ENV)
    if env_val:
        try: return json.loads(env_val)
        except: pass
    # גיבוי: קובץ
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE) as f:
            return json.load(f)
    return None

def save_token_data(token, refresh_token):
    data = {'token': token, 'refresh_token': refresh_token}
    # שמור לקובץ תמיד
    with open(TOKEN_FILE, 'w') as f:
        json.dump(data, f)
    # נסה לעדכן env var (עובד רק אם יש RENDER_API_KEY)
    _try_update_render_env(data)

def _try_update_render_env(data):
    """עדכן את ה-env var ב-Render דרך ה-API שלהם"""
    try:
        api_key = os.environ.get('RENDER_API_KEY')
        service_id = os.environ.get('RENDER_SERVICE_ID')
        if not api_key or not service_id:
            return
        import urllib.request
        payload = json.dumps({'key': TOKEN_ENV, 'value': json.dumps(data)}).encode()
        req = urllib.request.Request(
            f'https://api.render.com/v1/services/{service_id}/env-vars',
            data=payload,
            headers={'Authorization': f'Bearer {api_key}',
                     'Content-Type': 'application/json'},
            method='PUT'
        )
        urllib.request.urlopen(req, timeout=5)
    except: pass  # לא קריטי

def get_credentials():
    data = load_token_data()
    if not data: return None
    return Credentials(
        token=data.get('token'),
        refresh_token=data.get('refresh_token'),
        token_uri="https://oauth2.googleapis.com/token",
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET,
        scopes=SCOPES
    )

def get_client_config():
    return {"web": {"client_id": CLIENT_ID, "client_secret": CLIENT_SECRET,
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "redirect_uris": [REDIRECT_URI]}}

def get_service():
    creds = get_credentials()
    if not creds: return None
    try:
        svc = build('gmail', 'v1', credentials=creds)
        # רענן טוקן אם פג
        if creds.expired and creds.refresh_token:
            from google.auth.transport.requests import Request
            creds.refresh(Request())
            save_token_data(creds.token, creds.refresh_token)
        return svc
    except: return None

HOURS_RE = r'(\d{1,2}(?::\d{2})?[-–]\d{1,2}(?::\d{2})?)'
TODAY_KW = sorted([
    'ירד להיום','ירדה להיום','ההתקנה להיום','התקנה להיום',
    'שובץ להיום','תואם להיום','מוקלד להיום','להיום','היום'
], key=len, reverse=True)

@app.route('/')
def index():
    connected = get_credentials() is not None
    st = "✅ מחובר ל-Gmail" if connected else "❌ לא מחובר"
    link = "" if connected else "<br><br><a href='/oauth/start' style='background:#3b82f6;color:white;padding:14px 28px;border-radius:8px;text-decoration:none;font-size:16px;font-weight:bold'>🔐 התחבר ל-Gmail</a>"
    return f"<html dir='rtl'><body style='font-family:Arial;padding:40px;background:#0d1117;color:#e6edf3'><h1>שרת מכירות רובי</h1><p style='font-size:18px'>{st}</p>{link}</body></html>"

@app.route('/oauth/start')
def oauth_start():
    flow = Flow.from_client_config(get_client_config(), scopes=SCOPES, redirect_uri=REDIRECT_URI)
    auth_url, state = flow.authorization_url(access_type='offline', prompt='consent')
    with open('/tmp/oauth_state.txt', 'w') as f: f.write(state)
    return redirect(auth_url)

@app.route('/oauth/callback')
def oauth_callback():
    flow = Flow.from_client_config(get_client_config(), scopes=SCOPES, redirect_uri=REDIRECT_URI)
    flow.fetch_token(authorization_response=request.url)
    save_token_data(flow.credentials.token, flow.credentials.refresh_token)
    return "<html dir='rtl'><body style='background:#0d1117;color:#e6edf3;padding:40px;font-family:Arial;text-align:center'><h1>✅ התחברות הצליחה!</h1><p style='color:#10b981;font-size:18px'>המערכת מחוברת ל-Gmail.<br>סגור חלון זה וחזור למערכת.</p></body></html>"

@app.route('/api/status')
def api_status():
    return jsonify({'connected': get_credentials() is not None})

def parse_email_date(date_str):
    try:
        from email.utils import parsedate
        p = parsedate(date_str)
        if p:
            dt = datetime(p[0], p[1], p[2])
            return f'{p[2]:02d}/{p[1]:02d}/{p[0]}', f'{p[2]:02d}/{p[1]:02d}', dt
    except: pass
    return '', '', None

def ddmm_offset(ddmm, days):
    try:
        parts = ddmm.split('/')
        dt = datetime(2026, int(parts[1]), int(parts[0])) + timedelta(days=days)
        return f'{dt.day:02d}/{dt.month:02d}'
    except: return ''

def normalize_hours(h):
    h = h.replace('–', '-').replace(':00', '')
    parts = h.split('-')
    if len(parts) == 2:
        try: return f'{int(parts[0])}-{int(parts[1])}'
        except: pass
    return h

def normalize_hours(h):
    """8-10, 08:00-10:00, 1000-1200 → 8-10"""
    h = str(h).replace('–','-').strip()
    h = re.sub(r':00', '', h)
    h = re.sub(r'\b0?(\d{1,2})00\b', r'\1', h)
    parts = h.split('-')
    if len(parts) == 2:
        try: return f'{int(parts[0])}-{int(parts[1])}'
        except: pass
    return h

def normalize_date(d):
    """14/6 → 14/06"""
    parts = str(d).strip().split('/')
    return f'{parts[0].zfill(2)}/{parts[1].zfill(2)}' if len(parts)==2 else str(d)

def ddmm_offset(ddmm, days):
    try:
        p = ddmm.split('/')
        dt = datetime(2026, int(p[1]), int(p[0])) + timedelta(days=days)
        return f'{dt.day:02d}/{dt.month:02d}'
    except: return ''

STATUS_KW = r'(?:מוקלד|תואם|מתואם|שובץ|אושר|אישר|נקבע|מתוזמן|שיבוץ|תואמה)'
BANK_KW   = ['בנק התקנה','בהקדם','ממתין','אין תאריך','לא נקבע']

def find_updates_in_snippet(snippet, msg_ddmm, msg_dt):
    """מנוע סריקה מקיף — מזהה כל וריאציה של תיאום התקנה."""
    results = []
    text = snippet or ''
    today_pat = '|'.join(re.escape(k) for k in TODAY_KW)

    # 1. היום/ירד להיום + שעות
    for m in re.finditer(r'(?:' + today_pat + r')\s*' + HOURS_RE, text, re.IGNORECASE):
        results.append({'date': msg_ddmm, 'hours': normalize_hours(m.group(1)), 'kind': 'today', 'dt': msg_dt})
    for m in re.finditer(HOURS_RE + r'\s+(?:' + today_pat + r')', text, re.IGNORECASE):
        results.append({'date': msg_ddmm, 'hours': normalize_hours(m.group(1)), 'kind': 'today', 'dt': msg_dt})
    for m in re.finditer(HOURS_RE + r'\s+לעדכן\s+לקוח', text):
        results.append({'date': msg_ddmm, 'hours': normalize_hours(m.group(1)), 'kind': 'today', 'dt': msg_dt})

    # 2. מחר
    for m in re.finditer(r'(?:התקנה\s+ל)?מחר\s+' + HOURS_RE, text):
        tom = ddmm_offset(msg_ddmm, 1)
        if tom: results.append({'date': tom, 'hours': normalize_hours(m.group(1)), 'kind': 'tomorrow', 'dt': msg_dt})

    # 3. נרמול
    n = text
    n = re.sub(r'ב(\d{1,2}/\d{1,2})', r'\1', n)
    n = re.sub(r'ל(\d{1,2}/\d{1,2})', r'\1', n)
    n = re.sub(r'בין\s+', '', n)
    n = re.sub(r'מס\.?\s*לקוח', 'לקוח', n)
    n = re.sub(r'מתואם', 'תואם', n)
    n = re.sub(r'שיבוץ', 'שובץ', n)

    # 4. STATUS DD/MM HH-HH
    for m in re.finditer(STATUS_KW + r'\s+(\d{1,2}/\d{1,2})\s+' + HOURS_RE, n):
        results.append({'date': normalize_date(m.group(1)), 'hours': normalize_hours(m.group(2)), 'kind': 'recorded', 'dt': msg_dt})

    # 5. STATUS HH-HH DD/MM
    for m in re.finditer(STATUS_KW + r'\s+' + HOURS_RE + r'\s+(\d{1,2}/\d{1,2})', n):
        results.append({'date': normalize_date(m.group(2)), 'hours': normalize_hours(m.group(1)), 'kind': 'recorded', 'dt': msg_dt})

    # 6. לקוח XXXXXXX [STATUS] DD/MM HH-HH
    for m in re.finditer(r'לקוח\s+\d+\s+(?:' + STATUS_KW[3:-1] + r'\s+)?(\d{1,2}/\d{1,2})\s+' + HOURS_RE, n):
        results.append({'date': normalize_date(m.group(1)), 'hours': normalize_hours(m.group(2)), 'kind': 'recorded', 'dt': msg_dt})

    # 7. לקוח XXXXXXX HH-HH DD/MM
    for m in re.finditer(r'לקוח\s+\d+\s+' + HOURS_RE + r'\s+(\d{1,2}/\d{1,2})', n):
        results.append({'date': normalize_date(m.group(2)), 'hours': normalize_hours(m.group(1)), 'kind': 'recorded', 'dt': msg_dt})

    # 8. DD/MM HH-HH גרידא
    for m in re.finditer(r'(\d{1,2}/\d{1,2})\s+' + HOURS_RE, n):
        results.append({'date': normalize_date(m.group(1)), 'hours': normalize_hours(m.group(2)), 'kind': 'explicit', 'dt': msg_dt})

    # 9. HH-HH DD/MM גרידא
    for m in re.finditer(HOURS_RE + r'\s+(\d{1,2}/\d{1,2})', n):
        results.append({'date': normalize_date(m.group(2)), 'hours': normalize_hours(m.group(1)), 'kind': 'explicit', 'dt': msg_dt})

    # 10. בנק התקנה
    if not results and any(k in text for k in BANK_KW):
        results.append({'date': 'ממתין', 'hours': '', 'kind': 'waiting', 'dt': msg_dt})

    return results


def find_attachments(payload, attachments, message_id):
    fn = payload.get('filename', '')
    if fn and fn.lower().endswith('.pdf'):
        att_id = payload.get('body', {}).get('attachmentId')
        if att_id: attachments.append({'filename': fn, 'attachmentId': att_id, 'messageId': message_id})
    for part in payload.get('parts', []): find_attachments(part, attachments, message_id)

def best_date_in_message(text, msg_ddmm, msg_dt):
    """מחזיר (תאריך, שעות, עדיפות) הכי אמין בהודעה אחת, או None."""
    candidates = []
    n = text
    n = re.sub(r'ב(\d{1,2}/\d{1,2})', r'\1', n)
    n = re.sub(r'ל(\d{1,2}/\d{1,2})', r'\1', n)
    n = re.sub(r'בין\s+', '', n)
    n = re.sub(r'מתואם', 'תואם', n)
    n = re.sub(r'שיבוץ', 'שובץ', n)
    today_pat = '|'.join(re.escape(k) for k in TODAY_KW)

    # עדיפות 4: STATUS + תאריך + שעות (תואם 3/06 14-16)
    for m in re.finditer(r'(?:מוקלד|תואם|שובץ|אושר|אישר|נקבע)\s+(\d{1,2}/\d{1,2})\s+' + HOURS_RE, n):
        candidates.append((normalize_date(m.group(1)), normalize_hours(m.group(2)), 4))
    # עדיפות 4: STATUS + שעות + תאריך (שובץ 4-6 14/6)
    for m in re.finditer(r'(?:מוקלד|תואם|שובץ|אושר|אישר|נקבע)\s+' + HOURS_RE + r'\s+(\d{1,2}/\d{1,2})', n):
        candidates.append((normalize_date(m.group(2)), normalize_hours(m.group(1)), 4))
    # עדיפות 3: היום/ירד להיום + שעות
    for m in re.finditer(r'(?:' + today_pat + r')\s*' + HOURS_RE, text, re.IGNORECASE):
        candidates.append((msg_ddmm, normalize_hours(m.group(1)), 3))
    for m in re.finditer(HOURS_RE + r'\s+(?:' + today_pat + r')', text, re.IGNORECASE):
        candidates.append((msg_ddmm, normalize_hours(m.group(1)), 3))
    # עדיפות 3: מחר + שעות
    for m in re.finditer(r'(?:התקנה\s+ל)?מחר\s+' + HOURS_RE, text):
        tom = ddmm_offset(msg_ddmm, 1)
        if tom: candidates.append((tom, normalize_hours(m.group(1)), 3))
    # עדיפות 2: תאריך + שעות גרידא
    for m in re.finditer(r'(\d{1,2}/\d{1,2})\s+' + HOURS_RE, n):
        candidates.append((normalize_date(m.group(1)), normalize_hours(m.group(2)), 2))
    for m in re.finditer(HOURS_RE + r'\s+(\d{1,2}/\d{1,2})', n):
        candidates.append((normalize_date(m.group(2)), normalize_hours(m.group(1)), 2))
    # עדיפות 1: תאריך לבד (צריך להזיז ל1.6)
    for m in re.finditer(r'(\d{1,2}/\d{1,2})', n):
        d = normalize_date(m.group(1))
        candidates.append((d, '', 1))

    if not candidates:
        return None
    candidates.sort(key=lambda x: x[2], reverse=True)
    return candidates[0]


def parse_sale_from_thread(messages, thread_id):
    full_text = ''
    first_subject = ''
    all_msgs = []

    for i, m in enumerate(messages):
        hdrs = {h['name']: h['value'] for h in m['payload']['headers']}
        subj = hdrs.get('Subject', '')
        date_str = hdrs.get('Date', '')
        date_full, date_ddmm, msg_dt = parse_email_date(date_str)
        if i == 0: first_subject = subj
        snippet = m.get('snippet', '')
        full_text += subj + '\n' + snippet + '\n'
        all_msgs.append({'i': i, 'snippet': snippet, 'date_full': date_full,
                         'date_ddmm': date_ddmm, 'dt': msg_dt})

    # זהה thread מכירה — מילת מפתח או מספר לקוח
    sale_kw = ['תואם','לאשר חוזה','ממירים','שובץ','להקים','ממיר','מוקלד',
               'אישר','אושר','הוק','לעדכן לקוח','ירד להיום','ירדה להיום','ההתקנה להיום',
               'דאבל יס','דאבל סטינג','סיבים']
    has_cid = bool(re.search(r'\b30\d{5}\b', full_text))
    if not any(k in full_text for k in sale_kw) and not has_cid:
        return None

    # שם מהנושא
    subj_clean = re.sub(r'^(Re|Fwd|FW|RE):\s*', '', first_subject, flags=re.IGNORECASE).strip()
    subj_clean = re.sub(r'[-–].*', '', subj_clean).strip()
    if not subj_clean or len(subj_clean) < 2: return None

    # ביטול
    cancel_phrases = ['לא להקים','התחרטה','לבטל','ביטול','מבטל','בוטל','ביטל','מבטלת','ביטלה']
    is_cancelled = any(p in full_text for p in cancel_phrases)
    cancel_note = next((p for p in ['התחרטה','לא להקים','ביטל','ביטלה'] if p in full_text), '')

    # מספר לקוח
    cid = ''
    cid_m = re.search(r'(?:לקוח|מס\s*לקוח)\s+(\d{7})', full_text)
    if cid_m:
        cid = cid_m.group(1)
    else:
        any_cid = re.search(r'\b(30\d{5})\b', full_text)
        if any_cid: cid = any_cid.group(1)

    # ── לב המנוע: עבור על המיילים לפי סדר זמן, קח תאריך מכל אחד ──
    # התאריך מההודעה האחרונה (הכי מאוחרת) שמכילה תאריך = הסופי
    install_date, install_hours = '', ''
    first_date, first_hours = '', ''
    sorted_msgs = sorted(all_msgs, key=lambda x: x['dt'] or datetime.min)
    for msg in sorted_msgs:
        b = best_date_in_message(msg['snippet'], msg['date_ddmm'], msg['dt'])
        if b:
            if not first_date:
                first_date, first_hours = b[0], b[1]
            install_date = b[0]
            if b[1]: install_hours = b[1]  # שמור שעות אחרונות שאינן ריקות

    # שינוי תאריך
    has_change = bool(first_date and install_date and first_date != install_date)
    change_note = f'שונה מ-{first_date} ל-{install_date}' if has_change else ''

    # סטטוס
    if is_cancelled:
        status = 'בוטל'
    elif any(w in full_text for w in ['מוקלד','אושר','אישר','לעדכן לקוח']):
        status = 'מוקלד + הוק' if 'הוק' in full_text else 'מוקלד'
    elif 'לאשר חוזה' in full_text and 'הוק' in full_text:
        status = 'לאשר חוזה + הוק'
    elif 'לאשר חוזה' in full_text:
        status = 'לאשר חוזה'
    elif 'שובץ' in full_text or install_date:
        status = 'שובץ'
    else:
        status = 'בהקדם'

    # ממירים
    mirrors = 0
    mm = re.search(r'(\d+)\s*ממירים?', full_text)
    if mm:
        mirrors = int(mm.group(1))
    elif 'טריפל' in full_text: mirrors = 3
    elif 'דאבל' in full_text: mirrors = 2

    # חודש לפי תאריך התקנה
    install_month = ''
    if install_date and install_date not in ('—', 'ממתין', ''):
        parts = install_date.split('/')
        if len(parts) == 2:
            try: install_month = f'{int(parts[1]):02d}/2026'
            except: pass

    sale_date = sorted_msgs[0]['date_full'] if sorted_msgs else ''
    if not install_month and sale_date:
        sp = sale_date.split('/')
        if len(sp) >= 2:
            try: install_month = f'{int(sp[1]):02d}/2026'
            except: pass

    if not install_date: install_date = 'ממתין'

    return {'name': subj_clean, 'customerId': cid, 'saleDate': sale_date,
            'installDate': install_date, 'installMonth': install_month,
            'hours': install_hours, 'mirrors': mirrors, 'status': status,
            'hasChange': has_change, 'changeNote': change_note,
            'isCancelled': is_cancelled, 'cancelNote': cancel_note,
            'isToday': False, 'isApril': False, 'threadId': thread_id}


@app.route('/api/scan')
def scan():
    service = get_service()
    if not service:
        return jsonify({'error': 'not_authenticated',
                       'auth_url': 'https://sales-server-egdf.onrender.com/oauth/start'}), 401
    try:
        sales, invoices = [], []
        results = service.users().messages().list(userId='me',
            q='(from:oshrityes2901@gmail.com OR to:oshrityes2901@gmail.com OR '
              'from:oritapiro22@gmail.com OR to:oritapiro22@gmail.com OR '
              'from:avielv014@gmail.com OR to:avielv014@gmail.com) after:2026/1/1',
            maxResults=300).execute()

        threads_seen = set()
        for msg in results.get('messages', []):
            tid = msg['threadId']
            if tid in threads_seen: continue
            threads_seen.add(tid)
            try:
                thread = service.users().threads().get(userId='me', id=tid, format='full').execute()
                sale = parse_sale_from_thread(thread.get('messages', []), tid)
                if sale: sales.append(sale)
            except: continue

        inv_results = service.users().messages().list(userId='me',
            q='(חשבונית OR invoice OR receipt OR morning.co OR render.com OR icount) '
              'has:attachment after:2026/1/1', maxResults=50).execute()
        for msg in inv_results.get('messages', []):
            try:
                md = service.users().messages().get(userId='me', id=msg['id'], format='full').execute()
                hdrs = {h['name']: h['value'] for h in md['payload']['headers']}
                subj = hdrs.get('Subject', '')
                sender = hdrs.get('From', '')
                date_str = hdrs.get('Date', '')
                date_full, _, _ = parse_email_date(date_str)
                month = f'{date_full[3:5]}/{date_full[6:10]}' if len(date_full) >= 10 else ''
                biz = ['ר.א.מ','ליד מנג','render','funnelly','stripe','morning','חשבונית ירוקה','atp','icount','קבלה']
                if not any(k.lower() in (subj+sender).lower() for k in biz): continue
                atts = []
                find_attachments(md['payload'], atts, msg['id'])
                inv_type = 'income' if any(k in (subj+sender) for k in ['ר.א.מ','עמלות','may zalah','May zalah']) \
                    else ('subscription' if any(k in (subj+sender) for k in ['חשבונית ירוקה','morning']) else 'expense')
                num_m = re.search(r'(\d{4,})', subj)
                name_m = re.match(r'^"?([^"<]+)', sender)
                invoices.append({'id': msg['id'], 'date': date_full, 'month': month,
                    'from': name_m.group(1).strip() if name_m else sender,
                    'subject': subj, 'description': subj[:50],
                    'invoiceNum': num_m.group(1) if num_m else '', 'type': inv_type,
                    'gmailLink': f'https://mail.google.com/mail/u/0/#all/{msg["id"]}',
                    'hasAttachment': len(atts) > 0, 'attachments': atts})
            except: continue

        def d2i(d):
            try: p=d.split('/'); return int(p[2])*10000+int(p[1])*100+int(p[0])
            except: return 0
        sales.sort(key=lambda x: d2i(x.get('saleDate','')), reverse=True)
        return jsonify({'success': True, 'sales': sales, 'invoices': invoices})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/attachment/<message_id>/<attachment_id>')
def get_attachment(message_id, attachment_id):
    service = get_service()
    if not service: return jsonify({'error': 'not_authenticated'}), 401
    try:
        att = service.users().messages().attachments().get(
            userId='me', messageId=message_id, id=attachment_id).execute()
        data = base64.urlsafe_b64decode(att['data'])
        return Response(data, mimetype='application/pdf',
                       headers={'Content-Disposition': 'attachment; filename=invoice.pdf',
                                'Access-Control-Allow-Origin': '*'})
    except Exception as e: return jsonify({'error': str(e)}), 500

@app.route('/api/send-to-accountant', methods=['POST'])
def send_to_accountant():
    service = get_service()
    if not service: return jsonify({'error': 'not_authenticated'}), 401
    data = request.json
    month_name = data.get('monthName', '')
    invoices = data.get('invoices', [])
    try:
        msg = MIMEMultipart()
        msg['To'] = 'ei@eicpa.co.il'
        msg['Subject'] = f'חשבוניות {month_name} — ראובן חגג'
        body = f'שלום,\n\nמצורפות חשבוניות לחודש {month_name}.\n\n'
        for inv in invoices: body += f'• {inv.get("from","")} — {inv.get("description","")}\n'
        body += '\nבברכה,\nרובי חגג'
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        for inv in invoices:
            for att in inv.get('attachments', []):
                if att.get('attachmentId'):
                    try:
                        ad = service.users().messages().attachments().get(
                            userId='me', messageId=att['messageId'], id=att['attachmentId']).execute()
                        pdf = base64.urlsafe_b64decode(ad['data'])
                        part = MIMEBase('application', 'pdf')
                        part.set_payload(pdf)
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename="{att["filename"]}"')
                        msg.attach(part)
                    except: pass
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(userId='me', body={'raw': raw}).execute()
        return jsonify({'success': True})
    except Exception as e: return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=False)
