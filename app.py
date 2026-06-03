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

def find_updates_in_snippet(snippet, msg_ddmm, msg_dt):
    results = []
    text = snippet
    today_pat = '|'.join(re.escape(k) for k in TODAY_KW)

    # "היום/ירד להיום + שעות"
    for m in re.finditer(r'(?:' + today_pat + r')\s*' + HOURS_RE, text, re.IGNORECASE):
        results.append({'date': msg_ddmm, 'hours': normalize_hours(m.group(1)), 'kind': 'today', 'dt': msg_dt})
    # "שעות + להיום"
    for m in re.finditer(HOURS_RE + r'\s+(?:' + today_pat + r')', text, re.IGNORECASE):
        results.append({'date': msg_ddmm, 'hours': normalize_hours(m.group(1)), 'kind': 'today', 'dt': msg_dt})
    # "שעות לעדכן לקוח" (בלי תאריך = היום)
    for m in re.finditer(HOURS_RE + r'\s+לעדכן\s+לקוח', text):
        results.append({'date': msg_ddmm, 'hours': normalize_hours(m.group(1)), 'kind': 'today', 'dt': msg_dt})
    # "מחר שעות"
    for m in re.finditer(r'מחר\s+' + HOURS_RE, text):
        tom = ddmm_offset(msg_ddmm, 1)
        if tom: results.append({'date': tom, 'hours': normalize_hours(m.group(1)), 'kind': 'tomorrow', 'dt': msg_dt})
    # "מוקלד/תואם/שובץ DD/MM שעות"
    for m in re.finditer(r'(?:מוקלד|תואם|שובץ|אושר|אישר|נקבע|לקוח\s+\d+)\s+(\d{1,2}/\d{2})\s+' + HOURS_RE, text):
        results.append({'date': m.group(1), 'hours': normalize_hours(m.group(2)), 'kind': 'recorded', 'dt': msg_dt})
    # "DD/MM שעות"
    for m in re.finditer(r'(\d{1,2}/\d{2})\s+' + HOURS_RE, text):
        results.append({'date': m.group(1), 'hours': normalize_hours(m.group(2)), 'kind': 'explicit', 'dt': msg_dt})
    return results

def find_attachments(payload, attachments, message_id):
    fn = payload.get('filename', '')
    if fn and fn.lower().endswith('.pdf'):
        att_id = payload.get('body', {}).get('attachmentId')
        if att_id: attachments.append({'filename': fn, 'attachmentId': att_id, 'messageId': message_id})
    for part in payload.get('parts', []): find_attachments(part, attachments, message_id)

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

    sale_kw = ['תואם','לאשר חוזה','ממירים','שובץ','להקים','ממיר','מוקלד',
               'אישר','אושר','הוק','לעדכן לקוח','ירד להיום','ירדה להיום','ההתקנה להיום']
    if not any(k in full_text for k in sale_kw): return None

    subj_clean = re.sub(r'^(Re|Fwd|FW|RE):\s*', '', first_subject, flags=re.IGNORECASE).strip()
    subj_clean = re.sub(r'[-–].*', '', subj_clean).strip()
    if not subj_clean or len(subj_clean) < 2: return None

    cancel_phrases = ['לא להקים','התחרטה','לבטל','ביטול','מבטל','בוטל','ביטל','מבטלת','ביטלה']
    is_cancelled = any(p in full_text for p in cancel_phrases)
    cancel_note = next((p for p in ['התחרטה','לא להקים','ביטל','ביטלה'] if p in full_text), '')

    cid = ''
    cid_m = re.search(r'לקוח\s+(\d{7})', full_text)
    if cid_m: cid = cid_m.group(1)

    all_updates = []
    for msg in all_msgs:
        for u in find_updates_in_snippet(msg['snippet'], msg['date_ddmm'], msg['dt']):
            u['msg_i'] = msg['i']
            all_updates.append(u)

    kind_priority = {'today': 4, 'tomorrow': 3, 'recorded': 3, 'update': 2, 'explicit': 1}
    all_updates.sort(key=lambda u: (u['dt'] or datetime.min, kind_priority.get(u['kind'], 0)))

    install_date, install_hours, has_change, change_note, latest_kind = '', '', False, '', ''
    if all_updates:
        first_u, latest_u = all_updates[0], all_updates[-1]
        install_date = latest_u['date']
        install_hours = latest_u['hours']
        latest_kind = latest_u['kind']
        if first_u['date'] != latest_u['date'] or first_u['hours'] != latest_u['hours']:
            has_change = True
            change_note = f"שונה מ-{first_u['date']} {first_u['hours']} ל-{latest_u['date']} {latest_u['hours']}"

    if is_cancelled: status = 'בוטל'
    elif latest_kind in ('today','tomorrow','recorded','update') or \
         any(w in full_text for w in ['מוקלד','אושר','אישר','לעדכן לקוח']):
        status = 'מוקלד + הוק' if 'הוק' in full_text else 'מוקלד'
    elif 'לאשר חוזה' in full_text and 'הוק' in full_text: status = 'לאשר חוזה + הוק'
    elif 'לאשר חוזה' in full_text: status = 'לאשר חוזה'
    elif 'שובץ' in full_text: status = 'שובץ'
    else: status = 'בהקדם'

    mirrors = 0
    mm = re.search(r'(\d+)\s*ממירים?', full_text)
    if mm: mirrors = int(mm.group(1))

    install_month, is_april = '', False
    if install_date:
        parts = install_date.split('/')
        if len(parts) == 2:
            install_month = f'{int(parts[1]):02d}/2026'
            is_april = int(parts[1]) == 4

    sale_date = all_msgs[0]['date_full'] if all_msgs else ''
    return {'name': subj_clean, 'customerId': cid, 'saleDate': sale_date,
            'installDate': install_date, 'installMonth': install_month,
            'hours': install_hours, 'mirrors': mirrors, 'status': status,
            'hasChange': has_change, 'changeNote': change_note,
            'isCancelled': is_cancelled, 'cancelNote': cancel_note,
            'isToday': False, 'isApril': is_april, 'threadId': thread_id}

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
              'from:avielv014@gmail.com OR to:avielv014@gmail.com) after:2026/3/15',
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
