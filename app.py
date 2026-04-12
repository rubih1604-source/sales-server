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

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET', 'sales-rubi-2026')
CORS(app, origins="*")

CLIENT_ID     = os.environ.get('GOOGLE_CLIENT_ID')
CLIENT_SECRET = os.environ.get('GOOGLE_CLIENT_SECRET')
REDIRECT_URI  = 'https://sales-server-egdf.onrender.com/oauth/callback'
SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.compose',
]
TOKEN_FILE = '/tmp/gmail_token.json'

def get_client_config():
    return {"web": {
        "client_id": CLIENT_ID, "client_secret": CLIENT_SECRET,
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "redirect_uris": [REDIRECT_URI]
    }}

def get_credentials():
    if not os.path.exists(TOKEN_FILE): return None
    with open(TOKEN_FILE) as f: data = json.load(f)
    return Credentials(token=data.get('token'), refresh_token=data.get('refresh_token'),
        token_uri="https://oauth2.googleapis.com/token",
        client_id=CLIENT_ID, client_secret=CLIENT_SECRET, scopes=SCOPES)

def save_credentials(creds):
    with open(TOKEN_FILE, 'w') as f:
        json.dump({'token': creds.token, 'refresh_token': creds.refresh_token}, f)

def get_service():
    creds = get_credentials()
    return build('gmail', 'v1', credentials=creds) if creds else None

@app.route('/')
def index():
    connected = get_credentials() is not None
    st = "✅ מחובר ל-Gmail" if connected else "❌ לא מחובר"
    link = "" if connected else "<br><br><a href='/oauth/start' style='background:#3b82f6;color:white;padding:12px 24px;border-radius:8px;text-decoration:none;font-weight:bold'>🔐 התחבר ל-Gmail</a>"
    return f"<html dir='rtl'><body style='font-family:Arial;padding:40px;background:#0d1117;color:#e6edf3'><h1>שרת מכירות רובי</h1><p>סטטוס: <strong>{st}</strong></p>{link}</body></html>"

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
    save_credentials(flow.credentials)
    return "<html dir='rtl'><body style='font-family:Arial;padding:40px;background:#0d1117;color:#e6edf3'><h1>✅ התחברות הצליחה!</h1><p style='color:#10b981'>המערכת מחוברת. אפשר לסגור.</p></body></html>"

@app.route('/api/status')
def api_status():
    return jsonify({'connected': get_credentials() is not None})

def extract_body(payload, depth=0):
    if depth > 5: return ''
    text = ''
    if payload.get('body', {}).get('data'):
        try: text += base64.urlsafe_b64decode(payload['body']['data']).decode('utf-8', errors='ignore')
        except: pass
    for part in payload.get('parts', []):
        if part.get('mimeType', '') in ('text/plain', 'text/html'):
            if part.get('body', {}).get('data'):
                try: text += base64.urlsafe_b64decode(part['body']['data']).decode('utf-8', errors='ignore')
                except: pass
        text += extract_body(part, depth+1)
    return text[:5000]

def parse_date_str(date_str):
    try:
        from email.utils import parsedate
        p = parsedate(date_str)
        if p: return f'{p[2]:02d}/{p[1]:02d}/{p[0]}', f'{p[1]:02d}/{p[0]}'
    except: pass
    return '', ''

def find_attachments(payload, attachments, message_id):
    fn = payload.get('filename', '')
    if fn and fn.lower().endswith('.pdf'):
        att_id = payload.get('body', {}).get('attachmentId')
        if att_id: attachments.append({'filename': fn, 'attachmentId': att_id, 'messageId': message_id})
    for part in payload.get('parts', []): find_attachments(part, attachments, message_id)

def parse_sale_from_thread(messages, thread_id):
    full_text = ''
    first_subject = ''
    last_date = ''

    for i, m in enumerate(messages):
        hdrs = {h['name']: h['value'] for h in m['payload']['headers']}
        subj = hdrs.get('Subject', '')
        if i == 0: first_subject = subj
        last_date = hdrs.get('Date', last_date)
        full_text += subj + '\n' + m.get('snippet', '') + '\n' + extract_body(m['payload']) + '\n'

    # בדוק שזה thread מכירה
    sale_keywords = ['תואם', 'לאשר חוזה', 'ממירים', 'דאבל יס', 'דרבל יס', 'שובץ',
                     'להקים', 'ממיר', 'מוקלד', 'אישר', 'אושר', 'חוזה', 'הוק', 'YES']
    if not any(k in full_text for k in sale_keywords): return None

    # שם מנושא
    subj = re.sub(r'^(Re|Fwd|FW|RE):\s*', '', first_subject, flags=re.IGNORECASE).strip()
    subj = re.sub(r'[-–].*', '', subj).strip()
    if not subj or len(subj) < 2: return None

    # ביטול
    cancel_phrases = ['לא להקים', 'התחרטה', 'לבטל', 'ביטול', 'מבטל', 'בוטל', 'ביטל', 'מבטלת', 'ביטלה']
    is_cancelled = any(p in full_text for p in cancel_phrases)
    cancel_note = next((p for p in ['התחרטה', 'לא להקים', 'ביטל', 'ביטלה'] if p in full_text), '')

    # מספר לקוח
    cid = ''
    m = re.search(r'לקוח\s+(\d{7})', full_text)
    if m: cid = m.group(1)

    # תאריך ושעות — מילות מפתח לסטטוס סופי
    # מוקלד/אישר/אושר = סטטוס הגבוה ביותר
    confirmed_patterns = [
        r'מוקלד\s+(\d{1,2}/\d{2})\s+(\d{1,2}-\d{2})',
        r'אושר\s+(\d{1,2}/\d{2})\s+(\d{1,2}-\d{2})',
        r'אישר\s+(\d{1,2}/\d{2})\s+(\d{1,2}-\d{2})',
        r'מאושר\s+(\d{1,2}/\d{2})\s+(\d{1,2}-\d{2})',
        r'נקבע\s+(\d{1,2}/\d{2})\s+(\d{1,2}-\d{2})',
    ]
    scheduled_patterns = [
        r'תואם\s+(\d{1,2}/\d{2})\s+(\d{1,2}-\d{2})',
        r'שובץ\s+ל[-]?(\d{1,2}/\d{2})\s+(\d{1,2}-\d{2})',
        r'(\d{1,2}/\d{2})\s+(\d{1,2}-\d{2})',
    ]

    all_confirmed = []
    for pat in confirmed_patterns:
        for match in re.finditer(pat, full_text):
            all_confirmed.append((match.start(), match.group(1), match.group(2), 'confirmed'))

    all_scheduled = []
    for pat in scheduled_patterns:
        for match in re.finditer(pat, full_text):
            all_scheduled.append((match.start(), match.group(1), match.group(2), 'scheduled'))

    all_found = all_confirmed + all_scheduled
    all_found.sort(key=lambda x: x[0])

    install_date, install_hours, has_change, change_note, is_confirmed = '', '', False, '', False

    if all_found:
        latest = all_found[-1]
        install_date = latest[1]
        install_hours = latest[2]
        is_confirmed = latest[3] == 'confirmed'

        if len(all_found) > 1:
            first = all_found[0]
            if first[2] != latest[2] or first[1] != latest[1]:
                has_change = True
                change_note = f'שונה מ-{first[1]} {first[2]} ל-{latest[1]} {latest[2]}'

    # סטטוס — לפי מילות מפתח
    if is_cancelled:
        status = 'בוטל'
    elif is_confirmed or any(w in full_text for w in ['מוקלד', 'אושר', 'אישר', 'מאושר', 'נקבע']):
        # מוקלד = שובץ לוח זמנים
        if 'הוק' in full_text or 'HOK' in full_text.upper():
            status = 'מוקלד + הוק'
        else:
            status = 'מוקלד'
    elif any(w in full_text for w in ['לאשר חוזה', 'ממתין לחוזה', 'לאשר']) and ('הוק' in full_text or 'HOK' in full_text.upper()):
        status = 'לאשר חוזה + הוק'
    elif any(w in full_text for w in ['לאשר חוזה', 'ממתין לחוזה']):
        status = 'לאשר חוזה'
    elif 'שובץ' in full_text:
        status = 'שובץ'
    elif 'בנק התקנה' in full_text or 'בהקדם' in full_text:
        status = 'בהקדם'
    else:
        status = 'בהקדם'

    # ממירים
    mirrors = 0
    mm = re.search(r'(\d+)\s*ממירים?', full_text)
    if mm: mirrors = int(mm.group(1))

    # חודש
    install_month, is_april = '', False
    if install_date:
        parts = install_date.split('/')
        if len(parts) == 2:
            install_month = f'{parts[1]}/2026'
            is_april = parts[1] == '04'

    # תאריך המייל
    date, _ = parse_date_str(last_date)

    return {
        'name': subj,
        'customerId': cid,
        'installDate': install_date,
        'installMonth': install_month,
        'hours': install_hours,
        'mirrors': mirrors,
        'status': status,
        'hasChange': has_change,
        'changeNote': change_note,
        'isCancelled': is_cancelled,
        'cancelNote': cancel_note,
        'isConfirmed': is_confirmed,
        'isToday': False,
        'isApril': is_april,
        'threadId': thread_id,
        'lastDate': date
    }


@app.route('/api/scan')
def scan():
    service = get_service()
    if not service:
        return jsonify({'error': 'not_authenticated', 'auth_url': 'https://sales-server-egdf.onrender.com/oauth/start'}), 401
    try:
        sales, invoices = [], []

        results = service.users().messages().list(
            userId='me',
            q='(from:oshrityes2901@gmail.com OR to:oshrityes2901@gmail.com OR from:oritapiro22@gmail.com OR to:oritapiro22@gmail.com OR from:avielv014@gmail.com OR to:avielv014@gmail.com) after:2026/3/15',
            maxResults=300
        ).execute()

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

        # חשבוניות
        inv_results = service.users().messages().list(
            userId='me',
            q='(חשבונית OR invoice OR receipt OR morning.co OR render.com OR cardcom OR icount) has:attachment after:2026/1/1',
            maxResults=50
        ).execute()

        for msg in inv_results.get('messages', []):
            try:
                md = service.users().messages().get(userId='me', id=msg['id'], format='full').execute()
                hdrs = {h['name']: h['value'] for h in md['payload']['headers']}
                subj = hdrs.get('Subject', '')
                sender = hdrs.get('From', '')
                date_str = hdrs.get('Date', '')
                date, month = parse_date_str(date_str)
                biz = ['ר.א.מ','ליד מנג','render','funnelly','stripe','morning','חשבונית ירוקה','atp','icount','חשבונית מס','קבלה']
                if not any(k.lower() in (subj+sender).lower() for k in biz): continue
                atts = []
                find_attachments(md['payload'], atts, msg['id'])
                inv_type = 'income' if any(k in (subj+sender) for k in ['ר.א.מ','עמלות','may zalah','May zalah']) else \
                           ('subscription' if any(k in (subj+sender) for k in ['חשבונית ירוקה','morning','חשבון חודשי']) else 'expense')
                num_m = re.search(r'(\d{4,})', subj)
                name_m = re.match(r'^"?([^"<]+)', sender)
                invoices.append({
                    'id': msg['id'], 'date': date, 'month': month,
                    'from': name_m.group(1).strip() if name_m else sender,
                    'subject': subj, 'description': subj[:50],
                    'invoiceNum': num_m.group(1) if num_m else '',
                    'type': inv_type, 'amount': None, 'currency': 'ILS',
                    'gmailLink': f'https://mail.google.com/mail/u/0/#all/{msg["id"]}',
                    'hasAttachment': len(atts) > 0, 'attachments': atts
                })
            except: continue

        sales.sort(key=lambda x: (x.get('installMonth',''), x.get('installDate','')), reverse=True)
        return jsonify({'success': True, 'sales': sales, 'invoices': invoices,
                       'counts': {'sales': len(sales), 'invoices': len(invoices)}})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/attachment/<message_id>/<attachment_id>')
def get_attachment(message_id, attachment_id):
    service = get_service()
    if not service: return jsonify({'error': 'not_authenticated'}), 401
    try:
        att = service.users().messages().attachments().get(userId='me', messageId=message_id, id=attachment_id).execute()
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
        income = [i for i in invoices if i.get('type') == 'income']
        expense = [i for i in invoices if i.get('type') != 'income']
        if income:
            body += f'הכנסות ({len(income)}):\n'
            for inv in income: body += f'• {inv.get("from","")} — {inv.get("description","")}\n'
        if expense:
            body += f'\nהוצאות ({len(expense)}):\n'
            for inv in expense: body += f'• {inv.get("from","")} — {inv.get("description","")}\n'
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
