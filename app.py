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
SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.compose',
]
TOKEN_FILE = '/tmp/gmail_token.json'

# כל תבנית שעות אפשרית: 8-10, 10-12, 12-14, 14-16, 16-18, 18-20 וכו'
HOURS_RE = r'(\d{1,2}[-–]\d{2})'

# מילות מפתח "היום" — מחוברות לתאריך המייל
TODAY_KW = [
    'ירד להיום', 'ירדה להיום', 'ההתקנה להיום', 'התקנה להיום',
    'להיום', 'היום', 'שובץ להיום', 'תואם להיום', 'מוקלד להיום',
    'ירד ל-היום', 'ירדה ל-היום',
]

def get_client_config():
    return {"web": {"client_id": CLIENT_ID, "client_secret": CLIENT_SECRET,
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "redirect_uris": [REDIRECT_URI]}}

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
    st = "✅ מחובר" if connected else "❌ לא מחובר"
    link = "" if connected else "<br><a href='/oauth/start' style='background:#3b82f6;color:white;padding:12px 24px;border-radius:8px;text-decoration:none'>התחבר</a>"
    return f"<html dir='rtl'><body style='font-family:Arial;padding:40px;background:#0d1117;color:#e6edf3'><h1>שרת מכירות רובי</h1><p>{st}</p>{link}</body></html>"

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
    return "<html dir='rtl'><body style='background:#0d1117;color:#e6edf3;padding:40px;font-family:Arial'><h1>✅ מחובר!</h1><p style='color:#10b981'>סגור חלון זה.</p></body></html>"

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

def parse_email_date(date_str):
    """פרסר תאריך מייל → (DD/MM/YYYY, DD/MM, datetime)"""
    try:
        from email.utils import parsedate
        p = parsedate(date_str)
        if p:
            dt = datetime(p[0], p[1], p[2])
            return f'{p[2]:02d}/{p[1]:02d}/{p[0]}', f'{p[2]:02d}/{p[1]:02d}', dt
    except: pass
    return '', '', None

def ddmm_offset(ddmm, days):
    """DD/MM + N ימים → DD/MM"""
    try:
        parts = ddmm.split('/')
        dt = datetime(2026, int(parts[1]), int(parts[0])) + timedelta(days=days)
        return f'{dt.day:02d}/{dt.month:02d}'
    except: return ''

def find_attachments(payload, attachments, message_id):
    fn = payload.get('filename', '')
    if fn and fn.lower().endswith('.pdf'):
        att_id = payload.get('body', {}).get('attachmentId')
        if att_id: attachments.append({'filename': fn, 'attachmentId': att_id, 'messageId': message_id})
    for part in payload.get('parts', []): find_attachments(part, attachments, message_id)

def extract_updates_from_message(snippet_and_body, msg_ddmm, msg_dt):
    """
    חלץ עדכוני תאריך+שעה ממייל אחד.
    
    הכלל החשוב: כשכתוב "היום 16-18" — התאריך הוא תאריך המייל.
    כשכתוב "23/04 16-18" — התאריך הוא 23/04.
    
    מחזיר רשימה של dict: {date, hours, kind, msg_dt}
    """
    text = snippet_and_body
    results = []

    # ── קודם כל: דפוסים עם "היום" / "ירד להיום" / "ההתקנה להיום" ──
    # אלה מחוברים תמיד לתאריך המייל
    today_pattern = '|'.join(re.escape(k) for k in TODAY_KW)
    
    # "ירד להיום 16-18" / "היום 16-18" / "ההתקנה להיום 8-10"
    for m in re.finditer(r'(?:' + today_pattern + r')\s*' + HOURS_RE, text):
        results.append({
            'date': msg_ddmm,
            'hours': m.group(1),
            'kind': 'today',
            'msg_dt': msg_dt,
        })
    
    # "16-18 להיום" (סדר הפוך)
    for m in re.finditer(HOURS_RE + r'\s+(?:' + today_pattern + r')', text):
        results.append({
            'date': msg_ddmm,
            'hours': m.group(1),
            'kind': 'today',
            'msg_dt': msg_dt,
        })

    # ── "מחר HH-HH" ────────────────────────────────────────
    for m in re.finditer(r'מחר\s+' + HOURS_RE, text):
        tomorrow = ddmm_offset(msg_ddmm, 1)
        if tomorrow:
            results.append({
                'date': tomorrow,
                'hours': m.group(1),
                'kind': 'tomorrow',
                'msg_dt': msg_dt,
            })

    # ── דפוסים עם תאריך מפורש ──────────────────────────────
    # "שובץ 23/04 16-18 לעדכן לקוח"
    # "תואם 14/04 10-12"
    # "מוקלד 13/04 8-10"
    # "לקוח XXXXXXX תואם DD/MM HH-HH"
    # "DD/MM HH-HH לעדכן לקוח"
    # "DD/MM HH-HH לאשר חוזה"
    explicit = re.findall(
        r'(?:מוקלד|תואם|שובץ|אושר|אישר|נקבע|שיבוץ)?\s*(\d{1,2}/\d{2})\s+' + HOURS_RE,
        text
    )
    for date, hours in explicit:
        results.append({
            'date': date,
            'hours': hours,
            'kind': 'explicit',
            'msg_dt': msg_dt,
        })

    # "DD/MM HH-HH לעדכן" / "DD/MM HH-HH לאשר"
    explicit2 = re.findall(
        r'(\d{1,2}/\d{2})\s+' + HOURS_RE + r'\s+(?:לעדכן|לאשר)',
        text
    )
    for date, hours in explicit2:
        results.append({
            'date': date,
            'hours': hours,
            'kind': 'explicit',
            'msg_dt': msg_dt,
        })

    return results

def parse_sale_from_thread(messages, thread_id):
    """
    פרסר thread מכירה.
    
    לוגיקת עדכון תאריך:
    - כל עדכון שעות מקושר לתאריך המייל שבו הוא הופיע
    - העדכון האחרון (לפי תאריך המייל) מנצח
    - "היום 16-18" במייל מ-14/04 = 14/04 בשעות 16-18
    """
    full_text = ''
    first_subject = ''
    all_messages_info = []

    for i, m in enumerate(messages):
        hdrs = {h['name']: h['value'] for h in m['payload']['headers']}
        subj = hdrs.get('Subject', '')
        date_str = hdrs.get('Date', '')
        date_full, date_ddmm, msg_dt = parse_email_date(date_str)
        if i == 0: first_subject = subj
        
        # קח רק את ה-snippet וחלק קטן מהגוף — לא threadים ישנים
        snippet = m.get('snippet', '')
        # חלץ גוף ראשון בלבד (לא ציטוטים)
        body = extract_body(m['payload'])
        # חתוך ציטוטים ישנים — הם מתחילים ב "בתאריך יום" או "> "
        clean_body = re.split(r'בתאריך יום\s+[אבגדהוז]', body)[0]
        clean_body = re.split(r'On [A-Za-z]', clean_body)[0]
        
        text_for_parse = snippet + '\n' + clean_body
        full_text += subj + '\n' + text_for_parse + '\n'
        
        all_messages_info.append({
            'index': i,
            'date_full': date_full,
            'date_ddmm': date_ddmm,
            'msg_dt': msg_dt,
            'text': text_for_parse,
            'subject': subj,
        })

    # בדוק שזה thread מכירה
    sale_kw = ['תואם', 'לאשר חוזה', 'ממירים', 'דאבל יס', 'דרבל יס', 'שובץ',
               'להקים', 'ממיר', 'מוקלד', 'אישר', 'אושר', 'הוק',
               'לעדכן לקוח', 'ירד להיום', 'ירדה להיום', 'ההתקנה להיום']
    if not any(k in full_text for k in sale_kw): return None

    # שם מנושא
    subj_clean = re.sub(r'^(Re|Fwd|FW|RE):\s*', '', first_subject, flags=re.IGNORECASE).strip()
    subj_clean = re.sub(r'[-–].*', '', subj_clean).strip()
    if not subj_clean or len(subj_clean) < 2: return None

    # ביטול
    cancel_phrases = ['לא להקים', 'התחרטה', 'לבטל', 'ביטול', 'מבטל', 'בוטל', 'ביטל', 'מבטלת', 'ביטלה']
    is_cancelled = any(p in full_text for p in cancel_phrases)
    cancel_note = next((p for p in ['התחרטה', 'לא להקים', 'ביטל', 'ביטלה'] if p in full_text), '')

    # מספר לקוח
    cid = ''
    cid_m = re.search(r'לקוח\s+(\d{7})', full_text)
    if cid_m: cid = cid_m.group(1)

    # ── אסוף עדכונים מכל מייל בנפרד ──────────────────────
    all_updates = []
    for msg_info in all_messages_info:
        updates = extract_updates_from_message(
            msg_info['text'],
            msg_info['date_ddmm'],
            msg_info['msg_dt']
        )
        for u in updates:
            u['msg_index'] = msg_info['index']
            all_updates.append(u)

    # ── בחר את העדכון הנכון ────────────────────────────────
    # מיין לפי תאריך המייל (msg_dt) — האחרון מנצח
    # אם אין msg_dt — לפי אינדקס
    def sort_key(u):
        dt = u.get('msg_dt')
        if dt: return (dt, u['msg_index'])
        return (datetime.min, u['msg_index'])
    
    all_updates.sort(key=sort_key)

    install_date, install_hours, has_change, change_note = '', '', False, ''
    first_update = None
    latest_update = None

    if all_updates:
        first_update = all_updates[0]
        latest_update = all_updates[-1]
        install_date = latest_update['date']
        install_hours = latest_update['hours']

        if len(all_updates) > 1:
            if first_update['date'] != latest_update['date'] or \
               first_update['hours'] != latest_update['hours']:
                has_change = True
                change_note = f"שונה מ-{first_update['date']} {first_update['hours']} ל-{latest_update['date']} {latest_update['hours']}"

    # ── סטטוס ────────────────────────────────────────────
    has_today_kw = any(kw in full_text for kw in TODAY_KW)
    latest_kind = latest_update['kind'] if latest_update else ''

    if is_cancelled:
        status = 'בוטל'
    elif latest_kind in ('today', 'tomorrow') or any(w in full_text for w in ['מוקלד','אושר','אישר','לעדכן לקוח']):
        status = 'מוקלד + הוק' if ('הוק' in full_text) else 'מוקלד'
    elif 'לאשר חוזה' in full_text and 'הוק' in full_text:
        status = 'לאשר חוזה + הוק'
    elif 'לאשר חוזה' in full_text:
        status = 'לאשר חוזה'
    elif 'שובץ' in full_text:
        status = 'שובץ'
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

    sale_date = all_messages_info[0]['date_full'] if all_messages_info else ''

    return {
        'name': subj_clean,
        'customerId': cid,
        'saleDate': sale_date,
        'installDate': install_date,
        'installMonth': install_month,
        'hours': install_hours,
        'mirrors': mirrors,
        'status': status,
        'hasChange': has_change,
        'changeNote': change_note,
        'isCancelled': is_cancelled,
        'cancelNote': cancel_note,
        'isToday': False,
        'isApril': is_april,
        'threadId': thread_id
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
                date_full, _, _ = parse_email_date(date_str)
                month = date_full[3:10] if date_full else ''
                biz = ['ר.א.מ','ליד מנג','render','funnelly','stripe','morning','חשבונית ירוקה','atp','icount','חשבונית מס','קבלה']
                if not any(k.lower() in (subj+sender).lower() for k in biz): continue
                atts = []
                find_attachments(md['payload'], atts, msg['id'])
                inv_type = 'income' if any(k in (subj+sender) for k in ['ר.א.מ','עמלות','may zalah','May zalah']) else \
                           ('subscription' if any(k in (subj+sender) for k in ['חשבונית ירוקה','morning','חשבון חודשי']) else 'expense')
                num_m = re.search(r'(\d{4,})', subj)
                name_m = re.match(r'^"?([^"<]+)', sender)
                invoices.append({
                    'id': msg['id'], 'date': date_full, 'month': month,
                    'from': name_m.group(1).strip() if name_m else sender,
                    'subject': subj, 'description': subj[:50],
                    'invoiceNum': num_m.group(1) if num_m else '',
                    'type': inv_type, 'amount': None, 'currency': 'ILS',
                    'gmailLink': f'https://mail.google.com/mail/u/0/#all/{msg["id"]}',
                    'hasAttachment': len(atts) > 0, 'attachments': atts
                })
            except: continue

        sales.sort(key=lambda x: x.get('saleDate') or '', reverse=True)
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
