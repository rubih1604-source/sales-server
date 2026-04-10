import os
import json
import base64
from flask import Flask, request, jsonify, redirect, session, Response
from flask_cors import CORS
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build

os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET', 'sales-rubi-2026-fixed')
CORS(app, origins="*")

CLIENT_ID     = os.environ.get('GOOGLE_CLIENT_ID')
CLIENT_SECRET = os.environ.get('GOOGLE_CLIENT_SECRET')
REDIRECT_URI  = 'https://sales-server-egdf.onrender.com/oauth/callback'
SCOPES        = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.send',
]
TOKEN_FILE = '/tmp/gmail_token.json'
STATE_FILE = '/tmp/oauth_state.txt'

def get_client_config():
    return {"web": {"client_id": CLIENT_ID, "client_secret": CLIENT_SECRET,
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "redirect_uris": [REDIRECT_URI]}}

def get_credentials():
    if not os.path.exists(TOKEN_FILE): return None
    with open(TOKEN_FILE) as f: token_data = json.load(f)
    return Credentials(token=token_data.get('token'), refresh_token=token_data.get('refresh_token'),
        token_uri="https://oauth2.googleapis.com/token", client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET, scopes=SCOPES)

def save_credentials(creds):
    with open(TOKEN_FILE, 'w') as f:
        json.dump({'token': creds.token, 'refresh_token': creds.refresh_token}, f)

def get_service():
    creds = get_credentials()
    return build('gmail', 'v1', credentials=creds) if creds else None

@app.route('/')
def index():
    creds = get_credentials()
    s = "✅ מחובר ל-Gmail" if creds else "❌ לא מחובר"
    btn = "" if creds else '<a href="/oauth/start" style="background:#3b82f6;color:white;padding:12px 24px;border-radius:8px;text-decoration:none;font-weight:bold;display:inline-block;margin-top:20px">🔐 התחבר ל-Gmail</a>'
    return f'<html dir="rtl"><body style="font-family:Arial;padding:40px;background:#0d1117;color:#e6edf3"><h1>🔧 שרת מכירות רובי</h1><p>סטטוס: <strong>{s}</strong></p>{btn}</body></html>'

@app.route('/oauth/start')
def oauth_start():
    flow = Flow.from_client_config(get_client_config(), scopes=SCOPES, redirect_uri=REDIRECT_URI)
    auth_url, state = flow.authorization_url(access_type='offline', prompt='consent')
    with open(STATE_FILE, 'w') as f: f.write(state)
    return redirect(auth_url)

@app.route('/oauth/callback')
def oauth_callback():
    try:
        state = open(STATE_FILE).read().strip() if os.path.exists(STATE_FILE) else None
        flow = Flow.from_client_config(get_client_config(), scopes=SCOPES, redirect_uri=REDIRECT_URI, state=state)
        url = request.url.replace('http://', 'https://')
        flow.fetch_token(authorization_response=url)
        save_credentials(flow.credentials)
        return '<html dir="rtl"><body style="font-family:Arial;padding:40px;background:#0d1117;color:#e6edf3;text-align:center"><h1 style="color:#10b981">✅ התחברות הצליחה!</h1><p style="font-size:18px">השרת מחובר ל-Gmail שלך.</p><p style="color:#10b981">המערכת מוכנה לשימוש! 🎉</p></body></html>'
    except Exception as e:
        return f'<html dir="rtl"><body style="font-family:Arial;padding:40px;background:#0d1117;color:#e6edf3"><h1 style="color:#ef4444">❌ שגיאה: {str(e)}</h1><br><a href="/oauth/start" style="background:#3b82f6;color:white;padding:12px 24px;border-radius:8px;text-decoration:none">נסה שוב</a></body></html>'

@app.route('/api/status')
def api_status():
    return jsonify({'connected': get_credentials() is not None})

@app.route('/api/sales')
def get_sales():
    svc = get_service()
    if not svc: return jsonify({'error': 'not_authenticated'}), 401
    try:
        res = svc.users().messages().list(userId='me',
            q='from:rubih1604@gmail.com to:(oshrityes2901@gmail.com OR oritapiro22@gmail.com OR avielv014@gmail.com) after:2026/3/15',
            maxResults=100).execute()
        sales = []
        for msg in res.get('messages', []):
            m = svc.users().messages().get(userId='me', id=msg['id'], format='full').execute()
            headers = {h['name']: h['value'] for h in m['payload']['headers']}
            subj = headers.get('Subject', '')
            if subj.startswith('Re:') or subj.startswith('RE:'): continue
            body = ''
            pl = m['payload']
            if pl.get('body', {}).get('data'):
                body = base64.urlsafe_b64decode(pl['body']['data']).decode('utf-8', errors='ignore')
            elif pl.get('parts'):
                for p in pl['parts']:
                    if p.get('mimeType') == 'text/plain' and p.get('body', {}).get('data'):
                        body = base64.urlsafe_b64decode(p['body']['data']).decode('utf-8', errors='ignore')
                        break
            sales.append({'id': msg['id'], 'subject': subj, 'date': headers.get('Date',''), 'body': body[:600]})
        return jsonify({'sales': sales, 'count': len(sales)})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/invoices')
def get_invoices():
    svc = get_service()
    if not svc: return jsonify({'error': 'not_authenticated'}), 401
    try:
        res = svc.users().messages().list(userId='me',
            q='(חשבונית OR invoice OR receipt) has:attachment after:2026/1/1', maxResults=50).execute()
        invoices = []
        for msg in res.get('messages', []):
            m = svc.users().messages().get(userId='me', id=msg['id'], format='full').execute()
            headers = {h['name']: h['value'] for h in m['payload']['headers']}
            atts = []
            def find_atts(pl):
                if pl.get('filename','').lower().endswith('.pdf'):
                    atts.append({'filename': pl['filename'], 'attachmentId': pl.get('body',{}).get('attachmentId'), 'messageId': msg['id']})
                for p in pl.get('parts',[]): find_atts(p)
            find_atts(m['payload'])
            invoices.append({'id': msg['id'], 'subject': headers.get('Subject',''), 'from': headers.get('From',''),
                'date': headers.get('Date',''), 'attachments': atts, 'gmailLink': f'https://mail.google.com/mail/u/0/#all/{msg["id"]}'})
        return jsonify({'invoices': invoices, 'count': len(invoices)})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/attachment/<mid>/<aid>')
def get_attachment(mid, aid):
    svc = get_service()
    if not svc: return jsonify({'error': 'not_authenticated'}), 401
    try:
        att = svc.users().messages().attachments().get(userId='me', messageId=mid, id=aid).execute()
        data = base64.urlsafe_b64decode(att['data'])
        return Response(data, mimetype='application/pdf', headers={'Content-Disposition': 'attachment; filename=invoice.pdf'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/send-to-accountant', methods=['POST'])
def send_to_accountant():
    svc = get_service()
    if not svc: return jsonify({'error': 'not_authenticated'}), 401
    data = request.json
    month = data.get('month', '')
    invoices = data.get('invoices', [])
    try:
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText
        from email.mime.base import MIMEBase
        from email import encoders
        msg = MIMEMultipart()
        msg['To'] = 'ei@eicpa.co.il'
        msg['Subject'] = f'חשבוניות {month} — ראובן חגג'
        body = f'שלום,\n\nמצורפות חשבוניות לחודש {month}.\n\n'
        for inv in invoices:
            body += f'• {inv.get("from","")} — {inv.get("subject","")}\n'
        body += '\nבברכה,\nרובי חגג'
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        for inv in invoices:
            for att in inv.get('attachments', []):
                if att.get('attachmentId'):
                    try:
                        ad = svc.users().messages().attachments().get(userId='me', messageId=att['messageId'], id=att['attachmentId']).execute()
                        pdf = base64.urlsafe_b64decode(ad['data'])
                        part = MIMEBase('application', 'pdf')
                        part.set_payload(pdf)
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename="{att["filename"]}"')
                        msg.attach(part)
                    except: pass
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        svc.users().messages().send(userId='me', body={'raw': raw}).execute()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
