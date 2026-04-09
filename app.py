import os
import json
import base64
from flask import Flask, request, jsonify, redirect, session
from flask_cors import CORS
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET', 'sales-rubi-secret-2026')
CORS(app, origins="*")

# ── ENV VARS (set in Render dashboard) ──────────────────────
CLIENT_ID     = os.environ.get('GOOGLE_CLIENT_ID')
CLIENT_SECRET = os.environ.get('GOOGLE_CLIENT_SECRET')
REDIRECT_URI  = 'https://sales-rubi.onrender.com/oauth/callback'
SCOPES        = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.send',
]

TOKEN_FILE = '/tmp/gmail_token.json'

# ── helpers ──────────────────────────────────────────────────
def get_client_config():
    return {
        "web": {
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": [REDIRECT_URI]
        }
    }

def get_credentials():
    if not os.path.exists(TOKEN_FILE):
        return None
    with open(TOKEN_FILE) as f:
        token_data = json.load(f)
    creds = Credentials(
        token=token_data.get('token'),
        refresh_token=token_data.get('refresh_token'),
        token_uri="https://oauth2.googleapis.com/token",
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET,
        scopes=SCOPES
    )
    return creds

def save_credentials(creds):
    with open(TOKEN_FILE, 'w') as f:
        json.dump({
            'token': creds.token,
            'refresh_token': creds.refresh_token,
        }, f)

def get_gmail_service():
    creds = get_credentials()
    if not creds:
        return None
    return build('gmail', 'v1', credentials=creds)

# ── ROUTES ───────────────────────────────────────────────────

@app.route('/')
def index():
    creds = get_credentials()
    status = "✅ מחובר ל-Gmail" if creds else "❌ לא מחובר"
    return f"""
    <html dir="rtl"><body style="font-family:Arial;padding:40px;background:#0d1117;color:#e6edf3">
    <h1>🔧 שרת מכירות רובי</h1>
    <p>סטטוס: <strong>{status}</strong></p>
    {"<a href='/oauth/start' style='background:#3b82f6;color:white;padding:12px 24px;border-radius:8px;text-decoration:none;font-weight:bold'>🔐 התחבר ל-Gmail</a>" if not creds else "<p>השרת פעיל ומחובר!</p>"}
    </body></html>
    """

@app.route('/oauth/start')
def oauth_start():
    flow = Flow.from_client_config(
        get_client_config(),
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI
    )
    auth_url, state = flow.authorization_url(
        access_type='offline',
        include_granted_scopes='true',
        prompt='consent'
    )
    session['oauth_state'] = state
    return redirect(auth_url)

@app.route('/oauth/callback')
def oauth_callback():
    flow = Flow.from_client_config(
        get_client_config(),
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI
    )
    flow.fetch_token(authorization_response=request.url)
    creds = flow.credentials
    save_credentials(creds)
    return """
    <html dir="rtl"><body style="font-family:Arial;padding:40px;background:#0d1117;color:#e6edf3">
    <h1>✅ התחברות הצליחה!</h1>
    <p>השרת מחובר ל-Gmail שלך. אפשר לסגור את החלון הזה.</p>
    <p style="color:#10b981">המערכת שלך מוכנה לשימוש!</p>
    </body></html>
    """

# ── API: SCAN SALES ──────────────────────────────────────────
@app.route('/api/sales')
def get_sales():
    service = get_gmail_service()
    if not service:
        return jsonify({'error': 'not_authenticated', 'auth_url': '/oauth/start'}), 401

    try:
        # Search for sales emails
        results = service.users().messages().list(
            userId='me',
            q='from:rubih1604@gmail.com to:(oshrityes2901@gmail.com OR oritapiro22@gmail.com OR avielv014@gmail.com) after:2026/3/15',
            maxResults=100
        ).execute()

        messages = results.get('messages', [])
        sales = []

        for msg in messages:
            msg_data = service.users().messages().get(
                userId='me', id=msg['id'], format='full'
            ).execute()

            headers = {h['name']: h['value'] for h in msg_data['payload']['headers']}
            snippet = msg_data.get('snippet', '')
            subject = headers.get('Subject', '')
            date = headers.get('Date', '')

            # Skip RE: emails (replies) — only original sale emails
            if subject.startswith('Re:') or subject.startswith('RE:'):
                continue

            # Extract body
            body = ''
            payload = msg_data['payload']
            if payload.get('body', {}).get('data'):
                body = base64.urlsafe_b64decode(payload['body']['data']).decode('utf-8', errors='ignore')
            elif payload.get('parts'):
                for part in payload['parts']:
                    if part.get('mimeType') == 'text/plain' and part.get('body', {}).get('data'):
                        body = base64.urlsafe_b64decode(part['body']['data']).decode('utf-8', errors='ignore')
                        break

            sales.append({
                'id': msg['id'],
                'subject': subject,
                'date': date,
                'snippet': snippet,
                'body': body[:500],
            })

        return jsonify({'sales': sales, 'count': len(sales)})

    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── API: SCAN INVOICES ───────────────────────────────────────
@app.route('/api/invoices')
def get_invoices():
    service = get_gmail_service()
    if not service:
        return jsonify({'error': 'not_authenticated'}), 401

    try:
        results = service.users().messages().list(
            userId='me',
            q='(חשבונית OR invoice OR receipt) has:attachment after:2026/1/1',
            maxResults=50
        ).execute()

        messages = results.get('messages', [])
        invoices = []

        for msg in messages:
            msg_data = service.users().messages().get(
                userId='me', id=msg['id'], format='full'
            ).execute()

            headers = {h['name']: h['value'] for h in msg_data['payload']['headers']}
            subject = headers.get('Subject', '')
            sender = headers.get('From', '')
            date = headers.get('Date', '')

            # Find attachments
            attachments = []
            def find_attachments(payload):
                if payload.get('filename') and payload['filename'].endswith('.pdf'):
                    attachments.append({
                        'filename': payload['filename'],
                        'attachmentId': payload.get('body', {}).get('attachmentId'),
                        'messageId': msg['id']
                    })
                for part in payload.get('parts', []):
                    find_attachments(part)

            find_attachments(msg_data['payload'])

            invoices.append({
                'id': msg['id'],
                'subject': subject,
                'from': sender,
                'date': date,
                'attachments': attachments,
                'gmailLink': f'https://mail.google.com/mail/u/0/#all/{msg["id"]}'
            })

        return jsonify({'invoices': invoices, 'count': len(invoices)})

    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── API: DOWNLOAD ATTACHMENT ─────────────────────────────────
@app.route('/api/attachment/<message_id>/<attachment_id>')
def get_attachment(message_id, attachment_id):
    service = get_gmail_service()
    if not service:
        return jsonify({'error': 'not_authenticated'}), 401

    try:
        attachment = service.users().messages().attachments().get(
            userId='me',
            messageId=message_id,
            id=attachment_id
        ).execute()

        data = base64.urlsafe_b64decode(attachment['data'])
        
        from flask import Response
        return Response(
            data,
            mimetype='application/pdf',
            headers={'Content-Disposition': 'attachment; filename=invoice.pdf'}
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── API: SEND TO ACCOUNTANT ──────────────────────────────────
@app.route('/api/send-to-accountant', methods=['POST'])
def send_to_accountant():
    service = get_gmail_service()
    if not service:
        return jsonify({'error': 'not_authenticated'}), 401

    data = request.json
    month = data.get('month', '')
    invoices = data.get('invoices', [])
    accountant_email = 'ei@eicpa.co.il'

    try:
        # Build email with real PDF attachments
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText
        from email.mime.base import MIMEBase
        from email import encoders

        msg = MIMEMultipart()
        msg['To'] = accountant_email
        msg['Subject'] = f'חשבוניות {month} — ראובן חגג'

        # Body
        body_text = f'שלום,\n\nמצורפות חשבוניות לחודש {month}.\n\n'
        income = [i for i in invoices if i.get('type') == 'income']
        expense = [i for i in invoices if i.get('type') != 'income']

        if income:
            body_text += f'הכנסות ({len(income)}):\n'
            for inv in income:
                body_text += f'• {inv.get("from", "")} — {inv.get("subject", "")}\n'

        if expense:
            body_text += f'\nהוצאות ({len(expense)}):\n'
            for inv in expense:
                body_text += f'• {inv.get("from", "")} — {inv.get("subject", "")}\n'

        body_text += '\nבברכה,\nרובי חגג'
        msg.attach(MIMEText(body_text, 'plain', 'utf-8'))

        # Attach PDFs
        for inv in invoices:
            for att in inv.get('attachments', []):
                if att.get('attachmentId'):
                    try:
                        attachment_data = service.users().messages().attachments().get(
                            userId='me',
                            messageId=att['messageId'],
                            id=att['attachmentId']
                        ).execute()

                        pdf_data = base64.urlsafe_b64decode(attachment_data['data'])
                        part = MIMEBase('application', 'pdf')
                        part.set_payload(pdf_data)
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename="{att["filename"]}"')
                        msg.attach(part)
                    except:
                        pass

        # Send
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(
            userId='me',
            body={'raw': raw}
        ).execute()

        return jsonify({'success': True, 'message': f'נשלח לרואה חשבון ({accountant_email})'})

    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── API: STATUS ──────────────────────────────────────────────
@app.route('/api/status')
def status():
    creds = get_credentials()
    return jsonify({
        'connected': creds is not None,
        'auth_url': 'https://sales-rubi.onrender.com/oauth/start'
    })

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
