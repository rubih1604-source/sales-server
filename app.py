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
TOKEN_FILE = '/tmp/gmail_token.json'

# ── תבניות ────────────────────────────────────────────────
HOURS_RE = r'(\d{1,2}(?::?\d{2})?[-–]\d{1,2}(?::?\d{2})?)'
TODAY_KW = sorted(['ירד להיום','ירדה להיום','ההתקנה להיום','התקנה להיום',
                   'שובץ להיום','תואם להיום','מוקלד להיום','להיום','היום'],key=len,reverse=True)
STATUS_KW = r'(?:מוקלד|תואם|מתואם|שובץ|אושר|אישר|נקבע|מתוזמן|שיבוץ|תואמה)'

# ימי שבוע עברית → Python weekday (Mon=0..Sun=6)
WEEKDAYS = {'ראשון':6,'שני':0,'שלישי':1,'רביעי':2,'חמישי':3,'שישי':4,'שבת':5}
DAY_PAT = r'(?:יום\s+)?(?:ל)?(ראשון|שני|שלישי|רביעי|חמישי|שישי|שבת)(?:\s+הבא)?'

def get_client_config():
    return {"web":{"client_id":CLIENT_ID,"client_secret":CLIENT_SECRET,
        "auth_uri":"https://accounts.google.com/o/oauth2/auth",
        "token_uri":"https://oauth2.googleapis.com/token","redirect_uris":[REDIRECT_URI]}}

def get_credentials():
    if not os.path.exists(TOKEN_FILE): return None
    with open(TOKEN_FILE) as f: data=json.load(f)
    return Credentials(token=data.get('token'),refresh_token=data.get('refresh_token'),
        token_uri="https://oauth2.googleapis.com/token",
        client_id=CLIENT_ID,client_secret=CLIENT_SECRET,scopes=SCOPES)

def save_credentials(creds):
    with open(TOKEN_FILE,'w') as f:
        json.dump({'token':creds.token,'refresh_token':creds.refresh_token},f)

def get_service():
    creds=get_credentials()
    if not creds: return None
    try:
        if creds.expired and creds.refresh_token:
            from google.auth.transport.requests import Request
            creds.refresh(Request())
            save_credentials(creds)
        return build('gmail','v1',credentials=creds)
    except: return None

@app.route('/')
def index():
    c=get_credentials() is not None
    st="✅ מחובר ל-Gmail" if c else "❌ לא מחובר"
    link="" if c else "<br><br><a href='/oauth/start' style='background:#3b82f6;color:white;padding:14px 28px;border-radius:8px;text-decoration:none;font-size:16px;font-weight:bold'>🔐 התחבר ל-Gmail</a>"
    return f"<html dir='rtl'><body style='font-family:Arial;padding:40px;background:#0d1117;color:#e6edf3'><h1>שרת מכירות רובי</h1><p style='font-size:18px'>{st}</p>{link}</body></html>"

@app.route('/oauth/start')
def oauth_start():
    flow=Flow.from_client_config(get_client_config(),scopes=SCOPES,redirect_uri=REDIRECT_URI)
    url,state=flow.authorization_url(access_type='offline',prompt='consent')
    return redirect(url)

@app.route('/oauth/callback')
def oauth_callback():
    flow=Flow.from_client_config(get_client_config(),scopes=SCOPES,redirect_uri=REDIRECT_URI)
    flow.fetch_token(authorization_response=request.url)
    save_credentials(flow.credentials)
    return "<html dir='rtl'><body style='background:#0d1117;color:#e6edf3;padding:40px;font-family:Arial;text-align:center'><h1>✅ התחברות הצליחה!</h1><p style='color:#10b981;font-size:18px'>סגור חלון זה וחזור למערכת.</p></body></html>"

@app.route('/api/status')
def api_status():
    return jsonify({'connected':get_credentials() is not None})

# ── עזרי תאריך ────────────────────────────────────────────
def parse_email_date(date_str):
    try:
        from email.utils import parsedate
        p=parsedate(date_str)
        if p:
            dt=datetime(p[0],p[1],p[2])
            return f'{p[2]:02d}/{p[1]:02d}/{p[0]}', f'{p[2]:02d}/{p[1]:02d}', dt
    except: pass
    return '','',None

def nh(h):
    """נרמל שעות: 08:00-10:00, 0800-1000, 8-10 → 8-10"""
    h=str(h).replace('–','-').strip()
    h=re.sub(r':00','',h)
    h=re.sub(r'\b0?(\d{1,2})00\b',r'\1',h)
    p=h.split('-')
    if len(p)==2:
        try: return f'{int(p[0])}-{int(p[1])}'
        except: pass
    return h

def nd(d):
    """14/6 → 14/06"""
    p=str(d).strip().split('/')
    return f'{p[0].zfill(2)}/{p[1].zfill(2)}' if len(p)==2 else str(d)

def offset_ddmm(ddmm, days):
    try:
        p=ddmm.split('/'); dt=datetime(2026,int(p[1]),int(p[0]))+timedelta(days=days)
        return f'{dt.day:02d}/{dt.month:02d}'
    except: return ''

def next_weekday_ddmm(msg_dt, day_name):
    """היום הקרוב מסוג day_name אחרי תאריך המייל → DD/MM"""
    if not msg_dt or day_name not in WEEKDAYS: return ''
    target=WEEKDAYS[day_name]
    days_ahead=(target - msg_dt.weekday()) % 7
    # אם 0 — זה אותו יום, אבל "ראשון הבא" בד"כ הכוונה לקרוב, נשאיר 0 = היום
    result=msg_dt + timedelta(days=days_ahead)
    return f'{result.day:02d}/{result.month:02d}'

def find_dates_in_message(text, msg_ddmm, msg_dt):
    """
    מחזיר את כל ה-(תאריך, שעות, עדיפות) שמופיעים בהודעה אחת.
    עדיפות גבוהה = אמין יותר.
    """
    found = []
    n = text
    # נרמול
    n = re.sub(r'ב(\d{1,2}/\d{1,2})', r'\1', n)
    n = re.sub(r'ל(\d{1,2}/\d{1,2})', r'\1', n)
    n = re.sub(r'בין\s+', '', n)
    n = re.sub(r'מתואם', 'תואם', n)
    n = re.sub(r'שיבוץ', 'שובץ', n)
    today_pat = '|'.join(re.escape(k) for k in TODAY_KW)

    # עדיפות 5: STATUS + תאריך + שעות
    for m in re.finditer(STATUS_KW + r'\s+(\d{1,2}/\d{1,2})\s+' + HOURS_RE, n):
        found.append((nd(m.group(1)), nh(m.group(2)), 5))
    # עדיפות 5: STATUS + שעות + תאריך
    for m in re.finditer(STATUS_KW + r'\s+' + HOURS_RE + r'\s+(\d{1,2}/\d{1,2})', n):
        found.append((nd(m.group(2)), nh(m.group(1)), 5))
    # עדיפות 4: היום/ירד להיום + שעות
    for m in re.finditer(r'(?:' + today_pat + r')\s*' + HOURS_RE, text, re.IGNORECASE):
        found.append((msg_ddmm, nh(m.group(1)), 4))
    for m in re.finditer(HOURS_RE + r'\s+(?:' + today_pat + r')', text, re.IGNORECASE):
        found.append((msg_ddmm, nh(m.group(1)), 4))
    # עדיפות 4: מחר + שעות
    for m in re.finditer(r'(?:התקנה\s+ל)?מחר\s+' + HOURS_RE, text):
        t=offset_ddmm(msg_ddmm,1)
        if t: found.append((t, nh(m.group(1)), 4))
    # עדיפות 4: STATUS/התקנה + יום בשבוע + שעות (שובץ ליום ראשון 16-18)
    for m in re.finditer(r'(?:' + STATUS_KW[3:-1] + r'|התקנה)\s+' + DAY_PAT + r'\s+' + HOURS_RE, n):
        wd = next_weekday_ddmm(msg_dt, m.group(1))
        if wd: found.append((wd, nh(m.group(2)), 4))
    # עדיפות 4: STATUS/התקנה + יום בשבוע + מילה + שעות (שובץ ליום ראשון צהריים 16-18)
    for m in re.finditer(r'(?:' + STATUS_KW[3:-1] + r'|התקנה)\s+' + DAY_PAT + r'\s+[א-ת]+\s+' + HOURS_RE, n):
        wd = next_weekday_ddmm(msg_dt, m.group(1))
        if wd: found.append((wd, nh(m.group(2)), 4))
    # עדיפות 3: יום בשבוע + שעות (ליום ראשון 8-10)
    for m in re.finditer(DAY_PAT + r'\s+' + HOURS_RE, n):
        wd = next_weekday_ddmm(msg_dt, m.group(1))
        if wd: found.append((wd, nh(m.group(2)), 3))
    # עדיפות 3: STATUS/התקנה + יום בשבוע (בלי שעות)
    for m in re.finditer(r'(?:' + STATUS_KW[3:-1] + r'|התקנה)\s+' + DAY_PAT, n):
        wd = next_weekday_ddmm(msg_dt, m.group(1))
        if wd: found.append((wd, '', 3))
    # עדיפות 2: תאריך + שעות גרידא
    for m in re.finditer(r'(\d{1,2}/\d{1,2})\s+' + HOURS_RE, n):
        found.append((nd(m.group(1)), nh(m.group(2)), 2))
    for m in re.finditer(HOURS_RE + r'\s+(\d{1,2}/\d{1,2})', n):
        found.append((nd(m.group(2)), nh(m.group(1)), 2))
    # עדיפות 1: תאריך לבד
    for m in re.finditer(r'(\d{1,2}/\d{1,2})', n):
        d = nd(m.group(1))
        found.append((d, '', 1))

    return found

def find_attachments(payload, attachments, message_id):
    fn=payload.get('filename','')
    if fn and fn.lower().endswith('.pdf'):
        aid=payload.get('body',{}).get('attachmentId')
        if aid: attachments.append({'filename':fn,'attachmentId':aid,'messageId':message_id})
    for part in payload.get('parts',[]): find_attachments(part,attachments,message_id)

def parse_sale_from_thread(messages, thread_id):
    full_text=''
    first_subject=''
    all_msgs=[]

    for i,m in enumerate(messages):
        hdrs={h['name']:h['value'] for h in m['payload']['headers']}
        subj=hdrs.get('Subject','')
        date_str=hdrs.get('Date','')
        date_full,date_ddmm,msg_dt=parse_email_date(date_str)
        if i==0: first_subject=subj
        snippet=m.get('snippet','')
        full_text+=subj+'\n'+snippet+'\n'
        all_msgs.append({'i':i,'snippet':snippet,'date_full':date_full,'date_ddmm':date_ddmm,'dt':msg_dt})

    # זהה thread מכירה
    sale_kw=['תואם','לאשר חוזה','ממירים','שובץ','להקים','ממיר','מוקלד','אישר','אושר','הוק',
             'לעדכן לקוח','ירד להיום','דאבל יס','דאבל סטינג','רק יס','סיבים']
    has_cid=bool(re.search(r'\b30\d{5}\b',full_text))
    if not any(k in full_text for k in sale_kw) and not has_cid:
        return None

    # שם מהנושא
    subj_clean=re.sub(r'^(Re|Fwd|FW|RE):\s*','',first_subject,flags=re.IGNORECASE).strip()
    subj_clean=re.sub(r'[-–].*','',subj_clean).strip()
    if not subj_clean or len(subj_clean)<2: return None

    # ביטול
    cancel_phrases=['לא להקים','התחרטה','לבטל','ביטול','מבטל','בוטל','ביטל','מבטלת','ביטלה','לא מעוניין','לא עבר שיקוף']
    is_cancelled=any(p in full_text for p in cancel_phrases)
    cancel_note=next((p for p in ['התחרטה','לא להקים','ביטל','לא מעוניין','לא עבר שיקוף'] if p in full_text),'')

    # מספר לקוח
    cid=''
    cm=re.search(r'(?:לקוח|מס\s*לקוח)\s+(\d{7})',full_text)
    if cm: cid=cm.group(1)
    else:
        am=re.search(r'\b(30\d{5})\b',full_text)
        if am: cid=am.group(1)

    # ── מנוע התאריך: עבור על מיילים לפי זמן, התאריך האחרון מנצח ──
    sorted_msgs=sorted(all_msgs,key=lambda x:x['dt'] or datetime.min)
    install_date,install_hours='',''
    first_date,first_hours='',''
    needs_review=False

    for msg in sorted_msgs:
        dates=find_dates_in_message(msg['snippet'],msg['date_ddmm'],msg['dt'])
        if dates:
            # קח את העדיפות הגבוהה ביותר בהודעה
            dates.sort(key=lambda x:x[2],reverse=True)
            best=dates[0]
            if not first_date:
                first_date,first_hours=best[0],best[1]
            install_date=best[0]
            if best[1]: install_hours=best[1]

    # שינוי תאריך
    has_change=bool(first_date and install_date and first_date!=install_date)
    change_note=f'שונה מ-{first_date} ל-{install_date}' if has_change else ''

    # סטטוס
    if is_cancelled:
        status='בוטל'
    elif any(w in full_text for w in ['מוקלד','אושר','אישר','לעדכן לקוח']):
        status='מוקלד + הוק' if 'הוק' in full_text else 'מוקלד'
    elif 'לאשר חוזה' in full_text and 'הוק' in full_text:
        status='לאשר חוזה + הוק'
    elif 'לאשר חוזה' in full_text:
        status='לאשר חוזה'
    elif 'שובץ' in full_text or install_date:
        status='שובץ'
    else:
        status='בהקדם'

    # אם אין תאריך התקנה כלל ולא בוטל — דרוש עריכה
    if not install_date and not is_cancelled:
        needs_review=True

    # ממירים
    mirrors=0
    mm=re.search(r'(\d+)\s*ממירים?',full_text)
    if mm: mirrors=int(mm.group(1))
    elif 'טריפל' in full_text: mirrors=3
    elif 'דאבל' in full_text: mirrors=2
    elif 'רק יס' in full_text or 'ממיר' in full_text: mirrors=1

    # חודש לפי תאריך התקנה
    install_month=''
    if install_date and install_date not in ('—','ממתין',''):
        parts=install_date.split('/')
        if len(parts)==2:
            try: install_month=f'{int(parts[1]):02d}/2026'
            except: pass

    sale_date=sorted_msgs[0]['date_full'] if sorted_msgs else ''
    if not install_month and sale_date:
        sp=sale_date.split('/')
        if len(sp)>=2:
            try: install_month=f'{int(sp[1]):02d}/2026'
            except: pass

    if not install_date: install_date='ממתין'

    return {'name':subj_clean,'customerId':cid,'saleDate':sale_date,
            'installDate':install_date,'installMonth':install_month,
            'hours':install_hours,'mirrors':mirrors,'status':status,
            'hasChange':has_change,'changeNote':change_note,
            'isCancelled':is_cancelled,'cancelNote':cancel_note,
            'needsReview':needs_review,'isToday':False,'isApril':False,'threadId':thread_id}

@app.route('/api/scan')
def scan():
    service=get_service()
    if not service:
        return jsonify({'error':'not_authenticated','auth_url':'https://sales-server-egdf.onrender.com/oauth/start'}),401
    try:
        sales,invoices=[],[]
        results=service.users().messages().list(userId='me',
            q='(from:oshrityes2901@gmail.com OR to:oshrityes2901@gmail.com OR from:oritapiro22@gmail.com OR to:oritapiro22@gmail.com OR from:avielv014@gmail.com OR to:avielv014@gmail.com) after:2026/1/1',
            maxResults=400).execute()

        seen=set()
        for msg in results.get('messages',[]):
            tid=msg['threadId']
            if tid in seen: continue
            seen.add(tid)
            try:
                thread=service.users().threads().get(userId='me',id=tid,format='full').execute()
                sale=parse_sale_from_thread(thread.get('messages',[]),tid)
                if sale: sales.append(sale)
            except: continue

        inv_results=service.users().messages().list(userId='me',
            q='(חשבונית OR invoice OR receipt OR morning.co OR render.com OR icount) has:attachment after:2026/1/1',
            maxResults=50).execute()
        for msg in inv_results.get('messages',[]):
            try:
                md=service.users().messages().get(userId='me',id=msg['id'],format='full').execute()
                hdrs={h['name']:h['value'] for h in md['payload']['headers']}
                subj=hdrs.get('Subject',''); sender=hdrs.get('From',''); date_str=hdrs.get('Date','')
                date_full,_,_=parse_email_date(date_str)
                month=f'{date_full[3:5]}/{date_full[6:10]}' if len(date_full)>=10 else ''
                biz=['ר.א.מ','ליד מנג','render','funnelly','stripe','morning','חשבונית ירוקה','atp','icount','קבלה']
                if not any(k.lower() in (subj+sender).lower() for k in biz): continue
                atts=[]; find_attachments(md['payload'],atts,msg['id'])
                it='income' if any(k in (subj+sender) for k in ['ר.א.מ','עמלות','may zalah','May zalah']) else ('subscription' if any(k in (subj+sender) for k in ['חשבונית ירוקה','morning']) else 'expense')
                nm=re.search(r'(\d{4,})',subj); name_m=re.match(r'^"?([^"<]+)',sender)
                invoices.append({'id':msg['id'],'date':date_full,'month':month,
                    'from':name_m.group(1).strip() if name_m else sender,'subject':subj,'description':subj[:50],
                    'invoiceNum':nm.group(1) if nm else '','type':it,
                    'gmailLink':f'https://mail.google.com/mail/u/0/#all/{msg["id"]}',
                    'hasAttachment':len(atts)>0,'attachments':atts})
            except: continue

        def d2i(d):
            try: p=d.split('/'); return int(p[2])*10000+int(p[1])*100+int(p[0])
            except: return 0
        sales.sort(key=lambda x:d2i(x.get('saleDate','')),reverse=True)
        return jsonify({'success':True,'sales':sales,'invoices':invoices})
    except Exception as e:
        return jsonify({'error':str(e)}),500

@app.route('/api/attachment/<message_id>/<attachment_id>')
def get_attachment(message_id,attachment_id):
    service=get_service()
    if not service: return jsonify({'error':'not_authenticated'}),401
    try:
        att=service.users().messages().attachments().get(userId='me',messageId=message_id,id=attachment_id).execute()
        data=base64.urlsafe_b64decode(att['data'])
        return Response(data,mimetype='application/pdf',headers={'Content-Disposition':'attachment; filename=invoice.pdf','Access-Control-Allow-Origin':'*'})
    except Exception as e: return jsonify({'error':str(e)}),500

@app.route('/api/send-to-accountant',methods=['POST'])
def send_to_accountant():
    service=get_service()
    if not service: return jsonify({'error':'not_authenticated'}),401
    data=request.json; month_name=data.get('monthName',''); invoices=data.get('invoices',[])
    try:
        msg=MIMEMultipart(); msg['To']='ei@eicpa.co.il'; msg['Subject']=f'חשבוניות {month_name} — ראובן חגג'
        body=f'שלום,\n\nמצורפות חשבוניות לחודש {month_name}.\n\n'
        for inv in invoices: body+=f'• {inv.get("from","")} — {inv.get("description","")}\n'
        body+='\nבברכה,\nרובי חגג'
        msg.attach(MIMEText(body,'plain','utf-8'))
        for inv in invoices:
            for att in inv.get('attachments',[]):
                if att.get('attachmentId'):
                    try:
                        ad=service.users().messages().attachments().get(userId='me',messageId=att['messageId'],id=att['attachmentId']).execute()
                        pdf=base64.urlsafe_b64decode(ad['data'])
                        part=MIMEBase('application','pdf'); part.set_payload(pdf); encoders.encode_base64(part)
                        part.add_header('Content-Disposition',f'attachment; filename="{att["filename"]}"')
                        msg.attach(part)
                    except: pass
        raw=base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(userId='me',body={'raw':raw}).execute()
        return jsonify({'success':True})
    except Exception as e: return jsonify({'error':str(e)}),500

if __name__=='__main__':
    app.run(host='0.0.0.0',port=int(os.environ.get('PORT',5000)),debug=False)
