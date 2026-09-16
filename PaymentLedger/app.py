import os,re
from pathlib import Path
from datetime import datetime
import cv2,numpy as np,pandas as pd,pytesseract
from flask import Flask,render_template,request,redirect,url_for,session,send_file,flash
from werkzeug.utils import secure_filename
from openpyxl import load_workbook,Workbook

BASE=Path(__file__).resolve().parent
EXCEL=BASE/'database'/'PaymentLedger.xlsx'
UPLOADS=BASE/'uploads'
OWNER_USERNAME=os.getenv('OWNER_USERNAME','owner')
OWNER_PASSWORD=os.getenv('OWNER_PASSWORD','change-this-password')
app=Flask(__name__); app.secret_key=os.getenv('FLASK_SECRET_KEY','change-this-secret-key')
COLS=['Record ID','Date','Time','Status','Paid To','UPI ID','Amount','Transaction ID','Debited From','UTR','What Did You Purchase?','Created At']

def init_db():
    EXCEL.parent.mkdir(exist_ok=True); UPLOADS.mkdir(exist_ok=True)
    if not EXCEL.exists():
        wb=Workbook(); ws=wb.active; ws.title='Payments'; ws.append(COLS); ws.freeze_panes='A2'; wb.save(EXCEL)

def ocr(path):
    img=cv2.imread(str(path));
    if img is None: raise ValueError('Could not read image')
    g=cv2.cvtColor(img,cv2.COLOR_BGR2GRAY); g=cv2.resize(g,None,fx=2,fy=2,interpolation=cv2.INTER_CUBIC)
    g=cv2.normalize(g,None,0,255,cv2.NORM_MINMAX); g=cv2.GaussianBlur(g,(3,3),0)
    d=pytesseract.image_to_data(g,config='--oem 3 --psm 6',lang='eng',output_type=pytesseract.Output.DATAFRAME)
    d=d.dropna(subset=['text']); d['text']=d['text'].astype(str).str.strip(); return d[d.text!=''].reset_index(drop=True)

def extract(d):
    w=d.text.tolist(); t=' '.join(w); r={c:'' for c in COLS[1:-1]}
    m=re.search(r'\b(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})\b',t,re.I)
    if m:r['Date']=' '.join(m.groups())
    m=re.search(r'\b\d{1,2}:\d{2}\b',t)
    if m:r['Time']=m.group()
    m=re.search(r'\b(successful|failed|pending)\b',t,re.I)
    if m:r['Status']=m.group().title()
    m=re.search(r'\b[A-Za-z0-9._-]+@[A-Za-z0-9._-]+\b',t)
    if m:r['UPI ID']=m.group()
    m=re.search(r'\bT[A-Za-z0-9]{10,}\b',t)
    if m:r['Transaction ID']=m.group()
    m=re.search(r'\bUTR[:\s-]*([0-9]{8,})\b',t,re.I)
    if m:r['UTR']=m.group(1)
    m=re.search(r'(?:₹|Rs\.?|INR|%)\s*([\d,]+(?:\.\d{1,2})?)',t,re.I)
    if m:r['Amount']='Rs '+m.group(1).replace(',','')
    try:
        i=next(i for i,x in enumerate(w) if x.lower()=='paid' and i+1<len(w) and w[i+1].lower()=='to')
        stop={'payment','details','verified','transaction','id','debited','from','utr'}; p=[]
        for x in w[i+2:]:
            if x.lower() in stop or '@' in x or re.match(r'^(?:₹|Rs|INR|%)',x,re.I): break
            p.append(x)
            if len(p)>=5: break
        r['Paid To']=' '.join(p)
    except StopIteration: pass
    try:
        i=next(i for i,x in enumerate(w) if x.lower()=='debited' and i+1<len(w) and w[i+1].lower()=='from')
        p=[]
        for x in w[i+2:]:
            if re.search(r'[Xx*]{4,}\d{2,}$',x) or re.search(r'\d{4}$',x): p.append(x); break
            if len(p)<2:p.append(x)
        r['Debited From']=' '.join(p)
    except StopIteration: pass
    return r

def save(r):
    init_db()
    try: df=pd.read_excel(EXCEL,sheet_name='Payments')
    except: df=pd.DataFrame(columns=COLS)
    for c in COLS:
        if c not in df: df[c]=''
    ids=pd.to_numeric(df['Record ID'],errors='coerce').dropna(); rid=int(ids.max())+1 if len(ids) else 1
    r['Record ID']=rid; r['Created At']=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    df=pd.concat([df[COLS],pd.DataFrame([{c:r.get(c,'') for c in COLS}])],ignore_index=True)
    with pd.ExcelWriter(EXCEL,engine='openpyxl') as x: df.to_excel(x,sheet_name='Payments',index=False)
    wb=load_workbook(EXCEL); ws=wb['Payments']; ws.freeze_panes='A2'; wb.save(EXCEL); return rid

@app.route('/',methods=['GET','POST'])
def user():
    if request.method=='POST':
        f=request.files.get('payment_image'); purchase=request.form.get('purchase','').strip()
        if not f or not f.filename: flash('Please upload a payment screenshot.','error'); return redirect(url_for('user'))
        if not purchase: flash('Please enter what you purchased.','error'); return redirect(url_for('user'))
        path=UPLOADS/(datetime.now().strftime('%Y%m%d%H%M%S%f')+'_'+secure_filename(f.filename)); f.save(path)
        try:
            r=extract(ocr(path)); r['What Did You Purchase?']=purchase; rid=save(r)
        except Exception as e: flash('OCR/database error: '+str(e),'error'); return redirect(url_for('user'))
        finally:
            if path.exists(): path.unlink()
        return render_template('user.html',submitted=True,record_id=rid)
    return render_template('user.html',submitted=False)

@app.route('/owner/login',methods=['GET','POST'])
def login():
    if request.method=='POST':
        if request.form.get('username')==OWNER_USERNAME and request.form.get('password')==OWNER_PASSWORD:
            session['owner']=True; return redirect(url_for('owner'))
        flash('Invalid owner credentials.','error')
    return render_template('login.html')

@app.route('/owner')
def owner():
    if not session.get('owner'): return redirect(url_for('login'))
    init_db(); df=pd.read_excel(EXCEL,sheet_name='Payments'); records=df.fillna('').to_dict('records') if not df.empty else []
    amounts=pd.to_numeric(df['Amount'].astype(str).str.replace(r'[^0-9.]','',regex=True),errors='coerce').fillna(0) if not df.empty else pd.Series(dtype=float)
    return render_template('owner.html',records=records,count=len(records),total=f'{amounts.sum():,.2f}')

@app.route('/owner/download')
def download():
    if not session.get('owner'): return redirect(url_for('login'))
    init_db(); return send_file(EXCEL,as_attachment=True,download_name='PaymentLedger.xlsx')

@app.route('/owner/logout')
def logout(): session.clear(); return redirect(url_for('login'))

if __name__=='__main__': init_db(); app.run(host='0.0.0.0',port=int(os.getenv('PORT',8501)),debug=False)
