import os, sqlite3, smtplib, threading, json, re, math
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from functools import wraps
from flask import (Flask, render_template, request, redirect, url_for,
                   session, flash, send_from_directory, abort, jsonify)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from dotenv import load_dotenv

load_dotenv()

# ── Config ────────────────────────────────────────────────────────────────────
SECRET_KEY     = os.getenv('SECRET_KEY', '')
GMAIL_USER     = os.getenv('GMAIL_USER', '')
GMAIL_PASSWORD = os.getenv('GMAIL_APP_PASSWORD', '')
GROQ_API_KEY   = os.getenv('GROQ_API_KEY', '')
GROQ_MODEL     = os.getenv('GROQ_MODEL', 'llama-3.3-70b-versatile')
FLASK_ENV      = os.getenv('FLASK_ENV', 'development')

if FLASK_ENV == 'production' and not SECRET_KEY:
    raise RuntimeError('SECRET_KEY must be set in production. Check your .env file.')

app = Flask(__name__)
app.secret_key = SECRET_KEY or 'dev-only-insecure-key-set-env'
app.config['WTF_CSRF_ENABLED']     = True
app.config['WTF_CSRF_SECRET_KEY']  = (SECRET_KEY or 'csrf-dev-key') + '-csrf'
app.config['MAX_CONTENT_LENGTH']   = 15 * 1024 * 1024

UPLOAD_FOLDER      = 'uploads'
ALLOWED_EXTENSIONS = {'pdf', 'doc', 'docx', 'png', 'jpg', 'jpeg', 'xlsx', 'xls'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
DATABASE = 'portal.db'
MAX_EXTRACTED_BYTES = 50_000  # 50 KB cap on stored text

# ── Optional imports ──────────────────────────────────────────────────────────
try:
    import pdfplumber; PDF_OK = True
except: PDF_OK = False
try:
    from docx import Document as DocxDoc; DOCX_OK = True
except: DOCX_OK = False
try:
    import openpyxl; XLSX_OK = True
except: XLSX_OK = False
try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity; SKLEARN_OK = True
except: SKLEARN_OK = False
try:
    from groq import Groq as GroqClient
    groq_client = GroqClient(api_key=GROQ_API_KEY) if GROQ_API_KEY else None
    GROQ_OK = bool(GROQ_API_KEY)
except: groq_client = None; GROQ_OK = False

# ── Rate limiting ─────────────────────────────────────────────────────────────
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
limiter = Limiter(key_func=get_remote_address, app=app,
                  default_limits=[], storage_uri='memory://')

# ── CSRF ──────────────────────────────────────────────────────────────────────
from flask_wtf.csrf import CSRFProtect, generate_csrf, CSRFError
csrf = CSRFProtect(app)

@app.after_request
def inject_csrf_cookie(response):
    response.set_cookie('csrf_token', generate_csrf())
    return response

@app.errorhandler(CSRFError)
def csrf_error(e):
    flash('Session expired or invalid request. Please try again.', 'error')
    return redirect(request.referrer or url_for('login'))

# ── APScheduler (guard against double-start in debug reloader) ────────────────
from apscheduler.schedulers.background import BackgroundScheduler
_scheduler_started = False
scheduler = BackgroundScheduler(daemon=True)

# ── Database ──────────────────────────────────────────────────────────────────
def get_db():
    c = sqlite3.connect(DATABASE)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA foreign_keys = ON")
    return c

def init_db():
    with get_db() as db:
        db.executescript('''
            CREATE TABLE IF NOT EXISTS users (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                fullname        TEXT NOT NULL,
                email           TEXT UNIQUE NOT NULL,
                password        TEXT NOT NULL,
                role            TEXT NOT NULL DEFAULT 'student',
                matric          TEXT UNIQUE,
                phone           TEXT,
                department_id   INTEGER REFERENCES departments(id) ON DELETE SET NULL,
                faculty         TEXT,
                level           TEXT,
                program         TEXT,
                gender          TEXT,
                session_year    TEXT,
                semester        TEXT,
                state_of_origin TEXT,
                is_superadmin   INTEGER DEFAULT 0,
                created_at      TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS departments (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                name       TEXT UNIQUE NOT NULL,
                faculty    TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS lecturer_departments (
                lecturer_id   INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                department_id INTEGER NOT NULL REFERENCES departments(id) ON DELETE CASCADE,
                PRIMARY KEY (lecturer_id, department_id)
            );
            CREATE TABLE IF NOT EXISTS student_courses (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id     INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                course_code TEXT NOT NULL,
                course_name TEXT NOT NULL,
                credit_unit INTEGER DEFAULT 0,
                semester    TEXT
            );
            CREATE TABLE IF NOT EXISTS assignments (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                title           TEXT NOT NULL,
                description     TEXT,
                course_code     TEXT NOT NULL,
                course_name     TEXT NOT NULL,
                deadline        TEXT NOT NULL,
                department_id   INTEGER REFERENCES departments(id) ON DELETE SET NULL,
                reminder_sent   INTEGER DEFAULT 0,
                created_by      INTEGER NOT NULL REFERENCES users(id),
                created_at      TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS submissions (
                id                INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id           INTEGER NOT NULL REFERENCES users(id),
                assignment_id     INTEGER NOT NULL REFERENCES assignments(id),
                filename          TEXT NOT NULL,
                original_filename TEXT NOT NULL,
                file_ext          TEXT NOT NULL,
                file_size         INTEGER,
                is_late           INTEGER DEFAULT 0,
                extracted_text    TEXT,
                submitted_at      TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(user_id, assignment_id)
            );
            CREATE TABLE IF NOT EXISTS plagiarism_scores (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                assignment_id INTEGER NOT NULL REFERENCES assignments(id) ON DELETE CASCADE,
                sub_id_a      INTEGER NOT NULL REFERENCES submissions(id) ON DELETE CASCADE,
                sub_id_b      INTEGER NOT NULL REFERENCES submissions(id) ON DELETE CASCADE,
                score         REAL NOT NULL,
                checked_at    TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS email_log (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                event       TEXT NOT NULL,
                recipients  INTEGER NOT NULL,
                detail      TEXT,
                sent_at     TEXT DEFAULT CURRENT_TIMESTAMP
            );
        ''')
        # Safe migrations for existing DBs
        _safe_add_columns(db, 'users', [
            ('matric','TEXT'),('session_year','TEXT'),('semester','TEXT'),
            ('state_of_origin','TEXT'),('department_id','INTEGER'),
            ('faculty','TEXT'),('level','TEXT'),('program','TEXT'),
            ('phone','TEXT'),('gender','TEXT'),('is_superadmin','INTEGER DEFAULT 0')
        ])
        _safe_add_columns(db, 'assignments', [('reminder_sent','INTEGER DEFAULT 0')])
        _safe_add_columns(db, 'submissions',  [('extracted_text','TEXT')])

        # First admin = superadmin
        if not db.execute("SELECT id FROM users WHERE email='admin@imap.edu.ng'").fetchone():
            db.execute("INSERT INTO users (fullname,email,password,role,is_superadmin) VALUES (?,?,?,?,1)",
                ('Portal Admin','admin@imap.edu.ng',
                 generate_password_hash('Admin@1234'),'admin'))

        # Sample departments
        for name, fac in [
            ('Computer Science','School of Science and Technology'),
            ('Software and Web Development','School of Science and Technology'),
            ('Electrical Engineering','School of Engineering'),
            ('Business Administration','School of Business and Management'),
            ('Accountancy','School of Business and Management'),
            ('Mass Communication','School of Communication and Arts'),
        ]:
            if not db.execute("SELECT id FROM departments WHERE name=?",(name,)).fetchone():
                db.execute("INSERT INTO departments (name,faculty) VALUES (?,?)",(name,fac))
        db.commit()

def _safe_add_columns(db, table, cols):
    existing = {r[1] for r in db.execute(f"PRAGMA table_info({table})")}
    for col, typ in cols:
        if col not in existing:
            try: db.execute(f"ALTER TABLE {table} ADD COLUMN {col} {typ}")
            except: pass

init_db()

# ── Helpers ───────────────────────────────────────────────────────────────────
def allowed_file(f):
    return '.' in f and f.rsplit('.',1)[1].lower() in ALLOWED_EXTENSIONS

def file_icon(ext):
    return {'pdf':'pdf','doc':'word','docx':'word','xlsx':'excel','xls':'excel',
            'png':'image','jpg':'image','jpeg':'image'}.get(ext.lower(),'file')

def human_size(s):
    if not s: return '—'
    if s < 1024: return f'{s} B'
    if s < 1048576: return f'{s/1024:.1f} KB'
    return f'{s/1048576:.1f} MB'

def profile_pct(u):
    fields = ['matric','phone','department_id','level','program','gender','session_year','semester']
    return int(sum(1 for f in fields if u[f]) / len(fields) * 100)

def fmt_dt(dt_str):
    if not dt_str: return '—'
    try:
        return datetime.strptime(dt_str[:16], '%Y-%m-%dT%H:%M').strftime('%d %b %Y, %I:%M %p')
    except: return dt_str[:16]

def now_str(): return datetime.now().strftime('%Y-%m-%dT%H:%M')

def is_past(deadline): return now_str() > deadline[:16]

app.jinja_env.globals.update(
    file_icon=file_icon, human_size=human_size, profile_pct=profile_pct,
    fmt_dt=fmt_dt, now=datetime.now, GROQ_OK=GROQ_OK, is_past=is_past
)

# ── Email ─────────────────────────────────────────────────────────────────────
def _html_email(heading, rows, note=None):
    body = ''.join(f'<p style="margin:0 0 10px;color:#374151;font-size:14px;line-height:1.7">{r}</p>' for r in rows)
    footer_note = f'<p style="font-size:12px;color:#9ca3af;margin-top:12px">{note}</p>' if note else ''
    return f'''<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f3f4f6;font-family:Arial,sans-serif">
<div style="max-width:560px;margin:32px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 16px rgba(0,0,0,.08)">
  <div style="background:linear-gradient(135deg,#4f46e5,#2563eb);padding:24px 28px">
    <div style="font-size:18px;font-weight:800;color:#fff">TJs-Code Assignment Portal</div>
    <div style="font-size:11px;color:rgba(255,255,255,.75);margin-top:2px">Isa Mustapha Agwai I Polytechnic, Lafia</div>
  </div>
  <div style="padding:24px 28px">
    <h2 style="margin:0 0 14px;font-size:17px;color:#111827">{heading}</h2>
    {body}{footer_note}
  </div>
  <div style="background:#f9fafb;border-top:1px solid #e5e7eb;padding:14px 28px;font-size:11px;color:#9ca3af;text-align:center">
    &copy; {datetime.now().year} TJs-Code Assignment Portal &mdash; IMAP, Lafia
  </div>
</div></body></html>'''

def send_email_bg(to_list, subject, html, event=''):
    def _do():
        if not GMAIL_USER or not GMAIL_PASSWORD:
            print('[Email] No credentials'); return
        sent = 0
        for addr in to_list:
            try:
                msg = MIMEMultipart('alternative')
                msg['Subject'] = subject
                msg['From']    = f'TJs-Code Portal <{GMAIL_USER}>'
                msg['To']      = addr
                msg.attach(MIMEText(html, 'html'))
                with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
                    s.login(GMAIL_USER, GMAIL_PASSWORD)
                    s.sendmail(GMAIL_USER, addr, msg.as_string())
                sent += 1
            except Exception as e:
                print(f'[Email] {addr}: {e}')
        if event:
            with get_db() as db:
                db.execute("INSERT INTO email_log (event,recipients,detail) VALUES (?,?,?)",
                           (event, sent, subject)); db.commit()
        print(f'[Email] Sent {sent}/{len(to_list)}')
    threading.Thread(target=_do, daemon=True).start()

def get_dept_student_emails(dept_id, exclude_submitted=None):
    with get_db() as db:
        if dept_id:
            rows = db.execute("SELECT email FROM users WHERE role='student' AND department_id=?",(dept_id,)).fetchall()
        else:
            rows = db.execute("SELECT email FROM users WHERE role='student'").fetchall()
        emails = [r['email'] for r in rows]
        if exclude_submitted:
            submitted = {r['email'] for r in db.execute(
                "SELECT u.email FROM submissions s JOIN users u ON s.user_id=u.id WHERE s.assignment_id=?",
                (exclude_submitted,)).fetchall()}
            emails = [e for e in emails if e not in submitted]
    return emails

def notify_new_assignment(asgn, dept_name, emails):
    if not emails: return
    html = _html_email('New Assignment Posted', [
        f'A new assignment has been posted for <strong>{dept_name}</strong>.',
        f'<strong>Title:</strong> {asgn["title"]}',
        f'<strong>Course:</strong> {asgn["course_code"]} — {asgn["course_name"]}',
        f'<strong>Deadline:</strong> <span style="color:#dc2626;font-weight:700">{fmt_dt(asgn["deadline"])}</span>',
        asgn["description"] or '',
        'Log in to the portal to submit your work before the deadline.',
    ])
    send_email_bg(emails, f'New Assignment: {asgn["title"]}', html, 'new_assignment')

def notify_reminder(asgn, dept_name, emails):
    if not emails: return
    html = _html_email('Deadline Reminder — 24 Hours Left', [
        'This is a reminder that your assignment deadline is <strong>in less than 24 hours</strong>.',
        f'<strong>Assignment:</strong> {asgn["title"]}',
        f'<strong>Course:</strong> {asgn["course_code"]}',
        f'<strong>Deadline:</strong> <span style="color:#dc2626;font-weight:700">{fmt_dt(asgn["deadline"])}</span>',
        'If you have not submitted yet, please log in immediately.',
    ], note='You received this because you have not submitted this assignment yet.')
    send_email_bg(emails, f'Reminder: "{asgn["title"]}" due in 24hrs!', html, 'reminder')

def notify_plagiarism(email, name, asgn_title, score):
    html = _html_email('Similarity Warning on Your Submission', [
        f'Dear {name},',
        f'Your submission for <strong>{asgn_title}</strong> has been flagged for high similarity with another submission.',
        f'<strong>Similarity Score: <span style="color:#dc2626">{score:.0f}%</span></strong>',
        'A score above 70% may indicate copied content. Please review your work and consider re-submitting before the deadline.',
        'Academic integrity is taken seriously at IMAP.',
    ])
    send_email_bg([email], f'Similarity Warning — {asgn_title}', html, 'plagiarism_warning')

def notify_missing_reminder(email, name, asgn_title, deadline):
    html = _html_email('Assignment Submission Reminder', [
        f'Dear {name},',
        f'You have not yet submitted your assignment: <strong>{asgn_title}</strong>.',
        f'<strong>Deadline:</strong> <span style="color:#dc2626;font-weight:700">{fmt_dt(deadline)}</span>',
        'Please log in and submit as soon as possible.',
    ])
    send_email_bg([email], f'Reminder: Submit "{asgn_title}"', html, 'missing_reminder')

# ── Scheduler ─────────────────────────────────────────────────────────────────
def check_deadline_reminders():
    with app.app_context():
        try:
            now = now_str()
            in24 = (datetime.now() + timedelta(hours=24)).strftime('%Y-%m-%dT%H:%M')
            with get_db() as db:
                asgns = db.execute(
                    "SELECT * FROM assignments WHERE reminder_sent=0 AND deadline>? AND deadline<=?",(now,in24)
                ).fetchall()
                for a in asgns:
                    emails = get_dept_student_emails(a['department_id'], exclude_submitted=a['id'])
                    dept = db.execute("SELECT name FROM departments WHERE id=?",(a['department_id'],)).fetchone()
                    notify_reminder(a, dept['name'] if dept else 'All', emails)
                    db.execute("UPDATE assignments SET reminder_sent=1 WHERE id=?",(a['id'],))
                db.commit()
        except Exception as e:
            print(f'[Scheduler] {e}')

# Only start scheduler once (avoid double-start in Flask debug mode)
if os.environ.get('WERKZEUG_RUN_MAIN') != 'false':
    scheduler.add_job(check_deadline_reminders, 'interval', hours=1, id='reminder')
    scheduler.start()

# ── Text extraction ───────────────────────────────────────────────────────────
def extract_text(filepath, ext):
    ext = ext.lower()
    try:
        if ext == 'pdf' and PDF_OK:
            with pdfplumber.open(filepath) as pdf:
                return '\n'.join(p.extract_text() or '' for p in pdf.pages)
        if ext in ('doc','docx') and DOCX_OK:
            return '\n'.join(p.text for p in DocxDoc(filepath).paragraphs if p.text.strip())
        if ext in ('xlsx','xls') and XLSX_OK:
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            return '\n'.join(
                ' '.join(str(c) for c in row if c is not None)
                for ws in wb.worksheets for row in ws.iter_rows(values_only=True)
            )
    except Exception as e:
        print(f'[Extract] {e}')
    return ''

def extract_and_store(sub_id, filepath, ext):
    def _do():
        text = extract_text(filepath, ext)
        if text.strip():
            text = text[:MAX_EXTRACTED_BYTES]
            with get_db() as db:
                db.execute("UPDATE submissions SET extracted_text=? WHERE id=?",(text, sub_id))
                db.commit()
    threading.Thread(target=_do, daemon=True).start()

# ── Plagiarism ────────────────────────────────────────────────────────────────
def run_plagiarism_check(assignment_id):
    with get_db() as db:
        subs = db.execute(
            "SELECT s.id,s.extracted_text,u.fullname,u.email,u.matric "
            "FROM submissions s JOIN users u ON s.user_id=u.id "
            "WHERE s.assignment_id=? AND s.extracted_text IS NOT NULL AND s.extracted_text!=''",
            (assignment_id,)).fetchall()
    if len(subs) < 2: return []
    ids=[s['id'] for s in subs]; texts=[s['extracted_text'] for s in subs]
    names=[s['fullname'] for s in subs]; emails=[s['email'] for s in subs]
    results=[]
    try:
        if SKLEARN_OK:
            vec=TfidfVectorizer(stop_words='english',max_features=5000)
            sim=cosine_similarity(vec.fit_transform(texts))
        else:
            def jaccard(a,b):
                wa=set(re.findall(r'\w+',a.lower())); wb=set(re.findall(r'\w+',b.lower()))
                return len(wa&wb)/len(wa|wb) if wa|wb else 0
            n=len(texts)
            import numpy as np
            sim=np.array([[jaccard(texts[i],texts[j]) for j in range(n)] for i in range(n)])
        with get_db() as db:
            db.execute("DELETE FROM plagiarism_scores WHERE assignment_id=?",(assignment_id,))
            for i in range(len(subs)):
                for j in range(i+1,len(subs)):
                    score=round(float(sim[i][j])*100,1)
                    db.execute("INSERT INTO plagiarism_scores (assignment_id,sub_id_a,sub_id_b,score) VALUES (?,?,?,?)",
                               (assignment_id,ids[i],ids[j],score))
                    results.append({'sub_a':ids[i],'name_a':names[i],'email_a':emails[i],
                                    'sub_b':ids[j],'name_b':names[j],'email_b':emails[j],'score':score})
            db.commit()
        asgn_title=''
        with get_db() as db:
            row=db.execute("SELECT title FROM assignments WHERE id=?",(assignment_id,)).fetchone()
            if row: asgn_title=row['title']
        for r in results:
            if r['score']>=70:
                for em,nm in [(r['email_a'],r['name_a']),(r['email_b'],r['name_b'])]:
                    notify_plagiarism(em,nm,asgn_title,r['score'])
    except Exception as e:
        print(f'[Plagiarism] {e}')
    return sorted(results, key=lambda x: x['score'], reverse=True)

# ── Groq AI ───────────────────────────────────────────────────────────────────
def call_groq(messages, system):
    if not GROQ_OK or not groq_client:
        return 'AI is not available. Please check your GROQ_API_KEY in the .env file.'
    try:
        resp = groq_client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[{'role':'system','content':system}]+messages,
            max_tokens=1024, temperature=0.7)
        return resp.choices[0].message.content
    except Exception as e:
        return f'AI error: {str(e)}'

def build_student_context(uid):
    with get_db() as db:
        u = db.execute(
            "SELECT u.*,d.name as dept_name FROM users u "
            "LEFT JOIN departments d ON u.department_id=d.id WHERE u.id=?",(uid,)).fetchone()
        subs = db.execute(
            "SELECT s.id,s.is_late,s.submitted_at,s.file_ext,s.original_filename,"
            "a.title,a.course_code,a.course_name,a.deadline "
            "FROM submissions s JOIN assignments a ON s.assignment_id=a.id "
            "WHERE s.user_id=? ORDER BY s.submitted_at DESC",(uid,)).fetchall()
        available = db.execute(
            "SELECT a.id,a.title,a.course_code,a.course_name,a.deadline,a.description,"
            "(SELECT COUNT(*) FROM submissions WHERE assignment_id=a.id AND user_id=?) as done "
            "FROM assignments a "
            "WHERE a.department_id IS NULL OR a.department_id=? ORDER BY a.deadline",(uid,u['department_id'] or 0)).fetchall()
        flagged = db.execute(
            "SELECT p.score,a.title FROM plagiarism_scores p "
            "JOIN submissions s ON (s.id=p.sub_id_a OR s.id=p.sub_id_b) "
            "JOIN assignments a ON a.id=p.assignment_id "
            "WHERE s.user_id=? AND p.score>=70",(uid,)).fetchall()
        courses = db.execute(
            "SELECT course_code,course_name,credit_unit,semester FROM student_courses WHERE user_id=?",(uid,)).fetchall()

    lines = []
    lines.append("=== STUDENT PROFILE ===")
    lines.append(f"Name: {u['fullname']}")
    lines.append(f"Matric Number: {u['matric'] or 'NOT SET'}")
    lines.append(f"Email: {u['email']}")
    lines.append(f"Department: {u['dept_name'] or 'NOT SET'}")
    lines.append(f"Faculty: {u['faculty'] or 'not set'}")
    lines.append(f"Level: {u['level'] or 'not set'}")
    lines.append(f"Programme: {u['program'] or 'not set'}")
    lines.append(f"Academic Session: {u['session_year'] or 'not set'}")
    lines.append(f"Current Semester: {u['semester'] or 'not set'}")

    if courses:
        lines.append("")
        lines.append("=== REGISTERED COURSES ===")
        for c in courses:
            lines.append(f"- {c['course_code']}: {c['course_name']} ({c['credit_unit']} units, {c['semester'] or 'semester not set'})")

    lines.append("")
    lines.append("=== ASSIGNMENTS ===")
    pending  = [a for a in available if not a['done']]
    done_list= [a for a in available if a['done']]

    if not available:
        lines.append("No assignments posted for your department yet.")
    else:
        if pending:
            lines.append(f"PENDING ({len(pending)} not yet submitted):")
            for a in pending:
                od = " *** OVERDUE ***" if is_past(a['deadline']) else ""
                lines.append("  - " + a['title'] + " | Course: " + a['course_code'] + " (" + a['course_name'] + ")")
                lines.append(f"    Deadline: {fmt_dt(a['deadline'])}{od}")
                if a['description']:
                    lines.append(f"    Instructions: {a['description'][:200]}")
        if done_list:
            lines.append(f"SUBMITTED ({len(done_list)}):")
            for a in done_list:
                # find matching submission
                match = next((s for s in subs if s['title']==a['title']), None)
                late_tag = " [LATE]" if (match and match['is_late']) else " [On Time]"
                lines.append("  - " + a['title'] + " | " + a['course_code'] + late_tag)

    if subs:
        lines.append("")
        lines.append("=== SUBMISSION HISTORY ===")
        for s in subs:
            status = "LATE" if s['is_late'] else "On Time"
            submitted_at = s['submitted_at'][:16].replace('T',' ') if s['submitted_at'] else 'unknown'
            lines.append("- " + s['title'] + " (" + s['course_code'] + ") | File: " + s['original_filename'] + " | Status: " + status + " | Submitted: " + submitted_at)

    if flagged:
        lines.append("")
        lines.append("=== PLAGIARISM WARNINGS ===")
        for f in flagged:
            lines.append("- " + f['title'] + ": " + str(round(f['score'])) + "% similarity flagged. Student should review and re-submit.")

    return chr(10).join(lines)

def build_admin_context(uid):
    with get_db() as db:
        u = db.execute("SELECT * FROM users WHERE id=?",(uid,)).fetchone()
        is_super = bool(u['is_superadmin'])

        if is_super:
            dept_ids   = None
            dept_names = 'All departments (superadmin)'
            assignments = db.execute(
                "SELECT a.*,d.name as dept_name,"
                "(SELECT COUNT(*) FROM submissions WHERE assignment_id=a.id) as sub_count,"
                "(SELECT COUNT(*) FROM submissions WHERE assignment_id=a.id AND is_late=1) as late_count,"
                "(SELECT COUNT(*) FROM users WHERE role='student' AND department_id=a.department_id) as dept_students "
                "FROM assignments a LEFT JOIN departments d ON a.department_id=d.id "
                "ORDER BY a.deadline DESC LIMIT 10").fetchall()
            recent_subs = db.execute(
                "SELECT s.*,u.fullname,u.matric,a.title,a.course_code "
                "FROM submissions s JOIN users u ON s.user_id=u.id "
                "JOIN assignments a ON s.assignment_id=a.id "
                "ORDER BY s.submitted_at DESC LIMIT 10").fetchall()
            total  = db.execute("SELECT COUNT(*) FROM submissions").fetchone()[0]
            late   = db.execute("SELECT COUNT(*) FROM submissions WHERE is_late=1").fetchone()[0]
            stud_c = db.execute("SELECT COUNT(*) FROM users WHERE role='student'").fetchone()[0]
            flagged_pairs = db.execute(
                "SELECT p.score,ua.fullname as name_a,ub.fullname as name_b,a.title "
                "FROM plagiarism_scores p "
                "JOIN submissions sa ON p.sub_id_a=sa.id JOIN users ua ON sa.user_id=ua.id "
                "JOIN submissions sb ON p.sub_id_b=sb.id JOIN users ub ON sb.user_id=ub.id "
                "JOIN assignments a ON p.assignment_id=a.id "
                "WHERE p.score>=70 ORDER BY p.score DESC LIMIT 5").fetchall()
        else:
            my_depts = db.execute(
                "SELECT d.id,d.name FROM lecturer_departments ld "
                "JOIN departments d ON ld.department_id=d.id WHERE ld.lecturer_id=?",(uid,)).fetchall()
            dept_ids   = [d['id'] for d in my_depts]
            dept_names = ', '.join(d['name'] for d in my_depts) or 'none assigned'
            if dept_ids:
                ph = ','.join('?'*len(dept_ids))
                assignments = db.execute(
                    f"SELECT a.*,d.name as dept_name,"
                    "(SELECT COUNT(*) FROM submissions WHERE assignment_id=a.id) as sub_count,"
                    "(SELECT COUNT(*) FROM submissions WHERE assignment_id=a.id AND is_late=1) as late_count "
                    f"FROM assignments a LEFT JOIN departments d ON a.department_id=d.id "
                    f"WHERE a.department_id IN ({ph}) OR a.created_by=? "
                    "ORDER BY a.deadline DESC LIMIT 10",dept_ids+[uid]).fetchall()
                recent_subs = db.execute(
                    f"SELECT s.*,u.fullname,u.matric,a.title,a.course_code "
                    "FROM submissions s JOIN users u ON s.user_id=u.id "
                    f"JOIN assignments a ON s.assignment_id=a.id "
                    f"WHERE a.department_id IN ({ph}) "
                    "ORDER BY s.submitted_at DESC LIMIT 10",dept_ids).fetchall()
                total  = db.execute(f"SELECT COUNT(*) FROM submissions s JOIN assignments a ON s.assignment_id=a.id WHERE a.department_id IN ({ph})",dept_ids).fetchone()[0]
                late   = db.execute(f"SELECT COUNT(*) FROM submissions s JOIN assignments a ON s.assignment_id=a.id WHERE a.department_id IN ({ph}) AND s.is_late=1",dept_ids).fetchone()[0]
                stud_c = db.execute(f"SELECT COUNT(*) FROM users WHERE role='student' AND department_id IN ({ph})",dept_ids).fetchone()[0]
                flagged_pairs = []
            else:
                assignments=recent_subs=[]; total=late=stud_c=0; flagged_pairs=[]

    lines = []
    lines.append("=== ADMIN / LECTURER PROFILE ===")
    lines.append(f"Name: {u['fullname']}")
    lines.append(f"Role: {'Superadmin (full access)' if is_super else 'Lecturer'}")
    lines.append(f"Departments: {dept_names}")
    lines.append(f"Total submissions: {total} | Late: {late} | Students: {stud_c}")

    if assignments:
        lines.append("")
        lines.append("=== ASSIGNMENTS (latest 10) ===")
        for a in assignments:
            status = "CLOSED" if is_past(a['deadline']) else "OPEN"
            missing = ""
            if not is_past(a['deadline']):
                with get_db() as db:
                    dept_total = db.execute(
                        "SELECT COUNT(*) FROM users WHERE role='student' AND department_id=?",
                        (a['department_id'],)).fetchone()[0] if a['department_id'] else 0
                missing_count = dept_total - a['sub_count']
                if missing_count > 0:
                    missing = f" | {missing_count} student(s) NOT YET SUBMITTED"
            lines.append("- " + a['title'] + " [" + status + "]")
            lines.append("  Course: " + a['course_code'] + " (" + (a['course_name'] if 'course_name' in a.keys() else '') + ")")
            lines.append("  Dept: " + (a['dept_name'] or 'All') + " | Deadline: " + fmt_dt(a['deadline']))
            lines.append("  Submissions: " + str(a['sub_count']) + " | Late: " + str(a['late_count']) + missing)

    if recent_subs:
        lines.append("")
        lines.append("=== RECENT SUBMISSIONS (latest 10) ===")
        for s in recent_subs:
            status  = "LATE" if s['is_late'] else "On Time"
            matric  = s['matric'] or 'no matric'
            subtime = s['submitted_at'][:16].replace('T',' ') if s['submitted_at'] else '?'
            lines.append("- " + s['fullname'] + " (" + matric + "): " + s['title'] + " | " + s['course_code'] + " | " + status + " | " + subtime)

    if flagged_pairs:
        lines.append("")
        lines.append("=== PLAGIARISM FLAGS (>=70%) ===")
        for p in flagged_pairs:
            lines.append("- " + p['name_a'] + " & " + p['name_b'] + ": " + str(round(p['score'])) + "% similarity on " + p['title'])

    return chr(10).join(lines)

# ── Auth decorators ───────────────────────────────────────────────────────────
def login_required(f):
    @wraps(f)
    def d(*a,**k):
        if 'user_id' not in session:
            flash('Please log in to continue.','info'); return redirect(url_for('login'))
        return f(*a,**k)
    return d

def admin_required(f):
    @wraps(f)
    def d(*a,**k):
        if session.get('role') != 'admin': abort(403)
        return f(*a,**k)
    return d

def superadmin_required(f):
    @wraps(f)
    def d(*a,**k):
        if session.get('role') != 'admin' or not session.get('is_superadmin'): abort(403)
        return f(*a,**k)
    return d

def student_required(f):
    @wraps(f)
    def d(*a,**k):
        if session.get('role') != 'student': abort(403)
        return f(*a,**k)
    return d

# ── Auth routes ───────────────────────────────────────────────────────────────
@app.route('/', methods=['GET','POST'])
@limiter.limit("20 per minute")
def login():
    if 'user_id' in session:
        return redirect(url_for('admin_dashboard' if session['role']=='admin' else 'dashboard'))
    if request.method == 'POST':
        ident = request.form.get('identifier','').strip()
        pw    = request.form.get('password','')
        with get_db() as db:
            u = db.execute("SELECT * FROM users WHERE (email=? OR matric=?) AND role='student'",(ident,ident)).fetchone()
        if u and check_password_hash(u['password'],pw):
            session.update(user_id=u['id'],fullname=u['fullname'],role=u['role'],
                           email=u['email'],matric=u['matric'])
            return redirect(url_for('dashboard'))
        flash('Invalid credentials.','error')
    return render_template('login.html')

@app.route('/admin-login', methods=['GET','POST'])
@limiter.limit("5 per minute")
def admin_login():
    if 'user_id' in session:
        return redirect(url_for('admin_dashboard' if session['role']=='admin' else 'dashboard'))
    if request.method == 'POST':
        email = request.form.get('email','').strip().lower()
        pw    = request.form.get('password','')
        with get_db() as db:
            u = db.execute("SELECT * FROM users WHERE email=? AND role='admin'",(email,)).fetchone()
        if u and check_password_hash(u['password'],pw):
            session.update(user_id=u['id'],fullname=u['fullname'],role=u['role'],
                           email=u['email'],is_superadmin=bool(u['is_superadmin']))
            return redirect(url_for('admin_dashboard'))
        flash('Invalid admin credentials.','error')
    return render_template('admin_login.html')

@app.route('/register', methods=['GET','POST'])
def register():
    if 'user_id' in session: return redirect(url_for('dashboard'))
    with get_db() as db:
        depts = db.execute("SELECT * FROM departments ORDER BY name").fetchall()
    if request.method == 'POST':
        fullname = request.form.get('fullname','').strip()
        email    = request.form.get('email','').strip().lower()
        matric   = request.form.get('matric','').strip().upper()
        pw       = request.form.get('password','')
        confirm  = request.form.get('confirm','')
        dept_id  = request.form.get('department_id','') or None
        errors=[]
        if not fullname: errors.append('Full name is required.')
        if not email:    errors.append('Email is required.')
        if not matric:   errors.append('Matric number is required.')
        if len(pw)<6:    errors.append('Password must be at least 6 characters.')
        if pw!=confirm:  errors.append('Passwords do not match.')
        if errors:
            for e in errors: flash(e,'error')
        else:
            try:
                with get_db() as db:
                    db.execute("INSERT INTO users (fullname,email,password,matric,department_id,role) VALUES (?,?,?,?,?,'student')",
                               (fullname,email,generate_password_hash(pw),matric,dept_id))
                    db.commit()
                flash('Account created. You can now log in.','success')
                return redirect(url_for('login'))
            except sqlite3.IntegrityError:
                flash('Email or matric number already exists.','error')
    return render_template('register.html', depts=depts)

@app.route('/logout')
def logout():
    session.clear(); return redirect(url_for('login'))

# ── Student: Dashboard ────────────────────────────────────────────────────────
@app.route('/dashboard')
@login_required
@student_required
def dashboard():
    uid = session['user_id']
    with get_db() as db:
        user = db.execute("SELECT u.*,d.name as dept_name FROM users u LEFT JOIN departments d ON u.department_id=d.id WHERE u.id=?",(uid,)).fetchone()
        subs = db.execute(
            "SELECT s.*,a.title,a.course_code,a.course_name,a.deadline FROM submissions s "
            "JOIN assignments a ON s.assignment_id=a.id WHERE s.user_id=? ORDER BY s.submitted_at DESC",(uid,)).fetchall()
        available = db.execute(
            "SELECT a.*,d.name as dept_name,"
            "(SELECT COUNT(*) FROM submissions WHERE assignment_id=a.id AND user_id=?) as submitted "
            "FROM assignments a LEFT JOIN departments d ON a.department_id=d.id "
            "WHERE a.department_id IS NULL OR a.department_id=? ORDER BY a.deadline ASC",
            (uid,user['department_id'] or 0)).fetchall()
        flagged = db.execute(
            "SELECT p.score,a.title,a.course_code FROM plagiarism_scores p "
            "JOIN submissions s ON (s.id=p.sub_id_a OR s.id=p.sub_id_b) "
            "JOIN assignments a ON a.id=p.assignment_id "
            "WHERE s.user_id=? AND p.score>=70 ORDER BY p.score DESC LIMIT 5",(uid,)).fetchall()
    return render_template('user/dashboard.html', user=user, subs=subs, available=available,
                           completion=profile_pct(user), flagged=flagged)

@app.route('/submit/<int:asgn_id>', methods=['POST'])
@login_required
@student_required
def submit(asgn_id):
    uid = session['user_id']
    with get_db() as db:
        asgn    = db.execute("SELECT * FROM assignments WHERE id=?",(asgn_id,)).fetchone()
        student = db.execute("SELECT department_id FROM users WHERE id=?",(uid,)).fetchone()
        existing= db.execute("SELECT * FROM submissions WHERE user_id=? AND assignment_id=?",(uid,asgn_id)).fetchone()
    if not asgn: abort(404)
    # Department validation
    if asgn['department_id'] and student['department_id'] != asgn['department_id']:
        flash('You are not authorised to submit to this assignment.','error')
        return redirect(url_for('dashboard'))
    file = request.files.get('file')
    if not file or file.filename == '':
        flash('Please select a file.','error'); return redirect(url_for('dashboard'))
    if not allowed_file(file.filename):
        flash('Unsupported file type.','error'); return redirect(url_for('dashboard'))
    original = secure_filename(file.filename)
    ext = original.rsplit('.',1)[1].lower()
    ts  = datetime.now().strftime('%Y%m%d_%H%M%S')
    stored = f"u{uid}_a{asgn_id}_{ts}_{original}"
    path   = os.path.join(app.config['UPLOAD_FOLDER'], stored)
    file.save(path)
    try: deadline=datetime.strptime(asgn['deadline'][:16],'%Y-%m-%dT%H:%M'); is_late=1 if datetime.now()>deadline else 0
    except: is_late=0
    with get_db() as db:
        if existing:
            # Replace existing submission (before deadline)
            old_path = os.path.join(app.config['UPLOAD_FOLDER'], existing['filename'])
            if os.path.exists(old_path): os.remove(old_path)
            db.execute("UPDATE submissions SET filename=?,original_filename=?,file_ext=?,file_size=?,is_late=?,extracted_text=NULL,submitted_at=? WHERE id=?",
                       (stored,original,ext,os.path.getsize(path),is_late,datetime.now().strftime('%Y-%m-%dT%H:%M:%S'),existing['id']))
            db.execute("DELETE FROM plagiarism_scores WHERE sub_id_a=? OR sub_id_b=?",(existing['id'],existing['id']))
            db.commit()
            sub_id = existing['id']
        else:
            cur = db.execute("INSERT INTO submissions (user_id,assignment_id,filename,original_filename,file_ext,file_size,is_late) VALUES (?,?,?,?,?,?,?)",
                             (uid,asgn_id,stored,original,ext,os.path.getsize(path),is_late))
            db.commit(); sub_id=cur.lastrowid
    extract_and_store(sub_id, path, ext)
    flash('Submitted — marked LATE.' if is_late else 'Assignment submitted successfully.', 'error' if is_late else 'success')
    return redirect(url_for('dashboard'))

@app.route('/my/download/<int:sub_id>')
@login_required
@student_required
def student_download(sub_id):
    with get_db() as db:
        s = db.execute("SELECT * FROM submissions WHERE id=? AND user_id=?",(sub_id,session['user_id'])).fetchone()
    if not s: abort(404)
    return send_from_directory(app.config['UPLOAD_FOLDER'],s['filename'],as_attachment=True,download_name=s['original_filename'])

# ── Student: Profile ──────────────────────────────────────────────────────────
@app.route('/profile', methods=['GET','POST'])
@login_required
@student_required
def profile():
    uid = session['user_id']
    with get_db() as db:
        depts = db.execute("SELECT * FROM departments ORDER BY name").fetchall()
    if request.method == 'POST':
        action = request.form.get('action','profile')
        if action == 'profile':
            fields = {k:request.form.get(k,'').strip() for k in
                ['fullname','phone','level','program','gender','session_year','semester','state_of_origin']}
            dept_id = request.form.get('department_id','') or None
            if not fields['fullname']:
                flash('Full name cannot be empty.','error')
            else:
                with get_db() as db:
                    db.execute("""UPDATE users SET fullname=:fn,phone=:ph,level=:lv,program=:pg,
                        gender=:gn,session_year=:sy,semester=:sm,state_of_origin=:so,department_id=:di WHERE id=:id""",
                        dict(fn=fields['fullname'],ph=fields['phone'],lv=fields['level'],pg=fields['program'],
                             gn=fields['gender'],sy=fields['session_year'],sm=fields['semester'],
                             so=fields['state_of_origin'],di=dept_id,id=uid))
                    db.commit()
                session['fullname']=fields['fullname']
                flash('Profile updated.','success')
        elif action == 'change_password':
            cur=request.form.get('current_password',''); new=request.form.get('new_password',''); conf=request.form.get('confirm_password','')
            with get_db() as db:
                u=db.execute("SELECT password FROM users WHERE id=?",(uid,)).fetchone()
            if not check_password_hash(u['password'],cur): flash('Current password incorrect.','error')
            elif len(new)<6: flash('Password must be at least 6 characters.','error')
            elif new!=conf:  flash('Passwords do not match.','error')
            else:
                with get_db() as db:
                    db.execute("UPDATE users SET password=? WHERE id=?",(generate_password_hash(new),uid)); db.commit()
                flash('Password changed.','success')
        elif action == 'add_course':
            code=request.form.get('course_code','').strip().upper(); name=request.form.get('course_name','').strip()
            unit=request.form.get('credit_unit','0').strip(); sem=request.form.get('semester_c','').strip()
            if not code or not name: flash('Course code and name are required.','error')
            else:
                with get_db() as db:
                    if db.execute("SELECT id FROM student_courses WHERE user_id=? AND course_code=?",(uid,code)).fetchone():
                        flash(f'{code} already in your list.','error')
                    else:
                        db.execute("INSERT INTO student_courses (user_id,course_code,course_name,credit_unit,semester) VALUES (?,?,?,?,?)",
                                   (uid,code,name,int(unit) if unit.isdigit() else 0,sem)); db.commit()
                        flash(f'{code} added.','success')
        elif action == 'del_course':
            cid=request.form.get('course_id')
            with get_db() as db:
                db.execute("DELETE FROM student_courses WHERE id=? AND user_id=?",(cid,uid)); db.commit()
            flash('Course removed.','success')
        return redirect(url_for('profile'))
    with get_db() as db:
        user    = db.execute("SELECT u.*,d.name as dept_name FROM users u LEFT JOIN departments d ON u.department_id=d.id WHERE u.id=?",(uid,)).fetchone()
        courses = db.execute("SELECT * FROM student_courses WHERE user_id=? ORDER BY course_code",(uid,)).fetchall()
    return render_template('user/profile.html', user=user, courses=courses,
                           depts=depts, completion=profile_pct(user))

# ── Admin helpers ─────────────────────────────────────────────────────────────
def get_lecturer_dept_ids(uid, is_super):
    if is_super: return None  # None = all departments
    with get_db() as db:
        rows = db.execute("SELECT department_id FROM lecturer_departments WHERE lecturer_id=?",(uid,)).fetchall()
    return [r['department_id'] for r in rows]

# ── Admin: Dashboard ──────────────────────────────────────────────────────────
@app.route('/admin')
@admin_required
def admin_dashboard():
    uid=session['user_id']; is_super=session.get('is_superadmin',False)
    dept_ids=get_lecturer_dept_ids(uid,is_super)
    search=request.args.get('q','').strip(); dept_f=request.args.get('dept','').strip()
    late_f=request.args.get('late','').strip(); page=max(1,int(request.args.get('page',1)))
    PER_PAGE=25
    with get_db() as db:
        q=("SELECT s.*,u.fullname,u.matric,u.email,a.title,a.course_code,a.course_name,a.deadline,d.name as dept_name "
           "FROM submissions s JOIN users u ON s.user_id=u.id JOIN assignments a ON s.assignment_id=a.id "
           "LEFT JOIN departments d ON a.department_id=d.id")
        params=[]; where=[]
        if dept_ids is not None:
            placeholders=','.join('?'*len(dept_ids)) if dept_ids else '0'
            where.append(f"a.department_id IN ({placeholders})"); params+=dept_ids
        if search:
            where.append("(u.fullname LIKE ? OR u.matric LIKE ? OR a.title LIKE ?)"); params+=[f'%{search}%']*3
        if dept_f:
            where.append("a.department_id=?"); params.append(dept_f)
        if late_f=='1':
            where.append("s.is_late=1")
        if where: q+=" WHERE "+" AND ".join(where)
        q+=" ORDER BY s.submitted_at DESC"
        all_subs = db.execute(q,params).fetchall()
        total_subs=len(all_subs)
        subs=all_subs[(page-1)*PER_PAGE:page*PER_PAGE]
        total_pages=max(1,math.ceil(total_subs/PER_PAGE))
        # Stats
        s_params=dept_ids if dept_ids is not None else []
        if dept_ids is not None and dept_ids:
            ph=','.join('?'*len(dept_ids))
            total  =db.execute(f"SELECT COUNT(*) FROM submissions s JOIN assignments a ON s.assignment_id=a.id WHERE a.department_id IN ({ph})",dept_ids).fetchone()[0]
            late   =db.execute(f"SELECT COUNT(*) FROM submissions s JOIN assignments a ON s.assignment_id=a.id WHERE a.department_id IN ({ph}) AND s.is_late=1",dept_ids).fetchone()[0]
            stud_c =db.execute(f"SELECT COUNT(*) FROM users WHERE role='student' AND department_id IN ({ph})",dept_ids).fetchone()[0]
            asgn_c =db.execute(f"SELECT COUNT(*) FROM assignments WHERE department_id IN ({ph})",dept_ids).fetchone()[0]
        elif dept_ids is None:
            total  =db.execute("SELECT COUNT(*) FROM submissions").fetchone()[0]
            late   =db.execute("SELECT COUNT(*) FROM submissions WHERE is_late=1").fetchone()[0]
            stud_c =db.execute("SELECT COUNT(*) FROM users WHERE role='student'").fetchone()[0]
            asgn_c =db.execute("SELECT COUNT(*) FROM assignments").fetchone()[0]
        else:
            total=late=stud_c=asgn_c=0
        today=db.execute("SELECT COUNT(*) FROM submissions WHERE DATE(submitted_at)=DATE('now')").fetchone()[0]
        depts=db.execute("SELECT * FROM departments ORDER BY name").fetchall()
    return render_template('admin/dashboard.html', subs=subs, total=total, today=today,
                           late=late, stud_c=stud_c, asgn_c=asgn_c, depts=depts,
                           search=search, dept_f=dept_f, late_f=late_f,
                           page=page, total_pages=total_pages, total_subs=total_subs)

@app.route('/admin/download/<int:sid>')
@admin_required
def admin_download(sid):
    with get_db() as db:
        s=db.execute("SELECT * FROM submissions WHERE id=?",(sid,)).fetchone()
    if not s: abort(404)
    return send_from_directory(app.config['UPLOAD_FOLDER'],s['filename'],as_attachment=True,download_name=s['original_filename'])

@app.route('/admin/delete-sub/<int:sid>', methods=['POST'])
@admin_required
def admin_del_sub(sid):
    with get_db() as db:
        s=db.execute("SELECT * FROM submissions WHERE id=?",(sid,)).fetchone()
        if s:
            fp=os.path.join(app.config['UPLOAD_FOLDER'],s['filename'])
            if os.path.exists(fp): os.remove(fp)
            db.execute("DELETE FROM plagiarism_scores WHERE sub_id_a=? OR sub_id_b=?",(sid,sid))
            db.execute("DELETE FROM submissions WHERE id=?",(sid,)); db.commit()
    flash('Submission deleted.','success'); return redirect(url_for('admin_dashboard'))

# ── Admin: Assignments ────────────────────────────────────────────────────────
@app.route('/admin/assignments', methods=['GET','POST'])
@admin_required
def admin_assignments():
    uid=session['user_id']; is_super=session.get('is_superadmin',False)
    dept_ids=get_lecturer_dept_ids(uid,is_super)
    with get_db() as db:
        if dept_ids is None:
            depts=db.execute("SELECT * FROM departments ORDER BY name").fetchall()
        elif dept_ids:
            ph=','.join('?'*len(dept_ids))
            depts=db.execute(f"SELECT * FROM departments WHERE id IN ({ph}) ORDER BY name",dept_ids).fetchall()
        else:
            depts=[]
    if request.method=='POST':
        action=request.form.get('action','')
        if action=='create':
            title=request.form.get('title','').strip(); desc=request.form.get('description','').strip()
            ccode=request.form.get('course_code','').strip().upper(); cname=request.form.get('course_name','').strip()
            deadline=request.form.get('deadline','').strip(); dept_id=request.form.get('department_id','') or None
            if not title or not ccode or not cname or not deadline:
                flash('Title, course code, course name and deadline are required.','error')
            else:
                with get_db() as db:
                    cur=db.execute("INSERT INTO assignments (title,description,course_code,course_name,deadline,department_id,created_by) VALUES (?,?,?,?,?,?,?)",
                                   (title,desc,ccode,cname,deadline,dept_id,uid)); db.commit()
                    new_id=cur.lastrowid
                    emails=get_dept_student_emails(dept_id)
                    asgn=db.execute("SELECT * FROM assignments WHERE id=?",(new_id,)).fetchone()
                    dept=db.execute("SELECT name FROM departments WHERE id=?",(dept_id,)).fetchone() if dept_id else None
                notify_new_assignment(asgn, dept['name'] if dept else 'All Departments', emails)
                flash(f'Assignment "{title}" created. {len(emails)} student(s) notified.','success')
        elif action=='delete':
            aid=request.form.get('assignment_id')
            with get_db() as db:
                for s in db.execute("SELECT filename FROM submissions WHERE assignment_id=?",(aid,)).fetchall():
                    fp=os.path.join(app.config['UPLOAD_FOLDER'],s['filename'])
                    if os.path.exists(fp): os.remove(fp)
                db.execute("DELETE FROM submissions WHERE assignment_id=?",(aid,))
                db.execute("DELETE FROM plagiarism_scores WHERE assignment_id=?",(aid,))
                db.execute("DELETE FROM assignments WHERE id=?",(aid,)); db.commit()
            flash('Assignment deleted.','success')
        return redirect(url_for('admin_assignments'))
    with get_db() as db:
        if dept_ids is None:
            assignments=db.execute(
                "SELECT a.*,d.name as dept_name,"
                "(SELECT COUNT(*) FROM submissions WHERE assignment_id=a.id) as sub_count,"
                "(SELECT COUNT(*) FROM submissions WHERE assignment_id=a.id AND is_late=1) as late_count "
                "FROM assignments a LEFT JOIN departments d ON a.department_id=d.id ORDER BY a.deadline DESC").fetchall()
        elif dept_ids:
            ph=','.join('?'*len(dept_ids))
            assignments=db.execute(
                f"SELECT a.*,d.name as dept_name,"
                "(SELECT COUNT(*) FROM submissions WHERE assignment_id=a.id) as sub_count,"
                "(SELECT COUNT(*) FROM submissions WHERE assignment_id=a.id AND is_late=1) as late_count "
                f"FROM assignments a LEFT JOIN departments d ON a.department_id=d.id "
                f"WHERE a.department_id IN ({ph}) OR a.created_by=? ORDER BY a.deadline DESC",
                dept_ids+[uid]).fetchall()
        else:
            assignments=[]
    return render_template('admin/assignments.html', assignments=assignments, depts=depts, now=datetime.now())

@app.route('/admin/assignment/<int:aid>')
@admin_required
def admin_assignment_detail(aid):
    with get_db() as db:
        asgn=db.execute("SELECT a.*,d.name as dept_name FROM assignments a LEFT JOIN departments d ON a.department_id=d.id WHERE a.id=?",(aid,)).fetchone()
        if not asgn: abort(404)
        subs=db.execute("SELECT s.*,u.fullname,u.matric,u.email FROM submissions s JOIN users u ON s.user_id=u.id WHERE s.assignment_id=? ORDER BY s.submitted_at",(aid,)).fetchall()
        plag=db.execute(
            "SELECT p.*,sa.original_filename as fn_a,sb.original_filename as fn_b,"
            "ua.fullname as name_a,ub.fullname as name_b,ua.matric as matric_a,ub.matric as matric_b "
            "FROM plagiarism_scores p "
            "JOIN submissions sa ON p.sub_id_a=sa.id JOIN users ua ON sa.user_id=ua.id "
            "JOIN submissions sb ON p.sub_id_b=sb.id JOIN users ub ON sb.user_id=ub.id "
            "WHERE p.assignment_id=? ORDER BY p.score DESC",(aid,)).fetchall()
        extractable=db.execute("SELECT COUNT(*) FROM submissions WHERE assignment_id=? AND extracted_text IS NOT NULL AND extracted_text!=''",(aid,)).fetchone()[0]
        # Missing: students in dept who haven't submitted
        if asgn['department_id']:
            missing=db.execute(
                "SELECT u.id,u.fullname,u.matric,u.email FROM users u "
                "WHERE u.role='student' AND u.department_id=? "
                "AND u.id NOT IN (SELECT user_id FROM submissions WHERE assignment_id=?)",(asgn['department_id'],aid)).fetchall()
        else:
            missing=db.execute(
                "SELECT u.id,u.fullname,u.matric,u.email FROM users u WHERE u.role='student' "
                "AND u.id NOT IN (SELECT user_id FROM submissions WHERE assignment_id=?)",(aid,)).fetchall()
    return render_template('admin/assignment_detail.html', asgn=asgn, subs=subs,
                           plag=plag, extractable=extractable, missing=missing)

@app.route('/admin/assignment/<int:aid>/check-plagiarism', methods=['POST'])
@admin_required
def run_check(aid):
    results=run_plagiarism_check(aid)
    flagged=sum(1 for r in results if r['score']>=70)
    flash(f'Check complete. {len(results)} pair(s). {flagged} flagged ≥70%.','success')
    return redirect(url_for('admin_assignment_detail',aid=aid))

@app.route('/admin/assignment/<int:aid>/remind-missing', methods=['POST'])
@admin_required
def remind_missing(aid):
    with get_db() as db:
        asgn=db.execute("SELECT * FROM assignments WHERE id=?",(aid,)).fetchone()
        if asgn['department_id']:
            missing=db.execute(
                "SELECT u.fullname,u.email FROM users u WHERE u.role='student' AND u.department_id=? "
                "AND u.id NOT IN (SELECT user_id FROM submissions WHERE assignment_id=?)",(asgn['department_id'],aid)).fetchall()
        else:
            missing=db.execute(
                "SELECT u.fullname,u.email FROM users u WHERE u.role='student' "
                "AND u.id NOT IN (SELECT user_id FROM submissions WHERE assignment_id=?)",(aid,)).fetchall()
    for s in missing:
        notify_missing_reminder(s['email'],s['fullname'],asgn['title'],asgn['deadline'])
    flash(f'Reminder sent to {len(missing)} student(s) who haven\'t submitted.','success')
    return redirect(url_for('admin_assignment_detail',aid=aid))

# ── Admin: Departments ────────────────────────────────────────────────────────
@app.route('/admin/departments', methods=['GET','POST'])
@superadmin_required
def admin_departments():
    if request.method=='POST':
        action=request.form.get('action','')
        if action=='add':
            name=request.form.get('name','').strip(); fac=request.form.get('faculty','').strip()
            if not name: flash('Department name is required.','error')
            else:
                try:
                    with get_db() as db:
                        db.execute("INSERT INTO departments (name,faculty) VALUES (?,?)",(name,fac)); db.commit()
                    flash(f'{name} added.','success')
                except sqlite3.IntegrityError:
                    flash('Department already exists.','error')
        elif action=='delete':
            did=request.form.get('dept_id')
            with get_db() as db:
                db.execute("DELETE FROM departments WHERE id=?",(did,)); db.commit()
            flash('Department removed.','success')
        return redirect(url_for('admin_departments'))
    with get_db() as db:
        depts=db.execute(
            "SELECT d.*,COUNT(u.id) as student_count FROM departments d "
            "LEFT JOIN users u ON u.department_id=d.id AND u.role='student' "
            "GROUP BY d.id ORDER BY d.faculty,d.name").fetchall()
    return render_template('admin/departments.html', depts=depts)

# ── Admin: Students ───────────────────────────────────────────────────────────
@app.route('/admin/students')
@admin_required
def admin_students():
    uid=session['user_id']; is_super=session.get('is_superadmin',False)
    dept_ids=get_lecturer_dept_ids(uid,is_super)
    q=request.args.get('q','').strip()
    with get_db() as db:
        query=("SELECT u.*,d.name as dept_name,COUNT(s.id) as sub_count "
               "FROM users u LEFT JOIN departments d ON u.department_id=d.id "
               "LEFT JOIN submissions s ON u.id=s.user_id WHERE u.role='student'")
        params=[]
        if dept_ids is not None and dept_ids:
            ph=','.join('?'*len(dept_ids))
            query+=f" AND u.department_id IN ({ph})"; params+=dept_ids
        elif dept_ids is not None and not dept_ids:
            query+=" AND 1=0"
        if q:
            query+=" AND (u.fullname LIKE ? OR u.matric LIKE ? OR u.email LIKE ?)"; params+=[f'%{q}%']*3
        query+=" GROUP BY u.id ORDER BY u.fullname"
        students=db.execute(query,params).fetchall()
    return render_template('admin/students.html', students=students, q=q)

@app.route('/admin/student/<int:uid>')
@admin_required
def admin_student_detail(uid):
    with get_db() as db:
        student=db.execute("SELECT u.*,d.name as dept_name FROM users u LEFT JOIN departments d ON u.department_id=d.id WHERE u.id=? AND u.role='student'",(uid,)).fetchone()
        if not student: abort(404)
        subs=db.execute("SELECT s.*,a.title,a.course_code,a.deadline FROM submissions s JOIN assignments a ON s.assignment_id=a.id WHERE s.user_id=? ORDER BY s.submitted_at DESC",(uid,)).fetchall()
        courses=db.execute("SELECT * FROM student_courses WHERE user_id=? ORDER BY course_code",(uid,)).fetchall()
    return render_template('admin/student_detail.html', student=student, subs=subs,
                           courses=courses, completion=profile_pct(student))

# ── Admin: Manage Admins ──────────────────────────────────────────────────────
@app.route('/admin/admins', methods=['GET','POST'])
@superadmin_required
def admin_manage():
    if request.method=='POST':
        action=request.form.get('action','')
        if action=='create':
            fn=request.form.get('fullname','').strip(); em=request.form.get('email','').strip().lower()
            pw=request.form.get('password',''); cf=request.form.get('confirm','')
            dept_ids_new=request.form.getlist('dept_ids')
            errors=[]
            if not fn: errors.append('Full name required.')
            if not em: errors.append('Email required.')
            if len(pw)<6: errors.append('Password min 6 chars.')
            if pw!=cf: errors.append('Passwords do not match.')
            if errors:
                for e in errors: flash(e,'error')
            else:
                try:
                    with get_db() as db:
                        cur=db.execute("INSERT INTO users (fullname,email,password,role) VALUES (?,?,?,'admin')",
                                       (fn,em,generate_password_hash(pw))); db.commit()
                        new_id=cur.lastrowid
                        for did in dept_ids_new:
                            db.execute("INSERT OR IGNORE INTO lecturer_departments (lecturer_id,department_id) VALUES (?,?)",(new_id,did))
                        db.commit()
                    flash(f'Admin created for {fn}.','success')
                except sqlite3.IntegrityError:
                    flash('Email already exists.','error')
        elif action=='update_depts':
            target_id=int(request.form.get('target_id',0))
            new_depts=request.form.getlist('dept_ids')
            with get_db() as db:
                db.execute("DELETE FROM lecturer_departments WHERE lecturer_id=?",(target_id,))
                for did in new_depts:
                    db.execute("INSERT OR IGNORE INTO lecturer_departments (lecturer_id,department_id) VALUES (?,?)",(target_id,did))
                db.commit()
            flash('Departments updated.','success')
        elif action=='delete':
            tid=int(request.form.get('target_id',0))
            if tid==session['user_id']: flash('Cannot delete your own account.','error')
            else:
                with get_db() as db:
                    db.execute("DELETE FROM lecturer_departments WHERE lecturer_id=?",(tid,))
                    db.execute("DELETE FROM users WHERE id=? AND role='admin'",(tid,)); db.commit()
                flash('Admin removed.','success')
        return redirect(url_for('admin_manage'))
    with get_db() as db:
        admins=db.execute("SELECT * FROM users WHERE role='admin' ORDER BY is_superadmin DESC,created_at").fetchall()
        depts=db.execute("SELECT * FROM departments ORDER BY name").fetchall()
        lecturer_depts={} 
        for a in admins:
            rows=db.execute("SELECT department_id FROM lecturer_departments WHERE lecturer_id=?",(a['id'],)).fetchall()
            lecturer_depts[a['id']]=[r['department_id'] for r in rows]
    return render_template('admin/admins.html', admins=admins, depts=depts, lecturer_depts=lecturer_depts)

# ── AI routes (CSRF exempt — JSON API) ───────────────────────────────────────
@app.route('/api/chat', methods=['POST'])
@login_required
@csrf.exempt
def ai_chat():
    data    = request.get_json() or {}
    message = data.get('message','').strip()
    history = data.get('history',[])
    if not message: return jsonify(reply='Please type a message.')

    uid  = session['user_id']
    role = session.get('role','student')

    # Check if the message is asking to read/review a file
    read_keywords = ['read','review','check','what did i write','what does it say',
                     'content of','my submission','my file','read my']
    wants_file_read = any(kw in message.lower() for kw in read_keywords)

    file_content_ctx = ''
    if wants_file_read:
        with get_db() as db:
            if role == 'student':
                # Get the most recent submission with extracted text
                sub = db.execute(
                    "SELECT s.extracted_text,s.original_filename,a.title,a.course_code "
                    "FROM submissions s JOIN assignments a ON s.assignment_id=a.id "
                    "WHERE s.user_id=? AND s.extracted_text IS NOT NULL AND s.extracted_text!='' "
                    "ORDER BY s.submitted_at DESC LIMIT 1",(uid,)).fetchone()
                if sub:
                    file_content_ctx = (
                        "\n\n=== STUDENT'S MOST RECENT SUBMISSION (extracted text) ===\n"
                        "Assignment: " + sub['title'] + " (" + sub['course_code'] + ")\n"
                        "File: " + sub['original_filename'] + "\n"
                        "Content:\n" + (sub['extracted_text'] or '')[:3000]
                    )
                else:
                    file_content_ctx = "\n\n[No extractable text found in submissions. File may be an image or scanned document.]"
            else:
                # Admin asking to read — get most recent submission overall (their dept)
                sub = db.execute(
                    "SELECT s.extracted_text,s.original_filename,u.fullname,a.title,a.course_code "
                    "FROM submissions s JOIN users u ON s.user_id=u.id "
                    "JOIN assignments a ON s.assignment_id=a.id "
                    "WHERE s.extracted_text IS NOT NULL AND s.extracted_text!='' "
                    "ORDER BY s.submitted_at DESC LIMIT 1").fetchone()
                if sub:
                    file_content_ctx = (
                        "\n\n=== MOST RECENT SUBMISSION WITH TEXT ===\n"
                        "Student: " + sub['fullname'] + "\n"
                        "Assignment: " + sub['title'] + " (" + sub['course_code'] + ")\n"
                        "File: " + sub['original_filename'] + "\n"
                        "Content:\n" + (sub['extracted_text'] or '')[:3000]
                    )

    if role == 'student':
        context = build_student_context(uid)
        system  = (
            'You are an AI assistant for the TJs-Code Assignment Portal at IMAP, Lafia, Nigeria.\n'
            'You are speaking with a student. Here is everything you know about them right now:\n\n'
            + context + file_content_ctx +
            '\n\nIMPORTANT RULES:\n'
            '- Answer using the real data above. Never say "not specified" or "anonymous" — the data is there.\n'
            '- If asked about assignments, use the exact titles, courses, and deadlines from the context.\n'
            '- If asked about their submission or file content, use the extracted text above.\n'
            '- Be friendly, specific, and concise. Address the student by their first name.\n'
            '- If asked how to change password: tell them to go to Profile > Change Password tab.'
        )
    else:
        context = build_admin_context(uid)
        system  = (
            'You are an AI assistant for the TJs-Code Assignment Portal at IMAP, Lafia, Nigeria.\n'
            'You are speaking with a lecturer/admin. Here is their current portal data:\n\n'
            + context + file_content_ctx +
            '\n\nIMPORTANT RULES:\n'
            '- Use the real names, numbers, and dates from the context above.\n'
            '- Never say "anonymous" or "not specified" — actual data is in the context.\n'
            '- If asked about a submission content, use the extracted text provided.\n'
            '- Be professional and specific. Reference actual student names and assignment titles.\n'
            '- You can suggest running plagiarism checks, emailing missing students, or using AI Review buttons.'
        )

    reply = call_groq(history[-6:] + [{'role':'user','content':message}], system)
    return jsonify(reply=reply)

@app.route('/admin/submission/<int:sid>/ai-review', methods=['POST'])
@admin_required
@csrf.exempt
def ai_review_submission(sid):
    with get_db() as db:
        sub=db.execute("SELECT s.*,u.fullname,a.title,a.course_code FROM submissions s JOIN users u ON s.user_id=u.id JOIN assignments a ON s.assignment_id=a.id WHERE s.id=?",(sid,)).fetchone()
    if not sub: return jsonify(review='Submission not found.')
    text=sub['extracted_text'] or ''
    if not text.strip(): return jsonify(review='No extractable text in this submission (may be image or scanned PDF).')
    prompt=(f'Review this student submission:\nStudent: {sub["fullname"]}\nAssignment: {sub["title"]} ({sub["course_code"]})\n\n'
            f'Content:\n{text[:3000]}\n\nProvide:\n1. Brief summary (2-3 sentences)\n2. Key strengths\n3. Areas for improvement\n4. Suggested grade (A/B/C/D/F) with justification')
    return jsonify(review=call_groq([{'role':'user','content':prompt}],
        'You are an experienced Nigerian polytechnic lecturer providing fair, constructive academic feedback.'))

@app.route('/admin/assignment/<int:aid>/plagiarism-explain', methods=['POST'])
@admin_required
@csrf.exempt
def plagiarism_explain(aid):
    data=request.get_json() or {}
    with get_db() as db:
        sa=db.execute("SELECT s.*,u.fullname FROM submissions s JOIN users u ON s.user_id=u.id WHERE s.id=?",(data.get('sub_a'),)).fetchone()
        sb=db.execute("SELECT s.*,u.fullname FROM submissions s JOIN users u ON s.user_id=u.id WHERE s.id=?",(data.get('sub_b'),)).fetchone()
    if not sa or not sb: return jsonify(explanation='Submissions not found.')
    ta=(sa['extracted_text'] or '')[:1500]; tb=(sb['extracted_text'] or '')[:1500]
    if not ta or not tb: return jsonify(explanation='One or both submissions have no extractable text.')
    prompt=(f'Two submissions have {data.get("score",0)}% similarity.\n\n'
            f'Submission A ({sa["fullname"]}):\n{ta}\n\nSubmission B ({sb["fullname"]}):\n{tb}\n\n'
            f'Identify: 1. Matching phrases/sections 2. Whether direct copy or paraphrase 3. Integrity concern level (Low/Medium/High)')
    return jsonify(explanation=call_groq([{'role':'user','content':prompt}],
        'You are an academic integrity officer analysing student submissions for plagiarism.'))

@app.route('/admin/ai-brief', methods=['POST'])
@admin_required
@csrf.exempt
def ai_brief():
    data=request.get_json() or {}
    if not data.get('topic'): return jsonify(brief='Please provide a topic.')
    prompt=(f'Write a professional assignment brief for a Nigerian polytechnic.\n'
            f'Topic: {data.get("topic","")}\nCourse: {data.get("course","")}\n'
            f'Level: {data.get("level","HND")}\nDeadline: {data.get("deadline","TBD")}\n\n'
            f'Include: title, background, objectives, 3-5 tasks/questions, submission format, marking scheme. '
            f'Keep it appropriate for Nigerian polytechnic students.')
    return jsonify(brief=call_groq([{'role':'user','content':prompt}],
        'You are an experienced Nigerian polytechnic lecturer writing assignment briefs.'))

# ── Error handlers ────────────────────────────────────────────────────────────
@app.errorhandler(403)
def forbidden(e): return render_template('errors/403.html'),403
@app.errorhandler(404)
def not_found(e): return render_template('errors/404.html'),404
@app.errorhandler(413)
def too_large(e):
    flash('File too large. Max 15 MB.','error')
    return redirect(request.referrer or url_for('dashboard'))
@app.errorhandler(429)
def rate_limited(e):
    flash('Too many attempts. Please wait a minute and try again.','error')
    return redirect(request.referrer or url_for('login'))

if __name__ == '__main__':
    app.run(debug=True, use_reloader=True)
