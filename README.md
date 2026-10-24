# TJs-Code Departmental Assignment Portal
**Isa Mustapha Agwai I Polytechnic, Lafia — Nasarawa State**

---

## Quick Start

```bash
cd imap_portal
pip install -r requirements.txt
python app.py
```

Open: **http://127.0.0.1:5000**

---

## Default Admin Login

| Field    | Value               |
|----------|---------------------|
| Email    | admin@imap.edu.ng   |
| Password | Admin@1234          |

> Change this immediately after first login via **Admin → Admins → Add Admin**, then remove the default.

---

## Student Login

Students log in with their **matric number** (e.g. `IMAP/SST/HSWD/25/26/0006`) or email address.

---

## Features

### Student
- Register with matric number + department
- Dashboard with IMAP-style info tiles (Level, Programme, Session, Semester)
- View available assignments (filtered to their department)
- Submit files (PDF / Word / Images) up to 15 MB
- Late submissions auto-flagged with red LATE badge (still accepted)
- View submission history with on-time / late status
- Full profile: personal info, academic details, course registration, password change

### Admin
- **Dashboard** — all submissions with search, department filter, late-only filter
- **Assignments** — create with title, course code/name, deadline, department restriction; view per-assignment submissions; delete
- **Students** — browse all students, search by name/matric, view full profile + courses + submissions
- **Departments** — add/remove departments that students register under
- **Admins** — add new admin accounts, remove existing ones (cannot remove self)

---

## File Structure

```
imap_portal/
├── app.py                          # All routes + DB logic
├── requirements.txt
├── portal.db                       # SQLite (auto-created on first run)
├── uploads/                        # Uploaded files (auto-created)
├── static/
│   ├── css/style.css               # IMAP-themed stylesheet
│   └── img/crest.png               # ← Place IMAP crest PNG here
└── templates/
    ├── base.html                   # White top navbar shell
    ├── login.html                  # Purple-blue gradient login
    ├── register.html
    ├── user/
    │   ├── dashboard.html          # Student dashboard + assignment submit
    │   └── profile.html            # Profile + courses + password
    └── admin/
        ├── dashboard.html          # Submissions overview
        ├── assignments.html        # Manage assignments
        ├── assignment_detail.html  # Per-assignment submissions
        ├── students.html           # Student list
        ├── student_detail.html     # Student full profile
        ├── departments.html        # Manage departments
        └── admins.html             # Manage admin accounts
```

---

## Adding the IMAP Crest

Place the official IMAP crest image at:
```
static/img/crest.png
```
It will appear on the login page and navbar automatically.
If the file is missing, a fallback icon renders in its place.

---

## Production Deployment

```bash
# Set a strong secret key
export SECRET_KEY="your-very-long-random-string"

# Install gunicorn
pip install gunicorn

# Run
gunicorn app:app --bind 0.0.0.0:8000 --workers 2
```

For Nginx, proxy to port 8000 and serve `uploads/` as a static directory.

---

## Built by
**TJS Code** — *Olatunde Jonah Olabanjo*
