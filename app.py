from flask import Flask, request, redirect, url_for, render_template, render_template_string, session
import sqlite3
import os
import re
from datetime import date, timedelta
from openpyxl import load_workbook
import pandas as pd

app = Flask(__name__)
app.secret_key = "KITAPHANA_SECRET_2026"

# =========================================================
# ОРТАҚ ДИЗАЙН ЖӘНЕ НАВИГАЦИЯ (БАРЛЫҚ БЕТКЕ АВТОМАТТЫ ҚОСЫЛАДЫ)
# =========================================================
GLOBAL_CSS = """
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
    :root {
        --primary: #4f46e5;
        --primary-hover: #4338ca;
        --secondary: #0ea5e9;
        --success: #10b981;
        --danger: #ef4444;
        --warning: #f59e0b;
        --dark: #0f172a;
        --slate: #475569;
        --light: #f8fafc;
        --border: #e2e8f0;
    }
    * {
        box-sizing: border-box;
    }
    body {
        font-family: 'Inter', 'Segoe UI', sans-serif !important;
        background: #f1f5f9;
        color: var(--dark);
        margin: 0;
        padding-top: 85px !important;
        padding-bottom: 40px;
        min-height: 100vh;
    }
    /* Жоғарғы Навигация */
    .top-navbar {
        position: fixed;
        top: 0; left: 0; right: 0;
        height: 70px;
        background: rgba(15, 23, 42, 0.92);
        backdrop-filter: blur(14px);
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 0 28px;
        z-index: 99999;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.15);
        border-bottom: 1px solid rgba(255, 255, 255, 0.1);
    }
    .nav-brand {
        display: flex;
        align-items: center;
        gap: 12px;
        color: #ffffff;
        font-weight: 800;
        font-size: 18px;
        text-decoration: none;
        letter-spacing: 0.5px;
    }
    .nav-brand span {
        background: linear-gradient(135deg, #60a5fa, #c084fc);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .nav-links {
        display: flex;
        align-items: center;
        gap: 6px;
        overflow-x: auto;
        padding: 4px 0;
    }
    .nav-item {
        color: #cbd5e1;
        text-decoration: none;
        padding: 8px 14px;
        border-radius: 10px;
        font-size: 13.5px;
        font-weight: 600;
        transition: all 0.2s ease;
        white-space: nowrap;
    }
    .nav-item:hover {
        color: #ffffff;
        background: rgba(255, 255, 255, 0.08);
    }
    .nav-item.active {
        color: #ffffff;
        background: linear-gradient(135deg, #4f46e5, #3b82f6);
        box-shadow: 0 4px 12px rgba(79, 70, 229, 0.35);
    }
    .nav-user {
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .role-badge {
        background: rgba(255, 255, 255, 0.12);
        color: #e2e8f0;
        padding: 6px 12px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
        border: 1px solid rgba(255, 255, 255, 0.15);
    }
    .nav-logout {
        background: linear-gradient(135deg, #ef4444, #dc2626);
        color: #ffffff;
        text-decoration: none;
        padding: 8px 16px;
        border-radius: 10px;
        font-size: 13.5px;
        font-weight: 600;
        box-shadow: 0 4px 12px rgba(239, 68, 68, 0.3);
        transition: all 0.2s ease;
        white-space: nowrap;
    }
    .nav-logout:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(239, 68, 68, 0.45);
    }

    /* Ортақ контейнерлер мен элементтер */
    .page-container {
        max-width: 1220px;
        margin: 0 auto;
        padding: 0 20px;
    }
    .page-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 16px;
        margin-bottom: 28px;
        background: white;
        padding: 22px 28px;
        border-radius: 20px;
        box-shadow: 0 4px 20px rgba(15, 23, 42, 0.04);
        border: 1px solid var(--border);
    }
    .page-title {
        margin: 0;
        font-size: 24px;
        font-weight: 800;
        color: var(--dark);
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .page-subtitle {
        margin: 4px 0 0 0;
        font-size: 14px;
        color: var(--slate);
    }

    /* Іздеу жолағы */
    .search-box {
        position: relative;
        min-width: 280px;
        flex: 1;
        max-width: 380px;
    }
    .search-input {
        width: 100%;
        padding: 12px 16px 12px 42px;
        border: 1.5px solid var(--border);
        border-radius: 12px;
        font-size: 14px;
        background: var(--light);
        color: var(--dark);
        transition: all 0.2s;
        margin: 0 !important;
    }
    .search-input:focus {
        background: white;
        border-color: var(--primary);
        outline: none;
        box-shadow: 0 0 0 4px rgba(79, 70, 229, 0.12);
    }
    .search-icon {
        position: absolute;
        left: 14px;
        top: 50%;
        transform: translateY(-50%);
        color: #94a3b8;
        font-size: 16px;
        pointer-events: none;
    }

    /* Кестелер */
    .table-card {
        background: white;
        border-radius: 20px;
        box-shadow: 0 10px 30px rgba(15, 23, 42, 0.05);
        border: 1px solid var(--border);
        overflow: hidden;
    }
    .table-responsive {
        overflow-x: auto;
    }
    table.modern-table {
        width: 100%;
        border-collapse: collapse;
        text-align: left;
    }
    table.modern-table th {
        background: #f8fafc;
        color: #475569;
        font-weight: 700;
        font-size: 12.5px;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        padding: 16px 20px;
        border-bottom: 2px solid var(--border);
    }
    table.modern-table td {
        padding: 15px 20px;
        border-bottom: 1px solid #f1f5f9;
        font-size: 14.5px;
        color: #1e293b;
        vertical-align: middle;
    }
    table.modern-table tr:hover td {
        background: #f8fafc;
    }
    table.modern-table tr:last-child td {
        border-bottom: none;
    }

    /* Капсулалар (Badges) */
    .badge {
        display: inline-flex;
        align-items: center;
        gap: 5px;
        padding: 5px 12px;
        border-radius: 999px;
        font-size: 12.5px;
        font-weight: 700;
        letter-spacing: 0.2px;
    }
    .badge-success { background: #dcfce7; color: #15803d; }
    .badge-danger { background: #fee2e2; color: #b91c1c; }
    .badge-warning { background: #fef3c7; color: #b45309; }
    .badge-info { background: #e0f2fe; color: #0369a1; }
    .badge-purple { background: #f3e8ff; color: #6b21a8; }

    /* Батырмалар */
    .btn-modern {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        gap: 8px;
        padding: 11px 20px;
        border-radius: 12px;
        font-weight: 600;
        font-size: 14px;
        text-decoration: none;
        border: none;
        cursor: pointer;
        transition: all 0.2s ease;
    }
    .btn-primary {
        background: linear-gradient(135deg, #4f46e5, #3b82f6);
        color: white;
        box-shadow: 0 4px 14px rgba(79, 70, 229, 0.28);
    }
    .btn-primary:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(79, 70, 229, 0.4);
    }
    .btn-danger {
        background: linear-gradient(135deg, #ef4444, #dc2626);
        color: white;
        box-shadow: 0 4px 14px rgba(239, 68, 68, 0.25);
    }
    .btn-danger:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(239, 68, 68, 0.35);
    }
    .btn-outline {
        background: white;
        color: var(--dark);
        border: 1.5px solid var(--border);
    }
    .btn-outline:hover {
        border-color: var(--primary);
        color: var(--primary);
    }

    @media (max-width: 900px) {
        .top-navbar {
            padding: 0 14px;
            height: auto;
            min-height: 64px;
            flex-wrap: wrap;
            gap: 8px;
            padding-top: 10px;
            padding-bottom: 10px;
        }
        body {
            padding-top: 115px !important;
        }
        .role-badge { display: none; }
    }
</style>
<script>
function filterTable(inputId, tableId) {
    const input = document.getElementById(inputId);
    const filter = input.value.toLowerCase();
    const table = document.getElementById(tableId);
    if (!table) return;
    const tr = table.getElementsByTagName("tr");
    for (let i = 1; i < tr.length; i++) {
        let text = tr[i].textContent || tr[i].innerText;
        tr[i].style.display = text.toLowerCase().indexOf(filter) > -1 ? "" : "none";
    }
}
</script>
"""

@app.after_request
def add_global_navigation(response):
    if request.path in ("/login", "/logout"):
        return response
    if "username" not in session:
        return response
    if not response.content_type or "text/html" not in response.content_type:
        return response
    html = response.get_data(as_text=True)

    path = request.path
    role_name = "Кітапханашы" if session.get("role") == "librarian" else "Әкімші"

    nav_items = [
        ("/", "🏠 Басты бет"),
        ("/readers", "👥 Оқырмандар"),
        ("/books", "📚 Кітаптар"),
        ("/loans", "🔄 Кітап беру"),
        ("/rating", "👨🎓 Оқушылар"),
        ("/teacher-rating", "👩🏫 Мұғалімдер"),
        ("/book-rating", "🥇 Кітап рейтингі"),
        ("/report", "📊 Есеп"),
    ]
    links_html = ""
    for href, label in nav_items:
        active_cls = "active" if (path == href or (href != "/" and path.startswith(href))) else ""
        links_html += f'<a class="nav-item {active_cls}" href="{href}">{label}</a>'

    navigation = f"""
    {GLOBAL_CSS}
    <nav class="top-navbar">
        <a class="nav-brand" href="/">
            📚 <span>SMART КІТАПХАНА</span>
        </a>
        <div class="nav-links">
            {links_html}
        </div>
        <div class="nav-user">
            <span class="role-badge">👤 {role_name}</span>
            <a class="nav-logout" href="/logout">🚪 Шығу</a>
        </div>
    </nav>
    """
    if "</body>" in html:
        html = html.replace("</body>", navigation + "</body>", 1)
    else:
        html += navigation
    response.set_data(html)
    return response

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

USERS = {
    "library": {
        "password": "library123",
        "role": "librarian"
    },
    "admin": {
        "password": "admin123",
        "role": "admin"
    }
}
DB_NAME = "kitaphana.db"
DATABASE_URL = os.environ.get("DATABASE_URL")

def get_db():
    if DATABASE_URL:
        import psycopg
        from psycopg.rows import dict_row
        return psycopg.connect(
            DATABASE_URL,
            row_factory=dict_row
        )
    db = sqlite3.connect(DB_NAME)
    db.row_factory = sqlite3.Row
    return db

def init_db():
    db = get_db()
    if DATABASE_URL:
        id_type = "INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY"
    else:
        id_type = "INTEGER PRIMARY KEY AUTOINCREMENT"
    
    # ОҚЫРМАНДАР
    db.execute(f"""
        CREATE TABLE IF NOT EXISTS okuyandar (
            id {id_type},
            aty TEXT NOT NULL,
            synyp TEXT,
            turi TEXT,
            tirkelgen_kuni DATE
        )
    """)
    if DATABASE_URL:
        db.execute("""
            ALTER TABLE okuyandar
            ADD COLUMN IF NOT EXISTS tirkelgen_kuni DATE
        """)
    else:
        columns = [
            row["name"]
            for row in db.execute("PRAGMA table_info(okuyandar)").fetchall()
        ]
        if "tirkelgen_kuni" not in columns:
            db.execute("""
                ALTER TABLE okuyandar
                ADD COLUMN tirkelgen_kuni TEXT
            """)
    db.execute("""
        UPDATE okuyandar
        SET tirkelgen_kuni = '2026-09-01'
        WHERE tirkelgen_kuni IS NULL
    """)
    
    # КІТАПТАР
    db.execute(f"""
        CREATE TABLE IF NOT EXISTS kitaptar (
            id {id_type},
            kitap_id TEXT NOT NULL UNIQUE,
            ataui TEXT NOT NULL,
            avtor TEXT,
            sanat TEXT,
            jalpy_sany INTEGER DEFAULT 0
        )
    """)
    
    # КІТАП БЕРУ
    db.execute(f"""
        CREATE TABLE IF NOT EXISTS kitap_beru (
            id {id_type},
            kuni TEXT NOT NULL,
            okuyman_id INTEGER NOT NULL,
            kitap_id INTEGER NOT NULL,
            berildi TEXT DEFAULT 'Иә',
            kaitaryldy TEXT DEFAULT 'Жоқ',
            kaitaru_kuni TEXT,
            merzim TEXT
        )
    """)
    if DATABASE_URL:
        db.execute("""
            ALTER TABLE kitaptar
            ADD COLUMN IF NOT EXISTS baspa TEXT
        """)
        db.execute("""
            ALTER TABLE kitaptar
            ADD COLUMN IF NOT EXISTS zhyl TEXT
        """)
    else:
        columns = {
            row["name"]
            for row in db.execute(
                "PRAGMA table_info(kitaptar)"
            ).fetchall()
        }
        if "baspa" not in columns:
            db.execute(
                "ALTER TABLE kitaptar ADD COLUMN baspa TEXT"
            )
        if "zhyl" not in columns:
            db.execute(
                "ALTER TABLE kitaptar ADD COLUMN zhyl TEXT"
            )
    db.commit()

    # personal.xlsx ФАЙЛЫНАН ОҚЫРМАНДАРДЫ ЖҮКТЕУ (Тек база бос болса немесе жаңа адам болса)
    teachers_file = "personal.xlsx"
    if os.path.exists(teachers_file):
        existing_rows = db.execute("SELECT aty, synyp FROM okuyandar").fetchall()
        existing = {((r["aty"] or "").strip().casefold(), (r["synyp"] or "").strip().casefold()) for r in existing_rows}
        df_t = pd.read_excel(teachers_file)
        df_t.columns = [str(col).strip() for col in df_t.columns]
        for _, row in df_t.iterrows():
            raw_aty = str(row.get("Аты-жөні", "")).strip()
            aty = re.sub(r"^\d+[\.\)]\s*", "", raw_aty).strip()
            raw_synyp = row.get("Сыныбы/пәні", "")
            synyp = "" if pd.isna(raw_synyp) else str(raw_synyp).strip()
            if not aty or aty.lower() == "nan":
                continue
            if not synyp or synyp.lower() == "nan":
                continue
            if synyp.startswith("0") or "даярлық" in synyp.lower():
                continue
            key = (aty.casefold(), synyp.casefold())
            if key in existing:
                continue
            if re.match(r"^\d+\s*[-]?\s*[а-яәғқңөұүһі]?$", synyp.lower()):
                turi = "Оқушы"
            else:
                turi = "Мұғалім"
            if DATABASE_URL:
                db.execute("""
                    INSERT INTO okuyandar
                    (aty, synyp, turi, tirkelgen_kuni)
                    VALUES (%s, %s, %s, %s)
                """, (aty, synyp, turi, "2026-09-01"))
            else:
                db.execute("""
                    INSERT INTO okuyandar
                    (aty, synyp, turi, tirkelgen_kuni)
                    VALUES (?, ?, ?, ?)
                """, (aty, synyp, turi, "2026-09-01"))
            existing.add(key)
        db.commit()
    db.close()

# =========================================================
# АВТОРИЗАЦИЯ (LOGIN)
# =========================================================
@app.route("/login", methods=["GET", "POST"])
def login():
    error = ""
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]
        user = USERS.get(username)
        if user and user["password"] == password:
            session["username"] = username
            session["role"] = user["role"]
            return redirect(url_for("home"))
        error = "Логин немесе құпиясөз қате!"
    return render_template_string("""
    <!DOCTYPE html>
    <html lang="kk">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Кітапханаға кіру</title>
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap" rel="stylesheet">
        <style>
            * { box-sizing: border-box; }
            body {
                font-family: 'Inter', sans-serif;
                background: radial-gradient(circle at top left, #1e1b4b, #0f172a 60%, #020617);
                display: flex;
                justify-content: center;
                align-items: center;
                min-height: 100vh;
                margin: 0;
                padding: 20px;
                position: relative;
                overflow: hidden;
            }
            /* Декоративті жарық шарлар */
            .blob-1, .blob-2 {
                position: absolute;
                border-radius: 50%;
                filter: blur(80px);
                z-index: 0;
            }
            .blob-1 {
                width: 350px; height: 350px;
                background: rgba(79, 70, 229, 0.4);
                top: -80px; left: -80px;
            }
            .blob-2 {
                width: 300px; height: 300px;
                background: rgba(14, 165, 233, 0.3);
                bottom: -60px; right: -60px;
            }
            .login-card {
                position: relative;
                z-index: 1;
                background: rgba(255, 255, 255, 0.07);
                backdrop-filter: blur(20px);
                width: 100%;
                max-width: 420px;
                padding: 45px 35px;
                border-radius: 28px;
                box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
                border: 1px solid rgba(255, 255, 255, 0.15);
                text-align: center;
                color: white;
            }
            .logo-circle {
                width: 72px; height: 72px;
                background: linear-gradient(135deg, #4f46e5, #0ea5e9);
                border-radius: 20px;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 34px;
                margin: 0 auto 20px auto;
                box-shadow: 0 10px 25px rgba(79, 70, 229, 0.4);
            }
            h2 {
                margin: 0 0 8px 0;
                font-size: 24px;
                font-weight: 800;
                letter-spacing: -0.5px;
            }
            p.sub {
                margin: 0 0 30px 0;
                color: #94a3b8;
                font-size: 14px;
            }
            .input-group {
                text-align: left;
                margin-bottom: 18px;
            }
            .input-group label {
                display: block;
                font-size: 13px;
                font-weight: 600;
                color: #cbd5e1;
                margin-bottom: 8px;
            }
            input {
                width: 100%;
                padding: 14px 16px;
                border: 1px solid rgba(255, 255, 255, 0.18);
                border-radius: 14px;
                font-size: 15px;
                background: rgba(15, 23, 42, 0.6);
                color: white;
                transition: all 0.25s;
            }
            input::placeholder { color: #64748b; }
            input:focus {
                border-color: #60a5fa;
                background: rgba(15, 23, 42, 0.9);
                outline: none;
                box-shadow: 0 0 0 4px rgba(59, 130, 246, 0.2);
            }
            button {
                width: 100%;
                padding: 15px;
                margin-top: 10px;
                background: linear-gradient(135deg, #4f46e5, #2563eb);
                color: white;
                border: none;
                border-radius: 14px;
                font-size: 16px;
                font-weight: 700;
                cursor: pointer;
                box-shadow: 0 10px 20px rgba(37, 99, 235, 0.35);
                transition: all 0.25s;
            }
            button:hover {
                transform: translateY(-2px);
                box-shadow: 0 14px 28px rgba(37, 99, 235, 0.5);
            }
            .error {
                background: rgba(239, 68, 68, 0.15);
                border: 1px solid rgba(239, 68, 68, 0.4);
                color: #fca5a5;
                padding: 12px;
                border-radius: 12px;
                margin-bottom: 20px;
                font-size: 14px;
                font-weight: 600;
            }
        </style>
    </head>
    <body>
        <div class="blob-1"></div>
        <div class="blob-2"></div>
        <div class="login-card">
            <div class="logo-circle">📚</div>
            <h2>Мектеп кітапханасын</h2>
            <p class="sub">Цифрлық басқару жүйесіне қош келдіңіз</p>
            {% if error %}
                <div class="error">⚠️ {{ error }}</div>
            {% endif %}
            <form method="post">
                <div class="input-group">
                    <label>Логин</label>
                    <input type="text" name="username" placeholder="Мысалы: library" required>
                </div>
                <div class="input-group">
                    <label>Құпиясөз</label>
                    <input type="password" name="password" placeholder="••••••••" required>
                </div>
                <button type="submit">Жүйеге кіру →</button>
            </form>
        </div>
    </body>
    </html>
    """, error=error)

def is_logged_in():
    return "username" in session

def is_librarian():
    return session.get("role") == "librarian"

# =========================================================
# БАСТЫ БЕТ (HOME DASHBOARD)
# =========================================================
@app.route("/")
def home():
    if "username" not in session:
        return redirect(url_for("login"))
    db = get_db()
    total_books = db.execute("SELECT COALESCE(SUM(jalpy_sany), 0) AS value FROM kitaptar").fetchone()["value"]
    total_readers = db.execute("SELECT COUNT(*) AS value FROM okuyandar WHERE (synyp NOT LIKE '0%' AND synyp != '0' AND LOWER(synyp) NOT LIKE '%даярлық%') OR turi = 'Мұғалім'").fetchone()["value"]
    currently_out = db.execute("SELECT COUNT(*) AS value FROM kitap_beru WHERE berildi = 'Иә' AND kaitaryldy = 'Жоқ'").fetchone()["value"]
    if DATABASE_URL:
        overdue = db.execute("SELECT COUNT(*) AS value FROM kitap_beru WHERE berildi = 'Иә' AND kaitaryldy = 'Жоқ' AND merzim::date < CURRENT_DATE").fetchone()["value"]
    else:
        overdue = db.execute("SELECT COUNT(*) FROM kitap_beru WHERE berildi = 'Иә' AND kaitaryldy = 'Жоқ' AND date(merzim) < date('now', 'localtime')").fetchone()[0]
    db.close()
    
    return render_template_string("""
    <!DOCTYPE html>
    <html lang="kk">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Кітапхана басқару жүйесі</title>
        <style>
            body {
                background: linear-gradient(180deg, #0f172a 0%, #1e293b 340px, #f1f5f9 340px, #f1f5f9 100%) !important;
            }
            .hero-banner {
                background: linear-gradient(135deg, rgba(79, 70, 229, 0.25), rgba(14, 165, 233, 0.15)),
                            url('/static/library-bg.jpg') center/cover no-repeat;
                border-radius: 26px;
                padding: 40px 35px;
                color: white;
                display: flex;
                align-items: center;
                gap: 28px;
                margin-bottom: 30px;
                box-shadow: 0 20px 40px rgba(0, 0, 0, 0.25);
                border: 1px solid rgba(255, 255, 255, 0.15);
                position: relative;
                overflow: hidden;
            }
            .hero-banner::before {
                content: "";
                position: absolute;
                inset: 0;
                background: rgba(15, 23, 42, 0.72);
                backdrop-filter: blur(6px);
                z-index: 0;
            }
            .hero-content {
                position: relative;
                z-index: 1;
                display: flex;
                align-items: center;
                gap: 25px;
                flex-wrap: wrap;
            }
            .header-logo {
                width: 95px;
                height: 95px;
                border-radius: 24px;
                object-fit: cover;
                background: white;
                padding: 4px;
                box-shadow: 0 10px 25px rgba(0,0,0,0.3);
            }
            .hero-text h1 {
                margin: 0 0 8px 0;
                font-size: 28px;
                font-weight: 800;
                letter-spacing: 0.5px;
            }
            .hero-text p {
                margin: 0;
                color: #cbd5e1;
                font-size: 15px;
            }

            /* Статистика карточкалары */
            .stats-grid {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
                gap: 20px;
                margin-bottom: 35px;
            }
            .stat-card {
                background: white;
                padding: 24px;
                border-radius: 22px;
                box-shadow: 0 10px 25px rgba(15, 23, 42, 0.06);
                border: 1px solid #e2e8f0;
                display: flex;
                align-items: center;
                justify-content: space-between;
                transition: transform 0.25s ease, box-shadow 0.25s ease;
            }
            .stat-card:hover {
                transform: translateY(-5px);
                box-shadow: 0 16px 35px rgba(15, 23, 42, 0.1);
            }
            .stat-info span {
                font-size: 13px;
                font-weight: 700;
                color: #64748b;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }
            .stat-info h3 {
                font-size: 36px;
                font-weight: 800;
                margin: 8px 0 0 0;
                color: #0f172a;
            }
            .stat-icon {
                width: 60px; height: 60px;
                border-radius: 18px;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 28px;
            }
            .ic-blue { background: #e0e7ff; color: #4f46e5; }
            .ic-green { background: #dcfce7; color: #16a34a; }
            .ic-amber { background: #fef3c7; color: #d97706; }
            .ic-red { background: #fee2e2; color: #dc2626; }

            /* Жылдам мәзір плиткалары */
            .section-heading {
                font-size: 18px;
                font-weight: 800;
                color: #1e293b;
                margin: 0 0 18px 0;
            }
            .menu-grid {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
                gap: 18px;
            }
            .menu-tile {
                background: white;
                padding: 22px;
                border-radius: 20px;
                border: 1px solid #e2e8f0;
                text-decoration: none;
                color: #0f172a;
                display: flex;
                align-items: center;
                gap: 18px;
                box-shadow: 0 4px 15px rgba(0, 0, 0, 0.03);
                transition: all 0.25s ease;
            }
            .menu-tile:hover {
                transform: translateY(-4px);
                border-color: #4f46e5;
                box-shadow: 0 12px 28px rgba(79, 70, 229, 0.12);
            }
            .tile-icon {
                width: 54px; height: 54px;
                border-radius: 16px;
                background: #f8fafc;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 26px;
                transition: all 0.25s;
            }
            .menu-tile:hover .tile-icon {
                background: #4f46e5;
                color: white;
                transform: scale(1.08);
            }
            .tile-text h4 {
                margin: 0 0 4px 0;
                font-size: 16px;
                font-weight: 700;
            }
            .tile-text p {
                margin: 0;
                font-size: 13px;
                color: #64748b;
            }
        </style>
    </head>
    <body>
        <div class="page-container">
            <div class="hero-banner">
                <div class="hero-content">
                    <img src="/static/emblem.png" onerror="this.style.display='none'" alt="Эмблема" class="header-logo">
                    <div class="hero-text">
                        <h1>МЕКТЕП КІТАПХАНАСЫ БАСҚАРУ ЖҮЙЕСІ</h1>
                        <p>Кітап қорын, оқырмандарды және кітап айналымын автоматтандырылған бақылау панелі</p>
                    </div>
                </div>
            </div>
            
            <div class="stats-grid">
                <div class="stat-card">
                    <div class="stat-info">
                        <span>Барлық кітап қоры</span>
                        <h3>{{ total_books }}</h3>
                    </div>
                    <div class="stat-icon ic-blue">📖</div>
                </div>
                <div class="stat-card">
                    <div class="stat-info">
                        <span>Тіркелген оқырмандар</span>
                        <h3>{{ total_readers }}</h3>
                    </div>
                    <div class="stat-icon ic-green">👥</div>
                </div>
                <div class="stat-card">
                    <div class="stat-info">
                        <span>Қазір қолда жүрген</span>
                        <h3>{{ currently_out }}</h3>
                    </div>
                    <div class="stat-icon ic-amber">📤</div>
                </div>
                <div class="stat-card">
                    <div class="stat-info">
                        <span>Мерзімі өткендер</span>
                        <h3 style="color:#dc2626;">{{ overdue }}</h3>
                    </div>
                    <div class="stat-icon ic-red">⚠️</div>
                </div>
            </div>
            
            <h3 class="section-heading">⚡ Негізгі бөлімдер</h3>
            <div class="menu-grid">
                <a class="menu-tile" href="/loans">
                    <div class="tile-icon">🔄</div>
                    <div class="tile-text">
                        <h4>Кітап беру және қайтару</h4>
                        <p>Оқырманға кітап рәсімдеу журналы</p>
                    </div>
                </a>
                <a class="menu-tile" href="/readers">
                    <div class="tile-icon">👥</div>
                    <div class="tile-text">
                        <h4>Оқырмандар тізімі</h4>
                        <p>Оқушылар мен мұғалімдер базасы</p>
                    </div>
                </a>
                <a class="menu-tile" href="/books">
                    <div class="tile-icon">📚</div>
                    <div class="tile-text">
                        <h4>Кітаптар қоры</h4>
                        <p>Барлық кітаптар мен қалдық саны</p>
                    </div>
                </a>
                <a class="menu-tile" href="/rating">
                    <div class="tile-icon">👨🎓</div>
                    <div class="tile-text">
                        <h4>Оқушылар рейтингісі</h4>
                        <p>Ең көп кітап оқитын үздік оқушылар</p>
                    </div>
                </a>
                <a class="menu-tile" href="/teacher-rating">
                    <div class="tile-icon">👩🏫</div>
                    <div class="tile-text">
                        <h4>Мұғалімдер рейтингісі</h4>
                        <p>Белсенді оқырман ұстаздар тізімі</p>
                    </div>
                </a>
                <a class="menu-tile" href="/book-rating">
                    <div class="tile-icon">🥇</div>
                    <div class="tile-text">
                        <h4>Кітаптар рейтингісі</h4>
                        <p>Сұранысы ең жоғары кітаптар</p>
                    </div>
                </a>
                <a class="menu-tile" href="/report">
                    <div class="tile-icon">📊</div>
                    <div class="tile-text">
                        <h4>Толық есептер</h4>
                        <p>Күндік, айлық және жалпы статистика</p>
                    </div>
                </a>
            </div>
        </div>
    </body>
    </html>
    """,
    total_books=total_books,
    total_readers=total_readers,
    currently_out=currently_out,
    overdue=overdue
    )

# =========================================================
# ЕСЕП БӨЛІМІ
# =========================================================
@app.route("/report")
def report():
    if not is_logged_in():
        return redirect(url_for("login"))
    db = get_db()
    if DATABASE_URL:
        today_sql = "kuni::date = CURRENT_DATE"
        month_sql = "DATE_TRUNC('month', kuni::date) = DATE_TRUNC('month', CURRENT_DATE)"
        overdue_sql = "merzim::date < CURRENT_DATE"
    else:
        today_sql = "date(kuni) = date('now', 'localtime')"
        month_sql = "strftime('%Y-%m', kuni) = strftime('%Y-%m', 'now', 'localtime')"
        overdue_sql = "date(merzim) < date('now', 'localtime')"
    
    def count_value(sql):
        return db.execute(sql).fetchone()["value"]
    
    today_issued = count_value(f"SELECT COUNT(*) AS value FROM kitap_beru WHERE {today_sql} AND berildi = 'Иә'")
    today_students = count_value(f"SELECT COUNT(DISTINCT kb.okuyman_id) AS value FROM kitap_beru kb JOIN okuyandar ok ON kb.okuyman_id = ok.id WHERE {today_sql.replace('kuni', 'kb.kuni')} AND kb.berildi = 'Иә' AND ok.turi = 'Оқушы'")
    today_teachers = count_value(f"SELECT COUNT(DISTINCT kb.okuyman_id) AS value FROM kitap_beru kb JOIN okuyandar ok ON kb.okuyman_id = ok.id WHERE {today_sql.replace('kuni', 'kb.kuni')} AND kb.berildi = 'Иә' AND ok.turi = 'Мұғалім'")
    total_books = count_value("SELECT COALESCE(SUM(jalpy_sany), 0) AS value FROM kitaptar")
    total_readers = count_value("SELECT COUNT(*) AS value FROM okuyandar WHERE (synyp NOT LIKE '0%' AND synyp != '0' AND LOWER(synyp) NOT LIKE '%даярлық%') OR turi = 'Мұғалім'")
    returned_books = count_value("SELECT COUNT(*) AS value FROM kitap_beru WHERE berildi = 'Иә' AND kaitaryldy = 'Иә'")
    currently_out = count_value("SELECT COUNT(*) AS value FROM kitap_beru WHERE berildi = 'Иә' AND kaitaryldy = 'Жоқ'")
    overdue = count_value(f"SELECT COUNT(*) AS value FROM kitap_beru WHERE berildi = 'Иә' AND kaitaryldy = 'Жоқ' AND merzim IS NOT NULL AND {overdue_sql}")
    this_month = count_value(f"SELECT COUNT(*) AS value FROM kitap_beru WHERE {month_sql} AND berildi = 'Иә'")
    db.close()
    
    return render_template_string("""
    <!DOCTYPE html>
    <html lang="kk">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Кітапхана есебі</title>
        <style>
            .cards {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
                gap: 20px;
            }
            .card {
                background: white;
                padding: 26px;
                border-radius: 20px;
                box-shadow: 0 8px 24px rgba(15, 23, 42, 0.05);
                border: 1px solid #e2e8f0;
                position: relative;
                overflow: hidden;
                transition: transform 0.2s;
            }
            .card::after {
                content: "";
                position: absolute;
                top: 0; left: 0; width: 6px; height: 100%;
                background: #4f46e5;
            }
            .card.danger-card::after { background: #ef4444; }
            .card.success-card::after { background: #10b981; }
            .card:hover { transform: translateY(-4px); }
            .card-title {
                color: #64748b;
                font-size: 13.5px;
                font-weight: 700;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }
            .number {
                font-size: 40px;
                font-weight: 800;
                color: #0f172a;
                margin-top: 12px;
            }
            .danger { color: #dc2626; }
            .success { color: #059669; }
        </style>
    </head>
    <body>
        <div class="page-container">
            <div class="page-header">
                <div>
                    <h1 class="page-title">📊 Кітапхананың толық есебі</h1>
                    <p class="page-subtitle">Кітап қоры мен оқырмандар белсенділігінің статистикалық көрсеткіштері</p>
                </div>
                <a class="btn-modern btn-outline" href="/">← Басты бетке</a>
            </div>

            <div class="cards">
                <div class="card">
                    <div class="card-title">📚 Барлық кітап саны</div>
                    <div class="number">{{ total_books }}</div>
                </div>
                <div class="card">
                    <div class="card-title">👥 Барлық оқырман саны</div>
                    <div class="number">{{ total_readers }}</div>
                </div>
                <div class="card success-card">
                    <div class="card-title">✅ Қайтарылған кітаптар</div>
                    <div class="number success">{{ returned_books }}</div>
                </div>
                <div class="card">
                    <div class="card-title">📚 Бүгін берілген кітаптар</div>
                    <div class="number">{{ today_issued }}</div>
                </div>
                <div class="card">
                    <div class="card-title">👨🎓 Бүгін кітап алған оқушылар</div>
                    <div class="number">{{ today_students }}</div>
                </div>
                <div class="card">
                    <div class="card-title">👩🏫 Бүгін кітап алған мұғалімдер</div>
                    <div class="number">{{ today_teachers }}</div>
                </div>
                <div class="card">
                    <div class="card-title">📤 Қазір қолда берілген</div>
                    <div class="number">{{ currently_out }}</div>
                </div>
                <div class="card danger-card">
                    <div class="card-title">⚠️ Қайтару мерзімі өткен</div>
                    <div class="number danger">{{ overdue }}</div>
                </div>
                <div class="card">
                    <div class="card-title">📅 Осы айда берілген кітаптар</div>
                    <div class="number">{{ this_month }}</div>
                </div>
            </div>
        </div>
    </body>
    </html>
    """,
    today_issued=today_issued,
    today_students=today_students,
    today_teachers=today_teachers,
    currently_out=currently_out,
    overdue=overdue,
    this_month=this_month,
    total_books=total_books,
    total_readers=total_readers,
    returned_books=returned_books,                             
    )

# =========================================================
# КІТАП БЕРУ (LOANS) БӨЛІМІ
# =========================================================
@app.route("/loans", methods=["GET", "POST"])
def loans():
    if not is_logged_in():
        return redirect(url_for("login"))
    db = get_db()
    
    if request.method == "POST" and is_librarian():
        okuyman_id = request.form.get("okuyman_id")
        kitap_id = request.form.get("kitap_id")
        kuni = request.form.get("kuni") or date.today().isoformat()
        merzim = request.form.get("merzim") or (date.fromisoformat(kuni) + timedelta(days=14)).isoformat()
        
        if okuyman_id and kitap_id:
            if DATABASE_URL:
                db.execute("""
                    INSERT INTO kitap_beru (kuni, okuyman_id, kitap_id, berildi, kaitaryldy, kaitaru_kuni, merzim)
                    VALUES (%s, %s, %s, 'Иә', 'Жоқ', '', %s)
                """, (kuni, okuyman_id, kitap_id, merzim))
            else:
                db.execute("""
                    INSERT INTO kitap_beru (kuni, okuyman_id, kitap_id, berildi, kaitaryldy, kaitaru_kuni, merzim)
                    VALUES (?, ?, ?, 'Иә', 'Жоқ', '', ?)
                """, (kuni, okuyman_id, kitap_id, merzim))
            db.commit()
        return redirect(url_for("loans"))
        
    loans_list = db.execute("""
        SELECT kb.id, kb.kuni, kb.merzim, kb.kaitaryldy,
               ok.aty as okuyhan_aty, ok.synyp, ok.turi,
               k.ataui as kitap_ataui, k.kitap_id
        FROM kitap_beru kb
        JOIN okuyandar ok ON kb.okuyman_id = ok.id
        JOIN kitaptar k ON kb.kitap_id = k.id
        ORDER BY kb.id DESC
    """).fetchall()
    
    readers_list = db.execute("""
        SELECT id, aty, synyp, turi FROM okuyandar
        WHERE (synyp NOT LIKE '0%' AND synyp != '0' AND LOWER(synyp) NOT LIKE '%даярлық%')
           OR turi = 'Мұғалім'
        ORDER BY aty ASC
    """).fetchall()
    
    books_list = db.execute("SELECT id, ataui, kitap_id FROM kitaptar ORDER BY ataui ASC").fetchall()
    db.close()
    
    return render_template_string("""
    <!DOCTYPE html>
    <html lang="kk">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Кітап беру және қайтару</title>
        <style>
            .form-card {
                background: white;
                padding: 28px;
                border-radius: 20px;
                box-shadow: 0 8px 25px rgba(15, 23, 42, 0.05);
                margin-bottom: 30px;
                border: 1px solid #e2e8f0;
            }
            .form-grid {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
                gap: 18px;
                align-items: end;
            }
            .field label {
                display: block;
                font-size: 13px;
                font-weight: 700;
                color: #475569;
                margin-bottom: 8px;
            }
            .field select, .field input {
                width: 100%;
                padding: 12px 14px;
                border: 1.5px solid #cbd5e1;
                border-radius: 12px;
                background: #f8fafc;
                font-size: 14px;
                color: #0f172a;
                transition: all 0.2s;
            }
            .field select:focus, .field input:focus {
                border-color: #4f46e5;
                background: #ffffff;
                outline: none;
                box-shadow: 0 0 0 4px rgba(79, 70, 229, 0.12);
            }
        </style>
    </head>
    <body>
        <div class="page-container">
            <div class="page-header">
                <div>
                    <h1 class="page-title">🔄 Кітап беру және қайтару</h1>
                    <p class="page-subtitle">Оқырмандарға берілген кітаптарды тіркеу және қайтарылуын бақылау</p>
                </div>
                <div class="search-box">
                    <span class="search-icon">🔍</span>
                    <input type="text" id="loanSearch" class="search-input" onkeyup="filterTable('loanSearch', 'loanTable')" placeholder="Оқырман немесе кітап атауы...">
                </div>
            </div>
            
            {% if session.get("role") == "librarian" %}
            <div class="form-card">
                <h3 style="margin:0 0 18px 0; color:#0f172a; font-size:18px;">➕ Жаңа кітап беруді рәсімдеу</h3>
                <form method="POST" class="form-grid">
                    <div class="field">
                        <label>Оқырманды таңдаңыз</label>
                        <select name="okuyman_id" required>
                            <option value="">-- Оқырманды таңдаңыз --</option>
                            {% for r in readers_list %}
                                <option value="{{ r['id'] }}">{{ r['aty'] }} ({{ r['synyp'] }})</option>
                            {% endfor %}
                        </select>
                    </div>

                    <div class="field">
                        <label>Кітапты таңдаңыз</label>
                        <select name="kitap_id" required>
                            <option value="">-- Кітапты таңдаңыз --</option>
                            {% for b in books_list %}
                                <option value="{{ b['id'] }}">{{ b['ataui'] }} (ID: {{ b['kitap_id'] }})</option>
                            {% endfor %}
                        </select>
                    </div>

                    <div class="field">
                        <label>Берілген күні</label>
                        <input type="date" name="kuni" value="{{ today }}">
                    </div>

                    <div class="field">
                        <label>Қайтару мерзімі</label>
                        <input type="date" name="merzim" value="{{ default_merzim }}">
                    </div>

                    <div class="field">
                        <button type="submit" class="btn-modern btn-primary" style="width:100%; padding:13px;">📥 Кітапты беру</button>
                    </div>
                </form>
            </div>
            {% endif %}

            <div class="table-card">
                <div class="table-responsive">
                    <table class="modern-table" id="loanTable">
                        <thead>
                            <tr>
                                <th>ID</th>
                                <th>Оқырман</th>
                                <th>Сыныбы / Пәні</th>
                                <th>Кітап атауы</th>
                                <th>Берілген күні</th>
                                <th>Қайтару мерзімі</th>
                                <th>Статус</th>
                                <th>Әрекет</th>
                            </tr>
                        </thead>
                        <tbody>
                            {% for l in loans_list %}
                            <tr>
                                <td style="color:#64748b; font-weight:600;">#{{ l['id'] }}</td>
                                <td><b>{{ l['okuyhan_aty'] }}</b></td>
                                <td><span class="badge badge-info">{{ l['synyp'] }}</span></td>
                                <td style="font-weight:600;">📖 {{ l['kitap_ataui'] }}</td>
                                <td>{{ l['kuni'] }}</td>
                                <td>{{ l['merzim'] }}</td>
                                <td>
                                    {% if l['kaitaryldy'] == 'Иә' %}
                                        <span class="badge badge-success">✅ Қайтарылды</span>
                                    {% else %}
                                        <span class="badge badge-danger">⏳ Қолда</span>
                                    {% endif %}
                                </td>
                                <td>
                                    {% if session.get("role") == "librarian" and l['kaitaryldy'] == 'Жоқ' %}
                                        <form method="POST" action="/loans/return/{{ l['id'] }}" style="margin:0;">
                                            <button type="submit" class="btn-modern" style="background:#10b981; color:white; padding:7px 14px; font-size:12.5px;">✓ Қайтару</button>
                                        </form>
                                    {% else %}
                                        <span style="color:#94a3b8; font-size:13px; font-weight:500;">Аяқталды</span>
                                    {% endif %}
                                </td>
                            </tr>
                            {% else %}
                            <tr>
                                <td colspan="8" style="text-align:center; padding:30px; color:#64748b;">Берілген кітаптар тізімі бос.</td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </body>
    </html>
    """,
    loans_list=loans_list,
    readers_list=readers_list,
    books_list=books_list,
    today=date.today().isoformat(),
    default_merzim=(date.today() + timedelta(days=14)).isoformat()
    )

@app.route("/loans/return/<int:loan_id>", methods=["POST"])
def return_book(loan_id):
    if not is_librarian():
        return "Бұл бөлімге тек кітапханашы кіре алады", 403
    db = get_db()
    today_str = date.today().isoformat()
    if DATABASE_URL:
        db.execute("UPDATE kitap_beru SET kaitaryldy = 'Иә', kaitaru_kuni = %s WHERE id = %s", (today_str, loan_id))
    else:
        db.execute("UPDATE kitap_beru SET kaitaryldy = 'Иә', kaitaru_kuni = ? WHERE id = ?", (today_str, loan_id))
    db.commit()
    db.close()
    return redirect(url_for("loans"))

# =========================================================
# ОҚЫРМАНДАР БӨЛІМІ
# =========================================================
@app.route("/readers/delete/<int:reader_id>", methods=["POST"])
def delete_reader(reader_id):
    if not is_librarian():
        return "Бұл бөлімге тек кітапханашы кіре алады", 403
    db = get_db()
    db.execute("DELETE FROM okuyandar WHERE id = %s" if DATABASE_URL else "DELETE FROM okuyandar WHERE id = ?", (reader_id,))
    db.commit()
    db.close()
    return redirect(url_for("readers"))

@app.route("/readers/reset", methods=["POST"])
def reset_readers():
    if not is_librarian():
        return "Бұл бөлімге тек кітапханашы кіре алады", 403
    db = get_db()
    try:
        db.execute("DELETE FROM kitap_beru")
        db.execute("DELETE FROM okuyandar")
        db.commit()
    finally:
        db.close()
    return redirect(url_for("readers"))

@app.route("/readers/import", methods=["GET", "POST"])
def readers_import():
    if not is_logged_in():
        return redirect(url_for("login"))
    if not is_librarian():
        return "Бұл бөлімге тек кітапханашы кіре алады", 403
    message = ""
    if request.method == "POST":
        file = request.files.get("file")
        if not file or not file.filename:
            message = "Excel файлын таңдаңыз."
        elif not file.filename.lower().endswith(".xlsx"):
            message = "Тек .xlsx форматындағы файлды жүктеңіз."
        else:
            db = get_db()
            try:
                df = pd.read_excel(file)
                df.columns = [re.sub(r"\s+", "", str(col)).casefold() for col in df.columns]
                name_col = next((c for c in df.columns if c in ("аты-жөні", "атыжөні", "атыжони")), None)
                class_col = next((c for c in df.columns if c in ("сыныбы/пәні", "сыныбы/пэні", "сыныбыпәні", "сыныбы", "сыныбы/пән")), None)
                if not name_col or not class_col:
                    message = "Excel бағандарын таба алмадым. Баған атаулары: Аты-жөні және Сыныбы/пәні болуы керек."
                else:
                    existing_rows = db.execute("SELECT aty, synyp FROM okuyandar").fetchall()
                    existing = {((r["aty"] or "").strip().casefold(), (r["synyp"] or "").strip().casefold()) for r in existing_rows}
                    added = 0
                    skipped = 0
                    for _, row in df.iterrows():
                        raw_name = row[name_col]
                        raw_class = row[class_col]
                        if pd.isna(raw_name) or pd.isna(raw_class):
                            skipped += 1
                            continue
                        name = re.sub(r"^\s*\d+\s*[\.\)]\s*", "", str(raw_name)).strip()
                        synyp = str(raw_class).strip()
                        if not name or not synyp:
                            skipped += 1
                            continue
                        if synyp.startswith("0") or "даярлық" in synyp.casefold():
                            skipped += 1
                            continue
                        key = (name.casefold(), synyp.casefold())
                        if key in existing:
                            skipped += 1
                            continue
                        if re.match(r"^\d{1,2}\s*[-]?\s*[а-яәғқңөұүһі]?$", synyp.casefold()):
                            turi = "Оқушы"
                        else:
                            turi = "Мұғалім"
                        if DATABASE_URL:
                            db.execute("INSERT INTO okuyandar (aty, synyp, turi, tirkelgen_kuni) VALUES (%s, %s, %s, %s)", (name, synyp, turi, "2026-09-01"))
                        else:
                            db.execute("INSERT INTO okuyandar (aty, synyp, turi, tirkelgen_kuni) VALUES (?, ?, ?, ?)", (name, synyp, turi, "2026-09-01"))
                        existing.add(key)
                        added += 1
                    db.commit()
                    message = f"Импорт аяқталды! Қосылды: {added}. Өткізілді немесе бұрыннан бар: {skipped}."
            except Exception as e:
                db.rollback()
                message = f"Импорт қатесі: {e}"
            finally:
                db.close()
    return render_template_string("""
    <!DOCTYPE html>
    <html lang="kk">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Оқырмандарды Excel-ден жүктеу</title>
        <style>
            .import-card {
                background: white;
                max-width: 580px;
                margin: 20px auto;
                padding: 35px;
                border-radius: 22px;
                box-shadow: 0 10px 30px rgba(15, 23, 42, 0.06);
                border: 1px solid #e2e8f0;
            }
            .file-drop {
                border: 2px dashed #94a3b8;
                border-radius: 16px;
                padding: 30px 20px;
                text-align: center;
                background: #f8fafc;
                margin: 20px 0;
                transition: border-color 0.2s;
            }
            .file-drop:hover { border-color: #4f46e5; }
            .message {
                margin-top: 20px;
                padding: 14px 18px;
                background: #eff6ff;
                border: 1px solid #bfdbfe;
                color: #1e40af;
                border-radius: 12px;
                font-weight: 600;
            }
        </style>
    </head>
    <body>
        <div class="page-container">
            <div class="import-card">
                <h2 style="margin-top:0;">📥 Оқырмандарды Excel-ден жүктеу</h2>
                <p style="color:#64748b; font-size:14.5px;">Жүктелетін <b>.xlsx</b> файлында келесі екі баған болуы қажет:</p>
                <div style="display:flex; gap:10px; margin-bottom:15px;">
                    <span class="badge badge-info">Аты-жөні</span>
                    <span class="badge badge-purple">Сыныбы/пәні</span>
                </div>
                <form method="POST" enctype="multipart/form-data">
                    <div class="file-drop">
                        <input type="file" name="file" accept=".xlsx" required style="font-size:14px;">
                    </div>
                    <button type="submit" class="btn-modern btn-primary" style="width:100%; padding:14px;">🚀 Файлды жүктеу және сақтау</button>
                </form>
                {% if message %}
                    <div class="message">{{ message }}</div>
                {% endif %}
                <div style="margin-top:22px;">
                    <a class="btn-modern btn-outline" href="/readers">← Оқырмандар тізіміне қайту</a>
                </div>
            </div>
        </div>
    </body>
    </html>
    """, message=message)

@app.route("/readers")
def readers():
    if not is_logged_in():
        return redirect(url_for("login"))
    db = get_db()
    rows = db.execute("SELECT id, aty FROM okuyandar").fetchall()
    for row in rows:
        old_name = row["aty"] or ""
        new_name = re.sub(r'^\s*\d+\s*[\.\)]\s*', '', old_name).strip()
        if new_name != old_name:
            if DATABASE_URL:
                db.execute("UPDATE okuyandar SET aty = %s WHERE id = %s", (new_name, row["id"]))
            else:
                db.execute("UPDATE okuyandar SET aty = ? WHERE id = ?", (new_name, row["id"]))
    db.commit()
    try:
        readers = db.execute("""
            SELECT * FROM okuyandar
            WHERE (synyp NOT LIKE '0%' AND synyp != '0' AND LOWER(synyp) NOT LIKE '%даярлық%')
               OR turi = 'Мұғалім'
            ORDER BY aty ASC
        """).fetchall()
    finally:
        db.close()
        
    return render_template_string("""
    <!DOCTYPE html>
    <html lang="kk">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Оқырмандар тізімі</title>
        <style>
            .toolbar {
                display: flex;
                justify-content: space-between;
                align-items: center;
                flex-wrap: wrap;
                gap: 12px;
                margin-bottom: 20px;
            }
            .delete-btn {
                background: #fee2e2;
                color: #dc2626;
                border: 0;
                padding: 7px 14px;
                border-radius: 9px;
                cursor: pointer;
                font-weight: 600;
                font-size: 13px;
                transition: all 0.2s;
            }
            .delete-btn:hover {
                background: #dc2626;
                color: white;
            }
        </style>
    </head>
    <body>
        <div class="page-container">
            <div class="page-header">
                <div>
                    <h1 class="page-title">👥 Оқырмандар тізімі</h1>
                    <p class="page-subtitle">Мектеп оқушылары мен мұғалімдерінің тізімі ({{ readers|length }} оқырман)</p>
                </div>
                <div class="search-box">
                    <span class="search-icon">🔍</span>
                    <input type="text" id="readerSearch" class="search-input" onkeyup="filterTable('readerSearch', 'readerTable')" placeholder="Аты-жөні немесе сыныбы бойынша іздеу...">
                </div>
            </div>

            {% if session.get("role") == "librarian" %}
            <div class="toolbar">
                <div style="display:flex; gap:12px; flex-wrap:wrap;">
                    <a class="btn-modern btn-primary" href="/readers/import">📥 Excel-ден жүктеу</a>
                    <form method="POST" action="{{ url_for('reset_readers') }}" style="margin:0;" onsubmit="return confirm('Барлық оқырман мен кітап беру тарихын өшіресіз бе?');">
                        <button type="submit" class="btn-modern btn-danger">🗑 Барлығын тазалау</button>
                    </form>
                </div>
            </div>
            {% endif %}

            <div class="table-card">
                <div class="table-responsive">
                    <table class="modern-table" id="readerTable">
                        <thead>
                            <tr>
                                <th>ID</th>
                                <th>Аты-жөні</th>
                                <th>Сыныбы / Пәні</th>
                                <th>Санаты</th>
                                {% if session.get("role") == "librarian" %}
                                    <th style="text-align:right;">Әрекет</th>
                                {% endif %}
                            </tr>
                        </thead>
                        <tbody>
                            {% for r in readers %}
                            <tr>
                                <td style="color:#64748b; font-weight:600;">#{{ r['id'] }}</td>
                                <td><b>{{ r['aty'] }}</b></td>
                                <td>{{ r['synyp'] }}</td>
                                <td>
                                    {% if r['turi'] == 'Мұғалім' %}
                                        <span class="badge badge-purple">👩🏫 Мұғалім</span>
                                    {% else %}
                                        <span class="badge badge-info">👨🎓 Оқушы</span>
                                    {% endif %}
                                </td>
                                {% if session.get("role") == "librarian" %}
                                <td style="text-align:right;">
                                    <form method="POST" action="/readers/delete/{{ r['id'] }}" style="margin:0;" onsubmit="return confirm('Осы оқырманды өшіресіз бе?');">
                                        <button type="submit" class="delete-btn">Өшіру</button>
                                    </form>
                                </td>
                                {% endif %}
                            </tr>
                            {% else %}
                            <tr>
                                <td colspan="5" style="text-align:center; padding:30px; color:#64748b;">Оқырмандар тізімі бос.</td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </body>
    </html>
    """, readers=readers)

# =========================================================
# КІТАПТАР ТІЗІМІ
# =========================================================
@app.route("/books")
def books():
    if not is_logged_in():
        return redirect(url_for("login"))
    db = get_db()
    try:
        books = db.execute("""
            SELECT kitaptar.*,
                (SELECT COUNT(*) FROM kitap_beru WHERE kitap_beru.kitap_id = kitaptar.id AND kitap_beru.kaitaryldy = 'Жоқ') AS berilgen_sany,
                (kitaptar.jalpy_sany - (SELECT COUNT(*) FROM kitap_beru WHERE kitap_beru.kitap_id = kitaptar.id AND kitap_beru.kaitaryldy = 'Жоқ')) AS korda_bar
            FROM kitaptar
            ORDER BY kitaptar.id DESC
        """).fetchall()
    finally:
        db.close()
    return render_template_string("""
    <!DOCTYPE html>
    <html lang="kk">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Кітаптар қоры</title>
    </head>
    <body>
        <div class="page-container">
            <div class="page-header">
                <div>
                    <h1 class="page-title">📚 Кітаптар қоры</h1>
                    <p class="page-subtitle">Кітапханадағы барлық кітаптар тізімі мен қолжетімді қалдықтар</p>
                </div>
                <div class="search-box">
                    <span class="search-icon">🔍</span>
                    <input type="text" id="bookSearch" class="search-input" onkeyup="filterTable('bookSearch', 'bookTable')" placeholder="Кітап атауы, авторы немесе ID...">
                </div>
            </div>

            <div class="table-card">
                <div class="table-responsive">
                    <table class="modern-table" id="bookTable">
                        <thead>
                            <tr>
                                <th>ID</th>
                                <th>Кітап коды</th>
                                <th>Кітап атауы</th>
                                <th>Авторы</th>
                                <th>Санаты</th>
                                <th>Жалпы саны</th>
                                <th>Берілгені</th>
                                <th>Қорда бар</th>
                            </tr>
                        </thead>
                        <tbody>
                            {% for book in books %}
                            <tr>
                                <td style="color:#64748b;">#{{ book["id"] }}</td>
                                <td><span class="badge badge-info">{{ book["kitap_id"] }}</span></td>
                                <td><b>{{ book["ataui"] }}</b></td>
                                <td>{{ book["avtor"] or '—' }}</td>
                                <td>{{ book["sanat"] or '—' }}</td>
                                <td style="font-weight:700;">{{ book["jalpy_sany"] }}</td>
                                <td><span class="badge badge-warning">{{ book["berilgen_sany"] }}</span></td>
                                <td><span class="badge badge-success">{{ book["korda_bar"] }}</span></td>
                            </tr>
                            {% else %}
                            <tr>
                                <td colspan="8" style="text-align:center; padding:30px; color:#64748b;">Кітаптар тізімі бос.</td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </body>
    </html>
    """, books=books)

# =========================================================
# РЕЙТИНГТЕР (ОҚУШЫЛАР, МҰҒАЛІМДЕР, КІТАПТАР)
# =========================================================
@app.route("/rating")
def rating():
    if not is_logged_in():
        return redirect("/login")
    db = get_db()
    cursor = db.cursor()
    cursor.execute("""
        SELECT ok.aty, ok.synyp, COUNT(kb.id) AS kitap_sany
        FROM okuyandar ok
        LEFT JOIN kitap_beru kb ON kb.okuyman_id = ok.id
        WHERE ok.turi != 'Мұғалім' OR ok.turi IS NULL
        GROUP BY ok.id, ok.aty, ok.synyp
        ORDER BY kitap_sany DESC, ok.aty
    """)
    rows = cursor.fetchall()
    db.close()
    
    return render_template_string("""
    <!DOCTYPE html>
    <html lang="kk">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Оқушылар рейтингісі</title>
    </head>
    <body>
        <div class="page-container">
            <div class="page-header">
                <div>
                    <h1 class="page-title">👨🎓 Оқушылар рейтингісі</h1>
                    <p class="page-subtitle">Кітап алу белсенділігі бойынша үздік оқушылар тізімі</p>
                </div>
                <div class="search-box">
                    <span class="search-icon">🔍</span>
                    <input type="text" id="stRateSearch" class="search-input" onkeyup="filterTable('stRateSearch', 'stRateTable')" placeholder="Оқушыны іздеу...">
                </div>
            </div>

            <div class="table-card">
                <div class="table-responsive">
                    <table class="modern-table" id="stRateTable">
                        <thead>
                            <tr>
                                <th style="width:90px;">Орын</th>
                                <th>Оқушының аты-жөні</th>
                                <th>Сыныбы</th>
                                <th>Алған кітап саны</th>
                            </tr>
                        </thead>
                        <tbody>
                            {% for row in rows %}
                            <tr>
                                <td style="font-weight:800; font-size:16px;">
                                    {% if loop.index == 1 %}🥇 1
                                    {% elif loop.index == 2 %}🥈 2
                                    {% elif loop.index == 3 %}🥉 3
                                    {% else %}{{ loop.index }}{% endif %}
                                </td>
                                <td><b>{{ row['aty'] }}</b></td>
                                <td><span class="badge badge-info">{{ row['synyp'] or '—' }}</span></td>
                                <td><span class="badge badge-success" style="font-size:14px;">📚 {{ row['kitap_sany'] }} кітап</span></td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </body>
    </html>
    """, rows=rows)

@app.route("/teacher-rating")
def teacher_rating():
    if not is_logged_in():
        return redirect("/login")
    db = get_db()
    cursor = db.cursor()
    cursor.execute("""
        SELECT ok.aty, ok.synyp, COUNT(kb.id) AS kitap_sany
        FROM okuyandar ok
        LEFT JOIN kitap_beru kb ON kb.okuyman_id = ok.id
        WHERE ok.turi = 'Мұғалім'
        GROUP BY ok.id, ok.aty, ok.synyp
        ORDER BY kitap_sany DESC, ok.aty
    """)
    rows = cursor.fetchall()
    db.close()
    
    return render_template_string("""
    <!DOCTYPE html>
    <html lang="kk">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Мұғалімдер рейтингісі</title>
    </head>
    <body>
        <div class="page-container">
            <div class="page-header">
                <div>
                    <h1 class="page-title">👩🏫 Мұғалімдер рейтингісі</h1>
                    <p class="page-subtitle">Кітап алу белсенділігі бойынша ұстаздар рейтингі</p>
                </div>
                <div class="search-box">
                    <span class="search-icon">🔍</span>
                    <input type="text" id="tcRateSearch" class="search-input" onkeyup="filterTable('tcRateSearch', 'tcRateTable')" placeholder="Мұғалімді іздеу...">
                </div>
            </div>

            <div class="table-card">
                <div class="table-responsive">
                    <table class="modern-table" id="tcRateTable">
                        <thead>
                            <tr>
                                <th style="width:90px;">Орын</th>
                                <th>Мұғалімнің аты-жөні</th>
                                <th>Пәні / Қызметі</th>
                                <th>Алған кітап саны</th>
                            </tr>
                        </thead>
                        <tbody>
                            {% for row in rows %}
                            <tr>
                                <td style="font-weight:800; font-size:16px;">
                                    {% if loop.index == 1 %}🥇 1
                                    {% elif loop.index == 2 %}🥈 2
                                    {% elif loop.index == 3 %}🥉 3
                                    {% else %}{{ loop.index }}{% endif %}
                                </td>
                                <td><b>{{ row['aty'] }}</b></td>
                                <td><span class="badge badge-purple">{{ row['synyp'] or '—' }}</span></td>
                                <td><span class="badge badge-success" style="font-size:14px;">📚 {{ row['kitap_sany'] }} кітап</span></td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </body>
    </html>
    """, rows=rows)

@app.route("/book-rating")
def book_rating():
    if not is_logged_in():
        return redirect("/login")
    db = get_db()
    cursor = db.cursor()
    cursor.execute("""
        SELECT k.kitap_id, k.ataui, k.avtor, COUNT(kb.id) AS alu_sany
        FROM kitaptar k
        LEFT JOIN kitap_beru kb ON kb.kitap_id = k.id
        GROUP BY k.id, k.kitap_id, k.ataui, k.avtor
        ORDER BY alu_sany DESC, k.ataui
    """)
    rows = cursor.fetchall()
    db.close()
    
    return render_template_string("""
    <!DOCTYPE html>
    <html lang="kk">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Кітаптар рейтингісі</title>
    </head>
    <body>
        <div class="page-container">
            <div class="page-header">
                <div>
                    <h1 class="page-title">🥇 Кітаптар рейтингісі</h1>
                    <p class="page-subtitle">Оқырмандар арасында ең көп сұранысқа ие кітаптар</p>
                </div>
                <div class="search-box">
                    <span class="search-icon">🔍</span>
                    <input type="text" id="bkRateSearch" class="search-input" onkeyup="filterTable('bkRateSearch', 'bkRateTable')" placeholder="Кітапты іздеу...">
                </div>
            </div>

            <div class="table-card">
                <div class="table-responsive">
                    <table class="modern-table" id="bkRateTable">
                        <thead>
                            <tr>
                                <th style="width:90px;">Орын</th>
                                <th>Кітап ID</th>
                                <th>Кітап атауы</th>
                                <th>Авторы</th>
                                <th>Алынған саны</th>
                            </tr>
                        </thead>
                        <tbody>
                            {% for row in rows %}
                            <tr>
                                <td style="font-weight:800; font-size:16px;">
                                    {% if loop.index == 1 %}🥇 1
                                    {% elif loop.index == 2 %}🥈 2
                                    {% elif loop.index == 3 %}🥉 3
                                    {% else %}{{ loop.index }}{% endif %}
                                </td>
                                <td><span class="badge badge-info">{{ row['kitap_id'] }}</span></td>
                                <td><b>{{ row['ataui'] }}</b></td>
                                <td>{{ row['avtor'] or '—' }}</td>
                                <td><span class="badge badge-success" style="font-size:14px;">🔥 {{ row['alu_sany'] }} рет</span></td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </body>
    </html>
    """, rows=rows)

if __name__ == '__main__':
    init_db()
    app.run(debug=True)
