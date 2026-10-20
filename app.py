from flask import Flask, request, redirect, url_for, render_template_string, session
import sqlite3
import os
import re
from datetime import date, timedelta
import pandas as pd

app = Flask(__name__)
app.secret_key = "KITAPHANA_READING_NATION_2026"

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

def is_logged_in():
    return "username" in session

def is_librarian():
    return session.get("role") == "librarian"

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("home"))

@app.route("/login", methods=["GET", "POST"])
def login():
    error = ""
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
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
        <title>Жүйеге кіру — Оқитын ұлт</title>
        <style>
            body {
                font-family: 'Inter', -apple-system, sans-serif;
                background: #0a1329;
                color: #ffffff;
                display: flex;
                justify-content: center;
                align-items: center;
                min-height: 100vh;
                margin: 0;
            }
            .login-card {
                background: rgba(15, 23, 42, 0.9);
                border: 1px solid rgba(255, 255, 255, 0.15);
                border-radius: 20px;
                padding: 40px;
                width: 100%;
                max-width: 400px;
                box-shadow: 0 20px 50px rgba(0, 0, 0, 0.5);
                text-align: center;
            }
            .login-card h2 {
                color: #f59e0b;
                font-weight: 800;
                margin-bottom: 25px;
            }
            input {
                width: 100%;
                padding: 14px 16px;
                margin-bottom: 18px;
                border-radius: 12px;
                border: 1px solid #334155;
                background: #1e293b;
                color: white;
                font-size: 15px;
                box-sizing: border-box;
            }
            input:focus {
                border-color: #f59e0b;
                outline: none;
            }
            button {
                width: 100%;
                padding: 14px;
                border-radius: 12px;
                border: none;
                background: #f59e0b;
                color: #0f172a;
                font-weight: 700;
                font-size: 16px;
                cursor: pointer;
                transition: 0.2s;
            }
            button:hover {
                background: #d97706;
            }
            .error {
                background: rgba(239, 68, 68, 0.2);
                border: 1px solid #ef4444;
                color: #fca5a5;
                padding: 10px;
                border-radius: 10px;
                margin-bottom: 20px;
                font-size: 14px;
            }
        </style>
    </head>
    <body>
        <div class="login-card">
            <h2>📖 ОҚИТЫН ҰЛТ</h2>
            {% if error %}<div class="error">{{ error }}</div>{% endif %}
            <form method="post">
                <input type="text" name="username" placeholder="Логин" required>
                <input type="password" name="password" placeholder="Құпиясөз" required>
                <button type="submit">Жүйеге кіру</button>
            </form>
        </div>
    </body>
    </html>
    """, error=error)

# =========================================================
# БАСТЫ БЕТ - «ОҚИТЫН ҰЛТ» СТИЛІНДЕГІ ДИЗАЙН
# =========================================================
@app.route("/")
def home():
    db = get_db()
    total_books = db.execute("SELECT COALESCE(SUM(jalpy_sany), 0) AS value FROM kitaptar").fetchone()["value"]
    total_readers = db.execute("SELECT COUNT(*) AS value FROM okuyandar WHERE (synyp NOT LIKE '0%' AND synyp != '0' AND LOWER(synyp) NOT LIKE '%даярлық%') OR turi = 'Мұғалім'").fetchone()["value"]
    currently_out = db.execute("SELECT COUNT(*) AS value FROM kitap_beru WHERE berildi = 'Иә' AND kaitaryldy = 'Жоқ'").fetchone()["value"]
    db.close()
    
    return render_template_string("""
    <!DOCTYPE html>
    <html lang="kk">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>«Оқитын ұлт» Республикалық кітап оқу марафоны</title>
        <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800;900&display=swap" rel="stylesheet">
        <style>
            * {
                box-sizing: border-box;
                margin: 0;
                padding: 0;
            }
            body {
                font-family: 'Montserrat', sans-serif;
                background-color: #070e20;
                color: #ffffff;
                min-height: 100vh;
                overflow-x: hidden;
            }

            /* Шапка (Header/Navbar) */
            header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 20px 60px;
                background: rgba(7, 14, 32, 0.95);
                backdrop-filter: blur(10px);
                position: sticky;
                top: 0;
                z-index: 1000;
                border-bottom: 1px solid rgba(255, 255, 255, 0.08);
            }
            .logo {
                display: flex;
                align-items: center;
                gap: 12px;
                font-weight: 900;
                font-size: 20px;
                color: #ffffff;
                text-decoration: none;
                letter-spacing: 1px;
            }
            .logo span {
                color: #f59e0b;
            }
            .logo-icon {
                font-size: 28px;
            }
            nav {
                display: flex;
                align-items: center;
                gap: 30px;
            }
            nav a {
                color: #cbd5e1;
                text-decoration: none;
                font-size: 14px;
                font-weight: 600;
                transition: color 0.3s;
            }
            nav a:hover {
                color: #f59e0b;
            }
            .header-right {
                display: flex;
                align-items: center;
                gap: 20px;
            }
            .lang-select {
                font-size: 13px;
                font-weight: 700;
                color: #94a3b8;
            }
            .lang-select span.active {
                color: #ffffff;
            }
            .btn-login {
                padding: 10px 22px;
                border: 1px solid rgba(255, 255, 255, 0.3);
                border-radius: 8px;
                color: white;
                text-decoration: none;
                font-weight: 700;
                font-size: 14px;
                transition: 0.3s;
            }
            .btn-login:hover {
                background: rgba(255, 255, 255, 0.1);
            }
            .btn-start {
                padding: 10px 24px;
                background: #f59e0b;
                border-radius: 8px;
                color: #070e20;
                text-decoration: none;
                font-weight: 800;
                font-size: 14px;
                transition: 0.3s;
            }
            .btn-start:hover {
                background: #d97706;
            }

            /* Басты марафон бөлімі (Hero) */
            .hero-container {
                position: relative;
                padding: 60px 60px 80px 60px;
                background: linear-gradient(180deg, #070e20 0%, #0a1738 100%);
                display: flex;
                align-items: center;
                justify-content: space-between;
                min-height: 520px;
                background-image: radial-gradient(circle at 20% 30%, rgba(30, 58, 138, 0.3) 0%, transparent 50%);
            }
            .hero-content {
                max-width: 650px;
                z-index: 2;
            }
            .hero-title {
                font-size: 42px;
                font-weight: 900;
                line-height: 1.25;
                color: #f59e0b;
                text-transform: uppercase;
                letter-spacing: 1.5px;
                margin-bottom: 30px;
                text-shadow: 0 4px 20px rgba(0,0,0,0.5);
            }
            
            /* Прогресс бары */
            .progress-block {
                margin-bottom: 35px;
                max-width: 500px;
            }
            .progress-info {
                font-size: 11px;
                font-weight: 700;
                color: #94a3b8;
                text-transform: uppercase;
                letter-spacing: 1px;
                margin-bottom: 10px;
            }
            .progress-bar-bg {
                width: 100%;
                height: 12px;
                background: rgba(255, 255, 255, 0.1);
                border-radius: 10px;
                position: relative;
                overflow: visible;
            }
            .progress-bar-fill {
                width: 20%;
                height: 100%;
                background: linear-gradient(90deg, #d97706, #f59e0b);
                border-radius: 10px;
                position: relative;
            }
            .progress-badge {
                position: absolute;
                right: -15px;
                top: -10px;
                background: #f59e0b;
                color: #070e20;
                font-size: 11px;
                font-weight: 900;
                padding: 4px 8px;
                border-radius: 20px;
                box-shadow: 0 4px 10px rgba(0,0,0,0.3);
            }

            .hero-buttons {
                display: flex;
                gap: 15px;
            }
            .btn-main {
                padding: 16px 32px;
                background: #f59e0b;
                color: #070e20;
                text-decoration: none;
                font-weight: 800;
                border-radius: 8px;
                font-size: 15px;
                transition: 0.3s;
                display: inline-flex;
                align-items: center;
                gap: 8px;
            }
            .btn-main:hover {
                background: #e08e00;
                transform: translateY(-2px);
            }
            .btn-secondary {
                padding: 16px 32px;
                background: #ffffff;
                color: #070e20;
                text-decoration: none;
                font-weight: 800;
                border-radius: 8px;
                font-size: 15px;
                transition: 0.3s;
                display: inline-flex;
                align-items: center;
                gap: 8px;
            }
            .btn-secondary:hover {
                background: #f1f5f9;
                transform: translateY(-2px);
            }

            .hero-image {
                position: relative;
                z-index: 2;
                max-width: 500px;
            }
            .hero-image img {
                width: 100%;
                border-radius: 20px;
                box-shadow: 0 20px 50px rgba(0,0,0,0.6);
            }

            /* Ұлттық ойю-өрнек сәндік жолағы */
            .ornament-strip {
                width: 100%;
                height: 35px;
                background-color: #f59e0b;
                background-image: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="60" height="35" viewBox="0 0 60 35"><path d="M30 0 C20 10 10 0 0 17.5 C10 35 20 25 30 35 C40 25 50 35 60 17.5 C50 0 40 10 30 0 Z" fill="none" stroke="%20070e20" stroke-width="3"/></svg>');
                background-repeat: repeat-x;
                background-size: 60px 35px;
                box-shadow: 0 4px 15px rgba(0,0,0,0.4);
            }

            /* Басқару панелі және статистика */
            .main-content {
                padding: 60px;
                max-width: 1400px;
                margin: 0 auto;
            }
            .section-title {
                font-size: 24px;
                font-weight: 800;
                color: #ffffff;
                margin-bottom: 30px;
                display: flex;
                align-items: center;
                gap: 12px;
            }

            /* Статистика көрсеткіштері */
            .stats-grid {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
                gap: 20px;
                margin-bottom: 50px;
            }
            .stat-box {
                background: rgba(15, 23, 42, 0.8);
                border: 1px solid rgba(255, 255, 255, 0.08);
                border-radius: 16px;
                padding: 25px;
                text-align: center;
                backdrop-filter: blur(10px);
            }
            .stat-number {
                font-size: 42px;
                font-weight: 900;
                color: #f59e0b;
                margin-bottom: 6px;
            }
            .stat-label {
                font-size: 13px;
                font-weight: 700;
                color: #94a3b8;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }

            /* Навигациялық блогтар (Батырмасыз карталар) */
            .cards-grid {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
                gap: 25px;
            }
            .nav-card {
                background: rgba(15, 23, 42, 0.7);
                border: 1px solid rgba(255, 255, 255, 0.08);
                border-radius: 20px;
                padding: 30px;
                text-decoration: none;
                color: white;
                display: flex;
                align-items: flex-start;
                gap: 20px;
                transition: all 0.3s;
            }
            .nav-card:hover {
                background: rgba(30, 41, 59, 0.9);
                border-color: #f59e0b;
                transform: translateY(-5px);
                box-shadow: 0 15px 30px rgba(0,0,0,0.4);
            }
            .card-icon {
                font-size: 36px;
                background: rgba(245, 158, 11, 0.1);
                color: #f59e0b;
                width: 65px;
                height: 65px;
                border-radius: 16px;
                display: flex;
                align-items: center;
                justify-content: center;
                flex-shrink: 0;
            }
            .card-info h3 {
                font-size: 18px;
                font-weight: 800;
                margin-bottom: 8px;
                color: #ffffff;
            }
            .card-info p {
                font-size: 13px;
                color: #94a3b8;
                line-height: 1.5;
            }

            @media (max-width: 900px) {
                header, .hero-container, .main-content {
                    padding: 30px 20px;
                }
                .hero-container {
                    flex-direction: column;
                    text-align: center;
                }
                .hero-title {
                    font-size: 30px;
                }
                .hero-image {
                    margin-top: 40px;
                }
            }
        </style>
    </head>
    <body>

        <!-- Жоғарғы мәзір (Header) -->
        <header>
            <a href="/" class="logo">
                <span class="logo-icon">📖</span>
                <div>КІТАП <span>ОҚИУЫН ҰЛТ</span></div>
            </a>
            <nav>
                <a href="#">О проекте</a>
                <a href="/books">Книги</a>
                <a href="#">Правила</a>
                <a href="/rating">Рейтинг</a>
                <a href="#">Отзывы</a>
                <a href="#">Новости</a>
            </nav>
            <div class="header-right">
                <div class="lang-select">
                    <span class="active">Қаз</span> | <span>Рус</span>
                </div>
                {% if session.get('username') %}
                    <a href="/logout" class="btn-login">🚪 Шығу</a>
                {% else %}
                    <a href="/login" class="btn-login">Войти</a>
                {% endif %}
                <a href="/loans" class="btn-start">▸ Старт</a>
            </div>
        </header>

        <!-- Главный баннер (Hero section) -->
        <div class="hero-container">
            <div class="hero-content">
                <h1 class="hero-title">
                    Республиканский<br>марафон чтения книг<br>«Читающая нация»
                </h1>
                
                <div class="progress-block">
                    <div class="progress-info">Прогресс марафона | Сентябрь 2026 - Март 2027 | Осталось: 140 дней</div>
                    <div class="progress-bar-bg">
                        <div class="progress-bar-fill">
                            <div class="progress-badge">20%</div>
                        </div>
                    </div>
                </div>

                <div class="hero-buttons">
                    <a href="/loans" class="btn-main">Начать участие ▸</a>
                    <a href="/report" class="btn-secondary">Узнать о проекте ▸</a>
                </div>
            </div>

            <div class="hero-image">
                <img src="https://images.unsplash.com/photo-1544717305-2782549b5136?auto=format&fit=crop&w=800&q=80" alt="Отбасылық кітап оқу">
            </div>
        </div>

        <!-- Алтын ұлттық ою-өрнек жолағы -->
        <div class="ornament-strip"></div>

        <!-- Негізгі навигация және жүйелік модульдер -->
        <div class="main-content">
            <div class="section-title">📊 Марафон көрсеткіштері</div>
            
            <div class="stats-grid">
                <div class="stat-box">
                    <div class="stat-number">{{ total_books }}</div>
                    <div class="stat-label">Барлық кітап қоры</div>
                </div>
                <div class="stat-box">
                    <div class="stat-number">{{ total_readers }}</div>
                    <div class="stat-label">Тіркелген оқырмандар</div>
                </div>
                <div class="stat-box">
                    <div class="stat-number">{{ currently_out }}</div>
                    <div class="stat-label">Қазір оқылудағы кітаптар</div>
                </div>
            </div>

            <div class="section-title">🚀 Жүйені басқару бөлімдері</div>

            <div class="cards-grid">
                <a href="/readers" class="nav-card">
                    <div class="card-icon">👥</div>
                    <div class="card-info">
                        <h3>Оқырмандар тізімі</h3>
                        <p>Оқушылар мен мұғалімдердің толық базасын қарау және басқару</p>
                    </div>
                </a>

                <a href="/books" class="nav-card">
                    <div class="card-icon">📚</div>
                    <div class="card-info">
                        <h3>Кітаптар қоры</h3>
                        <p>Кітапхана каталогы, авторы және дана санын жүргізу</p>
                    </div>
                </a>

                <a href="/loans" class="nav-card">
                    <div class="card-icon">🔄</div>
                    <div class="card-info">
                        <h3>Кітап беру және қайтару</h3>
                        <p>Оқырмандарға кітап беріп, қайтару мерзімін бақылау</p>
                    </div>
                </a>

                <a href="/rating" class="nav-card">
                    <div class="card-icon">👨‍🎓</div>
                    <div class="card-info">
                        <h3>Оқушылар рейтингісі</h3>
                        <p>Ең көп кітап оқыған белсенді оқушылар топтамасы</p>
                    </div>
                </a>

                <a href="/teacher-rating" class="nav-card">
                    <div class="card-icon">👩‍🏫</div>
                    <div class="card-info">
                        <h3>Мұғалімдер рейтингісі</h3>
                        <p>Марафонға белсене қатысқан ұстаздар тізімі</p>
                    </div>
                </a>

                <a href="/report" class="nav-card">
                    <div class="card-icon">📊</div>
                    <div class="card-info">
                        <h3>Толық есептер</h3>
                        <p>Кітапхана жұмысы бойынша статистика мен есептер</p>
                    </div>
                </a>
            </div>
        </div>

    </body>
    </html>
    """,
    total_books=total_books,
    total_readers=total_readers,
    currently_out=currently_out
    )

if __name__ == "__main__":
    app.run(debug=True, port=5000)
