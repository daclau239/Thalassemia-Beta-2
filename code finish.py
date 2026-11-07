import io
import json
import math
import os
import re
import hashlib
import secrets
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from urllib.parse import quote

import requests
import streamlit as st
from docx import Document

# ===== CODE BY NGUYEN DAC LAU 24XN1 =====
# Yêu cầu: streamlit>=1.45, psycopg[binary], python-docx, XlsxWriter, requests

# set_page_config PHẢI là lệnh Streamlit đầu tiên.
st.set_page_config(
    page_title="Hệ thống sàng lọc Thalassemia",
    page_icon=":material/bloodtype:",
    layout="wide",
)

st.markdown("""
<style>
/* ===== Professional Hematology UI — adaptive Light / Dark ===== */
:root {
  --heme-navy:#17324D;
  --heme-red:#8E2C3A;
  --heme-accent:#A33A49;
  --heme-border:color-mix(in srgb, var(--text-color) 18%, transparent);
  --heme-muted:color-mix(in srgb, var(--text-color) 66%, transparent);
  --heme-surface:var(--secondary-background-color);
  --heme-surface-soft:color-mix(in srgb, var(--secondary-background-color) 86%, var(--background-color));
}
.main .block-container { max-width:1180px; padding-top:2rem; padding-bottom:3rem; }
[data-testid="stHeader"] { background:transparent; }
.hematology-brand { border-left:5px solid var(--heme-red); padding:.15rem 0 .15rem 1rem; margin-bottom:.35rem; }
.hematology-brand .eyebrow { font-size:.78rem; letter-spacing:.16em; color:var(--heme-muted); font-weight:700; text-transform:uppercase; }
.hematology-brand .title { font-size:2rem; line-height:1.15; color:var(--text-color); font-weight:750; margin-top:.18rem; }
.hematology-brand .title-main { font-size:2.35rem; font-weight:850; letter-spacing:-0.035em; line-height:1.12; text-transform:uppercase; }
.hematology-brand .subtitle { color:var(--heme-muted); font-size:.95rem; margin-top:.45rem; }
.section-rule { height:1px; background:var(--heme-border); margin:1.2rem 0; }
[data-testid="stExpander"] { border:1px solid var(--heme-border); border-radius:10px; background:var(--heme-surface-soft); }
[data-testid="stExpander"] summary p { color:var(--text-color); font-weight:700; }
[data-testid="stSidebar"] { border-right:1px solid var(--heme-border); }
[data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { color:var(--text-color); }
.stButton > button { border-radius:7px; font-weight:600; }
.contact-card { border:1px solid var(--heme-border); border-radius:10px; padding:1rem 1.15rem; background:var(--heme-surface-soft); margin-top:1rem; }
.contact-card .label { color:var(--heme-muted); font-size:.78rem; text-transform:uppercase; letter-spacing:.08em; font-weight:700; }
.contact-card .name { color:var(--text-color); font-size:1.1rem; font-weight:750; margin:.2rem 0 .55rem; }
.contact-card a { color:var(--heme-accent); text-decoration:none; font-weight:600; }
.contact-card a:hover { text-decoration:underline; }
.small-note { color:var(--heme-muted); font-size:.82rem; }

/* Huy hiệu mức nguy cơ */
.risk-badge { display:inline-block; padding:.2rem .7rem; border-radius:999px; font-size:.78rem; font-weight:700; letter-spacing:.06em; }
.risk-low  { background:color-mix(in srgb,#2E7D32 16%,transparent); color:#2E7D32; }
.risk-mid  { background:color-mix(in srgb,#B26A00 16%,transparent); color:#B26A00; }
.risk-high { background:color-mix(in srgb,#B3261E 16%,transparent); color:#B3261E; }

@media (prefers-color-scheme: dark) {
  :root { --heme-red:#C45A69; --heme-accent:#E07987; }
  .risk-low  { color:#7BC47F; }
  .risk-mid  { color:#E5A64B; }
  .risk-high { color:#F2877F; }
  [data-testid="stExpander"], .contact-card { box-shadow:0 1px 0 rgba(255,255,255,.025) inset; }
}
</style>
""", unsafe_allow_html=True)

# ============================================================
# THALASSEMIA SCREENING PREVIEW
# ============================================================
# 1) Hồ sơ bệnh nhân
# 2) Vòng 1: 20 câu hỏi (phân tầng + khuyến nghị; mọi người đều được vào Vòng 2)
# 3) Vòng 2: độ cao + CBC, hiệu chỉnh Hb (WHO 2024), Mentzer, khuyến nghị
# 4) Vòng 3: kết quả xét nghiệm chuyên sâu / theo dõi
#
# IMPORTANT:
# - Điểm nguy cơ là prototype, chưa validation trên người Việt Nam.
# - "Khuyến nghị" là hỗ trợ sàng lọc, không chẩn đoán.
# - Q19/Q20 về tiếp cận xét nghiệm = 0 điểm.
# - Một số điện thoại = một hồ sơ; mỗi lần sàng lọc là một bản ghi lịch sử riêng.
#
# SECRETS cần cấu hình (Streamlit Cloud → Settings → Secrets):
#   DB_HOST, DB_PORT, DB_NAME, DB_USER, DB_PASSWORD   (hoặc DATABASE_URL)
#   ADMIN_USERNAME, ADMIN_PASSWORD                    (tài khoản quản trị đầu tiên)
#   ADMIN_FULL_NAME, ADMIN_EMAIL                      (tuỳ chọn)
# ============================================================


# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------

def _secret(name, default=""):
    try:
        value = st.secrets.get(name, default)
    except Exception:
        value = default
    if value is None or value == "":
        value = os.environ.get(name, default)
    return str(value).strip() if value is not None else ""


DB_HOST = _secret("DB_HOST")
DB_PORT = int(_secret("DB_PORT", "6543") or "6543")
DB_NAME = _secret("DB_NAME", "postgres")
DB_USER = _secret("DB_USER")
DB_PASSWORD = _secret("DB_PASSWORD")
DATABASE_URL = _secret("DATABASE_URL") or _secret("SUPABASE_DB_URL")

# Tài khoản quản trị lấy từ secrets — KHÔNG ghi mật khẩu trong code.
ADMIN_USERNAME = _secret("ADMIN_USERNAME")
ADMIN_PASSWORD = _secret("ADMIN_PASSWORD")
ADMIN_FULL_NAME = _secret("ADMIN_FULL_NAME", "Quản trị viên hệ thống")
ADMIN_EMAIL = _secret("ADMIN_EMAIL", "admin@thalassemia.local")

LOCAL_SQLITE_PATH = "thalassemia_patients.db"
CONSENT_VERSION = "DACLAU239-BETA5"
ADMIN_DATA_URL = "https://raw.githubusercontent.com/open-admin-data/vietnam-administrative-divisions/main/data/hierarchy.json"
ADMIN_DATA_SOURCE_URL = "https://github.com/open-admin-data/vietnam-administrative-divisions"

ROUND1_MAX_SCORE = 20
ROUND1_HIGH_THRESHOLD = 8  # chỉ dùng để phân tầng Vòng 1, KHÔNG khóa Vòng 2
FOLLOWUP_DAYS = 30
BIRTH_DATE_MIN = date(1900, 1, 1)
BIRTH_DATE_MAX = date.today()

ALTITUDE_OPTIONS = {
    "<500 m": 0.0,
    "500–999 m": 0.4,
    "1.000–1.499 m": 0.8,
    "1.500–1.999 m": 1.1,
    "2.000–2.499 m": 1.4,
    "2.500–2.999 m": 1.8,
    "3.000–3.499 m": 2.1,
    "3.500–3.999 m": 2.5,
    "4.000–4.499 m": 2.9,
    "4.500–4.999 m": 3.3,
}
ALTITUDE_LOOKUP_URL = "https://elevationfinder.net/"

PBKDF2_ITERATIONS = 310_000
DB_SCHEMA_READY = "db_schema_ready_v23"

# Hằng số điều hướng (dùng chung để không sai chuỗi so sánh)
PAGE_SCREENING = ":material/edit_note: Nhập sàng lọc"
PAGE_ADMIN = ":material/admin_panel_settings: Quản trị hệ thống"

# Bộ icon thống nhất cho mức nguy cơ
RISK_STYLE = {
    "THẤP": ("success", ":material/check_circle:", "risk-low"),
    "TRUNG BÌNH": ("warning", ":material/warning:", "risk-mid"),
    "CAO": ("error", ":material/error:", "risk-high"),
}


def risk_alert(title, category, body):
    kind, icon, _ = RISK_STYLE.get(category, RISK_STYLE["CAO"])
    getattr(st, kind)(f"**{title}: {category}**\n\n{body}", icon=icon)


def risk_badge(category):
    cls = RISK_STYLE.get(category, RISK_STYLE["CAO"])[2]
    return f'<span class="risk-badge {cls}">{category}</span>'


def bullet_list(items):
    for item in items:
        st.markdown(f"- {item}")


# ------------------------------------------------------------
# DATABASE
# ------------------------------------------------------------

try:
    import psycopg
    DB_INTEGRITY_ERRORS = (psycopg.IntegrityError,)
except ModuleNotFoundError:
    psycopg = None
    DB_INTEGRITY_ERRORS = ()


class PersistentPGConnection:
    """Lớp bọc psycopg: dùng host/port/user/password riêng hoặc DATABASE_URL."""

    def __init__(self, dsn=None):
        if psycopg is None:
            raise RuntimeError(
                "Thiếu thư viện psycopg. Hãy thêm `psycopg[binary]` vào requirements.txt."
            )
        if DB_HOST and DB_USER and DB_PASSWORD:
            kwargs = {
                "host": DB_HOST, "port": DB_PORT, "dbname": DB_NAME or "postgres",
                "user": DB_USER, "password": DB_PASSWORD,
            }
        elif dsn:
            kwargs = {"conninfo": dsn}
        else:
            raise RuntimeError(
                "Chưa cấu hình database. Cần DB_HOST, DB_PORT, DB_NAME, DB_USER, DB_PASSWORD "
                "trong Streamlit Secrets, hoặc một DATABASE_URL hợp lệ."
            )
        try:
            self._conn = psycopg.connect(**kwargs)
        except Exception as exc:
            raise RuntimeError(
                "Không thể kết nối PostgreSQL. Kiểm tra lại cấu hình database trong Secrets."
            ) from exc
        self._conn.autocommit = False

    @staticmethod
    def _translate(sql):
        return sql.replace("?", "%s")

    def execute(self, sql, params=None):
        cur = self._conn.cursor()
        cur.execute(self._translate(sql), params or ())
        return cur

    def commit(self):
        self._conn.commit()

    def rollback(self):
        self._conn.rollback()

    def close(self):
        try:
            self._conn.close()
        except Exception:
            pass


def _pg_table_columns(conn, table_name):
    rows = conn.execute(
        "SELECT column_name FROM information_schema.columns "
        "WHERE table_schema = 'public' AND table_name = ?",
        (table_name,),
    ).fetchall()
    return {row[0] for row in rows}


def _create_postgres_schema(conn):
    conn.execute("""
        CREATE TABLE IF NOT EXISTS patient_profiles (
            phone TEXT PRIMARY KEY,
            full_name TEXT NOT NULL,
            birth_date TEXT NOT NULL,
            gender TEXT NOT NULL,
            current_address TEXT NOT NULL,
            province TEXT NOT NULL,
            commune TEXT NOT NULL,
            research_consent INTEGER NOT NULL DEFAULT 0,
            consent_version TEXT,
            consent_at TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )""")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS user_accounts (
            id BIGSERIAL PRIMARY KEY,
            username TEXT NOT NULL UNIQUE,
            full_name TEXT NOT NULL,
            email TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            password_salt TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('admin', 'staff')),
            status TEXT NOT NULL CHECK(status IN ('pending', 'approved', 'rejected', 'disabled')),
            created_at TEXT NOT NULL,
            approved_by TEXT,
            approved_at TEXT,
            last_login_at TEXT
        )""")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS screening_records (
            id BIGSERIAL PRIMARY KEY,
            phone TEXT NOT NULL,
            screening_at TEXT NOT NULL,
            entered_by_username TEXT,
            entry_mode TEXT NOT NULL CHECK(entry_mode IN ('self', 'assisted')),
            consent_version TEXT NOT NULL,
            consent_at TEXT NOT NULL,
            round1_score INTEGER NOT NULL,
            round1_category TEXT NOT NULL,
            round1_conclusion TEXT,
            round1_reasons TEXT,
            answers_json TEXT,
            round2_completed INTEGER NOT NULL DEFAULT 0,
            altitude_choice TEXT,
            altitude_adjustment REAL,
            hb REAL,
            hb_adjusted REAL,
            mcv REAL,
            mch REAL,
            rbc REAL,
            rdw REAL,
            mentzer REAL,
            round2_score INTEGER,
            round2_category TEXT,
            round2_conclusion TEXT,
            round2_reasons TEXT,
            findings_json TEXT,
            advice_json TEXT,
            round3_completed INTEGER NOT NULL DEFAULT 0,
            round3_test_date TEXT,
            round3_facility TEXT,
            round3_test_type TEXT,
            round3_hba REAL,
            round3_hba2 REAL,
            round3_hbf REAL,
            round3_hbe REAL,
            round3_ferritin REAL,
            round3_serum_iron REAL,
            round3_transferrin_saturation REAL,
            round3_genetic_result TEXT,
            round3_lab_conclusion TEXT,
            round3_followup_status TEXT,
            round3_counseling_note TEXT,
            round3_followup_date TEXT,
            FOREIGN KEY(phone) REFERENCES patient_profiles(phone)
        )""")
    conn.execute("CREATE TABLE IF NOT EXISTS system_meta (key TEXT PRIMARY KEY, value TEXT)")

    # Migration an toàn cho các cột thêm sau bản prototype đầu.
    patient_columns = _pg_table_columns(conn, "patient_profiles")
    for column, definition in [
        ("research_consent", "INTEGER NOT NULL DEFAULT 0"),
        ("consent_version", "TEXT"),
        ("consent_at", "TEXT"),
    ]:
        if column not in patient_columns:
            conn.execute(f"ALTER TABLE patient_profiles ADD COLUMN {column} {definition}")

    record_columns = _pg_table_columns(conn, "screening_records")
    for column, definition in [
        ("round3_completed", "INTEGER NOT NULL DEFAULT 0"),
        ("round3_test_date", "TEXT"), ("round3_facility", "TEXT"), ("round3_test_type", "TEXT"),
        ("round3_hba", "REAL"), ("round3_hba2", "REAL"), ("round3_hbf", "REAL"), ("round3_hbe", "REAL"),
        ("round3_ferritin", "REAL"), ("round3_serum_iron", "REAL"),
        ("round3_transferrin_saturation", "REAL"), ("round3_genetic_result", "TEXT"),
        ("round3_lab_conclusion", "TEXT"), ("round3_followup_status", "TEXT"),
        ("round3_counseling_note", "TEXT"), ("round3_followup_date", "TEXT"),
    ]:
        if column not in record_columns:
            conn.execute(f"ALTER TABLE screening_records ADD COLUMN {column} {definition}")


def _copy_rows(conn, local, table, conflict_clause, defaults=None):
    """Chép dữ liệu bảng từ SQLite cũ sang PostgreSQL theo các cột chung."""
    pg_cols = _pg_table_columns(conn, table)
    rows = local.execute(f"SELECT * FROM {table}").fetchall()
    for r in rows:
        data = {k: r[k] for k in r.keys() if k in pg_cols}
        for k, v in (defaults or {}).items():
            if data.get(k) is None:
                data[k] = v
        cols = list(data.keys())
        sql = (
            f"INSERT INTO {table} ({', '.join(cols)}) "
            f"VALUES ({', '.join('?' for _ in cols)}) {conflict_clause}"
        )
        conn.execute(sql, tuple(data[c] for c in cols))


def _migrate_local_sqlite_once(conn):
    """Chuyển DB SQLite prototype cũ (nếu có) sang database bền vững, đúng một lần."""
    if not os.path.exists(LOCAL_SQLITE_PATH):
        return
    marker = conn.execute(
        "SELECT value FROM system_meta WHERE key = ?", ("local_sqlite_migrated",)
    ).fetchone()
    if marker and str(marker[0]) == "1":
        return
    try:
        import sqlite3
        local = sqlite3.connect(LOCAL_SQLITE_PATH)
        local.row_factory = sqlite3.Row
        now = datetime.now().isoformat(timespec="seconds")

        _copy_rows(conn, local, "patient_profiles", "ON CONFLICT (phone) DO NOTHING",
                   {"research_consent": 0, "created_at": now, "updated_at": now})
        try:
            _copy_rows(conn, local, "user_accounts", "ON CONFLICT (username) DO NOTHING")
        except Exception:
            pass
        _copy_rows(conn, local, "screening_records", "ON CONFLICT (id) DO NOTHING", {
            "entry_mode": "self", "consent_version": CONSENT_VERSION, "consent_at": now,
            "round1_score": 0, "round1_category": "THẤP",
            "round2_completed": 0, "round3_completed": 0,
        })
        local.close()
        conn.execute(
            "INSERT INTO system_meta(key,value) VALUES(?,?) "
            "ON CONFLICT(key) DO UPDATE SET value=EXCLUDED.value",
            ("local_sqlite_migrated", "1"),
        )
        conn.commit()
    except Exception as exc:
        try:
            conn.rollback()
        except Exception:
            pass
        st.warning(
            "Đã kết nối database bền vững nhưng không thể tự động chuyển dữ liệu SQLite cũ: "
            f"{exc}"
        )


def _sync_postgres_sequences(conn):
    """Đồng bộ sequence sau khi migrate các ID cũ."""
    for table in ("user_accounts", "screening_records"):
        conn.execute(
            f"SELECT setval(pg_get_serial_sequence('{table}', 'id'), "
            f"COALESCE((SELECT MAX(id) FROM {table}), 1), TRUE)"
        )


def _ensure_schema_once():
    if st.session_state.get(DB_SCHEMA_READY):
        return
    conn = PersistentPGConnection(DATABASE_URL or None)
    try:
        _create_postgres_schema(conn)
        _migrate_local_sqlite_once(conn)
        _sync_postgres_sequences(conn)
        conn.commit()
        st.session_state[DB_SCHEMA_READY] = True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _database_is_configured():
    return bool((DB_HOST and DB_USER and DB_PASSWORD) or DATABASE_URL)


def get_db():
    if not _database_is_configured():
        raise RuntimeError(
            "Chưa cấu hình database bền vững. Hãy thêm thông tin kết nối vào Streamlit Secrets."
        )
    _ensure_schema_once()
    return PersistentPGConnection(DATABASE_URL or None)


# ------------------------------------------------------------
# PASSWORD / VALIDATION
# ------------------------------------------------------------

def hash_password(password, salt_hex=None):
    if salt_hex:
        salt, salt_value = bytes.fromhex(salt_hex), salt_hex
    else:
        salt = secrets.token_bytes(16)
        salt_value = salt.hex()
    digest = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), salt, PBKDF2_ITERATIONS
    ).hex()
    return digest, salt_value


def verify_password(password, stored_hash, stored_salt):
    digest, _ = hash_password(password, stored_salt)
    return secrets.compare_digest(digest, stored_hash)


def normalize_phone(value):
    """Chuẩn hóa số điện thoại Việt Nam về chuỗi 10 chữ số bắt đầu bằng 0."""
    digits = re.sub(r"\D", "", str(value or ""))
    if digits.startswith("84") and len(digits) == 11:
        digits = "0" + digits[2:]
    return digits


def valid_vietnam_phone(phone):
    return bool(re.fullmatch(r"0\d{9}", str(phone or "")))


def valid_email(email):
    return bool(re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", (email or "").strip()))


def valid_username(username):
    return bool(re.fullmatch(r"[A-Za-z0-9_.-]{4,32}", (username or "").strip()))


# ------------------------------------------------------------
# ACCOUNTS
# ------------------------------------------------------------

def ensure_default_admin():
    """Tạo/sửa tài khoản quản trị đầu tiên từ secrets (ADMIN_USERNAME / ADMIN_PASSWORD).

    Không có secrets thì không tạo gì, để tránh mật khẩu mặc định nằm trong code.
    """
    if not (ADMIN_USERNAME and ADMIN_PASSWORD):
        return
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT id, role FROM user_accounts WHERE username = ? LIMIT 1",
            (ADMIN_USERNAME,),
        ).fetchone()
        now = datetime.now().isoformat(timespec="seconds")

        if row:
            if str(row[1] or "").lower() != "admin":
                password_hash, password_salt = hash_password(ADMIN_PASSWORD)
                conn.execute(
                    "UPDATE user_accounts SET full_name = ?, password_hash = ?, password_salt = ?, "
                    "role = 'admin', status = 'approved', approved_by = ?, approved_at = ? WHERE id = ?",
                    (ADMIN_FULL_NAME, password_hash, password_salt, ADMIN_USERNAME, now, row[0]),
                )
                conn.commit()
            return

        if conn.execute("SELECT id FROM user_accounts WHERE role = 'admin' LIMIT 1").fetchone():
            return

        password_hash, password_salt = hash_password(ADMIN_PASSWORD)
        email_taken = conn.execute(
            "SELECT id FROM user_accounts WHERE lower(email) = lower(?) LIMIT 1",
            (ADMIN_EMAIL,),
        ).fetchone()
        admin_email = f"{ADMIN_USERNAME}@thalassemia.local" if email_taken else ADMIN_EMAIL

        conn.execute(
            "INSERT INTO user_accounts (username, full_name, email, password_hash, password_salt, "
            "role, status, created_at, approved_by, approved_at) "
            "VALUES (?, ?, ?, ?, ?, 'admin', 'approved', ?, ?, ?) ON CONFLICT (username) DO NOTHING",
            (ADMIN_USERNAME, ADMIN_FULL_NAME, admin_email, password_hash, password_salt,
             now, ADMIN_USERNAME, now),
        )
        conn.commit()
    except DB_INTEGRITY_ERRORS:
        # Phiên khác có thể vừa tạo admin cùng lúc.
        conn.rollback()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def admin_exists():
    conn = get_db()
    try:
        row = conn.execute("SELECT 1 FROM user_accounts WHERE role = 'admin' LIMIT 1").fetchone()
    finally:
        conn.close()
    return row is not None


def register_staff_account(username, full_name, email, password):
    username = username.strip()
    email = email.strip().lower()
    now = datetime.now().isoformat(timespec="seconds")
    password_hash, password_salt = hash_password(password)
    conn = get_db()
    try:
        conn.execute(
            "INSERT INTO user_accounts (username, full_name, email, password_hash, password_salt, "
            "role, status, created_at) VALUES (?, ?, ?, ?, ?, 'staff', 'pending', ?)",
            (username, full_name.strip(), email, password_hash, password_salt, now),
        )
        conn.commit()
        return True, ""
    except DB_INTEGRITY_ERRORS:
        conn.rollback()
        return False, "Tên đăng nhập hoặc email đã được sử dụng."
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def authenticate_user(login_value, password):
    login_value = (login_value or "").strip().lower()
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT id, username, full_name, email, password_hash, password_salt, role, status "
            "FROM user_accounts WHERE lower(username) = ? OR lower(email) = ? LIMIT 1",
            (login_value, login_value),
        ).fetchone()

        if not row or not verify_password(password, row[4], row[5]):
            return None, "Sai tên đăng nhập/email hoặc mật khẩu."

        if row[7] != "approved":
            if row[7] == "pending":
                return None, "Tài khoản đang chờ quản trị viên phê duyệt."
            if row[7] == "disabled":
                return None, "Tài khoản đã bị vô hiệu hóa."
            return None, "Tài khoản chưa được phép truy cập."

        conn.execute(
            "UPDATE user_accounts SET last_login_at = ? WHERE id = ?",
            (datetime.now().isoformat(timespec="seconds"), row[0]),
        )
        conn.commit()
    finally:
        conn.close()

    return {
        "id": row[0], "username": row[1], "full_name": row[2],
        "email": row[3], "role": row[6], "status": row[7],
    }, ""


def list_user_accounts():
    conn = get_db()
    try:
        return conn.execute(
            "SELECT id, username, full_name, email, role, status, created_at, "
            "approved_by, approved_at, last_login_at FROM user_accounts "
            "ORDER BY CASE role WHEN 'admin' THEN 0 ELSE 1 END, "
            "CASE status WHEN 'pending' THEN 0 ELSE 1 END, created_at DESC"
        ).fetchall()
    finally:
        conn.close()


def update_staff_status(user_id, status, approved_by):
    if status not in {"approved", "rejected", "disabled", "pending"}:
        raise ValueError("Trạng thái không hợp lệ.")
    conn = get_db()
    try:
        if status == "approved":
            conn.execute(
                "UPDATE user_accounts SET status = ?, approved_by = ?, approved_at = ? "
                "WHERE id = ? AND role = 'staff'",
                (status, approved_by, datetime.now().isoformat(timespec="seconds"), user_id),
            )
        else:
            conn.execute(
                "UPDATE user_accounts SET status = ? WHERE id = ? AND role = 'staff'",
                (status, user_id),
            )
        conn.commit()
    finally:
        conn.close()


# ------------------------------------------------------------
# PATIENT / SCREENING DATA
# ------------------------------------------------------------

_SCREENING_SELECT = """
    s.id, s.screening_at, s.entered_by_username, s.entry_mode,
    p.phone, p.full_name, p.birth_date, p.gender,
    p.current_address, p.province, p.commune,
    s.consent_version, s.consent_at,
    s.round1_score, s.round1_category,
    s.round2_completed, s.altitude_choice, s.altitude_adjustment,
    s.hb, s.hb_adjusted, s.mcv, s.mch, s.rbc, s.rdw, s.mentzer,
    s.round2_score, s.round2_category, s.round2_conclusion,
    s.round1_conclusion,
    s.round3_completed, s.round3_test_date, s.round3_facility, s.round3_test_type,
    s.round3_hba, s.round3_hba2, s.round3_hbf, s.round3_hbe,
    s.round3_ferritin, s.round3_serum_iron, s.round3_transferrin_saturation,
    s.round3_genetic_result, s.round3_lab_conclusion, s.round3_followup_status,
    s.round3_counseling_note, s.round3_followup_date
"""


def list_patient_profiles_for_staff():
    conn = get_db()
    try:
        return conn.execute(
            "SELECT phone, full_name, birth_date, gender, current_address, province, commune, "
            "research_consent, consent_version, consent_at, created_at, updated_at "
            "FROM patient_profiles WHERE research_consent = 1 ORDER BY updated_at DESC"
        ).fetchall()
    finally:
        conn.close()


def upsert_patient(profile):
    conn = get_db()
    now = datetime.now().isoformat(timespec="seconds")
    try:
        exists = conn.execute(
            "SELECT 1 FROM patient_profiles WHERE phone = ?", (profile["phone"],)
        ).fetchone()
        if exists:
            conn.execute(
                "UPDATE patient_profiles SET full_name = ?, birth_date = ?, gender = ?, "
                "current_address = ?, province = ?, commune = ?, research_consent = ?, "
                "consent_version = ?, consent_at = ?, updated_at = ? WHERE phone = ?",
                (profile["full_name"], profile["birth_date"], profile["gender"],
                 profile["current_address"], profile["province"], profile["commune"],
                 1, CONSENT_VERSION, profile["consent_at"], now, profile["phone"]),
            )
            action = "updated"
        else:
            conn.execute(
                "INSERT INTO patient_profiles (phone, full_name, birth_date, gender, current_address, "
                "province, commune, research_consent, consent_version, consent_at, created_at, updated_at) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (profile["phone"], profile["full_name"], profile["birth_date"], profile["gender"],
                 profile["current_address"], profile["province"], profile["commune"],
                 1, CONSENT_VERSION, profile["consent_at"], now, now),
            )
            action = "inserted"
        conn.commit()
        return action
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def create_screening_record(patient, answers, score1, category1, conclusion1, reasons1,
                            entry_mode, entered_by_username=None):
    """Lưu một lượt sàng lọc sau khi hoàn thành Vòng 1. Không xóa lượt cũ."""
    conn = get_db()
    try:
        cur = conn.execute(
            "INSERT INTO screening_records (phone, screening_at, entered_by_username, entry_mode, "
            "consent_version, consent_at, round1_score, round1_category, round1_conclusion, "
            "round1_reasons, answers_json) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) RETURNING id",
            (
                patient["phone"], datetime.now().isoformat(timespec="seconds"),
                entered_by_username, entry_mode, CONSENT_VERSION, patient["consent_at"],
                int(score1), category1, conclusion1,
                json.dumps(reasons1, ensure_ascii=False),
                json.dumps(answers, ensure_ascii=False),
            ),
        )
        record_id = int(cur.fetchone()[0])
        conn.commit()
        return record_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def update_screening_round2(record_id, r2):
    if not record_id:
        return
    conn = get_db()
    try:
        conn.execute(
            "UPDATE screening_records SET round2_completed = 1, altitude_choice = ?, "
            "altitude_adjustment = ?, hb = ?, hb_adjusted = ?, mcv = ?, mch = ?, rbc = ?, rdw = ?, "
            "mentzer = ?, round2_score = ?, round2_category = ?, round2_conclusion = ?, "
            "round2_reasons = ?, findings_json = ?, advice_json = ? WHERE id = ?",
            (
                r2["altitude_choice"], r2["adjustment"], r2["hb"], r2["hb_adjusted"],
                r2["mcv"], r2["mch"], r2["rbc"], r2["rdw"], r2["mentzer"],
                r2["score"], r2["category"], r2["conclusion"],
                json.dumps(r2["reasons"], ensure_ascii=False),
                json.dumps(r2["findings"], ensure_ascii=False),
                json.dumps(r2["advice"], ensure_ascii=False),
                int(record_id),
            ),
        )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def update_screening_round3(record_id, r3):
    """Lưu kết quả xét nghiệm chuyên sâu và thông tin theo dõi."""
    if not record_id:
        return
    conn = get_db()
    try:
        conn.execute(
            "UPDATE screening_records SET round3_completed = 1, round3_test_date = ?, "
            "round3_facility = ?, round3_test_type = ?, round3_hba = ?, round3_hba2 = ?, "
            "round3_hbf = ?, round3_hbe = ?, round3_ferritin = ?, round3_serum_iron = ?, "
            "round3_transferrin_saturation = ?, round3_genetic_result = ?, round3_lab_conclusion = ?, "
            "round3_followup_status = ?, round3_counseling_note = ?, round3_followup_date = ? "
            "WHERE id = ?",
            (
                r3.get("test_date"), r3.get("facility"), r3.get("test_type"),
                r3.get("hba"), r3.get("hba2"), r3.get("hbf"), r3.get("hbe"),
                r3.get("ferritin"), r3.get("serum_iron"), r3.get("transferrin_saturation"),
                r3.get("genetic_result"), r3.get("lab_conclusion"),
                r3.get("followup_status"), r3.get("counseling_note"), r3.get("followup_date"),
                int(record_id),
            ),
        )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def list_screening_records_for_staff():
    """TOÀN BỘ lịch sử sàng lọc của người tham gia đã đồng ý."""
    conn = get_db()
    try:
        return conn.execute(
            f"SELECT {_SCREENING_SELECT} FROM screening_records s "
            "JOIN patient_profiles p ON p.phone = s.phone "
            "WHERE p.research_consent = 1 ORDER BY s.screening_at DESC, s.id DESC"
        ).fetchall()
    finally:
        conn.close()


def list_latest_screening_records_for_staff():
    """Một lượt mới nhất cho mỗi người (hiển thị nhanh); không xóa lịch sử."""
    conn = get_db()
    try:
        return conn.execute(
            f"SELECT {_SCREENING_SELECT} FROM ("
            "  SELECT sr.*, ROW_NUMBER() OVER (PARTITION BY sr.phone "
            "         ORDER BY sr.screening_at DESC, sr.id DESC) AS rn FROM screening_records sr"
            ") s JOIN patient_profiles p ON p.phone = s.phone "
            "WHERE s.rn = 1 AND p.research_consent = 1 ORDER BY s.screening_at DESC, s.id DESC"
        ).fetchall()
    finally:
        conn.close()


def _blank(value):
    return "" if value is None else value


def export_screening_xlsx(patient_rows, screening_rows):
    """Tạo file Excel 2 sheet trong bộ nhớ, không ghi file tạm lên server."""
    try:
        import xlsxwriter
    except ModuleNotFoundError as exc:
        raise RuntimeError(
            "Thiếu thư viện XlsxWriter. Hãy thêm `XlsxWriter>=3.2` vào requirements.txt."
        ) from exc

    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    header_fmt = workbook.add_format({
        "bold": True, "bg_color": "#D9EAF7", "border": 1,
        "align": "center", "valign": "vcenter", "text_wrap": True,
    })
    cell_fmt = workbook.add_format({"border": 1, "valign": "top"})

    # Sheet 1: hồ sơ hiện tại
    ws = workbook.add_worksheet("Ho_so_nguoi_tham_gia")
    patient_headers = [
        "Số điện thoại", "Họ và tên", "Ngày sinh", "Giới tính", "Địa chỉ hiện tại",
        "Tỉnh/thành", "Phường/xã/đặc khu", "Đồng ý nghiên cứu", "Phiên bản consent",
        "Thời điểm đồng ý", "Tạo lúc", "Cập nhật lúc",
    ]
    for c, h in enumerate(patient_headers):
        ws.write(0, c, h, header_fmt)
    for r, row in enumerate(patient_rows, start=1):
        values = [row[0], row[1], row[2], row[3], row[4], row[5], row[6],
                  "Có" if row[7] else "Không", row[8] or "", row[9] or "", row[10], row[11]]
        for c, value in enumerate(values):
            ws.write(r, c, value, cell_fmt)
    ws.freeze_panes(1, 0)
    for c, w in enumerate([15, 24, 13, 12, 32, 20, 26, 16, 25, 21, 21, 21]):
        ws.set_column(c, c, w)
    ws.autofilter(0, 0, max(len(patient_rows), 1), len(patient_headers) - 1)

    # Sheet 2: lịch sử sàng lọc
    ws2 = workbook.add_worksheet("Lich_su_sang_loc")
    screening_headers = [
        "ID lượt sàng lọc", "Thời điểm", "Người nhập", "Hình thức nhập",
        "Số điện thoại", "Họ và tên", "Ngày sinh", "Giới tính",
        "Địa chỉ", "Tỉnh/thành", "Phường/xã/đặc khu", "Phiên bản consent",
        "Thời điểm đồng ý", "Điểm Vòng 1", "Nguy cơ Vòng 1",
        "Đã hoàn thành Vòng 2", "Khoảng độ cao", "Hiệu chỉnh Hb (g/dL)",
        "Hb (g/dL)", "Hb sau hiệu chỉnh (g/dL)", "MCV (fL)", "MCH (pg)",
        "RBC (T/L)", "RDW-CV (%)", "Mentzer Index", "Điểm CBC",
        "Nguy cơ Vòng 2", "Kết luận Vòng 2", "Kết luận Vòng 1",
        "Vòng 3", "Ngày xét nghiệm chuyên sâu", "Cơ sở thực hiện", "Loại xét nghiệm",
        "HbA (%)", "HbA2 (%)", "HbF (%)", "HbE (%)", "Ferritin", "Sắt huyết thanh",
        "Độ bão hòa transferrin (%)", "Kết quả di truyền", "Kết luận trên phiếu xét nghiệm",
        "Trạng thái theo dõi", "Ghi chú tư vấn", "Ngày hẹn theo dõi",
    ]
    for c, h in enumerate(screening_headers):
        ws2.write(0, c, h, header_fmt)
    for r, row in enumerate(screening_rows, start=1):
        values = [_blank(v) for v in row]
        values[3] = "Tự nhập" if row[3] == "self" else "Nhập giúp người tham gia"
        values[15] = "Có" if row[15] else "Chưa"
        values[29] = "Có" if row[29] else "Chưa"
        for c, value in enumerate(values):
            ws2.write(r, c, value, cell_fmt)
    ws2.freeze_panes(1, 0)
    for c in range(len(screening_headers)):
        ws2.set_column(c, c, 20)
    ws2.set_column(4, 4, 15)
    ws2.set_column(5, 5, 24)
    ws2.set_column(8, 8, 32)
    ws2.set_column(9, 10, 22)
    ws2.set_column(27, 28, 38)
    ws2.autofilter(0, 0, max(len(screening_rows), 1), len(screening_headers) - 1)

    workbook.close()
    output.seek(0)
    return output


# ------------------------------------------------------------
# DATABASE GATE
# ------------------------------------------------------------

if not _database_is_configured():
    st.error("Database chưa được cấu hình — hệ thống đang khóa ghi dữ liệu để tránh mất dữ liệu.")
    with st.expander("Cấu hình database bền vững", expanded=True):
        st.code(
            'DB_HOST = "<host của Supabase pooler>"\n'
            'DB_PORT = 6543\n'
            'DB_NAME = "postgres"\n'
            'DB_USER = "<user của Supabase>"\n'
            'DB_PASSWORD = "<mật khẩu database>"\n\n'
            'ADMIN_USERNAME = "<tên đăng nhập quản trị>"\n'
            'ADMIN_PASSWORD = "<mật khẩu quản trị mạnh>"',
            language="toml",
        )
        st.caption(
            "Đặt các giá trị này trong Streamlit Cloud → Manage app → Settings → Secrets. "
            "Không cần percent-encode mật khẩu khi dùng DB_PASSWORD riêng."
        )
    st.stop()

try:
    if not st.session_state.get("default_admin_ready_v23"):
        ensure_default_admin()
        st.session_state["default_admin_ready_v23"] = True
except Exception as exc:
    st.error("Không thể kết nối/khởi tạo database bền vững.")
    st.code(str(exc))
    st.info("Kiểm tra lại cấu hình database trong Streamlit Secrets.")
    st.stop()


# ------------------------------------------------------------
# AUTH UI
# ------------------------------------------------------------

def current_auth_user():
    return st.session_state.get("auth_user")


def logout_user():
    st.session_state.pop("auth_user", None)


def render_auth_sidebar():
    user = current_auth_user()
    st.sidebar.header(":material/lock: Tài khoản và quyền truy cập")

    if user:
        role_label = "Quản trị viên" if user["role"] == "admin" else "Nhân sự được duyệt"
        st.sidebar.success(f"Đang đăng nhập: **{user['full_name']}**\n\n{role_label}",
                           icon=":material/verified_user:")

        if user["role"] == "admin":
            page = st.sidebar.radio("Khu vực làm việc", [PAGE_SCREENING, PAGE_ADMIN],
                                    key="auth_page_admin")
        else:
            page = PAGE_SCREENING
            st.sidebar.info(
                "Bạn có thể nhập hồ sơ/sàng lọc giúp người tham gia. "
                "Danh sách hồ sơ chỉ xem được ở khu vực quản trị có quyền.",
                icon=":material/info:",
            )
        st.session_state["auth_page"] = page

        if st.sidebar.button("Đăng xuất", icon=":material/logout:", use_container_width=True):
            logout_user()
            st.session_state.pop("auth_page", None)
            st.rerun()
        return

    st.sidebar.caption(
        "Người tham gia có thể tự nhập hồ sơ.\n"
        "Quản trị viên/nhân sự được duyệt đăng nhập để quản lý hoặc nhập giúp người tham gia."
    )
    auth_mode = st.sidebar.radio(
        "Chế độ", ["Đăng nhập", "Đăng ký nhân sự"],
        key="auth_mode_v9", label_visibility="collapsed",
    )

    if auth_mode == "Đăng nhập":
        with st.sidebar.form("login_form"):
            login_value = st.text_input("Tên đăng nhập hoặc email")
            password = st.text_input("Mật khẩu", type="password")
            login_submit = st.form_submit_button("Đăng nhập", icon=":material/login:",
                                                 use_container_width=True)
        if login_submit:
            user_result, message = authenticate_user(login_value, password)
            if user_result:
                st.session_state["auth_user"] = user_result
                st.session_state["auth_page"] = PAGE_SCREENING
                st.rerun()
            else:
                st.sidebar.error(message, icon=":material/error:")

        try:
            if not admin_exists():
                st.sidebar.warning(
                    "Chưa có tài khoản quản trị. Hãy đặt ADMIN_USERNAME và ADMIN_PASSWORD "
                    "trong Streamlit Secrets rồi tải lại ứng dụng.",
                    icon=":material/warning:",
                )
        except Exception:
            pass
    else:
        with st.sidebar.form("staff_register_form"):
            staff_username = st.text_input("Tên đăng nhập")
            staff_full_name = st.text_input("Họ và tên")
            staff_email = st.text_input("Email")
            staff_password = st.text_input("Mật khẩu", type="password")
            staff_password2 = st.text_input("Nhập lại mật khẩu", type="password")
            register_submit = st.form_submit_button(
                "Gửi yêu cầu tạo tài khoản", icon=":material/person_add:", use_container_width=True
            )

        if register_submit:
            if not valid_username(staff_username):
                st.sidebar.error("Tên đăng nhập 4–32 ký tự, không có khoảng trắng.")
            elif not staff_full_name.strip():
                st.sidebar.error("Vui lòng nhập họ và tên.")
            elif not valid_email(staff_email):
                st.sidebar.error("Email chưa đúng định dạng.")
            elif len(staff_password) < 8:
                st.sidebar.error("Mật khẩu phải có ít nhất 8 ký tự.")
            elif staff_password != staff_password2:
                st.sidebar.error("Hai mật khẩu không khớp.")
            else:
                ok, msg = register_staff_account(
                    staff_username, staff_full_name, staff_email, staff_password
                )
                if ok:
                    st.sidebar.success(
                        "Đã gửi tài khoản. Quản trị viên phải phê duyệt trước khi đăng nhập.",
                        icon=":material/check_circle:",
                    )
                else:
                    st.sidebar.error(msg)

        st.sidebar.caption(
            "Tài khoản nhân sự không được xem dữ liệu người tham gia cho đến khi quản trị viên phê duyệt."
        )


def render_admin_console(user):
    st.header(":material/admin_panel_settings: Quản trị hệ thống")
    st.success(
        f"Xin chào **{user['full_name']}** — quyền: "
        f"{'Quản trị viên' if user['role'] == 'admin' else 'Nhân sự được duyệt'}",
        icon=":material/verified_user:",
    )

    patients = list_patient_profiles_for_staff()
    users = list_user_accounts()
    screening_rows = list_latest_screening_records_for_staff()
    screening_history_rows = list_screening_records_for_staff()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Hồ sơ đã đồng ý", len(patients))
    c2.metric("Chờ phê duyệt", sum(1 for r in users if r[5] == "pending" and r[4] == "staff"))
    c3.metric("Nhân sự được duyệt", sum(1 for r in users if r[4] == "staff" and r[5] == "approved"))
    c4.metric("Người có lượt sàng lọc", len(screening_rows))

    if user["role"] == "admin":
        st.subheader(":material/group: Phê duyệt tài khoản nhân sự")
        pending_users = [r for r in users if r[4] == "staff" and r[5] == "pending"]
        if not pending_users:
            st.info("Hiện không có tài khoản nào đang chờ phê duyệt.", icon=":material/inbox:")
        else:
            for row in pending_users:
                with st.container(border=True):
                    a, b, c = st.columns([2, 2, 1])
                    with a:
                        st.write(f"**{row[2]}**")
                        st.caption(f"Username: {row[1]} · Email: {row[3]}")
                    with b:
                        st.caption(f"Đăng ký: {row[6]}")
                    with c:
                        b1, b2 = st.columns(2)
                        if b1.button("", key=f"approve_{row[0]}", icon=":material/check:",
                                     help="Duyệt tài khoản"):
                            update_staff_status(row[0], "approved", user["username"])
                            st.rerun()
                        if b2.button("", key=f"reject_{row[0]}", icon=":material/close:",
                                     help="Từ chối"):
                            update_staff_status(row[0], "rejected", user["username"])
                            st.rerun()

        st.subheader(":material/manage_accounts: Tài khoản nhân sự")
        for row in [r for r in users if r[4] == "staff"]:
            with st.container(border=True):
                s1, s2, s3 = st.columns([2, 2, 1])
                with s1:
                    st.write(f"**{row[2]}**")
                    st.caption(f"{row[1]} · {row[3]}")
                with s2:
                    st.write(f"Trạng thái: **{row[5]}**")
                with s3:
                    if row[5] == "approved":
                        if st.button("Vô hiệu hóa", key=f"disable_{row[0]}", icon=":material/block:"):
                            update_staff_status(row[0], "disabled", user["username"])
                            st.rerun()
                    elif row[5] in {"disabled", "rejected"}:
                        if st.button("Mở lại", key=f"enable_{row[0]}", icon=":material/lock_open:"):
                            update_staff_status(row[0], "approved", user["username"])
                            st.rerun()

    st.divider()
    st.subheader(":material/table_view: Dữ liệu người tham gia")
    st.caption(
        "Chỉ quản trị viên và nhân sự đã được phê duyệt mới xem được dữ liệu này. "
        "Bảng trên web hiển thị lượt mới nhất cho mỗi số điện thoại; toàn bộ lịch sử vẫn được giữ "
        "trong database và file Excel nghiên cứu."
    )

    tab1, tab2 = st.tabs([":material/person: Hồ sơ hiện tại",
                          ":material/science: Lượt sàng lọc gần nhất"])

    with tab1:
        patient_table = [{
            "Số điện thoại": r[0], "Họ tên": r[1], "Ngày sinh": r[2], "Giới tính": r[3],
            "Địa chỉ hiện tại": r[4], "Tỉnh/thành": r[5], "Phường/xã/đặc khu": r[6],
            "Đồng ý nghiên cứu": "Có" if r[7] else "Không", "Phiên bản consent": r[8] or "",
            "Thời điểm đồng ý": r[9] or "", "Tạo lúc": r[10], "Cập nhật lúc": r[11],
        } for r in patients]
        if patient_table:
            st.dataframe(patient_table, use_container_width=True, hide_index=True)
        else:
            st.info("Chưa có hồ sơ nào đã đồng ý tham gia.")

    with tab2:
        screening_table = [{
            "ID": r[0], "Thời điểm": r[1], "Người nhập": r[2] or "",
            "Hình thức nhập": "Tự nhập" if r[3] == "self" else "Nhập giúp người tham gia",
            "Số điện thoại": r[4], "Họ tên": r[5], "Tỉnh/thành": r[9], "Phường/xã/đặc khu": r[10],
            "Điểm Vòng 1": r[13], "Nguy cơ Vòng 1": r[14],
            "Vòng 2": "Có" if r[15] else "Chưa", "Độ cao": r[16] or "",
            "Hb": _blank(r[18]), "Hb sau hiệu chỉnh": _blank(r[19]),
            "MCV": _blank(r[20]), "MCH": _blank(r[21]), "RBC": _blank(r[22]),
            "RDW": _blank(r[23]), "Mentzer": _blank(r[24]), "Điểm CBC": _blank(r[25]),
            "Nguy cơ Vòng 2": r[26] or "", "Kết luận": r[27] or r[28] or "",
            "Vòng 3": "Có" if r[29] else "Chưa", "Ngày xét nghiệm": r[30] or "",
            "Cơ sở thực hiện": r[31] or "", "Loại xét nghiệm": r[32] or "",
            "HbA2": _blank(r[34]), "HbF": _blank(r[35]), "HbE": _blank(r[36]),
            "Ferritin": _blank(r[37]), "Kết quả di truyền": r[40] or "",
            "Kết luận xét nghiệm": r[41] or "", "Trạng thái theo dõi": r[42] or "",
            "Ghi chú tư vấn": r[43] or "", "Ngày hẹn theo dõi": r[44] or "",
        } for r in screening_rows]
        if screening_table:
            st.dataframe(screening_table, use_container_width=True, hide_index=True)
        else:
            st.info("Chưa có lượt sàng lọc nào được lưu.")

    if patients or screening_history_rows:
        try:
            excel_data = export_screening_xlsx(patients, screening_history_rows)
            st.download_button(
                "Xuất dữ liệu Excel — toàn bộ lịch sử (.xlsx)",
                data=excel_data.getvalue(),
                file_name=f"Thalassemia_du_lieu_{date.today().isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
                use_container_width=True,
            )
            st.caption(
                "File Excel gồm 2 sheet: Hồ sơ hiện tại và toàn bộ lịch sử sàng lọc. "
                "Không xuất mật khẩu/tài khoản nhân sự."
            )
        except RuntimeError as exc:
            st.error(str(exc))


# ------------------------------------------------------------
# ADMINISTRATIVE DATA (34 PROVINCES / 3,321 COMMUNES)
# ------------------------------------------------------------

@st.cache_data(ttl=86400, show_spinner=False)
def load_admin_hierarchy():
    """Tải danh mục địa giới 2 cấp hiện hành: 34 tỉnh/thành, 3.321 đơn vị cấp xã."""
    response = requests.get(ADMIN_DATA_URL, timeout=20)
    response.raise_for_status()
    payload = response.json()

    records = payload.get("data", payload)
    if not isinstance(records, list) or not records:
        raise ValueError("Dữ liệu địa giới không đúng định dạng.")

    provinces = {}
    for province in records:
        province_name = province.get("name", {}).get("local")
        if not province_name:
            continue
        wards = []
        for ward in province.get("ward", []) or province.get("wards", []):
            ward_name = ward.get("name", {}).get("local")
            if ward_name:
                wards.append(ward_name)
        provinces[province_name] = sorted(set(wards), key=str.casefold)

    if len(provinces) != 34:
        raise ValueError(f"Dữ liệu địa giới hiện trả về {len(provinces)} tỉnh/thành, không phải 34.")
    total = sum(len(items) for items in provinces.values())
    if total != 3321:
        raise ValueError(f"Dữ liệu địa giới hiện có {total} đơn vị cấp xã, không phải 3.321.")
    return provinces


# ------------------------------------------------------------
# GENERAL HELPERS
# ------------------------------------------------------------

def calculate_age(birth_date):
    today = date.today()
    return today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))


def safe_filename(text):
    text = re.sub(r"[^0-9A-Za-zÀ-ỹĐđ _-]", "_", (text or "").strip())
    return text.strip(" _") or "nguoi_sang_loc"


def reset_results():
    st.session_state.pop("screening_id", None)
    for key in list(st.session_state.keys()):
        if key.startswith(("round1_", "round2_", "round3_", "low_cbc_", "google_")):
            del st.session_state[key]


# ------------------------------------------------------------
# ALTITUDE
# ------------------------------------------------------------

def altitude_selector(key_prefix):
    choice = st.radio(
        "Chọn khoảng độ cao nơi đang sinh sống",
        list(ALTITUDE_OPTIONS.keys()),
        index=0,
        key=f"{key_prefix}_choice",
    )
    st.link_button("Tra cứu độ cao nơi ở", ALTITUDE_LOOKUP_URL, icon=":material/terrain:")

    adjustment = ALTITUDE_OPTIONS.get(choice, 0.0)
    st.info(f"Hiệu chỉnh Hb tham khảo: **-{adjustment:.1f} g/dL**", icon=":material/tune:")

    if choice.startswith(("2.500", "3.000", "3.500", "4.000", "4.500")):
        st.warning(
            "Ở độ cao ≥2.500 m, WHO lưu ý mức độ không chắc chắn của hiệu chỉnh cao hơn.",
            icon=":material/warning:",
        )
    return choice, adjustment


# ------------------------------------------------------------
# CBC UNIT CONVERSION
# ------------------------------------------------------------

def _decimal(value):
    try:
        return Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError):
        raise ValueError("Giá trị số không hợp lệ.")


def _round_number(value, places=4):
    q = Decimal("1").scaleb(-places)
    return float(_decimal(value).quantize(q, rounding=ROUND_HALF_UP))


def hb_to_g_dl(value, unit):
    value = _decimal(value)
    if unit == "g/dL":
        return _round_number(value, 4)
    if unit == "g/L":
        return _round_number(value / Decimal("10"), 4)
    raise ValueError("Đơn vị Hb không hợp lệ.")


def rbc_to_t_l(value, unit):
    # T/L, 10^12/L và 10^6/µL có cùng trị số quy đổi.
    if unit in ("T/L", "10^12/L", "10^6/µL"):
        return _round_number(value, 4)
    raise ValueError("Đơn vị RBC không hợp lệ.")


# ------------------------------------------------------------
# ROUND 1
# ------------------------------------------------------------

def calculate_round1_score(a):
    score = 0
    reasons = []
    weighted_items = [
        ("q1", 3, "Có người thân/dòng họ mắc Thalassemia"),
        ("q2", 3, "Có người thân/dòng họ mang gen Thalassemia/hemoglobinopathy"),
        ("q3", 1, "Cha/mẹ từng xét nghiệm Thalassemia/hemoglobinopathy"),
        ("q4", 2, "Anh/chị/em từng thiếu máu hoặc hồng cầu nhỏ"),
        ("q5", 2, "Gia đình có trẻ từng truyền máu nhiều lần/định kỳ"),
        ("q6", 1, "Từng được thông báo thiếu máu"),
        ("q7", 2, "Từng được thông báo MCV thấp/hồng cầu nhỏ"),
        ("q8", 1, "Từng được thông báo MCH thấp/hồng cầu nhược sắc"),
        ("q10", 2, "Từng được chẩn đoán HbE/hemoglobinopathy khác"),
        ("q11", 1, "Bản thân từng truyền máu nhiều lần/định kỳ"),
        ("q12", 2, "Thiếu máu kéo dài từ nhỏ/tuổi thiếu niên"),
        ("q13", 1, "Mệt mỏi/giảm sức hoạt động"),
        ("q14", 1, "Hoa mắt/chóng mặt không rõ nguyên nhân"),
        ("q15", 1, "Da/niêm nhợt"),
        ("q16", 1, "Vàng da/vàng mắt không rõ nguyên nhân"),
        ("q17", 2, "Từng được ghi nhận lách to/gan lách to"),
        ("q18", 1, "Có tiền sử/biến chứng bệnh huyết học mạn"),
    ]
    for key, weight, label in weighted_items:
        if a.get(key) == "Có":
            score += weight
            reasons.append(label)

    if a.get("q9") == "Đã nghi ngờ":
        score += 2
        reasons.append("Từng có kết quả nghi ngờ Thalassemia/hemoglobinopathy")
    elif a.get("q9") == "Đã xác định mang gen":
        score += 4
        reasons.append("Từng được xác định mang gen")

    return min(score, ROUND1_MAX_SCORE), reasons


def round1_category(score):
    """Phân tầng thông tin Vòng 1; không dùng để loại người tham gia khỏi Vòng 2."""
    if score >= ROUND1_HIGH_THRESHOLD:
        return ("CAO", "Có nhiều yếu tố đáng lưu ý qua sàng lọc ban đầu; nên ưu tiên thực hiện/nhập CBC và được tư vấn y tế phù hợp.")
    if score >= 4:
        return ("TRUNG BÌNH", "Có một số yếu tố đáng lưu ý; nên cân nhắc CBC nếu chưa có kết quả và tiếp tục theo dõi.")
    return ("THẤP", "Chưa ghi nhận nhiều yếu tố đáng lưu ý qua bộ câu hỏi ban đầu; kết quả này không loại trừ khả năng mang gen Thalassemia.")


def round1_guidance(category):
    common = [
        "Vòng 1 chỉ là sàng lọc ban đầu, không dùng để chẩn đoán hoặc loại trừ Thalassemia.",
        "Bạn vẫn có thể tiếp tục Vòng 2 nếu đã có CBC hoặc muốn nhập CBC sau khi xét nghiệm.",
        "Nếu chưa có CBC, có thể trao đổi với cơ sở y tế về việc thực hiện công thức máu khi phù hợp với tình trạng sức khỏe.",
    ]
    if category == "CAO":
        action = [
            "Ưu tiên có/nhập CBC gồm Hb, RBC, MCV, MCH và RDW.",
            "Nếu CBC có kiểu hình hồng cầu nhỏ/nhược sắc, nên được đánh giá nguyên nhân, đặc biệt phân biệt thiếu sắt với Thalassemia/hemoglobinopathy.",
            "Khi có chỉ định chuyên môn, có thể cần ferritin/đánh giá sắt và xét nghiệm hemoglobin như HPLC hoặc điện di Hb; xét nghiệm gen tùy trường hợp.",
        ]
    elif category == "TRUNG BÌNH":
        action = [
            "Nên lưu ý các yếu tố đã ghi nhận và cân nhắc CBC nếu chưa có kết quả.",
            "Nếu đã có CBC, hãy nhập kết quả vào Vòng 2 để hệ thống phân tích các chỉ số hồng cầu.",
            "Nếu bất thường kéo dài hoặc có tiền sử gia đình đáng chú ý, nên trao đổi với nhân viên y tế về nhu cầu xét nghiệm chuyên sâu.",
        ]
    else:
        action = [
            "Tiếp tục theo dõi sức khỏe và khám theo nhu cầu/hướng dẫn của cơ sở y tế.",
            "Nếu đã có CBC, nên nhập vào Vòng 2 vì Vòng 1 thấp không có nghĩa là đã loại trừ Thalassemia.",
            "Nếu sau này xuất hiện thiếu máu, MCV/MCH thấp, tiền sử gia đình đáng chú ý hoặc có nhu cầu tư vấn trước hôn nhân/sinh sản, nên đánh giá lại.",
        ]
    return common + action


# ------------------------------------------------------------
# ROUND 2
# ------------------------------------------------------------

def calculate_round2_score(mcv, mch, rbc, rdw):
    """Sàng lọc CBC theo Phụ lục "Quy trình xét nghiệm sàng lọc, chẩn đoán Thalassemia"
    (Hướng dẫn chẩn đoán và điều trị một số bệnh lý huyết học, Bộ Y tế 2022):
    MCV < 85 fL và/hoặc MCH < 28 pg. Mentzer chỉ là chỉ số hỗ trợ."""
    reasons = []
    score = 0

    if mcv < 85:
        reasons.append("MCV <85 fL — đạt tiêu chí sàng lọc CBC theo phụ lục hướng dẫn BYT 2022")
        score += 2
    if mch < 28:
        reasons.append("MCH <28 pg — đạt tiêu chí sàng lọc CBC theo phụ lục hướng dẫn BYT 2022")
        score += 2
    if mcv < 80:
        score += 1
        reasons.append("MCV <80 fL — hồng cầu nhỏ")
    if mch < 27:
        score += 1
        reasons.append("MCH <27 pg — xu hướng hồng cầu nhược sắc")
    if rdw > 14:
        reasons.append("RDW >14% — kích thước hồng cầu không đồng đều theo hướng dẫn BYT 2022")
    if mcv < 85 and rbc >= 5.0:
        score += 1
        reasons.append("RBC tương đối cao trong bối cảnh MCV giảm")

    mentzer = _round_number(Decimal(str(mcv)) / Decimal(str(rbc)), 2) if rbc > 0 else None
    if mentzer is not None:
        reasons.append(f"Mentzer Index = {mentzer:.2f} — chỉ số hỗ trợ, không phải tiêu chuẩn chẩn đoán")
    return score, mentzer, reasons


def round2_category(score, mcv, mch=None):
    screening_trigger = (mcv < 85) or (mch is not None and mch < 28)
    if not screening_trigger:
        return ("THẤP", "CBC hiện tại chưa đạt tiêu chí sàng lọc MCV <85 fL và/hoặc MCH <28 pg trong phụ lục hướng dẫn BYT 2022; điều này không loại trừ hoàn toàn Thalassemia.")
    if score >= 6:
        return ("CAO", "CBC có tiêu chí sàng lọc đáng lưu ý (MCV <85 fL và/hoặc MCH <28 pg). Theo quy trình BYT 2022, nên đánh giá tình trạng sắt và xác định thành phần huyết sắc tố bằng điện di/HPLC; xét nghiệm gen tùy trường hợp.")
    return ("TRUNG BÌNH", "CBC đạt tiêu chí sàng lọc MCV <85 fL và/hoặc MCH <28 pg. Cần đánh giá thiếu sắt và nguyên nhân hồng cầu nhỏ; nếu phù hợp, thực hiện điện di/HPLC huyết sắc tố.")


def narrative_cbc_advice(hb_adjusted, mcv, mch, rbc, rdw, mentzer):
    findings, advice = [], []

    if mcv < 85:
        findings.append("MCV <85 fL — đạt tiêu chí sàng lọc CBC theo phụ lục BYT 2022.")
    if mch < 28:
        findings.append("MCH <28 pg — đạt tiêu chí sàng lọc CBC theo phụ lục BYT 2022.")
    if rdw > 14:
        findings.append("RDW tăng; cần lưu ý thiếu sắt hoặc các nguyên nhân khác của microcytosis.")
    if mcv < 85 and rbc >= 5.0:
        findings.append("RBC tương đối cao trong bối cảnh MCV thấp.")

    if mcv < 85 and mentzer < 13:
        findings.append("Mentzer Index <13: mẫu hình sàng lọc nghiêng về Thalassemia hơn thiếu sắt.")
        advice.append("Trao đổi với nhân viên y tế về HPLC/điện di hemoglobin; tùy trường hợp có thể cần xét nghiệm phân tử.")
    elif mcv < 85 and mentzer >= 13:
        findings.append("Mentzer Index ≥13: mẫu hình sàng lọc nghiêng về thiếu sắt hoặc nguyên nhân microcytosis khác.")
        advice.append("Cân nhắc đánh giá tình trạng sắt, đặc biệt Ferritin, theo chỉ định của nhân viên y tế.")

    if hb_adjusted < 8:
        advice.append("Hb sau hiệu chỉnh rất thấp: nên được đánh giá y tế sớm.")
    elif hb_adjusted < 10:
        advice.append("Hb sau hiệu chỉnh thấp đáng kể: nên khám và đánh giá nguyên nhân thiếu máu.")
    elif hb_adjusted < 12:
        advice.append("Hb sau hiệu chỉnh thấp/giáp ranh tùy nhóm đối tượng; cần đối chiếu tuổi, giới và khoảng tham chiếu của labo.")

    if not findings:
        findings.append("Chưa ghi nhận microcytosis/nhược sắc rõ trên CBC.")
    if not advice:
        advice.append("Tiếp tục đối chiếu với khoảng tham chiếu trên phiếu xét nghiệm và hướng dẫn của nhân viên y tế.")
    return findings, advice


# ------------------------------------------------------------
# ROUND 3 — BYT 2022 INTERPRETATION SUPPORT
# ------------------------------------------------------------

def round3_byt_interpretation(r3):
    notes = []
    hba, hba2, hbf, hbe = r3.get("hba"), r3.get("hba2"), r3.get("hbf"), r3.get("hbe")
    ferritin, tsat = r3.get("ferritin"), r3.get("transferrin_saturation")

    if hba is not None and 96.5 <= hba <= 98 and hba2 is not None and 2 <= hba2 <= 3.5 and hbf is not None and hbf < 1:
        notes.append("Thành phần Hb nằm trong khoảng HbA 96,5–98%, HbA2 2–3,5%, HbF <1% được nêu trong phụ lục BYT 2022; cần đối chiếu toàn bộ phiếu xét nghiệm và lâm sàng.")
    if hba2 is not None and hba2 > 3.5:
        notes.append("HbA2 >3,5% — bất thường được phụ lục BYT 2022 nêu là gợi ý Beta-thalassemia.")
    if hbf is not None and hbf >= 1:
        notes.append("HbF tăng so với ngưỡng <1% nêu trong phụ lục BYT 2022; cần đánh giá trong bối cảnh toàn bộ thành phần Hb.")
    if hbe is not None and hbe > 0:
        notes.append("Có HbE — phụ lục BYT 2022 nêu HbE dương tính là dấu hiệu cần đánh giá bệnh huyết sắc tố E/hemoglobinopathy.")

    combined = ((r3.get("genetic_result") or "") + " " + (r3.get("lab_conclusion") or "")).lower()
    for marker, label in [("hbh", "HbH"), ("hb bart", "Hb Bart’s"), ("hbcs", "HbCS"), ("constant spring", "Hb Constant Spring")]:
        if marker in combined:
            notes.append(f"Phiếu có đề cập {label}; theo hướng dẫn BYT 2022 cần đối chiếu với đánh giá Alpha-thalassemia và/hoặc xét nghiệm DNA.")

    if ferritin is not None and ferritin < 30:
        notes.append("Ferritin <30 ng/mL — mức được hướng dẫn BYT 2022 nêu trong chẩn đoán thiếu máu thiếu sắt; cần đối chiếu lâm sàng và các chỉ số sắt khác.")
    if tsat is not None and tsat < 30:
        notes.append("Độ bão hòa transferrin <30% — mức được hướng dẫn BYT 2022 nêu trong chẩn đoán thiếu máu thiếu sắt.")
    if not notes:
        notes.append("Chưa có mẫu hình tự động để đối chiếu từ các trường đã nhập. Cần xem toàn bộ phiếu xét nghiệm và đánh giá chuyên môn.")
    return notes


# ------------------------------------------------------------
# MEDICAL FACILITIES
# ------------------------------------------------------------
# Danh mục điều hướng cho prototype, KHÔNG phải chỉ định điều trị.
# Phân hạng/tên đơn vị có thể thay đổi; luôn kèm nút tra cứu Google Maps.

def _f(name, tier, note, query=None):
    q = quote(query or name)
    return {"name": name, "tier": tier, "note": note,
            "maps": f"https://www.google.com/maps/search/?api=1&query={q}"}


_TW = "Hạng đặc biệt – tuyến Trung ương"
_SPEC = "Ưu tiên khi cần đánh giá chuyên sâu huyết học"

MEDICAL_FACILITIES = {
    "Hà Nội": [
        _f("Bệnh viện Bạch Mai", _TW, _SPEC, "Bệnh viện Bạch Mai Hà Nội"),
        _f("Bệnh viện Hữu nghị Việt Đức", "Bệnh viện tuyến Trung ương", "Cơ sở chuyên sâu; phù hợp khi cần tuyến trên", "Bệnh viện Hữu nghị Việt Đức Hà Nội"),
        _f("Bệnh viện Trung ương Quân đội 108", _TW, "Cơ sở tuyến trên với nhiều chuyên khoa sâu"),
    ],
    "Hồ Chí Minh": [
        _f("Bệnh viện Chợ Rẫy", _TW, _SPEC, "Bệnh viện Chợ Rẫy Hồ Chí Minh"),
        _f("Bệnh viện Thống Nhất", "Hạng I – tuyến Trung ương", "Bệnh viện đa khoa tuyến trên", "Bệnh viện Thống Nhất Hồ Chí Minh"),
        _f("Bệnh viện Nhân Dân 115", "Hạng I", "Bệnh viện đa khoa tuyến chuyên sâu", "Bệnh viện Nhân Dân 115 Hồ Chí Minh"),
        _f("Bệnh viện Đại học Y Dược TP. Hồ Chí Minh", "Hạng I", "Bệnh viện trường đại học tuyến chuyên sâu"),
    ],
    "Đà Nẵng": [
        _f("Bệnh viện C Đà Nẵng", "Bệnh viện tuyến Trung ương", "Cơ sở tuyến Trung ương tại Đà Nẵng"),
        _f("Bệnh viện Đà Nẵng", "Hạng I", "Bệnh viện đa khoa lớn của thành phố"),
        _f("Bệnh viện Đại học Y – Dược, Đại học Đà Nẵng", "Bệnh viện trường đại học", "Có thể lựa chọn khi cần đánh giá chuyên khoa", "Bệnh viện Đại học Y Dược Đại học Đà Nẵng"),
    ],
    "Huế": [
        _f("Bệnh viện Trung ương Huế", _TW, "Ưu tiên khi cần đánh giá chuyên sâu"),
        _f("Bệnh viện Trường Đại học Y – Dược Huế", "Hạng I", "Bệnh viện trường đại học tuyến chuyên sâu", "Bệnh viện Trường Đại học Y Dược Huế"),
    ],
    "Cần Thơ": [
        _f("Bệnh viện Đa khoa Trung ương Cần Thơ", "Bệnh viện tuyến Trung ương", "Ưu tiên khi cần tuyến chuyên sâu tại ĐBSCL"),
        _f("Bệnh viện Đại học Y Dược Cần Thơ", "Hạng I", "Bệnh viện trường đại học tuyến chuyên sâu"),
    ],
    "Thái Nguyên": [_f("Bệnh viện Trung ương Thái Nguyên", _TW, "Ưu tiên khi cần đánh giá chuyên sâu")],
    "Hải Phòng": [_f("Bệnh viện Hữu nghị Việt Tiệp", "Hạng I", "Bệnh viện đa khoa tuyến chuyên sâu", "Bệnh viện Hữu nghị Việt Tiệp Hải Phòng")],
    "Quảng Ninh": [
        _f("Bệnh viện Bãi Cháy", "Hạng I", "Bệnh viện đa khoa lớn của tỉnh", "Bệnh viện Bãi Cháy Quảng Ninh"),
        _f("Bệnh viện Đa khoa tỉnh Quảng Ninh", "Hạng I", "Bệnh viện đa khoa tuyến tỉnh"),
    ],
    "Đắk Lắk": [_f("Bệnh viện Đa khoa vùng Tây Nguyên", "Hạng I", "Cơ sở tuyến chuyên sâu khu vực Tây Nguyên")],
    "Nghệ An": [_f("Bệnh viện Hữu nghị Đa khoa Nghệ An", "Hạng I", "Bệnh viện đa khoa tuyến tỉnh")],
    "Đồng Nai": [_f("Bệnh viện Đa khoa Đồng Nai", "Hạng I", "Bệnh viện đa khoa tuyến tỉnh")],
    "Đồng Tháp": [_f("Bệnh viện Đa khoa Đồng Tháp", "Hạng I", "Bệnh viện đa khoa tuyến tỉnh")],
    "An Giang": [_f("Bệnh viện Đa khoa khu vực An Giang", "Hạng I", "Bệnh viện đa khoa tuyến tỉnh/khu vực")],
}

# Các tỉnh có một bệnh viện đa khoa tuyến tỉnh hạng I.
for _p in ["Thanh Hóa", "Hà Tĩnh", "Khánh Hòa", "Lâm Đồng", "Quảng Ngãi", "Gia Lai", "Bắc Ninh",
           "Phú Thọ", "Lào Cai", "Hưng Yên", "Ninh Bình", "Tây Ninh", "Vĩnh Long", "Cà Mau"]:
    MEDICAL_FACILITIES[_p] = [_f(f"Bệnh viện Đa khoa tỉnh {_p}", "Hạng I", "Bệnh viện đa khoa tuyến tỉnh")]

# Các tỉnh còn lại: cơ sở đa khoa tuyến tỉnh.
for _p in ["Điện Biên", "Lai Châu", "Sơn La", "Tuyên Quang"]:
    MEDICAL_FACILITIES[_p] = [_f(f"Bệnh viện Đa khoa tỉnh {_p}", "Tuyến tỉnh", "Cơ sở đa khoa tuyến tỉnh")]

MEDICAL_FACILITIES.update({
    "Cao Bằng": [
        _f("Bệnh viện Đa khoa tỉnh Cao Bằng", "Tuyến tỉnh", "Cơ sở đa khoa tuyến tỉnh; có thể liên hệ khoa Huyết học/Truyền máu để hỏi khả năng xét nghiệm chuyên sâu"),
        _f("Bệnh viện Y học cổ truyền Cao Bằng", "Tuyến tỉnh", "Cơ sở y tế trên địa bàn tỉnh"),
    ],
    "Lạng Sơn": [
        _f("Bệnh viện Đa khoa tỉnh Lạng Sơn", "Tuyến tỉnh", "Cơ sở đa khoa tuyến tỉnh"),
        _f("Bệnh viện Phổi Lạng Sơn", "Chuyên khoa", "Cơ sở chuyên khoa trên địa bàn tỉnh"),
    ],
    "Quảng Trị": [
        _f("Bệnh viện Hữu nghị Việt Nam – Cuba Đồng Hới", "Tuyến Trung ương", "Cơ sở tuyến trên phục vụ khu vực; nên xác nhận khoa Huyết học/Truyền máu trước khi đến", "Bệnh viện Hữu nghị Việt Nam Cuba Đồng Hới"),
        _f("Bệnh viện Đa khoa tỉnh Quảng Trị", "Tuyến tỉnh", "Cơ sở đa khoa tuyến tỉnh"),
    ],
})

# Tên tỉnh cũ → tỉnh/thành dùng để tra danh mục cơ sở.
PROVINCE_FACILITY_ALIASES = {
    "Thừa Thiên Huế": "Huế", "Quảng Nam": "Đà Nẵng", "Bình Định": "Gia Lai",
    "Ninh Thuận": "Khánh Hòa", "Phú Yên": "Đắk Lắk", "Đắk Nông": "Lâm Đồng",
    "Bình Thuận": "Lâm Đồng", "Kon Tum": "Quảng Ngãi", "Yên Bái": "Lào Cai",
    "Bắc Kạn": "Thái Nguyên", "Hòa Bình": "Phú Thọ", "Vĩnh Phúc": "Phú Thọ",
    "Hà Nam": "Ninh Bình", "Nam Định": "Ninh Bình", "Quảng Bình": "Quảng Trị",
    "Bà Rịa - Vũng Tàu": "Hồ Chí Minh", "Bình Dương": "Hồ Chí Minh", "Long An": "Tây Ninh",
    "Tiền Giang": "Đồng Tháp", "Bến Tre": "Vĩnh Long", "Trà Vinh": "Vĩnh Long",
    "Sóc Trăng": "Cần Thơ", "Hậu Giang": "Cần Thơ", "Kiên Giang": "An Giang",
}


def recommended_facilities(province):
    key = PROVINCE_FACILITY_ALIASES.get(province, province)
    return MEDICAL_FACILITIES.get(key, [])[:5]


def all_medical_facilities():
    rows = []
    for province, facilities in MEDICAL_FACILITIES.items():
        for facility in facilities:
            item = dict(facility)
            item["province"] = province
            rows.append(item)
    return rows


# ------------------------------------------------------------
# WORD REPORT
# ------------------------------------------------------------

def make_word(patient, r1_score, r1_category, r1_reasons, r2=None):
    doc = Document()
    doc.add_heading("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", level=3)
    doc.add_heading("PHIẾU SÀNG LỌC VÀ PHÂN TẦNG NGUY CƠ THALASSEMIA", level=1)
    doc.add_paragraph("Công cụ hỗ trợ sàng lọc; không thay thế chẩn đoán hoặc chỉ định của nhân viên y tế.")

    doc.add_heading("I. THÔNG TIN BỆNH NHÂN", level=2)
    for label, value in [
        ("Họ và tên", patient["full_name"]), ("Ngày sinh", patient["birth_date"]),
        ("Tuổi", patient["age"]), ("Giới tính", patient["gender"]),
        ("Số điện thoại", patient["phone"]), ("Địa chỉ hiện tại", patient["current_address"]),
        ("Tỉnh/thành", patient["province"]), ("Phường/xã/đặc khu", patient["commune"]),
    ]:
        doc.add_paragraph(f"{label}: {value}")

    doc.add_heading("II. VÒNG 1", level=2)
    doc.add_paragraph(f"Điểm: {r1_score}/{ROUND1_MAX_SCORE}")
    doc.add_paragraph(f"Mức nguy cơ: {r1_category}")
    for item in r1_reasons:
        doc.add_paragraph(f"- {item}")

    if r2:
        doc.add_heading("III. VÒNG 2", level=2)
        for line in [
            f"Khoảng độ cao: {r2['altitude_choice']}",
            f"Hiệu chỉnh Hb: -{r2['adjustment']:.1f} g/dL",
            f"Hb thực đo sau quy đổi: {r2['hb']:.1f} g/dL",
            f"Hb sau hiệu chỉnh: {r2['hb_adjusted']:.1f} g/dL",
            f"MCV: {r2['mcv']:.1f} fL",
            f"MCH: {r2['mch']:.1f} pg",
            f"RBC: {r2['rbc']:.2f} T/L",
            f"RDW-CV: {r2['rdw']:.1f}%",
            f"Mentzer Index: {r2['mentzer']:.2f}",
            f"Điểm CBC prototype: {r2['score']}",
            f"Mức nguy cơ Vòng 2: {r2['category']}",
            f"Nhận định: {r2['conclusion']}",
        ]:
            doc.add_paragraph(line)

        doc.add_heading("IV. PHÂN TÍCH SƠ BỘ", level=2)
        for finding in r2["findings"]:
            doc.add_paragraph(f"- {finding}")
        doc.add_heading("V. KHUYẾN NGHỊ", level=2)
        for advice in r2["advice"]:
            doc.add_paragraph(f"- {advice}")
    else:
        doc.add_heading("III. THEO DÕI", level=2)
        doc.add_paragraph("Vòng 2 chưa được mở. Tiếp tục theo dõi sức khỏe và đánh giá lại nếu xuất hiện bất thường.")

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out


# ============================================================
# RESEARCH OVERVIEW DIALOG
# ============================================================

@st.dialog("Tìm hiểu về nghiên cứu và cơ sở khoa học", width="large")
def show_research_overview():
    st.markdown("""
### 1. Lý do khoa học – y tế để phát triển hệ thống

Thalassemia là nhóm bệnh lý di truyền do giảm tổng hợp chuỗi globin, có thể biểu hiện từ người mang gen gần như không triệu chứng đến các thể thiếu máu nặng. Trong thực hành sàng lọc, **công thức máu (CBC) và các chỉ số hồng cầu** có thể cung cấp dấu hiệu ban đầu của kiểu hình hồng cầu nhỏ, nhược sắc; khi có nghi ngờ, người bệnh cần được đánh giá tiếp bằng các xét nghiệm hemoglobin chuyên sâu và/hoặc xét nghiệm phân tử tùy trường hợp.

Vấn đề thực tiễn là khả năng tiếp cận các tầng xét nghiệm không giống nhau. CBC thường dễ tiếp cận hơn so với HPLC/điện di hemoglobin hoặc xét nghiệm gen. Vì vậy hệ thống này được xây dựng như **một lớp hỗ trợ sàng lọc – giải thích CBC – phân tầng – điều hướng**, không thay thế bác sĩ và không đưa ra chẩn đoán xác định.

### 2. Vì sao Vòng 1 có 20 câu hỏi?

Vòng 1 không nhằm loại người tham gia khỏi quá trình sàng lọc. Nó nhằm thu thập những thông tin có thể làm thay đổi mức độ cần lưu ý, đồng thời đưa ra khuyến nghị phù hợp trước khi xem CBC. Các nhóm câu hỏi được lựa chọn dựa trên 4 nhóm thông tin: **tiền sử gia đình, tiền sử huyết học cá nhân, dấu hiệu hỗ trợ và khả năng tiếp cận xét nghiệm**.

**Q1–Q5 – Tiền sử gia đình:** Thalassemia có tính di truyền. Thông tin về người thân mắc bệnh, mang gen, thiếu máu/hồng cầu nhỏ hoặc truyền máu nhiều lần có thể làm tăng lý do cần xem xét sàng lọc.

**Q6–Q12 – Tiền sử bản thân:** từng được thông báo thiếu máu, MCV/MCH thấp, từng xét nghiệm hemoglobinopathy, HbE, truyền máu hoặc thiếu máu kéo dài giúp hệ thống biết người tham gia đã có những dấu hiệu/lịch sử nào cần được đối chiếu với CBC hiện tại.

**Q13–Q18 – Dấu hiệu hỗ trợ:** mệt mỏi, chóng mặt, da niêm nhợt, vàng da, lách to hoặc tiền sử biến chứng huyết học có thể gợi ý vấn đề huyết học, nhưng **không đặc hiệu cho Thalassemia**. Vì vậy các câu này chỉ có vai trò hỗ trợ, không được dùng để chẩn đoán.

**Q19–Q20 – Khả năng tiếp cận xét nghiệm:** giúp hệ thống hiểu người tham gia đã có CBC hay gặp khó khăn khi tiếp cận xét nghiệm chuyên sâu. Hai câu này **không được cộng vào điểm nguy cơ sinh học**, vì chi phí, khoảng cách và thời gian không phải là đặc điểm bệnh sinh của Thalassemia.

### 3. Ý nghĩa của các thông số huyết học ở Vòng 2

**Hb – Hemoglobin:** phản ánh lượng hemoglobin trong máu và là chỉ số quan trọng để đánh giá thiếu máu. Hb cần được diễn giải theo tuổi, giới, thai kỳ và các yếu tố bối cảnh; hệ thống không dùng một giá trị Hb đơn độc để chẩn đoán Thalassemia.

**MCV – Mean Corpuscular Volume:** thể tích trung bình của hồng cầu. MCV giảm cho thấy hồng cầu nhỏ (microcytosis), là một dấu hiệu quan trọng khi xem xét Thalassemia nhưng cũng gặp trong thiếu sắt và các nguyên nhân khác.

**MCH – Mean Corpuscular Hemoglobin:** lượng hemoglobin trung bình trong mỗi hồng cầu. MCH giảm biểu hiện xu hướng nhược sắc và thường đi cùng microcytosis trong Thalassemia trait.

**RBC – số lượng hồng cầu:** cho biết số lượng hồng cầu. Trong một số trường hợp Thalassemia trait, RBC có thể tương đối cao dù MCV/MCH giảm. Vì vậy RBC giúp đặt MCV/MCH vào bối cảnh thay vì nhìn một chỉ số đơn lẻ.

**RDW – Red Cell Distribution Width:** phản ánh mức độ biến thiên kích thước hồng cầu. RDW tăng có thể gặp trong thiếu sắt và nhiều tình trạng khác; RDW không đủ đặc hiệu để phân biệt Thalassemia với thiếu sắt.

**Mentzer Index = MCV / RBC:** là chỉ số sàng lọc đơn giản được dùng để định hướng giữa kiểu hình gợi ý Thalassemia và thiếu sắt. Đây chỉ là công cụ hỗ trợ, không phải xét nghiệm xác định và có thể sai trong các trường hợp phối hợp bệnh lý.

### 4. Vì sao vẫn cần xét nghiệm chuyên sâu?

CBC chỉ cho thấy **kiểu hình huyết học**, không trực tiếp xác định loại hemoglobin bất thường hay biến thể gen. Khi có nghi ngờ phù hợp, các bước tiếp theo có thể bao gồm đánh giá tình trạng sắt (ví dụ ferritin), phân tích hemoglobin bằng HPLC/điện di và xét nghiệm phân tử khi có chỉ định. Với β-thalassemia trait, HbA₂ tăng có thể là dấu hiệu hỗ trợ; trong một số thể alpha-thalassemia, xét nghiệm phân tử có vai trò quan trọng vì điện di có thể không phát hiện được.

### 5. Vì sao tất cả người tham gia đều được vào Vòng 2?

Vòng 1 là **sàng lọc ban đầu và giáo dục sức khỏe**, không phải một phép loại trừ. Một Vòng 1 “thấp” không có nghĩa là không mang gen. Vì vậy nếu người tham gia đã có CBC, họ vẫn có thể nhập CBC để hệ thống phân tích các chỉ số huyết học.

### 6. Ý nghĩa của kết quả hệ thống

Kết quả của hệ thống được diễn đạt theo hướng **“gợi ý – cần đánh giá thêm – nên trao đổi với cơ sở y tế”**, không phải “mắc bệnh/không mắc bệnh”. Mục tiêu là giúp người dùng hiểu kết quả CBC, nhận biết khi nào cần đánh giá tiếp và tìm đúng cơ sở y tế có năng lực phù hợp.

### 7. Giá trị cộng đồng của nghiên cứu

Giá trị của hệ thống không nằm ở việc thay thế bệnh viện. Giá trị nằm ở việc tận dụng những dữ liệu huyết học cơ bản đã có, đặc biệt là CBC, để **giảm khoảng cách giữa cộng đồng và xét nghiệm chuyên sâu**. Hệ thống hướng tới sử dụng nguồn lực hợp lý hơn: người có ít dấu hiệu đáng lưu ý có thể được hướng dẫn theo dõi phù hợp; người có CBC gợi ý cần đánh giá thêm có thể được hướng dẫn đến cơ sở chuyên môn thay vì tự tìm kiếm hoặc di chuyển không cần thiết.

### 8. Giới hạn cần biết

Các điểm số và ngưỡng của Vòng 1/Vòng 2 trong phiên bản hiện tại là **prototype**, chưa được thẩm định trên một quần thể người Việt Nam đủ lớn. Vì vậy hệ thống chỉ có giá trị hỗ trợ sàng lọc và nghiên cứu phát triển, không thay thế chẩn đoán lâm sàng.
    """)
    st.divider()
    st.subheader("Tài liệu chuyên môn tham khảo")
    st.link_button("WHO 2024 – Guideline on haemoglobin cutoffs", "https://www.who.int/publications/i/item/9789240088542", icon=":material/open_in_new:", use_container_width=True)
    st.link_button("GeneReviews – Beta-Thalassemia", "https://www.ncbi.nlm.nih.gov/books/NBK1426/", icon=":material/open_in_new:", use_container_width=True)
    st.link_button("ACOG – Carrier Screening for Genetic Conditions", "https://www.acog.org/clinical/clinical-guidance/committee-opinion/articles/2017/03/carrier-screening-for-genetic-conditions", icon=":material/open_in_new:", use_container_width=True)
    st.link_button("ACOG – Carrier Screening for Hemoglobinopathies", "https://www.acog.org/womens-health/faqs/carrier-screening-for-hemoglobinopathies", icon=":material/open_in_new:", use_container_width=True)


# ============================================================
# HEADER
# ============================================================

st.markdown(
    '<div class="hematology-brand"><div class="eyebrow">HEMATOLOGY · COMMUNITY SCREENING RESEARCH</div>'
    '<div class="title title-main">HỆ THỐNG HỖ TRỢ SÀNG LỌC THALASSEMIA</div></div>'
    '<div class="section-rule"></div>',
    unsafe_allow_html=True,
)

with st.expander("Giới thiệu đề tài và cơ sở nghiên cứu", expanded=False, icon=":material/menu_book:"):
    st.markdown("## Lời nói đầu")
    st.markdown("""
    Thalassemia là nhóm bệnh lý huyết sắc tố di truyền do giảm hoặc mất khả năng tổng hợp một hoặc nhiều chuỗi globin. Phổ biểu hiện rất rộng, từ người mang gen có biểu hiện huyết học nhẹ hoặc gần như không có triệu chứng đến các thể bệnh thiếu máu nặng cần chăm sóc y tế lâu dài. Vì có tính di truyền, việc nhận diện người mang gen không chỉ có ý nghĩa đối với cá nhân mà còn có giá trị trong tư vấn di truyền và dự phòng nguy cơ cho thế hệ sau.

    Trong thực hành sàng lọc, công thức máu (CBC) và các chỉ số hồng cầu là những dữ liệu ban đầu có giá trị để nhận diện kiểu hình hồng cầu nhỏ, nhược sắc. Tuy nhiên, CBC không thể tự xác định loại hemoglobinopathy hay biến thể gen. Khi có dấu hiệu nghi ngờ, người tham gia cần được đánh giá tiếp bằng tình trạng sắt, HPLC/điện di hemoglobin và trong những trường hợp phù hợp là xét nghiệm phân tử.

    Từ nhu cầu kết nối giữa sàng lọc ban đầu – dữ liệu CBC – giải thích kết quả – khuyến nghị – tiếp cận cơ sở y tế, đề tài “Hệ thống hỗ trợ sàng lọc Thalassemia trong cộng đồng” được xây dựng. Hệ thống được định hướng như một công cụ hỗ trợ sức khỏe cộng đồng, không thay thế bác sĩ và không đưa ra chẩn đoán xác định.
    """)

    st.markdown("## 1. Lý do chọn đề tài")
    st.markdown("""
    Thứ nhất, Thalassemia là bệnh lý di truyền có thể tồn tại âm thầm trong cộng đồng. Người mang gen có thể không có biểu hiện lâm sàng rõ ràng nhưng vẫn có ý nghĩa về mặt di truyền. Do đó, chỉ dựa vào triệu chứng để nhận diện là không đủ; cần có chiến lược sàng lọc phù hợp để phát hiện những trường hợp cần được đánh giá sâu hơn.

    Thứ hai, CBC là xét nghiệm có tính nền tảng trong thực hành huyết học. Các thông số như Hb, MCV, MCH, RBC và RDW cung cấp thông tin về số lượng và đặc điểm hình thái của quần thể hồng cầu. Mẫu hình hồng cầu nhỏ, nhược sắc có thể gặp trong Thalassemia nhưng cũng gặp trong thiếu sắt và nhiều nguyên nhân khác. Vì vậy, vấn đề không chỉ là “đọc một con số”, mà là đặt các chỉ số vào đúng bối cảnh sàng lọc.

    Thứ ba, các xét nghiệm chuyên sâu như HPLC, điện di hemoglobin và xét nghiệm phân tử có vai trò quan trọng nhưng không phải lúc nào cũng là bước đầu tiên hoặc dễ tiếp cận đối với mọi người dân. Một hệ thống hỗ trợ có thể giúp người sử dụng hiểu dữ liệu CBC đang có, nhận biết khi nào cần đánh giá thêm và tránh tâm lý tự kết luận chỉ từ một chỉ số.

    Thứ tư, việc xây dựng hệ thống theo hướng không loại người tham gia chỉ dựa trên bảng câu hỏi Vòng 1 giúp giảm nguy cơ bỏ sót người mang gen không có triệu chứng. Điểm Vòng 1 chỉ được sử dụng để phân tầng và đưa ra khuyến nghị; người tham gia vẫn có thể tiếp tục Vòng 2 để nhập CBC nếu có.

    Từ những cơ sở trên, đề tài được lựa chọn với mục tiêu xây dựng một cầu nối hỗ trợ sàng lọc trong cộng đồng, bắt đầu từ thông tin nguy cơ và CBC, sau đó định hướng đánh giá chuyên sâu và tiếp cận cơ sở y tế khi cần thiết.
    """)

    st.markdown("## 2. Mục tiêu đề tài")
    st.markdown("### 2.1. Mục tiêu tổng quát")
    st.write("Xây dựng hệ thống hỗ trợ sàng lọc Thalassemia trong cộng đồng dựa trên thông tin tiền sử và các chỉ số huyết học cơ bản, nhằm hỗ trợ nhận diện trường hợp cần được đánh giá thêm và định hướng tiếp cận cơ sở y tế phù hợp.")
    st.markdown("### 2.2. Mục tiêu cụ thể")
    st.markdown("""
    1. Xây dựng bộ câu hỏi sàng lọc ban đầu nhằm thu thập tiền sử gia đình, tiền sử huyết học và các yếu tố liên quan.
    2. Phân tầng mức độ cần lưu ý ở Vòng 1 nhưng không sử dụng điểm số để chẩn đoán hoặc loại trừ Thalassemia.
    3. Cho phép tất cả người tham gia tiếp tục Vòng 2 và nhập kết quả CBC khi có.
    4. Phân tích Hb, MCV, MCH, RBC, RDW và Mentzer theo mô hình sàng lọc prototype.
    5. Định hướng các bước đánh giá tiếp theo như tình trạng sắt, HPLC/điện di hemoglobin hoặc xét nghiệm phân tử khi phù hợp.
    6. Hỗ trợ người tham gia tiếp cận cơ sở y tế có năng lực chuyên môn phù hợp.
    """)

    st.markdown("## 3. Cơ sở khoa học")
    st.markdown("### 3.1. Cơ sở bệnh sinh và ý nghĩa của sàng lọc")
    st.markdown("""
    Thalassemia phát sinh từ các biến thể di truyền làm giảm hoặc mất tổng hợp chuỗi globin. Sự mất cân bằng chuỗi globin có thể dẫn đến sinh hồng cầu không hiệu quả và mức độ thiếu máu khác nhau tùy kiểu gen, kiểu hình và thể bệnh. Ở người mang gen, biểu hiện thường nhẹ hơn đáng kể so với các thể bệnh Thalassemia có triệu chứng.

    Vì tình trạng mang gen có thể không biểu hiện rõ trên lâm sàng, sàng lọc dựa trên dữ liệu huyết học và tiền sử có vai trò nhận diện những trường hợp nên được xác nhận bằng xét nghiệm chuyên sâu. Đây là lý do hệ thống được thiết kế theo mô hình nhiều tầng thay vì cố gắng đưa ra chẩn đoán từ một bảng điểm duy nhất.
    """)
    st.markdown("### 3.2. Vai trò của công thức máu (CBC)")
    st.markdown("""
    Hb (hemoglobin): phản ánh lượng hemoglobin trong máu và là chỉ số quan trọng khi đánh giá thiếu máu. Giá trị Hb cần được diễn giải theo tuổi, giới, thai kỳ, độ cao và các yếu tố liên quan; hệ thống sử dụng hiệu chỉnh độ cao theo hướng dẫn WHO 2024 ở mức prototype.

    MCV (Mean Corpuscular Volume): phản ánh thể tích trung bình của hồng cầu. MCV giảm cho thấy hồng cầu nhỏ (microcytosis). Đây là một dấu hiệu thường gặp trong Thalassemia trait nhưng cũng có thể gặp trong thiếu sắt.

    MCH (Mean Corpuscular Hemoglobin): phản ánh lượng hemoglobin trung bình trong một hồng cầu. MCH thấp thường đi cùng kiểu hình hồng cầu nhược sắc và có thể hỗ trợ nhận diện trường hợp cần đánh giá thêm.

    RBC (số lượng hồng cầu): cung cấp thêm thông tin về số lượng hồng cầu. Trong một số trường hợp microcytosis, số lượng hồng cầu tương đối cao có thể hỗ trợ phân biệt Thalassemia trait với thiếu sắt, nhưng không đủ để xác định chẩn đoán.

    RDW (Red Cell Distribution Width): phản ánh mức độ biến thiên kích thước hồng cầu. RDW có thể hỗ trợ diễn giải bối cảnh microcytosis, nhưng không có tính đặc hiệu đủ cao để dùng riêng cho chẩn đoán Thalassemia.
    """)
    st.markdown("### 3.3. Chỉ số Mentzer và các chỉ số phân biệt")
    st.markdown("""
    Mentzer Index = MCV / RBC là một chỉ số kinh điển được sử dụng như công cụ định hướng giữa thiếu sắt và Thalassemia trong một số bối cảnh sàng lọc. Tuy nhiên, đây chỉ là chỉ số hỗ trợ, không phải xét nghiệm xác nhận. Độ chính xác của các chỉ số phân biệt thay đổi theo quần thể, tuổi, tình trạng thiếu sắt đồng thời và thể bệnh.

    Vì vậy, hệ thống không sử dụng Mentzer đơn độc. Kết quả được đặt cùng MCV, MCH, RBC, RDW và bối cảnh Vòng 1 để tạo ra một mức nguy cơ sàng lọc prototype.
    """)
    st.markdown("### 3.4. Vì sao phải đánh giá thiếu sắt?")
    st.markdown("""
    Thiếu sắt là một nguyên nhân phổ biến của thiếu máu hồng cầu nhỏ và có thể tạo ra kiểu hình CBC tương tự Thalassemia. Do đó, khi phát hiện microcytosis/nhược sắc, cần xem xét tình trạng sắt, thường với ferritin và các xét nghiệm chuyển hóa sắt phù hợp theo đánh giá chuyên môn.

    Đặc biệt, người có khả năng mang Thalassemia không nên tự sử dụng sắt kéo dài chỉ vì thấy Hb thấp hoặc MCV thấp. Việc bổ sung sắt nên dựa trên bằng chứng thiếu sắt và chỉ định phù hợp.
    """)
    st.markdown("### 3.5. Vì sao CBC chưa đủ để xác định Thalassemia?")
    st.markdown("""
    CBC phản ánh kiểu hình huyết học, trong khi Thalassemia là một nhóm rối loạn có nguyên nhân di truyền. Vì vậy, cùng một mẫu hình hồng cầu nhỏ có thể xuất hiện trong nhiều tình trạng khác nhau.

    Khi kết quả sàng lọc gợi ý, bước tiếp theo có thể bao gồm HPLC hoặc điện di hemoglobin để đánh giá thành phần hemoglobin. Trong những trường hợp cần xác định biến thể hoặc khi kết quả xét nghiệm huyết học chưa giải thích được tình trạng nghi ngờ, xét nghiệm phân tử có thể được cân nhắc theo chỉ định chuyên môn. Với α-thalassemia, xét nghiệm di truyền có thể đặc biệt hữu ích vì các xét nghiệm hemoglobin thường quy không phải lúc nào cũng xác định được tình trạng mang gen.
    """)
    st.markdown("### 3.6. Cơ sở thiết kế Vòng 1 → Vòng 2")
    st.markdown("""
    Vòng 1 thu thập thông tin nguy cơ và tiền sử để tạo bối cảnh trước khi xem CBC. Điểm số chỉ nhằm phân tầng ban đầu.

    Vòng 2 tiếp nhận CBC và phân tích các chỉ số hồng cầu. Đây là tầng có tính khách quan hơn vì sử dụng dữ liệu xét nghiệm thực tế của người tham gia.

    Không có bước nào được thiết kế để “chẩn đoán Thalassemia”. Kết quả cuối cùng được diễn đạt bằng các mức “thấp – trung bình – cao” về mức độ cần đánh giá thêm, kèm khuyến nghị bước tiếp theo. Người tham gia vẫn có thể cần xét nghiệm chuyên sâu và đánh giá bởi nhân viên y tế.
    """)
    st.markdown("### 3.7. Phạm vi và giới hạn của mô hình")
    st.markdown("""
    Phiên bản hiện tại là prototype hỗ trợ sàng lọc, trong đó các điểm số và ngưỡng phân tầng chưa được thẩm định trên một quần thể người Việt Nam đủ lớn. Do đó, kết quả không được sử dụng như tiêu chuẩn chẩn đoán, không thay thế xét nghiệm chuyên sâu và không được dùng để tự quyết định điều trị.

    Mục tiêu của hệ thống là tạo ra một quy trình dễ tiếp cận: nhận diện nguy cơ → xem CBC → giải thích → khuyến nghị → điều hướng, sau đó việc xác nhận và xử trí thuộc về cơ sở y tế có thẩm quyền.
    """)

    st.markdown("## 4. Tài liệu tham khảo")
    refs = [
        ("World Health Organization. Guideline on haemoglobin cutoffs to define anaemia in individuals and populations. 2024.", "https://www.who.int/publications/i/item/9789240088542"),
        ("World Health Organization. WHO guidelines on best practices in the measurement of haemoglobin. 2024.", "https://www.who.int/publications/b/71578"),
        ("GeneReviews®. Beta-Thalassemia. University of Washington, Seattle. Updated February 12, 2026.", "https://www.ncbi.nlm.nih.gov/books/NBK1426/"),
        ("American College of Obstetricians and Gynecologists (ACOG). Carrier Screening for Genetic Conditions.", "https://www.acog.org/clinical/clinical-guidance/committee-opinion/articles/2017/03/carrier-screening-for-genetic-conditions"),
        ("American College of Obstetricians and Gynecologists (ACOG). Hemoglobinopathies in Pregnancy.", "https://www.acog.org/clinical/clinical-guidance/practice-advisory/articles/2022/08/hemoglobinopathies-in-pregnancy"),
        ("American College of Obstetricians and Gynecologists (ACOG). Carrier Screening for Hemoglobinopathies.", "https://www.acog.org/womens-health/faqs/carrier-screening-for-hemoglobinopathies"),
        ("Prevalence of Thalassemia in the Vietnamese Population and Building a Clinical Decision Support System for Prenatal Screening for Thalassemia.", "https://pmc.ncbi.nlm.nih.gov/articles/PMC10171208/"),
        ("Thalassemia among ethnic minorities in Vietnam: prevalence and molecular characterization.", "https://pmc.ncbi.nlm.nih.gov/articles/PMC5496844/"),
        ("A Review of Artificial Intelligence and Machine Learning Applications in Thalassemia Screening and Diagnosis.", "https://pmc.ncbi.nlm.nih.gov/articles/PMC10177591/"),
        ("Thalassaemia International Federation (TIF). Guidelines and Reports.", "https://thalassaemia.org.cy/publications/tif-publications/guidelines-and-reports/"),
    ]
    for i, (title, url) in enumerate(refs, 1):
        st.markdown(f"{i}. [{title}]({url})")
        st.caption(url)

if st.button("Tìm hiểu về nghiên cứu và cơ sở khoa học", icon=":material/school:"):
    show_research_overview()

st.markdown(
    '<div class="contact-card"><div class="label">Liên hệ tác giả</div>'
    '<div class="name">Nguyễn Đắc Lâu</div>'
    '<div>Điện thoại: <a href="tel:0357930820">0357930820</a></div>'
    '<div>Email: <a href="mailto:nguyendaclau2309@gmail.com">nguyendaclau2309@gmail.com</a></div>'
    '<div class="small-note">Tác giả phát triển nguyên mẫu hệ thống hỗ trợ sàng lọc Thalassemia.</div></div>',
    unsafe_allow_html=True,
)

# ============================================================
# ACCESS / SIDEBAR / ADMIN
# ============================================================

render_auth_sidebar()

auth_user = current_auth_user()

if auth_user and st.session_state.get("auth_page") == PAGE_ADMIN:
    render_admin_console(auth_user)
    st.stop()

operator_mode = auth_user is not None
operator_username = auth_user["username"] if auth_user else None

if operator_mode:
    st.success(
        f"**Chế độ nhập giúp người tham gia** — thao tác đang được ghi nhận bởi "
        f"**{auth_user['full_name']}** (#{auth_user['username']}).",
        icon=":material/assignment_ind:",
    )

# ============================================================
# PATIENT PROFILE
# ============================================================

st.header(":material/badge: Thông tin bệnh nhân")
st.caption(
    "Nhập hồ sơ người tham gia. "
    + ("Chế độ nhập giúp: hệ thống ghi nhận tài khoản nhân sự thực hiện."
       if operator_mode else "Người tham gia tự nhập thông tin của mình.")
)

# A. Thông tin cơ bản — form riêng
with st.form("patient_basic_form_v22", clear_on_submit=False):
    with st.container(border=True):
        p1, p2, p3 = st.columns(3)
        with p1:
            full_name = st.text_input("Họ và tên *", placeholder="Nguyễn Văn A", key="v22_full_name")
            birth_date = st.date_input(
                "Ngày sinh *", value=date(2000, 1, 1),
                min_value=BIRTH_DATE_MIN, max_value=BIRTH_DATE_MAX,
                format="DD/MM/YYYY", key="v22_birth_date",
            )
        with p2:
            phone_raw = st.text_input("Số điện thoại *", placeholder="09xxxxxxxx", key="v22_phone")
            gender = st.selectbox("Giới tính *", ["Nam", "Nữ", "Khác"], key="v22_gender")
        with p3:
            current_address = st.text_input("Địa chỉ hiện tại *", placeholder="Số nhà/thôn/tổ/đường", key="v22_address")

        st.caption(
            "Nhập thông tin cơ bản rồi chọn địa giới hành chính bên dưới. "
            "Danh sách phường/xã sẽ cập nhật ngay khi đổi tỉnh/thành."
        )
        basic_continue = st.form_submit_button(
            "Tiếp tục", icon=":material/arrow_forward:", type="primary", use_container_width=True
        )

# B. Địa giới — cố ý đặt NGOÀI form để selectbox tỉnh rerun ngay khi đổi.
st.subheader(":material/location_on: Địa giới hành chính hiện tại")

try:
    admin_hierarchy = load_admin_hierarchy()
    admin_data_ok = True
except Exception as exc:
    admin_hierarchy = {}
    admin_data_ok = False
    st.error("Không tải được danh mục tỉnh/thành và phường/xã/đặc khu. Vui lòng tải lại trang hoặc thử lại sau.")
    st.caption(f"Nguồn dữ liệu: {ADMIN_DATA_SOURCE_URL}")
    st.code(str(exc))

selected_province = ""
commune_value = ""

if admin_data_ok:
    provinces_list = sorted(admin_hierarchy.keys(), key=str.casefold)
    province_choice = st.selectbox(
        "Tỉnh / thành phố *", ["— Chọn tỉnh/thành —"] + provinces_list, key="v22_profile_province"
    )

    if province_choice == "— Chọn tỉnh/thành —":
        st.selectbox("Phường / xã / đặc khu *", ["— Chọn tỉnh/thành trước —"],
                     disabled=True, key="v22_profile_commune_empty")
    else:
        selected_province = province_choice
        available_communes = list(admin_hierarchy.get(selected_province, []))
        # Key phụ thuộc tỉnh để không tái sử dụng lựa chọn của tỉnh trước đó.
        safe_province = selected_province.replace(" ", "_").replace("/", "_").replace("\\", "_")
        commune_choice = st.selectbox(
            "Phường / xã / đặc khu *",
            ["— Chọn phường/xã/đặc khu —"] + available_communes,
            key=f"v22_profile_commune_{safe_province}",
        )
        if commune_choice != "— Chọn phường/xã/đặc khu —":
            commune_value = commune_choice
        if not available_communes:
            st.error("Tỉnh/thành này chưa có danh mục phường/xã/đặc khu trong dữ liệu hiện tại.")

    st.caption(
        "Danh mục địa giới được tải theo cấu trúc 2 cấp hiện hành; "
        "không nhập tay tên phường/xã để tránh sai địa danh."
    )

if selected_province and commune_value:
    st.success(f"Đã chọn: **{commune_value}, {selected_province}**", icon=":material/location_on:")

# C. Đồng ý + lưu hồ sơ — form riêng
with st.form("patient_consent_save_form_v22", clear_on_submit=False):
    with st.container(border=True):
        st.markdown("### :material/verified_user: Đồng ý tham gia sàng lọc và nghiên cứu")

        if operator_mode:
            st.info(
                "Bạn đang nhập giúp người tham gia. Chỉ tiếp tục khi người tham gia đã được giải thích "
                "nội dung và đã đồng ý cho sử dụng dữ liệu theo mục đích nghiên cứu/sàng lọc.",
                icon=":material/info:",
            )
        else:
            st.info(
                "Để tiếp tục, người tham gia cần đọc và đồng ý với nội dung dưới đây. "
                "Nếu không đồng ý, hệ thống sẽ không thực hiện sàng lọc và không lưu hồ sơ.",
                icon=":material/info:",
            )

        st.markdown(
            "**Tôi đồng ý cho nghiên cứu sinh sử dụng thông tin cá nhân, thông tin khảo sát và "
            "thông tin sức khỏe/xét nghiệm do tôi cung cấp cho mục đích sàng lọc cộng đồng "
            "Thalassemia và nghiên cứu khoa học.**"
        )
        st.markdown(
            "Tôi hiểu rằng việc tham gia là tự nguyện; kết quả của hệ thống chỉ có tính chất sàng lọc, "
            "không thay thế chẩn đoán của cơ sở y tế; dữ liệu được lưu phục vụ mục đích nêu trên theo "
            "phiên bản chấp thuận của nghiên cứu."
        )

        research_consent = st.checkbox(
            "Tôi đã đọc, hiểu và đồng ý tham gia." if not operator_mode
            else "Tôi xác nhận người tham gia đã đọc/được giải thích và đã đồng ý.",
            key="v22_research_consent",
        )
        if operator_mode:
            st.caption(f"Người nhập: {auth_user['full_name']} ({auth_user['username']})")
        st.caption(f"Phiên bản nội dung chấp thuận: {CONSENT_VERSION}")

        save_profile = st.form_submit_button(
            "Lưu / cập nhật hồ sơ", icon=":material/save:", type="primary", use_container_width=True
        )

phone = normalize_phone(phone_raw)

if save_profile:
    if not research_consent:
        st.error("Bạn cần đồng ý tham gia sàng lọc và nghiên cứu trước khi tiếp tục.")
    elif not admin_data_ok:
        st.error("Chưa tải được danh mục địa giới hiện hành nên chưa thể lưu hồ sơ.")
    elif not full_name.strip():
        st.error("Vui lòng nhập họ và tên.")
    elif not valid_vietnam_phone(phone):
        st.error("Số điện thoại phải là số Việt Nam 10 chữ số.")
    elif not current_address.strip():
        st.error("Vui lòng nhập địa chỉ hiện tại.")
    elif not selected_province:
        st.error("Vui lòng chọn tỉnh/thành phố.")
    elif not commune_value:
        st.error("Vui lòng chọn phường/xã/đặc khu.")
    else:
        profile = {
            "phone": phone,
            "full_name": full_name.strip(),
            "birth_date": birth_date.isoformat(),
            "gender": gender,
            "current_address": current_address.strip(),
            "province": selected_province,
            "commune": commune_value,
            "age": calculate_age(birth_date),
            "consent_at": datetime.now().isoformat(timespec="seconds"),
            "entry_mode": "assisted" if operator_mode else "self",
            "entered_by_username": operator_username,
        }
        action = upsert_patient(profile)
        st.session_state["patient_profile"] = profile
        st.session_state["research_consent"] = True
        reset_results()

        if action == "updated":
            st.success("Đã cập nhật hồ sơ bằng lần nhập sau cùng.", icon=":material/check_circle:")
        else:
            st.success("Đã tạo hồ sơ bệnh nhân.", icon=":material/check_circle:")

patient = st.session_state.get("patient_profile")

if not st.session_state.get("research_consent", False):
    st.warning(
        "Vui lòng đồng ý tham gia để được tham gia các vòng khảo sát sắp tới",
        icon=":material/lock:",
    )
    st.stop()

if not patient:
    st.info("Hãy lưu hồ sơ bệnh nhân trước khi bắt đầu Vòng 1.", icon=":material/arrow_upward:")
    st.stop()

st.success(
    f"**{patient['full_name']}** · {patient['age']} tuổi · {patient['phone']} · "
    f"{patient['commune']}, {patient['province']}",
    icon=":material/person:",
)

# ============================================================
# ROUND 1
# ============================================================

st.divider()
st.header(":material/quiz: Vòng 1 — 20 câu hỏi sàng lọc")
st.caption(
    "Vòng 1 dùng để ghi nhận các yếu tố cần lưu ý và đưa khuyến nghị. "
    "Tất cả người tham gia đều được tiếp tục Vòng 2."
)

YNK = ["Không", "Có", "Không biết"]

ROUND1_SECTIONS = [
    (":material/family_restroom: A. Tiền sử gia đình", [
        ("q1", "1. Trong gia đình/dòng họ có người từng được chẩn đoán Thalassemia không?"),
        ("q2", "2. Trong gia đình/dòng họ có người từng được thông báo mang gen Thalassemia/hemoglobinopathy không?"),
        ("q3", "3. Cha hoặc mẹ bạn có từng được xét nghiệm Thalassemia/hemoglobinopathy không?"),
        ("q4", "4. Anh/chị/em ruột có từng được chẩn đoán thiếu máu hoặc hồng cầu nhỏ không?"),
        ("q5", "5. Gia đình có trẻ từng phải truyền máu nhiều lần hoặc định kỳ không?"),
    ]),
    (":material/history: B. Tiền sử bản thân", [
        ("q6", "6. Bạn từng được nhân viên y tế thông báo bị thiếu máu chưa?"),
        ("q7", "7. Bạn từng được thông báo MCV thấp/hồng cầu nhỏ chưa?"),
        ("q8", "8. Bạn từng được thông báo MCH thấp/hồng cầu nhược sắc chưa?"),
        ("q9", None),  # selectbox riêng
        ("q10", "10. Bạn từng được chẩn đoán HbE hoặc hemoglobinopathy khác chưa?"),
        ("q11", "11. Bản thân từng truyền máu nhiều lần hoặc định kỳ chưa?"),
        ("q12", "12. Bạn có tiền sử thiếu máu kéo dài từ nhỏ hoặc từ tuổi thiếu niên không?"),
    ]),
    (":material/monitor_heart: C. Dấu hiệu hỗ trợ", [
        ("q13", "13. Bạn có thường xuyên mệt mỏi hoặc giảm khả năng hoạt động không?"),
        ("q14", "14. Bạn có thường xuyên hoa mắt/chóng mặt không rõ nguyên nhân không?"),
        ("q15", "15. Bạn từng được nhận xét da hoặc niêm mạc nhợt hơn bình thường chưa?"),
        ("q16", "16. Bạn từng có vàng da/vàng mắt không rõ nguyên nhân chưa?"),
        ("q17", "17. Bạn từng được bác sĩ ghi nhận lách to hoặc gan lách to chưa?"),
        ("q18", "18. Bạn từng được bác sĩ lưu ý có biến chứng liên quan bệnh huyết học mạn chưa?"),
    ]),
    (":material/travel_explore: D. Khả năng tiếp cận xét nghiệm — không tính điểm", [
        ("q19", "19. Bạn hiện có CBC trong vòng 6–12 tháng gần đây không?"),
        ("q20", "20. Bạn có gặp khó khăn khi đến cơ sở có xét nghiệm chuyên sâu do khoảng cách, chi phí hoặc thời gian di chuyển không?"),
    ]),
]

Q9_OPTIONS = [
    "Chưa xét nghiệm", "Đã xét nghiệm, bình thường", "Đã nghi ngờ",
    "Đã xác định mang gen", "Không nhớ",
]

with st.form("round1_form", clear_on_submit=False):
    answers = {}
    for section_title, questions in ROUND1_SECTIONS:
        with st.container(border=True):
            st.subheader(section_title)
            for qkey, qtext in questions:
                if qkey == "q9":
                    answers[qkey] = st.selectbox(
                        "9. Bạn từng xét nghiệm Thalassemia/hemoglobinopathy chưa?",
                        Q9_OPTIONS, key="r1w_q9",
                    )
                else:
                    answers[qkey] = st.radio(qtext, YNK, horizontal=True, key=f"r1w_{qkey}")
            if section_title.startswith(":material/travel_explore:"):
                st.info(
                    "Q19–Q20 được lưu để hỗ trợ điều hướng y tế, **không ảnh hưởng đến điểm nguy cơ Thalassemia**.",
                    icon=":material/info:",
                )

    round1_submit = st.form_submit_button(
        "Đánh giá Vòng 1", icon=":material/play_arrow:", type="primary", use_container_width=True
    )

if round1_submit:
    score1, reasons1 = calculate_round1_score(answers)
    category1, conclusion1 = round1_category(score1)

    # Vòng 1 mới → xóa kết quả Vòng 2/3 cũ trong phiên.
    for key in list(st.session_state.keys()):
        if key.startswith(("round2_", "round3_", "google_")):
            del st.session_state[key]

    st.session_state["round1_score"] = score1
    st.session_state["round1_reasons"] = reasons1
    st.session_state["round1_category"] = category1
    st.session_state["round1_conclusion"] = conclusion1
    st.session_state["round1_completed"] = True
    st.session_state["screening_id"] = create_screening_record(
        patient=patient, answers=answers, score1=score1, category1=category1,
        conclusion1=conclusion1, reasons1=reasons1,
        entry_mode="assisted" if operator_mode else "self",
        entered_by_username=operator_username,
    )

# ------------------------------------------------------------
# ROUND 1 RESULT
# ------------------------------------------------------------

if st.session_state.get("round1_completed", False):
    st.subheader(":material/fact_check: Kết quả Vòng 1")

    score1 = st.session_state["round1_score"]
    category1 = st.session_state["round1_category"]
    conclusion1 = st.session_state["round1_conclusion"]
    reasons1 = st.session_state["round1_reasons"]

    a1, a2 = st.columns(2)
    a1.metric("Điểm sàng lọc ban đầu", f"{score1}/{ROUND1_MAX_SCORE}")
    a2.metric("Vòng 2", "Được tiếp tục")

    risk_alert("Mức độ đáng lưu ý", category1, conclusion1)

    if reasons1:
        with st.expander("Những điểm Vòng 1 cần lưu ý", expanded=True, icon=":material/search:"):
            bullet_list(reasons1)
    else:
        st.info("Chưa ghi nhận yếu tố đáng lưu ý từ các câu trả lời.")

    st.markdown("### :material/route: Sau Vòng 1, bạn nên làm gì?")
    for i, item in enumerate(round1_guidance(category1), 1):
        st.markdown(f"{i}. {item}")

    st.info(
        "**Vòng 2 được mở cho tất cả người tham gia.** Nếu bạn đã có CBC, hãy nhập kết quả để hệ thống "
        "phân tích các chỉ số huyết học. Nếu chưa có, bạn có thể thực hiện CBC tại cơ sở y tế phù hợp "
        "rồi quay lại nhập kết quả.",
        icon=":material/lock_open:",
    )
    st.session_state["round2_unlocked"] = True

# ============================================================
# ROUND 2
# ============================================================

if st.session_state.get("round2_unlocked", False):

    st.divider()
    st.header(":material/biotech: Vòng 2 — Độ cao và công thức máu")

    with st.form("round2_form_v23", clear_on_submit=False):
        with st.container(border=True):
            st.subheader(":material/terrain: 1. Nơi cư trú và độ cao")
            st.write(f"**Tỉnh/thành:** {patient['province']}")
            st.write(f"**Phường/xã/đặc khu:** {patient['commune']}")
            st.write(f"**Địa chỉ hiện tại:** {patient['current_address']}")
            altitude_choice, altitude_adjustment = altitude_selector("round2")

        with st.container(border=True):
            st.subheader(":material/water_drop: 2. Nhập công thức máu")
            c1, c2, c3, c4, c5 = st.columns(5)
            with c1:
                hb_unit = st.selectbox("Đơn vị Hb", ["g/dL", "g/L"], key="round2_hb_unit")
                hb_raw = st.number_input("Hb", min_value=0.0, max_value=100000.0, value=None,
                                         step=0.1, key="round2_hb_raw")
            with c2:
                mcv = st.number_input("MCV (fL)", 30.0, 150.0, None, 0.1, key="round2_mcv")
            with c3:
                mch = st.number_input("MCH (pg)", 10.0, 50.0, None, 0.1, key="round2_mch")
            with c4:
                rbc_unit = st.selectbox("Đơn vị RBC", ["T/L", "10^12/L", "10^6/µL"], key="round2_rbc_unit")
                rbc_raw = st.number_input("RBC", 1.0, 10.0, None, 0.1, key="round2_rbc_raw")
            with c5:
                rdw = st.number_input("RDW-CV (%)", 5.0, 40.0, None, 0.1, key="round2_rdw")

            st.caption("Hệ thống tự chuyển Hb → g/dL và RBC → T/L trước khi tính.")
            st.caption(
                "Tiêu chí sàng lọc CBC theo Phụ lục ‘Quy trình xét nghiệm sàng lọc, chẩn đoán Thalassemia’ "
                "của Hướng dẫn chẩn đoán và điều trị một số bệnh lý huyết học (Bộ Y tế, 2022): "
                "MCV <85 fL và/hoặc MCH <28 pg. Đây là tiêu chí sàng lọc, không phải tiêu chuẩn chẩn đoán xác định."
            )

        round2_submit = st.form_submit_button(
            "Phân tích Vòng 2", icon=":material/biotech:", type="primary", use_container_width=True
        )

    if round2_submit:
        required_cbc = {"Hb": hb_raw, "MCV": mcv, "MCH": mch, "RBC": rbc_raw, "RDW-CV": rdw}
        missing = [name for name, value in required_cbc.items() if value is None]

        if missing:
            st.error("Vui lòng nhập đầy đủ số liệu CBC trước khi phân tích: " + ", ".join(missing) + ".")
        elif rbc_raw <= 0:
            st.error("RBC phải lớn hơn 0.")
        else:
            hb = hb_to_g_dl(hb_raw, hb_unit)
            rbc = rbc_to_t_l(rbc_raw, rbc_unit)
            hb_adjusted = hb - altitude_adjustment

            score2, mentzer, reasons2 = calculate_round2_score(mcv, mch, rbc, rdw)
            category2, conclusion2 = round2_category(score2, mcv, mch)
            findings, advice = narrative_cbc_advice(hb_adjusted, mcv, mch, rbc, rdw, mentzer)

            st.session_state.update({
                "round2_altitude_choice": altitude_choice,
                "round2_adjustment": altitude_adjustment,
                "round2_score": score2,
                "round2_mentzer": mentzer,
                "round2_category": category2,
                "round2_conclusion": conclusion2,
                "round2_reasons": reasons2,
                "round2_hb": hb,
                "round2_hb_adjusted": hb_adjusted,
                "round2_mcv_result": mcv,
                "round2_mch_result": mch,
                "round2_rbc_result": rbc,
                "round2_rdw_result": rdw,
                "round2_findings": findings,
                "round2_advice": advice,
            })

            update_screening_round2(
                st.session_state.get("screening_id"),
                {
                    "altitude_choice": altitude_choice, "adjustment": altitude_adjustment,
                    "hb": hb, "hb_adjusted": hb_adjusted, "mcv": mcv, "mch": mch,
                    "rbc": rbc, "rdw": rdw, "mentzer": mentzer, "score": score2,
                    "category": category2, "conclusion": conclusion2, "reasons": reasons2,
                    "findings": findings, "advice": advice,
                },
            )
            st.session_state["round2_completed"] = True

    # --------------------------------------------------------
    # ROUND 2 RESULTS
    # --------------------------------------------------------

    if st.session_state.get("round2_completed", False):

        st.subheader(":material/monitoring: Kết quả Vòng 2")

        a, b, c, d = st.columns(4)
        a.metric("Mentzer Index", f"{st.session_state['round2_mentzer']:.2f}")
        b.metric("Điểm CBC", str(st.session_state["round2_score"]))
        c.metric("Hb thực đo", f"{st.session_state['round2_hb']:.1f} g/dL")
        d.metric("Hb sau hiệu chỉnh", f"{st.session_state['round2_hb_adjusted']:.1f} g/dL")

        category2 = st.session_state["round2_category"]
        conclusion2 = st.session_state["round2_conclusion"]
        risk_alert("Nguy cơ Vòng 2", category2, conclusion2)

        with st.expander("Các yếu tố từ CBC", expanded=True, icon=":material/search:"):
            bullet_list(st.session_state["round2_reasons"])

        st.markdown("### :material/analytics: Phân tích sơ bộ")
        bullet_list(st.session_state["round2_findings"])

        st.markdown("### :material/clinical_notes: Khuyến nghị cho người được sàng lọc")
        bullet_list(st.session_state["round2_advice"])

        st.caption(
            "Những nhận định trên chỉ hỗ trợ sàng lọc và phải được đối chiếu với lâm sàng, "
            "xét nghiệm chuyên sâu và nhân viên y tế."
        )

        # ----------------------------------------------------
        # ROUND 3
        # ----------------------------------------------------

        st.subheader(":material/lab_panel: 3. Theo dõi sau sàng lọc — kết quả xét nghiệm chuyên sâu")
        st.caption(
            "Nếu người tham gia đã thực hiện điện di huyết sắc tố, HPLC, xét nghiệm sắt hoặc xét nghiệm "
            "di truyền, có thể nhập bổ sung tại đây để phục vụ theo dõi và tư vấn về sau. Vòng 3 không "
            "chấm điểm và không tự xác lập chẩn đoán."
        )

        with st.form("round3_followup_form", clear_on_submit=False):
            c1, c2, c3 = st.columns(3)
            with c1:
                r3_date = st.date_input("Ngày thực hiện xét nghiệm", value=None, key="round3_test_date")
            with c2:
                r3_type = st.selectbox(
                    "Loại xét nghiệm",
                    ["Điện di huyết sắc tố", "HPLC huyết sắc tố", "Xét nghiệm di truyền",
                     "Ferritin / đánh giá sắt", "Kết hợp nhiều xét nghiệm", "Khác"],
                    key="round3_test_type",
                )
            with c3:
                r3_facility = st.text_input("Cơ sở thực hiện", key="round3_facility",
                                            placeholder="Bệnh viện / phòng xét nghiệm")

            st.markdown("**Kết quả huyết sắc tố (nếu có trên phiếu xét nghiệm)**")
            h1, h2, h3, h4 = st.columns(4)
            with h1:
                r3_hba = st.number_input("HbA (%)", min_value=0.0, max_value=100.0, value=None, step=0.1, key="round3_hba")
            with h2:
                r3_hba2 = st.number_input("HbA2 (%)", min_value=0.0, max_value=30.0, value=None, step=0.1, key="round3_hba2")
            with h3:
                r3_hbf = st.number_input("HbF (%)", min_value=0.0, max_value=100.0, value=None, step=0.1, key="round3_hbf")
            with h4:
                r3_hbe = st.number_input("HbE (%)", min_value=0.0, max_value=100.0, value=None, step=0.1, key="round3_hbe")

            st.markdown("**Xét nghiệm liên quan khác (nếu có)**")
            f1, f2, f3 = st.columns(3)
            with f1:
                r3_ferritin = st.number_input("Ferritin", min_value=0.0, max_value=10000.0, value=None, step=0.1, key="round3_ferritin")
            with f2:
                r3_iron = st.number_input("Sắt huyết thanh", min_value=0.0, max_value=1000.0, value=None, step=0.1, key="round3_iron")
            with f3:
                r3_tsat = st.number_input("Độ bão hòa transferrin (%)", min_value=0.0, max_value=100.0, value=None, step=0.1, key="round3_tsat")

            r3_genetic = st.text_area("Kết quả xét nghiệm di truyền (nếu có)", key="round3_genetic",
                                      placeholder="Ghi theo phiếu xét nghiệm hoặc tóm tắt chính xác kết quả...")
            r3_lab = st.text_area("Kết luận trên phiếu xét nghiệm / nhận xét của cơ sở thực hiện", key="round3_lab",
                                  placeholder="Ưu tiên ghi đúng nội dung của cơ sở xét nghiệm, không tự suy diễn.")

            s1, s2 = st.columns(2)
            with s1:
                r3_status = st.selectbox(
                    "Trạng thái theo dõi",
                    ["Đã có kết quả", "Đã thực hiện — đang chờ kết quả", "Chưa thực hiện"],
                    key="round3_status",
                )
            with s2:
                r3_follow_date = st.date_input("Ngày hẹn theo dõi tiếp theo (nếu có)", value=None, key="round3_follow_date")

            r3_note = st.text_area(
                "Ghi chú tư vấn / theo dõi", key="round3_note",
                placeholder="Ví dụ: mang kết quả đến bác sĩ Huyết học; cân nhắc tư vấn di truyền khi phù hợp...",
            )

            save_r3 = st.form_submit_button(
                "Lưu kết quả theo dõi", icon=":material/save:", type="primary", use_container_width=True
            )

        if save_r3:
            has_result = (
                any(v is not None for v in [r3_hba, r3_hba2, r3_hbf, r3_hbe, r3_ferritin, r3_iron, r3_tsat])
                or bool(r3_genetic.strip()) or bool(r3_lab.strip())
            )
            if r3_status == "Đã có kết quả" and not has_result:
                st.warning("Bạn chọn 'Đã có kết quả' nhưng chưa nhập dữ liệu. Vui lòng kiểm tra lại.")
            else:
                update_screening_round3(
                    st.session_state.get("screening_id"),
                    {
                        "test_date": r3_date.isoformat() if r3_date else None,
                        "facility": r3_facility.strip() or None, "test_type": r3_type,
                        "hba": r3_hba, "hba2": r3_hba2, "hbf": r3_hbf, "hbe": r3_hbe,
                        "ferritin": r3_ferritin, "serum_iron": r3_iron, "transferrin_saturation": r3_tsat,
                        "genetic_result": r3_genetic.strip() or None,
                        "lab_conclusion": r3_lab.strip() or None,
                        "followup_status": r3_status,
                        "counseling_note": r3_note.strip() or None,
                        "followup_date": r3_follow_date.isoformat() if r3_follow_date else None,
                    },
                )
                st.session_state["round3_completed"] = True
                st.success("Đã lưu kết quả theo dõi sau sàng lọc.", icon=":material/check_circle:")

        if st.session_state.get("round3_completed", False):
            st.info(
                "Hồ sơ đã có dữ liệu Vòng 3. Khi có kết quả mới, bạn có thể cập nhật lại để duy trì hồ sơ theo dõi hiện hành.",
                icon=":material/info:",
            )
            st.markdown("### :material/rule: Đối chiếu kết quả theo hướng dẫn BYT 2022")
            current_r3 = {
                "hba": st.session_state.get("round3_hba"), "hba2": st.session_state.get("round3_hba2"),
                "hbf": st.session_state.get("round3_hbf"), "hbe": st.session_state.get("round3_hbe"),
                "ferritin": st.session_state.get("round3_ferritin"),
                "transferrin_saturation": st.session_state.get("round3_tsat"),
                "genetic_result": st.session_state.get("round3_genetic", ""),
                "lab_conclusion": st.session_state.get("round3_lab", ""),
            }
            bullet_list(round3_byt_interpretation(current_r3))
            st.caption(
                "Đây là đối chiếu hỗ trợ dựa trên dữ liệu đã nhập và nội dung tài liệu BYT 2022; "
                "không thay thế kết luận của cơ sở xét nghiệm/bác sĩ Huyết học."
            )

        # ----------------------------------------------------
        # MEDICAL FACILITIES
        # ----------------------------------------------------

        st.subheader(":material/local_hospital: Danh sách cơ sở y tế")
        st.caption(
            "Tra cứu bệnh viện/cơ sở y tế theo tỉnh, tên cơ sở hoặc tuyến. Danh mục là dữ liệu điều hướng "
            "của prototype và cần được đối soát với bệnh viện trước khi sử dụng thực tế."
        )

        all_facilities = all_medical_facilities()
        province_options = ["Tất cả tỉnh/thành"] + sorted({x["province"] for x in all_facilities})
        default_province = PROVINCE_FACILITY_ALIASES.get(patient["province"], patient["province"])
        selected_facility_province = st.selectbox(
            "Tỉnh/thành", province_options,
            index=province_options.index(default_province) if default_province in province_options else 0,
            key="facility_province_filter",
        )
        facility_keyword = st.text_input(
            "Tìm bệnh viện", placeholder="Ví dụ: Bạch Mai, Huyết học, Đa khoa...",
            key="facility_keyword", icon=":material/search:",
        )

        filtered = all_facilities
        if selected_facility_province != "Tất cả tỉnh/thành":
            filtered = [x for x in filtered if x["province"] == selected_facility_province]
        if facility_keyword.strip():
            kw = facility_keyword.strip().lower()
            filtered = [
                x for x in filtered
                if kw in " ".join([x["name"], x["tier"], x["note"], x["province"]]).lower()
            ]

        st.write(f"**{len(filtered)} cơ sở** phù hợp với bộ lọc.")
        if filtered:
            for i, facility in enumerate(filtered, start=1):
                with st.container(border=True):
                    st.markdown(f"**{i}. {facility['name']}**")
                    st.caption(f"{facility['province']} · {facility['tier']}")
                    st.write(facility["note"])
                    st.link_button("Xem vị trí / chỉ đường", facility["maps"], icon=":material/map:")
        else:
            st.info("Chưa tìm thấy cơ sở phù hợp. Hãy thử bỏ bộ lọc hoặc tìm bằng tên bệnh viện khác.")

        st.divider()
        st.markdown("### :material/star: Cơ sở ưu tiên theo nơi ở của người bệnh")

        facilities = recommended_facilities(patient["province"])
        if facilities:
            st.success(f"Đã tìm thấy {len(facilities)} cơ sở ưu tiên cho **{patient['province']}**.",
                       icon=":material/check_circle:")
            for i, facility in enumerate(facilities, start=1):
                with st.container(border=True):
                    st.markdown(f"**{i}. {facility['name']}**")
                    st.caption(f"{facility['tier']} · {facility['note']}")
                    st.link_button("Xem vị trí / chỉ đường", facility["maps"], icon=":material/map:")
        else:
            maps_url = ("https://www.google.com/maps/search/?api=1&query="
                        + quote(f"bệnh viện {patient['province']}"))
            st.info(
                f"Chưa có cơ sở được chọn sẵn cho **{patient['province']}**. "
                "Bạn có thể mở Google Maps để xem các bệnh viện gần khu vực."
            )
            st.link_button("Tìm bệnh viện trên Google Maps", maps_url, icon=":material/map:")

        st.caption(
            "Danh mục chỉ hỗ trợ định hướng. Khả năng thực hiện CBC, HPLC/điện di Hb, xét nghiệm gen "
            "Thalassemia và tiếp nhận chuyên khoa có thể khác nhau theo cơ sở; nên gọi xác nhận trước khi đến."
        )

        # ----------------------------------------------------
        # WORD REPORT
        # ----------------------------------------------------

        st.subheader(":material/description: Phiếu kết quả")

        report = make_word(
            patient=patient,
            r1_score=st.session_state["round1_score"],
            r1_category=st.session_state["round1_category"],
            r1_reasons=st.session_state["round1_reasons"],
            r2={
                "altitude_choice": st.session_state["round2_altitude_choice"],
                "adjustment": st.session_state["round2_adjustment"],
                "hb": st.session_state["round2_hb"],
                "hb_adjusted": st.session_state["round2_hb_adjusted"],
                "mcv": st.session_state["round2_mcv_result"],
                "mch": st.session_state["round2_mch_result"],
                "rbc": st.session_state["round2_rbc_result"],
                "rdw": st.session_state["round2_rdw_result"],
                "mentzer": st.session_state["round2_mentzer"],
                "score": st.session_state["round2_score"],
                "category": st.session_state["round2_category"],
                "conclusion": st.session_state["round2_conclusion"],
                "findings": st.session_state["round2_findings"],
                "advice": st.session_state["round2_advice"],
            },
        )

        st.download_button(
            "Tải phiếu Word",
            data=report,
            file_name=f"Phieu_Thalassemia_{safe_filename(patient['full_name'])}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            icon=":material/download:",
            use_container_width=True,
        )

# ============================================================
# FOOTER
# ============================================================

st.divider()
st.caption(
    "V5 Prototype — các trọng số/ngưỡng cần validation trên dữ liệu người Việt Nam "
    "trước khi sử dụng trong nghiên cứu lâm sàng."
)
