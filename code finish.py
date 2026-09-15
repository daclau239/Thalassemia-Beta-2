
import io
import json
import math
import os
import re
import hashlib
import secrets
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

import requests
import streamlit as st

# ===== PROFESSIONAL HEMATOLOGY UI =====
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

/* Dark-mode refinement. Streamlit exposes theme colors through CSS variables;
   this media rule only adjusts the hematology accent for contrast. */
@media (prefers-color-scheme: dark) {
  :root { --heme-red:#C45A69; --heme-accent:#E07987; }
  [data-testid="stExpander"], .contact-card {
    box-shadow:0 1px 0 rgba(255,255,255,.025) inset;
  }
}
</style>
""", unsafe_allow_html=True)

from docx import Document

# ============================================================
# THALASSEMIA SCREENING V8
# ============================================================
# 1) Hồ sơ bệnh nhân
# 2) Vòng 1: 20 câu hỏi
# 3) Vòng 1 phân tầng và đưa khuyến nghị; tất cả người tham gia đều được vào Vòng 2
# 4) Vòng 2:
#      - chọn tỉnh + phường/xã/đặc khu
#      - chọn khoảng độ cao
#      - CBC + đơn vị
#      - chuẩn hóa đơn vị
#      - Hb hiệu chỉnh theo độ cao (WHO 2024 prototype)
#      - Mentzer + phân tích sơ bộ
#      - lời khuyên
#      - gợi ý cơ sở y tế qua Google Places nếu có API key
#
# IMPORTANT:
# - Risk score hiện tại là prototype, chưa validation trên người Việt Nam.
# - "Khuyến nghị" là hỗ trợ sàng lọc, không chẩn đoán.
# - Q19/Q20 về tiếp cận xét nghiệm = 0 điểm.
# - Một số điện thoại = một hồ sơ hiện tại; mỗi lần sàng lọc được lưu thành một bản ghi lịch sử riêng.
# - SQLite chỉ phù hợp prototype; Streamlit Cloud có thể reset filesystem.
# ============================================================


# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------

st.set_page_config(
    page_title="Hệ thống Sàng lọc Thalassemia",
    page_icon="🩸",
    layout="wide",
)

# Persistent database for Streamlit Cloud.
# Configure one of these secrets in Streamlit Cloud:
#   # ============================================================
# PERSISTENT DATABASE CONFIG
# Prefer separate secrets instead of a URI so passwords containing
# @, #, ?, &, etc. never break URL parsing.
#
# Streamlit Secrets:
# DB_HOST = "aws-0-ap-northeast-2.pooler.supabase.com"
# DB_PORT = 6543
# DB_NAME = "postgres"
# DB_USER = "postgres.cqltdtigenzqoqrfzgej"
# DB_PASSWORD = "YOUR_DATABASE_PASSWORD"
#
# DATABASE_URL is still supported as a fallback.
# ============================================================

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

LOCAL_SQLITE_PATH = "thalassemia_patients.db"
CONSENT_VERSION = "DACLAU239-BETA5"
ADMIN_DATA_URL = "https://raw.githubusercontent.com/open-admin-data/vietnam-administrative-divisions/main/data/hierarchy.json"
ADMIN_DATA_SOURCE_URL = "https://github.com/open-admin-data/vietnam-administrative-divisions"

ROUND1_MAX_SCORE = 20
ROUND1_HIGH_THRESHOLD = 8  # chỉ dùng để phân tầng Vòng 1, KHÔNG khóa Vòng 2

FOLLOWUP_DAYS = 30

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


# ------------------------------------------------------------
# GOOGLE KEY
# ------------------------------------------------------------

def get_google_key():
    """Read Google API key only from Streamlit secrets."""
    try:
        value = st.secrets["GOOGLE_API_KEY"]
    except Exception:
        return ""
    return str(value).strip() if value else ""

GOOGLE_API_KEY = get_google_key()


# ------------------------------------------------------------
# DATABASE
# ------------------------------------------------------------

PBKDF2_ITERATIONS = 310_000

DB_SCHEMA_READY_V21 = "db_schema_ready_v21"


# Tài khoản quản trị mặc định theo yêu cầu của chủ hệ thống.
DEFAULT_ADMIN_USERNAME = "daclau239"
DEFAULT_ADMIN_PASSWORD = "23092002"
DEFAULT_ADMIN_FULL_NAME = "Quản trị viên hệ thống"
DEFAULT_ADMIN_EMAIL = "admin@thalassemia.local"



class PersistentPGConnection:
    """Compatibility wrapper around psycopg using explicit host/port/user/password."""
    def __init__(self, dsn=None):
        try:
            import psycopg
        except ModuleNotFoundError as exc:
            raise RuntimeError(
                "Thiếu thư viện psycopg. Hãy thêm `psycopg[binary]` vào requirements.txt."
            ) from exc

        kwargs = None

        if DB_HOST and DB_USER and DB_PASSWORD:
            kwargs = {
                "host": DB_HOST,
                "port": DB_PORT,
                "dbname": DB_NAME or "postgres",
                "user": DB_USER,
                "password": DB_PASSWORD,
            }
        elif dsn:
            kwargs = {"conninfo": dsn}

        if not kwargs:
            raise RuntimeError(
                "Chưa cấu hình database. Cần DB_HOST, DB_PORT, DB_NAME, DB_USER, DB_PASSWORD "
                "trong Streamlit Secrets (khuyến nghị), hoặc một DATABASE_URL hợp lệ."
            )

        try:
            self._conn = psycopg.connect(**kwargs)
        except Exception as exc:
            host = DB_HOST or "<DATABASE_URL>"
            raise RuntimeError(
                f"Không thể kết nối PostgreSQL tới host '{host}'. "
                "Kiểm tra DB_HOST, DB_PORT, DB_USER, DB_PASSWORD và database Supabase."
            ) from exc

        self._conn.autocommit = False

    @staticmethod
    def _translate(sql: str) -> str:
        return sql.replace("?", "%s")

    def execute(self, sql: str, params=None):
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
        """
        SELECT column_name
        FROM information_schema.columns
        WHERE table_schema = 'public' AND table_name = ?
        """,
        (table_name,),
    ).fetchall()
    return {row[0] for row in rows}


def _create_postgres_schema(conn):
    conn.execute(
        """
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
        )
        """
    )
    conn.execute(
        """
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
        )
        """
    )
    conn.execute(
        """
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
        )
        """
    )
    conn.execute("""
        CREATE TABLE IF NOT EXISTS system_meta (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    """)

    # Safe migrations for columns added after the first prototype.
    patient_columns = _pg_table_columns(conn, "patient_profiles")
    patient_migrations = [
        ("research_consent", "INTEGER NOT NULL DEFAULT 0"),
        ("consent_version", "TEXT"),
        ("consent_at", "TEXT"),
    ]
    for column, definition in patient_migrations:
        if column not in patient_columns:
            conn.execute(f"ALTER TABLE patient_profiles ADD COLUMN {column} {definition}")

    record_columns = _pg_table_columns(conn, "screening_records")
    record_migrations = [
        ("round3_completed", "INTEGER NOT NULL DEFAULT 0"),
        ("round3_test_date", "TEXT"),
        ("round3_facility", "TEXT"),
        ("round3_test_type", "TEXT"),
        ("round3_hba", "REAL"),
        ("round3_hba2", "REAL"),
        ("round3_hbf", "REAL"),
        ("round3_hbe", "REAL"),
        ("round3_ferritin", "REAL"),
        ("round3_serum_iron", "REAL"),
        ("round3_transferrin_saturation", "REAL"),
        ("round3_genetic_result", "TEXT"),
        ("round3_lab_conclusion", "TEXT"),
        ("round3_followup_status", "TEXT"),
        ("round3_counseling_note", "TEXT"),
        ("round3_followup_date", "TEXT"),
    ]
    for column, definition in record_migrations:
        if column not in record_columns:
            conn.execute(f"ALTER TABLE screening_records ADD COLUMN {column} {definition}")


def _migrate_local_sqlite_once(conn):
    """Migrate an existing local SQLite prototype DB into the persistent DB once, when present."""
    if not os.path.exists(LOCAL_SQLITE_PATH):
        return
    marker = conn.execute(
        "SELECT value FROM system_meta WHERE key = ?",
        ("local_sqlite_migrated",),
    ).fetchone()
    if marker and str(marker[0]) == "1":
        return

    try:
        import sqlite3
        local = sqlite3.connect(LOCAL_SQLITE_PATH)
        local.row_factory = sqlite3.Row

        # Copy parent tables first.
        patient_rows = local.execute("SELECT * FROM patient_profiles").fetchall()
        for r in patient_rows:
            conn.execute(
                """
                INSERT INTO patient_profiles
                (phone, full_name, birth_date, gender, current_address, province, commune,
                 research_consent, consent_version, consent_at, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT (phone) DO UPDATE SET
                    full_name=EXCLUDED.full_name,
                    birth_date=EXCLUDED.birth_date,
                    gender=EXCLUDED.gender,
                    current_address=EXCLUDED.current_address,
                    province=EXCLUDED.province,
                    commune=EXCLUDED.commune,
                    research_consent=EXCLUDED.research_consent,
                    consent_version=EXCLUDED.consent_version,
                    consent_at=EXCLUDED.consent_at,
                    updated_at=EXCLUDED.updated_at
                """,
                (
                    r["phone"], r["full_name"], r["birth_date"], r["gender"],
                    r["current_address"], r["province"], r["commune"],
                    r["research_consent"] if "research_consent" in r.keys() else 0,
                    r["consent_version"] if "consent_version" in r.keys() else None,
                    r["consent_at"] if "consent_at" in r.keys() else None,
                    r["created_at"], r["updated_at"],
                ),
            )

        # Copy users.
        try:
            user_rows = local.execute("SELECT * FROM user_accounts").fetchall()
            for r in user_rows:
                cols = set(r.keys())
                conn.execute(
                    """
                    INSERT INTO user_accounts
                    (id, username, full_name, email, password_hash, password_salt,
                     role, status, created_at, approved_by, approved_at, last_login_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT (username) DO NOTHING
                    """,
                    (
                        r["id"], r["username"], r["full_name"], r["email"],
                        r["password_hash"], r["password_salt"], r["role"], r["status"],
                        r["created_at"], r["approved_by"] if "approved_by" in cols else None,
                        r["approved_at"] if "approved_at" in cols else None,
                        r["last_login_at"] if "last_login_at" in cols else None,
                    ),
                )
        except Exception:
            pass

        # Copy ALL screening history. Nothing is deleted during migration.
        screening_rows = local.execute("SELECT * FROM screening_records").fetchall()
        record_columns = _pg_table_columns(conn, "screening_records")
        allowed = [c for c in record_columns if c != "id"]
        # Use explicit compatible columns and NULL for absent legacy fields.
        for r in screening_rows:
            def rv(name):
                return r[name] if name in r.keys() else None

            conn.execute(
                """
                INSERT INTO screening_records
                (id, phone, screening_at, entered_by_username, entry_mode,
                 consent_version, consent_at, round1_score, round1_category,
                 round1_conclusion, round1_reasons, answers_json,
                 round2_completed, altitude_choice, altitude_adjustment, hb, hb_adjusted,
                 mcv, mch, rbc, rdw, mentzer, round2_score, round2_category,
                 round2_conclusion, round2_reasons, findings_json, advice_json,
                 round3_completed, round3_test_date, round3_facility, round3_test_type,
                 round3_hba, round3_hba2, round3_hbf, round3_hbe, round3_ferritin,
                 round3_serum_iron, round3_transferrin_saturation, round3_genetic_result,
                 round3_lab_conclusion, round3_followup_status, round3_counseling_note,
                 round3_followup_date)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ON CONFLICT (id) DO NOTHING
                """,
                (
                    r["id"], rv("phone"), rv("screening_at"), rv("entered_by_username"),
                    rv("entry_mode"), rv("consent_version"), rv("consent_at"),
                    rv("round1_score") or 0, rv("round1_category") or "THẤP",
                    rv("round1_conclusion"), rv("round1_reasons"), rv("answers_json"),
                    rv("round2_completed") or 0, rv("altitude_choice"), rv("altitude_adjustment"),
                    rv("hb"), rv("hb_adjusted"), rv("mcv"), rv("mch"), rv("rbc"), rv("rdw"),
                    rv("mentzer"), rv("round2_score"), rv("round2_category"), rv("round2_conclusion"),
                    rv("round2_reasons"), rv("findings_json"), rv("advice_json"),
                    rv("round3_completed") or 0, rv("round3_test_date"), rv("round3_facility"),
                    rv("round3_test_type"), rv("round3_hba"), rv("round3_hba2"), rv("round3_hbf"),
                    rv("round3_hbe"), rv("round3_ferritin"), rv("round3_serum_iron"),
                    rv("round3_transferrin_saturation"), rv("round3_genetic_result"),
                    rv("round3_lab_conclusion"), rv("round3_followup_status"),
                    rv("round3_counseling_note"), rv("round3_followup_date"),
                ),
            )

        local.close()
        conn.execute(
            "INSERT INTO system_meta(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=EXCLUDED.value",
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
    """Đồng bộ sequence sau khi migrate các ID cũ từ SQLite."""
    for table in ("user_accounts", "screening_records"):
        conn.execute(
            f"""
            SELECT setval(
                pg_get_serial_sequence('{table}', 'id'),
                COALESCE((SELECT MAX(id) FROM {table}), 1),
                TRUE
            )
            """
        )



def _ensure_schema_once_v21():
    if st.session_state.get(DB_SCHEMA_READY_V21):
        return
    conn = PersistentPGConnection(DATABASE_URL or None)
    try:
        _create_postgres_schema(conn)
        _migrate_local_sqlite_once(conn)
        _sync_postgres_sequences(conn)
        conn.commit()
        st.session_state[DB_SCHEMA_READY_V21] = True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def get_db():
    _ensure_schema_once_v21()
    return PersistentPGConnection(DATABASE_URL or None)


def get_secret_value(*paths, env_name=""):
    for path in paths:
        try:
            value = st.secrets
            for part in path.split("."):
                value = value[part]
            if value:
                return str(value).strip()
        except Exception:
            pass

    if env_name:
        return os.environ.get(env_name, "").strip()
    return ""


def hash_password(password, salt_hex=None):
    if salt_hex:
        salt = bytes.fromhex(salt_hex)
        salt_value = salt_hex
    else:
        salt = secrets.token_bytes(16)
        salt_value = salt.hex()

    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        PBKDF2_ITERATIONS,
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
    """Kiểm tra số điện thoại Việt Nam dạng cơ bản cho prototype."""
    return bool(re.fullmatch(r"0\d{9}", str(phone or "")))

def valid_email(email):
    return bool(
        re.fullmatch(
            r"[^@\s]+@[^@\s]+\.[^@\s]+",
            (email or "").strip(),
        )
    )


def valid_username(username):
    return bool(
        re.fullmatch(
            r"[A-Za-z0-9_.-]{4,32}",
            (username or "").strip(),
        )
    )


def ensure_default_admin():
    """Create/repair the default admin safely on Streamlit reruns/concurrent starts.

    Uses an atomic INSERT OR IGNORE to avoid UNIQUE/IntegrityError when two
    sessions initialize the database at nearly the same time. If the default
    email is already occupied by a non-admin account, a deterministic fallback
    email is used because username is the login identifier.
    """
    conn = get_db()
    try:
        # First, if the default admin username already exists, nothing to do.
        row = conn.execute(
            "SELECT id, role, status FROM user_accounts WHERE username = ? LIMIT 1",
            (DEFAULT_ADMIN_USERNAME,),
        ).fetchone()
        if row:
            # Repair a legacy/default account if it exists but is not admin.
            if str(row[1] or '').lower() != 'admin':
                now = datetime.now().isoformat(timespec="seconds")
                password_hash, password_salt = hash_password(DEFAULT_ADMIN_PASSWORD)
                conn.execute(
                    """
                    UPDATE user_accounts
                    SET full_name = ?, password_hash = ?, password_salt = ?,
                        role = 'admin', status = 'approved',
                        approved_by = ?, approved_at = ?
                    WHERE id = ?
                    """,
                    (DEFAULT_ADMIN_FULL_NAME, password_hash, password_salt,
                     DEFAULT_ADMIN_USERNAME, now, row[0]),
                )
                conn.commit()
            return

        # If any admin already exists, do not create a second one.
        admin_row = conn.execute(
            "SELECT id FROM user_accounts WHERE role = 'admin' LIMIT 1"
        ).fetchone()
        if admin_row:
            return

        now = datetime.now().isoformat(timespec="seconds")
        password_hash, password_salt = hash_password(DEFAULT_ADMIN_PASSWORD)

        # Avoid a crash if another account already uses the default email.
        email_row = conn.execute(
            "SELECT id FROM user_accounts WHERE lower(email) = lower(?) LIMIT 1",
            (DEFAULT_ADMIN_EMAIL,),
        ).fetchone()
        admin_email = (
            f"{DEFAULT_ADMIN_USERNAME}@thalassemia.local"
            if email_row else DEFAULT_ADMIN_EMAIL
        )

        conn.execute(
            """
            INSERT INTO user_accounts
            (username, full_name, email, password_hash, password_salt,
             role, status, created_at, approved_by, approved_at)
            VALUES (?, ?, ?, ?, ?, 'admin', 'approved', ?, ?, ?)
            ON CONFLICT (username) DO NOTHING
            """,
            (
                DEFAULT_ADMIN_USERNAME,
                DEFAULT_ADMIN_FULL_NAME,
                admin_email,
                password_hash,
                password_salt,
                now,
                DEFAULT_ADMIN_USERNAME,
                now,
            ),
        )
        conn.commit()
    except sqlite3.IntegrityError:
        # Another session/process may have created the admin concurrently.
        # Re-check instead of letting the whole app crash.
        try:
            admin_row = conn.execute(
                "SELECT id FROM user_accounts WHERE role = 'admin' LIMIT 1"
            ).fetchone()
            if admin_row:
                conn.commit()
                return
        except Exception:
            pass
        raise
    finally:
        conn.close()

def admin_exists():
    conn = get_db()
    row = conn.execute(
        "SELECT 1 FROM user_accounts WHERE role = 'admin' LIMIT 1"
    ).fetchone()
    conn.close()
    return row is not None


def create_admin_account(username, full_name, email, password):
    username = username.strip()
    email = email.strip().lower()
    now = datetime.now().isoformat(timespec="seconds")
    password_hash, password_salt = hash_password(password)

    conn = get_db()
    try:
        conn.execute(
            """
            INSERT INTO user_accounts
            (username, full_name, email, password_hash, password_salt,
             role, status, created_at, approved_by, approved_at)
            VALUES (?, ?, ?, ?, ?, 'admin', 'approved', ?, ?, ?)
            """,
            (
                username,
                full_name.strip(),
                email,
                password_hash,
                password_salt,
                now,
                username,
                now,
            ),
        )
        conn.commit()
        return True, ""
    except sqlite3.IntegrityError as exc:
        conn.rollback()
        return False, f"Tài khoản/email đã tồn tại hoặc dữ liệu không hợp lệ: {exc}"
    finally:
        conn.close()


def register_staff_account(username, full_name, email, password):
    username = username.strip()
    email = email.strip().lower()
    now = datetime.now().isoformat(timespec="seconds")
    password_hash, password_salt = hash_password(password)

    conn = get_db()
    try:
        conn.execute(
            """
            INSERT INTO user_accounts
            (username, full_name, email, password_hash, password_salt,
             role, status, created_at)
            VALUES (?, ?, ?, ?, ?, 'staff', 'pending', ?)
            """,
            (
                username,
                full_name.strip(),
                email,
                password_hash,
                password_salt,
                now,
            ),
        )
        conn.commit()
        return True, ""
    except sqlite3.IntegrityError:
        conn.rollback()
        return False, "Tên đăng nhập hoặc email đã được sử dụng."
    finally:
        conn.close()


def authenticate_user(login_value, password):
    login_value = (login_value or "").strip().lower()
    conn = get_db()
    row = conn.execute(
        """
        SELECT id, username, full_name, email, password_hash,
               password_salt, role, status
        FROM user_accounts
        WHERE lower(username) = ? OR lower(email) = ?
        LIMIT 1
        """,
        (login_value, login_value),
    ).fetchone()

    if not row:
        conn.close()
        return None, "Sai tên đăng nhập/email hoặc mật khẩu."

    if not verify_password(password, row[4], row[5]):
        conn.close()
        return None, "Sai tên đăng nhập/email hoặc mật khẩu."

    if row[7] != "approved":
        conn.close()
        if row[7] == "pending":
            return None, "Tài khoản đang chờ quản trị viên phê duyệt."
        if row[7] == "disabled":
            return None, "Tài khoản đã bị vô hiệu hóa."
        return None, "Tài khoản chưa được phép truy cập."

    now = datetime.now().isoformat(timespec="seconds")
    conn.execute(
        "UPDATE user_accounts SET last_login_at = ? WHERE id = ?",
        (now, row[0]),
    )
    conn.commit()
    conn.close()

    return {
        "id": row[0],
        "username": row[1],
        "full_name": row[2],
        "email": row[3],
        "role": row[6],
        "status": row[7],
    }, ""


def list_user_accounts():
    conn = get_db()
    rows = conn.execute(
        """
        SELECT id, username, full_name, email, role, status,
               created_at, approved_by, approved_at, last_login_at
        FROM user_accounts
        ORDER BY
            CASE role WHEN 'admin' THEN 0 ELSE 1 END,
            CASE status WHEN 'pending' THEN 0 ELSE 1 END,
            created_at DESC
        """
    ).fetchall()
    conn.close()
    return rows


def update_staff_status(user_id, status, approved_by):
    if status not in {"approved", "rejected", "disabled", "pending"}:
        raise ValueError("Trạng thái không hợp lệ.")

    conn = get_db()
    now = datetime.now().isoformat(timespec="seconds")
    if status == "approved":
        conn.execute(
            """
            UPDATE user_accounts
            SET status = ?, approved_by = ?, approved_at = ?
            WHERE id = ? AND role = 'staff'
            """,
            (status, approved_by, now, user_id),
        )
    else:
        conn.execute(
            """
            UPDATE user_accounts
            SET status = ?
            WHERE id = ? AND role = 'staff'
            """,
            (status, user_id),
        )
    conn.commit()
    conn.close()


def list_patient_profiles_for_staff():
    conn = get_db()
    rows = conn.execute(
        """
        SELECT phone, full_name, birth_date, gender,
               current_address, province, commune,
               research_consent, consent_version, consent_at,
               created_at, updated_at
        FROM patient_profiles
        WHERE research_consent = 1
        ORDER BY updated_at DESC
        """
    ).fetchall()
    conn.close()
    return rows


def phone_exists(phone):
    conn = get_db()

    row = conn.execute(
        """
        SELECT 1
        FROM patient_profiles
        WHERE phone = ?
        """,
        (phone,),
    ).fetchone()

    conn.close()
    return row is not None


def upsert_patient(profile):
    conn = get_db()

    now = datetime.now().isoformat(
        timespec="seconds"
    )

    exists = conn.execute(
        """
        SELECT 1
        FROM patient_profiles
        WHERE phone = ?
        """,
        (profile["phone"],),
    ).fetchone()

    if exists:
        conn.execute(
            """
            UPDATE patient_profiles
            SET
                full_name = ?,
                birth_date = ?,
                gender = ?,
                current_address = ?,
                province = ?,
                commune = ?,
                research_consent = ?,
                consent_version = ?,
                consent_at = ?,
                updated_at = ?
            WHERE phone = ?
            """,
            (
                profile["full_name"],
                profile["birth_date"],
                profile["gender"],
                profile["current_address"],
                profile["province"],
                profile["commune"],
                1,
                CONSENT_VERSION,
                profile["consent_at"],
                now,
                profile["phone"],
            ),
        )
        action = "updated"
    else:
        conn.execute(
            """
            INSERT INTO patient_profiles (
                phone,
                full_name,
                birth_date,
                gender,
                current_address,
                province,
                commune,
                research_consent,
                consent_version,
                consent_at,
                created_at,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                profile["phone"],
                profile["full_name"],
                profile["birth_date"],
                profile["gender"],
                profile["current_address"],
                profile["province"],
                profile["commune"],
                1,
                CONSENT_VERSION,
                profile["consent_at"],
                now,
                now,
            ),
        )
        action = "inserted"

    conn.commit()
    conn.close()

    return action


def create_screening_record(
    patient,
    answers,
    score1,
    category1,
    conclusion1,
    reasons1,
    entry_mode,
    entered_by_username=None,
):
    """Lưu một lượt sàng lọc sau khi hoàn thành Vòng 1."""
    now = datetime.now().isoformat(timespec="seconds")
    conn = get_db()
    cur = conn.execute(
        """
        INSERT INTO screening_records (
            phone, screening_at, entered_by_username, entry_mode,
            consent_version, consent_at, round1_score, round1_category,
            round1_conclusion, round1_reasons, answers_json
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        RETURNING id
        """,
        (
            patient["phone"],
            now,
            entered_by_username,
            entry_mode,
            CONSENT_VERSION,
            patient["consent_at"],
            int(score1),
            category1,
            conclusion1,
            json.dumps(reasons1, ensure_ascii=False),
            json.dumps(answers, ensure_ascii=False),
        ),
    )
    record_row = cur.fetchone()
    record_id = int(record_row[0])

    # KHÔNG xóa lượt cũ. Mọi lần sàng lọc được giữ để nghiên cứu theo dõi dọc.
    conn.commit()
    conn.close()
    return record_id


def update_screening_round2(record_id, r2):
    if not record_id:
        return

    conn = get_db()
    conn.execute(
        """
        UPDATE screening_records
        SET round2_completed = 1,
            altitude_choice = ?,
            altitude_adjustment = ?,
            hb = ?,
            hb_adjusted = ?,
            mcv = ?,
            mch = ?,
            rbc = ?,
            rdw = ?,
            mentzer = ?,
            round2_score = ?,
            round2_category = ?,
            round2_conclusion = ?,
            round2_reasons = ?,
            findings_json = ?,
            advice_json = ?
        WHERE id = ?
        """,
        (
            r2["altitude_choice"],
            r2["adjustment"],
            r2["hb"],
            r2["hb_adjusted"],
            r2["mcv"],
            r2["mch"],
            r2["rbc"],
            r2["rdw"],
            r2["mentzer"],
            r2["score"],
            r2["category"],
            r2["conclusion"],
            json.dumps(r2["reasons"], ensure_ascii=False),
            json.dumps(r2["findings"], ensure_ascii=False),
            json.dumps(r2["advice"], ensure_ascii=False),
            int(record_id),
        ),
    )
    conn.commit()
    conn.close()



def update_screening_round3(record_id, r3):
    """Lưu kết quả xét nghiệm chuyên sâu và thông tin theo dõi sau sàng lọc."""
    if not record_id:
        return
    conn = get_db()
    conn.execute(
        """UPDATE screening_records SET
            round3_completed = 1, round3_test_date = ?, round3_facility = ?,
            round3_test_type = ?, round3_hba = ?, round3_hba2 = ?, round3_hbf = ?,
            round3_hbe = ?, round3_ferritin = ?, round3_serum_iron = ?,
            round3_transferrin_saturation = ?, round3_genetic_result = ?,
            round3_lab_conclusion = ?, round3_followup_status = ?,
            round3_counseling_note = ?, round3_followup_date = ?
        WHERE id = ?""",
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
    conn.close()

def list_screening_records_for_staff():
    """Lấy TOÀN BỘ lịch sử sàng lọc của người tham gia đã đồng ý."""
    conn = get_db()
    rows = conn.execute(
        """
        SELECT
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
        FROM screening_records s
        JOIN patient_profiles p ON p.phone = s.phone
        WHERE p.research_consent = 1
        ORDER BY s.screening_at DESC, s.id DESC
        """
    ).fetchall()
    conn.close()
    return rows



def list_latest_screening_records_for_staff():
    """Lấy một lượt mới nhất/người để hiển thị nhanh trên giao diện; không xóa lịch sử."""
    conn = get_db()
    rows = conn.execute(
        """
        SELECT
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
        FROM (
            SELECT sr.*,
                   ROW_NUMBER() OVER (
                       PARTITION BY sr.phone
                       ORDER BY sr.screening_at DESC, sr.id DESC
                   ) AS rn
            FROM screening_records sr
        ) s
        JOIN patient_profiles p ON p.phone = s.phone
        WHERE s.rn = 1 AND p.research_consent = 1
        ORDER BY s.screening_at DESC, s.id DESC
        """
    ).fetchall()
    conn.close()
    return rows


def export_screening_xlsx(patient_rows, screening_rows):
    """Tạo một file Excel nhiều sheet, không lưu file tạm trên server."""
    try:
        import xlsxwriter
    except ModuleNotFoundError as exc:
        raise RuntimeError(
            "Thiếu thư viện XlsxWriter. Hãy thêm `XlsxWriter>=3.2` vào requirements.txt rồi redeploy ứng dụng."
        ) from exc

    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})

    header_fmt = workbook.add_format({
        "bold": True,
        "bg_color": "#D9EAF7",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "text_wrap": True,
    })
    cell_fmt = workbook.add_format({"border": 1, "valign": "top"})
    date_fmt = workbook.add_format({"border": 1, "num_format": "dd/mm/yyyy hh:mm"})

    # Sheet 1: hồ sơ hiện tại
    ws = workbook.add_worksheet("Ho_so_nguoi_tham_gia")
    patient_headers = [
        "Số điện thoại", "Họ và tên", "Ngày sinh", "Giới tính",
        "Địa chỉ hiện tại", "Tỉnh/thành", "Phường/xã/đặc khu",
        "Đồng ý nghiên cứu", "Phiên bản consent", "Thời điểm đồng ý",
        "Tạo lúc", "Cập nhật lúc",
    ]
    for c, h in enumerate(patient_headers):
        ws.write(0, c, h, header_fmt)
    for r, row in enumerate(patient_rows, start=1):
        values = [
            row[0], row[1], row[2], row[3], row[4], row[5], row[6],
            "Có" if row[7] else "Không", row[8] or "", row[9] or "", row[10], row[11],
        ]
        for c, value in enumerate(values):
            fmt = date_fmt if c in (9, 10, 11) and isinstance(value, datetime) else cell_fmt
            ws.write(r, c, value, fmt)

    ws.freeze_panes(1, 0)
    widths = [15, 24, 13, 12, 32, 20, 26, 16, 25, 21, 21, 21]
    for c, w in enumerate(widths):
        ws.set_column(c, c, w)
    ws.autofilter(0, 0, max(len(patient_rows), 1), len(patient_headers)-1)

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
        values = [
            row[0], row[1], row[2] or "",
            "Tự nhập" if row[3] == "self" else "Nhập giúp người tham gia",
            row[4], row[5], row[6], row[7], row[8], row[9], row[10],
            row[11], row[12], row[13], row[14],
            "Có" if row[15] else "Chưa", row[16] or "", row[17] if row[17] is not None else "",
            row[18] if row[18] is not None else "", row[19] if row[19] is not None else "",
            row[20] if row[20] is not None else "", row[21] if row[21] is not None else "",
            row[22] if row[22] is not None else "", row[23] if row[23] is not None else "",
            row[24] if row[24] is not None else "", row[25] if row[25] is not None else "",
            row[26] or "", row[27] or "", row[28] or "",
            "Có" if row[29] else "Chưa", row[30] or "", row[31] or "", row[32] or "",
            row[33] if row[33] is not None else "", row[34] if row[34] is not None else "",
            row[35] if row[35] is not None else "", row[36] if row[36] is not None else "",
            row[37] if row[37] is not None else "", row[38] if row[38] is not None else "",
            row[39] if row[39] is not None else "", row[40] or "", row[41] or "",
            row[42] or "", row[43] or "", row[44] or "",
        ]
        for c, value in enumerate(values):
            fmt = date_fmt if c in (1, 12) and isinstance(value, datetime) else cell_fmt
            ws2.write(r, c, value, fmt)

    ws2.freeze_panes(1, 0)
    for c in range(len(screening_headers)):
        ws2.set_column(c, c, 20)
    ws2.set_column(4, 4, 15)
    ws2.set_column(5, 5, 24)
    ws2.set_column(8, 8, 32)
    ws2.set_column(9, 10, 22)
    ws2.set_column(27, 28, 38)
    ws2.autofilter(0, 0, max(len(screening_rows), 1), len(screening_headers)-1)

    workbook.close()
    output.seek(0)
    return output


# ------------------------------------------------------------
# ACCESS CONTROL / ADMIN CONSOLE
# ------------------------------------------------------------

def _database_is_configured():
    return bool((DB_HOST and DB_USER and DB_PASSWORD) or DATABASE_URL)


if not _database_is_configured():
    st.error("Database chưa được cấu hình — hệ thống đang khóa ghi dữ liệu để tránh mất dữ liệu.")
    with st.expander("Cấu hình database bền vững", expanded=True):
        st.code(
            'DB_HOST = "aws-0-ap-northeast-2.pooler.supabase.com"\n'
            'DB_PORT = 6543\n'
            'DB_NAME = "postgres"\n'
            'DB_USER = "postgres.cqltdtigenzqoqrfzgej"\n'
            'DB_PASSWORD = "MẬT_KHẨU_DATABASE_CỦA_EM"',
            language="toml",
        )
        st.caption(
            "Đặt các giá trị này trong Streamlit Cloud → Manage app → Settings → Secrets. "
            "Không cần percent-encode mật khẩu khi dùng DB_PASSWORD riêng."
        )
    st.stop()

try:
    ensure_default_admin()
except Exception as exc:
    st.error("Không thể kết nối/khởi tạo database bền vững.")
    st.code(str(exc))
    st.info(
        "Kiểm tra lại DB_HOST/DB_PORT/DB_NAME/DB_USER/DB_PASSWORD trong Streamlit Secrets."
    )
    st.stop()


def current_auth_user():
    return st.session_state.get("auth_user")


def logout_user():
    st.session_state.pop("auth_user", None)


def render_auth_sidebar():
    user = current_auth_user()

    st.sidebar.header("🔐 TÀI KHOẢN & QUYỀN TRUY CẬP")

    if user:
        role_label = "Quản trị viên" if user["role"] == "admin" else "Nhân sự được duyệt"
        st.sidebar.success(
            f"Đang đăng nhập: **{user['full_name']}**\n\n{role_label}"
        )

        if user["role"] == "admin":
            page = st.sidebar.radio(
                "Khu vực làm việc",
                ["📝 Nhập sàng lọc", "🛡️ Quản trị hệ thống"],
                key="auth_page_admin",
            )
        else:
            page = "📝 Nhập sàng lọc"
            st.sidebar.info(
                "Bạn có thể nhập hồ sơ/sàng lọc giúp người tham gia. "
                "Danh sách hồ sơ chỉ xem được ở khu vực quản trị có quyền."
            )

        st.session_state["auth_page"] = page

        if st.sidebar.button("🚪 Đăng xuất", use_container_width=True):
            logout_user()
            st.session_state.pop("auth_page", None)
            st.rerun()
        return

    st.sidebar.caption(
        "Người tham gia có thể tự nhập hồ sơ.\n"
        "Quản trị viên/nhân sự được duyệt đăng nhập để quản lý hoặc nhập giúp người tham gia."
    )

    auth_mode = st.sidebar.radio(
        "",
        ["Tài khoản quản trị", "Đăng ký nhân sự"],
        key="auth_mode_v8",
    )

    if auth_mode == "Tài khoản quản trị":
        with st.sidebar.form("login_form"):
            login_value = st.text_input("Tên đăng nhập hoặc email")
            password = st.text_input("Mật khẩu", type="password")
            login_submit = st.form_submit_button(
                "🔑 ĐĂNG NHẬP",
                use_container_width=True,
            )

        if login_submit:
            user_result, message = authenticate_user(login_value, password)
            if user_result:
                st.session_state["auth_user"] = user_result
                st.session_state["auth_page"] = "📝 Nhập sàng lọc"
                st.rerun()
            else:
                st.sidebar.error(message)

        st.sidebar.info(
            "Tài khoản quản trị đã được cấu hình sẵn cho hệ thống. "
            "Đăng nhập bằng tài khoản được cấp cho quản trị viên."
        )
    else:
        with st.sidebar.form("staff_register_form"):
            staff_username = st.text_input("Tên đăng nhập")
            staff_full_name = st.text_input("Họ và tên")
            staff_email = st.text_input("Email")
            staff_password = st.text_input("Mật khẩu", type="password")
            staff_password2 = st.text_input("Nhập lại mật khẩu", type="password")
            register_submit = st.form_submit_button(
                "📝 GỬI YÊU CẦU TẠO TÀI KHOẢN",
                use_container_width=True,
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
                        "✅ Đã gửi tài khoản. Quản trị viên phải phê duyệt trước khi đăng nhập."
                    )
                else:
                    st.sidebar.error(msg)

        st.sidebar.caption(
            "Tài khoản nhân sự không được xem dữ liệu người tham gia cho đến khi quản trị viên phê duyệt."
        )


def render_admin_console(user):
    st.header("🛡️ QUẢN TRỊ HỆ THỐNG")
    st.success(
        f"Xin chào **{user['full_name']}** — quyền: "
        f"{'Quản trị viên' if user['role'] == 'admin' else 'Nhân sự được duyệt'}"
    )

    patients = list_patient_profiles_for_staff()
    users = list_user_accounts()
    screening_rows = list_latest_screening_records_for_staff()
    screening_history_rows = list_screening_records_for_staff()

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Hồ sơ đã đồng ý", len(patients))
    with c2:
        pending_count = sum(1 for row in users if row[5] == "pending" and row[4] == "staff")
        st.metric("Chờ phê duyệt", pending_count)
    with c3:
        staff_count = sum(1 for row in users if row[4] == "staff" and row[5] == "approved")
        st.metric("Nhân sự được duyệt", staff_count)
    with c4:
        st.metric("Người có lượt sàng lọc", len(screening_rows))

    if user["role"] == "admin":
        st.subheader("👥 Phê duyệt tài khoản nhân sự")
        pending_users = [row for row in users if row[4] == "staff" and row[5] == "pending"]
        if not pending_users:
            st.info("Hiện không có tài khoản nào đang chờ phê duyệt.")
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
                        with b1:
                            if st.button("✅ Duyệt", key=f"approve_{row[0]}"):
                                update_staff_status(row[0], "approved", user["username"])
                                st.rerun()
                        with b2:
                            if st.button("❌ Từ chối", key=f"reject_{row[0]}"):
                                update_staff_status(row[0], "rejected", user["username"])
                                st.rerun()

        st.subheader("⚙️ Tài khoản nhân sự")
        staff_rows = [row for row in users if row[4] == "staff"]
        for row in staff_rows:
            with st.container(border=True):
                s1, s2, s3 = st.columns([2, 2, 1])
                with s1:
                    st.write(f"**{row[2]}**")
                    st.caption(f"{row[1]} · {row[3]}")
                with s2:
                    st.write(f"Trạng thái: **{row[5]}**")
                with s3:
                    if row[5] == "approved":
                        if st.button("Vô hiệu hóa", key=f"disable_{row[0]}"):
                            update_staff_status(row[0], "disabled", user["username"])
                            st.rerun()
                    elif row[5] in {"disabled", "rejected"}:
                        if st.button("Mở lại", key=f"enable_{row[0]}"):
                            update_staff_status(row[0], "approved", user["username"])
                            st.rerun()

    st.divider()
    st.subheader("📊 DỮ LIỆU NGƯỜI THAM GIA — DẠNG BẢNG")
    st.caption(
        "🔒 Chỉ quản trị viên và nhân sự đã được quản trị viên phê duyệt mới xem được dữ liệu này. "
        "Bảng trên web hiển thị lượt mới nhất cho mỗi số điện thoại; toàn bộ lịch sử vẫn được giữ trong database và file Excel nghiên cứu."
    )

    tab1, tab2 = st.tabs(["👤 Hồ sơ hiện tại", "🧪 Lượt sàng lọc gần nhất"])

    with tab1:
        patient_columns = [
            "Số điện thoại", "Họ tên", "Ngày sinh", "Giới tính",
            "Địa chỉ hiện tại", "Tỉnh/thành", "Phường/xã/đặc khu",
            "Đồng ý nghiên cứu", "Phiên bản consent", "Thời điểm đồng ý",
            "Tạo lúc", "Cập nhật lúc",
        ]
        patient_table = []
        for row in patients:
            patient_table.append({
                "Số điện thoại": row[0],
                "Họ tên": row[1],
                "Ngày sinh": row[2],
                "Giới tính": row[3],
                "Địa chỉ hiện tại": row[4],
                "Tỉnh/thành": row[5],
                "Phường/xã/đặc khu": row[6],
                "Đồng ý nghiên cứu": "Có" if row[7] else "Không",
                "Phiên bản consent": row[8] or "",
                "Thời điểm đồng ý": row[9] or "",
                "Tạo lúc": row[10],
                "Cập nhật lúc": row[11],
            })
        if patient_table:
            st.dataframe(patient_table, use_container_width=True, hide_index=True)
        else:
            st.info("Chưa có hồ sơ nào đã đồng ý tham gia.")

    with tab2:
        screening_columns = [
            "ID", "Thời điểm", "Người nhập", "Hình thức nhập", "Số điện thoại",
            "Họ tên", "Tỉnh/thành", "Phường/xã/đặc khu", "Điểm Vòng 1",
            "Nguy cơ Vòng 1", "Vòng 2", "Độ cao", "Hb", "Hb sau hiệu chỉnh",
            "MCV", "MCH", "RBC", "RDW", "Mentzer", "Điểm CBC",
            "Nguy cơ Vòng 2", "Kết luận", "Vòng 3", "Ngày xét nghiệm",
            "Cơ sở thực hiện", "Loại xét nghiệm", "HbA2", "HbF", "HbE", "Ferritin",
            "Kết quả di truyền", "Kết luận xét nghiệm", "Trạng thái theo dõi", "Ghi chú tư vấn",
            "Ngày hẹn theo dõi",
        ]
        screening_table = []
        for row in screening_rows:
            screening_table.append({
                "ID": row[0],
                "Thời điểm": row[1],
                "Người nhập": row[2] or "",
                "Hình thức nhập": "Tự nhập" if row[3] == "self" else "Nhập giúp người tham gia",
                "Số điện thoại": row[4],
                "Họ tên": row[5],
                "Tỉnh/thành": row[9],
                "Phường/xã/đặc khu": row[10],
                "Điểm Vòng 1": row[13],
                "Nguy cơ Vòng 1": row[14],
                "Vòng 2": "Có" if row[15] else "Chưa",
                "Độ cao": row[16] or "",
                "Hb": row[18] if row[18] is not None else "",
                "Hb sau hiệu chỉnh": row[19] if row[19] is not None else "",
                "MCV": row[20] if row[20] is not None else "",
                "MCH": row[21] if row[21] is not None else "",
                "RBC": row[22] if row[22] is not None else "",
                "RDW": row[23] if row[23] is not None else "",
                "Mentzer": row[24] if row[24] is not None else "",
                "Điểm CBC": row[25] if row[25] is not None else "",
                "Nguy cơ Vòng 2": row[26] or "",
                "Kết luận": row[27] or row[28] or "",
                "Vòng 3": "Có" if row[29] else "Chưa",
                "Ngày xét nghiệm": row[30] or "",
                "Cơ sở thực hiện": row[31] or "",
                "Loại xét nghiệm": row[32] or "",
                "HbA2": row[34] if row[34] is not None else "",
                "HbF": row[35] if row[35] is not None else "",
                "HbE": row[36] if row[36] is not None else "",
                "Ferritin": row[37] if row[37] is not None else "",
                "Kết quả di truyền": row[40] or "",
                "Kết luận xét nghiệm": row[41] or "",
                "Trạng thái theo dõi": row[42] or "",
                "Ghi chú tư vấn": row[43] or "",
                "Ngày hẹn theo dõi": row[44] or "",
            })
        if screening_table:
            st.dataframe(screening_table, use_container_width=True, hide_index=True)
        else:
            st.info("Chưa có lượt sàng lọc nào được lưu.")

    if patients or screening_history_rows:
        excel_data = export_screening_xlsx(patients, screening_history_rows)
        st.download_button(
            "📊 XUẤT DỮ LIỆU EXCEL — TOÀN BỘ LỊCH SỬ (.xlsx)",
            data=excel_data.getvalue(),
            file_name=f"Thalassemia_du_lieu_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption(
            "File Excel gồm 2 sheet: Hồ sơ hiện tại và TOÀN BỘ lịch sử sàng lọc. Bảng trên web chỉ hiển thị lượt mới nhất của từng số điện thoại. "
            "Không xuất mật khẩu/tài khoản nhân sự."
        )


# ------------------------------------------------------------
# CURRENT VIETNAM ADMINISTRATIVE DATA (34 PROVINCES / 3,321 COMMUNES)
# ------------------------------------------------------------

@st.cache_data(ttl=86400, show_spinner=False)
def load_admin_hierarchy():
    """Load the current 2-level administrative hierarchy.

    Source dataset follows Vietnam's post-2025 structure: 34 provincial
    units directly administering 3,321 commune/ward/special-area units.
    The prototype deliberately does NOT fall back to manual commune entry.
    """
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
        raise ValueError(
            f"Dữ liệu địa giới hiện trả về {len(provinces)} tỉnh/thành, không phải 34."
        )

    total_communes = sum(len(items) for items in provinces.values())
    if total_communes != 3321:
        raise ValueError(
            f"Dữ liệu địa giới hiện có {total_communes} đơn vị cấp xã, không phải 3.321."
        )

    return provinces


# ------------------------------------------------------------
# GENERAL HELPERS
# ------------------------------------------------------------

def calculate_age(birth_date):
    today = date.today()

    return (
        today.year
        - birth_date.year
        - (
            (today.month, today.day)
            < (birth_date.month, birth_date.day)
        )
    )


def safe_filename(text):
    text = re.sub(
        r"[^0-9A-Za-zÀ-ỹĐđ _-]",
        "_",
        (text or "").strip(),
    )
    return text.strip(" _") or "nguoi_sang_loc"


def reset_results():
    st.session_state.pop("screening_id", None)
    for key in list(st.session_state.keys()):
        if (
            key.startswith("round1_")
            or key.startswith("round2_")
            or key.startswith("low_cbc_")
            or key.startswith("google_")
        ):
            del st.session_state[key]


# ------------------------------------------------------------
# ALTITUDE
# ------------------------------------------------------------

def altitude_adjustment_from_choice(choice):
    return ALTITUDE_OPTIONS.get(
        choice,
        0.0,
    )


def altitude_selector(key_prefix):
    choice = st.radio(
        "Chọn khoảng độ cao nơi đang sinh sống",
        list(ALTITUDE_OPTIONS.keys()),
        index=0,
        key=f"{key_prefix}_choice",
    )

    st.link_button(
        "🔎 Tra cứu độ cao nơi ở",
        ALTITUDE_LOOKUP_URL,
    )

    adjustment = altitude_adjustment_from_choice(
        choice
    )

    st.info(
        f"Hiệu chỉnh Hb tham khảo: **-{adjustment:.1f} g/dL**"
    )

    if choice.startswith(
        (
            "2.500",
            "3.000",
            "3.500",
            "4.000",
            "4.500",
        )
    ):
        st.warning(
            "Ở độ cao ≥2.500 m, WHO lưu ý mức độ không chắc chắn "
            "của hiệu chỉnh cao hơn."
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
# ROUND 1 SCORE
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
        reasons.append(
            "Từng có kết quả nghi ngờ Thalassemia/hemoglobinopathy"
        )

    elif a.get("q9") == "Đã xác định mang gen":
        score += 4
        reasons.append(
            "Từng được xác định mang gen"
        )

    return min(
        score,
        ROUND1_MAX_SCORE,
    ), reasons


def round1_category(score):
    """Phân tầng thông tin Vòng 1; không dùng để loại người tham gia khỏi Vòng 2."""
    if score >= ROUND1_HIGH_THRESHOLD:
        return (
            "CAO",
            "Có nhiều yếu tố đáng lưu ý qua sàng lọc ban đầu; nên ưu tiên thực hiện/nhập CBC và được tư vấn y tế phù hợp.",
        )
    if score >= 4:
        return (
            "TRUNG BÌNH",
            "Có một số yếu tố đáng lưu ý; nên cân nhắc CBC nếu chưa có kết quả và tiếp tục theo dõi.",
        )
    return (
        "THẤP",
        "Chưa ghi nhận nhiều yếu tố đáng lưu ý qua bộ câu hỏi ban đầu; kết quả này không loại trừ khả năng mang gen Thalassemia.",
    )


def round1_guidance(category, reasons):
    """Tạo lời khuyên hành động rõ ràng sau Vòng 1."""
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

# ============================================================
# ROUND 2
# ============================================================

if st.session_state.get("round2_unlocked", False):

    st.divider()
    st.header("🟧 VÒNG 2 — ĐỘ CAO + CBC")
    st.caption(
        "Nhập toàn bộ thông số một lần rồi bấm PHÂN TÍCH VÒNG 2. "
        "Trong lúc nhập, form không gửi dữ liệu lên database."
    )

    with st.form("round2_form_v21", clear_on_submit=False):
        with st.container(border=True):
            st.subheader("1. Nơi cư trú và độ cao")

            st.write(f"**Tỉnh/thành:** {patient['province']}")
            st.write(f"**Phường/xã/đặc khu:** {patient['commune']}")
            st.write(f"**Địa chỉ hiện tại:** {patient['current_address']}")

            altitude_choice = st.selectbox(
                "Chọn khoảng độ cao nơi đang sinh sống",
                list(ALTITUDE_OPTIONS.keys()),
                index=0,
                key="round2_altitude_choice_v21",
            )
            altitude_adjustment = altitude_adjustment_from_choice(altitude_choice)

            st.link_button("🔎 Tra cứu độ cao nơi ở", ALTITUDE_LOOKUP_URL)
            st.info(f"Hiệu chỉnh Hb tham khảo: **-{altitude_adjustment:.1f} g/dL**")

        with st.container(border=True):
            st.subheader("2. Nhập Công thức máu")

            c1, c2, c3, c4, c5 = st.columns(5)

            with c1:
                hb_unit = st.selectbox(
                    "Đơn vị Hb",
                    ["g/dL", "g/L"],
                    key="round2_hb_unit_v21",
                )
                hb_raw = st.number_input(
                    "Hb",
                    min_value=0.0,
                    max_value=100000.0,
                    value=None,
                    step=0.1,
                    key="round2_hb_raw_v21",
                )

            with c2:
                mcv = st.number_input(
                    "MCV (fL)",
                    30.0, 150.0, None, 0.1,
                    key="round2_mcv_v21",
                )

            with c3:
                mch = st.number_input(
                    "MCH (pg)",
                    10.0, 50.0, None, 0.1,
                    key="round2_mch_v21",
                )

            with c4:
                rbc_unit = st.selectbox(
                    "Đơn vị RBC",
                    ["T/L", "10^12/L", "10^6/µL"],
                    key="round2_rbc_unit_v21",
                )
                rbc_raw = st.number_input(
                    "RBC",
                    1.0, 10.0, None, 0.1,
                    key="round2_rbc_raw_v21",
                )

            with c5:
                rdw = st.number_input(
                    "RDW-CV (%)",
                    5.0, 40.0, None, 0.1,
                    key="round2_rdw_v21",
                )

            st.caption("Hệ thống tự chuyển Hb → g/dL và RBC → T/L trước khi tính.")
            st.caption(
                "Tiêu chí sàng lọc CBC theo Hướng dẫn của Bộ Y tế 2022: "
                "MCV <85 fL và/hoặc MCH <28 pg. Đây là tiêu chí sàng lọc, "
                "không phải tiêu chuẩn chẩn đoán xác định."
            )

        round2_submit = st.form_submit_button(
            "🩸 PHÂN TÍCH VÒNG 2",
            type="primary",
            use_container_width=True,
        )

    if round2_submit:
        required_cbc = {
            "Hb": hb_raw,
            "MCV": mcv,
            "MCH": mch,
            "RBC": rbc_raw,
            "RDW-CV": rdw,
        }
        missing = [name for name, value in required_cbc.items() if value is None]

        if missing:
            st.error(
                "Vui lòng nhập đầy đủ số liệu CBC trước khi phân tích: "
                + ", ".join(missing) + "."
            )
        elif rbc_raw <= 0:
            st.error("RBC phải lớn hơn 0.")
        else:
            hb = hb_to_g_dl(hb_raw, hb_unit)
            rbc = rbc_to_t_l(rbc_raw, rbc_unit)

            hb_adjusted = hb - altitude_adjustment

            score, mentzer, reasons = calculate_round2_score(
                mcv, mch, rbc, rdw
            )
            category, conclusion = round2_category(
                score, mcv, mch
            )
            findings, advice = narrative_cbc_advice(
                hb_adjusted, mcv, mch, rbc, rdw, mentzer
            )

            r2 = {
                "altitude_choice": altitude_choice,
                "adjustment": altitude_adjustment,
                "hb": hb,
                "hb_adjusted": hb_adjusted,
                "mcv": mcv,
                "mch": mch,
                "rbc": rbc,
                "rdw": rdw,
                "mentzer": mentzer,
                "score": score,
                "category": category,
                "conclusion": conclusion,
                "reasons": reasons,
                "findings": findings,
                "advice": advice,
            }

            st.session_state["round2_altitude_choice"] = altitude_choice
            st.session_state["round2_adjustment"] = altitude_adjustment
            st.session_state["round2_hb"] = hb
            st.session_state["round2_hb_adjusted"] = hb_adjusted
            st.session_state["round2_mcv_result"] = mcv
            st.session_state["round2_mch_result"] = mch
            st.session_state["round2_rbc_result"] = rbc
            st.session_state["round2_rdw_result"] = rdw
            st.session_state["round2_mentzer"] = mentzer
            st.session_state["round2_score"] = score
            st.session_state["round2_category"] = category
            st.session_state["round2_conclusion"] = conclusion
            st.session_state["round2_reasons"] = reasons
            st.session_state["round2_findings"] = findings
            st.session_state["round2_advice"] = advice

            update_screening_round2(
                st.session_state.get("screening_id"),
                r2,
            )

            st.session_state["round2_completed"] = True

# ROUND 3 — FOLLOW-UP / SPECIALIZED RESULTS
        # ----------------------------------------------------

        st.subheader("3. THEO DÕI SAU SÀNG LỌC — KẾT QUẢ XÉT NGHIỆM CHUYÊN SÂU")
        st.caption(
            "Nếu người tham gia đã thực hiện điện di huyết sắc tố, HPLC, xét nghiệm sắt hoặc xét nghiệm di truyền, "
            "có thể nhập bổ sung tại đây để phục vụ theo dõi và tư vấn về sau. Vòng 3 không chấm điểm và không tự xác lập chẩn đoán."
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
                r3_facility = st.text_input("Cơ sở thực hiện", key="round3_facility", placeholder="Bệnh viện / phòng xét nghiệm")

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

            r3_genetic = st.text_area("Kết quả xét nghiệm di truyền (nếu có)", key="round3_genetic", placeholder="Ghi theo phiếu xét nghiệm hoặc tóm tắt chính xác kết quả...")
            r3_lab = st.text_area("Kết luận trên phiếu xét nghiệm / nhận xét của cơ sở thực hiện", key="round3_lab", placeholder="Ưu tiên ghi đúng nội dung của cơ sở xét nghiệm, không tự suy diễn.")

            q1, q2 = st.columns(2)
            with q1:
                r3_status = st.selectbox(
                    "Trạng thái theo dõi",
                    ["Đã có kết quả", "Đã thực hiện — đang chờ kết quả", "Chưa thực hiện"],
                    key="round3_status",
                )
            with q2:
                r3_follow_date = st.date_input("Ngày hẹn theo dõi tiếp theo (nếu có)", value=None, key="round3_follow_date")

            r3_note = st.text_area(
                "Ghi chú tư vấn / theo dõi",
                key="round3_note",
                placeholder="Ví dụ: mang kết quả đến bác sĩ Huyết học; cân nhắc tư vấn di truyền khi phù hợp...",
            )

            save_r3 = st.form_submit_button("LƯU KẾT QUẢ THEO DÕI", type="primary", use_container_width=True)

        if save_r3:
            has_result = any(v is not None for v in [r3_hba, r3_hba2, r3_hbf, r3_hbe, r3_ferritin, r3_iron, r3_tsat]) or bool(r3_genetic.strip()) or bool(r3_lab.strip())
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
                        "genetic_result": r3_genetic.strip() or None, "lab_conclusion": r3_lab.strip() or None,
                        "followup_status": r3_status, "counseling_note": r3_note.strip() or None,
                        "followup_date": r3_follow_date.isoformat() if r3_follow_date else None,
                    },
                )
                st.session_state["round3_completed"] = True
                st.success("Đã lưu kết quả theo dõi sau sàng lọc.")

        if st.session_state.get("round3_completed", False):
            st.info("Hồ sơ đã có dữ liệu Vòng 3. Khi có kết quả mới, bạn có thể cập nhật lại để duy trì hồ sơ theo dõi hiện hành.")

        # ----------------------------------------------------
        # ROUND 3 — BYT 2022 INTERPRETATION SUPPORT
        # ----------------------------------------------------
        def _round3_byt_interpretation(r3):
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
            gen = (r3.get("genetic_result") or "").lower()
            lab = (r3.get("lab_conclusion") or "").lower()
            combined = gen + " " + lab
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

        if st.session_state.get("round3_completed", False):
            st.markdown("### Đối chiếu kết quả theo hướng dẫn BYT 2022")
            current_r3 = {
                "hba": st.session_state.get("round3_hba"), "hba2": st.session_state.get("round3_hba2"),
                "hbf": st.session_state.get("round3_hbf"), "hbe": st.session_state.get("round3_hbe"),
                "ferritin": st.session_state.get("round3_ferritin"),
                "transferrin_saturation": st.session_state.get("round3_tsat"),
                "genetic_result": st.session_state.get("round3_genetic", ""),
                "lab_conclusion": st.session_state.get("round3_lab", ""),
            }
            for note in _round3_byt_interpretation(current_r3):
                st.write("• " + note)
            st.caption("Đây là đối chiếu hỗ trợ dựa trên dữ liệu đã nhập và nội dung tài liệu BYT 2022; không thay thế kết luận của cơ sở xét nghiệm/bác sĩ Huyết học.")

        # ----------------------------------------------------
        # MEDICAL FACILITIES
        # ----------------------------------------------------

        st.subheader("🏥 DANH SÁCH CƠ SỞ Y TẾ")
        st.caption(
            "Tra cứu bệnh viện/cơ sở y tế theo tỉnh, tên cơ sở hoặc tuyến. "
            "Danh mục là dữ liệu điều hướng của prototype và cần được đối soát "
            "với bệnh viện trước khi sử dụng thực tế."
        )

        all_facilities = all_medical_facilities()
        province_options = ["Tất cả tỉnh/thành"] + sorted(
            {x["province"] for x in all_facilities}
        )
        selected_facility_province = st.selectbox(
            "Tỉnh/thành",
            province_options,
            index=(province_options.index(patient["province"])
                   if patient["province"] in province_options else 0),
            key="facility_province_filter",
        )
        facility_keyword = st.text_input(
            "🔎 Tìm bệnh viện",
            placeholder="Ví dụ: Bạch Mai, Huyết học, Đa khoa...",
            key="facility_keyword",
        )

        filtered_facilities = all_facilities
        if selected_facility_province != "Tất cả tỉnh/thành":
            province_key = PROVINCE_FACILITY_ALIASES.get(
                selected_facility_province,
                selected_facility_province,
            )
            filtered_facilities = [
                x for x in filtered_facilities
                if x["province"] == province_key
                or x["province"] == selected_facility_province
            ]

        if facility_keyword.strip():
            kw = facility_keyword.strip().lower()
            filtered_facilities = [
                x for x in filtered_facilities
                if kw in " ".join([
                    x.get("name", ""),
                    x.get("tier", ""),
                    x.get("note", ""),
                    x.get("province", ""),
                ]).lower()
            ]

        st.write(f"**{len(filtered_facilities)} cơ sở** phù hợp với bộ lọc.")

        if filtered_facilities:
            for i, facility in enumerate(filtered_facilities, start=1):
                with st.container(border=True):
                    st.markdown(f"### {i}. {facility['name']}")
                    st.write(f"**Tỉnh/thành:** {facility['province']}")
                    st.write(f"**Phân loại:** {facility['tier']}")
                    st.write(f"**Gợi ý:** {facility['note']}")
                    st.link_button(
                        "🗺️ Xem vị trí / chỉ đường",
                        facility["maps"],
                    )
        else:
            st.info(
                "Chưa tìm thấy cơ sở phù hợp. Hãy thử bỏ bộ lọc hoặc tìm bằng "
                "tên bệnh viện khác."
            )

        st.divider()
        st.markdown("### ⭐ Cơ sở ưu tiên theo nơi ở của người bệnh")

        facilities = recommended_facilities(patient["province"])
        if facilities:
            st.success(
                f"Đã tìm thấy {len(facilities)} cơ sở ưu tiên cho **{patient['province']}**."
            )
            for i, facility in enumerate(facilities, start=1):
                with st.container(border=True):
                    st.markdown(f"**{i}. {facility['name']}**")
                    st.caption(f"{facility['tier']} · {facility['note']}")
                    st.link_button(
                        "🗺️ Xem vị trí / chỉ đường",
                        facility["maps"],
                        key=f"recommended_facility_{i}_{patient['phone']}",
                    )
        else:
            maps_query = f"bệnh viện {patient['province']}"
            maps_url = (
                "https://www.google.com/maps/search/?api=1&query="
                + requests.utils.quote(maps_query)
            )
            st.info(
                f"Chưa có cơ sở được chọn sẵn cho **{patient['province']}**. "
                "Bạn có thể mở Google Maps để xem các bệnh viện gần khu vực."
            )
            st.link_button("🗺️ Tìm bệnh viện trên Google Maps", maps_url)

        st.caption(
            "⚠️ Danh mục chỉ hỗ trợ định hướng. Khả năng thực hiện CBC, HPLC/điện di Hb, "
            "xét nghiệm gen Thalassemia và tiếp nhận chuyên khoa có thể khác nhau theo cơ sở; "
            "nên gọi xác nhận trước khi đến."
        )

        # ----------------------------------------------------
        # WORD
        # ----------------------------------------------------

        st.subheader(
            "📄 PHIẾU KẾT QUẢ"
        )

        report = make_word(
            patient=patient,
            r1_score=st.session_state[
                "round1_score"
            ],
            r1_category=st.session_state[
                "round1_category"
            ],
            r1_reasons=st.session_state[
                "round1_reasons"
            ],
            r2={
                "altitude_choice": st.session_state[
                    "round2_altitude_choice"
                ],
                "adjustment": st.session_state[
                    "round2_adjustment"
                ],
                "hb": st.session_state[
                    "round2_hb"
                ],
                "hb_adjusted": st.session_state[
                    "round2_hb_adjusted"
                ],
                "mcv": st.session_state[
                    "round2_mcv_result"
                ],
                "mch": st.session_state[
                    "round2_mch_result"
                ],
                "rbc": st.session_state[
                    "round2_rbc_result"
                ],
                "rdw": st.session_state[
                    "round2_rdw_result"
                ],
                "mentzer": st.session_state[
                    "round2_mentzer"
                ],
                "score": st.session_state[
                    "round2_score"
                ],
                "category": st.session_state[
                    "round2_category"
                ],
                "conclusion": st.session_state[
                    "round2_conclusion"
                ],
                "findings": st.session_state[
                    "round2_findings"
                ],
                "advice": st.session_state[
                    "round2_advice"
                ],
            },
        )

        st.download_button(
            "📥 TẢI PHIẾU WORD",
            data=report,
            file_name=(
                "Phieu_Thalassemia_"
                f"{safe_filename(patient['full_name'])}.docx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True,
        )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "V5 Prototype — các trọng số/ngưỡng cần validation trên dữ liệu "
    "người Việt Nam trước khi sử dụng trong nghiên cứu lâm sàng."
)
