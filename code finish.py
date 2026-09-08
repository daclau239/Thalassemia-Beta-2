
import io
import json
import math
import os
import re
import sqlite3
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
# - Một số điện thoại = một hồ sơ; nhập lại sẽ ghi nhận lần cuối.
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

DB_PATH = "thalassemia_patients.db"
CONSENT_VERSION = "THAL-RS-CONSENT-v1-2026-09-05"
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

# Tài khoản quản trị mặc định theo yêu cầu của chủ hệ thống.
DEFAULT_ADMIN_USERNAME = "daclau239"
DEFAULT_ADMIN_PASSWORD = "23092002"
DEFAULT_ADMIN_FULL_NAME = "Quản trị viên hệ thống"
DEFAULT_ADMIN_EMAIL = "admin@thalassemia.local"


def get_db():
    conn = sqlite3.connect(
        DB_PATH,
        check_same_thread=False,
    )

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

    # Tài khoản quản trị / người được phê duyệt.
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS user_accounts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
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

    # Nhật ký từng lần sàng lọc: giữ lịch sử theo lượt, tách khỏi hồ sơ hiện tại.
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS screening_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
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

    # Chỉ giữ lại 01 lượt sàng lọc mới nhất cho mỗi số điện thoại.
    # Các lượt cũ được xóa khỏi bảng screening_records để dữ liệu nghiên cứu
    # luôn có đúng 01 bản ghi hiện hành cho mỗi người. Hồ sơ patient_profiles
    # vẫn được giữ nguyên.
    conn.execute(
        """
        DELETE FROM screening_records
        WHERE id NOT IN (
            SELECT MAX(id)
            FROM screening_records
            GROUP BY phone
        )
        """
    )

    # Migrate prototype databases created before consent fields existed.
    existing = {
        row[1]
        for row in conn.execute(
            "PRAGMA table_info(patient_profiles)"
        ).fetchall()
    }
    migrations = [
        ("research_consent", "INTEGER NOT NULL DEFAULT 0"),
        ("consent_version", "TEXT"),
        ("consent_at", "TEXT"),
    ]
    for column, definition in migrations:
        if column not in existing:
            conn.execute(
                f"ALTER TABLE patient_profiles ADD COLUMN {column} {definition}"
            )

    # Migration cho dữ liệu theo dõi Vòng 3 / xét nghiệm chuyên sâu.
    record_columns = {row[1] for row in conn.execute("PRAGMA table_info(screening_records)").fetchall()}
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

    conn.commit()
    return conn


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
    """Tạo admin mặc định một lần nếu database chưa có tài khoản admin."""
    conn = get_db()
    row = conn.execute(
        "SELECT id FROM user_accounts WHERE username = ? LIMIT 1",
        (DEFAULT_ADMIN_USERNAME,),
    ).fetchone()
    if row:
        conn.close()
        return

    admin_row = conn.execute(
        "SELECT id FROM user_accounts WHERE role = 'admin' LIMIT 1"
    ).fetchone()
    if admin_row:
        conn.close()
        return

    now = datetime.now().isoformat(timespec="seconds")
    password_hash, password_salt = hash_password(DEFAULT_ADMIN_PASSWORD)
    conn.execute(
        """
        INSERT INTO user_accounts
        (username, full_name, email, password_hash, password_salt,
         role, status, created_at, approved_by, approved_at)
        VALUES (?, ?, ?, ?, ?, 'admin', 'approved', ?, ?, ?)
        """,
        (
            DEFAULT_ADMIN_USERNAME,
            DEFAULT_ADMIN_FULL_NAME,
            DEFAULT_ADMIN_EMAIL,
            password_hash,
            password_salt,
            now,
            DEFAULT_ADMIN_USERNAME,
            now,
        ),
    )
    conn.commit()
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
    record_id = cur.lastrowid

    # Khi một người (cùng số điện thoại) sàng lọc lại, lượt mới thay thế
    # hoàn toàn lượt cũ. Không xóa hồ sơ người tham gia, chỉ xóa bản ghi
    # sàng lọc cũ để kho dữ liệu nghiên cứu có 01 dòng/người.
    conn.execute(
        "DELETE FROM screening_records WHERE phone = ? AND id <> ?",
        (patient["phone"], int(record_id)),
    )

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
    """Lấy 01 lượt sàng lọc duy nhất cho mỗi người.

    Hệ thống chủ động xóa lượt cũ khi có lượt mới, nên CSDL và file Excel
    đều phục vụ tập dữ liệu nghiên cứu theo nguyên tắc 01 người/01 dòng.
    """
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
    ws2 = workbook.add_worksheet("Luot_sang_loc_moi_nhat")
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

ensure_default_admin()


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
    screening_rows = list_screening_records_for_staff()

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
        "Bảng sàng lọc chỉ giữ một bản ghi hiện hành cho mỗi số điện thoại; khi sàng lọc lại, bản ghi cũ được thay thế để tạo dataset nghiên cứu sạch."
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

    if patients or screening_rows:
        excel_data = export_screening_xlsx(patients, screening_rows)
        st.download_button(
            "📊 XUẤT DỮ LIỆU EXCEL (.xlsx)",
            data=excel_data.getvalue(),
            file_name=f"Thalassemia_du_lieu_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption(
            "File Excel gồm 2 sheet: Hồ sơ hiện tại và lượt sàng lọc gần nhất của từng số điện thoại. "
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
# ROUND 2 CBC ANALYSIS
# ------------------------------------------------------------

def calculate_round2_score(
    mcv,
    mch,
    rbc,
    rdw,
):
    """
    Sàng lọc CBC theo quy trình trong Hướng dẫn chẩn đoán và điều trị
    một số bệnh lý huyết học (Bộ Y tế, 2022), Phụ lục: Quy trình xét nghiệm
    sàng lọc, chẩn đoán Thalassemia.

    Tiêu chí gợi ý bước đánh giá tiếp theo trong tài liệu: MCV < 85 fL
    và/hoặc MCH < 28 pg. Mentzer được lưu như chỉ số hỗ trợ, nhưng không
    được dùng làm tiêu chuẩn chẩn đoán của Bộ Y tế.
    """
    reasons = []
    score = 0

    # Tiêu chí sàng lọc theo BYT 2022: MCV <85 và/hoặc MCH <28.
    if mcv < 85:
        reasons.append("MCV <85 fL — đạt tiêu chí sàng lọc CBC theo phụ lục hướng dẫn BYT 2022")
        score += 2
    if mch < 28:
        reasons.append("MCH <28 pg — đạt tiêu chí sàng lọc CBC theo phụ lục hướng dẫn BYT 2022")
        score += 2

    # Các đặc điểm bổ sung để mô tả kiểu hình, không thay thế tiêu chuẩn BYT.
    if mcv < 80:
        score += 1
        reasons.append("MCV <80 fL — hồng cầu nhỏ")
    if mch < 27:
        score += 1
        reasons.append("MCH <27 pg — xu hướng hồng cầu nhược sắc")
    if rdw > 14:
        reasons.append("RDW >14% — kích thước hồng cầu không đồng đều theo hướng dẫn BYT 2022")

    # RBC chỉ là thông tin hỗ trợ bối cảnh.
    if mcv < 85 and rbc >= 5.0:
        score += 1
        reasons.append("RBC tương đối cao trong bối cảnh MCV giảm")

    mentzer = (_round_number(Decimal(str(mcv)) / Decimal(str(rbc)), 2)
               if rbc > 0 else None)
    if mentzer is not None:
        reasons.append(f"Mentzer Index = {mentzer:.2f} — chỉ số hỗ trợ, không phải tiêu chuẩn chẩn đoán")

    return score, mentzer, reasons


def round2_category(score, mcv, mch=None):
    """Phân tầng nguy cơ sàng lọc; tiêu chí CBC cốt lõi bám theo BYT 2022."""
    screening_trigger = (mcv < 85) or (mch is not None and mch < 28)
    if not screening_trigger:
        return (
            "THẤP",
            "CBC hiện tại chưa đạt tiêu chí sàng lọc MCV <85 fL và/hoặc MCH <28 pg trong phụ lục hướng dẫn BYT 2022; điều này không loại trừ hoàn toàn Thalassemia.",
        )
    if score >= 6:
        return (
            "CAO",
            "CBC có tiêu chí sàng lọc đáng lưu ý (MCV <85 fL và/hoặc MCH <28 pg). Theo quy trình BYT 2022, nên đánh giá tình trạng sắt và xác định thành phần huyết sắc tố bằng điện di/HPLC; xét nghiệm gen tùy trường hợp.",
        )
    return (
        "TRUNG BÌNH",
        "CBC đạt tiêu chí sàng lọc MCV <85 fL và/hoặc MCH <28 pg. Cần đánh giá thiếu sắt và nguyên nhân hồng cầu nhỏ; nếu phù hợp, thực hiện điện di/HPLC huyết sắc tố.",
    )


def narrative_cbc_advice(
    hb_adjusted,
    mcv,
    mch,
    rbc,
    rdw,
    mentzer,
):
    findings = []
    advice = []

    if mcv < 85:
        findings.append(
            "MCV <85 fL — đạt tiêu chí sàng lọc CBC theo phụ lục BYT 2022."
        )

    if mch < 28:
        findings.append(
            "MCH <28 pg — đạt tiêu chí sàng lọc CBC theo phụ lục BYT 2022."
        )

    if rdw > 14:
        findings.append(
            "RDW tăng; cần lưu ý thiếu sắt hoặc các nguyên nhân "
            "khác của microcytosis."
        )

    if mcv < 85 and rbc >= 5.0:
        findings.append(
            "RBC tương đối cao trong bối cảnh MCV thấp."
        )

    if mcv < 85 and mentzer < 13:
        findings.append(
            "Mentzer Index <13: mẫu hình sàng lọc nghiêng về "
            "Thalassemia hơn thiếu sắt."
        )

        advice.append(
            "Trao đổi với nhân viên y tế về HPLC/điện di hemoglobin; "
            "tùy trường hợp có thể cần xét nghiệm phân tử."
        )

    elif mcv < 85 and mentzer >= 13:
        findings.append(
            "Mentzer Index ≥13: mẫu hình sàng lọc nghiêng về "
            "thiếu sắt hoặc nguyên nhân microcytosis khác."
        )

        advice.append(
            "Cân nhắc đánh giá tình trạng sắt, đặc biệt Ferritin, "
            "theo chỉ định của nhân viên y tế."
        )

    if hb_adjusted < 8:
        advice.append(
            "Hb sau hiệu chỉnh rất thấp: nên được đánh giá y tế sớm."
        )
    elif hb_adjusted < 10:
        advice.append(
            "Hb sau hiệu chỉnh thấp đáng kể: nên khám và đánh giá nguyên nhân thiếu máu."
        )
    elif hb_adjusted < 12:
        advice.append(
            "Hb sau hiệu chỉnh thấp/giáp ranh tùy nhóm đối tượng; "
            "cần đối chiếu tuổi, giới và khoảng tham chiếu của labo."
        )

    if not findings:
        findings.append(
            "Chưa ghi nhận microcytosis/nhược sắc rõ trên CBC."
        )

    if not advice:
        advice.append(
            "Tiếp tục đối chiếu với khoảng tham chiếu trên phiếu xét nghiệm "
            "và hướng dẫn của nhân viên y tế."
        )

    return findings, advice


# ------------------------------------------------------------
# CURATED MEDICAL FACILITIES BY PROVINCE/CITY
# ------------------------------------------------------------
# Mục tiêu của prototype:
#   1) Ưu tiên bệnh viện hạng I trong chính tỉnh/thành người dùng đang sống.
#   2) Ưu tiên bệnh viện tuyến Trung ương / hạng đặc biệt nếu có tại địa bàn.
#   3) Hiển thị 3–5 cơ sở phù hợp để người dùng chủ động lựa chọn.
#
# Lưu ý:
# - Đây là danh mục điều hướng cho prototype, KHÔNG phải danh sách chỉ định điều trị.
# - Phân hạng/cấp quản lý có thể thay đổi; bản triển khai nghiên cứu chính thức nên
#   đồng bộ định kỳ từ nguồn dữ liệu Bộ Y tế/website chính thức của bệnh viện.
# - Có thêm alias tên tỉnh cũ để app vẫn hoạt động nếu dữ liệu địa giới cũ còn trong DB.

MEDICAL_FACILITIES = {
    "Hà Nội": [
        {"name": "Bệnh viện Bạch Mai", "tier": "Hạng đặc biệt – tuyến Trung ương", "note": "Ưu tiên khi cần đánh giá chuyên sâu huyết học", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Bạch+Mai+Hà+Nội"},
        {"name": "Bệnh viện Hữu nghị Việt Đức", "tier": "Bệnh viện tuyến Trung ương", "note": "Cơ sở chuyên sâu; phù hợp khi cần tuyến trên", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Hữu+nghị+Việt+Đức+Hà+Nội"},
        {"name": "Bệnh viện Trung ương Quân đội 108", "tier": "Hạng đặc biệt – tuyến Trung ương", "note": "Cơ sở tuyến trên với nhiều chuyên khoa sâu", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Trung+ương+Quân+đội+108"},
    ],
    "Hồ Chí Minh": [
        {"name": "Bệnh viện Chợ Rẫy", "tier": "Hạng đặc biệt – tuyến Trung ương", "note": "Ưu tiên khi cần đánh giá chuyên sâu huyết học", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Chợ+Rẫy+Hồ+Chí+Minh"},
        {"name": "Bệnh viện Thống Nhất", "tier": "Hạng I – tuyến Trung ương", "note": "Bệnh viện đa khoa tuyến trên", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Thống+Nhất+Hồ+Chí+Minh"},
        {"name": "Bệnh viện Nhân Dân 115", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến chuyên sâu", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Nhân+Dân+115+Hồ+Chí+Minh"},
        {"name": "Bệnh viện Đại học Y Dược TP. Hồ Chí Minh", "tier": "Hạng I", "note": "Bệnh viện trường đại học tuyến chuyên sâu", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đại+học+Y+Dược+TP+Hồ+Chí+Minh"},
    ],
    "Đà Nẵng": [
        {"name": "Bệnh viện C Đà Nẵng", "tier": "Bệnh viện tuyến Trung ương", "note": "Cơ sở tuyến Trung ương tại Đà Nẵng", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+C+Đà+Nẵng"},
        {"name": "Bệnh viện Đà Nẵng", "tier": "Hạng I", "note": "Bệnh viện đa khoa lớn của thành phố", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đà+Nẵng"},
        {"name": "Bệnh viện Đại học Y – Dược, Đại học Đà Nẵng", "tier": "Bệnh viện trường đại học", "note": "Có thể lựa chọn khi cần đánh giá chuyên khoa", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đại+học+Y+Dược+Đại+học+Đà+Nẵng"},
    ],
    "Thừa Thiên Huế": [],
    "Huế": [
        {"name": "Bệnh viện Trung ương Huế", "tier": "Hạng đặc biệt – tuyến Trung ương", "note": "Ưu tiên khi cần đánh giá chuyên sâu", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Trung+ương+Huế"},
        {"name": "Bệnh viện Trường Đại học Y – Dược Huế", "tier": "Hạng I", "note": "Bệnh viện trường đại học tuyến chuyên sâu", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Trường+Đại+học+Y+Dược+Huế"},
    ],
    "Cần Thơ": [
        {"name": "Bệnh viện Đa khoa Trung ương Cần Thơ", "tier": "Bệnh viện tuyến Trung ương", "note": "Ưu tiên khi cần tuyến chuyên sâu tại ĐBSCL", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đa+khoa+Trung+ương+Cần+Thơ"},
        {"name": "Bệnh viện Đại học Y Dược Cần Thơ", "tier": "Hạng I", "note": "Bệnh viện trường đại học tuyến chuyên sâu", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đại+học+Y+Dược+Cần+Thơ"},
    ],
    "Thái Nguyên": [
        {"name": "Bệnh viện Trung ương Thái Nguyên", "tier": "Hạng đặc biệt – tuyến Trung ương", "note": "Ưu tiên khi cần đánh giá chuyên sâu", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Trung+ương+Thái+Nguyên"},
    ],
    "Hải Phòng": [
        {"name": "Bệnh viện Hữu nghị Việt Tiệp", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến chuyên sâu", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Hữu+nghị+Việt+Tiệp+Hải+Phòng"},
    ],
    "Quảng Ninh": [
        {"name": "Bệnh viện Bãi Cháy", "tier": "Hạng I", "note": "Bệnh viện đa khoa lớn của tỉnh", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Bãi+Cháy+Quảng+Ninh"},
        {"name": "Bệnh viện Đa khoa tỉnh Quảng Ninh", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến tỉnh", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đa+khoa+tỉnh+Quảng+Ninh"},
    ],
    "Thanh Hóa": [
        {"name": "Bệnh viện Đa khoa tỉnh Thanh Hóa", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến tỉnh", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đa+khoa+tỉnh+Thanh+Hóa"},
    ],
    "Nghệ An": [
        {"name": "Bệnh viện Hữu nghị Đa khoa Nghệ An", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến tỉnh", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Hữu+nghị+Đa+khoa+Nghệ+An"},
    ],
    "Hà Tĩnh": [
        {"name": "Bệnh viện Đa khoa tỉnh Hà Tĩnh", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến tỉnh", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đa+khoa+tỉnh+Hà+Tĩnh"},
    ],
    "Khánh Hòa": [
        {"name": "Bệnh viện Đa khoa tỉnh Khánh Hòa", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến tỉnh", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đa+khoa+tỉnh+Khánh+Hòa"},
    ],
    "Đắk Lắk": [
        {"name": "Bệnh viện Đa khoa vùng Tây Nguyên", "tier": "Hạng I", "note": "Cơ sở tuyến chuyên sâu khu vực Tây Nguyên", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đa+khoa+vùng+Tây+Nguyên"},
    ],
    "Lâm Đồng": [
        {"name": "Bệnh viện Đa khoa tỉnh Lâm Đồng", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến tỉnh", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đa+khoa+tỉnh+Lâm+Đồng"},
    ],
    "Quảng Ngãi": [
        {"name": "Bệnh viện Đa khoa tỉnh Quảng Ngãi", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến tỉnh", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đa+khoa+tỉnh+Quảng+Ngãi"},
    ],
    "Gia Lai": [
        {"name": "Bệnh viện Đa khoa tỉnh Gia Lai", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến tỉnh", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đa+khoa+tỉnh+Gia+Lai"},
    ],
    "Bắc Ninh": [
        {"name": "Bệnh viện Đa khoa tỉnh Bắc Ninh", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến tỉnh", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đa+khoa+tỉnh+Bắc+Ninh"},
    ],
    "Phú Thọ": [
        {"name": "Bệnh viện Đa khoa tỉnh Phú Thọ", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến tỉnh", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đa+khoa+tỉnh+Phú+Thọ"},
    ],
    "Lào Cai": [
        {"name": "Bệnh viện Đa khoa tỉnh Lào Cai", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến tỉnh", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đa+khoa+tỉnh+Lào+Cai"},
    ],
    "Hưng Yên": [
        {"name": "Bệnh viện Đa khoa tỉnh Hưng Yên", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến tỉnh", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đa+khoa+tỉnh+Hưng+Yên"},
    ],
    "Ninh Bình": [
        {"name": "Bệnh viện Đa khoa tỉnh Ninh Bình", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến tỉnh", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đa+khoa+tỉnh+Ninh+Bình"},
    ],
    "Tây Ninh": [
        {"name": "Bệnh viện Đa khoa tỉnh Tây Ninh", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến tỉnh", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đa+khoa+tỉnh+Tây+Ninh"},
    ],
    "Đồng Nai": [
        {"name": "Bệnh viện Đa khoa Đồng Nai", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến tỉnh", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đa+khoa+Đồng+Nai"},
    ],
    "Vĩnh Long": [
        {"name": "Bệnh viện Đa khoa tỉnh Vĩnh Long", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến tỉnh", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đa+khoa+tỉnh+Vĩnh+Long"},
    ],
    "Đồng Tháp": [
        {"name": "Bệnh viện Đa khoa Đồng Tháp", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến tỉnh", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đa+khoa+Đồng+Tháp"},
    ],
    "An Giang": [
        {"name": "Bệnh viện Đa khoa khu vực An Giang", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến tỉnh/khu vực", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đa+khoa+khu+vực+An+Giang"},
    ],
    "Cà Mau": [
        {"name": "Bệnh viện Đa khoa tỉnh Cà Mau", "tier": "Hạng I", "note": "Bệnh viện đa khoa tuyến tỉnh", "maps": "https://www.google.com/maps/search/?api=1&query=Bệnh+viện+Đa+khoa+tỉnh+Cà+Mau"},
    ],
}

# Tên tỉnh cũ → tên tỉnh/thành dùng để tra danh mục cơ sở trong prototype.
PROVINCE_FACILITY_ALIASES = {
    "Thừa Thiên Huế": "Huế",
    "Quảng Nam": "Đà Nẵng",
    "Bình Định": "Gia Lai",
    "Ninh Thuận": "Khánh Hòa",
    "Phú Yên": "Đắk Lắk",
    "Đắk Nông": "Lâm Đồng",
    "Bình Thuận": "Lâm Đồng",
    "Kon Tum": "Quảng Ngãi",
    "Yên Bái": "Lào Cai",
    "Bắc Kạn": "Thái Nguyên",
    "Hòa Bình": "Phú Thọ",
    "Vĩnh Phúc": "Phú Thọ",
    "Hà Nam": "Ninh Bình",
    "Nam Định": "Ninh Bình",
    "Quảng Bình": "Quảng Trị",
    "Bà Rịa - Vũng Tàu": "Hồ Chí Minh",
    "Bình Dương": "Hồ Chí Minh",
    "Long An": "Tây Ninh",
    "Tiền Giang": "Đồng Tháp",
    "Bến Tre": "Vĩnh Long",
    "Trà Vinh": "Vĩnh Long",
    "Sóc Trăng": "Cần Thơ",
    "Hậu Giang": "Cần Thơ",
    "Kiên Giang": "An Giang",
}


def recommended_facilities(province):
    """Trả về 3–5 cơ sở ưu tiên theo tỉnh/thành hiện tại."""
    province_key = PROVINCE_FACILITY_ALIASES.get(province, province)
    facilities = MEDICAL_FACILITIES.get(province_key, [])
    return facilities[:5]


# ------------------------------------------------------------
# GOOGLE PLACES
# ------------------------------------------------------------

@st.cache_data(ttl=3600)
def nearby_medical(
    lat,
    lng,
    api_key,
):
    if not api_key:
        return []

    url = (
        "https://places.googleapis.com/v1/places:searchNearby"
    )

    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": api_key,
        "X-Goog-FieldMask": (
            "places.displayName,"
            "places.formattedAddress,"
            "places.location,"
            "places.googleMapsUri,"
            "places.rating,"
            "places.userRatingCount"
        ),
    }

    payload = {
        "includedTypes": [
            "hospital",
            "medical_center",
            "medical_clinic",
            "medical_lab",
        ],
        "maxResultCount": 20,
        "rankPreference": "DISTANCE",
        "locationRestriction": {
            "circle": {
                "center": {
                    "latitude": lat,
                    "longitude": lng,
                },
                "radius": 50000.0,
            }
        },
    }

    try:
        response = requests.post(
            url,
            headers=headers,
            json=payload,
            timeout=15,
        )
        response.raise_for_status()
        return response.json().get(
            "places",
            [],
        )
    except Exception:
        return []


def haversine_km(
    lat1,
    lon1,
    lat2,
    lon2,
):
    radius = 6371.0

    p1 = math.radians(lat1)
    p2 = math.radians(lat2)

    dp = math.radians(
        lat2 - lat1
    )

    dl = math.radians(
        lon2 - lon1
    )

    a = (
        math.sin(dp / 2) ** 2
        + math.cos(p1)
        * math.cos(p2)
        * math.sin(dl / 2) ** 2
    )

    return (
        2
        * radius
        * math.asin(
            math.sqrt(a)
        )
    )


def rank_facilities(
    places,
    lat,
    lng,
):
    rows = []

    for place in places:
        loc = place.get(
            "location",
            {},
        )

        lat2 = loc.get(
            "latitude"
        )
        lng2 = loc.get(
            "longitude"
        )

        if (
            lat2 is None
            or lng2 is None
        ):
            continue

        rows.append(
            {
                "name": place.get(
                    "displayName",
                    {},
                ).get(
                    "text",
                    "Cơ sở y tế",
                ),
                "address": place.get(
                    "formattedAddress",
                    "Chưa có địa chỉ",
                ),
                "distance": haversine_km(
                    lat,
                    lng,
                    lat2,
                    lng2,
                ),
                "rating": place.get(
                    "rating"
                ),
                "rating_count": place.get(
                    "userRatingCount"
                ),
                "maps_uri": place.get(
                    "googleMapsUri"
                ),
            }
        )

    rows.sort(
        key=lambda x: x["distance"]
    )

    return rows[:5]


# ------------------------------------------------------------
# LOW RISK PANEL
# ------------------------------------------------------------

def low_risk_panel():

    st.success(
        "🟢 **NGUY CƠ SÀNG LỌC BAN ĐẦU: THẤP**"
    )

    st.markdown(
        """
### 📅 Theo dõi sức khỏe

Kết quả Vòng 1 hiện chưa cho thấy nhiều yếu tố nguy cơ rõ ràng.

Bạn nên tiếp tục **theo dõi tình trạng sức khỏe và khám định kỳ
theo hướng dẫn của cơ sở y tế**.

Trong prototype, hệ thống đặt mốc xem xét lại khoảng **30 ngày**.
Đây là mốc theo dõi của ứng dụng, không phải chỉ định bắt buộc
mọi người nguy cơ thấp phải khám hàng tháng.

### 🩸 Khi nào nên kiểm tra CBC lại?

Nếu xuất hiện mệt mỏi kéo dài, da/niêm nhợt, chóng mặt, vàng da/vàng mắt
hoặc phiếu công thức máu có bất thường, hãy đưa kết quả cho nhân viên y tế.
Bạn có thể nhập CBC bên dưới để hệ thống **sàng lọc lại**.
"""
    )

    reminder_date = (
        date.today()
        + timedelta(
            days=FOLLOWUP_DAYS
        )
    )

    st.info(
        f"🗓️ Mốc nhắc prototype: "
        f"**{reminder_date.strftime('%d/%m/%Y')}**"
    )

    with st.expander(
        "🩸 Nhập CBC nếu lần xét nghiệm sau có bất thường",
        expanded=False,
    ):

        a, b, c, d = st.columns(4)

        with a:
            hb_unit = st.selectbox(
                "Đơn vị Hb",
                ["g/dL", "g/L"],
                key="low_cbc_hb_unit",
            )
            hb_raw = st.number_input(
                "Hb",
                min_value=3.0 if hb_unit == "g/dL" else 30.0,
                max_value=25.0 if hb_unit == "g/dL" else 250.0,
                value=13.0 if hb_unit == "g/dL" else 130.0,
                step=0.1,
                key="low_cbc_hb",
            )

        with b:
            mcv = st.number_input(
                "MCV (fL)",
                30.0,
                150.0,
                85.0,
                0.1,
                key="low_cbc_mcv",
            )

        with c:
            mch = st.number_input(
                "MCH (pg)",
                10.0,
                50.0,
                29.0,
                0.1,
                key="low_cbc_mch",
            )

        with d:
            rbc_unit = st.selectbox(
                "Đơn vị RBC",
                ["T/L", "10^12/L", "10^6/µL"],
                key="low_cbc_rbc_unit",
            )
            rbc_raw = st.number_input(
                "RBC",
                1.0,
                10.0,
                4.8,
                0.1,
                key="low_cbc_rbc",
            )

        rdw = st.number_input(
            "RDW-CV (%)",
            5.0,
            40.0,
            13.0,
            0.1,
            key="low_cbc_rdw",
        )

        if st.button(
            "🔎 PHÂN TÍCH CBC BẤT THƯỜNG",
            key="low_cbc_analyze",
        ):

            hb = hb_to_g_dl(
                hb_raw,
                hb_unit,
            )

            rbc = rbc_to_t_l(
                rbc_raw,
                rbc_unit,
            )

            score, mentzer, reasons = (
                calculate_round2_score(
                    mcv,
                    mch,
                    rbc,
                    rdw,
                )
            )

            category, conclusion = (
                round2_category(
                    score,
                    mcv,
                    mch,
                )
            )

            if category == "THẤP":
                st.success(
                    f"🟢 {category} — {conclusion}"
                )
            elif category == "TRUNG BÌNH":
                st.warning(
                    f"🟡 {category} — {conclusion}"
                )
            else:
                st.error(
                    f"🟠 {category} — {conclusion}"
                )

            st.metric(
                "Mentzer Index",
                f"{mentzer:.2f}",
            )

            findings, advice = narrative_cbc_advice(
                hb,
                mcv,
                mch,
                rbc,
                rdw,
                mentzer,
            )

            st.markdown(
                "### 🧠 Phân tích sơ bộ"
            )

            for finding in findings:
                st.write(
                    f"• {finding}"
                )

            st.markdown(
                "### 💡 Khuyến nghị"
            )

            for item in advice:
                st.write(
                    f"→ {item}"
                )


# ------------------------------------------------------------
# WORD REPORT
# ------------------------------------------------------------

def make_word(
    patient,
    r1_score,
    r1_category,
    r1_reasons,
    r2=None,
):
    doc = Document()

    doc.add_heading(
        "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM",
        level=3,
    )

    doc.add_heading(
        "PHIẾU SÀNG LỌC VÀ PHÂN TẦNG NGUY CƠ THALASSEMIA",
        level=1,
    )

    doc.add_paragraph(
        "Công cụ hỗ trợ sàng lọc; không thay thế chẩn đoán "
        "hoặc chỉ định của nhân viên y tế."
    )

    doc.add_heading(
        "I. THÔNG TIN BỆNH NHÂN",
        level=2,
    )

    for label, value in [
        ("Họ và tên", patient["full_name"]),
        ("Ngày sinh", patient["birth_date"]),
        ("Tuổi", patient["age"]),
        ("Giới tính", patient["gender"]),
        ("Số điện thoại", patient["phone"]),
        ("Địa chỉ hiện tại", patient["current_address"]),
        ("Tỉnh/thành", patient["province"]),
        ("Phường/xã/đặc khu", patient["commune"]),
    ]:
        doc.add_paragraph(
            f"{label}: {value}"
        )

    doc.add_heading(
        "II. VÒNG 1",
        level=2,
    )

    doc.add_paragraph(
        f"Điểm: {r1_score}/{ROUND1_MAX_SCORE}"
    )

    doc.add_paragraph(
        f"Mức nguy cơ: {r1_category}"
    )

    for item in r1_reasons:
        doc.add_paragraph(
            f"- {item}"
        )

    if r2:

        doc.add_heading(
            "III. VÒNG 2",
            level=2,
        )

        doc.add_paragraph(
            f"Khoảng độ cao: {r2['altitude_choice']}"
        )

        doc.add_paragraph(
            f"Hiệu chỉnh Hb: -{r2['adjustment']:.1f} g/dL"
        )

        doc.add_paragraph(
            f"Hb thực đo sau quy đổi: {r2['hb']:.1f} g/dL"
        )

        doc.add_paragraph(
            f"Hb sau hiệu chỉnh: {r2['hb_adjusted']:.1f} g/dL"
        )

        doc.add_paragraph(
            f"MCV: {r2['mcv']:.1f} fL"
        )

        doc.add_paragraph(
            f"MCH: {r2['mch']:.1f} pg"
        )

        doc.add_paragraph(
            f"RBC: {r2['rbc']:.2f} T/L"
        )

        doc.add_paragraph(
            f"RDW-CV: {r2['rdw']:.1f}%"
        )

        doc.add_paragraph(
            f"Mentzer Index: {r2['mentzer']:.2f}"
        )

        doc.add_paragraph(
            f"Điểm CBC prototype: {r2['score']}"
        )

        doc.add_paragraph(
            f"Mức nguy cơ Vòng 2: {r2['category']}"
        )

        doc.add_paragraph(
            f"Nhận định: {r2['conclusion']}"
        )

        doc.add_heading(
            "IV. PHÂN TÍCH SƠ BỘ",
            level=2,
        )

        for finding in r2["findings"]:
            doc.add_paragraph(
                f"- {finding}"
            )

        doc.add_heading(
            "V. KHUYẾN NGHỊ",
            level=2,
        )

        for advice in r2["advice"]:
            doc.add_paragraph(
                f"- {advice}"
            )

    else:

        doc.add_heading(
            "III. THEO DÕI",
            level=2,
        )

        doc.add_paragraph(
            "Vòng 2 chưa được mở. Tiếp tục theo dõi sức khỏe "
            "và đánh giá lại nếu xuất hiện bất thường."
        )

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)

    return out


# ============================================================
# HEADER
# ============================================================

# ============================================================
# CỬA SỔ GIẢI THÍCH NGHIÊN CỨU
# ============================================================

@st.dialog("📚 Tìm hiểu về nghiên cứu và cơ sở khoa học", width="large")
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
    st.subheader("🔗 Tài liệu chuyên môn tham khảo")
    st.link_button("WHO 2024 – Guideline on haemoglobin cutoffs", "https://www.who.int/publications/i/item/9789240088542", use_container_width=True)
    st.link_button("GeneReviews – Beta-Thalassemia", "https://www.ncbi.nlm.nih.gov/books/NBK1426/", use_container_width=True)
    st.link_button("ACOG – Carrier Screening for Genetic Conditions", "https://www.acog.org/clinical/clinical-guidance/committee-opinion/articles/2017/03/carrier-screening-for-genetic-conditions", use_container_width=True)
    st.link_button("ACOG – Carrier Screening for Hemoglobinopathies", "https://www.acog.org/womens-health/faqs/carrier-screening-for-hemoglobinopathies", use_container_width=True)


st.markdown("""<div class=\"hematology-brand\"><div class=\"eyebrow\">HEMATOLOGY · COMMUNITY SCREENING RESEARCH</div><div class=\"title\">Hệ thống hỗ trợ sàng lọc Thalassemia</div><div class=\"subtitle\">Sàng lọc ban đầu · Công thức máu · Phân tích huyết học · Khuyến nghị · Điều hướng cơ sở y tế</div></div><div class=\"section-rule\"></div>""", unsafe_allow_html=True)



# ============================================================
# GIỚI THIỆU ĐỀ TÀI — EXPANDER
with st.expander("GIỚI THIỆU ĐỀ TÀI & CƠ SỞ NGHIÊN CỨU · BẤM ĐỂ XEM", expanded=False):
    st.markdown("## LỜI NÓI ĐẦU")
    st.markdown("""
    Thalassemia là nhóm bệnh lý huyết sắc tố di truyền do giảm hoặc mất khả năng tổng hợp một hoặc nhiều chuỗi globin. Phổ biểu hiện rất rộng, từ người mang gen có biểu hiện huyết học nhẹ hoặc gần như không có triệu chứng đến các thể bệnh thiếu máu nặng cần chăm sóc y tế lâu dài. Vì có tính di truyền, việc nhận diện người mang gen không chỉ có ý nghĩa đối với cá nhân mà còn có giá trị trong tư vấn di truyền và dự phòng nguy cơ cho thế hệ sau.

    Trong thực hành sàng lọc, **công thức máu (CBC) và các chỉ số hồng cầu** là những dữ liệu ban đầu có giá trị để nhận diện kiểu hình hồng cầu nhỏ, nhược sắc. Tuy nhiên, CBC không thể tự xác định loại hemoglobinopathy hay biến thể gen. Khi có dấu hiệu nghi ngờ, người tham gia cần được đánh giá tiếp bằng tình trạng sắt, HPLC/điện di hemoglobin và trong những trường hợp phù hợp là xét nghiệm phân tử.

    Từ nhu cầu kết nối giữa **sàng lọc ban đầu – dữ liệu CBC – giải thích kết quả – khuyến nghị – tiếp cận cơ sở y tế**, đề tài **“Hệ thống hỗ trợ sàng lọc Thalassemia trong cộng đồng”** được xây dựng. Hệ thống được định hướng như một công cụ hỗ trợ sức khỏe cộng đồng, không thay thế bác sĩ và không đưa ra chẩn đoán xác định.
    """)

    st.markdown("## 1. LÝ DO CHỌN ĐỀ TÀI")
    st.markdown("""
    Thứ nhất, Thalassemia là bệnh lý di truyền có thể tồn tại âm thầm trong cộng đồng. Người mang gen có thể không có biểu hiện lâm sàng rõ ràng nhưng vẫn có ý nghĩa về mặt di truyền. Do đó, chỉ dựa vào triệu chứng để nhận diện là không đủ; cần có chiến lược sàng lọc phù hợp để phát hiện những trường hợp cần được đánh giá sâu hơn.

    Thứ hai, **CBC là xét nghiệm có tính nền tảng trong thực hành huyết học**. Các thông số như Hb, MCV, MCH, RBC và RDW cung cấp thông tin về số lượng và đặc điểm hình thái của quần thể hồng cầu. Mẫu hình hồng cầu nhỏ, nhược sắc có thể gặp trong Thalassemia nhưng cũng gặp trong thiếu sắt và nhiều nguyên nhân khác. Vì vậy, vấn đề không chỉ là “đọc một con số”, mà là đặt các chỉ số vào đúng bối cảnh sàng lọc.

    Thứ ba, các xét nghiệm chuyên sâu như **HPLC, điện di hemoglobin và xét nghiệm phân tử** có vai trò quan trọng nhưng không phải lúc nào cũng là bước đầu tiên hoặc dễ tiếp cận đối với mọi người dân. Một hệ thống hỗ trợ có thể giúp người sử dụng hiểu dữ liệu CBC đang có, nhận biết khi nào cần đánh giá thêm và tránh tâm lý tự kết luận chỉ từ một chỉ số.

    Thứ tư, việc xây dựng hệ thống theo hướng **không loại người tham gia chỉ dựa trên bảng câu hỏi Vòng 1** giúp giảm nguy cơ bỏ sót người mang gen không có triệu chứng. Điểm Vòng 1 chỉ được sử dụng để phân tầng và đưa ra khuyến nghị; người tham gia vẫn có thể tiếp tục Vòng 2 để nhập CBC nếu có.

    Từ những cơ sở trên, đề tài được lựa chọn với mục tiêu xây dựng một **cầu nối hỗ trợ sàng lọc trong cộng đồng**, bắt đầu từ thông tin nguy cơ và CBC, sau đó định hướng đánh giá chuyên sâu và tiếp cận cơ sở y tế khi cần thiết.
    """)

    st.markdown("## 2. MỤC TIÊU ĐỀ TÀI")
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

    st.markdown("## 3. CƠ SỞ KHOA HỌC")
    st.markdown("### 3.1. Cơ sở bệnh sinh và ý nghĩa của sàng lọc")
    st.markdown("""
    Thalassemia phát sinh từ các biến thể di truyền làm giảm hoặc mất tổng hợp chuỗi globin. Sự mất cân bằng chuỗi globin có thể dẫn đến sinh hồng cầu không hiệu quả và mức độ thiếu máu khác nhau tùy kiểu gen, kiểu hình và thể bệnh. Ở người mang gen, biểu hiện thường nhẹ hơn đáng kể so với các thể bệnh Thalassemia có triệu chứng.

    Vì tình trạng mang gen có thể không biểu hiện rõ trên lâm sàng, **sàng lọc dựa trên dữ liệu huyết học và tiền sử** có vai trò nhận diện những trường hợp nên được xác nhận bằng xét nghiệm chuyên sâu. Đây là lý do hệ thống được thiết kế theo mô hình nhiều tầng thay vì cố gắng đưa ra chẩn đoán từ một bảng điểm duy nhất.
    """)

    st.markdown("### 3.2. Vai trò của công thức máu (CBC)")
    st.markdown("""
    **Hb (hemoglobin):** phản ánh lượng hemoglobin trong máu và là chỉ số quan trọng khi đánh giá thiếu máu. Giá trị Hb cần được diễn giải theo tuổi, giới, thai kỳ, độ cao và các yếu tố liên quan; hệ thống sử dụng hiệu chỉnh độ cao theo hướng dẫn WHO 2024 ở mức prototype.

    **MCV (Mean Corpuscular Volume):** phản ánh thể tích trung bình của hồng cầu. MCV giảm cho thấy hồng cầu nhỏ (microcytosis). Đây là một dấu hiệu thường gặp trong Thalassemia trait nhưng cũng có thể gặp trong thiếu sắt.

    **MCH (Mean Corpuscular Hemoglobin):** phản ánh lượng hemoglobin trung bình trong một hồng cầu. MCH thấp thường đi cùng kiểu hình hồng cầu nhược sắc và có thể hỗ trợ nhận diện trường hợp cần đánh giá thêm.

    **RBC (số lượng hồng cầu):** cung cấp thêm thông tin về số lượng hồng cầu. Trong một số trường hợp microcytosis, số lượng hồng cầu tương đối cao có thể hỗ trợ phân biệt Thalassemia trait với thiếu sắt, nhưng không đủ để xác định chẩn đoán.

    **RDW (Red Cell Distribution Width):** phản ánh mức độ biến thiên kích thước hồng cầu. RDW có thể hỗ trợ diễn giải bối cảnh microcytosis, nhưng không có tính đặc hiệu đủ cao để dùng riêng cho chẩn đoán Thalassemia.
    """)

    st.markdown("### 3.3. Chỉ số Mentzer và các chỉ số phân biệt")
    st.markdown("""
    **Mentzer Index = MCV / RBC** là một chỉ số kinh điển được sử dụng như công cụ định hướng giữa thiếu sắt và Thalassemia trong một số bối cảnh sàng lọc. Tuy nhiên, đây chỉ là **chỉ số hỗ trợ**, không phải xét nghiệm xác nhận. Độ chính xác của các chỉ số phân biệt thay đổi theo quần thể, tuổi, tình trạng thiếu sắt đồng thời và thể bệnh.

    Vì vậy, hệ thống không sử dụng Mentzer đơn độc. Kết quả được đặt cùng MCV, MCH, RBC, RDW và bối cảnh Vòng 1 để tạo ra một mức **nguy cơ sàng lọc prototype**.
    """)

    st.markdown("### 3.4. Vì sao phải đánh giá thiếu sắt?")
    st.markdown("""
    Thiếu sắt là một nguyên nhân phổ biến của thiếu máu hồng cầu nhỏ và có thể tạo ra kiểu hình CBC tương tự Thalassemia. Do đó, khi phát hiện microcytosis/nhược sắc, cần xem xét tình trạng sắt, thường với **ferritin và các xét nghiệm chuyển hóa sắt phù hợp** theo đánh giá chuyên môn.

    Đặc biệt, người có khả năng mang Thalassemia không nên tự sử dụng sắt kéo dài chỉ vì thấy Hb thấp hoặc MCV thấp. Việc bổ sung sắt nên dựa trên bằng chứng thiếu sắt và chỉ định phù hợp.
    """)

    st.markdown("### 3.5. Vì sao CBC chưa đủ để xác định Thalassemia?")
    st.markdown("""
    CBC phản ánh **kiểu hình huyết học**, trong khi Thalassemia là một nhóm rối loạn có nguyên nhân di truyền. Vì vậy, cùng một mẫu hình hồng cầu nhỏ có thể xuất hiện trong nhiều tình trạng khác nhau.

    Khi kết quả sàng lọc gợi ý, bước tiếp theo có thể bao gồm **HPLC hoặc điện di hemoglobin** để đánh giá thành phần hemoglobin. Trong những trường hợp cần xác định biến thể hoặc khi kết quả xét nghiệm huyết học chưa giải thích được tình trạng nghi ngờ, **xét nghiệm phân tử** có thể được cân nhắc theo chỉ định chuyên môn. Với α-thalassemia, xét nghiệm di truyền có thể đặc biệt hữu ích vì các xét nghiệm hemoglobin thường quy không phải lúc nào cũng xác định được tình trạng mang gen.
    """)

    st.markdown("### 3.6. Cơ sở thiết kế Vòng 1 → Vòng 2")
    st.markdown("""
    **Vòng 1** thu thập thông tin nguy cơ và tiền sử để tạo bối cảnh trước khi xem CBC. Điểm số chỉ nhằm phân tầng ban đầu.

    **Vòng 2** tiếp nhận CBC và phân tích các chỉ số hồng cầu. Đây là tầng có tính khách quan hơn vì sử dụng dữ liệu xét nghiệm thực tế của người tham gia.

    **Không có bước nào được thiết kế để “chẩn đoán Thalassemia”.** Kết quả cuối cùng được diễn đạt bằng các mức “thấp – trung bình – cao” về mức độ cần đánh giá thêm, kèm khuyến nghị bước tiếp theo. Người tham gia vẫn có thể cần xét nghiệm chuyên sâu và đánh giá bởi nhân viên y tế.
    """)

    st.markdown("### 3.7. Phạm vi và giới hạn của mô hình")
    st.markdown("""
    Phiên bản hiện tại là **prototype hỗ trợ sàng lọc**, trong đó các điểm số và ngưỡng phân tầng chưa được thẩm định trên một quần thể người Việt Nam đủ lớn. Do đó, kết quả không được sử dụng như tiêu chuẩn chẩn đoán, không thay thế xét nghiệm chuyên sâu và không được dùng để tự quyết định điều trị.

    Mục tiêu của hệ thống là tạo ra một quy trình dễ tiếp cận: **nhận diện nguy cơ → xem CBC → giải thích → khuyến nghị → điều hướng**, sau đó việc xác nhận và xử trí thuộc về cơ sở y tế có thẩm quyền.
    """)

    st.markdown("## 4. TÀI LIỆU THAM KHẢO")
    refs = [
        ("World Health Organization. Guideline on haemoglobin cutoffs to define anaemia in individuals and populations. 2024.",
         "https://www.who.int/publications/i/item/9789240088542"),
        ("World Health Organization. WHO guidelines on best practices in the measurement of haemoglobin. 2024.",
         "https://www.who.int/publications/b/71578"),
        ("GeneReviews®. Beta-Thalassemia. University of Washington, Seattle. Updated February 12, 2026.",
         "https://www.ncbi.nlm.nih.gov/books/NBK1426/"),
        ("American College of Obstetricians and Gynecologists (ACOG). Carrier Screening for Genetic Conditions.",
         "https://www.acog.org/clinical/clinical-guidance/committee-opinion/articles/2017/03/carrier-screening-for-genetic-conditions"),
        ("American College of Obstetricians and Gynecologists (ACOG). Hemoglobinopathies in Pregnancy.",
         "https://www.acog.org/clinical/clinical-guidance/practice-advisory/articles/2022/08/hemoglobinopathies-in-pregnancy"),
        ("American College of Obstetricians and Gynecologists (ACOG). Carrier Screening for Hemoglobinopathies.",
         "https://www.acog.org/womens-health/faqs/carrier-screening-for-hemoglobinopathies"),
        ("Prevalence of Thalassemia in the Vietnamese Population and Building a Clinical Decision Support System for Prenatal Screening for Thalassemia.",
         "https://pmc.ncbi.nlm.nih.gov/articles/PMC10171208/"),
        ("Thalassemia among ethnic minorities in Vietnam: prevalence and molecular characterization.",
         "https://pmc.ncbi.nlm.nih.gov/articles/PMC5496844/"),
        ("A Review of Artificial Intelligence and Machine Learning Applications in Thalassemia Screening and Diagnosis.",
         "https://pmc.ncbi.nlm.nih.gov/articles/PMC10177591/"),
        ("Thalassaemia International Federation (TIF). Guidelines and Reports.",
         "https://thalassaemia.org.cy/publications/tif-publications/guidelines-and-reports/"),
    ]
    for i, (title, url) in enumerate(refs, 1):
        st.markdown(f"**{i}. [{title}]({url})**")
        st.caption(url)

st.markdown("""<div class=\"contact-card\"><div class=\"label\">Liên hệ tác giả</div><div class=\"name\">Nguyễn Đắc Lâu</div><div>Điện thoại: <a href=\"tel:0357930820\">0357930820</a></div><div>Email: <a href=\"mailto:nguyendaclau2309@gmail.com\">nguyendaclau2309@gmail.com</a></div><div class=\"small-note\">Tác giả và người phát triển nguyên mẫu hệ thống hỗ trợ sàng lọc Thalassemia.</div></div>""", unsafe_allow_html=True)

st.divider()
st.markdown("## QUY TRÌNH THAM GIA SÀNG LỌC")
st.caption("Hồ sơ → Đồng ý tham gia → Vòng 1 → CBC/Vòng 2 → Kết quả sàng lọc → Khuyến nghị → Điều hướng cơ sở y tế")


# ============================================================
# ACCESS / SIDEBAR
# ============================================================

with st.sidebar:
    st.header("QUY TRÌNH")
    st.caption("Hồ sơ → Vòng 1 → CBC → Phân tích → Khuyến nghị → Cơ sở y tế")
    if GOOGLE_API_KEY:
        st.caption("🟢 Google Places đã cấu hình")
    else:
        st.caption("ℹ️ Cơ sở y tế vẫn được gợi ý theo danh mục tỉnh/thành.")

render_auth_sidebar()

# ============================================================
# ADMIN CONSOLE
# ============================================================

auth_user = current_auth_user()

if auth_user and st.session_state.get("auth_page") == "🛡️ Quản trị hệ thống":
    render_admin_console(auth_user)
    st.stop()

operator_mode = auth_user is not None
operator_username = auth_user["username"] if auth_user else None

if operator_mode:
    st.success(
        f"👩‍🔬 **Chế độ nhập giúp người tham gia** — thao tác đang được ghi nhận bởi **{auth_user['full_name']}** (#{auth_user['username']})."
    )

# ============================================================
# PATIENT PROFILE
# ============================================================

st.header(
    "👤 THÔNG TIN BỆNH NHÂN"
)
st.caption(
    "Nhập hồ sơ người tham gia. "
    + (
        "Chế độ nhập giúp: hệ thống ghi nhận tài khoản nhân sự thực hiện."
        if operator_mode
        else "Người tham gia tự nhập thông tin của mình."
    )
)

with st.container(border=True):

    p1, p2, p3 = st.columns(3)

    with p1:

        full_name = st.text_input(
            "Họ và tên *",
            placeholder="Nguyễn Văn A",
        )

        birth_date = st.date_input(
            "Ngày sinh *",
            value=date(
                2000,
                1,
                1,
            ),
            min_value=date(
                1900,
                1,
                1,
            ),
            max_value=date.today(),
            format="DD/MM/YYYY",
        )

    with p2:

        phone_raw = st.text_input(
            "Số điện thoại *",
            placeholder="09xxxxxxxx",
        )

        gender = st.selectbox(
            "Giới tính *",
            [
                "Nam",
                "Nữ",
                "Khác",
            ],
        )

    with p3:

        current_address = st.text_input(
            "Địa chỉ hiện tại *",
            placeholder="Số nhà/thôn/tổ/đường",
        )

    st.subheader(
        "📍 Địa giới hành chính hiện tại"
    )

    try:
        admin_hierarchy = load_admin_hierarchy()
        admin_data_ok = True
    except Exception as exc:
        admin_hierarchy = {}
        admin_data_ok = False
        st.error(
            "Không tải được danh mục tỉnh/thành và phường/xã hiện hành. "
            "Vui lòng tải lại trang hoặc thử lại sau."
        )
        st.caption(
            f"Nguồn dữ liệu: {ADMIN_DATA_SOURCE_URL}"
        )
        st.code(str(exc))

    if admin_data_ok:
        provinces_list = sorted(
            admin_hierarchy.keys(),
            key=str.casefold,
        )

        selected_province = st.selectbox(
            "Tỉnh / thành phố *",
            ["— Chọn tỉnh/thành —"] + provinces_list,
            key="profile_province",
        )

        if selected_province == "— Chọn tỉnh/thành —":
            commune_value = ""
            st.selectbox(
                "Phường / xã / đặc khu *",
                ["— Chọn tỉnh/thành trước —"],
                disabled=True,
                key="profile_commune_disabled",
            )
        else:
            available_communes = admin_hierarchy[selected_province]
            commune_choice = st.selectbox(
                "Phường / xã / đặc khu *",
                ["— Chọn phường/xã/đặc khu —"] + available_communes,
                key="profile_commune",
            )
            commune_value = (
                ""
                if commune_choice == "— Chọn phường/xã/đặc khu —"
                else commune_choice
            )

        st.caption(
            "Danh mục địa giới được tải theo cấu trúc 2 cấp hiện hành; "
            "không nhập tay tên phường/xã để tránh sai địa danh."
        )
    else:
        selected_province = ""
        commune_value = ""

    if selected_province and commune_value:
        st.success(
            f"📍 Đã chọn: **{commune_value}, {selected_province}**"
        )

    st.markdown("### 🔐 Đồng ý tham gia sàng lọc và nghiên cứu")
    if operator_mode:
        st.info(
            "Bạn đang nhập giúp người tham gia. Chỉ tiếp tục khi người tham gia đã được giải thích "
            "nội dung, đồng ý cho nghiên cứu sinh sử dụng dữ liệu theo mục đích nghiên cứu/sàng lọc, "
            "và bạn có cơ sở hợp lý để ghi nhận sự đồng ý đó."
        )
    else:
        st.info(
            "Để tiếp tục, người tham gia cần đọc và đồng ý với nội dung dưới đây. "
            "Nếu không đồng ý, hệ thống sẽ **không thực hiện sàng lọc và không lưu hồ sơ/thông tin sức khỏe**."
        )
    with st.container(border=True):
        st.markdown(
            "**Tôi đồng ý cho nghiên cứu sinh sử dụng thông tin cá nhân, "
            "thông tin khảo sát và thông tin sức khỏe/xét nghiệm do tôi cung cấp "
            "cho mục đích sàng lọc cộng đồng Thalassemia và nghiên cứu khoa học.**"
        )
        st.markdown(
            "Tôi hiểu rằng việc tham gia là tự nguyện; kết quả của hệ thống chỉ có "
            "tính chất sàng lọc, không thay thế chẩn đoán của cơ sở y tế; dữ liệu "
            "được lưu phục vụ mục đích nêu trên theo phiên bản chấp thuận của nghiên cứu. "
            "Tôi có thể dừng tham gia bằng cách không tiếp tục sử dụng hệ thống."
        )
        research_consent = st.checkbox(
            "✅ Tôi đã đọc, hiểu và đồng ý tham gia." if not operator_mode
            else "✅ Tôi xác nhận người tham gia đã đọc/được giải thích và đã đồng ý.",
            key="research_consent",
        )
        if operator_mode:
            st.caption(
                f"Người nhập: {auth_user['full_name']} ({auth_user['username']})"
            )
        st.caption(
            f"Phiên bản nội dung chấp thuận: {CONSENT_VERSION}"
        )


phone = normalize_phone(
    phone_raw
)

if (
    research_consent
    and phone
    and valid_vietnam_phone(phone)
    and phone_exists(phone)
):

    st.warning(
        "📌 Số điện thoại này đã tồn tại. "
        "Lưu lại sẽ **ghi đè bằng lần nhập sau cùng**."
    )


if st.button(
    "💾 LƯU / CẬP NHẬT HỒ SƠ",
    type="secondary",
    disabled=not (research_consent and admin_data_ok),
):

    if not research_consent:

        st.error(
            "Bạn cần đồng ý tham gia sàng lọc và nghiên cứu trước khi tiếp tục."
        )

    elif not admin_data_ok:

        st.error(
            "Chưa tải được danh mục địa giới hiện hành nên chưa thể lưu hồ sơ."
        )

    elif not full_name.strip():

        st.error(
            "Vui lòng nhập họ và tên."
        )

    elif not valid_vietnam_phone(phone):

        st.error(
            "Số điện thoại phải là số Việt Nam 10 chữ số."
        )

    elif not current_address.strip():

        st.error(
            "Vui lòng nhập địa chỉ hiện tại."
        )

    elif not selected_province:

        st.error(
            "Vui lòng chọn tỉnh/thành phố."
        )

    elif not commune_value:

        st.error(
            "Vui lòng chọn phường/xã/đặc khu."
        )

    else:

        profile = {
            "phone": phone,
            "full_name": full_name.strip(),
            "birth_date": birth_date.isoformat(),
            "gender": gender,
            "current_address": current_address.strip(),
            "province": selected_province,
            "commune": commune_value,
            "age": calculate_age(
                birth_date
            ),
            "consent_at": datetime.now().isoformat(timespec="seconds"),
            "entry_mode": "assisted" if operator_mode else "self",
            "entered_by_username": operator_username,
        }

        action = upsert_patient(
            profile
        )

        st.session_state[
            "patient_profile"
        ] = profile

        reset_results()

        if action == "updated":
            st.success(
                "✅ Đã cập nhật hồ sơ bằng lần nhập sau cùng."
            )
        else:
            st.success(
                "✅ Đã tạo hồ sơ bệnh nhân."
            )


patient = st.session_state.get(
    "patient_profile"
)

if not st.session_state.get("research_consent", False):
    st.warning(
        "🔒 Bạn chưa đồng ý tham gia. Hệ thống không mở Vòng 1 và không xử lý/lưu dữ liệu sàng lọc."
    )
    st.stop()

if not patient:

    st.info(
        "👆 Hãy lưu hồ sơ bệnh nhân trước khi bắt đầu Vòng 1."
    )

    st.stop()


# ============================================================
# PROFILE SUMMARY
# ============================================================

st.success(
    f"✅ **{patient['full_name']}** · "
    f"{patient['age']} tuổi · "
    f"{patient['phone']} · "
    f"{patient['commune']}, {patient['province']}"
)


# ============================================================
# ROUND 1 QUESTIONS
# ============================================================

st.divider()

st.header(
    "🟦 VÒNG 1 — 20 CÂU HỎI SÀNG LỌC"
)

st.caption("Vòng 1 dùng để ghi nhận các yếu tố cần lưu ý và đưa khuyến nghị. Tất cả người tham gia đều được tiếp tục Vòng 2.")


with st.form("round1_form", clear_on_submit=False):
    with st.container(border=True):

        st.subheader(
            "A. Tiền sử gia đình"
        )

        q1 = st.radio(
            "1. Trong gia đình/dòng họ có người từng được chẩn đoán Thalassemia không?",
            ["Không", "Có", "Không biết"],
            horizontal=True,
        )

        q2 = st.radio(
            "2. Trong gia đình/dòng họ có người từng được thông báo mang gen Thalassemia/hemoglobinopathy không?",
            ["Không", "Có", "Không biết"],
            horizontal=True,
        )

        q3 = st.radio(
            "3. Cha hoặc mẹ bạn có từng được xét nghiệm Thalassemia/hemoglobinopathy không?",
            ["Không", "Có", "Không biết"],
            horizontal=True,
        )

        q4 = st.radio(
            "4. Anh/chị/em ruột có từng được chẩn đoán thiếu máu hoặc hồng cầu nhỏ không?",
            ["Không", "Có", "Không biết"],
            horizontal=True,
        )

        q5 = st.radio(
            "5. Gia đình có trẻ từng phải truyền máu nhiều lần hoặc định kỳ không?",
            ["Không", "Có", "Không biết"],
            horizontal=True,
        )


    with st.container(border=True):

        st.subheader(
            "B. Tiền sử bản thân"
        )

        q6 = st.radio(
            "6. Bạn từng được nhân viên y tế thông báo bị thiếu máu chưa?",
            ["Không", "Có", "Không biết"],
            horizontal=True,
        )

        q7 = st.radio(
            "7. Bạn từng được thông báo MCV thấp/hồng cầu nhỏ chưa?",
            ["Không", "Có", "Không biết"],
            horizontal=True,
        )

        q8 = st.radio(
            "8. Bạn từng được thông báo MCH thấp/hồng cầu nhược sắc chưa?",
            ["Không", "Có", "Không biết"],
            horizontal=True,
        )

        q9 = st.selectbox(
            "9. Bạn từng xét nghiệm Thalassemia/hemoglobinopathy chưa?",
            [
                "Chưa xét nghiệm",
                "Đã xét nghiệm, bình thường",
                "Đã nghi ngờ",
                "Đã xác định mang gen",
                "Không nhớ",
            ],
        )

        q10 = st.radio(
            "10. Bạn từng được chẩn đoán HbE hoặc hemoglobinopathy khác chưa?",
            ["Không", "Có", "Không biết"],
            horizontal=True,
        )

        q11 = st.radio(
            "11. Bản thân từng truyền máu nhiều lần hoặc định kỳ chưa?",
            ["Không", "Có", "Không biết"],
            horizontal=True,
        )

        q12 = st.radio(
            "12. Bạn có tiền sử thiếu máu kéo dài từ nhỏ hoặc từ tuổi thiếu niên không?",
            ["Không", "Có", "Không biết"],
            horizontal=True,
        )


    with st.container(border=True):

        st.subheader(
            "C. Dấu hiệu hỗ trợ"
        )

        q13 = st.radio(
            "13. Bạn có thường xuyên mệt mỏi hoặc giảm khả năng hoạt động không?",
            ["Không", "Có", "Không biết"],
            horizontal=True,
        )

        q14 = st.radio(
            "14. Bạn có thường xuyên hoa mắt/chóng mặt không rõ nguyên nhân không?",
            ["Không", "Có", "Không biết"],
            horizontal=True,
        )

        q15 = st.radio(
            "15. Bạn từng được nhận xét da hoặc niêm mạc nhợt hơn bình thường chưa?",
            ["Không", "Có", "Không biết"],
            horizontal=True,
        )

        q16 = st.radio(
            "16. Bạn từng có vàng da/vàng mắt không rõ nguyên nhân chưa?",
            ["Không", "Có", "Không biết"],
            horizontal=True,
        )

        q17 = st.radio(
            "17. Bạn từng được bác sĩ ghi nhận lách to hoặc gan lách to chưa?",
            ["Không", "Có", "Không biết"],
            horizontal=True,
        )

        q18 = st.radio(
            "18. Bạn từng được bác sĩ lưu ý có biến chứng liên quan bệnh huyết học mạn chưa?",
            ["Không", "Có", "Không biết"],
            horizontal=True,
        )


    with st.container(border=True):

        st.subheader(
            "D. Khả năng tiếp cận xét nghiệm — KHÔNG TÍNH ĐIỂM"
        )

        q19 = st.radio(
            "19. Bạn hiện có CBC trong vòng 6–12 tháng gần đây không?",
            ["Không", "Có", "Không biết"],
            horizontal=True,
        )

        q20 = st.radio(
            "20. Bạn có gặp khó khăn khi đến cơ sở có xét nghiệm chuyên sâu "
            "do khoảng cách, chi phí hoặc thời gian di chuyển không?",
            ["Không", "Có", "Không biết"],
            horizontal=True,
        )

        st.info(
            "Q19–Q20 được lưu để hỗ trợ điều hướng y tế, **không ảnh hưởng "
            "đến điểm nguy cơ Thalassemia**."
        )


    if st.form_submit_button(
        "🔎 ĐÁNH GIÁ VÒNG 1",
        type="primary",
        use_container_width=True,
    ):

        answers = {
            "q1": q1,
            "q2": q2,
            "q3": q3,
            "q4": q4,
            "q5": q5,
            "q6": q6,
            "q7": q7,
            "q8": q8,
            "q9": q9,
            "q10": q10,
            "q11": q11,
            "q12": q12,
            "q13": q13,
            "q14": q14,
            "q15": q15,
            "q16": q16,
            "q17": q17,
            "q18": q18,
            "q19": q19,
            "q20": q20,
        }

        score1, reasons1 = (
            calculate_round1_score(
                answers
            )
        )

        category1, conclusion1 = (
            round1_category(
                score1
            )
        )

        st.session_state[
            "round1_score"
        ] = score1

        st.session_state[
            "round1_reasons"
        ] = reasons1

        # Lưu một lượt sàng lọc hoàn chỉnh Vòng 1. Nếu có tài khoản đăng nhập,
        # ghi nhận người nhập là nhân sự/ quản trị viên đang thực hiện thao tác.
        st.session_state["screening_id"] = create_screening_record(
            patient=patient,
            answers=answers,
            score1=score1,
            category1=category1,
            conclusion1=conclusion1,
            reasons1=reasons1,
            entry_mode="assisted" if operator_mode else "self",
            entered_by_username=operator_username,
        )

        st.session_state[
            "round1_category"
        ] = category1

        st.session_state[
            "round1_conclusion"
        ] = conclusion1

        st.session_state[
            "round1_completed"
        ] = True

        # Vòng 1 mới -> xóa Vòng 2 cũ.
        for key in list(
            st.session_state.keys()
        ):
            if (
                key.startswith("round2_")
                or key.startswith("google_")
            ):
                del st.session_state[key]


# ============================================================
# ROUND 1 RESULT
# ============================================================

if st.session_state.get(
    "round1_completed",
    False,
):

    st.subheader(
        "📋 KẾT QUẢ VÒNG 1"
    )

    score1 = st.session_state[
        "round1_score"
    ]

    category1 = st.session_state[
        "round1_category"
    ]

    conclusion1 = st.session_state[
        "round1_conclusion"
    ]

    reasons1 = st.session_state[
        "round1_reasons"
    ]

    a1, a2 = st.columns(2)
    with a1:
        st.metric("Điểm sàng lọc ban đầu", f"{score1}/{ROUND1_MAX_SCORE}")
    with a2:
        st.metric("Vòng 2", "Được tiếp tục")

    # Vòng 1 chỉ phân tầng và đưa lời khuyên; KHÔNG khóa Vòng 2.
    if category1 == "CAO":
        st.error(f"🔴 **MỨC ĐỘ ĐÁNG LƯU Ý: CAO**\n\n{conclusion1}")
    elif category1 == "TRUNG BÌNH":
        st.warning(f"🟡 **MỨC ĐỘ ĐÁNG LƯU Ý: TRUNG BÌNH**\n\n{conclusion1}")
    else:
        st.success(f"🟢 **MỨC ĐỘ ĐÁNG LƯU Ý: THẤP**\n\n{conclusion1}")

    if reasons1:
        with st.expander("🔎 Những điểm Vòng 1 cần lưu ý", expanded=True):
            for item in reasons1:
                st.write(f"• {item}")
    else:
        st.info("Chưa ghi nhận yếu tố đáng lưu ý từ các câu trả lời có tính điểm.")

    st.markdown("### 🧭 Sau Vòng 1, bạn nên làm gì?")
    guidance = round1_guidance(category1, reasons1)
    for i, item in enumerate(guidance, 1):
        st.write(f"**{i}.** {item}")

    st.info("🩸 **Vòng 2 được mở cho tất cả người tham gia.** Nếu bạn đã có CBC, hãy nhập kết quả để hệ thống phân tích các chỉ số huyết học. Nếu chưa có, bạn có thể thực hiện CBC tại cơ sở y tế phù hợp rồi quay lại nhập kết quả.")
    st.session_state["round2_unlocked"] = True


# ============================================================
# ROUND 2
# ============================================================

if st.session_state.get(
    "round2_unlocked",
    False,
):

    st.divider()

    st.header(
        "🟧 VÒNG 2 — ĐỘ CAO + CBC"
    )

    # --------------------------------------------------------
    # LOCATION + ALTITUDE
    # --------------------------------------------------------

    with st.container(border=True):

        st.subheader(
            "1. Nơi cư trú và độ cao"
        )

        st.write(
            f"**Tỉnh/thành:** {patient['province']}"
        )

        st.write(
            f"**Phường/xã/đặc khu:** {patient['commune']}"
        )

        st.write(
            f"**Địa chỉ hiện tại:** {patient['current_address']}"
        )

        altitude_choice, altitude_adjustment = (
            altitude_selector(
                "round2"
            )
        )

        st.session_state[
            "round2_altitude_choice"
        ] = altitude_choice

        st.session_state[
            "round2_adjustment"
        ] = altitude_adjustment

    # --------------------------------------------------------
    # CBC
    # --------------------------------------------------------

    with st.container(border=True):

        st.subheader(
            "2. Nhập Công thức máu"
        )

        c1, c2, c3, c4, c5 = st.columns(5)

        with c1:

            hb_unit = st.selectbox(
                "Đơn vị Hb",
                ["g/dL", "g/L"],
                key="round2_hb_unit",
            )

            hb_raw = st.number_input(
                "Hb",
                min_value=3.0
                if hb_unit == "g/dL"
                else 30.0,
                max_value=25.0
                if hb_unit == "g/dL"
                else 250.0,
                value=None,
                step=0.1,
                key="round2_hb_raw",
            )

        with c2:

            mcv = st.number_input(
                "MCV (fL)",
                30.0,
                150.0,
                None,
                0.1,
                key="round2_mcv",
            )

        with c3:

            mch = st.number_input(
                "MCH (pg)",
                10.0,
                50.0,
                None,
                0.1,
                key="round2_mch",
            )

        with c4:

            rbc_unit = st.selectbox(
                "Đơn vị RBC",
                [
                    "T/L",
                    "10^12/L",
                    "10^6/µL",
                ],
                key="round2_rbc_unit",
            )

            rbc_raw = st.number_input(
                "RBC",
                1.0,
                10.0,
                None,
                0.1,
                key="round2_rbc_raw",
            )

        with c5:

            rdw = st.number_input(
                "RDW-CV (%)",
                5.0,
                40.0,
                None,
                0.1,
                key="round2_rdw",
            )

        st.caption(
            "Hệ thống tự chuyển Hb → g/dL và RBC → T/L trước khi tính."
        )

        st.caption(
            "Tiêu chí sàng lọc CBC theo Phụ lục ‘Quy trình xét nghiệm sàng lọc, chẩn đoán Thalassemia’ của Hướng dẫn chẩn đoán và điều trị một số bệnh lý huyết học (Bộ Y tế, 2022): MCV <85 fL và/hoặc MCH <28 pg. Đây là tiêu chí sàng lọc, không phải tiêu chuẩn chẩn đoán xác định."
        )


    # --------------------------------------------------------
    # ANALYZE
    # --------------------------------------------------------

    if st.button(
        "🩸 PHÂN TÍCH VÒNG 2",
        type="primary",
        use_container_width=True,
    ):

        required_cbc = {
            "Hb": hb_raw,
            "MCV": mcv,
            "MCH": mch,
            "RBC": rbc_raw,
            "RDW-CV": rdw,
        }
        missing = [name for name, value in required_cbc.items() if value is None]
        if missing:
            st.error("Vui lòng nhập đầy đủ số liệu CBC trước khi phân tích: " + ", ".join(missing) + ".")
            st.stop()

        if rbc_raw <= 0:

            st.error(
                "RBC phải lớn hơn 0."
            )

        else:

            hb = hb_to_g_dl(
                hb_raw,
                hb_unit,
            )

            rbc = rbc_to_t_l(
                rbc_raw,
                rbc_unit,
            )

            adjustment = st.session_state[
                "round2_adjustment"
            ]

            hb_adjusted = (
                hb - adjustment
            )

            score2, mentzer, reasons2 = (
                calculate_round2_score(
                    mcv,
                    mch,
                    rbc,
                    rdw,
                )
            )

            category2, conclusion2 = (
                round2_category(
                    score2,
                    mcv,
                    mch,
                )
            )

            findings, advice = (
                narrative_cbc_advice(
                    hb_adjusted,
                    mcv,
                    mch,
                    rbc,
                    rdw,
                    mentzer,
                )
            )

            st.session_state[
                "round2_score"
            ] = score2

            st.session_state[
                "round2_mentzer"
            ] = mentzer

            st.session_state[
                "round2_category"
            ] = category2

            st.session_state[
                "round2_conclusion"
            ] = conclusion2

            st.session_state[
                "round2_reasons"
            ] = reasons2

            st.session_state[
                "round2_hb"
            ] = hb

            st.session_state[
                "round2_hb_adjusted"
            ] = hb_adjusted

            st.session_state[
                "round2_mcv_result"
            ] = mcv

            st.session_state[
                "round2_mch_result"
            ] = mch

            st.session_state[
                "round2_rbc_result"
            ] = rbc

            st.session_state[
                "round2_rdw_result"
            ] = rdw

            st.session_state[
                "round2_findings"
            ] = findings

            st.session_state[
                "round2_advice"
            ] = advice

            update_screening_round2(
                st.session_state.get("screening_id"),
                {
                    "altitude_choice": st.session_state["round2_altitude_choice"],
                    "adjustment": adjustment,
                    "hb": hb,
                    "hb_adjusted": hb_adjusted,
                    "mcv": mcv,
                    "mch": mch,
                    "rbc": rbc,
                    "rdw": rdw,
                    "mentzer": mentzer,
                    "score": score2,
                    "category": category2,
                    "conclusion": conclusion2,
                    "reasons": reasons2,
                    "findings": findings,
                    "advice": advice,
                },
            )

            st.session_state[
                "round2_completed"
            ] = True


    # --------------------------------------------------------
    # RESULTS
    # --------------------------------------------------------

    if st.session_state.get(
        "round2_completed",
        False,
    ):

        st.subheader(
            "📊 KẾT QUẢ VÒNG 2"
        )

        a, b, c, d = st.columns(4)

        with a:
            st.metric(
                "Mentzer Index",
                f"{st.session_state['round2_mentzer']:.2f}",
            )

        with b:
            st.metric(
                "Điểm CBC",
                str(
                    st.session_state[
                        "round2_score"
                    ]
                ),
            )

        with c:
            st.metric(
                "Hb thực đo",
                f"{st.session_state['round2_hb']:.1f} g/dL",
            )

        with d:
            st.metric(
                "Hb sau hiệu chỉnh",
                f"{st.session_state['round2_hb_adjusted']:.1f} g/dL",
            )

        category2 = st.session_state[
            "round2_category"
        ]

        conclusion2 = st.session_state[
            "round2_conclusion"
        ]

        if category2 == "THẤP":

            st.success(
                f"🟢 **NGUY CƠ VÒNG 2: THẤP**\n\n"
                f"{conclusion2}"
            )

        elif category2 == "TRUNG BÌNH":

            st.warning(
                f"🟡 **NGUY CƠ VÒNG 2: TRUNG BÌNH**\n\n"
                f"{conclusion2}"
            )

        else:

            st.error(
                f"🔴 **NGUY CƠ VÒNG 2: {category2}**\n\n"
                f"{conclusion2}"
            )

        with st.expander(
            "🔎 Các yếu tố từ CBC",
            expanded=True,
        ):

            for item in st.session_state[
                "round2_reasons"
            ]:
                st.write(
                    f"• {item}"
                )

        st.markdown(
            "### 🧠 Phân tích sơ bộ"
        )

        for finding in st.session_state[
            "round2_findings"
        ]:
            st.write(
                f"• {finding}"
            )

        st.markdown(
            "### 💡 Khuyến nghị cho người được sàng lọc"
        )

        for item in st.session_state[
            "round2_advice"
        ]:
            st.write(
                f"→ {item}"
            )

        st.caption(
            "Những nhận định trên chỉ hỗ trợ sàng lọc và phải được "
            "đối chiếu với lâm sàng, xét nghiệm chuyên sâu và nhân viên y tế."
        )

        # ----------------------------------------------------
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

        st.subheader(
            "🏥 CƠ SỞ Y TẾ GỢI Ý"
        )

        facilities = recommended_facilities(patient["province"])

        if facilities:

            st.success(
                f"Đã tìm thấy {len(facilities)} cơ sở ưu tiên trong "
                f"**{patient['province']}**."
            )

            st.caption(
                "Ưu tiên bệnh viện hạng I và bệnh viện tuyến Trung ương/hạng đặc biệt "
                "đang có trong danh mục prototype của tỉnh/thành. "
                "Khi đi khám, người bệnh nên hỏi trước khoa Huyết học/Truyền máu "
                "và khả năng thực hiện xét nghiệm chuyên sâu."
            )

            for i, facility in enumerate(facilities, start=1):
                with st.container(border=True):
                    st.markdown(
                        f"### {i}. {facility['name']}"
                    )
                    st.write(
                        f"**Phân loại:** {facility['tier']}"
                    )
                    st.write(
                        f"**Gợi ý:** {facility['note']}"
                    )
                    st.link_button(
                        "🗺️ Xem vị trí / chỉ đường",
                        facility["maps"],
                    )

        else:

            st.info(
                f"Prototype chưa có danh mục bệnh viện ưu tiên cho **{patient['province']}**. "
                "Bạn có thể dùng Google Maps để tìm bệnh viện hạng I hoặc cơ sở tuyến Trung ương "
                "trong chính tỉnh/thành."
            )

            maps_query = (
                f"bệnh viện hạng I bệnh viện Trung ương {patient['province']}"
            )
            maps_url = (
                "https://www.google.com/maps/search/?api=1&query="
                + requests.utils.quote(maps_query)
            )
            st.link_button(
                "🗺️ Tìm bệnh viện tuyến trên trong tỉnh",
                maps_url,
            )

        st.caption(
            "Danh mục cơ sở dùng cho điều hướng prototype; cần cập nhật/đối soát định kỳ "
            "với nguồn chính thức trước khi dùng trong nghiên cứu hoặc triển khai thực tế."
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
