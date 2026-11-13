import io
import json
import math
import os
import re
import hashlib
import secrets
import time
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from urllib.parse import quote

import requests
import streamlit as st

# ===== CODE BY NGUYEN DAC LAU 24XN1 =====
# Yêu cầu: xem requirements.txt (streamlit>=1.45, psycopg[pool], python-docx, XlsxWriter, openpyxl, pyreadstat, pyreadr...)

# set_page_config PHẢI là lệnh Streamlit đầu tiên.
st.set_page_config(
    page_title="Hệ thống sàng lọc Thalassemia",
    page_icon=":material/bloodtype:",
    layout="wide",
)

GLASS_CSS = r"""
<style>
/* =====================================================================
   LIQUID GLASS UI — kính mờ, viền sáng, nền sắc màu; tự thích ứng sáng/tối
   (dùng biến màu của chính theme Streamlit: --text-color, --background-color)
   ===================================================================== */
:root{
  --lg-accent:#C0394F;          /* đỏ huyết học */
  --lg-accent-dark:#8E2C3A;
  --lg-blue:#4F7CFF;
  --lg-teal:#2BB3A3;
  --lg-text:var(--text-color,#1d1d1f);
  --lg-bg:var(--background-color,#ffffff);
  --lg-fill:color-mix(in srgb,var(--lg-bg) 66%,transparent);
  --lg-fill-soft:color-mix(in srgb,var(--lg-bg) 30%,transparent);
  --lg-fill-strong:color-mix(in srgb,var(--lg-bg) 74%,transparent);
  --lg-pop:color-mix(in srgb,var(--lg-bg) 90%,transparent);
  --lg-stroke:color-mix(in srgb,var(--lg-text) 13%,transparent);
  --lg-muted:color-mix(in srgb,var(--lg-text) 62%,transparent);
  --lg-radius:22px;
  --lg-fx:blur(18px);      /* chỉ dùng cho thanh bên, popup, hộp thoại */
  --lg-fx-sm:none;         /* thẻ/nút/ô nhập/tab: KHÔNG blur (nặng GPU) */
  --lg-spec:inset 0 1px 0 rgba(255,255,255,.55),inset 0 -1px 0 rgba(255,255,255,.08),inset 0 0 0 1px rgba(255,255,255,.08);
  --lg-shadow:0 8px 22px -14px rgba(20,25,45,.30);
}
@media (prefers-color-scheme: dark){
  :root{ --lg-spec:inset 0 1px 0 rgba(255,255,255,.22),inset 0 -1px 0 rgba(255,255,255,.04),inset 0 0 0 1px rgba(255,255,255,.05);
         --lg-shadow:0 8px 22px -14px rgba(0,0,0,.6); }
}

/* ---------- Nền aurora: đặt trực tiếp lên .stApp (KHÔNG đổi position/isolation để không làm vỡ bố cục) ---------- */
.stApp,[data-testid="stApp"]{
  font-family:-apple-system,BlinkMacSystemFont,"SF Pro Display","SF Pro Text","Inter","Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif;
  background:
    radial-gradient(900px 620px at 6% -6%,color-mix(in srgb,var(--lg-accent) 30%,transparent),transparent 62%),
    radial-gradient(820px 580px at 102% 2%,color-mix(in srgb,var(--lg-blue) 26%,transparent),transparent 62%),
    radial-gradient(920px 700px at 72% 108%,color-mix(in srgb,var(--lg-teal) 24%,transparent),transparent 62%),
    radial-gradient(700px 520px at 0% 100%,color-mix(in srgb,#9B6BFF 16%,transparent),transparent 62%),
    var(--lg-bg) !important; }
[data-testid="stAppViewContainer"],[data-testid="stMain"],.main,[data-testid="stHeader"]{ background:transparent !important; }
.main .block-container{ max-width:1180px; padding-top:2rem; padding-bottom:4rem; }
h1,h2,h3,h4{ letter-spacing:-0.02em; }
hr{ border:0 !important; height:1px !important; background:linear-gradient(90deg,transparent,var(--lg-stroke),transparent) !important; }
[data-testid="stCaptionContainer"],.stCaption{ color:var(--lg-muted); }
a{ color:var(--lg-accent); }
::-webkit-scrollbar{ width:10px; height:10px; }
::-webkit-scrollbar-thumb{ background:color-mix(in srgb,var(--lg-text) 22%,transparent); border-radius:999px; border:3px solid transparent; background-clip:padding-box; }

/* ---------- Thẻ kính: container, expander, form, metric, uploader ---------- */
[data-testid="stVerticalBlockBorderWrapper"],
[data-testid="stExpander"],
[data-testid="stMetric"],
[data-testid="stFileUploaderDropzone"],
.contact-card{
  background:var(--lg-fill) !important;
  border:1px solid var(--lg-stroke) !important;
  border-radius:var(--lg-radius) !important;
  box-shadow:var(--lg-spec),var(--lg-shadow);
}
[data-testid="stForm"]{
  background:var(--lg-fill-soft) !important;
  -webkit-backdrop-filter:var(--lg-fx-sm); backdrop-filter:var(--lg-fx-sm);
  border:1px solid var(--lg-stroke) !important;
  border-radius:calc(var(--lg-radius) + 4px) !important;
  box-shadow:var(--lg-spec);
}
[data-testid="stForm"] [data-testid="stVerticalBlockBorderWrapper"]{ background:var(--lg-fill-strong) !important; box-shadow:var(--lg-spec); }
[data-testid="stExpander"] details{ border:0 !important; background:transparent !important; }
[data-testid="stExpander"] summary{ border-radius:var(--lg-radius); }
[data-testid="stExpander"] summary p{ font-weight:700; }
[data-testid="stMetric"]{ padding:14px 18px; }
[data-testid="stMetricValue"]{ font-weight:750; letter-spacing:-0.02em; }
[data-testid="stFileUploaderDropzone"]{ border-style:dashed !important; }
[data-testid="stDataFrame"]{ border-radius:18px; overflow:hidden; border:1px solid var(--lg-stroke); box-shadow:var(--lg-spec); }

/* ---------- Thanh bên nổi ---------- */
[data-testid="stSidebar"]{
  background:var(--lg-fill) !important;
  -webkit-backdrop-filter:var(--lg-fx); backdrop-filter:var(--lg-fx);
  border-right:1px solid var(--lg-stroke); box-shadow:var(--lg-spec); }
[data-testid="stSidebarContent"],[data-testid="stSidebar"]>div{ background:transparent !important; }

/* ---------- Hộp thông báo ---------- */
[data-testid="stAlert"]>div,[data-testid="stAlert"] [data-baseweb="notification"]{
  border-radius:18px !important; border:1px solid var(--lg-stroke) !important;
  -webkit-backdrop-filter:var(--lg-fx-sm); backdrop-filter:var(--lg-fx-sm); box-shadow:var(--lg-spec); }

/* ---------- Nút bấm dạng viên thuốc kính ---------- */
[data-testid^="stBaseButton"],[data-testid^="stBaseLinkButton"],.stButton>button,.stDownloadButton>button{
  border-radius:999px !important; font-weight:600 !important;
  border:1px solid var(--lg-stroke) !important;
  background:var(--lg-fill) !important;
  -webkit-backdrop-filter:var(--lg-fx-sm); backdrop-filter:var(--lg-fx-sm);
  box-shadow:var(--lg-spec),0 6px 16px -8px rgba(20,25,45,.35);
  transition:transform .15s ease,background .15s ease; }
[data-testid^="stBaseButton"]:hover,[data-testid^="stBaseLinkButton"]:hover{
  transform:translateY(-1px); background:var(--lg-fill-strong) !important;
  box-shadow:var(--lg-spec),0 10px 22px -8px rgba(20,25,45,.45); }
[data-testid^="stBaseButton"]:active,[data-testid^="stBaseLinkButton"]:active{ transform:scale(.98); }
[data-testid^="stBaseButton"]:disabled{ opacity:.5; transform:none; }
[data-testid="stBaseButton-primary"],[data-testid="stBaseButton-primaryFormSubmit"],
button[kind="primary"],button[kind="primaryFormSubmit"]{
  background:linear-gradient(135deg,color-mix(in srgb,var(--lg-accent) 86%,#fff),var(--lg-accent-dark)) !important;
  border:1px solid rgba(255,255,255,.38) !important; color:#fff !important;
  box-shadow:inset 0 1px 0 rgba(255,255,255,.55),inset 0 -6px 14px rgba(0,0,0,.18),0 12px 26px -10px var(--lg-accent) !important; }
[data-testid="stBaseButton-primary"] *,[data-testid="stBaseButton-primaryFormSubmit"] *{ color:#fff !important; }
[data-testid="stBaseButton-primary"]:hover,[data-testid="stBaseButton-primaryFormSubmit"]:hover{
  background:linear-gradient(135deg,color-mix(in srgb,var(--lg-accent) 74%,#fff),var(--lg-accent)) !important; }

/* ---------- Ô nhập liệu kính ---------- */
[data-baseweb="input"],[data-baseweb="textarea"],[data-baseweb="select"]>div,[data-testid="stDateInput"] [data-baseweb="input"]{
  background:var(--lg-fill-soft) !important; border:1px solid var(--lg-stroke) !important;
  border-radius:14px !important; box-shadow:inset 0 1px 0 rgba(255,255,255,.18);
  -webkit-backdrop-filter:var(--lg-fx-sm); backdrop-filter:var(--lg-fx-sm); }
[data-baseweb="base-input"],[data-baseweb="base-input"] input,[data-baseweb="textarea"] textarea{ background:transparent !important; }
[data-baseweb="input"]:focus-within,[data-baseweb="textarea"]:focus-within,[data-baseweb="select"]>div:focus-within{
  border-color:var(--lg-accent) !important; box-shadow:0 0 0 3px color-mix(in srgb,var(--lg-accent) 28%,transparent) !important; }
[data-testid="stNumberInput"] button{ background:transparent !important; box-shadow:none !important; backdrop-filter:none; }
[data-baseweb="popover"]>div,[data-baseweb="popover"] [role="listbox"],[data-baseweb="menu"]{
  background:var(--lg-pop) !important; -webkit-backdrop-filter:var(--lg-fx); backdrop-filter:var(--lg-fx);
  border:1px solid var(--lg-stroke); border-radius:16px !important; box-shadow:var(--lg-shadow); }

/* ---------- Tab dạng viên kính ---------- */
[data-baseweb="tab-list"]{ gap:4px; padding:5px; width:fit-content; max-width:100%; border-radius:999px;
  background:var(--lg-fill); border:1px solid var(--lg-stroke); box-shadow:var(--lg-spec);
  -webkit-backdrop-filter:var(--lg-fx-sm); backdrop-filter:var(--lg-fx-sm); }
button[data-baseweb="tab"]{ border-radius:999px !important; padding:6px 16px !important; background:transparent !important; }
button[data-baseweb="tab"][aria-selected="true"]{ background:var(--lg-fill-strong) !important;
  box-shadow:var(--lg-spec),0 4px 12px -4px rgba(20,25,45,.35); }
[data-baseweb="tab-highlight"],[data-baseweb="tab-border"]{ display:none !important; }

/* ---------- Hộp thoại ---------- */
div[role="dialog"]{ background:var(--lg-pop) !important; -webkit-backdrop-filter:var(--lg-fx); backdrop-filter:var(--lg-fx);
  border:1px solid var(--lg-stroke); border-radius:28px !important; box-shadow:var(--lg-shadow),var(--lg-spec); }
[data-testid="stProgress"] [role="progressbar"],[data-testid="stProgress"] [role="progressbar"]>div{ border-radius:999px; }

/* ---------- Khối thương hiệu ---------- */
.hematology-brand{ display:flex; align-items:center; gap:1.2rem; position:relative; overflow:hidden;
  padding:1.35rem 1.7rem; margin-bottom:.4rem; border-radius:30px;
  background:var(--lg-fill); border:1px solid var(--lg-stroke);
  box-shadow:var(--lg-spec),var(--lg-shadow); }
.hematology-brand::after{ content:""; position:absolute; inset:0; pointer-events:none;
  background:linear-gradient(115deg,rgba(255,255,255,.38) 0%,rgba(255,255,255,0) 36%); opacity:.55; }
.hematology-brand .brand-text{ position:relative; z-index:1; min-width:0; }
.brand-orb{ position:relative; z-index:1; flex:none; width:60px; height:60px; border-radius:50%;
  background:radial-gradient(circle at 30% 24%,rgba(255,255,255,.95),rgba(255,255,255,.12) 38%,transparent 40%),
             linear-gradient(135deg,#E7707F,var(--lg-accent-dark));
  box-shadow:inset 0 2px 7px rgba(255,255,255,.65),inset 0 -9px 18px rgba(0,0,0,.28),0 14px 28px -10px var(--lg-accent); }
.brand-orb::before{ content:""; position:absolute; left:50%; top:56%; width:20px; height:20px;
  border-radius:50% 0 50% 50%; transform:translate(-50%,-50%) rotate(-45deg);
  background:rgba(255,255,255,.93); box-shadow:0 2px 6px rgba(0,0,0,.25); }
.hematology-brand .eyebrow{ font-size:.74rem; letter-spacing:.18em; color:var(--lg-muted); font-weight:700; text-transform:uppercase; }
.hematology-brand .title{ line-height:1.12; font-weight:800; margin-top:.2rem; color:var(--lg-text); }
.hematology-brand .title-main{ font-size:clamp(1.45rem,3.1vw,2.3rem); letter-spacing:-0.035em; text-transform:uppercase;
  background:linear-gradient(95deg,var(--lg-text) 10%,color-mix(in srgb,var(--lg-accent) 78%,var(--lg-text)) 90%);
  -webkit-background-clip:text; background-clip:text; -webkit-text-fill-color:transparent; }
.hematology-brand .chips{ display:flex; flex-wrap:wrap; gap:.4rem; margin-top:.7rem; }
.chip{ display:inline-block; padding:.18rem .7rem; font-size:.74rem; font-weight:650; border-radius:999px;
  color:var(--lg-text); background:var(--lg-fill-soft); border:1px solid var(--lg-stroke); box-shadow:var(--lg-spec); }
.section-rule{ height:1px; margin:1.2rem 0; background:linear-gradient(90deg,transparent,var(--lg-stroke),transparent); }
.contact-card{ padding:1rem 1.2rem; margin-top:1rem; }
.contact-card .label{ color:var(--lg-muted); font-size:.74rem; text-transform:uppercase; letter-spacing:.1em; font-weight:700; }
.contact-card .name{ color:var(--lg-text); font-size:1.1rem; font-weight:750; margin:.2rem 0 .55rem; }
.contact-card a{ color:var(--lg-accent); text-decoration:none; font-weight:600; }
.contact-card a:hover{ text-decoration:underline; }
.small-note{ color:var(--lg-muted); font-size:.82rem; }

/* ---------- Huy hiệu mức nguy cơ ---------- */
.risk-badge{ display:inline-block; padding:.2rem .75rem; border-radius:999px; font-size:.76rem; font-weight:700; letter-spacing:.06em;
  border:1px solid var(--lg-stroke); box-shadow:var(--lg-spec); }
.risk-low{ background:color-mix(in srgb,#2E7D32 18%,transparent); color:#2E7D32; }
.risk-mid{ background:color-mix(in srgb,#B26A00 18%,transparent); color:#B26A00; }
.risk-high{ background:color-mix(in srgb,#B3261E 18%,transparent); color:#B3261E; }
@media (prefers-color-scheme: dark){ .risk-low{ color:#7BC47F; } .risk-mid{ color:#E5A64B; } .risk-high{ color:#F2877F; } }

/* ---------- Ẩn khoảng trống do khối <style> ---------- */
[data-testid="stMainBlockContainer"]>[data-testid="stVerticalBlock"]>[data-testid="stElementContainer"]:first-child:has(style){ display:none; }

/* ---------- Tiết kiệm tài nguyên ---------- */
@media (max-width:768px){
  :root{ --lg-radius:18px; --lg-fx:blur(12px) saturate(1.4); --lg-fx-sm:blur(8px); }
  .hematology-brand{ padding:1rem 1.1rem; gap:.8rem; border-radius:24px; }
  .brand-orb{ width:46px; height:46px; } .brand-orb::before{ width:15px; height:15px; }
}
@media (prefers-reduced-motion:reduce){ *{ transition:none !important; } }
@media (prefers-reduced-transparency:reduce){
  :root{ --lg-fx:none; --lg-fx-sm:none; --lg-fill:color-mix(in srgb,var(--lg-bg) 94%,transparent);
         --lg-fill-soft:color-mix(in srgb,var(--lg-bg) 88%,transparent); --lg-fill-strong:var(--lg-bg); }
}
@supports not ((backdrop-filter:blur(1px)) or (-webkit-backdrop-filter:blur(1px))){
  :root{ --lg-fill:color-mix(in srgb,var(--lg-bg) 88%,transparent); --lg-fill-soft:color-mix(in srgb,var(--lg-bg) 80%,transparent); }
}
</style>
"""

# Khi người dùng tắt hiệu ứng kính (máy yếu): bỏ blur + nền aurora, tăng độ đặc để vẫn dễ đọc.
GLASS_OFF_CSS = r"""
<style>
:root{ --lg-fx:none; --lg-fx-sm:none;
  --lg-fill:color-mix(in srgb,var(--lg-bg) 92%,transparent);
  --lg-fill-soft:color-mix(in srgb,var(--lg-bg) 86%,transparent);
  --lg-fill-strong:var(--lg-bg); }
.stApp,[data-testid="stApp"]{ background:var(--lg-bg) !important; }
[data-testid="stVerticalBlockBorderWrapper"],[data-testid="stExpander"],[data-testid="stMetric"],[data-testid="stSidebar"],.hematology-brand,.contact-card{ box-shadow:none !important; }
</style>
"""


def _minify_css(css):
    css = re.sub(r"/\*.*?\*/", "", css, flags=re.S)
    return re.sub(r"\s+", " ", css).strip()


# Có thể tắt hiệu ứng kính ở thanh bên (máy yếu / mạng chậm) → nhẹ và nhanh hơn.
# Công tắc khẩn cấp: thêm ?glass=off vào cuối địa chỉ web để tắt toàn bộ CSS kính.
_glass_on = st.session_state.get("glass_fx", True) and st.query_params.get("glass") != "off"
if st.query_params.get("glass") != "off":
    st.markdown(_minify_css(GLASS_CSS + ("" if _glass_on else GLASS_OFF_CSS)), unsafe_allow_html=True)

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

# Dịch vụ AI đọc phiếu xét nghiệm (tùy chọn). Không có key thì chỉ dùng được Excel/CSV.
ANTHROPIC_API_KEY = _secret("ANTHROPIC_API_KEY")
ANTHROPIC_MODEL = _secret("ANTHROPIC_MODEL", "claude-sonnet-5-5")

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
PAGE_BATCH = ":material/batch_prediction: Phân tích hồ sơ hàng loạt"

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


def _connection_kwargs(dsn=None):
    # prepare_threshold=None: bắt buộc khi đi qua PgBouncer/Supabase pooler (transaction mode),
    # nếu không các kết nối giữ lâu sẽ lỗi "prepared statement already exists".
    if DB_HOST and DB_USER and DB_PASSWORD:
        return {
            "host": DB_HOST, "port": DB_PORT, "dbname": DB_NAME or "postgres",
            "user": DB_USER, "password": DB_PASSWORD, "prepare_threshold": None,
        }
    if dsn:
        return {"conninfo": dsn, "prepare_threshold": None}
    raise RuntimeError(
        "Chưa cấu hình database. Cần DB_HOST, DB_PORT, DB_NAME, DB_USER, DB_PASSWORD "
        "trong Streamlit Secrets, hoặc một DATABASE_URL hợp lệ."
    )


@st.cache_resource(show_spinner=False)
def _get_pool(dsn=None):
    """Pool kết nối dùng chung toàn server: bỏ chi phí TCP/TLS/đăng nhập ở mỗi lần truy vấn."""
    try:
        from psycopg_pool import ConnectionPool
    except ModuleNotFoundError:
        return None  # chưa cài psycopg_pool → quay về kết nối trực tiếp
    kwargs = dict(_connection_kwargs(dsn))
    conninfo = kwargs.pop("conninfo", "")
    opts = dict(conninfo=conninfo, kwargs=kwargs, min_size=1, max_size=6,
                max_idle=240, timeout=8, open=True)
    try:
        return ConnectionPool(check=ConnectionPool.check_connection, **opts)
    except (TypeError, AttributeError):
        return ConnectionPool(**opts)


class PersistentPGConnection:
    """Lớp bọc psycopg: dùng host/port/user/password riêng hoặc DATABASE_URL."""

    def __init__(self, dsn=None):
        if psycopg is None:
            raise RuntimeError(
                "Thiếu thư viện psycopg. Hãy thêm `psycopg[binary]` vào requirements.txt."
            )
        self._pool = None
        pool = _get_pool(dsn)
        try:
            if pool is not None:
                self._conn = pool.getconn()
                self._pool = pool
            else:
                self._conn = psycopg.connect(**_connection_kwargs(dsn))
                self._conn.autocommit = False
        except Exception as exc:
            raise RuntimeError(
                "Không thể kết nối PostgreSQL. Kiểm tra lại cấu hình database trong Secrets."
            ) from exc

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
        if self._pool is not None:
            try:
                self._conn.rollback()  # trả kết nối về pool ở trạng thái sạch
            except Exception:
                pass
            try:
                self._pool.putconn(self._conn)
            except Exception:
                pass
            return
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


@st.cache_resource(show_spinner="Đang khởi tạo cơ sở dữ liệu...")
def _init_database_once(_version):
    """Tạo bảng/migrate đúng 1 lần cho mỗi lần khởi động server (không phải mỗi phiên)."""
    conn = PersistentPGConnection(DATABASE_URL or None)
    try:
        _create_postgres_schema(conn)
        _migrate_local_sqlite_once(conn)
        _sync_postgres_sequences(conn)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return True


def _ensure_schema_once():
    state = _net_state()
    if time.time() - state["db_fail_at"] < 20:  # vừa lỗi → báo ngay, không chờ timeout lại
        raise RuntimeError(state["db_err"] or "Không kết nối được database.")
    try:
        _init_database_once(DB_SCHEMA_READY)
    except Exception as exc:
        state["db_fail_at"], state["db_err"] = time.time(), str(exc)
        raise


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


@st.cache_data(ttl=60, show_spinner=False)
def admin_exists_cached():
    return admin_exists()


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

@st.cache_resource(show_spinner=False)
def _init_admin_once(_username):
    ensure_default_admin()
    return True


try:
    _init_admin_once(ADMIN_USERNAME or "-")
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

        pages = [PAGE_SCREENING, PAGE_BATCH] + ([PAGE_ADMIN] if user["role"] == "admin" else [])
        page = st.sidebar.radio("Khu vực làm việc", pages, key="auth_page_radio")
        if user["role"] != "admin":
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
            if not admin_exists_cached():
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

    if patients:
        if st.button("Chuẩn bị file Excel — toàn bộ lịch sử", icon=":material/table_view:",
                     use_container_width=True):
            with st.spinner("Đang tạo file Excel..."):
                try:
                    st.session_state["admin_excel"] = export_screening_xlsx(
                        patients, list_screening_records_for_staff()
                    ).getvalue()
                except RuntimeError as exc:
                    st.session_state.pop("admin_excel", None)
                    st.error(str(exc))
        if st.session_state.get("admin_excel"):
            st.download_button(
                "Tải file Excel (.xlsx)",
                data=st.session_state["admin_excel"],
                file_name=f"Thalassemia_du_lieu_{date.today().isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
                use_container_width=True,
            )
            st.caption(
                "File Excel gồm 2 sheet: Hồ sơ hiện tại và toàn bộ lịch sử sàng lọc. "
                "Không xuất mật khẩu/tài khoản nhân sự. Bấm 'Chuẩn bị' lại để lấy dữ liệu mới nhất."
            )


# ------------------------------------------------------------
# ADMINISTRATIVE DATA (34 PROVINCES / 3,321 COMMUNES)
# ------------------------------------------------------------

try:
    _BASE_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    _BASE_DIR = os.getcwd()
ADMIN_LOCAL_CACHE = os.path.join(_BASE_DIR, "admin_hierarchy_cache.json")
ADMIN_LOCAL_CACHE_MAX_AGE_DAYS = 30


def _valid_admin_data(data):
    return (isinstance(data, dict) and len(data) == 34
            and sum(len(v) for v in data.values()) == 3321)


def _read_local_admin_cache(ignore_age=False):
    try:
        if not os.path.exists(ADMIN_LOCAL_CACHE):
            return None
        age_days = (time.time() - os.path.getmtime(ADMIN_LOCAL_CACHE)) / 86400
        if not ignore_age and age_days > ADMIN_LOCAL_CACHE_MAX_AGE_DAYS:
            return None
        with open(ADMIN_LOCAL_CACHE, encoding="utf-8") as f:
            data = json.load(f)
        return data if _valid_admin_data(data) else None
    except Exception:
        return None


def _write_local_admin_cache(provinces):
    try:
        tmp = ADMIN_LOCAL_CACHE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(provinces, f, ensure_ascii=False)
        os.replace(tmp, ADMIN_LOCAL_CACHE)
    except Exception:
        pass  # không ghi được (filesystem chỉ đọc) thì bỏ qua


def _parse_admin_payload(payload):
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


@st.cache_resource(show_spinner=False)
def _net_state():
    """Bộ nhớ dùng chung giữa các lần chạy lại: nhớ thời điểm lỗi mạng/DB gần nhất."""
    return {"admin_fail_at": 0.0, "admin_err": "", "db_fail_at": 0.0, "db_err": ""}


@st.cache_resource(ttl=86400, show_spinner="Đang tải danh mục địa giới...")
def load_admin_hierarchy():
    """Danh mục 34 tỉnh/thành, 3.321 đơn vị cấp xã.

    Thứ tự: (1) file cache cục bộ (nhanh, không cần mạng) → (2) tải từ GitHub rồi lưu file cache
    → (3) nếu mạng lỗi thì dùng file cache cũ. cache_resource trả về cùng một đối tượng,
    không sao chép lại mỗi lần rerun như cache_data."""
    cached = _read_local_admin_cache()
    if cached:
        return cached
    state = _net_state()
    stale = _read_local_admin_cache(ignore_age=True)
    # Vừa lỗi mạng cách đây <60 giây → không chờ timeout lần nữa (tránh đơ trang mỗi lần bấm).
    if time.time() - state["admin_fail_at"] < 60:
        if stale:
            return stale
        raise RuntimeError(state["admin_err"] or "Không tải được danh mục địa giới.")
    try:
        response = requests.get(ADMIN_DATA_URL, timeout=6)
        response.raise_for_status()
        provinces = _parse_admin_payload(response.json())
        _write_local_admin_cache(provinces)
        return provinces
    except Exception as exc:
        state["admin_fail_at"], state["admin_err"] = time.time(), str(exc)
        if stale:
            return stale
        raise


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
    from docx import Document  # import chậm → chỉ nạp khi xuất Word
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


@st.cache_data(show_spinner=False, max_entries=20, ttl=600)
def make_word_bytes(patient, r1_score, r1_category, r1_reasons, r2=None):
    """Chỉ dựng file Word khi dữ liệu thay đổi (trước đây dựng lại sau MỖI lần rerun)."""
    return make_word(patient, r1_score, r1_category, r1_reasons, r2).getvalue()


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


# ---------- Câu hỏi Vòng 1: định nghĩa SỚM vì trang phân tích hàng loạt cũng dùng ----------
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


# ============================================================
# BATCH ANALYSIS — hồ sơ ĐẦY ĐỦ nhiều người cùng lúc
#   Vòng 1 (20 câu tầm soát) + CTM 18 chỉ số + sắt/ferritin + HPLC + PCR/di truyền
# ============================================================
# Quy trình: tải file → đọc dữ liệu (bảng hoặc AI) → NGƯỜI DÙNG KIỂM TRA/SỬA
#            → phân tích theo các hàm của từng vòng → xem chi tiết / xuất Excel.
# File tải lên chỉ xử lý trong bộ nhớ, KHÔNG lưu vào database hay đĩa.
# Tối ưu: trang chạy trong st.fragment (tương tác không dựng lại cả app),
#         đọc bảng/AI có cache ngắn hạn theo nội dung file.

import base64
import numbers
import tempfile
import unicodedata
from collections import Counter

import pandas as pd

BATCH_MAX_FILES = 20
BATCH_MAX_MB_TABLE = 30      # bảng / phần mềm thống kê / Word
BATCH_MAX_MB_AI = 10         # PDF / ảnh gửi AI đọc
BATCH_MAX_TABLE_ROWS = 2000  # mỗi file
BATCH_MAX_TOTAL_ROWS = 2000  # mỗi lần tải
AI_FILE_TYPES = {            # AI nhìn trực tiếp (PDF/ảnh)
    ".pdf": "application/pdf", ".png": "image/png", ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg", ".webp": "image/webp",
}
TABLE_EXT = (".xlsx", ".xlsm", ".xls", ".ods", ".csv", ".tsv", ".txt", ".json", ".parquet")
STAT_EXT = (".sav", ".zsav", ".por", ".dta", ".sas7bdat", ".xpt", ".rds", ".rdata", ".rda")
WORD_EXT = (".docx",)
UPLOAD_TYPES = sorted({e.lstrip(".") for e in (*AI_FILE_TYPES, *TABLE_EXT, *STAT_EXT, *WORD_EXT)})

BATCH_YNK = ["Không", "Có", "Không biết"]
BATCH_Q9 = ["Chưa xét nghiệm", "Đã xét nghiệm, bình thường", "Đã nghi ngờ",
            "Đã xác định mang gen", "Không nhớ"]
BATCH_SEX = ["Nam", "Nữ", "Khác"]

# ---------- Danh sách cột (nguồn sự thật duy nhất) ----------
COL_LABEL, COL_SRC, COL_NOTE = "Tên/Mã hồ sơ", "Nguồn", "Ghi chú"
COL_AGE, COL_SEX, COL_ALT = "Tuổi", "Giới tính", "Khoảng độ cao"
COL_MARITAL, COL_PREG = "Tình trạng hôn nhân", "Đang mang thai"
BATCH_MARITAL = ["Độc thân", "Chuẩn bị kết hôn", "Đã kết hôn", "Đã ly hôn/góa", "Khác"]
Q_COLS = [f"Q{i}" for i in range(1, 21)]
CBC_COLS = ["WBC", "NEU%", "LYM%", "MONO%", "EOS%", "BASO%", "RBC", "HGB", "HCT",
            "MCV", "MCH", "MCHC", "RDW-CV", "RDW-SD", "PLT", "MPV", "PDW", "PCT"]
UNIT_COLS = ["Đơn vị HGB", "Đơn vị RBC"]
IRON_COLS = ["Ferritin", "Sắt huyết thanh", "TSAT (%)"]
HB_NUM_COLS = ["HbA (%)", "HbA2 (%)", "HbF (%)", "HbE (%)"]
COL_HPLC = "Kết luận HPLC/điện di"
COL_PCR = "Kết quả PCR/di truyền"
COL_LAB = "Kết luận phiếu XN"

NUM_COLS = [COL_AGE] + CBC_COLS + IRON_COLS + HB_NUM_COLS
TEXT_RESULT_COLS = [COL_HPLC, COL_PCR, COL_LAB]
META_COLS = {COL_LABEL, COL_SRC, COL_AGE, COL_SEX, COL_MARITAL, COL_PREG, COL_ALT, COL_NOTE, *UNIT_COLS}
COMPUTED_COLS = ["Điểm Vòng 1 (tự tính)", "Mức Vòng 1", "HGB g/dL quy đổi", "Hb sau hiệu chỉnh độ cao",
                 "Mentzer (MCV/RBC)", "Tiêu chí CTM BYT 2022", "HbA2 > 3,5%?", "Ferritin < 30?"]
SKIP_SHEETS = {"danhsach", "huongdan", "docao", "readme"}  # sheet phụ của file mẫu
BIRTH_ALIASES = {"ngaysinh", "namsinh", "dob", "birthdate", "birthyear", "ngaythangnamsinh"}
BATCH_COLUMNS = ([COL_LABEL, COL_SRC, COL_AGE, COL_SEX, COL_MARITAL, COL_PREG, COL_ALT] + Q_COLS + CBC_COLS + UNIT_COLS
                 + IRON_COLS + HB_NUM_COLS + [COL_HPLC, COL_PCR, COL_LAB, COL_NOTE])

COLUMN_GROUPS = {
    "Thông tin": [COL_LABEL, COL_SRC, COL_AGE, COL_SEX, COL_MARITAL, COL_PREG, COL_ALT, COL_NOTE],
    "Câu hỏi Vòng 1 (Q1–Q20)": [COL_LABEL] + Q_COLS,
    "Công thức máu (18 chỉ số)": [COL_LABEL] + CBC_COLS + UNIT_COLS,
    "Sắt · HPLC · PCR": [COL_LABEL] + IRON_COLS + HB_NUM_COLS + TEXT_RESULT_COLS,
}

COLUMN_ALIASES = {
    COL_LABEL: {"hotenmahoso", "hoten", "tenmahoso", "ten", "mahoso", "ma", "name", "id", "mabn", "mabenhnhan"},
    COL_AGE: {"tuoi", "age"},
    COL_SEX: {"gioitinh", "gioi", "sex", "gender", "gt"},
    COL_ALT: {"khoangdocao", "docao", "altitude"},
    COL_MARITAL: {"tinhtranghonnhan", "honnhan", "tinhtranggiadinh", "marital", "maritalstatus"},
    COL_PREG: {"mangthai", "dangmangthai", "thaiky", "pregnant", "pregnancy"},
    "WBC": {"wbc", "bachcau", "slbc"},
    "NEU%": {"neu", "neutrophil", "neutrophils", "neupct"},
    "LYM%": {"lym", "lymph", "lymphocyte", "lymphocytes", "lympct"},
    "MONO%": {"mono", "mon", "monocyte", "monocytes", "monopct"},
    "EOS%": {"eos", "eosinophil", "eosinophils", "eospct"},
    "BASO%": {"baso", "bas", "basophil", "basophils", "basopct"},
    "RBC": {"rbc", "hongcau", "slhc", "rbctl"},
    "HGB": {"hgb", "hb", "hemoglobin", "haemoglobin", "huyetsacto", "hbgdl", "hbgl"},
    "HCT": {"hct", "hematocrit", "haematocrit"},
    "MCV": {"mcv", "mcvfl"},
    "MCH": {"mch", "mchpg"},
    "MCHC": {"mchc"},
    "RDW-CV": {"rdwcv", "rdw"},
    "RDW-SD": {"rdwsd"},
    "PLT": {"plt", "tieucau", "platelet", "platelets"},
    "MPV": {"mpv"},
    "PDW": {"pdw"},
    "PCT": {"pct", "plateletcrit"},
    "Đơn vị HGB": {"donvihgb", "donvihb", "hbunit", "hgbunit"},
    "Đơn vị RBC": {"donvirbc", "rbcunit"},
    "Ferritin": {"ferritin", "ferritinngml"},
    "Sắt huyết thanh": {"sathuyetthanh", "serumiron", "iron", "sat", "fe"},
    "TSAT (%)": {"tsat", "transferrinsaturation", "dobaohoatransferrin"},
    "HbA (%)": {"hba"}, "HbA2 (%)": {"hba2"}, "HbF (%)": {"hbf"}, "HbE (%)": {"hbe"},
    COL_HPLC: {"ketluanhplcdiendi", "ketluanhplc", "hplc", "diendi", "ketluandiendi"},
    COL_PCR: {"ketquapcrditruyen", "pcr", "ditruyen", "ketquapcr", "xetnghiemgen", "ketquaditruyen", "gen"},
    COL_LAB: {"ketluanphieuxn", "ketluanphieu", "ketluan"},
    COL_NOTE: {"ghichu", "note", "notes"},
}
for _i in range(1, 21):
    COLUMN_ALIASES[f"Q{_i}"] = {f"q{_i}", f"cau{_i}"}


def _norm(text):
    text = str(text or "").replace("đ", "d").replace("Đ", "D")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]", "", text.lower())


def _s(value):
    """Chuỗi sạch hoặc None (xử lý None/NaN/chuỗi rỗng)."""
    if value is None:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    text = str(value).strip()
    return text if text and text.lower() not in {"nan", "none", "nat"} else None


def _to_float(value):
    if value is None:
        return None
    try:
        if isinstance(value, str):
            value = value.strip().replace(",", ".")
            if not value:
                return None
        number = float(value)
    except (TypeError, ValueError):
        return None
    return None if math.isnan(number) else number


def _is_code(value):
    """Giá trị dạng mã số (1/2/3...) thường gặp trong file SPSS/Stata chưa gắn nhãn."""
    if isinstance(value, bool) or value is None:
        return False
    if isinstance(value, numbers.Number):
        return not (isinstance(value, float) and math.isnan(value))
    return _s(value) in {"0", "1", "2", "3", "8", "9", "99"}


def _norm_choice(value, q9=False):
    s = _s(value)
    if s is None:
        return None
    if isinstance(value, numbers.Number) and not isinstance(value, bool) and float(value).is_integer():
        s = str(int(value))  # 1.0 -> "1"
    n = _norm(s)
    if q9:
        codes = {"1": "Chưa xét nghiệm", "2": "Đã xét nghiệm, bình thường", "3": "Đã nghi ngờ",
                 "4": "Đã xác định mang gen", "5": "Không nhớ"}
        if n in codes:
            return codes[n]
        if "nghi" in n:
            return "Đã nghi ngờ"
        if "mang" in n or "xacdinh" in n:
            return "Đã xác định mang gen"
        if "binhthuong" in n:
            return "Đã xét nghiệm, bình thường"
        if "chua" in n:
            return "Chưa xét nghiệm"
        if "khongnho" in n or "khongbiet" in n:
            return "Không nhớ"
        return None
    # Mã số thường gặp: 1=Có, 2=Không, 3/9=Không biết (cảnh báo người dùng kiểm tra khi đọc từ mã số)
    if s == "?" or n in {"khongbiet", "kb", "unknown", "3", "8", "9", "88", "99"}:
        return "Không biết"
    if n in {"co", "c", "yes", "y", "x", "1", "true", "dung"}:
        return "Có"
    if n in {"khong", "k", "no", "n", "0", "2", "false", "sai"}:
        return "Không"
    return None


def _norm_sex(value):
    s = _s(value)
    if s is None:
        return None
    if isinstance(value, numbers.Number) and not isinstance(value, bool) and float(value).is_integer():
        s = str(int(value))
    n = _norm(s)
    if n in {"nam", "male", "m", "1"}:
        return "Nam"
    if n in {"nu", "female", "f", "2"}:
        return "Nữ"
    return "Khác" if n in {"khac", "other", "3"} else None


def _norm_marital(value):
    s = _s(value)
    if s is None:
        return None
    n = _norm(s)
    if not n or n.isdigit():
        return None
    if "docthan" in n or ("chua" in n and "ket" in n) or "single" in n:
        return "Độc thân"
    if "chuanbi" in n or "sapket" in n or "dinhket" in n:
        return "Chuẩn bị kết hôn"
    if "lyhon" in n or "goa" in n or "lydi" in n or "divorc" in n or "widow" in n:
        return "Đã ly hôn/góa"
    if "kethon" in n or "vochong" in n or "married" in n or "cogiadinh" in n:
        return "Đã kết hôn"
    return "Khác"


def _age_from_birth(v):
    try:
        if v is None:
            return None
        if isinstance(v, numbers.Number) and not isinstance(v, bool):
            y = int(v)
            age = date.today().year - y if 1900 <= y <= date.today().year else None
        else:
            ts = pd.to_datetime(v, errors="coerce", dayfirst=True)
            if pd.isna(ts):
                return None
            t0 = date.today()
            age = t0.year - ts.year - ((t0.month, t0.day) < (ts.month, ts.day))
        return float(age) if age is not None and 0 <= age <= 120 else None
    except Exception:
        return None


def _norm_hb_unit(unit, value):
    u = _norm(unit)
    if "dl" in u:
        return "g/dL"
    if u.endswith("gl"):
        return "g/L"
    if value is not None and value > 30:  # Hb người 5–20 g/dL hay 50–200 g/L
        return "g/L"
    return "g/dL"


def _norm_rbc_unit(unit):
    raw = str(unit or "")
    if "12" in raw:
        return "10^12/L"
    if "6" in raw:
        return "10^6/µL"
    return "T/L"


def _blank_row(label, source, default_alt="<500 m", note=""):
    row = {c: None for c in BATCH_COLUMNS}
    row.update({COL_LABEL: label, COL_SRC: source, COL_ALT: default_alt, COL_NOTE: note,
                "Đơn vị HGB": "g/dL", "Đơn vị RBC": "T/L"})
    return row


def _normalize_row_units(row):
    """Tự quy đổi các đơn vị thường gặp; ghi chú lại để người dùng kiểm tra."""
    notes = []

    def conv(col, cond, fn, text):
        v = _to_float(row.get(col))
        if v is not None and cond(v):
            row[col] = fn(v)
            notes.append(text)

    conv("WBC", lambda v: v > 500, lambda v: v / 1000, "WBC quy đổi /µL → 10^9/L")
    conv("PLT", lambda v: v > 5000, lambda v: v / 1000, "PLT quy đổi /µL → 10^9/L")
    conv("HCT", lambda v: 0 < v < 1, lambda v: v * 100, "HCT quy đổi tỷ lệ → %")
    conv("MCHC", lambda v: v > 100, lambda v: v / 10, "MCHC quy đổi g/L → g/dL")
    conv("RBC", lambda v: v > 100000, lambda v: v / 1_000_000, "RBC quy đổi /µL → T/L")
    if notes:
        old = _s(row.get(COL_NOTE))
        row[COL_NOTE] = "; ".join(([old] if old else []) + notes)
    return row


# ------------------------------------------------------------
# AI: đọc phiếu xét nghiệm PDF/ảnh (CTM + sắt + HPLC + PCR + bảng câu hỏi nếu có)
# ------------------------------------------------------------

AI_PROMPT = """Bạn là công cụ trích xuất số liệu từ tài liệu (ảnh, PDF hoặc văn bản Word) gồm phiếu xét nghiệm (công thức máu CTM/CBC, ferritin/sắt, HPLC/điện di huyết sắc tố, PCR/xét nghiệm gen Thalassemia) và/hoặc phiếu khảo sát tầm soát cộng đồng.

QUY TẮC BẮT BUỘC:
- KHÔNG trích xuất, KHÔNG nhắc lại họ tên, ngày sinh, số điện thoại, địa chỉ, mã bệnh nhân, tên bác sĩ. Với các trường văn bản, chỉ chép phần KẾT QUẢ/KẾT LUẬN và bỏ mọi thông tin định danh.
- Chỉ lấy KẾT QUẢ của người bệnh, KHÔNG lấy khoảng tham chiếu/giá trị bình thường.
- Không thấy thì để null. Tuyệt đối không đoán, không tự tính, không suy diễn.
- Ghi đơn vị đúng như trên phiếu (hgb_unit: "g/dL" hoặc "g/L"; rbc_unit: "T/L", "10^12/L" hoặc "10^6/µL").
- Các chỉ số %: neu_pct, lym_pct, mono_pct, eos_pct, baso_pct là tỷ lệ phần trăm (không lấy số tuyệt đối #).
- Nếu chỉ có RDW-SD thì rdw_cv = null.
- "answers": chỉ điền khi tài liệu là phiếu khảo sát đã trả lời 20 câu tầm soát; q1..q20 nhận "Có", "Không" hoặc "Không biết"; q9 nhận một trong: "Chưa xét nghiệm", "Đã xét nghiệm, bình thường", "Đã nghi ngờ", "Đã xác định mang gen", "Không nhớ". Không phải phiếu khảo sát thì answers = null.
- Thông tin nhân khẩu: chỉ lấy tuổi (age_years, số), giới tính (sex: "Nam"/"Nữ"/"Khác"), tình trạng hôn nhân/gia đình (marital_status: "Độc thân", "Chuẩn bị kết hôn", "Đã kết hôn", "Đã ly hôn/góa", "Khác"), đang mang thai (pregnant: "Có"/"Không"). TUYỆT ĐỐI không lấy tên, ngày sinh đầy đủ, địa chỉ, số điện thoại.
- Ảnh mờ/không phải phiếu xét nghiệm: legible=false và giải thích ngắn trong notes.

Chỉ trả về MỘT đối tượng JSON, không thêm chữ nào khác:
{"wbc":null,"neu_pct":null,"lym_pct":null,"mono_pct":null,"eos_pct":null,"baso_pct":null,
 "rbc":null,"rbc_unit":null,"hgb":null,"hgb_unit":null,"hct":null,"mcv":null,"mch":null,"mchc":null,
 "rdw_cv":null,"rdw_sd":null,"plt":null,"mpv":null,"pdw":null,"pct":null,
 "ferritin":null,"serum_iron":null,"tsat":null,
 "hba":null,"hba2":null,"hbf":null,"hbe":null,
 "hplc_conclusion":null,"pcr_result":null,"lab_conclusion":null,
 "age_years":null,"sex":null,"marital_status":null,"pregnant":null,
 "answers":null,"legible":true,"notes":"ghi chú ngắn bằng tiếng Việt"}"""

AI_NUM_MAP = {
    "wbc": "WBC", "neu_pct": "NEU%", "lym_pct": "LYM%", "mono_pct": "MONO%", "eos_pct": "EOS%",
    "baso_pct": "BASO%", "rbc": "RBC", "hgb": "HGB", "hct": "HCT", "mcv": "MCV", "mch": "MCH",
    "mchc": "MCHC", "rdw_cv": "RDW-CV", "rdw_sd": "RDW-SD", "plt": "PLT", "mpv": "MPV",
    "pdw": "PDW", "pct": "PCT", "ferritin": "Ferritin", "serum_iron": "Sắt huyết thanh",
    "tsat": "TSAT (%)", "hba": "HbA (%)", "hba2": "HbA2 (%)", "hbf": "HbF (%)", "hbe": "HbE (%)",
}
AI_TEXT_MAP = {"hplc_conclusion": COL_HPLC, "pcr_result": COL_PCR, "lab_conclusion": COL_LAB}


def _ai_request(user_content, max_tokens=2000):
    """Gọi Anthropic Messages API; trả về dict JSON hoặc raise RuntimeError."""
    if not ANTHROPIC_API_KEY:
        raise RuntimeError("Chưa cấu hình ANTHROPIC_API_KEY trong Streamlit Secrets.")
    payload = {"model": ANTHROPIC_MODEL, "max_tokens": max_tokens, "temperature": 0,
               "messages": [{"role": "user", "content": user_content}]}
    try:
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"x-api-key": ANTHROPIC_API_KEY, "anthropic-version": "2023-06-01",
                     "content-type": "application/json"},
            json=payload, timeout=120,
        )
    except requests.RequestException as exc:
        raise RuntimeError(f"Không kết nối được dịch vụ AI: {type(exc).__name__}") from exc

    if resp.status_code != 200:
        try:
            detail = resp.json().get("error", {}).get("message", "")
        except Exception:
            detail = ""
        raise RuntimeError(f"Dịch vụ AI trả lỗi {resp.status_code}. {detail[:200]}")

    text = "".join(b.get("text", "") for b in resp.json().get("content", []) if b.get("type") == "text")
    match = re.search(r"\{.*\}", text, flags=re.DOTALL)
    if not match:
        raise RuntimeError("AI không trả về dữ liệu đúng định dạng.")
    try:
        return json.loads(match.group(0))
    except json.JSONDecodeError as exc:
        raise RuntimeError("Không đọc được JSON do AI trả về.") from exc


def ai_extract_report(file_bytes, mime):
    """AI đọc PDF/ảnh phiếu xét nghiệm."""
    block_type = "document" if mime == "application/pdf" else "image"
    return _ai_request([
        {"type": block_type, "source": {"type": "base64", "media_type": mime,
                                        "data": base64.standard_b64encode(file_bytes).decode("ascii")}},
        {"type": "text", "text": AI_PROMPT},
    ])


def ai_extract_report_text(text):
    """AI đọc nội dung chữ của Word/văn bản."""
    return _ai_request([{"type": "text",
                         "text": 'NỘI DUNG TÀI LIỆU:\n"""\n' + text[:60000] + '\n"""\n\n' + AI_PROMPT}])


def ai_map_columns(headers, labels):
    """AI ghép TÊN CỘT (không gửi dữ liệu) vào các trường chuẩn."""
    texts = _question_texts()
    fields = [c for c in BATCH_COLUMNS if c not in (COL_SRC, COL_NOTE)]
    field_desc = "\n".join(f"- {c}" + (f": {texts[c]}" if c in texts else "") for c in fields)
    cols = "\n".join(f"- {h}" + (f" (nhãn biến: {l})" if l else "") for h, l in zip(headers, labels))
    prompt = (
        "Bạn ghép TÊN CỘT của một bộ dữ liệu sàng lọc Thalassemia vào danh sách TRƯỜNG CHUẨN. "
        "Bạn chỉ thấy tên cột và nhãn biến, không có dữ liệu. Chỉ ghép khi chắc chắn "
        "(ví dụ 'hemo_gdl' → HGB; 'fam_thal' → Q1). Không chắc thì trả null.\n\n"
        f"TRƯỜNG CHUẨN:\n{field_desc}\n\nCỘT CẦN GHÉP:\n{cols}\n\n"
        'Chỉ trả về MỘT đối tượng JSON dạng {"tên cột gốc": "trường chuẩn hoặc null"}, không thêm chữ nào khác.'
    )
    return _ai_request([{"type": "text", "text": prompt}], max_tokens=2500)


@st.cache_data(show_spinner=False, max_entries=30, ttl=600)
def ai_extract_report_cached(file_bytes, mime):
    # Cache ngắn (10 phút) theo nội dung file: bấm đọc lại cùng file không tốn thêm chi phí/thời gian.
    return ai_extract_report(file_bytes, mime)


@st.cache_data(show_spinner=False, max_entries=30, ttl=600)
def ai_extract_report_text_cached(text):
    return ai_extract_report_text(text)


@st.cache_data(show_spinner=False, max_entries=50, ttl=3600)
def ai_map_columns_cached(headers, labels):
    return ai_map_columns(headers, labels)


def row_from_ai(label, data, default_alt):
    row = _blank_row(label, "AI (PDF/ảnh/Word)", default_alt)
    row[COL_AGE] = _to_float(data.get("age_years"))
    row[COL_SEX] = _norm_sex(data.get("sex"))
    row[COL_MARITAL] = _norm_marital(data.get("marital_status"))
    preg = _norm_choice(data.get("pregnant"))
    row[COL_PREG] = preg if preg in ("Có", "Không") else None
    for key, col in AI_NUM_MAP.items():
        row[col] = _to_float(data.get(key))
    for key, col in AI_TEXT_MAP.items():
        row[col] = (_s(data.get(key)) or "")[:600] or None
    row["Đơn vị HGB"] = _norm_hb_unit(data.get("hgb_unit"), row["HGB"])
    row["Đơn vị RBC"] = _norm_rbc_unit(data.get("rbc_unit"))

    answers = data.get("answers")
    if isinstance(answers, dict):
        for i in range(1, 21):
            row[f"Q{i}"] = _norm_choice(answers.get(f"q{i}"), q9=(i == 9))

    notes = []
    if data.get("legible") is False:
        notes.append("AI báo phiếu khó đọc — kiểm tra kỹ")
    if data.get("notes"):
        notes.append(str(data["notes"])[:200])
    row[COL_NOTE] = "; ".join(notes) or "AI đọc — cần đối chiếu lại với phiếu gốc"
    return _normalize_row_units(row)


# ------------------------------------------------------------
# QUY ƯỚC MÃ HÓA — người dùng chủ động chọn cách đọc các biến dạng mã số
# ------------------------------------------------------------
# Mỗi nghiên cứu mã hóa khác nhau (giới tính 0/1 hay 1/2; Có/Không là 1/2 hay 1/0...).
# Khi gặp GIÁ TRỊ SỐ, hệ thống chỉ đọc theo bảng quy ước này (không tự đoán).
# Giá trị chữ (Nam, Nữ, Có, Không, Male, Yes...) vẫn được nhận diện tự động.

BLANK_MEANING = "— bỏ trống —"
BIRTH_TARGET = "Ngày/Năm sinh (tự tính tuổi)"
CB_COLS = ["Mã trong file", "Ý nghĩa"]
CB_TITLES = {
    "sex": "Giới tính", "yn": "Câu hỏi Có/Không (Q1–Q8, Q10–Q20)", "q9": "Câu 9 (xét nghiệm Thalassemia trước đây)",
    "marital": "Tình trạng hôn nhân", "preg": "Đang mang thai",
}
_Q9_ORDER = ["Chưa xét nghiệm", "Đã xét nghiệm, bình thường", "Đã nghi ngờ", "Đã xác định mang gen", "Không nhớ"]
CB_PRESETS = {  # mẫu đầu tiên của mỗi nhóm là mặc định
    "sex": {
        "1 = Nam, 2 = Nữ": [("1", "Nam"), ("2", "Nữ")],
        "0 = Nữ, 1 = Nam": [("0", "Nữ"), ("1", "Nam")],
        "0 = Nam, 1 = Nữ": [("0", "Nam"), ("1", "Nữ")],
        "1 = Nữ, 2 = Nam": [("1", "Nữ"), ("2", "Nam")],
        "M = Nam, F = Nữ": [("m", "Nam"), ("f", "Nữ")],
    },
    "yn": {
        "1 = Có, 2 = Không, 3/9 = Không biết": [("1", "Có"), ("2", "Không"), ("3", "Không biết"), ("9", "Không biết")],
        "1 = Có, 0 = Không, 9 = Không biết": [("1", "Có"), ("0", "Không"), ("9", "Không biết")],
        "0 = Không, 1 = Có, 2 = Không biết": [("0", "Không"), ("1", "Có"), ("2", "Không biết")],
        "1 = Không, 2 = Có, 3/9 = Không biết": [("1", "Không"), ("2", "Có"), ("3", "Không biết"), ("9", "Không biết")],
        "1 = Có, 0 = Không, 9 = thiếu (bỏ trống)": [("1", "Có"), ("0", "Không"), ("9", BLANK_MEANING)],
    },
    "q9": {
        "1–5 theo thứ tự lựa chọn": [(str(i + 1), v) for i, v in enumerate(_Q9_ORDER)],
        "0–4 theo thứ tự lựa chọn": [(str(i), v) for i, v in enumerate(_Q9_ORDER)],
    },
    "marital": {
        "1 = Độc thân, 2 = Đã kết hôn, 3 = Đã ly hôn/góa": [("1", "Độc thân"), ("2", "Đã kết hôn"), ("3", "Đã ly hôn/góa")],
        "0 = Độc thân, 1 = Đã kết hôn, 2 = Đã ly hôn/góa": [("0", "Độc thân"), ("1", "Đã kết hôn"), ("2", "Đã ly hôn/góa")],
        "1 = Độc thân, 2 = Chuẩn bị kết hôn, 3 = Đã kết hôn, 4 = Đã ly hôn/góa, 5 = Khác":
            [("1", "Độc thân"), ("2", "Chuẩn bị kết hôn"), ("3", "Đã kết hôn"), ("4", "Đã ly hôn/góa"), ("5", "Khác")],
    },
    "preg": {
        "1 = Có, 2 = Không": [("1", "Có"), ("2", "Không")],
        "1 = Có, 0 = Không": [("1", "Có"), ("0", "Không")],
    },
}
CB_MEANINGS = {"sex": BATCH_SEX, "yn": BATCH_YNK, "q9": BATCH_Q9, "marital": BATCH_MARITAL, "preg": ["Có", "Không"]}
HB_UNIT_CHOICES = {"Tự nhận diện": "auto", "g/dL": "g/dL", "g/L": "g/L"}
RBC_UNIT_CHOICES = {"Tự nhận diện (mặc định T/L)": "auto", "T/L = 10^12/L = 10^6/µL": "T/L",
                    "/µL (vd 4.800.000 → tự chia 1.000.000)": "per_ul"}


def _code_key(value):
    """Chuẩn hóa mã để so khớp: 1, 1.0, "1", " 1 " đều thành "1"; chữ thì viết thường."""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, bool):
        return str(int(value))
    if isinstance(value, numbers.Number):
        f = float(value)
        return str(int(f)) if f.is_integer() else str(f)
    s = _s(value)
    if s is None:
        return None
    try:
        f = float(s.replace(",", "."))
        return str(int(f)) if f.is_integer() else str(f)
    except ValueError:
        return s.lower()


def _is_numeric_key(key):
    try:
        float(key)
        return True
    except (TypeError, ValueError):
        return False


def _table_to_dict(df):
    out = {}
    if df is None:
        return out
    for _, r in df.iterrows():
        code, meaning = _code_key(r.get(CB_COLS[0])), _s(r.get(CB_COLS[1]))
        if code is None or meaning is None:
            continue
        out[code] = None if meaning == BLANK_MEANING else meaning
    return out


def default_codebook():
    cb = {k: {_code_key(c): m for c, m in next(iter(p.values()))} for k, p in CB_PRESETS.items()}
    cb.update({"num_missing": set(), "hb_unit": "auto", "rbc_unit": "auto", "manual": {},
               "hb_unit_label": "Tự nhận diện", "rbc_unit_label": "Tự nhận diện (mặc định T/L)"})
    return cb


def build_codebook():
    """Gom các lựa chọn của người dùng (đang lưu trong session) thành một bộ quy ước."""
    ss = st.session_state
    cb = default_codebook()
    for key in CB_PRESETS:
        df = ss.get(f"cb_{key}_val")
        if df is not None:
            cb[key] = _table_to_dict(df)
    cb["num_missing"] = {f for f in (_to_float(tok) for tok in re.split(r"[;,\s]+", str(ss.get("cb_num_missing") or ""))) if f is not None}
    hb_label = ss.get("cb_hb_unit") or "Tự nhận diện"
    rbc_label = ss.get("cb_rbc_unit") or "Tự nhận diện (mặc định T/L)"
    cb["hb_unit"], cb["hb_unit_label"] = HB_UNIT_CHOICES.get(hb_label, "auto"), hb_label
    cb["rbc_unit"], cb["rbc_unit_label"] = RBC_UNIT_CHOICES.get(rbc_label, "auto"), rbc_label
    mdf = ss.get("cb_manual_val")
    if mdf is not None:
        for _, r in mdf.iterrows():
            src, tgt = _s(r.get("Tên cột trong file")), _s(r.get("Trường chuẩn"))
            if src and tgt:
                cb["manual"][_norm(src)] = tgt
    return cb


def _decode_cat(value, table, text_fn, unmapped, group):
    """Số → tra bảng quy ước (mã lạ → để trống + ghi nhận); chữ → nhận diện tự động."""
    key = _code_key(value)
    if key is None:
        return None
    if key in table:
        return table[key]
    if _is_numeric_key(key):
        unmapped.setdefault(group, Counter())[key] += 1
        return None
    return text_fn(value)


def codebook_summary_text(cb):
    parts = []
    for key, title in CB_TITLES.items():
        items = ", ".join(f"{c}={m or 'bỏ trống'}" for c, m in cb[key].items())
        parts.append(f"{title.split(' (')[0]}: {items or '—'}")
    if cb["num_missing"]:
        parts.append("Mã thiếu số liệu: " + ", ".join(f"{x:g}" for x in sorted(cb["num_missing"])))
    parts.append(f"Hb: {cb['hb_unit_label']}; RBC: {cb['rbc_unit_label']}")
    return " · ".join(parts)


def codebook_summary_df(cb):
    rows = [{"Nhóm": CB_TITLES[k], "Mã": c, "Ý nghĩa": m or "(bỏ trống)"} for k in CB_TITLES for c, m in cb[k].items()]
    rows.append({"Nhóm": "Mã thiếu của số liệu", "Mã": ", ".join(f"{x:g}" for x in sorted(cb["num_missing"])) or "(không)", "Ý nghĩa": "bỏ trống"})
    rows.append({"Nhóm": "Đơn vị Hb", "Mã": cb["hb_unit_label"], "Ý nghĩa": ""})
    rows.append({"Nhóm": "Đơn vị RBC", "Mã": cb["rbc_unit_label"], "Ý nghĩa": ""})
    rows += [{"Nhóm": "Ghép cột thủ công", "Mã": k, "Ý nghĩa": v} for k, v in cb["manual"].items()]
    return pd.DataFrame(rows)


def _codebook_editor(key):
    presets = CB_PRESETS[key]
    names = list(presets)
    base_k, ver_k, val_k, sel_k = f"cb_{key}_base", f"cb_{key}_ver", f"cb_{key}_val", f"cb_{key}_preset"
    ss = st.session_state
    if base_k not in ss:
        ss[base_k] = pd.DataFrame(presets[names[0]], columns=CB_COLS)
        ss[ver_k] = 0

    def _load_preset():
        chosen = st.session_state.get(sel_k)
        if chosen in presets:
            st.session_state[base_k] = pd.DataFrame(presets[chosen], columns=CB_COLS)
            st.session_state[ver_k] += 1

    st.selectbox("Chọn mẫu quy ước", [*names, "Tùy chỉnh (tự sửa bảng bên dưới)"], key=sel_k, on_change=_load_preset)
    ss[val_k] = st.data_editor(
        ss[base_k], num_rows="dynamic", hide_index=True, use_container_width=True,
        key=f"cb_{key}_ed_{ss[ver_k]}",
        column_config={
            CB_COLS[0]: st.column_config.TextColumn("Mã trong file", required=True,
                                                    help="Số hoặc chữ đúng như trong file (ví dụ 0, 1, 2, 9, M, F)"),
            CB_COLS[1]: st.column_config.SelectboxColumn("Ý nghĩa", options=[*CB_MEANINGS[key], BLANK_MEANING], required=True),
        },
    )


def render_codebook_panel():
    with st.expander("Quy ước mã hóa & ghép cột — dùng khi file có mã số (0/1/2…) hoặc tên cột khác mẫu",
                     icon=":material/tune:"):
        st.caption(
            "Mỗi nghiên cứu mã hóa một kiểu (giới tính 0/1 hay 1/2, Có/Không là 1/2 hay 1/0, 9 là thiếu hay không biết…). "
            "Khi gặp **giá trị số**, hệ thống chỉ đọc theo các bảng dưới đây; mã nào không có trong bảng sẽ để trống và được báo lại. "
            "Giá trị chữ (Nam, Nữ, Có, Không, Male, Yes…) vẫn nhận diện tự động. Quy ước áp dụng cho các file Excel/CSV/SPSS/Stata/SAS/R/Word-bảng "
            "và được ghi vào file kết quả để lưu vết."
        )
        tabs = st.tabs(["Giới tính", "Có/Không (Q)", "Tình trạng xét nghiệm Thalassemia", "Hôn nhân", "Thai kỳ", "Số liệu & đơn vị", "Ghép cột thủ công"])
        for tab, key in zip(tabs[:5], ["sex", "yn", "q9", "marital", "preg"]):
            with tab:
                _codebook_editor(key)
        with tabs[2]:
            st.caption("Áp dụng cho cột Q9. Các câu Q khác dùng bảng 'Có/Không (Q)'. Một bảng dùng chung cho Q1–Q8 và Q10–Q20.")
        with tabs[5]:
            st.text_input("Mã thiếu của số liệu (CTM, ferritin, HbA2, tuổi…) — cách nhau bằng dấu phẩy",
                          key="cb_num_missing", placeholder="Ví dụ: -999, 9999",
                          help="Các giá trị này sẽ bị coi là bỏ trống. Chỉ áp dụng cho cột số liệu, KHÔNG áp dụng cho câu hỏi (dùng bảng quy ước ở các tab trước).")
            c1, c2 = st.columns(2)
            c1.selectbox("Đơn vị Hb của file", list(HB_UNIT_CHOICES), key="cb_hb_unit",
                         help="'Tự nhận diện': số >30 coi là g/L. Chọn cố định nếu file dùng một đơn vị thống nhất.")
            c2.selectbox("Đơn vị RBC của file", list(RBC_UNIT_CHOICES), key="cb_rbc_unit")
        with tabs[6]:
            st.caption("Khi tên cột không giống mẫu: gõ đúng tên cột trong file (không phân biệt hoa/thường, dấu) và chọn trường chuẩn tương ứng.")
            fields = [c for c in BATCH_COLUMNS if c != COL_SRC] + [BIRTH_TARGET]
            st.session_state["cb_manual_val"] = st.data_editor(
                pd.DataFrame({"Tên cột trong file": pd.Series(dtype="object"), "Trường chuẩn": pd.Series(dtype="object")}),
                num_rows="dynamic", hide_index=True, use_container_width=True, key="cb_manual_ed",
                column_config={
                    "Tên cột trong file": st.column_config.TextColumn("Tên cột trong file", required=True),
                    "Trường chuẩn": st.column_config.SelectboxColumn("Trường chuẩn", options=fields, required=True),
                },
            )
    return build_codebook()


# ------------------------------------------------------------
# ĐỌC FILE: Excel · CSV · Word · SPSS · Stata · SAS · R · JSON · Parquet
# ------------------------------------------------------------

def _decode_csv(data):
    for enc in ("utf-8-sig", "utf-16", "cp1258", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(data), sep=None, engine="python", encoding=enc)
        except UnicodeError:
            continue
    raise ValueError("Không đọc được bảng văn bản (sai mã hóa ký tự).")


def _docx_tables(data):
    from docx import Document
    out = []
    for ti, table in enumerate(Document(io.BytesIO(data)).tables, 1):
        grid = [[c.text.strip() for c in row.cells] for row in table.rows]
        if len(grid) < 2:
            continue
        seen, header = {}, []
        for i, h in enumerate(grid[0]):
            h = h or f"cot_{i}"
            k = seen.get(h, 0)
            seen[h] = k + 1
            header.append(h if k == 0 else f"{h}_{k}")
        out.append({"sheet": f"Bảng Word {ti}", "df": pd.DataFrame(grid[1:], columns=header), "labels": {}})
    return out


def docx_text(data):
    from docx import Document
    doc = Document(io.BytesIO(data))
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            cells = []
            for c in row.cells:
                txt = c.text.strip()
                if txt and (not cells or cells[-1] != txt):  # ô gộp bị lặp lại
                    cells.append(txt)
            if cells:
                parts.append(" | ".join(cells))
    return "\n".join(parts)


def _read_with_tempfile(ext, data):
    # Thư viện đọc SPSS/R cần đường dẫn file → ghi tạm vào thư mục tạm của hệ điều hành và XÓA NGAY khi xong.
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "data" + ext)
        with open(path, "wb") as f:
            f.write(data)
        if ext in (".sav", ".zsav", ".por"):
            import pyreadstat
            if ext == ".por":
                df, meta = pyreadstat.read_por(path)
            else:
                df, meta = pyreadstat.read_sav(path, apply_value_formats=True)
            labels = {k: v for k, v in zip(meta.column_names, meta.column_labels) if v}
            return [{"sheet": "", "df": df, "labels": labels}]
        import pyreadr
        result = pyreadr.read_r(path)
        return [{"sheet": str(k or ""), "df": v, "labels": {}} for k, v in result.items() if hasattr(v, "columns")]


def read_tables(name, data):
    """Trả về danh sách {sheet, df, labels}. Hỗ trợ nhiều sheet/bảng trong một file."""
    ext = os.path.splitext(name.lower())[1]
    bio = io.BytesIO(data)
    try:
        if ext in (".xlsx", ".xlsm", ".xls", ".ods"):
            return [{"sheet": k, "df": v, "labels": {}} for k, v in pd.read_excel(bio, sheet_name=None).items()]
        if ext in (".csv", ".tsv", ".txt"):
            return [{"sheet": "", "df": _decode_csv(data), "labels": {}}]
        if ext == ".json":
            try:
                df = pd.read_json(bio)
            except ValueError:
                df = pd.json_normalize(json.loads(data.decode("utf-8-sig")))
            return [{"sheet": "", "df": df, "labels": {}}]
        if ext == ".parquet":
            return [{"sheet": "", "df": pd.read_parquet(bio), "labels": {}}]
        if ext in WORD_EXT:
            return _docx_tables(data)
        if ext == ".dta":
            with pd.read_stata(bio, iterator=True) as reader:
                df = reader.read()
                labels = {k: v for k, v in reader.variable_labels().items() if v}
            return [{"sheet": "", "df": df, "labels": labels}]
        if ext in (".sas7bdat", ".xpt"):
            fmt = "xport" if ext == ".xpt" else "sas7bdat"
            try:
                df = pd.read_sas(io.BytesIO(data), format=fmt, encoding="utf-8")
            except UnicodeDecodeError:
                df = pd.read_sas(io.BytesIO(data), format=fmt, encoding="latin-1")
            return [{"sheet": "", "df": df, "labels": {}}]
        if ext in (".sav", ".zsav", ".por", ".rds", ".rdata", ".rda"):
            return _read_with_tempfile(ext, data)
    except ImportError as exc:
        raise RuntimeError(f"Thiếu thư viện để đọc định dạng {ext}: {exc.name or exc}. Hãy thêm vào requirements.txt.") from exc
    raise RuntimeError("Định dạng không được hỗ trợ")


@st.cache_data(show_spinner=False, max_entries=10, ttl=600)
def read_tables_cached(name, data):
    return read_tables(name, data)


def _map_columns(df, labels, ai_ok, notes, source, manual=None):
    """Ghép cột bằng bí danh; cột còn lại (nếu được phép) nhờ AI ghép theo TÊN cột."""
    mapping = {}
    computed_norm = {_norm(c) for c in COMPUTED_COLS}
    manual = manual or {}
    for col in df.columns:  # ưu tiên ghép cột do người dùng chỉ định
        tgt = manual.get(_norm(col))
        if tgt in BATCH_COLUMNS and tgt != COL_SRC and tgt not in mapping:
            mapping[tgt] = col
    for col in df.columns:
        key = _norm(col)
        for target in BATCH_COLUMNS:
            if target in mapping or target == COL_SRC:
                continue
            if key in COLUMN_ALIASES.get(target, set()) or key == _norm(target):
                mapping[target] = col
                break
    used = set(mapping.values())
    unmapped = [c for c in df.columns
                if c not in used and _s(c) and _norm(c) not in computed_norm and _norm(c) not in BIRTH_ALIASES
                and manual.get(_norm(c)) != BIRTH_TARGET]
    if unmapped and mapping:
        if ai_ok and ANTHROPIC_API_KEY:
            by_str = {str(c): c for c in unmapped[:200]}
            try:
                result = ai_map_columns_cached(tuple(by_str), tuple(str(labels.get(c, "")) for c in by_str.values()))
                added = []
                for src, target in result.items():
                    if src in by_str and target in BATCH_COLUMNS and target not in mapping and target != COL_SRC:
                        mapping[target] = by_str[src]
                        added.append(f"{src}→{target}")
                if added:
                    notes.append(f"{source}: AI đã ghép {len(added)} cột theo tên: " + ", ".join(added[:25])
                                 + (" ..." if len(added) > 25 else "") + ". Hãy kiểm tra lại.")
            except Exception as exc:
                notes.append(f"{source}: AI không ghép được cột ({exc}).")
        else:
            notes.append(f"{source}: {len(unmapped)} cột chưa nhận diện (ví dụ: "
                         + ", ".join(str(c) for c in unmapped[:5])
                         + "). Tick xác nhận AI để tự ghép theo tên cột, hoặc đổi tên cột theo file mẫu.")
    return mapping


def rows_from_dataframe(df, labels, filename, sheet, default_alt, ai_ok, notes, cb=None):
    cb = cb or default_codebook()
    source = filename + (f" [{sheet}]" if sheet else "")
    mapping = _map_columns(df, labels, ai_ok, notes, source, cb.get("manual"))
    data_cols = [c for c in mapping if c not in META_COLS]
    if not data_cols:
        return []
    manual = cb.get("manual") or {}
    birth_col = next((c for c in df.columns if _norm(c) in BIRTH_ALIASES or manual.get(_norm(c)) == BIRTH_TARGET), None)

    def preg_text(v):
        c = _norm_choice(v)
        return c if c in ("Có", "Không") else None

    rows, unmapped = [], {}
    for i, rec in df.head(BATCH_MAX_TABLE_ROWS).iterrows():
        row = _blank_row("", f"Bảng: {source}", default_alt)
        for col, src in mapping.items():
            v = rec[src]
            if col in NUM_COLS:
                f = _to_float(v)
                row[col] = None if f is not None and f in cb["num_missing"] else f
            elif col in Q_COLS:
                if col == "Q9":
                    row[col] = _decode_cat(v, cb["q9"], lambda x: _norm_choice(x, q9=True), unmapped, "Câu 9")
                else:
                    row[col] = _decode_cat(v, cb["yn"], _norm_choice, unmapped, "Câu Có/Không")
            elif col == COL_SEX:
                row[col] = _decode_cat(v, cb["sex"], _norm_sex, unmapped, "Giới tính")
            elif col == COL_MARITAL:
                row[col] = _decode_cat(v, cb["marital"], _norm_marital, unmapped, "Hôn nhân")
            elif col == COL_PREG:
                row[col] = _decode_cat(v, cb["preg"], preg_text, unmapped, "Thai kỳ")
            elif col == COL_ALT:
                row[col] = _s(v) if _s(v) in ALTITUDE_OPTIONS else default_alt
            elif col != COL_SRC:
                row[col] = _s(v)
        if row[COL_AGE] is None and birth_col is not None:
            row[COL_AGE] = _age_from_birth(rec[birth_col])
        if all(row[c] is None for c in data_cols):
            continue  # dòng trống (ví dụ dòng chỉ có công thức)
        row[COL_LABEL] = row[COL_LABEL] or f"{filename} · dòng {i + 2}"
        row["Đơn vị HGB"] = cb["hb_unit"] if cb["hb_unit"] != "auto" else _norm_hb_unit(row.get("Đơn vị HGB"), row["HGB"])
        row["Đơn vị RBC"] = _norm_rbc_unit(row.get("Đơn vị RBC"))
        if cb["rbc_unit"] == "per_ul" and row["RBC"] is not None:
            row["RBC"] = row["RBC"] / 1_000_000
        rows.append(_normalize_row_units(row))

    for group, counter in unmapped.items():
        examples = ", ".join(f"{k} (×{n})" for k, n in counter.most_common(6))
        notes.append(f"{source}: {group} có mã KHÔNG nằm trong bảng quy ước → để trống: {examples}. "
                     "Hãy bổ sung vào bảng quy ước rồi đọc lại.")
    if len(df) > BATCH_MAX_TABLE_ROWS:
        notes.append(f"{source}: chỉ đọc {BATCH_MAX_TABLE_ROWS} dòng đầu (file có {len(df)} dòng).")
    return rows


def load_file_rows(name, data, default_alt, ai_ok, cb=None):
    """Đọc một file bất kỳ được hỗ trợ → (danh sách dòng, danh sách ghi chú)."""
    ext = os.path.splitext(name.lower())[1]
    notes = []
    cb = cb or default_codebook()

    if ext in AI_FILE_TYPES:
        if len(data) > BATCH_MAX_MB_AI * 1024 * 1024:
            raise RuntimeError(f"File vượt {BATCH_MAX_MB_AI} MB")
        if not ai_ok:
            raise RuntimeError("Chưa tick xác nhận đồng ý gửi tới AI")
        return [row_from_ai(name, ai_extract_report_cached(data, AI_FILE_TYPES[ext]), default_alt)], notes

    if ext == ".doc":
        raise RuntimeError("Định dạng .doc cũ chưa hỗ trợ — hãy mở và Lưu thành .docx")
    if ext not in (*TABLE_EXT, *STAT_EXT, *WORD_EXT):
        raise RuntimeError("Định dạng không được hỗ trợ")
    if len(data) > BATCH_MAX_MB_TABLE * 1024 * 1024:
        raise RuntimeError(f"File vượt {BATCH_MAX_MB_TABLE} MB")

    tables = []
    try:
        tables = read_tables_cached(name, data)
    except Exception:
        if ext not in (".txt", *WORD_EXT):  # txt/docx còn đường lui: AI đọc văn bản
            raise

    rows = []
    for tb in tables:
        if _norm(tb["sheet"]) in SKIP_SHEETS:
            continue
        rows.extend(rows_from_dataframe(tb["df"], tb["labels"], name, tb["sheet"], default_alt, ai_ok, notes, cb))
    if rows:
        return rows, notes

    if ext in (".txt", *WORD_EXT):
        text = docx_text(data) if ext in WORD_EXT else data.decode("utf-8-sig", errors="replace")
        if not text.strip():
            raise RuntimeError("Tài liệu không có nội dung chữ")
        if not ai_ok:
            raise RuntimeError("Tài liệu dạng phiếu cần AI đọc — hãy tick xác nhận gửi nội dung tới AI")
        return [row_from_ai(name, ai_extract_report_text_cached(text), default_alt)], notes

    raise RuntimeError("Không nhận diện được cột dữ liệu nào (Q1–Q20, CTM, ferritin, HbA2, PCR...)")


# ------------------------------------------------------------
# PHÂN TÍCH
# ------------------------------------------------------------

RISK_RANK = {"THẤP": 1, "TRUNG BÌNH": 2, "CAO": 3}


def who_anemia_flag(hb_adj, age, sex, pregnant=False):
    """Ngưỡng Hb theo WHO 2024 (mực nước biển; Hb đã hiệu chỉnh độ cao). Thai kỳ: 11,0 g/dL
    (tam cá nguyệt 1 và 3; tam cá nguyệt 2 là 10,5)."""
    if age is None:
        return ""
    suffix = ""
    if 0.5 <= age < 5:
        cutoff = 11.0
    elif 5 <= age < 12:
        cutoff = 11.5
    elif 12 <= age < 15:
        cutoff = 12.0
    elif age >= 15 and pregnant:
        cutoff, suffix = 11.0, " — thai kỳ; TCN 2: 10,5"
    elif age >= 15 and sex == "Nam":
        cutoff = 13.0
    elif age >= 15 and sex == "Nữ":
        cutoff = 12.0
    else:
        return ""
    return (f"Có (<{cutoff:g} g/dL{suffix})" if hb_adj < cutoff else f"Không (≥{cutoff:g} g/dL{suffix})")


def integrated_next_steps(cat1, cat2, has_cbc, ferritin, hba2, hbe, hplc_done, pcr_done, anything,
                          marital=None, pregnant=False):
    """Gợi ý bước tiếp theo — chỉ là hỗ trợ sàng lọc, bám các ngưỡng BYT 2022 đã dùng trong app."""
    steps = []
    iron_low = ferritin is not None and ferritin < 30

    if cat1 == "CAO" and not has_cbc:
        steps.append("Vòng 1 mức cao nhưng chưa có CTM đầy đủ: nên ưu tiên thực hiện/nhập công thức máu.")
    if cat2 in ("TRUNG BÌNH", "CAO"):
        if ferritin is None:
            steps.append("CTM đạt tiêu chí sàng lọc: nên đánh giá tình trạng sắt (ferritin) để phân biệt với thiếu sắt.")
        elif iron_low:
            steps.append("Ferritin <30 ng/mL: gợi ý có thiếu sắt (có thể đi kèm Thalassemia). Thiếu sắt có thể làm giảm HbA2 "
                         "nên cần bác sĩ đối chiếu trước khi diễn giải HPLC; không tự bổ sung sắt kéo dài khi chưa có chỉ định.")
        elif not hplc_done:
            steps.append("CTM đạt tiêu chí sàng lọc và ferritin chưa thấp: nên làm HPLC/điện di huyết sắc tố.")
        if hplc_done and not pcr_done and hba2 is not None and hba2 <= 3.5 and not (hbe and hbe > 0) and not iron_low:
            steps.append("CTM gợi ý nhưng HbA2 không tăng: cần lưu ý alpha-thalassemia hoặc thể khác (điện di thường không phát hiện được); "
                         "bác sĩ có thể cân nhắc xét nghiệm gen/PCR.")
    if hba2 is not None and hba2 > 3.5:
        steps.append("HbA2 >3,5%: gợi ý Beta-thalassemia; nên gặp bác sĩ Huyết học, cân nhắc xét nghiệm gen và tư vấn di truyền "
                     "(đặc biệt khi chuẩn bị kết hôn/mang thai).")
    if hbe is not None and hbe > 0:
        steps.append("Có HbE: cần bác sĩ Huyết học đánh giá hemoglobinopathy và tư vấn di truyền.")
    if cat1 == "CAO" and cat2 == "THẤP":
        steps.append("Vòng 1 mức cao nhưng CTM chưa đạt tiêu chí: không loại trừ mang gen; nên trao đổi với bác sĩ nếu tiền sử gia đình rõ.")
    positive = (cat2 in ("TRUNG BÌNH", "CAO") or cat1 == "CAO" or (hba2 is not None and hba2 > 3.5)
                or (hbe is not None and hbe > 0))
    if positive and (pregnant or marital == "Chuẩn bị kết hôn"):
        steps.append("Có yếu tố gợi ý Thalassemia và đang chuẩn bị kết hôn/mang thai: nên được tư vấn di truyền sớm "
                     "và cân nhắc xét nghiệm cho người bạn đời.")
    if iron_low and cat2 not in ("TRUNG BÌNH", "CAO"):
        steps.append("Ferritin <30 ng/mL: gợi ý thiếu sắt (kể cả khi CTM chưa đạt tiêu chí sàng lọc); nên trao đổi với bác sĩ "
                     "về đánh giá và bổ sung sắt theo chỉ định.")
    if pcr_done:
        steps.append("Có kết quả PCR/di truyền: hệ thống không tự diễn giải; cần bác sĩ Huyết học/tư vấn di truyền xem phiếu gốc.")
    if not steps and anything:
        steps.append("Chưa có dấu hiệu cần thêm bước từ dữ liệu đã nhập; tiếp tục theo dõi và đối chiếu với bác sĩ khi có triệu chứng hoặc nhu cầu tư vấn.")
    return steps


def analyze_batch_row(idx, rec):
    label = _s(rec.get(COL_LABEL)) or f"Hồ sơ {idx + 1}"
    age = _to_float(rec.get(COL_AGE))
    sex = _s(rec.get(COL_SEX))
    marital = _s(rec.get(COL_MARITAL))
    pregnant = _s(rec.get(COL_PREG)) == "Có"
    warnings, levels, det = [], [], {}
    out = {"Hồ sơ": label, "Nguồn": _s(rec.get(COL_SRC)) or "",
           "Tuổi": age if age is not None else "", "Giới tính": sex or "",
           "Hôn nhân": marital or "", "Mang thai": "Có" if pregnant else (_s(rec.get(COL_PREG)) or "")}

    # ---------------- Vòng 1 ----------------
    answers = {f"q{i}": _s(rec.get(f"Q{i}")) for i in range(1, 21)}
    answered = sum(1 for v in answers.values() if v)
    cat1 = None
    if answered:
        score1, reasons1 = calculate_round1_score(answers)
        cat1, concl1 = round1_category(score1)
        levels.append(cat1)
        det["r1"] = {"score": score1, "category": cat1, "conclusion": concl1,
                     "reasons": reasons1, "answered": answered}
        out["Vòng 1 - điểm"], out["Vòng 1 - nguy cơ"] = f"{score1}/{ROUND1_MAX_SCORE}", cat1
        if answered < 20:
            warnings.append(f"Vòng 1 mới trả lời {answered}/20 câu — điểm có thể thấp hơn thực tế")

    # ---------------- CTM 18 chỉ số / Vòng 2 ----------------
    nums = {c: _to_float(rec.get(c)) for c in CBC_COLS}
    n_cbc = sum(v is not None for v in nums.values())
    out["Số chỉ số CTM"] = f"{n_cbc}/18" if n_cbc else ""
    cat2, has_cbc = None, False
    if n_cbc:
        core = ["HGB", "MCV", "MCH", "RBC", "RDW-CV"]
        missing = [c for c in core if nums[c] is None]
        if missing:
            warnings.append("CTM thiếu " + ", ".join(missing) + " — chưa chạy Vòng 2")
        else:
            try:
                hb = hb_to_g_dl(nums["HGB"], _s(rec.get("Đơn vị HGB")) or "g/dL")
                rbc = rbc_to_t_l(nums["RBC"], _s(rec.get("Đơn vị RBC")) or "T/L")
            except ValueError as exc:
                hb = rbc = None
                warnings.append(str(exc))
            if hb is not None:
                mcv, mch, rdw = nums["MCV"], nums["MCH"], nums["RDW-CV"]
                limits = [("HGB", hb, 1, 25), ("MCV", mcv, 30, 150), ("MCH", mch, 10, 50),
                          ("RBC", rbc, 1, 10), ("RDW-CV", rdw, 5, 40)]
                bad = [f"{n}={v:g} ngoài khoảng {lo}–{hi}" for n, v, lo, hi in limits if not lo <= v <= hi]
                if bad:
                    warnings.append("; ".join(bad) + " (kiểm tra đơn vị/số liệu) — chưa chạy Vòng 2")
                else:
                    alt_choice = rec.get(COL_ALT) if rec.get(COL_ALT) in ALTITUDE_OPTIONS else "<500 m"
                    hb_adj = hb - ALTITUDE_OPTIONS[alt_choice]
                    score2, mentzer, reasons2 = calculate_round2_score(mcv, mch, rbc, rdw)
                    cat2, concl2 = round2_category(score2, mcv, mch)
                    findings, advice = narrative_cbc_advice(hb_adj, mcv, mch, rbc, rdw, mentzer)
                    mchc = nums["MCHC"]
                    if mchc is not None and mchc < 32:
                        findings.append("MCHC <32 g/dL — gợi ý nhược sắc (đối chiếu khoảng tham chiếu của labo).")
                    anemia = who_anemia_flag(hb_adj, age, sex, pregnant)
                    has_cbc = True
                    levels.append(cat2)
                    det["r2"] = {"category": cat2, "conclusion": concl2, "reasons": reasons2,
                                 "findings": findings, "advice": advice, "altitude": alt_choice,
                                 "score": score2, "mentzer": mentzer}
                    out.update({
                        "Hb (g/dL)": round(hb, 2), "Hb sau hiệu chỉnh": round(hb_adj, 2),
                        "Thiếu máu (WHO 2024)": anemia, "MCV": mcv, "MCH": mch,
                        "MCHC": mchc if mchc is not None else "", "RDW-CV": rdw,
                        "RBC (T/L)": rbc, "Mentzer": mentzer, "Điểm CBC": score2,
                        "Vòng 2 - nguy cơ": cat2,
                    })
        det["cbc_values"] = {c: v for c, v in nums.items() if v is not None}

    # ---------------- Sắt · HPLC · PCR (Vòng 3) ----------------
    hba, hba2, hbf, hbe = (_to_float(rec.get(c)) for c in HB_NUM_COLS)
    ferritin, iron, tsat = (_to_float(rec.get(c)) for c in IRON_COLS)
    hplc_text, pcr_text, lab_text = _s(rec.get(COL_HPLC)), _s(rec.get(COL_PCR)), _s(rec.get(COL_LAB))
    num3 = [hba, hba2, hbf, hbe, ferritin, iron, tsat]
    has3 = any(v is not None for v in num3) or any((hplc_text, pcr_text, lab_text))
    flag3, notes3 = None, []
    hplc_done = any(v is not None for v in (hba, hba2, hbf, hbe)) or bool(hplc_text)
    if has3:
        for name, v, hi in [("HbA", hba, 100), ("HbA2", hba2, 30), ("HbF", hbf, 100),
                            ("HbE", hbe, 100), ("Ferritin", ferritin, 10000)]:
            if v is not None and not 0 <= v <= hi:
                warnings.append(f"{name}={v:g} ngoài khoảng 0–{hi}")
        combined_text = " ".join(filter(None, [hplc_text, pcr_text, lab_text])).lower()
        notes3 = round3_byt_interpretation({
            "hba": hba, "hba2": hba2, "hbf": hbf, "hbe": hbe, "ferritin": ferritin,
            "transferrin_saturation": tsat, "genetic_result": pcr_text,
            "lab_conclusion": " ".join(filter(None, [hplc_text, lab_text])),
        })
        marker = any(m in combined_text for m in ("hbh", "hb bart", "hbcs", "constant spring"))
        if (hba2 is not None and hba2 > 3.5) or (hbe is not None and hbe > 0) or marker:
            flag3 = "CAO"
        elif hbf is not None and hbf >= 1:
            flag3 = "TRUNG BÌNH"
        elif any(v is not None for v in (hba, hba2, hbf, hbe)):
            flag3 = "THẤP"
        if flag3:
            levels.append(flag3)
        det["r3"] = {"flag": flag3, "notes": notes3}
        out.update({
            "Ferritin": ferritin if ferritin is not None else "",
            "HbA2 (%)": hba2 if hba2 is not None else "",
            "HbE (%)": hbe if hbe is not None else "",
            "Vòng 3 - mức lưu ý": flag3 or "",
            "Vòng 3 - đối chiếu BYT": " | ".join(notes3),
        })

    anything = bool(answered or n_cbc or has3)
    steps = integrated_next_steps(cat1, cat2, has_cbc, ferritin, hba2, hbe, hplc_done, bool(pcr_text), anything,
                                  marital, pregnant)
    det["steps"], det["warnings"] = steps, warnings
    out["Mức lưu ý cao nhất"] = max(levels, key=RISK_RANK.get) if levels else ""
    out["Bước tiếp theo"] = " ".join(steps)
    out["Trạng thái"] = "OK" if levels or has3 else "Không đủ dữ liệu"
    out["Cảnh báo"] = "; ".join(filter(None, [_s(rec.get(COL_NOTE))] + warnings))
    out["_detail"] = det
    return out


# ------------------------------------------------------------
# XUẤT EXCEL / FILE MẪU CÓ CÔNG THỨC
# ------------------------------------------------------------

def results_to_excel(results_df, input_df, codebook_df=None):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        results_df.to_excel(writer, sheet_name="Ket_qua", index=False)
        input_df.to_excel(writer, sheet_name="Du_lieu_dau_vao", index=False)
        if codebook_df is not None:
            codebook_df.to_excel(writer, sheet_name="Quy_uoc_ma_hoa", index=False)
        wb = writer.book
        head = wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "text_wrap": True, "valign": "vcenter"})
        wide = {"Bước tiếp theo", "Vòng 3 - đối chiếu BYT", "Cảnh báo"}
        sheets = [("Ket_qua", results_df), ("Du_lieu_dau_vao", input_df)]
        if codebook_df is not None:
            sheets.append(("Quy_uoc_ma_hoa", codebook_df))
        for sheet, df in sheets:
            ws = writer.sheets[sheet]
            for c, name in enumerate(df.columns):
                ws.write(0, c, name, head)
                ws.set_column(c, c, 60 if name in wide else 16)
            ws.freeze_panes(1, 1)
        ws = writer.sheets["Ket_qua"]
        for name in ("Vòng 1 - nguy cơ", "Vòng 2 - nguy cơ", "Vòng 3 - mức lưu ý", "Mức lưu ý cao nhất"):
            if name in results_df.columns:
                col = list(results_df.columns).index(name)
                for text, color in (("THẤP", "#C8E6C9"), ("TRUNG BÌNH", "#FFE0B2"), ("CAO", "#FFCDD2")):
                    ws.conditional_format(1, col, max(len(results_df), 1), col, {
                        "type": "cell", "criteria": "==", "value": f'"{text}"',
                        "format": wb.add_format({"bg_color": color}),
                    })
    buf.seek(0)
    return buf


def _question_texts():
    texts = {"Q9": "9. Bạn từng xét nghiệm Thalassemia/hemoglobinopathy chưa?"}
    for _title, questions in ROUND1_SECTIONS:
        for qkey, qtext in questions:
            if qtext:
                texts[qkey.upper()] = qtext
    return texts


@st.cache_data(show_spinner=False)
def build_template_xlsx():
    """File mẫu ĐẦY ĐỦ có CÔNG THỨC: điểm Vòng 1, quy đổi Hb, hiệu chỉnh độ cao, Mentzer, tiêu chí BYT 2022."""
    import xlsxwriter
    from xlsxwriter.utility import xl_col_to_name

    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    head = wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "text_wrap": True, "valign": "vcenter", "align": "center"})
    head_q = wb.add_format({"bold": True, "bg_color": "#FDE9D9", "border": 1, "text_wrap": True, "valign": "vcenter", "align": "center"})
    head_c = wb.add_format({"bold": True, "bg_color": "#E8E8E8", "border": 1, "text_wrap": True, "valign": "vcenter", "align": "center"})
    cell = wb.add_format({"border": 1})
    calc = wb.add_format({"border": 1, "bg_color": "#F5F5F5"})

    inputs = ([COL_LABEL, COL_AGE, COL_SEX, COL_MARITAL, COL_PREG, COL_ALT] + Q_COLS + CBC_COLS + UNIT_COLS
              + IRON_COLS + HB_NUM_COLS + [COL_HPLC, COL_PCR, COL_LAB, COL_NOTE])
    computed = COMPUTED_COLS
    col = {name: xl_col_to_name(i) for i, name in enumerate(inputs)}
    n_in = len(inputs)

    ws = wb.add_worksheet("Du_lieu")
    texts = _question_texts()
    for c, name in enumerate(inputs):
        ws.write(0, c, name, head_q if name in Q_COLS else head)
        if name in texts:
            ws.write_comment(0, c, texts[name], {"x_scale": 3, "y_scale": 1.2})
    for c, name in enumerate(computed, start=n_in):
        ws.write(0, c, name, head_c)
    ccol = {name: xl_col_to_name(n_in + i) for i, name in enumerate(computed)}

    weights = {1: 3, 2: 3, 3: 1, 4: 2, 5: 2, 6: 1, 7: 2, 8: 1, 10: 2, 11: 1, 12: 2,
               13: 1, 14: 1, 15: 1, 16: 1, 17: 2, 18: 1}
    n_rows = 200
    for r in range(2, n_rows + 2):  # dòng Excel 2..201
        for c in range(n_in):
            ws.write_blank(r - 1, c, None, cell)
        terms = [f'{w}*({col[f"Q{q}"]}{r}="Có")' for q, w in weights.items()]
        terms.append(f'IF({col["Q9"]}{r}="Đã nghi ngờ",2,IF({col["Q9"]}{r}="Đã xác định mang gen",4,0))')
        score_f = (f'=IF(COUNTA({col["Q1"]}{r}:{col["Q20"]}{r})=0,"",'
                   f'MIN({ROUND1_MAX_SCORE},{"+".join(terms)}))')
        sc = f'{ccol["Điểm Vòng 1 (tự tính)"]}{r}'
        hgb, unit = f'{col["HGB"]}{r}', f'{col["Đơn vị HGB"]}{r}'
        hb_g = f'{ccol["HGB g/dL quy đổi"]}{r}'
        mcv, mch, rbc = f'{col["MCV"]}{r}', f'{col["MCH"]}{r}', f'{col["RBC"]}{r}'
        alt = f'{col[COL_ALT]}{r}'
        formulas = {
            "Điểm Vòng 1 (tự tính)": score_f,
            "Mức Vòng 1": f'=IF({sc}="","",IF({sc}>={ROUND1_HIGH_THRESHOLD},"CAO",IF({sc}>=4,"TRUNG BÌNH","THẤP")))',
            "HGB g/dL quy đổi": f'=IF({hgb}="","",IF({unit}="g/L",{hgb}/10,{hgb}))',
            "Hb sau hiệu chỉnh độ cao": f'=IF({hb_g}="","",{hb_g}-IFERROR(VLOOKUP({alt},Danh_sach!$A$2:$B$11,2,FALSE),0))',
            "Mentzer (MCV/RBC)": f'=IF(OR({mcv}="",{rbc}="",{rbc}=0),"",ROUND({mcv}/{rbc},2))',
            "Tiêu chí CTM BYT 2022": (f'=IF(OR({mcv}="",{mch}=""),"",IF(OR({mcv}<85,{mch}<28),'
                                      f'"Đạt tiêu chí sàng lọc (MCV<85 và/hoặc MCH<28)","Chưa đạt tiêu chí"))'),
            "HbA2 > 3,5%?": f'=IF({col["HbA2 (%)"]}{r}="","",IF({col["HbA2 (%)"]}{r}>3.5,"Có","Không"))',
            "Ferritin < 30?": f'=IF({col["Ferritin"]}{r}="","",IF({col["Ferritin"]}{r}<30,"Có","Không"))',
        }
        for name in computed:
            ws.write_formula(r - 1, n_in + computed.index(name), formulas[name], calc)

    def validate(name, source):
        c = inputs.index(name)
        ws.data_validation(1, c, n_rows, c, {"validate": "list", "source": source})

    validate(COL_SEX, BATCH_SEX)
    validate(COL_MARITAL, BATCH_MARITAL)
    validate(COL_PREG, ["Có", "Không"])
    validate("Đơn vị HGB", ["g/dL", "g/L"])
    validate("Đơn vị RBC", ["T/L", "10^12/L", "10^6/µL"])
    validate(COL_ALT, "=Danh_sach!$A$2:$A$11")
    for q in Q_COLS:
        validate(q, "=Danh_sach!$F$2:$F$6" if q == "Q9" else "=Danh_sach!$D$2:$D$4")

    ws.set_column(0, 0, 22)
    ws.set_column(1, inputs.index(COL_ALT), 15)
    ws.set_column(inputs.index("Q1"), inputs.index("Q20"), 7)
    ws.set_column(inputs.index("WBC"), inputs.index("Đơn vị RBC"), 9)
    ws.set_column(inputs.index("Ferritin"), inputs.index(COL_NOTE), 16)
    ws.set_column(inputs.index(COL_HPLC), inputs.index(COL_LAB), 28)
    ws.set_column(n_in, n_in + len(computed) - 1, 22)
    ws.set_row(0, 38)
    ws.freeze_panes(1, 1)

    wl = wb.add_worksheet("Danh_sach")
    wl.write(0, 0, "Khoảng độ cao", head)
    wl.write(0, 1, "Hiệu chỉnh Hb (g/dL)", head)
    for i, (k, v) in enumerate(ALTITUDE_OPTIONS.items(), start=1):
        wl.write(i, 0, k)
        wl.write(i, 1, v)
    wl.write(0, 3, "Lựa chọn Có/Không", head)
    for i, v in enumerate(BATCH_YNK, start=1):
        wl.write(i, 3, v)
    wl.write(0, 5, "Lựa chọn câu 9", head)
    for i, v in enumerate(BATCH_Q9, start=1):
        wl.write(i, 5, v)
    wl.set_column(0, 5, 26)

    wh = wb.add_worksheet("Huong_dan")
    lines = [
        "HƯỚNG DẪN",
        "• Mỗi dòng ở sheet Du_lieu là một người. Cột xám bên phải là công thức tự tính — không cần nhập.",
        "• Câu hỏi Q1–Q20 (cột cam): chọn Có / Không / Không biết (Q9 có danh sách riêng). Q19–Q20 không tính điểm.",
        "• CTM 18 chỉ số: WBC, NEU%, LYM%, MONO%, EOS%, BASO%, RBC, HGB, HCT, MCV, MCH, MCHC, RDW-CV, RDW-SD, PLT, MPV, PDW, PCT.",
        "  Bắt buộc để chạy Vòng 2: HGB, MCV, MCH, RBC, RDW-CV. Các chỉ số còn lại tùy chọn.",
        "• Sắt · HPLC · PCR: ferritin (ng/mL), sắt huyết thanh, TSAT; HbA/HbA2/HbF/HbE (%); kết luận HPLC/PCR ghi đúng theo phiếu.",
        "• Tuổi, Giới tính, Tình trạng hôn nhân, Đang mang thai (không bắt buộc): dùng đối chiếu ngưỡng thiếu máu WHO 2024 và gợi ý tư vấn di truyền.",
        "• Cũng có thể tải file Word, PDF, ảnh, SPSS (.sav), Stata (.dta), SAS, R... — hệ thống tự nhận cột (AI hỗ trợ ghép cột khi được phép).",
        "• Tiêu chí sàng lọc CTM: MCV < 85 fL và/hoặc MCH < 28 pg (Phụ lục hướng dẫn BYT 2022). Mentzer chỉ là chỉ số hỗ trợ.",
        "• Có thể tải trực tiếp file này lên trang Phân tích hồ sơ hàng loạt. Kết quả chỉ hỗ trợ sàng lọc, không chẩn đoán.",
        "",
        "CÂU HỎI TẦM SOÁT:",
    ]
    for q in Q_COLS:
        lines.append(f"{q}: {texts.get(q, '')}")
    for i, line in enumerate(lines):
        wh.write(i, 0, line)
    wh.set_column(0, 0, 120)
    wb.close()
    buf.seek(0)
    return buf.getvalue()


# ------------------------------------------------------------
# TRANG PHÂN TÍCH HÀNG LOẠT
# ------------------------------------------------------------

def _risk_cell_style(v):
    colors = {"THẤP": "rgba(46,125,50,.25)", "TRUNG BÌNH": "rgba(230,150,30,.28)", "CAO": "rgba(211,47,47,.28)"}
    return f"background-color: {colors[v]}; font-weight: 700" if v in colors else ""


def _editor_column_config():
    texts = _question_texts()
    cfg = {c: st.column_config.NumberColumn(c, format="%g", min_value=0.0) for c in NUM_COLS}
    cfg[COL_AGE] = st.column_config.NumberColumn("Tuổi", format="%g", min_value=0.0, max_value=120.0)
    cfg[COL_SEX] = st.column_config.SelectboxColumn("Giới tính", options=BATCH_SEX)
    cfg[COL_MARITAL] = st.column_config.SelectboxColumn("Tình trạng hôn nhân", options=BATCH_MARITAL)
    cfg[COL_PREG] = st.column_config.SelectboxColumn("Đang mang thai", options=["Không", "Có"])
    cfg[COL_ALT] = st.column_config.SelectboxColumn("Khoảng độ cao", options=list(ALTITUDE_OPTIONS.keys()))
    cfg["Đơn vị HGB"] = st.column_config.SelectboxColumn("Đơn vị HGB", options=["g/dL", "g/L"])
    cfg["Đơn vị RBC"] = st.column_config.SelectboxColumn("Đơn vị RBC", options=["T/L", "10^12/L", "10^6/µL"])
    for q in Q_COLS:
        cfg[q] = st.column_config.SelectboxColumn(
            q, options=BATCH_Q9 if q == "Q9" else BATCH_YNK, help=texts.get(q), width="small")
    for c in (COL_LABEL, COL_SRC, COL_NOTE, *TEXT_RESULT_COLS):
        cfg[c] = st.column_config.TextColumn(c, width="medium")
    return cfg


@st.fragment
def render_batch_page(user):
    # @st.fragment: mọi thao tác trên trang này chỉ chạy lại riêng trang này, không dựng lại cả ứng dụng.
    st.header(":material/batch_prediction: Phân tích hồ sơ hàng loạt")
    st.caption(
        "Phân tích nhiều hồ sơ đầy đủ cùng lúc: thông tin nhân khẩu, 20 câu tầm soát Vòng 1, công thức máu "
        "18 chỉ số, ferritin/sắt, HPLC và PCR/di truyền. Đọc được Excel, CSV, Word, PDF, ảnh và file phần mềm "
        "thống kê (SPSS, Stata, SAS, R); kiểm tra trong bảng trước khi phân tích."
    )

    cb = render_codebook_panel()

    with st.container(border=True):
        st.subheader(":material/upload_file: 1. Tải hồ sơ")
        files = st.file_uploader(
            f"Chọn tối đa {BATCH_MAX_FILES} file — Excel/CSV/Word/PDF/ảnh/SPSS/Stata/SAS/R "
            f"(bảng ≤ {BATCH_MAX_MB_TABLE} MB, PDF/ảnh ≤ {BATCH_MAX_MB_AI} MB)",
            type=UPLOAD_TYPES,
            accept_multiple_files=True, key="batch_files",
        )
        default_alt = st.selectbox(
            "Khoảng độ cao mặc định cho các hồ sơ mới", list(ALTITUDE_OPTIONS.keys()),
            key="batch_default_alt", help="Có thể sửa riêng từng dòng trong bảng bên dưới.",
        )
        ai_ok = st.checkbox(
            "Tôi xác nhận người bệnh đã đồng ý và cho phép gửi PDF/ảnh/nội dung Word, hoặc TÊN CỘT của file bảng, tới dịch vụ AI để đọc/ghép dữ liệu.",
            key="batch_ai_consent",
        )
        st.caption(
            "PDF/ảnh và Word dạng phiếu: nội dung được gửi tới dịch vụ AI của Anthropic để trích xuất CTM, ferritin, "
            "HPLC, PCR, thông tin nhân khẩu (và câu trả lời nếu là phiếu khảo sát). File bảng/SPSS/Stata/SAS/R được đọc "
            "ngay trên máy chủ; chỉ khi có cột lạ mới gửi TÊN CỘT (không gửi dữ liệu) cho AI ghép. Hệ thống không lưu file "
            "(file SPSS/R được ghi tạm vào thư mục tạm và xóa ngay). Nên che họ tên, ngày sinh trên PDF/ảnh. "
            "AI có thể đọc sai — luôn đối chiếu bảng với phiếu gốc."
        )
        b1, b2 = st.columns(2)
        scan = b1.button("Đọc dữ liệu từ file", icon=":material/document_scanner:",
                         type="primary", use_container_width=True)
        b2.download_button(
            "Tải file Excel mẫu đầy đủ (có công thức)", data=build_template_xlsx(),
            file_name="Mau_ho_so_Thalassemia_day_du.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:", use_container_width=True,
        )

    if scan:
        if not files:
            st.warning("Hãy chọn ít nhất một file.")
        elif len(files) > BATCH_MAX_FILES:
            st.error(f"Tối đa {BATCH_MAX_FILES} file mỗi lần.")
        else:
            rows, all_notes, progress = [], [], st.progress(0.0, text="Đang đọc file...")
            for i, f in enumerate(files):
                name, data = f.name, f.getvalue()
                progress.progress(i / len(files), text=f"Đang đọc {name} ({i + 1}/{len(files)})")
                try:
                    part, notes = load_file_rows(name, data, default_alt, ai_ok, cb)
                    rows.extend(part)
                    all_notes.extend(notes)
                except Exception as exc:
                    rows.append(_blank_row(name, "Lỗi đọc file", default_alt,
                                           f"Không đọc được: {exc}. Có thể nhập tay."))
            if any(str(r.get(COL_SRC, "")).startswith("Bảng:") for r in rows):
                all_notes.insert(0, "Quy ước mã hóa đã áp dụng — " + codebook_summary_text(cb))
            if len(rows) > BATCH_MAX_TOTAL_ROWS:
                all_notes.append(f"Chỉ giữ {BATCH_MAX_TOTAL_ROWS} dòng đầu (tổng {len(rows)} dòng).")
                rows = rows[:BATCH_MAX_TOTAL_ROWS]
            progress.progress(1.0, text="Hoàn tất đọc file.")
            st.session_state["batch_rows"] = rows
            st.session_state["batch_notes"] = all_notes
            st.session_state["batch_ver"] = st.session_state.get("batch_ver", 0) + 1
            st.session_state.pop("batch_results", None)

    # ---------------- Bảng kiểm tra/sửa ----------------
    st.subheader(":material/edit_note: 2. Kiểm tra và chỉnh sửa dữ liệu")
    st.caption(
        "Chọn nhóm cột để xem/sửa (dữ liệu ở các nhóm khác vẫn được giữ nguyên). Đối chiếu với phiếu gốc, "
        "thêm dòng mới hoặc xóa dòng (chọn dòng rồi nhấn Delete). Ô trống = chưa có dữ liệu."
    )
    for note in st.session_state.get("batch_notes", []):
        st.info(note, icon=":material/info:")
    group = st.radio("Nhóm cột", list(COLUMN_GROUPS.keys()), horizontal=True,
                     key="batch_group", label_visibility="collapsed")

    rows = st.session_state.get("batch_rows", [])
    df = pd.DataFrame(rows, columns=BATCH_COLUMNS)
    for col in NUM_COLS:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    edited = st.data_editor(
        df, num_rows="dynamic", hide_index=True, use_container_width=True,
        key=f"batch_editor_{st.session_state.get('batch_ver', 0)}",
        column_order=COLUMN_GROUPS[group], disabled=[COL_SRC],
        column_config=_editor_column_config(),
    )

    run = st.button("Phân tích tất cả hồ sơ", icon=":material/analytics:", type="primary",
                    use_container_width=True, disabled=edited is None or len(edited) == 0)

    if run:
        records = edited.to_dict("records")
        with st.spinner("Đang phân tích..."):
            st.session_state["batch_results"] = [analyze_batch_row(i, r) for i, r in enumerate(records)]
            st.session_state["batch_input_df"] = edited.copy()
            st.session_state["batch_codebook_df"] = codebook_summary_df(cb)

    results = st.session_state.get("batch_results")
    if not results:
        return

    # ---------------- Kết quả ----------------
    st.divider()
    st.subheader(":material/monitoring: 3. Kết quả phân tích")
    counts = Counter(r.get("Mức lưu ý cao nhất") for r in results if r.get("Trạng thái") == "OK")
    insufficient = sum(1 for r in results if r.get("Trạng thái") != "OK")

    m = st.columns(5)
    m[0].metric("Tổng hồ sơ", len(results))
    m[1].metric("Mức thấp", counts.get("THẤP", 0))
    m[2].metric("Mức trung bình", counts.get("TRUNG BÌNH", 0))
    m[3].metric("Mức cao", counts.get("CAO", 0))
    m[4].metric("Không đủ dữ liệu", insufficient)
    st.caption("“Mức lưu ý cao nhất” = mức cao nhất trong các vòng có dữ liệu (Vòng 1, CTM Vòng 2, dấu hiệu HPLC/PCR Vòng 3).")

    all_cols = ["Hồ sơ", "Nguồn", "Tuổi", "Giới tính", "Hôn nhân", "Mang thai", "Vòng 1 - điểm", "Vòng 1 - nguy cơ", "Số chỉ số CTM",
                "Hb (g/dL)", "Hb sau hiệu chỉnh", "Thiếu máu (WHO 2024)", "MCV", "MCH", "MCHC", "RDW-CV",
                "RBC (T/L)", "Mentzer", "Điểm CBC", "Vòng 2 - nguy cơ", "Ferritin", "HbA2 (%)", "HbE (%)",
                "Vòng 3 - mức lưu ý", "Vòng 3 - đối chiếu BYT", "Mức lưu ý cao nhất", "Bước tiếp theo",
                "Trạng thái", "Cảnh báo"]
    show_cols = ["Hồ sơ", "Vòng 1 - điểm", "Vòng 1 - nguy cơ", "Hb sau hiệu chỉnh", "MCV", "MCH", "Mentzer",
                 "Vòng 2 - nguy cơ", "Ferritin", "HbA2 (%)", "HbE (%)", "Vòng 3 - mức lưu ý",
                 "Mức lưu ý cao nhất", "Bước tiếp theo", "Trạng thái", "Cảnh báo"]
    full_df = pd.DataFrame([{k: v for k, v in r.items() if k != "_detail"} for r in results]).reindex(columns=all_cols)
    view_df = full_df[show_cols]

    risk_subset = ["Vòng 1 - nguy cơ", "Vòng 2 - nguy cơ", "Vòng 3 - mức lưu ý", "Mức lưu ý cao nhất"]
    try:
        styled = view_df.style.map(_risk_cell_style, subset=[c for c in risk_subset if c in view_df.columns])
    except AttributeError:
        styled = view_df.style.applymap(_risk_cell_style, subset=[c for c in risk_subset if c in view_df.columns])
    st.dataframe(styled, use_container_width=True, hide_index=True)

    if insufficient:
        st.warning(f"{insufficient} hồ sơ không đủ dữ liệu để phân tích. Xem cột Cảnh báo, bổ sung ở bảng bước 2 rồi phân tích lại.",
                   icon=":material/warning:")

    input_df = st.session_state.get("batch_input_df")
    if input_df is not None:
        st.download_button(
            "Tải kết quả Excel (đầy đủ cột + dữ liệu đầu vào)",
            data=results_to_excel(full_df, input_df, st.session_state.get("batch_codebook_df")).getvalue(),
            file_name=f"Ket_qua_phan_tich_hang_loat_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:", use_container_width=True,
        )

    ok_idx = [i for i, r in enumerate(results) if r.get("Trạng thái") == "OK"]
    if ok_idx:
        st.markdown("### :material/clinical_notes: Xem chi tiết từng hồ sơ")
        pick = st.selectbox(
            "Chọn hồ sơ", ok_idx, key="batch_detail_pick",
            format_func=lambda i: f"{i + 1}. {results[i]['Hồ sơ']} — {results[i].get('Mức lưu ý cao nhất') or 'chưa phân tầng'}",
        )
        det = results[pick]["_detail"]

        if "r1" in det:
            d = det["r1"]
            risk_alert(f"Vòng 1 ({d['score']}/{ROUND1_MAX_SCORE} điểm, {d['answered']}/20 câu)", d["category"], d["conclusion"])
            if d["reasons"]:
                bullet_list(d["reasons"])
        if "r2" in det:
            d = det["r2"]
            risk_alert("CTM — Vòng 2", d["category"], d["conclusion"])
            st.caption(f"Khoảng độ cao: {d['altitude']} · Mentzer {d['mentzer']:.2f} · Điểm CBC {d['score']}")
            with st.expander("Các yếu tố, phân tích sơ bộ và khuyến nghị CTM", icon=":material/search:"):
                st.markdown("**Các yếu tố từ CBC**")
                bullet_list(d["reasons"])
                st.markdown("**Phân tích sơ bộ**")
                bullet_list(d["findings"])
                st.markdown("**Khuyến nghị**")
                bullet_list(d["advice"])
        if det.get("cbc_values"):
            with st.expander(f"Các chỉ số CTM đã nhập ({len(det['cbc_values'])}/18)", icon=":material/water_drop:"):
                st.dataframe(pd.DataFrame([det["cbc_values"]]), hide_index=True, use_container_width=True)
        if "r3" in det:
            st.markdown("**Sắt · HPLC · PCR — đối chiếu BYT 2022**")
            bullet_list(det["r3"]["notes"])
        if det["steps"]:
            st.markdown("**Gợi ý bước tiếp theo**")
            bullet_list(det["steps"])
        if det["warnings"]:
            st.warning("; ".join(det["warnings"]), icon=":material/warning:")

    st.caption(
        "Kết quả chỉ hỗ trợ sàng lọc, không chẩn đoán; cần đối chiếu lâm sàng, xét nghiệm chuyên sâu "
        "và nhân viên y tế. Trang này không lưu dữ liệu vào database."
    )


FACILITY_DISPLAY_LIMIT = 12


@st.fragment
def render_facilities(patient):
    """Lọc/tìm cơ sở y tế chỉ chạy lại riêng khối này (không dựng lại toàn bộ trang)."""
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

    st.write(
        f"**{len(filtered)} cơ sở** phù hợp với bộ lọc."
        + (f" Chỉ hiển thị {FACILITY_DISPLAY_LIMIT} cơ sở đầu tiên; hãy lọc thêm để thu hẹp."
           if len(filtered) > FACILITY_DISPLAY_LIMIT else "")
    )
    if filtered:
        for i, facility in enumerate(filtered[:FACILITY_DISPLAY_LIMIT], start=1):
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


# ============================================================
# HEADER
# ============================================================

st.markdown(
    '<div class="hematology-brand"><div class="brand-orb"></div><div class="brand-text">'
    '<div class="eyebrow">HEMATOLOGY · COMMUNITY SCREENING RESEARCH</div>'
    '<div class="title title-main">HỆ THỐNG HỖ TRỢ SÀNG LỌC THALASSEMIA</div>'
    '<div class="chips"><span class="chip">Vòng 1 · Tầm soát</span><span class="chip">Vòng 2 · Công thức máu</span>'
    '<span class="chip">Vòng 3 · Theo dõi</span><span class="chip">Prototype</span></div></div></div>'
    '<div class="section-rule"></div>',
    unsafe_allow_html=True,
)

@st.dialog("Giới thiệu đề tài và cơ sở nghiên cứu", width="large")
def show_project_intro():
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

# Nội dung dài chỉ được dựng khi người dùng bấm nút → trang chính nhẹ hơn nhiều.
_bi1, _bi2, _ = st.columns([1, 1, 2])
if _bi1.button("Giới thiệu đề tài", icon=":material/menu_book:", use_container_width=True):
    show_project_intro()
if _bi2.button("Cơ sở khoa học", icon=":material/school:", use_container_width=True):
    show_research_overview()

st.markdown(
    '<div class="contact-card"><div class="label">Liên hệ tác giả</div>'
    '<div class="name">Nguyễn Đắc Lâu</div>'
    '<div>Điện thoại: <a href="tel:0357930820">0357930820</a></div>'
    '<div>Email: <a href="mailto:nguyendaclau2309@gmail.com">nguyendaclau2309@gmail.com</a></div>'
    '<div class="small-note">Tác giả và người phát triển nguyên mẫu hệ thống hỗ trợ sàng lọc Thalassemia.</div></div>',
    unsafe_allow_html=True,
)

# ============================================================
# ACCESS / SIDEBAR / ADMIN
# ============================================================

render_auth_sidebar()

st.sidebar.divider()
st.sidebar.toggle(
    "Hiệu ứng kính (Liquid Glass)", value=True, key="glass_fx",
    help="Tắt nếu máy yếu hoặc mạng chậm để trang nhẹ và mở nhanh hơn.",
)

auth_user = current_auth_user()

if auth_user and st.session_state.get("auth_page") == PAGE_ADMIN:
    render_admin_console(auth_user)
    st.stop()

if auth_user and st.session_state.get("auth_page") == PAGE_BATCH:
    render_batch_page(auth_user)
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

# B. Địa giới — fragment: đổi tỉnh/thành chỉ chạy lại riêng khối này, không chạy lại cả trang.
@st.fragment
def address_picker():
    st.subheader(":material/location_on: Địa giới hành chính hiện tại")
    province, commune, ok = "", "", False

    try:
        admin_hierarchy = load_admin_hierarchy()
        ok = True
    except Exception as exc:
        admin_hierarchy = {}
        st.error("Không tải được danh mục tỉnh/thành và phường/xã/đặc khu. Vui lòng tải lại trang hoặc thử lại sau.")
        st.caption(f"Nguồn dữ liệu: {ADMIN_DATA_SOURCE_URL}")
        st.code(str(exc))

    if ok:
        provinces_list = sorted(admin_hierarchy.keys(), key=str.casefold)
        province_choice = st.selectbox(
            "Tỉnh / thành phố *", ["— Chọn tỉnh/thành —"] + provinces_list, key="v22_profile_province"
        )
        if province_choice == "— Chọn tỉnh/thành —":
            st.selectbox("Phường / xã / đặc khu *", ["— Chọn tỉnh/thành trước —"],
                         disabled=True, key="v22_profile_commune_empty")
        else:
            province = province_choice
            available_communes = list(admin_hierarchy.get(province, []))
            # Key phụ thuộc tỉnh để không tái sử dụng lựa chọn của tỉnh trước đó.
            safe_province = province.replace(" ", "_").replace("/", "_").replace("\\", "_")
            commune_choice = st.selectbox(
                "Phường / xã / đặc khu *",
                ["— Chọn phường/xã/đặc khu —"] + available_communes,
                key=f"v22_profile_commune_{safe_province}",
            )
            if commune_choice != "— Chọn phường/xã/đặc khu —":
                commune = commune_choice
            if not available_communes:
                st.error("Tỉnh/thành này chưa có danh mục phường/xã/đặc khu trong dữ liệu hiện tại.")

        st.caption(
            "Danh mục địa giới được tải theo cấu trúc 2 cấp hiện hành; "
            "không nhập tay tên phường/xã để tránh sai địa danh."
        )

    st.session_state["addr_data_ok"] = ok
    st.session_state["addr_province"] = province
    st.session_state["addr_commune"] = commune

    if province and commune:
        st.success(f"Đã chọn: **{commune}, {province}**", icon=":material/location_on:")


address_picker()
admin_data_ok = st.session_state.get("addr_data_ok", False)
selected_province = st.session_state.get("addr_province", "")
commune_value = st.session_state.get("addr_commune", "")

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
        "Bạn chưa đồng ý tham gia. Hệ thống không mở Vòng 1 và không xử lý/lưu dữ liệu sàng lọc.",
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
        st.info("Chưa ghi nhận yếu tố đáng lưu ý từ các câu trả lời có tính điểm.")

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

        render_facilities(patient)

        # ----------------------------------------------------
        # WORD REPORT
        # ----------------------------------------------------

        st.subheader(":material/description: Phiếu kết quả")

        report = make_word_bytes(
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
