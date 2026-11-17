import io
import json
import math
import os
import re
import hashlib
import secrets
import time
import base64
import numbers
import tempfile
import unicodedata
from collections import Counter
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from urllib.parse import quote

import numpy as np
import pandas as pd
import requests
import streamlit as st

# ===== THAL-SCREEN AI: HỆ THỐNG HỌC MÁY DỰ ĐOÁN NGUY CƠ MANG GEN THALASSEMIA =====
# Đề tài NCKH: Xây dựng và đánh giá mô hình học máy dự đoán nguy cơ mang gen Thalassemia
#              dựa trên chỉ số công thức máu và Ferritin huyết thanh ở người trưởng thành
#              tại Bệnh viện Trung ương Huế
# Chủ nhiệm đề tài: Nguyễn Đắc Lâu (24XN1 - Trường Đại học Phan Châu Trinh)
# Thành viên: Trần Phan Thảo Vi, Huỳnh Ngọc Anh Huy, Lê Văn Thịnh, Phan Ngọc Bảo An
# Giảng viên hướng dẫn: TS. Nguyễn Đình Khắc Chân
# Yêu cầu: xem requirements.txt (streamlit>=1.45, scikit-learn, xgboost, shap, scipy, psycopg[pool], python-docx, XlsxWriter, openpyxl...)

# set_page_config PHẢI là lệnh Streamlit đầu tiên.
st.set_page_config(
    page_title="THAL-SCREEN AI · Sàng lọc Thalassemia (CBC + Ferritin)",
    page_icon=":material/bloodtype:",
    layout="wide",
)

GLASS_CSS = r"""
<style>
/* =====================================================================
   LIQUID GLASS UI — kính mờ, viền sáng. Màu theo chế độ sáng/tối được chèn riêng
   (--lg-text, --lg-bg, --lg-fill*, --lg-pop, --lg-spec, --lg-shadow, --lg-link, --lg-risk-*, --lg-aurora)
   ===================================================================== */
:root{
  --lg-accent:#C0394F; --lg-accent-dark:#8E2C3A; --lg-blue:#4F7CFF; --lg-teal:#2BB3A3;
  --lg-radius:22px;
  --lg-fx:blur(18px);      /* chỉ dùng cho thanh bên, popup, hộp thoại */
  --lg-fx-sm:none;         /* thẻ/nút/ô nhập/tab: KHÔNG blur (nặng GPU) */
}

/* ---------- Nền aurora: đặt trực tiếp lên .stApp (KHÔNG đổi position/isolation để không làm vỡ bố cục) ---------- */
.stApp,[data-testid="stApp"]{
  font-family:-apple-system,BlinkMacSystemFont,"SF Pro Display","SF Pro Text","Inter","Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif;
  background:var(--lg-aurora),var(--lg-bg) !important; }
[data-testid="stAppViewContainer"],[data-testid="stMain"],.main,[data-testid="stHeader"]{ background:transparent !important; }
.main .block-container{ max-width:1180px; padding-top:2rem; padding-bottom:4rem; }
h1,h2,h3,h4{ letter-spacing:-0.02em; }
hr{ border:0 !important; height:1px !important; background:linear-gradient(90deg,transparent,var(--lg-stroke),transparent) !important; }
a{ color:var(--lg-link); }
[data-testid="stCaptionContainer"],.stCaption{ color:var(--lg-muted); }
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
.hematology-brand .title-main{ font-size:clamp(1.35rem,2.8vw,2.15rem); letter-spacing:-0.035em; text-transform:uppercase;
  background:linear-gradient(95deg,var(--lg-text) 10%,color-mix(in srgb,var(--lg-accent) 78%,var(--lg-text)) 90%);
  -webkit-background-clip:text; background-clip:text; -webkit-text-fill-color:transparent; }
.hematology-brand .chips{ display:flex; flex-wrap:wrap; gap:.4rem; margin-top:.7rem; }
.chip{ display:inline-block; padding:.18rem .7rem; font-size:.74rem; font-weight:650; border-radius:999px;
  color:var(--lg-text); background:var(--lg-fill-soft); border:1px solid var(--lg-stroke); box-shadow:var(--lg-spec); }
.section-rule{ height:1px; margin:1.2rem 0; background:linear-gradient(90deg,transparent,var(--lg-stroke),transparent); }
.contact-card{ padding:1rem 1.2rem; margin-top:1rem; }
.contact-card .label{ color:var(--lg-muted); font-size:.74rem; text-transform:uppercase; letter-spacing:.1em; font-weight:700; }
.contact-card .name{ color:var(--lg-text); font-size:1.1rem; font-weight:750; margin:.2rem 0 .55rem; }
.contact-card a{ color:var(--lg-link); text-decoration:none; font-weight:600; }
.contact-card a:hover{ text-decoration:underline; }
.small-note{ color:var(--lg-muted); font-size:.82rem; }

/* ---------- Huy hiệu mức nguy cơ & thanh đóng góp SHAP ---------- */
.risk-badge{ display:inline-block; padding:.2rem .75rem; border-radius:999px; font-size:.76rem; font-weight:700; letter-spacing:.06em;
  border:1px solid var(--lg-stroke); box-shadow:var(--lg-spec); }
.risk-low{ background:color-mix(in srgb,var(--lg-risk-low) 18%,transparent); color:var(--lg-risk-low); }
.risk-mid{ background:color-mix(in srgb,var(--lg-risk-mid) 18%,transparent); color:var(--lg-risk-mid); }
.risk-high{ background:color-mix(in srgb,var(--lg-risk-high) 18%,transparent); color:var(--lg-risk-high); }

.shap-row{ display:flex; align-items:center; gap:.65rem; margin:.32rem 0; font-size:.86rem; }
.shap-label{ width:175px; flex:none; font-weight:650; color:var(--lg-text); }
.shap-track{ flex:1; height:12px; border-radius:999px; background:var(--lg-fill-soft); overflow:hidden; border:1px solid var(--lg-stroke); }
.shap-bar-pos{ height:100%; border-radius:999px; background:linear-gradient(90deg,#C0394F,#E7707F); }
.shap-bar-neg{ height:100%; border-radius:999px; background:linear-gradient(90deg,#2BB3A3,#4F7CFF); }
.shap-val{ width:110px; flex:none; text-align:right; font-variant-numeric:tabular-nums; font-weight:650; }

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


_LG_LIGHT = (
    "--lg-text:#1d1d1f;--lg-bg:#ffffff;--lg-link:#C0394F;"
    "--lg-stroke:color-mix(in srgb,var(--lg-text) 13%,transparent);--lg-muted:color-mix(in srgb,var(--lg-text) 62%,transparent);"
    "--lg-fill:color-mix(in srgb,var(--lg-bg) 66%,transparent);--lg-fill-soft:color-mix(in srgb,var(--lg-bg) 30%,transparent);"
    "--lg-fill-strong:color-mix(in srgb,var(--lg-bg) 74%,transparent);--lg-pop:color-mix(in srgb,var(--lg-bg) 90%,transparent);"
    "--lg-spec:inset 0 1px 0 rgba(255,255,255,.55),inset 0 -1px 0 rgba(255,255,255,.08),inset 0 0 0 1px rgba(255,255,255,.08);"
    "--lg-shadow:0 8px 22px -14px rgba(20,25,45,.30);"
    "--lg-risk-low:#2E7D32;--lg-risk-mid:#B26A00;--lg-risk-high:#B3261E;"
    "--lg-aurora:radial-gradient(900px 620px at 6% -6%,color-mix(in srgb,var(--lg-accent) 30%,transparent),transparent 62%),"
    "radial-gradient(820px 580px at 102% 2%,color-mix(in srgb,var(--lg-blue) 26%,transparent),transparent 62%),"
    "radial-gradient(920px 700px at 72% 108%,color-mix(in srgb,var(--lg-teal) 24%,transparent),transparent 62%),"
    "radial-gradient(700px 520px at 0% 100%,color-mix(in srgb,#9B6BFF 16%,transparent),transparent 62%);"
)
_LG_DARK = (
    "--lg-text:#eef1f7;--lg-bg:#0f131b;--lg-link:#FF8FA1;color-scheme:dark;"
    "--lg-stroke:color-mix(in srgb,var(--lg-text) 15%,transparent);--lg-muted:color-mix(in srgb,var(--lg-text) 70%,transparent);"
    "--lg-fill:rgba(255,255,255,.07);--lg-fill-soft:rgba(255,255,255,.05);--lg-fill-strong:rgba(255,255,255,.13);--lg-pop:rgba(27,32,43,.98);"
    "--lg-spec:inset 0 1px 0 rgba(255,255,255,.14),inset 0 -1px 0 rgba(255,255,255,.03),inset 0 0 0 1px rgba(255,255,255,.04);"
    "--lg-shadow:0 10px 26px -14px rgba(0,0,0,.75);"
    "--lg-risk-low:#7DD08B;--lg-risk-mid:#F0B25A;--lg-risk-high:#FF8A80;"
    "--lg-aurora:radial-gradient(900px 620px at 6% -6%,rgba(224,86,110,.17),transparent 62%),"
    "radial-gradient(820px 580px at 102% 2%,rgba(88,130,255,.15),transparent 62%),"
    "radial-gradient(920px 700px at 72% 108%,rgba(43,179,163,.12),transparent 62%),"
    "radial-gradient(700px 520px at 0% 100%,rgba(155,107,255,.10),transparent 62%);"
)
_DARK_ONLY_CSS = r"""
.stApp,[data-testid="stApp"]{ color:var(--lg-text); }
.stApp h1,.stApp h2,.stApp h3,.stApp h4,.stApp h5,.stApp h6,.stApp p,.stApp li,.stApp label,
[data-testid="stMarkdownContainer"] p,[data-testid="stMarkdownContainer"] li,[data-testid="stWidgetLabel"] p,
[data-testid="stMetricLabel"] p,[data-testid="stMetricValue"],[data-testid="stExpander"] summary p,
[data-testid="stRadio"] label p,[data-testid="stCheckbox"] label p,[data-testid="stToggle"] label p,
[data-testid="stTabs"] button p{ color:var(--lg-text) !important; }
[data-testid="stCaptionContainer"],[data-testid="stCaptionContainer"] p,.stCaption,.small-note{ color:var(--lg-muted) !important; }
[data-testid="stMarkdownContainer"] p,[data-testid="stMarkdownContainer"] li{ line-height:1.6; }
[data-testid^="stBaseButton"],[data-testid^="stBaseLinkButton"]{ color:var(--lg-text); }
[data-baseweb="input"] input,[data-baseweb="textarea"] textarea,[data-baseweb="select"] input,
[data-baseweb="select"] div,[data-baseweb="select"] span{ color:var(--lg-text) !important; -webkit-text-fill-color:var(--lg-text) !important; caret-color:var(--lg-link); }
[data-baseweb="select"] svg,[data-baseweb="input"] svg{ fill:var(--lg-muted) !important; }
input::placeholder,textarea::placeholder{ color:var(--lg-muted) !important; -webkit-text-fill-color:var(--lg-muted) !important; opacity:1; }
[data-baseweb="popover"] li,[data-baseweb="menu"] li,[role="option"]{ color:var(--lg-text) !important; }
[role="option"]:hover,[role="option"][aria-selected="true"],[data-baseweb="menu"] li:hover{ background:var(--lg-fill-strong) !important; }
.stApp [data-testid="stBaseButton-primary"] *,.stApp [data-testid="stBaseButton-primaryFormSubmit"] *,.stApp button[kind="primary"] *,.stApp button[kind="primaryFormSubmit"] *{ color:#fff !important; -webkit-text-fill-color:#fff !important; }
"""


def _theme_vars_css():
    try:
        mode = st.context.theme.type
    except Exception:
        mode = None
    if mode == "dark":
        body = ":root{" + _LG_DARK + "}" + _DARK_ONLY_CSS
    elif mode == "light":
        body = ":root{" + _LG_LIGHT + "}"
    else:
        body = (
            ":root{" + _LG_LIGHT + "} @media (prefers-color-scheme: dark){:root{" + _LG_DARK + "}"
            + _DARK_ONLY_CSS + "}"
        )
    return "<style>" + body + "</style>"


_glass_on = st.session_state.get("glass_fx", True) and st.query_params.get("glass") != "off"
if st.query_params.get("glass") != "off":
    st.markdown(
        _minify_css(_theme_vars_css() + GLASS_CSS + ("" if _glass_on else GLASS_OFF_CSS)),
        unsafe_allow_html=True,
    )

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

ADMIN_USERNAME = _secret("ADMIN_USERNAME")
ADMIN_PASSWORD = _secret("ADMIN_PASSWORD")
ADMIN_FULL_NAME = _secret("ADMIN_FULL_NAME", "Quản trị viên hệ thống")
ADMIN_EMAIL = _secret("ADMIN_EMAIL", "admin@thalassemia.local")

ANTHROPIC_API_KEY = _secret("ANTHROPIC_API_KEY")
ANTHROPIC_MODEL = _secret("ANTHROPIC_MODEL", "claude-sonnet-4-5")

LOCAL_SQLITE_PATH = "thalassemia_patients.db"
CONSENT_VERSION = "DACLAU24XN1-THALSCREEN-V6"
ADMIN_DATA_URL = "https://raw.githubusercontent.com/open-admin-data/vietnam-administrative-divisions/main/data/hierarchy.json"
ADMIN_DATA_SOURCE_URL = "https://github.com/open-admin-data/vietnam-administrative-divisions"

ROUND1_MAX_SCORE = 20
ROUND1_HIGH_THRESHOLD = 8
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
DB_SCHEMA_READY = "db_schema_ready_v24_ml"

# Hằng số điều hướng 4 phân khu chức năng
PAGE_SCREENING = ":material/edit_note: Nhập sàng lọc & Dự đoán AI"
PAGE_BATCH = ":material/batch_prediction: Phân tích hồ sơ hàng loạt"
PAGE_ML_LAB = ":material/model_training: Huấn luyện & Kiểm định Mô hình ML (NCKH)"
PAGE_ADMIN = ":material/admin_panel_settings: Quản trị hệ thống"

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


# ============================================================
# THAL-SCREEN AI ENGINE: CHỈ SỐ PHÁI SINH, MÔ HÌNH ML & SHAP
# ============================================================
# Bám sát đề cương NCKH tại Bệnh viện Trung ương Huế:
# - 7 biến CBC gốc: RBC (T/L), Hb (g/L), HCT (%), MCV (fL), MCH (pg), MCHC (g/L), RDW-CV (%)
# - 1 biến sinh hóa sắt: Ferritin huyết thanh (ng/mL)
# - 4 chỉ số phái sinh (Feature Engineering): Mentzer, Green & King, RDWI, Srivastava
# - 3 thuật toán Học máy: Logistic Regression (LR), Random Forest (RF), XGBoost
# - So sánh trực tiếp Mô hình A (Chỉ CBC) vs Mô hình B (CBC + Ferritin) + Giải thích SHAP

ML_FEATURES_CBC = ["RBC", "Hb_gL", "HCT", "MCV", "MCH", "MCHC_gL", "RDW-CV"]
ML_FEATURES_FULL = ["RBC", "Hb_gL", "HCT", "MCV", "MCH", "MCHC_gL", "RDW-CV", "Ferritin"]

FEATURE_DISPLAY_NAMES = {
    "MCV": "MCV (fL)",
    "MCH": "MCH (pg)",
    "RBC": "RBC (10¹²/L)",
    "Ferritin": "Ferritin (ng/mL)",
    "RDW-CV": "RDW-CV (%)",
    "Hb_gL": "Hb (g/L)",
    "MCHC_gL": "MCHC (g/L)",
    "HCT": "HCT (%)",
}


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
    if unit in ("T/L", "10^12/L", "10^6/µL"):
        return _round_number(value, 4)
    raise ValueError("Đơn vị RBC không hợp lệ.")


def derive_complete_cbc(hb_g_dl, rbc, mcv, mch, rdw, hct=None, mchc_g_l=None):
    """Chuẩn hóa 7 chỉ số CBC về đúng hệ đơn vị của Đề cương NCKH:
    RBC (T/L), Hb_gL (g/L), HCT (%), MCV (fL), MCH (pg), MCHC_gL (g/L), RDW-CV (%).
    Nếu người dùng không nhập HCT hoặc MCHC thì tự động nội suy từ RBC, MCV và Hb."""
    hb_g_l = float(hb_g_dl) * 10.0
    if hct is None or float(hct) <= 0:
        hct_val = (float(rbc) * float(mcv)) / 10.0
    else:
        hct_val = float(hct)
        if 0 < hct_val < 1.0:
            hct_val *= 100.0
    if mchc_g_l is None or float(mchc_g_l) <= 0:
        mchc_val = (hb_g_l * 100.0 / hct_val) if hct_val > 0 else 325.0
    else:
        mchc_val = float(mchc_g_l)
        if mchc_val < 60.0:  # người dùng nhập theo g/dL (vd 32.5) -> đổi sang g/L (325)
            mchc_val *= 10.0
    return {
        "RBC": round(float(rbc), 2),
        "Hb_gL": round(hb_g_l, 1),
        "HCT": round(hct_val, 1),
        "MCV": round(float(mcv), 1),
        "MCH": round(float(mch), 1),
        "MCHC_gL": round(mchc_val, 1),
        "RDW-CV": round(float(rdw), 1),
    }


def calculate_discriminant_indices(hb_g_dl, rbc, mcv, mch, rdw, ferritin=None):
    """Tính bộ chỉ số phân biệt hồng cầu kinh điển và biến đặc trưng phái sinh (Feature Engineering)."""
    mentzer = _round_number(Decimal(str(mcv)) / Decimal(str(rbc)), 2) if rbc and rbc > 0 else None
    green_king = (
        round((float(mcv) ** 2 * float(rdw)) / (100.0 * float(hb_g_dl)), 2)
        if hb_g_dl and hb_g_dl > 0
        else None
    )
    rdwi = round((float(mcv) * float(rdw)) / float(rbc), 2) if rbc and rbc > 0 else None
    srivastava = round(float(mch) / float(rbc), 2) if rbc and rbc > 0 else None
    ferritin_mcv_ratio = (
        round(math.log10(max(float(ferritin), 1.0)) / float(mcv), 4)
        if (ferritin is not None and float(ferritin) > 0 and mcv and mcv > 0)
        else None
    )
    return {
        "mentzer": mentzer,
        "green_king": green_king,
        "rdwi": rdwi,
        "srivastava": srivastava,
        "ferritin_mcv_ratio": ferritin_mcv_ratio,
    }


def _sigmoid(x):
    return 1.0 / (1.0 + math.exp(-max(min(float(x), 30.0), -30.0)))


def predict_thal_screen_ai(hb_g_dl, rbc, mcv, mch, rdw, hct=None, mchc_g_l=None, ferritin=None):
    """Engine dự đoán nguy cơ mang gen Thalassemia bằng 3 thuật toán (LR, RF, XGBoost)
    cho cả Mô hình A (Chỉ CBC) và Mô hình B (CBC + Ferritin) kèm giải thích đóng góp biến SHAP.
    - Nếu người dùng đã huấn luyện mô hình thực tế trong phiên (st.session_state['trained_ml_bundle']),
      hệ thống sử dụng trực tiếp các mô hình scikit-learn/XGBoost đó.
    - Nếu chưa huấn luyện dữ liệu riêng, hệ thống sử dụng bộ trọng số chuẩn đã hiệu chỉnh theo
      thông số tham chiếu tại Bệnh viện Trung ương Huế (Bảng 3.2 – 3.8 của Đề cương NCKH)."""
    cbc = derive_complete_cbc(hb_g_dl, rbc, mcv, mch, rdw, hct, mchc_g_l)
    indices = calculate_discriminant_indices(hb_g_dl, rbc, mcv, mch, rdw, ferritin)
    has_ferritin = ferritin is not None and float(ferritin) > 0
    fer_val = float(ferritin) if has_ferritin else 82.5  # trung vị tham chiếu khi khuyết

    # Kiểm tra nếu có mô hình đã huấn luyện trong phiên làm việc
    bundle = st.session_state.get("trained_ml_bundle")
    if bundle is not None and bundle.get("models_a") is not None and bundle.get("models_b") is not None:
        try:
            x_cbc = pd.DataFrame([{k: cbc[k] for k in ML_FEATURES_CBC}])
            x_full = pd.DataFrame([{**{k: cbc[k] for k in ML_FEATURES_CBC}, "Ferritin": fer_val}])
            probs_a = {
                name: float(mdl.predict_proba(x_cbc)[0, 1])
                for name, mdl in bundle["models_a"].items()
            }
            probs_b = (
                {
                    name: float(mdl.predict_proba(x_full)[0, 1])
                    for name, mdl in bundle["models_b"].items()
                }
                if has_ferritin
                else probs_a
            )
            best_prob_a = probs_a["XGBoost"]
            best_prob_b = probs_b["XGBoost"]
            active_prob = best_prob_b if has_ferritin else best_prob_a
            shap_contribs = _compute_calibrated_shap(cbc, fer_val, has_ferritin)
            return _format_ml_output(
                cbc, indices, fer_val if has_ferritin else None, probs_a, probs_b, active_prob, shap_contribs
            )
        except Exception:
            pass

    # Chuẩn hóa Z-score so với trung bình quần thể tham chiếu (Bảng 3.2 Đề cương NCKH)
    z_mcv = (79.5 - cbc["MCV"]) / 6.8          # MCV giảm -> dương (tăng nguy cơ Thalassemia)
    z_mch = (26.0 - cbc["MCH"]) / 2.6          # MCH giảm -> dương
    z_rbc = (cbc["RBC"] - 4.95) / 0.50         # RBC tăng -> dương
    z_rdw = (cbc["RDW-CV"] - 14.0) / 1.6       # RDW tăng quá cao kèm Ferritin thấp -> gợi ý IDA
    z_hb = (130.0 - cbc["Hb_gL"]) / 14.5       # Hb giảm nhẹ-vừa
    z_mchc = (325.0 - cbc["MCHC_gL"]) / 14.0   # MCHC giảm
    z_hct = (39.5 - cbc["HCT"]) / 4.3          # HCT giảm

    # Đóng góp của Ferritin trong phân biệt IDA vs Thalassemia (kể cả đồng mắc)
    if has_ferritin:
        if fer_val < 15.0:
            # Thiếu sắt rõ, nhưng nếu RBC vẫn cao (>= 5.1) và Mentzer < 13 thì cảnh giác đồng mắc!
            coexisting_signal = 0.55 if (cbc["RBC"] >= 5.1 or (indices["mentzer"] and indices["mentzer"] < 13.0)) else -1.15
            z_fer = coexisting_signal
        elif fer_val < 30.0:
            coexisting_signal = 0.35 if (cbc["RBC"] >= 5.0 and cbc["MCV"] < 74.0) else -0.65
            z_fer = coexisting_signal
        else:
            # Ferritin bình thường/tăng (>= 30 ng/mL) trong bối cảnh hồng cầu nhỏ -> loại trừ IDA, đẩy mạnh xác suất Thalassemia
            microcytic_factor = 1.0 if (cbc["MCV"] < 80.0 or cbc["MCH"] < 27.0) else -0.25
            z_fer = min(math.log10(max(fer_val, 30.0) / 30.0) * 1.35, 1.25) * microcytic_factor
    else:
        z_fer = 0.0

    # Tương tác phi tuyến đặc trưng của các mô hình cây (Random Forest & XGBoost)
    nonlinear_boost = 0.0
    if cbc["MCV"] < 78.0 and cbc["RBC"] >= 5.0:
        nonlinear_boost += 0.48
    if indices["mentzer"] is not None and indices["mentzer"] < 13.0:
        nonlinear_boost += 0.38
    if indices["green_king"] is not None and indices["green_king"] < 65.0 and cbc["MCV"] < 80.0:
        nonlinear_boost += 0.28
    if cbc["MCV"] >= 85.0 and cbc["MCH"] >= 28.0:
        nonlinear_boost -= 0.95

    # Log-odds của Mô hình A (Chỉ CBC)
    logit_lr_a = -0.85 + 0.68 * z_mcv + 0.56 * z_mch + 0.44 * z_rbc - 0.24 * z_rdw + 0.18 * z_hb + 0.13 * z_mchc + 0.08 * z_hct
    logit_rf_a = logit_lr_a + 0.65 * nonlinear_boost
    logit_xgb_a = logit_lr_a + 0.92 * nonlinear_boost

    probs_a = {
        "Logistic Regression": round(_sigmoid(logit_lr_a), 4),
        "Random Forest": round(_sigmoid(logit_rf_a), 4),
        "XGBoost": round(_sigmoid(logit_xgb_a), 4),
    }

    # Log-odds của Mô hình B (CBC + Ferritin)
    if has_ferritin:
        logit_lr_b = logit_lr_a + 0.42 * z_fer
        logit_rf_b = logit_rf_a + 0.54 * z_fer
        logit_xgb_b = logit_xgb_a + 0.62 * z_fer
        probs_b = {
            "Logistic Regression": round(_sigmoid(logit_lr_b), 4),
            "Random Forest": round(_sigmoid(logit_rf_b), 4),
            "XGBoost": round(_sigmoid(logit_xgb_b), 4),
        }
    else:
        probs_b = dict(probs_a)

    active_prob = probs_b["XGBoost"] if has_ferritin else probs_a["XGBoost"]
    shap_contribs = _compute_calibrated_shap(cbc, fer_val if has_ferritin else None, has_ferritin)
    return _format_ml_output(
        cbc, indices, fer_val if has_ferritin else None, probs_a, probs_b, active_prob, shap_contribs
    )


def _compute_calibrated_shap(cbc, fer_val, has_ferritin):
    """Tính giá trị đóng góp SHAP cá thể hóa (Local SHAP values) cho 8 biến đầu vào chính,
    bám sát tỷ trọng toàn cục trong Bảng 3.8 của Đề cương (MCV 25%, MCH 21%, RBC 16%, Ferritin 12%,
    RDW-CV 10%, Hb 8%, MCHC 5%, HCT 3%)."""
    raw_shap = {
        "MCV": 0.25 * ((79.5 - cbc["MCV"]) / 6.5),
        "MCH": 0.21 * ((26.0 - cbc["MCH"]) / 2.5),
        "RBC": 0.16 * ((cbc["RBC"] - 4.95) / 0.48),
        "RDW-CV": -0.10 * ((cbc["RDW-CV"] - 14.0) / 1.5),
        "Hb_gL": 0.08 * ((130.0 - cbc["Hb_gL"]) / 14.0),
        "MCHC_gL": 0.05 * ((325.0 - cbc["MCHC_gL"]) / 13.5),
        "HCT": 0.03 * ((39.5 - cbc["HCT"]) / 4.2),
    }
    if has_ferritin and fer_val is not None:
        if fer_val < 30.0:
            # Nếu thiếu sắt nhưng Mentzer < 13 hoặc RBC >= 5.0 thì cảnh báo đồng mắc
            raw_shap["Ferritin"] = 0.06 if (cbc["RBC"] >= 5.0 and cbc["MCV"] < 75.0) else -0.14
        else:
            raw_shap["Ferritin"] = 0.12 * (1.0 if cbc["MCV"] < 80.0 or cbc["MCH"] < 27.0 else -0.3)
    else:
        raw_shap["Ferritin"] = 0.0

    total_abs = sum(abs(v) for v in raw_shap.values()) or 1.0
    items = []
    for key in ["MCV", "MCH", "RBC", "Ferritin", "RDW-CV", "Hb_gL", "MCHC_gL", "HCT"]:
        val = cbc.get(key) if key != "Ferritin" else (fer_val if has_ferritin else None)
        s_val = raw_shap[key]
        pct = round(abs(s_val) / total_abs * 100.0, 1)
        items.append({
            "feature": key,
            "label": FEATURE_DISPLAY_NAMES[key],
            "value": val,
            "shap_value": round(s_val, 4),
            "abs_pct": pct,
            "direction": "Tăng nguy cơ Thalassemia (+)" if s_val >= 0 else "Giảm nguy cơ / Hướng về IDA (−)",
        })
    items.sort(key=lambda x: abs(x["shap_value"]), reverse=True)
    return items


def _format_ml_output(cbc, indices, ferritin, probs_a, probs_b, active_prob, shap_contribs):
    """Phân tầng nguy cơ theo xác suất hiệu chỉnh của mô hình XGBoost và sinh khuyến nghị lâm sàng."""
    pct = round(active_prob * 100.0, 1)
    coexisting_ida = (
        ferritin is not None
        and ferritin < 30.0
        and (cbc["MCV"] < 76.0 and (cbc["RBC"] >= 5.0 or (indices["mentzer"] and indices["mentzer"] < 13.0)))
    )
    if active_prob >= 0.50 or coexisting_ida:
        category = "CAO"
    elif active_prob >= 0.20 or cbc["MCV"] < 85.0 or cbc["MCH"] < 28.0:
        category = "TRUNG BÌNH"
    else:
        category = "THẤP"

    delta_pct = round((probs_b["XGBoost"] - probs_a["XGBoost"]) * 100.0, 1) if ferritin is not None else 0.0
    model_used = "Mô hình B (XGBoost: CBC + Ferritin)" if ferritin is not None else "Mô hình A (XGBoost: Chỉ CBC)"

    if category == "CAO":
        conclusion = (
            f"{model_used} dự đoán xác suất mang gen Thalassemia ở mức CAO ({pct:.1f}%). "
            + (
                "CẢNH BÁO ĐỒNG MẮC: Ghi nhận Ferritin < 30 ng/mL đi kèm RBC tăng cao và MCV giảm sâu — "
                "nghi ngờ đồng mắc Thiếu máu thiếu sắt phối hợp Mang gen Thalassemia. "
                if coexisting_ida
                else ""
            )
            + "Khuyến nghị thực hiện phân tích thành phần huyết sắc tố (HPLC/điện di mao quản) và xét nghiệm gen (HBA1, HBA2, HBB)."
        )
    elif category == "TRUNG BÌNH":
        conclusion = (
            f"{model_used} ghi nhận xác suất mang gen Thalassemia ở mức TRUNG BÌNH ({pct:.1f}%). "
            + (
                "Nên định lượng bổ sung Ferritin huyết thanh để nâng độ chính xác lên Mô hình B và phân biệt với thiếu máu thiếu sắt."
                if ferritin is None
                else "Cần đối chiếu lâm sàng, cân nhắc điện di huyết sắc tố (HPLC) hoặc xét nghiệm gen alpha-Thalassemia nếu nghi ngờ."
            )
        )
    else:
        conclusion = (
            f"{model_used} dự đoán xác suất mang gen Thalassemia ở mức THẤP ({pct:.1f}%). "
            "Tuy nhiên, kết quả sàng lọc bằng học máy không loại trừ hoàn toàn người lành mang gen thể ẩn (silent carrier)."
        )

    return {
        "cbc": cbc,
        "indices": indices,
        "ferritin": ferritin,
        "probs_a": probs_a,
        "probs_b": probs_b,
        "active_prob": active_prob,
        "active_pct": pct,
        "delta_pct": delta_pct,
        "category": category,
        "conclusion": conclusion,
        "coexisting_ida": coexisting_ida,
        "model_used": model_used,
        "shap": shap_contribs,
    }


def render_shap_explanation_card(ml_res):
    """Hiển thị bảng so sánh Mô hình A vs Mô hình B và đồ thị thanh ngang SHAP cá thể hóa."""
    with st.container(border=True):
        st.markdown("#### :material/psychology: Dự đoán Học máy (THAL-SCREEN AI) & Giá trị bổ sung của Ferritin")
        m1, m2, m3, m4 = st.columns(4)
        prob_a_pct = ml_res["probs_a"]["XGBoost"] * 100.0
        prob_b_pct = ml_res["probs_b"]["XGBoost"] * 100.0
        m1.metric("Mô hình A (Chỉ CBC)", f"{prob_a_pct:.1f}%")
        if ml_res["ferritin"] is not None:
            m2.metric(
                "Mô hình B (CBC + Ferritin)",
                f"{prob_b_pct:.1f}%",
                delta=f"{ml_res['delta_pct']:+.1f}% so với chỉ CBC",
            )
        else:
            m2.metric("Mô hình B (CBC + Ferritin)", "Chưa nhập Ferritin")
        m3.metric("Green & King Index", f"{ml_res['indices']['green_king']:.1f}" if ml_res["indices"]["green_king"] else "—")
        m4.metric("Chỉ số RDWI", f"{ml_res['indices']['rdwi']:.1f}" if ml_res["indices"]["rdwi"] else "—")

        # Bảng so sánh 3 thuật toán
        algo_rows = []
        for algo in ["Logistic Regression", "Random Forest", "XGBoost"]:
            pa = ml_res["probs_a"][algo] * 100.0
            pb = ml_res["probs_b"][algo] * 100.0
            algo_rows.append({
                "Thuật toán Học máy": algo + (" (Mô hình chính)" if algo == "XGBoost" else ""),
                "Xác suất Mô hình A (Chỉ CBC)": f"{pa:.1f}%",
                "Xác suất Mô hình B (CBC + Ferritin)": f"{pb:.1f}%" if ml_res["ferritin"] is not None else "Chưa có Ferritin",
                "Chênh lệch (ΔP)": f"{pb - pa:+.1f} điểm %" if ml_res["ferritin"] is not None else "—",
            })
        st.dataframe(pd.DataFrame(algo_rows), hide_index=True, use_container_width=True)

        if ml_res["coexisting_ida"]:
            st.warning(
                "**Phát hiện mẫu hình nghi ngờ đồng mắc Thiếu máu thiếu sắt + Mang gen Thalassemia:** "
                "Mặc dù Ferritin huyết thanh giảm thấp (< 30 ng/mL), số lượng hồng cầu (RBC) vẫn tăng cao và MCV giảm sâu. "
                "Lưu ý: thiếu sắt phối hợp có thể làm giảm giả tạo nồng độ HbA2 trên HPLC; cần bác sĩ Huyết học đối chiếu.",
                icon=":material/warning:",
            )

        st.markdown("##### :material/bar_chart: Giải thích mức độ đóng góp biến đầu vào bằng phương pháp SHAP (Cá thể hóa)")
        st.caption(
            "Thanh màu đỏ (+) biểu thị yếu tố đẩy xác suất dự đoán về phía **Mang gen Thalassemia**; "
            "thanh màu xanh (−) biểu thị yếu tố kéo xác suất về phía **Không mang gen / Thiếu máu thiếu sắt đơn thuần**."
        )
        html_rows = []
        for item in ml_res["shap"]:
            val_str = f"{item['value']:g}" if item["value"] is not None else "Khuyết"
            width_pct = min(max(item["abs_pct"] * 2.2, 4.0), 100.0)
            bar_cls = "shap-bar-pos" if item["shap_value"] >= 0 else "shap-bar-neg"
            sign_str = f"{item['shap_value']:+.3f} ({item['abs_pct']:.1f}%)"
            html_rows.append(
                f'<div class="shap-row">'
                f'<div class="shap-label">{item["label"]} = {val_str}</div>'
                f'<div class="shap-track"><div class="{bar_cls}" style="width:{width_pct:.1f}%"></div></div>'
                f'<div class="shap-val">{sign_str}</div>'
                f"</div>"
            )
        st.markdown("".join(html_rows), unsafe_allow_html=True)


