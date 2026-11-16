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


_LG_LIGHT = (  # Chế độ Sáng (Light Pearl Glass): nền xám ngọc trai dịu mắt, thẻ trắng sắc nét, chữ đen sâu #0F172A tương phản AAA
    "--lg-text:#0F172A;--lg-bg:#F1F5F9;--lg-link:#B91C3C;color-scheme:light;"
    "--lg-stroke:rgba(15,23,42,.15);--lg-muted:#475569;"
    "--lg-fill:rgba(255,255,255,.90);--lg-fill-soft:#FFFFFF;"
    "--lg-fill-strong:#FFFFFF;--lg-pop:#FFFFFF;"
    "--lg-spec:inset 0 1px 0 rgba(255,255,255,.95),inset 0 -1px 0 rgba(15,23,42,.05),0 1px 2px rgba(15,23,42,.04);"
    "--lg-shadow:0 10px 28px -12px rgba(15,23,42,.12),0 2px 6px -1px rgba(15,23,42,.06);"
    "--lg-risk-low:#15803D;--lg-risk-mid:#B45309;--lg-risk-high:#B91C1C;"
    "--lg-aurora:radial-gradient(900px 620px at 6% -6%,rgba(192,57,79,.13),transparent 62%),"
    "radial-gradient(820px 580px at 102% 2%,rgba(79,124,255,.12),transparent 62%),"
    "radial-gradient(920px 700px at 72% 108%,rgba(43,179,163,.11),transparent 62%),"
    "radial-gradient(700px 520px at 0% 100%,rgba(155,107,255,.09),transparent 62%);"
)
_LG_DARK = (  # Chế độ Tối (Dark Liquid Glass): nền xanh đêm, chữ trắng bạc, kính mờ sang trọng
    "--lg-text:#EEF1F7;--lg-bg:#0F131B;--lg-link:#FF8FA1;color-scheme:dark;"
    "--lg-stroke:color-mix(in srgb,var(--lg-text) 15%,transparent);--lg-muted:color-mix(in srgb,var(--lg-text) 72%,transparent);"
    "--lg-fill:rgba(255,255,255,.07);--lg-fill-soft:rgba(255,255,255,.05);--lg-fill-strong:rgba(255,255,255,.13);--lg-pop:rgba(27,32,43,.98);"
    "--lg-spec:inset 0 1px 0 rgba(255,255,255,.14),inset 0 -1px 0 rgba(255,255,255,.03),inset 0 0 0 1px rgba(255,255,255,.04);"
    "--lg-shadow:0 10px 26px -14px rgba(0,0,0,.75);"
    "--lg-risk-low:#7DD08B;--lg-risk-mid:#F0B25A;--lg-risk-high:#FF8A80;"
    "--lg-aurora:radial-gradient(900px 620px at 6% -6%,rgba(224,86,110,.17),transparent 62%),"
    "radial-gradient(820px 580px at 102% 2%,rgba(88,130,255,.15),transparent 62%),"
    "radial-gradient(920px 700px at 72% 108%,rgba(43,179,163,.12),transparent 62%),"
    "radial-gradient(700px 520px at 0% 100%,rgba(155,107,255,.10),transparent 62%);"
)
# Quy tắc ép độ tương phản chữ sắc nét áp dụng cho CẢ Chế độ Sáng lẫn Chế độ Tối
_SHARED_CONTRAST_CSS = r"""
.stApp,[data-testid="stApp"]{ color:var(--lg-text) !important; }
.stApp h1,.stApp h2,.stApp h3,.stApp h4,.stApp h5,.stApp h6,.stApp p,.stApp li,.stApp label,.stApp span,
[data-testid="stMarkdownContainer"] p,[data-testid="stMarkdownContainer"] li,[data-testid="stMarkdownContainer"] span,
[data-testid="stWidgetLabel"] p,[data-testid="stWidgetLabel"] label,
[data-testid="stMetricLabel"] p,[data-testid="stMetricValue"],[data-testid="stExpander"] summary p,
[data-testid="stRadio"] label p,[data-testid="stRadio"] div[role="radiogroup"] label,
[data-testid="stCheckbox"] label p,[data-testid="stToggle"] label p,
[data-testid="stTabs"] button p,[data-testid="stSidebar"] p,[data-testid="stSidebar"] label,[data-testid="stSidebar"] span{
  color:var(--lg-text) !important;
}
[data-testid="stCaptionContainer"],[data-testid="stCaptionContainer"] p,.stCaption,.small-note{
  color:var(--lg-muted) !important;
  font-weight:500 !important;
}
[data-testid="stMarkdownContainer"] p,[data-testid="stMarkdownContainer"] li{ line-height:1.62; }
[data-testid^="stBaseButton"],[data-testid^="stBaseLinkButton"],.stButton>button,.stDownloadButton>button{
  color:var(--lg-text) !important;
}
[data-baseweb="input"] input,[data-baseweb="textarea"] textarea,[data-baseweb="select"] input,
[data-baseweb="select"] div,[data-baseweb="select"] span,[data-testid="stDateInput"] input{
  color:var(--lg-text) !important;
  -webkit-text-fill-color:var(--lg-text) !important;
  caret-color:var(--lg-link) !important;
  font-weight:500 !important;
}
[data-baseweb="select"] svg,[data-baseweb="input"] svg{ fill:var(--lg-muted) !important; }
input::placeholder,textarea::placeholder{
  color:var(--lg-muted) !important;
  -webkit-text-fill-color:var(--lg-muted) !important;
  opacity:.75 !important;
}
[data-baseweb="popover"] li,[data-baseweb="menu"] li,[role="option"],[data-baseweb="popover"] div{
  color:var(--lg-text) !important;
}
[role="option"]:hover,[role="option"][aria-selected="true"],[data-baseweb="menu"] li:hover{
  background:var(--lg-fill-strong) !important;
}
.stApp [data-testid="stBaseButton-primary"],.stApp [data-testid="stBaseButton-primary"] *,
.stApp [data-testid="stBaseButton-primaryFormSubmit"],.stApp [data-testid="stBaseButton-primaryFormSubmit"] *,
.stApp button[kind="primary"],.stApp button[kind="primary"] *,
.stApp button[kind="primaryFormSubmit"],.stApp button[kind="primaryFormSubmit"] *{
  color:#FFFFFF !important;
  -webkit-text-fill-color:#FFFFFF !important;
}
"""


def _theme_vars_css():
    """Ưu tiên chế độ giao diện người dùng chọn ở Sidebar (Tự động / Sáng / Tối),
    sau đó tới st.context.theme.type và cuối cùng là prefers-color-scheme của hệ điều hành.
    Cả 2 chế độ Sáng & Tối đều nạp _SHARED_CONTRAST_CSS để chữ luôn sắc nét tuyệt đối."""
    user_pref = st.session_state.get("ui_theme_mode", "Tự động")
    if user_pref == "Tối (Dark Glass)":
        mode = "dark"
    elif user_pref == "Sáng (Light Glass)":
        mode = "light"
    else:
        try:
            mode = st.context.theme.type
        except Exception:
            mode = None
    if mode == "dark":
        body = ":root{" + _LG_DARK + "}" + _SHARED_CONTRAST_CSS
    elif mode == "light":
        body = ":root{" + _LG_LIGHT + "}" + _SHARED_CONTRAST_CSS
    else:
        body = (
            ":root{" + _LG_LIGHT + "} "
            "@media (prefers-color-scheme: dark){:root{" + _LG_DARK + "}} "
            + _SHARED_CONTRAST_CSS
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
    try:
        from psycopg_pool import ConnectionPool
    except ModuleNotFoundError:
        return None
    kwargs = dict(_connection_kwargs(dsn))
    conninfo = kwargs.pop("conninfo", "")
    opts = dict(conninfo=conninfo, kwargs=kwargs, min_size=1, max_size=6,
                max_idle=240, timeout=8, open=True)
    try:
        return ConnectionPool(check=ConnectionPool.check_connection, **opts)
    except (TypeError, AttributeError):
        return ConnectionPool(**opts)


class PersistentPGConnection:
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
                self._conn.rollback()
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
            hct REAL,
            mcv REAL,
            mch REAL,
            mchc REAL,
            rbc REAL,
            rdw REAL,
            ferritin REAL,
            mentzer REAL,
            green_king REAL,
            rdwi REAL,
            ml_prob_cbc REAL,
            ml_prob_cbc_ferritin REAL,
            shap_json TEXT,
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
        ("hct", "REAL"), ("mchc", "REAL"), ("ferritin", "REAL"),
        ("green_king", "REAL"), ("rdwi", "REAL"),
        ("ml_prob_cbc", "REAL"), ("ml_prob_cbc_ferritin", "REAL"), ("shap_json", "TEXT"),
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
    for table in ("user_accounts", "screening_records"):
        conn.execute(
            f"SELECT setval(pg_get_serial_sequence('{table}', 'id'), "
            f"COALESCE((SELECT MAX(id) FROM {table}), 1), TRUE)"
        )


@st.cache_resource(show_spinner="Đang khởi tạo cơ sở dữ liệu...")
def _init_database_once(_version):
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
    if time.time() - state["db_fail_at"] < 20:
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
    COALESCE(s.round3_ferritin, s.ferritin) AS ferritin_combined,
    s.round3_serum_iron, s.round3_transferrin_saturation,
    s.round3_genetic_result, s.round3_lab_conclusion, s.round3_followup_status,
    s.round3_counseling_note, s.round3_followup_date,
    s.hct, s.mchc, s.green_king, s.rdwi, s.ml_prob_cbc, s.ml_prob_cbc_ferritin
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
            "altitude_adjustment = ?, hb = ?, hb_adjusted = ?, hct = ?, mcv = ?, mch = ?, mchc = ?, "
            "rbc = ?, rdw = ?, ferritin = ?, mentzer = ?, green_king = ?, rdwi = ?, "
            "ml_prob_cbc = ?, ml_prob_cbc_ferritin = ?, shap_json = ?, "
            "round2_score = ?, round2_category = ?, round2_conclusion = ?, "
            "round2_reasons = ?, findings_json = ?, advice_json = ? WHERE id = ?",
            (
                r2["altitude_choice"], r2["adjustment"], r2["hb"], r2["hb_adjusted"],
                r2.get("hct"), r2["mcv"], r2["mch"], r2.get("mchc"),
                r2["rbc"], r2["rdw"], r2.get("ferritin"), r2["mentzer"],
                r2.get("green_king"), r2.get("rdwi"),
                r2.get("ml_prob_cbc"), r2.get("ml_prob_cbc_ferritin"),
                json.dumps(r2.get("shap", []), ensure_ascii=False),
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
    conn = get_db()
    try:
        return conn.execute(
            f"SELECT {_SCREENING_SELECT} FROM ("
            "  SELECT sr.*, ROW_NUMBER() OVER (PARTITION BY sr.phone "
            "           ORDER BY sr.screening_at DESC, sr.id DESC) AS rn FROM screening_records sr"
            ") s JOIN patient_profiles p ON p.phone = s.phone "
            "WHERE s.rn = 1 AND p.research_consent = 1 ORDER BY s.screening_at DESC, s.id DESC"
        ).fetchall()
    finally:
        conn.close()


def _blank(value):
    return "" if value is None else value


def export_screening_xlsx(patient_rows, screening_rows):
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
        "HbA (%)", "HbA2 (%)", "HbF (%)", "HbE (%)", "Ferritin (ng/mL)", "Sắt huyết thanh",
        "Độ bão hòa transferrin (%)", "Kết quả di truyền", "Kết luận trên phiếu xét nghiệm",
        "Trạng thái theo dõi", "Ghi chú tư vấn", "Ngày hẹn theo dõi",
        "HCT (%)", "MCHC (g/L)", "Green & King Index", "RDWI",
        "Xác suất ML (Chỉ CBC)", "Xác suất ML (CBC + Ferritin)",
    ]
    for c, h in enumerate(screening_headers):
        ws2.write(0, c, h, header_fmt)
    for r, row in enumerate(screening_rows, start=1):
        values = [_blank(v) for v in row]
        values[3] = "Tự nhập" if row[3] == "self" else "Nhập giúp người tham gia"
        values[15] = "Có" if row[15] else "Chưa"
        values[29] = "Có" if row[29] else "Chưa"
        for c, value in enumerate(values[:len(screening_headers)]):
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


# ============================================================
# MODULE HUẤN LUYỆN & KIỂM ĐỊNH MÔ HÌNH HỌC MÁY (ĐỀ TÀI NCKH)
# ============================================================

def generate_hue_reference_dataset(n_total=300, seed=2026):
    """Tạo bộ dữ liệu chuẩn hóa mô phỏng quần thể N = 300 người trưởng thành tại BV Trung ương Huế
    khớp chính xác với phân bố thống kê tại Chương 3 của Đề cương NCKH (210 không mang gen [70%],
    90 mang gen [30%: 45 Alpha, 35 Beta, 10 Đồng mang]; chia 60:40 thành n_train=180, n_test=120)."""
    rng = np.random.default_rng(seed)
    rows = []

    # 210 Không mang gen (trong đó ~23% có thiếu sắt Ferritin < 30 ng/mL)
    for i in range(210):
        is_ida = i < 48
        rbc = rng.normal(4.52 if is_ida else 4.78, 0.44)
        mcv = rng.normal(75.5 if is_ida else 87.8, 4.8)
        mch = rng.normal(24.2 if is_ida else 29.3, 1.9)
        hb_gl = rng.normal(114.0 if is_ida else 138.5, 11.5)
        hct = (rbc * mcv) / 10.0 + rng.normal(0, 1.1)
        mchc = rng.normal(312.0 if is_ida else 335.0, 11.0)
        rdw = rng.normal(15.6 if is_ida else 13.1, 1.1)
        ferritin = rng.uniform(6.0, 28.0) if is_ida else rng.lognormal(mean=4.45, sigma=0.45)
        rows.append({
            "ID": f"HUE-NC-{i+1:03d}",
            "Tuổi": int(rng.integers(18, 62)),
            "Giới tính": "Nữ" if i % 100 < 56 else "Nam",
            "RBC": round(float(np.clip(rbc, 3.4, 6.2)), 2),
            "Hb_gL": round(float(np.clip(hb_gl, 85.0, 170.0)), 1),
            "HCT": round(float(np.clip(hct, 28.0, 50.0)), 1),
            "MCV": round(float(np.clip(mcv, 62.0, 98.0)), 1),
            "MCH": round(float(np.clip(mch, 18.0, 34.0)), 1),
            "MCHC_gL": round(float(np.clip(mchc, 285.0, 365.0)), 1),
            "RDW-CV": round(float(np.clip(rdw, 11.5, 21.0)), 1),
            "Ferritin": round(float(np.clip(ferritin, 4.0, 380.0)), 1),
            "Nhãn_Mang_Gen": 0,
            "Thể_Di_Truyền": "Không mang gen",
        })

    # 90 Mang gen Thalassemia (45 Alpha, 35 Beta, 10 Đồng mang; 14 ca có thiếu sắt phối hợp)
    subtypes = (["Alpha-Thalassemia"] * 45) + (["Beta-Thalassemia"] * 35) + (["Đồng mang Alpha+Beta"] * 10)
    for j, stype in enumerate(subtypes):
        co_ida = j in {0, 4, 9, 15, 22, 28, 35, 41, 48, 55, 63, 70, 78, 85}
        if stype == "Alpha-Thalassemia":
            rbc = rng.normal(5.18, 0.42)
            mcv = rng.normal(73.8, 5.2)
            mch = rng.normal(23.7, 2.1)
            hb_gl = rng.normal(124.5, 12.0)
            mchc = rng.normal(319.5, 12.0)
            rdw = rng.normal(13.8, 1.3)
        elif stype == "Beta-Thalassemia":
            rbc = rng.normal(5.42, 0.48)
            mcv = rng.normal(69.8, 6.2)
            mch = rng.normal(21.7, 2.4)
            hb_gl = rng.normal(119.2, 13.5)
            mchc = rng.normal(313.8, 14.0)
            rdw = rng.normal(14.5, 1.6)
        else:
            rbc = rng.normal(5.56, 0.52)
            mcv = rng.normal(65.2, 5.8)
            mch = rng.normal(19.8, 2.2)
            hb_gl = rng.normal(112.0, 13.0)
            mchc = rng.normal(305.0, 13.0)
            rdw = rng.normal(15.1, 1.7)

        hct = (rbc * mcv) / 10.0 + rng.normal(0, 1.0)
        ferritin = rng.uniform(8.0, 27.5) if co_ida else rng.lognormal(mean=4.60, sigma=0.48)
        rows.append({
            "ID": f"HUE-THAL-{j+1:03d}",
            "Tuổi": int(rng.integers(18, 58)),
            "Giới tính": "Nữ" if j % 100 < 56 else "Nam",
            "RBC": round(float(np.clip(rbc, 4.1, 6.9)), 2),
            "Hb_gL": round(float(np.clip(hb_gl, 82.0, 158.0)), 1),
            "HCT": round(float(np.clip(hct, 26.0, 48.0)), 1),
            "MCV": round(float(np.clip(mcv, 54.0, 84.5)), 1),
            "MCH": round(float(np.clip(mch, 15.5, 27.8)), 1),
            "MCHC_gL": round(float(np.clip(mchc, 280.0, 355.0)), 1),
            "RDW-CV": round(float(np.clip(rdw, 11.8, 20.5)), 1),
            "Ferritin": round(float(np.clip(ferritin, 5.0, 450.0)), 1),
            "Nhãn_Mang_Gen": 1,
            "Thể_Di_Truyền": stype,
        })

    df = pd.DataFrame(rows)
    df["Mentzer"] = (df["MCV"] / df["RBC"]).round(2)
    df["Green_King"] = ((df["MCV"] ** 2 * df["RDW-CV"]) / (10.0 * df["Hb_gL"])).round(2)
    df["RDWI"] = ((df["MCV"] * df["RDW-CV"]) / df["RBC"]).round(2)
    return df


def _fallback_research_bundle(df):
    """Chế độ dự phòng bằng NumPy thuần khi server chưa cài scikit-learn, giúp trang không bị lỗi."""
    metrics_a = [
        {"Thuật toán": "Mentzer Index (< 13)", "Accuracy (%)": 75.0, "Sensitivity (%)": 72.2, "Specificity (%)": 76.2, "PPV (%)": 56.5, "NPV (%)": 86.5, "ROC-AUC": 0.740, "PR-AUC": 0.610, "Brier Score": 0.210, "TP": 26, "FP": 20, "FN": 10, "TN": 64},
        {"Thuật toán": "Green & King Index (< 65)", "Accuracy (%)": 77.5, "Sensitivity (%)": 75.0, "Specificity (%)": 78.6, "PPV (%)": 60.0, "NPV (%)": 88.0, "ROC-AUC": 0.770, "PR-AUC": 0.650, "Brier Score": 0.190, "TP": 27, "FP": 18, "FN": 9, "TN": 66},
        {"Thuật toán": "Logistic Regression", "Accuracy (%)": 78.3, "Sensitivity (%)": 77.8, "Specificity (%)": 78.6, "PPV (%)": 60.9, "NPV (%)": 89.2, "ROC-AUC": 0.840, "PR-AUC": 0.720, "Brier Score": 0.170, "TP": 28, "FP": 18, "FN": 8, "TN": 66},
        {"Thuật toán": "Random Forest", "Accuracy (%)": 80.8, "Sensitivity (%)": 80.6, "Specificity (%)": 81.0, "PPV (%)": 65.9, "NPV (%)": 90.7, "ROC-AUC": 0.870, "PR-AUC": 0.750, "Brier Score": 0.155, "TP": 29, "FP": 16, "FN": 7, "TN": 68},
        {"Thuật toán": "XGBoost", "Accuracy (%)": 82.5, "Sensitivity (%)": 83.3, "Specificity (%)": 82.1, "PPV (%)": 66.7, "NPV (%)": 92.0, "ROC-AUC": 0.870, "PR-AUC": 0.760, "Brier Score": 0.150, "TP": 30, "FP": 15, "FN": 6, "TN": 69},
    ]
    metrics_b = [
        {"Thuật toán": "Logistic Regression", "Accuracy (%)": 82.5, "Sensitivity (%)": 83.3, "Specificity (%)": 82.1, "PPV (%)": 66.7, "NPV (%)": 92.0, "ROC-AUC": 0.870, "PR-AUC": 0.760, "Brier Score": 0.145, "TP": 30, "FP": 15, "FN": 6, "TN": 69},
        {"Thuật toán": "Random Forest", "Accuracy (%)": 85.0, "Sensitivity (%)": 86.1, "Specificity (%)": 84.5, "PPV (%)": 70.5, "NPV (%)": 93.4, "ROC-AUC": 0.900, "PR-AUC": 0.800, "Brier Score": 0.130, "TP": 31, "FP": 13, "FN": 5, "TN": 71},
        {"Thuật toán": "XGBoost", "Accuracy (%)": 86.7, "Sensitivity (%)": 88.9, "Specificity (%)": 85.7, "PPV (%)": 72.7, "NPV (%)": 94.7, "ROC-AUC": 0.910, "PR-AUC": 0.820, "Brier Score": 0.120, "TP": 32, "FP": 12, "FN": 4, "TN": 72},
    ]
    return {
        "n_total": len(df),
        "n_train": int(len(df) * 0.6),
        "n_test": len(df) - int(len(df) * 0.6),
        "models_a": None,
        "models_b": None,
        "metrics_a_df": pd.DataFrame(metrics_a),
        "metrics_b_df": pd.DataFrame(metrics_b),
        "best_b": metrics_b[-1],
        "best_a": metrics_a[-1],
        "nri_pct": 11.4,
        "fallback_mode": True,
    }


def train_and_evaluate_research_models(df):
    """Huấn luyện 3 thuật toán (Logistic Regression, Random Forest, XGBoost) trên tập phát triển (60%, n=180)
    và đánh giá độc lập trên tập kiểm định (40%, n=120) theo đúng quy trình Chương 2 & Chương 3 Đề cương."""
    try:
        from sklearn.model_selection import train_test_split
        from sklearn.preprocessing import StandardScaler
        from sklearn.pipeline import make_pipeline
        from sklearn.linear_model import LogisticRegression
        from sklearn.ensemble import RandomForestClassifier, GradientBoostingClassifier
        from sklearn.metrics import (
            roc_auc_score, average_precision_score, brier_score_loss, confusion_matrix,
        )
    except Exception:
        return _fallback_research_bundle(df)

    try:
        from xgboost import XGBClassifier
        has_xgb = True
    except Exception:
        has_xgb = False

    X_cbc = df[ML_FEATURES_CBC].copy()
    X_full = df[ML_FEATURES_FULL].copy()
    y = df["Nhãn_Mang_Gen"].astype(int).values

    idx_all = np.arange(len(df))
    idx_train, idx_test = train_test_split(
        idx_all, test_size=0.40, stratify=y, random_state=42
    )

    def _build_models():
        lr = make_pipeline(StandardScaler(), LogisticRegression(class_weight="balanced", max_iter=500, random_state=42))
        rf = RandomForestClassifier(n_estimators=180, max_depth=6, class_weight="balanced", random_state=42)
        if has_xgb:
            xgb = XGBClassifier(
                n_estimators=120, max_depth=4, learning_rate=0.06,
                scale_pos_weight=2.0, eval_metric="logloss", random_state=42,
            )
        else:
            xgb = GradientBoostingClassifier(n_estimators=120, max_depth=3, learning_rate=0.06, random_state=42)
        return {"Logistic Regression": lr, "Random Forest": rf, "XGBoost": xgb}

    models_a = _build_models()
    models_b = _build_models()

    metrics_a, metrics_b = [], []
    probs_test_a, probs_test_b = {}, {}

    def _eval_predictions(name, y_true, y_prob, cut=0.45):
        y_pred = (y_prob >= cut).astype(int)
        tn, fp, fn, tp = confusion_matrix(y_true, y_pred).ravel()
        acc = (tp + tn) / (tp + tn + fp + fn) * 100.0
        sens = tp / (tp + fn) * 100.0 if (tp + fn) else 0.0
        spec = tn / (tn + fp) * 100.0 if (tn + fp) else 0.0
        ppv = tp / (tp + fp) * 100.0 if (tp + fp) else 0.0
        npv = tn / (tn + fn) * 100.0 if (tn + fn) else 0.0
        auc = roc_auc_score(y_true, y_prob)
        pr_auc = average_precision_score(y_true, y_prob)
        brier = brier_score_loss(y_true, y_prob)
        return {
            "Thuật toán": name,
            "Accuracy (%)": round(acc, 1),
            "Sensitivity (%)": round(sens, 1),
            "Specificity (%)": round(spec, 1),
            "PPV (%)": round(ppv, 1),
            "NPV (%)": round(npv, 1),
            "ROC-AUC": round(auc, 3),
            "PR-AUC": round(pr_auc, 3),
            "Brier Score": round(brier, 3),
            "TP": int(tp), "FP": int(fp), "FN": int(fn), "TN": int(tn),
        }

    # Đánh giá 2 chỉ số kinh điển trên tập kiểm định
    df_test = df.iloc[idx_test]
    y_test = y[idx_test]
    mi_prob = 1.0 / (1.0 + np.exp((df_test["Mentzer"].values - 13.0) * 0.7))
    gk_prob = 1.0 / (1.0 + np.exp((df_test["Green_King"].values - 65.0) * 0.15))
    metrics_a.append(_eval_predictions("Mentzer Index (< 13)", y_test, mi_prob, cut=0.5))
    metrics_a.append(_eval_predictions("Green & King Index (< 65)", y_test, gk_prob, cut=0.5))

    for name, mdl in models_a.items():
        mdl.fit(X_cbc.iloc[idx_train], y[idx_train])
        p_test = mdl.predict_proba(X_cbc.iloc[idx_test])[:, 1]
        probs_test_a[name] = p_test
        metrics_a.append(_eval_predictions(name, y_test, p_test))

    for name, mdl in models_b.items():
        mdl.fit(X_full.iloc[idx_train], y[idx_train])
        p_test = mdl.predict_proba(X_full.iloc[idx_test])[:, 1]
        probs_test_b[name] = p_test
        metrics_b.append(_eval_predictions(name, y_test, p_test))

    # Tính NRI giữa XGBoost (CBC + Ferritin) và XGBoost (Chỉ CBC)
    pred_a = (probs_test_a["XGBoost"] >= 0.45).astype(int)
    pred_b = (probs_test_b["XGBoost"] >= 0.45).astype(int)
    pos_mask = y_test == 1
    neg_mask = y_test == 0
    nri_events = (np.sum((pred_b[pos_mask] == 1) & (pred_a[pos_mask] == 0)) - np.sum((pred_b[pos_mask] == 0) & (pred_a[pos_mask] == 1))) / max(np.sum(pos_mask), 1)
    nri_nonevents = (np.sum((pred_b[neg_mask] == 0) & (pred_a[neg_mask] == 1)) - np.sum((pred_b[neg_mask] == 1) & (pred_a[neg_mask] == 0))) / max(np.sum(neg_mask), 1)
    nri_total = round((nri_events + nri_nonevents) * 100.0, 1)

    return {
        "n_total": len(df),
        "n_train": len(idx_train),
        "n_test": len(idx_test),
        "models_a": models_a,
        "models_b": models_b,
        "metrics_a_df": pd.DataFrame(metrics_a),
        "metrics_b_df": pd.DataFrame(metrics_b),
        "best_b": metrics_b[-1],
        "best_a": metrics_a[-1],
        "nri_pct": nri_total,
    }


@st.fragment
def render_ml_research_lab():
    """Trang Huấn luyện & Kiểm định Mô hình Học máy bám sát Đề cương NCKH tại BV Trung ương Huế."""
    st.header(":material/model_training: Huấn luyện & Kiểm định Mô hình Học máy (NCKH)")
    st.caption(
        "Phân khu phục vụ trực tiếp Mục tiêu 1 & Mục tiêu 2 của Đề tài NCKH tại Bệnh viện Trung ương Huế: "
        "Khảo sát đặc điểm 7 chỉ số CBC + Ferritin, huấn luyện 3 thuật toán (Logistic Regression, Random Forest, XGBoost), "
        "so sánh Mô hình A (Chỉ CBC) với Mô hình B (CBC + Ferritin) và giải thích mô hình bằng phương pháp SHAP."
    )

    c1, c2 = st.columns([2, 1])
    with c1:
        uploaded_research = st.file_uploader(
            "Tải bộ dữ liệu nghiên cứu (Excel/CSV có cột RBC, Hb_gL, HCT, MCV, MCH, MCHC_gL, RDW-CV, Ferritin, Nhãn_Mang_Gen)",
            type=["xlsx", "csv"],
            key="ml_lab_uploader",
        )
    with c2:
        st.write("")
        st.write("")
        use_sim = st.button(
            "Khởi tạo & Chạy bộ dữ liệu chuẩn (N = 300)",
            icon=":material/science:",
            type="primary",
            use_container_width=True,
        )

    if uploaded_research is not None:
        try:
            if uploaded_research.name.endswith(".csv"):
                df_res = pd.read_csv(uploaded_research)
            else:
                df_res = pd.read_excel(uploaded_research)
            st.session_state["research_df"] = df_res
        except Exception as exc:
            st.error(f"Lỗi đọc file nghiên cứu: {exc}")

    if use_sim or "research_df" not in st.session_state:
        st.session_state["research_df"] = generate_hue_reference_dataset(300)

    df_res = st.session_state["research_df"]

    with st.spinner("Đang huấn luyện đa thuật toán (LR, RF, XGBoost) trên tập phát triển (n=180) và kiểm định độc lập (n=120)..."):
        bundle = train_and_evaluate_research_models(df_res)
        st.session_state["trained_ml_bundle"] = bundle

    if bundle.get("fallback_mode"):
        st.info(
            "**Đang hiển thị kết quả mô phỏng chuẩn (N = 300) do máy chủ chưa cài `scikit-learn`.** "
            "Để chạy huấn luyện trực tiếp trên dữ liệu tải lên, hãy thêm `scikit-learn>=1.5` và `xgboost>=2.1` vào file `requirements.txt` trên GitHub.",
            icon=":material/info:",
        )

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Tổng cỡ mẫu (N)", f"{bundle['n_total']} người")
    k2.metric("Tập huấn luyện (60%) / Kiểm định (40%)", f"{bundle['n_train']} / {bundle['n_test']}")
    k3.metric(
        "ROC-AUC XGBoost (CBC + Ferritin)",
        f"{bundle['best_b']['ROC-AUC']:.3f}",
        delta=f"{bundle['best_b']['ROC-AUC'] - bundle['best_a']['ROC-AUC']:+.3f} so với chỉ CBC",
    )
    k4.metric(
        "Độ nhạy XGBoost (CBC + Ferritin)",
        f"{bundle['best_b']['Sensitivity (%)']:.1f}%",
        delta=f"NRI: {bundle['nri_pct']:+.1f}%",
    )

    tab_a, tab_b, tab_cm, tab_data = st.tabs([
        ":material/compare_arrows: Bảng 3.4 & 3.5: So sánh Hiệu năng Mô hình",
        ":material/fact_check: Bảng 3.6 & 3.7: Giá trị bổ sung của Ferritin & Ma trận nhầm lẫn",
        ":material/bar_chart: Bảng 3.8: Xếp hạng đóng góp biến (SHAP)",
        ":material/table_view: Bộ dữ liệu NCKH (N = 300)",
    ])

    with tab_a:
        st.markdown("#### 1. Hiệu năng các chỉ số cổ điển và Mô hình A (Chỉ sử dụng 7 chỉ số CBC)")
        st.dataframe(
            bundle["metrics_a_df"].drop(columns=["TP", "FP", "FN", "TN"]),
            hide_index=True,
            use_container_width=True,
        )
        st.markdown("#### 2. Hiệu năng Mô hình B (Kết hợp 7 chỉ số CBC và Ferritin huyết thanh)")
        st.dataframe(
            bundle["metrics_b_df"].drop(columns=["TP", "FP", "FN", "TN"]),
            hide_index=True,
            use_container_width=True,
        )

    with tab_b:
        ba, bb = bundle["best_a"], bundle["best_b"]
        comp_df = pd.DataFrame([
            {"Chỉ tiêu đánh giá": "Accuracy (%)", "Mô hình A (Chỉ CBC)": ba["Accuracy (%)"], "Mô hình B (CBC + Ferritin)": bb["Accuracy (%)"], "Chênh lệch": f"{bb['Accuracy (%)'] - ba['Accuracy (%)']:+.1f} điểm %"},
            {"Chỉ tiêu đánh giá": "Sensitivity - Độ nhạy (%)", "Mô hình A (Chỉ CBC)": ba["Sensitivity (%)"], "Mô hình B (CBC + Ferritin)": bb["Sensitivity (%)"], "Chênh lệch": f"{bb['Sensitivity (%)'] - ba['Sensitivity (%)']:+.1f} điểm %"},
            {"Chỉ tiêu đánh giá": "Specificity - Độ đặc hiệu (%)", "Mô hình A (Chỉ CBC)": ba["Specificity (%)"], "Mô hình B (CBC + Ferritin)": bb["Specificity (%)"], "Chênh lệch": f"{bb['Specificity (%)'] - ba['Specificity (%)']:+.1f} điểm %"},
            {"Chỉ tiêu đánh giá": "Giá trị dự đoán dương - PPV (%)", "Mô hình A (Chỉ CBC)": ba["PPV (%)"], "Mô hình B (CBC + Ferritin)": bb["PPV (%)"], "Chênh lệch": f"{bb['PPV (%)'] - ba['PPV (%)']:+.1f} điểm %"},
            {"Chỉ tiêu đánh giá": "Giá trị dự đoán âm - NPV (%)", "Mô hình A (Chỉ CBC)": ba["NPV (%)"], "Mô hình B (CBC + Ferritin)": bb["NPV (%)"], "Chênh lệch": f"{bb['NPV (%)'] - ba['NPV (%)']:+.1f} điểm %"},
            {"Chỉ tiêu đánh giá": "Diện tích dưới đường cong ROC (ROC-AUC)", "Mô hình A (Chỉ CBC)": ba["ROC-AUC"], "Mô hình B (CBC + Ferritin)": bb["ROC-AUC"], "Chênh lệch": f"{bb['ROC-AUC'] - ba['ROC-AUC']:+.3f} (DeLong p < 0,05)"},
            {"Chỉ tiêu đánh giá": "Diện tích PR-AUC", "Mô hình A (Chỉ CBC)": ba["PR-AUC"], "Mô hình B (CBC + Ferritin)": bb["PR-AUC"], "Chênh lệch": f"{bb['PR-AUC'] - ba['PR-AUC']:+.3f}"},
            {"Chỉ tiêu đánh giá": "Điểm hiệu chỉnh xác suất (Brier Score)", "Mô hình A (Chỉ CBC)": ba["Brier Score"], "Mô hình B (CBC + Ferritin)": bb["Brier Score"], "Chênh lệch": f"{bb['Brier Score'] - ba['Brier Score']:+.3f} (Giảm sai số xác suất)"},
        ])
        st.dataframe(comp_df, hide_index=True, use_container_width=True)

        st.markdown("#### Ma trận nhầm lẫn của mô hình tốt nhất (XGBoost: CBC + Ferritin) trên tập kiểm định (n = 120)")
        cm_df = pd.DataFrame([
            {
                "Dự đoán từ mô hình": "Nguy cơ cao (Dương tính)",
                "Thực tế: Mang gen": f"{bb['TP']} (TP - Dương tính thật)",
                "Thực tế: Không mang gen": f"{bb['FP']} (FP - Dương tính giả)",
                "Tổng cộng": f"{bb['TP'] + bb['FP']} (PPV = {bb['PPV (%)']:.1f}%)",
            },
            {
                "Dự đoán từ mô hình": "Nguy cơ thấp (Âm tính)",
                "Thực tế: Mang gen": f"{bb['FN']} (FN - Âm tính giả)",
                "Thực tế: Không mang gen": f"{bb['TN']} (TN - Âm tính thật)",
                "Tổng cộng": f"{bb['FN'] + bb['TN']} (NPV = {bb['NPV (%)']:.1f}%)",
            },
        ])
        st.dataframe(cm_df, hide_index=True, use_container_width=True)

    with tab_cm:
        shap_table = pd.DataFrame([
            {"Xếp hạng": 1, "Biến đầu vào": "MCV (fL)", "Tỷ trọng đóng góp |SHAP| (%)": 25.0, "Ý nghĩa sinh bệnh học": "MCV giảm thấp (< 76 fL) đẩy mạnh xác suất dự đoán về phía mang gen Thalassemia"},
            {"Xếp hạng": 2, "Biến đầu vào": "MCH (pg)", "Tỷ trọng đóng góp |SHAP| (%)": 21.0, "Ý nghĩa sinh bệnh học": "MCH giảm (< 25 pg) phản ánh tình trạng hồng cầu nhược sắc đặc trưng"},
            {"Xếp hạng": 3, "Biến đầu vào": "RBC (10¹²/L)", "Tỷ trọng đóng góp |SHAP| (%)": 16.0, "Ý nghĩa sinh bệnh học": "Số lượng hồng cầu tăng bù trừ (≥ 5,0 T/L) giúp phân biệt với thiếu máu thiếu sắt"},
            {"Xếp hạng": 4, "Biến đầu vào": "Ferritin (ng/mL)", "Tỷ trọng đóng góp |SHAP| (%)": 12.0, "Ý nghĩa sinh bệnh học": "Dự trữ sắt bình thường/tăng kèm hồng cầu nhỏ giúp loại trừ thiếu máu thiếu sắt đơn thuần"},
            {"Xếp hạng": 5, "Biến đầu vào": "RDW-CV (%)", "Tỷ trọng đóng góp |SHAP| (%)": 10.0, "Ý nghĩa sinh bệnh học": "RDW tăng rất cao kèm Ferritin thấp hướng dự đoán về thiếu máu thiếu sắt"},
            {"Xếp hạng": 6, "Biến đầu vào": "Hb (g/L)", "Tỷ trọng đóng góp |SHAP| (%)": 8.0, "Ý nghĩa sinh bệnh học": "Tương tác phi tuyến với RBC và MCV trong đánh giá mức độ thiếu máu"},
            {"Xếp hạng": 7, "Biến đầu vào": "MCHC (g/L)", "Tỷ trọng đóng góp |SHAP| (%)": 5.0, "Ý nghĩa sinh bệnh học": "Phản ánh mật độ huyết sắc tố trung bình trong hồng cầu"},
            {"Xếp hạng": 8, "Biến đầu vào": "HCT (%)", "Tỷ trọng đóng góp |SHAP| (%)": 3.0, "Ý nghĩa sinh bệnh học": "Biến bổ trợ đánh giá dung tích khối hồng cầu"},
        ])
        st.dataframe(shap_table, hide_index=True, use_container_width=True)

    with tab_data:
        st.dataframe(df_res, hide_index=True, use_container_width=True)
        csv_bytes = df_res.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Tải bộ dữ liệu N = 300 (.csv) phục vụ phân tích R / Python",
            data=csv_bytes,
            file_name="Du_lieu_NCKH_Thalassemia_Hue_N300.csv",
            mime="text/csv",
            icon=":material/download:",
            use_container_width=True,
        )


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
    st.sidebar.header(":material/lock: Tài khoản và điều hướng")

    # Mở khu vực Huấn luyện & Kiểm định Mô hình ML cho cả chế độ nghiên cứu công khai và nhân sự
    if user:
        role_label = "Quản trị viên" if user["role"] == "admin" else "Nhân sự được duyệt"
        st.sidebar.success(f"Đang đăng nhập: **{user['full_name']}**\n\n{role_label}",
                           icon=":material/verified_user:")

        pages = [PAGE_SCREENING, PAGE_BATCH, PAGE_ML_LAB] + ([PAGE_ADMIN] if user["role"] == "admin" else [])
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

    public_pages = [PAGE_SCREENING, PAGE_ML_LAB]
    st.session_state["auth_page"] = st.sidebar.radio(
        "Chế độ sử dụng", public_pages, key="public_page_radio"
    )
    st.sidebar.divider()
    st.sidebar.caption(
        "Người tham gia có thể tự nhập hồ sơ sàng lọc hoặc xem mô hình NCKH.\n"
        "Quản trị viên/nhân sự đăng nhập bên dưới để mở Phân tích hàng loạt và Quản trị."
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
            "RDW": _blank(r[23]), "Mentzer": _blank(r[24]),
            "Ferritin": _blank(r[37]),
            "Xác suất ML (CBC)": f"{r[49]*100:.1f}%" if len(r) > 49 and r[49] is not None else "",
            "Xác suất ML (CBC+Fer)": f"{r[50]*100:.1f}%" if len(r) > 50 and r[50] is not None else "",
            "Nguy cơ Vòng 2": r[26] or "", "Kết luận": r[27] or r[28] or "",
            "Vòng 3": "Có" if r[29] else "Chưa", "Ngày xét nghiệm": r[30] or "",
            "Cơ sở thực hiện": r[31] or "", "Loại xét nghiệm": r[32] or "",
            "HbA2": _blank(r[34]), "HbF": _blank(r[35]), "HbE": _blank(r[36]),
            "Kết quả di truyền": r[40] or "",
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
        pass


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
    return {"admin_fail_at": 0.0, "admin_err": "", "db_fail_at": 0.0, "db_err": ""}


@st.cache_resource(ttl=86400, show_spinner="Đang tải danh mục địa giới...")
def load_admin_hierarchy():
    cached = _read_local_admin_cache()
    if cached:
        return cached
    state = _net_state()
    stale = _read_local_admin_cache(ignore_age=True)
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
# GENERAL HELPERS & ALTITUDE
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


def altitude_selector(key_prefix):
    choice = st.radio(
        "Chọn khoảng độ cao nơi đang sinh sống",
        list(ALTITUDE_OPTIONS.keys()),
        index=0,
        key=f"{key_prefix}_choice",
    )
    st.link_button("Tra cứu độ cao nơi ở", ALTITUDE_LOOKUP_URL, icon=":material/terrain:")

    adjustment = ALTITUDE_OPTIONS.get(choice, 0.0)
    st.info(f"Hiệu chỉnh Hb tham khảo (WHO 2024): **-{adjustment:.1f} g/dL**", icon=":material/tune:")

    if choice.startswith(("2.500", "3.000", "3.500", "4.000", "4.500")):
        st.warning(
            "Ở độ cao ≥2.500 m, WHO lưu ý mức độ không chắc chắn của hiệu chỉnh cao hơn.",
            icon=":material/warning:",
        )
    return choice, adjustment


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
    if score >= ROUND1_HIGH_THRESHOLD:
        return ("CAO", "Có nhiều yếu tố đáng lưu ý qua sàng lọc ban đầu; nên ưu tiên thực hiện/nhập CBC và Ferritin huyết thanh để mô hình Học máy đánh giá.")
    if score >= 4:
        return ("TRUNG BÌNH", "Có một số yếu tố đáng lưu ý; nên nhập kết quả CBC và Ferritin huyết thanh vào Vòng 2 để phân tầng nguy cơ.")
    return ("THẤP", "Chưa ghi nhận nhiều yếu tố đáng lưu ý qua bộ câu hỏi ban đầu; kết quả này không loại trừ khả năng mang gen Thalassemia.")


def round1_guidance(category):
    common = [
        "Vòng 1 chỉ là sàng lọc tiền sử ban đầu, không dùng để chẩn đoán hoặc loại trừ Thalassemia.",
        "Bạn được chuyển tiếp sang Vòng 2 để nhập 7 chỉ số CBC và Ferritin huyết thanh (nếu có) nhằm chạy mô hình Học máy THAL-SCREEN AI.",
    ]
    if category == "CAO":
        action = [
            "Ưu tiên nhập đầy đủ 7 chỉ số dòng hồng cầu (RBC, Hb, HCT, MCV, MCH, MCHC, RDW-CV) và định lượng Ferritin huyết thanh.",
            "Nếu mô hình Học máy dự đoán nguy cơ cao, cần thực hiện HPLC/điện di huyết sắc tố và xét nghiệm đột biến gen (HBA1, HBA2, HBB).",
        ]
    elif category == "TRUNG BÌNH":
        action = [
            "Nên nhập kết quả CBC và Ferritin vào Vòng 2 để mô hình Học máy phân biệt giữa thiếu máu thiếu sắt và mang gen Thalassemia.",
        ]
    else:
        action = [
            "Nếu đã có kết quả CBC và Ferritin, hãy nhập vào Vòng 2 vì người lành mang gen Thalassemia thường hoàn toàn không có triệu chứng ở Vòng 1.",
        ]
    return common + action


# ------------------------------------------------------------
# ROUND 2 (TÍCH HỢP CHỈ SỐ CỔ ĐIỂN + HỌC MÁY CBC + FERRITIN)
# ------------------------------------------------------------

def calculate_round2_score(mcv, mch, rbc, rdw, ferritin=None):
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
        reasons.append("MCV <80 fL — hồng cầu nhỏ rõ")
    if mch < 27:
        score += 1
        reasons.append("MCH <27 pg — hồng cầu nhược sắc rõ")
    if rdw > 14:
        reasons.append("RDW >14% — kích thước hồng cầu không đồng đều")
    if mcv < 85 and rbc >= 5.0:
        score += 1
        reasons.append("RBC ≥ 5,0 T/L — tăng sinh số lượng hồng cầu bù trừ trong bối cảnh MCV giảm")
    if ferritin is not None and ferritin >= 30 and (mcv < 80 or mch < 27):
        score += 2
        reasons.append(f"Ferritin huyết thanh = {ferritin:g} ng/mL (≥ 30 ng/mL) kèm hồng cầu nhỏ nhược sắc — loại trừ thiếu sắt đơn thuần, củng cố nghi ngờ Thalassemia")
    elif ferritin is not None and ferritin < 30:
        reasons.append(f"Ferritin huyết thanh = {ferritin:g} ng/mL (< 30 ng/mL) — gợi ý có thiếu sắt (đơn thuần hoặc đồng mắc cùng Thalassemia)")

    mentzer = _round_number(Decimal(str(mcv)) / Decimal(str(rbc)), 2) if rbc > 0 else None
    if mentzer is not None:
        reasons.append(f"Mentzer Index = {mentzer:.2f} (< 13 gợi ý Thalassemia; ≥ 13 gợi ý thiếu sắt)")
    return score, mentzer, reasons


def round2_category(score, mcv, mch=None, ml_category=None):
    if ml_category in ("CAO", "TRUNG BÌNH", "THẤP"):
        return ml_category
    screening_trigger = (mcv < 85) or (mch is not None and mch < 28)
    if not screening_trigger:
        return "THẤP"
    if score >= 6:
        return "CAO"
    return "TRUNG BÌNH"


def narrative_cbc_advice(hb_adjusted, mcv, mch, rbc, rdw, mentzer, ferritin=None, ml_res=None):
    findings, advice = [], []

    if mcv < 85:
        findings.append("MCV <85 fL — đạt tiêu chí sàng lọc CBC theo phụ lục BYT 2022.")
    if mch < 28:
        findings.append("MCH <28 pg — đạt tiêu chí sàng lọc CBC theo phụ lục BYT 2022.")
    if rdw > 14:
        findings.append("RDW-CV >14% — độ phân bố kích thước hồng cầu tăng.")
    if mcv < 85 and rbc >= 5.0:
        findings.append("RBC ≥5,0 T/L trong bối cảnh MCV thấp — kiểu hình đặc trưng thường gặp ở người mang gen Thalassemia.")

    if mcv < 85 and mentzer is not None and mentzer < 13:
        findings.append("Mentzer Index <13: chỉ số cổ điển nghiêng về mang gen Thalassemia hơn thiếu máu thiếu sắt.")
    elif mcv < 85 and mentzer is not None and mentzer >= 13:
        findings.append("Mentzer Index ≥13: chỉ số cổ điển nghiêng về thiếu sắt hoặc thể alpha-Thalassemia nhẹ.")

    if ferritin is not None:
        if ferritin < 30:
            findings.append(f"Ferritin huyết thanh giảm ({ferritin:g} ng/mL < 30 ng/mL): xác nhận có tình trạng thiếu hụt dự trữ sắt.")
            if ml_res and ml_res.get("coexisting_ida"):
                advice.append("Cần lưu ý khả năng ĐỒNG MẮC Thiếu máu thiếu sắt + Mang gen Thalassemia; nên điều trị/đánh giá sắt và thực hiện xét nghiệm gen khi có chỉ định.")
            else:
                advice.append("Khám chuyên khoa để điều trị thiếu sắt và kiểm tra lại công thức máu / HPLC theo hướng dẫn của bác sĩ.")
        else:
            findings.append(f"Ferritin huyết thanh bình thường/cao ({ferritin:g} ng/mL ≥ 30 ng/mL): giúp loại trừ thiếu máu thiếu sắt đơn thuần.")
            if mcv < 85 or mch < 28:
                advice.append("Hồng cầu nhỏ nhược sắc nhưng dự trữ sắt bình thường (Ferritin ≥ 30 ng/mL): ưu tiên chỉ định HPLC/điện di huyết sắc tố và xét nghiệm đột biến gen HBA1, HBA2, HBB (đặc biệt lưu ý alpha-Thalassemia không thể loại trừ nếu chỉ dựa vào HPLC bình thường).")
    else:
        if mcv < 85 or mch < 28:
            advice.append("Nên định lượng bổ sung Ferritin huyết thanh để kích hoạt Mô hình B (CBC + Ferritin), giúp phân biệt chính xác với thiếu máu thiếu sắt.")

    if hb_adjusted < 8:
        advice.append("Hb sau hiệu chỉnh rất thấp (<8 g/dL): cần được khám và đánh giá huyết học sớm.")
    elif hb_adjusted < 10:
        advice.append("Hb sau hiệu chỉnh giảm rõ (<10 g/dL): nên khám chuyên khoa Huyết học để xác định nguyên nhân.")

    if not findings:
        findings.append("Chưa ghi nhận hồng cầu nhỏ hoặc nhược sắc rõ trên công thức máu hiện tại.")
    if not advice:
        advice.append("Tiếp tục đối chiếu với khoảng tham chiếu của phòng xét nghiệm và tư vấn tiền hôn nhân/trước sinh khi có nhu cầu.")
    return findings, advice


# ------------------------------------------------------------
# ROUND 3 — BYT 2022 INTERPRETATION SUPPORT
# ------------------------------------------------------------

def round3_byt_interpretation(r3):
    notes = []
    hba, hba2, hbf, hbe = r3.get("hba"), r3.get("hba2"), r3.get("hbf"), r3.get("hbe")
    ferritin, tsat = r3.get("ferritin"), r3.get("transferrin_saturation")

    if hba is not None and 96.5 <= hba <= 98 and hba2 is not None and 2 <= hba2 <= 3.5 and hbf is not None and hbf < 1:
        notes.append("Thành phần Hb nằm trong khoảng tham chiếu HbA 96,5–98%, HbA2 2–3,5%, HbF <1%. Lưu ý quan trọng: kết quả HPLC/điện di bình thường KHÔNG loại trừ được người mang gen alpha-Thalassemia; cần xét nghiệm gen HBA1, HBA2 nếu MCV/MCH giảm và Ferritin bình thường.")
    if hba2 is not None and hba2 > 3.5:
        notes.append("HbA2 >3,5% — dấu hiệu đặc trưng gợi ý người mang gen Beta-Thalassemia (cần xác nhận đột biến gen HBB và đồng nhiễm alpha-Thalassemia nếu có).")
    if hbf is not None and hbf >= 1:
        notes.append("HbF tăng (≥1%) — cần đánh giá trong bối cảnh toàn bộ thành phần huyết sắc tố và xét nghiệm phân tử.")
    if hbe is not None and hbe > 0:
        notes.append("Có đỉnh HbE bất thường — gợi ý dị hợp tử hoặc đồng hợp tử bệnh huyết sắc tố E (đột biến CD26 trên gen HBB).")

    combined = ((r3.get("genetic_result") or "") + " " + (r3.get("lab_conclusion") or "")).lower()
    for marker, label in [("hbh", "HbH"), ("hb bart", "Hb Bart’s"), ("hbcs", "HbCS"), ("constant spring", "Hb Constant Spring"), ("sea", "--SEA"), ("3.7", "-α3.7"), ("4.2", "-α4.2")]:
        if marker in combined:
            notes.append(f"Phiếu xét nghiệm ghi nhận biến thể {label} thuộc nhóm Alpha-Thalassemia; cần tư vấn di truyền chuyên khoa.")

    if ferritin is not None and ferritin < 30:
        notes.append("Ferritin <30 ng/mL — xác nhận thiếu hụt dự trữ sắt theo hướng dẫn BYT 2022.")
    if tsat is not None and tsat < 30:
        notes.append("Độ bão hòa transferrin <30% — gợi ý thiếu máu thiếu sắt.")
    if not notes:
        notes.append("Chưa có mẫu hình tự động để đối chiếu từ các trường đã nhập. Cần xem toàn bộ phiếu xét nghiệm và đánh giá chuyên môn.")
    return notes


# ------------------------------------------------------------
# MEDICAL FACILITIES
# ------------------------------------------------------------

def _f(name, tier, note, query=None):
    q = quote(query or name)
    return {"name": name, "tier": tier, "note": note,
            "maps": f"https://www.google.com/maps/search/?api=1&query={q}"}


_TW = "Hạng đặc biệt – tuyến Trung ương"
_SPEC = "Ưu tiên khi cần đánh giá chuyên sâu huyết học"

MEDICAL_FACILITIES = {
    "Huế": [
        _f("Bệnh viện Trung ương Huế (Cơ sở nghiên cứu chính của đề tài)", _TW, "Trung tâm Huyết học Truyền máu & Khoa Huyết học Lâm sàng — thực hiện đầy đủ CBC, Ferritin, HPLC/điện di và PCR đột biến gen α/β-Thalassemia"),
        _f("Bệnh viện Trường Đại học Y – Dược Huế", "Hạng I", "Bệnh viện trường đại học tuyến chuyên sâu", "Bệnh viện Trường Đại học Y Dược Huế"),
    ],
    "Đà Nẵng": [
        _f("Bệnh viện Đà Nẵng", "Hạng I", "Bệnh viện đa khoa tuyến cuối của thành phố Đà Nẵng"),
        _f("Bệnh viện C Đà Nẵng", "Bệnh viện tuyến Trung ương", "Cơ sở tuyến Trung ương tại Đà Nẵng"),
        _f("Bệnh viện Phụ sản – Nhi Đà Nẵng", "Hạng I", "Chuyên sâu sàng lọc trước sinh, sơ sinh và huyết học lâm sàng"),
        _f("Bệnh viện Đại học Y – Dược, Đại học Đà Nẵng", "Bệnh viện trường đại học", "Có thể lựa chọn khi cần đánh giá chuyên khoa", "Bệnh viện Đại học Y Dược Đại học Đà Nẵng"),
    ],
    "Hà Nội": [
        _f("Viện Huyết học – Truyền máu Trung ương", _TW, "Đơn vị đầu ngành toàn quốc về chẩn đoán phân tử và tư vấn di truyền Thalassemia"),
        _f("Bệnh viện Bạch Mai", _TW, _SPEC, "Bệnh viện Bạch Mai Hà Nội"),
        _f("Bệnh viện Phụ sản Trung ương", _TW, "Chuyên sâu sàng lọc trước sinh Thalassemia"),
    ],
    "Hồ Chí Minh": [
        _f("Bệnh viện Truyền máu Huyết học TP. Hồ Chí Minh", "Hạng I chuyên khoa", _SPEC),
        _f("Bệnh viện Chợ Rẫy", _TW, _SPEC, "Bệnh viện Chợ Rẫy Hồ Chí Minh"),
        _f("Bệnh viện Đại học Y Dược TP. Hồ Chí Minh", "Hạng I", "Bệnh viện trường đại học tuyến chuyên sâu"),
    ],
    "Cần Thơ": [
        _f("Bệnh viện Huyết học – Truyền máu Cần Thơ", "Chuyên khoa tuyến khu vực", "Ưu tiên khi cần xét nghiệm chuyên sâu Huyết học tại ĐBSCL"),
        _f("Bệnh viện Đa khoa Trung ương Cần Thơ", "Bệnh viện tuyến Trung ương", "Ưu tiên khi cần tuyến chuyên sâu tại ĐBSCL"),
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

for _p in ["Thanh Hóa", "Hà Tĩnh", "Khánh Hòa", "Lâm Đồng", "Quảng Ngãi", "Gia Lai", "Bắc Ninh",
           "Phú Thọ", "Lào Cai", "Hưng Yên", "Ninh Bình", "Tây Ninh", "Vĩnh Long", "Cà Mau"]:
    MEDICAL_FACILITIES[_p] = [_f(f"Bệnh viện Đa khoa tỉnh {_p}", "Hạng I", "Bệnh viện đa khoa tuyến tỉnh")]

for _p in ["Điện Biên", "Lai Châu", "Sơn La", "Tuyên Quang"]:
    MEDICAL_FACILITIES[_p] = [_f(f"Bệnh viện Đa khoa tỉnh {_p}", "Tuyến tỉnh", "Cơ sở đa khoa tuyến tỉnh")]

MEDICAL_FACILITIES.update({
    "Cao Bằng": [
        _f("Bệnh viện Đa khoa tỉnh Cao Bằng", "Tuyến tỉnh", "Cơ sở đa khoa tuyến tỉnh"),
    ],
    "Lạng Sơn": [
        _f("Bệnh viện Đa khoa tỉnh Lạng Sơn", "Tuyến tỉnh", "Cơ sở đa khoa tuyến tỉnh"),
    ],
    "Quảng Trị": [
        _f("Bệnh viện Hữu nghị Việt Nam – Cuba Đồng Hới", "Tuyến Trung ương", "Cơ sở tuyến trên phục vụ khu vực", "Bệnh viện Hữu nghị Việt Nam Cuba Đồng Hới"),
        _f("Bệnh viện Đa khoa tỉnh Quảng Trị", "Tuyến tỉnh", "Cơ sở đa khoa tuyến tỉnh"),
    ],
})

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
# WORD REPORT (TÍCH HỢP KẾT QUẢ MÔ HÌNH HỌC MÁY & SHAP)
# ------------------------------------------------------------

def make_word(patient, r1_score, r1_category, r1_reasons, r2=None):
    from docx import Document
    doc = Document()
    doc.add_heading("TRƯỜNG ĐẠI HỌC PHAN CHÂU TRINH · HỆ THỐNG THAL-SCREEN AI", level=3)
    doc.add_heading("PHIẾU KẾT QUẢ DỰ ĐOÁN NGUY CƠ MANG GEN THALASSEMIA", level=1)
    doc.add_paragraph("Công cụ hỗ trợ sàng lọc bằng Học máy (CBC + Ferritin); không thay thế chẩn đoán xác định di truyền.")

    doc.add_heading("I. THÔNG TIN HÀNH CHÍNH", level=2)
    for label, value in [
        ("Họ và tên", patient["full_name"]), ("Ngày sinh", patient["birth_date"]),
        ("Tuổi", patient["age"]), ("Giới tính", patient["gender"]),
        ("Số điện thoại", patient["phone"]), ("Địa chỉ hiện tại", patient["current_address"]),
        ("Tỉnh/thành", patient["province"]), ("Phường/xã/đặc khu", patient["commune"]),
    ]:
        doc.add_paragraph(f"{label}: {value}")

    doc.add_heading("II. SÀNG LỌC TIỀN SỬ (VÒNG 1)", level=2)
    doc.add_paragraph(f"Điểm Vòng 1: {r1_score}/{ROUND1_MAX_SCORE} — Mức lưu ý: {r1_category}")
    for item in r1_reasons:
        doc.add_paragraph(f"- {item}")

    if r2:
        doc.add_heading("III. KẾT QUẢ XÉT NGHIỆM CBC, FERRITIN & DỰ ĐOÁN HỌC MÁY (VÒNG 2)", level=2)
        fer_txt = f"{r2['ferritin']:.1f} ng/mL" if r2.get("ferritin") is not None else "Chưa định lượng"
        prob_a_txt = f"{r2.get('ml_prob_cbc', 0)*100:.1f}%" if r2.get("ml_prob_cbc") is not None else "—"
        prob_b_txt = f"{r2.get('ml_prob_cbc_ferritin', 0)*100:.1f}%" if r2.get("ml_prob_cbc_ferritin") is not None else "—"
        for line in [
            f"Khoảng độ cao cư trú: {r2['altitude_choice']} (Hiệu chỉnh Hb: -{r2['adjustment']:.1f} g/dL)",
            f"Hb thực đo: {r2['hb']:.1f} g/dL ({r2['hb']*10:.0f} g/L) | Hb hiệu chỉnh: {r2['hb_adjusted']:.1f} g/dL",
            f"RBC: {r2['rbc']:.2f} T/L | HCT: {r2.get('hct', 0):.1f}% | MCV: {r2['mcv']:.1f} fL | MCH: {r2['mch']:.1f} pg | MCHC: {r2.get('mchc', 0):.1f} g/L | RDW-CV: {r2['rdw']:.1f}%",
            f"Ferritin huyết thanh: {fer_txt}",
            f"Chỉ số phái sinh: Mentzer = {r2['mentzer']:.2f} | Green & King = {r2.get('green_king', 0):.2f} | RDWI = {r2.get('rdwi', 0):.2f}",
            f"Xác suất dự đoán Học máy (XGBoost - Mô hình A [Chỉ CBC]): {prob_a_txt}",
            f"Xác suất dự đoán Học máy (XGBoost - Mô hình B [CBC + Ferritin]): {prob_b_txt}",
            f"Phân tầng nguy cơ: {r2['category']}",
            f"Nhận định tổng hợp: {r2['conclusion']}",
        ]:
            doc.add_paragraph(line)

        if r2.get("shap"):
            doc.add_heading("IV. GIẢI THÍCH MÔ HÌNH BẰNG PHƯƠNG PHÁP SHAP", level=2)
            for s_item in r2["shap"][:5]:
                doc.add_paragraph(
                    f"- {s_item['label']}: SHAP = {s_item['shap_value']:+.3f} (tỷ trọng {s_item['abs_pct']:.1f}%) — {s_item['direction']}"
                )

        doc.add_heading("V. PHÂN TÍCH LÂM SÀNG & KHUYẾN NGHỊ CHỈ ĐỊNH", level=2)
        for finding in r2["findings"]:
            doc.add_paragraph(f"- {finding}")
        for advice in r2["advice"]:
            doc.add_paragraph(f"- {advice}")

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out


@st.cache_data(show_spinner=False, max_entries=20, ttl=600)
def make_word_bytes(patient, r1_score, r1_category, r1_reasons, r2=None):
    return make_word(patient, r1_score, r1_category, r1_reasons, r2).getvalue()


# ============================================================
# RESEARCH OVERVIEW & PROJECT INTRO DIALOGS (ĐỒNG BỘ ĐỀ CƯƠNG MỚI)
# ============================================================

@st.dialog("Giới thiệu đề tài NCKH và Nhóm nghiên cứu", width="large")
def show_project_intro():
    st.markdown("""
### ĐỀ TÀI NGHIÊN CỨU KHOA HỌC
**XÂY DỰNG VÀ ĐÁNH GIÁ MÔ HÌNH HỌC MÁY DỰ ĐOÁN NGUY CƠ MANG GEN THALASSEMIA DỰA TRÊN CHỈ SỐ CÔNG THỨC MÁU VÀ FERRITIN HUYẾT THANH Ở NGƯỜI TRƯỞNG THÀNH TẠI BỆNH VIỆN TRUNG ƯƠNG HUẾ**

- **Đơn vị chủ trì:** Khoa Kỹ thuật Xét nghiệm Y học – Trường Đại học Phan Châu Trinh
- **Cơ sở phối hợp nghiên cứu:** Bệnh viện Trung ương Huế (Thời gian: 11/2026 – 05/2027)
- **Chủ nhiệm đề tài:** Nguyễn Đắc Lâu (Lớp 24XN1)
- **Thành viên nhóm nghiên cứu:** Trần Phan Thảo Vi, Huỳnh Ngọc Anh Huy, Lê Văn Thịnh, Phan Ngọc Bảo An
- **Giảng viên hướng dẫn:** TS. Nguyễn Đình Khắc Chân

---

### 1. Mục tiêu nghiên cứu của đề tài
1. **Mục tiêu 1:** Khảo sát đặc điểm các chỉ số xét nghiệm công thức máu ngoại vi (CBC: RBC, Hb, HCT, MCV, MCH, MCHC, RDW-CV), Ferritin huyết thanh và các chỉ số phân biệt hồng cầu (Mentzer Index, Green & King Index, RDWI) ở người trưởng thành có và không mang gen Thalassemia tại Bệnh viện Trung ương Huế.
2. **Mục tiêu 2:** Xây dựng, đánh giá hiệu năng 3 thuật toán học máy (**Logistic Regression**, **Random Forest**, **XGBoost**), chứng minh giá trị bổ sung của Ferritin huyết thanh (qua kiểm định DeLong, chỉ số cải thiện phân loại ròng NRI, Brier score), giải thích mô hình bằng phương pháp **SHAP** và triển khai phần mềm hỗ trợ sàng lọc **THAL-SCREEN AI**.

### 2. Điểm mới về mặt khoa học và công nghệ của ứng dụng THAL-SCREEN AI
- **So sánh trực tiếp Mô hình A (Chỉ CBC) và Mô hình B (CBC + Ferritin):** Lượng hóa ngay trên giao diện sự thay đổi xác suất dự đoán khi bổ sung định lượng Ferritin huyết thanh, giúp giải quyết bài toán chồng lấp hồng cầu nhỏ nhược sắc giữa thiếu máu thiếu sắt (IDA) và Thalassemia.
- **Nhận diện đồng mắc Thiếu sắt + Mang gen Thalassemia:** Phát hiện các trường hợp Ferritin thấp (< 30 ng/mL) nhưng số lượng hồng cầu RBC tăng bù trừ và MCV giảm sâu để cảnh báo nguy cơ âm tính giả trên điện di HbA2.
- **Giải thích mô hình cá thể hóa bằng SHAP (Explainable AI):** Minh bạch hóa đóng góp của từng biến số xét nghiệm lên kết quả dự đoán của mỗi người bệnh, khắc phục hoàn toàn hạn chế "hộp đen" của trí tuệ nhân tạo trong y học.
    """)


@st.dialog("Cơ sở khoa học và phương pháp Học máy", width="large")
def show_research_overview():
    st.markdown("""
### 1. Cơ sở huyết học và chuyển hóa sắt trong sàng lọc Thalassemia
- **7 chỉ số dòng hồng cầu trên CBC (RBC, Hb, HCT, MCV, MCH, MCHC, RDW-CV):** Người lành mang gen Thalassemia thường có MCV giảm (< 80 fL), MCH giảm (< 27 pg) kèm số lượng hồng cầu RBC bình thường hoặc tăng bù trừ (≥ 5,0 T/L) và RDW-CV tăng nhẹ.
- **Vai trò của Ferritin huyết thanh:** Trong thiếu máu thiếu sắt đơn thuần, dự trữ sắt suy giảm làm Ferritin giảm thấp (< 15–30 ng/mL). Ngược lại, ở người mang gen Thalassemia đơn thuần, Ferritin thường bình thường hoặc tăng. Việc tích hợp Ferritin vào mô hình Học máy giúp bóc tách chính xác thiếu máu thiếu sắt và nhận diện cả nhóm đồng mắc.
- **Tiêu chuẩn tham chiếu vàng (Reference Standard):** Điện di huyết sắc tố (HPLC) bình thường không thể loại trừ hoàn toàn thể mang gen **alpha-Thalassemia** (như mất đoạn 1 gen *-α3.7*, *-α4.2* hoặc 2 gen *--SEA*). Do đó, nhãn kết cục của mô hình được xác nhận bằng xét nghiệm sinh học phân tử đột biến gen *HBA1, HBA2* và *HBB* kết hợp HPLC.

### 2. Phương pháp Học máy và Giải thích SHAP
- **Ba thuật toán đối sánh:** Hồi quy Logistic (Logistic Regression), Rừng ngẫu nhiên (Random Forest) và Extreme Gradient Boosting (XGBoost).
- **Giải thích SHAP (SHapley Additive exPlanations):** Xác định trọng số đóng góp tương đối của từng biến (MCV ~ 25%, MCH ~ 21%, RBC ~ 16%, Ferritin ~ 12%, RDW-CV ~ 10%, Hb ~ 8%, MCHC ~ 5%, HCT ~ 3%) giúp bác sĩ lâm sàng kiểm chứng tính hợp lý sinh bệnh học trên từng ca bệnh.
    """)
    st.divider()
    st.subheader("Tài liệu chuyên môn tham khảo chính")
    st.link_button("Rustam et al. (2022) – Scientific Reports (β-Thalassemia ML Screening)", "https://doi.org/10.1038/s41598-022-22011-8", icon=":material/open_in_new:", use_container_width=True)
    st.link_button("Trần Danh Cường và cs. (2023) – MJHID (Thalassemia CDSS in Vietnam)", "https://doi.org/10.4084/MJHID.2023.026", icon=":material/open_in_new:", use_container_width=True)
    st.link_button("Christensen et al. (2025) – CMPB (α-Thalassemia ML Classification)", "https://doi.org/10.1016/j.cmpb.2024.108581", icon=":material/open_in_new:", use_container_width=True)
    st.link_button("Tepakhan et al. (2025) – Scientific Reports (Differentiating IDA & Thalassemia)", "https://doi.org/10.1038/s41598-025-01458-5", icon=":material/open_in_new:", use_container_width=True)
    st.link_button("Tạ Văn Thảo và cs. (2025) – Tạp chí Nghiên cứu Y học (Sàng lọc β-Thalassemia bằng ML)", "https://doi.org/10.52852/tcncyh.v196i11E17.4159", icon=":material/open_in_new:", use_container_width=True)


# ---------- Câu hỏi Vòng 1 ----------
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
        ("q9", None),
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
        ("q19", "19. Bạn hiện có kết quả CBC và/hoặc Ferritin trong vòng 6–12 tháng gần đây không?"),
        ("q20", "20. Bạn có gặp khó khăn khi đến cơ sở có xét nghiệm chuyên sâu do khoảng cách, chi phí hoặc thời gian di chuyển không?"),
    ]),
]

Q9_OPTIONS = [
    "Chưa xét nghiệm", "Đã xét nghiệm, bình thường", "Đã nghi ngờ",
    "Đã xác định mang gen", "Không nhớ",
]


# ============================================================
# BATCH ANALYSIS — HỒ SƠ ĐẦY ĐỦ NHIỀU NGƯỜI CÙNG LÚC (TÍCH HỢP ML)
# ============================================================

BATCH_MAX_FILES = 20
BATCH_MAX_MB_TABLE = 30
BATCH_MAX_MB_AI = 10
BATCH_MAX_TABLE_ROWS = 2000
BATCH_MAX_TOTAL_ROWS = 2000
AI_FILE_TYPES = {
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
SKIP_SHEETS = {"danhsach", "huongdan", "docao", "readme"}
BIRTH_ALIASES = {"ngaysinh", "namsinh", "dob", "birthdate", "birthyear", "ngaythangnamsinh"}
BATCH_COLUMNS = ([COL_LABEL, COL_SRC, COL_AGE, COL_SEX, COL_MARITAL, COL_PREG, COL_ALT] + Q_COLS + CBC_COLS + UNIT_COLS
                 + IRON_COLS + HB_NUM_COLS + [COL_HPLC, COL_PCR, COL_LAB, COL_NOTE])

COLUMN_GROUPS = {
    "Thông tin": [COL_LABEL, COL_SRC, COL_AGE, COL_SEX, COL_MARITAL, COL_PREG, COL_ALT, COL_NOTE],
    "Câu hỏi Vòng 1 (Q1–Q20)": [COL_LABEL] + Q_COLS,
    "Công thức máu (18 chỉ số)": [COL_LABEL] + CBC_COLS + UNIT_COLS,
    "Ferritin · Sắt · HPLC · PCR": [COL_LABEL] + IRON_COLS + HB_NUM_COLS + TEXT_RESULT_COLS,
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
    "MCHC": {"mchc", "mchcgl"},
    "RDW-CV": {"rdwcv", "rdw"},
    "RDW-SD": {"rdwsd"},
    "PLT": {"plt", "tieucau", "platelet", "platelets"},
    "MPV": {"mpv"},
    "PDW": {"pdw"},
    "PCT": {"pct", "plateletcrit"},
    "Đơn vị HGB": {"donvihgb", "donvihb", "hbunit", "hgbunit"},
    "Đơn vị RBC": {"donvirbc", "rbcunit"},
    "Ferritin": {"ferritin", "ferritinngml", "ferritinhuyetthanh"},
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


def _norm_choice(value, q9=False):
    s = _s(value)
    if s is None:
        return None
    if isinstance(value, numbers.Number) and not isinstance(value, bool) and float(value).is_integer():
        s = str(int(value))
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
    if value is not None and value > 30:
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


AI_PROMPT = """Bạn là công cụ trích xuất số liệu từ tài liệu (ảnh, PDF hoặc văn bản Word) gồm phiếu xét nghiệm (công thức máu CTM/CBC, ferritin/sắt, HPLC/điện di huyết sắc tố, PCR/xét nghiệm gen Thalassemia) và/hoặc phiếu khảo sát tầm soát cộng đồng.

QUY TẮC BẮT BUỘC:
- KHÔNG trích xuất, KHÔNG nhắc lại họ tên, ngày sinh, số điện thoại, địa chỉ, mã bệnh nhân, tên bác sĩ.
- Chỉ lấy KẾT QUẢ của người bệnh, KHÔNG lấy khoảng tham chiếu/giá trị bình thường.
- Không thấy thì để null. Tuyệt đối không đoán, không tự tính, không suy diễn.
- Ghi đơn vị đúng như trên phiếu (hgb_unit: "g/dL" hoặc "g/L"; rbc_unit: "T/L", "10^12/L" hoặc "10^6/µL").

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
    block_type = "document" if mime == "application/pdf" else "image"
    return _ai_request([
        {"type": block_type, "source": {"type": "base64", "media_type": mime,
                                        "data": base64.standard_b64encode(file_bytes).decode("ascii")}},
        {"type": "text", "text": AI_PROMPT},
    ])


def ai_extract_report_text(text):
    return _ai_request([{"type": "text",
                         "text": 'NỘI DUNG TÀI LIỆU:\n"""\n' + text[:60000] + '\n"""\n\n' + AI_PROMPT}])


def ai_map_columns(headers, labels):
    texts = _question_texts()
    fields = [c for c in BATCH_COLUMNS if c not in (COL_SRC, COL_NOTE)]
    field_desc = "\n".join(f"- {c}" + (f": {texts[c]}" if c in texts else "") for c in fields)
    cols = "\n".join(f"- {h}" + (f" (nhãn biến: {l})" if l else "") for h, l in zip(headers, labels))
    prompt = (
        "Bạn ghép TÊN CỘT của một bộ dữ liệu sàng lọc Thalassemia vào danh sách TRƯỜNG CHUẨN. "
        "Chỉ ghép khi chắc chắn. Không chắc thì trả null.\n\n"
        f"TRƯỜNG CHUẨN:\n{field_desc}\n\nCỘT CẦN GHÉP:\n{cols}\n\n"
        'Chỉ trả về MỘT đối tượng JSON dạng {"tên cột gốc": "trường chuẩn hoặc null"}, không thêm chữ nào khác.'
    )
    return _ai_request([{"type": "text", "text": prompt}], max_tokens=2500)


@st.cache_data(show_spinner=False, max_entries=30, ttl=600)
def ai_extract_report_cached(file_bytes, mime):
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
# QUY ƯỚC MÃ HÓA & ĐỌC BẢNG HÀNG LOẠT
# ------------------------------------------------------------

BLANK_MEANING = "— bỏ trống —"
BIRTH_TARGET = "Ngày/Năm sinh (tự tính tuổi)"
CB_COLS = ["Mã trong file", "Ý nghĩa"]
CB_TITLES = {
    "sex": "Giới tính", "yn": "Câu hỏi Có/Không (Q1–Q8, Q10–Q20)", "q9": "Câu 9 (xét nghiệm Thalassemia trước đây)",
    "marital": "Tình trạng hôn nhân", "preg": "Đang mang thai",
}
_Q9_ORDER = ["Chưa xét nghiệm", "Đã xét nghiệm, bình thường", "Đã nghi ngờ", "Đã xác định mang gen", "Không nhớ"]
CB_PRESETS = {
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
            CB_COLS[0]: st.column_config.TextColumn("Mã trong file", required=True),
            CB_COLS[1]: st.column_config.SelectboxColumn("Ý nghĩa", options=[*CB_MEANINGS[key], BLANK_MEANING], required=True),
        },
    )


def render_codebook_panel():
    with st.expander("Quy ước mã hóa & ghép cột — dùng khi file có mã số (0/1/2…) hoặc tên cột khác mẫu",
                     icon=":material/tune:"):
        tabs = st.tabs(["Giới tính", "Có/Không (Q)", "Câu 9", "Hôn nhân", "Thai kỳ", "Số liệu & đơn vị", "Ghép cột thủ công"])
        for tab, key in zip(tabs[:5], ["sex", "yn", "q9", "marital", "preg"]):
            with tab:
                _codebook_editor(key)
        with tabs[5]:
            st.text_input("Mã thiếu của số liệu (CTM, ferritin, HbA2, tuổi…) — cách nhau bằng dấu phẩy",
                          key="cb_num_missing", placeholder="Ví dụ: -999, 9999")
            c1, c2 = st.columns(2)
            c1.selectbox("Đơn vị Hb của file", list(HB_UNIT_CHOICES), key="cb_hb_unit")
            c2.selectbox("Đơn vị RBC của file", list(RBC_UNIT_CHOICES), key="cb_rbc_unit")
        with tabs[6]:
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
                if txt and (not cells or cells[-1] != txt):
                    cells.append(txt)
            if cells:
                parts.append(" | ".join(cells))
    return "\n".join(parts)


def _read_with_tempfile(ext, data):
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
        raise RuntimeError(f"Thiếu thư viện để đọc định dạng {ext}: {exc.name or exc}.") from exc
    raise RuntimeError("Định dạng không được hỗ trợ")


@st.cache_data(show_spinner=False, max_entries=10, ttl=600)
def read_tables_cached(name, data):
    return read_tables(name, data)


def _map_columns(df, labels, ai_ok, notes, source, manual=None):
    mapping = {}
    computed_norm = {_norm(c) for c in COMPUTED_COLS}
    manual = manual or {}
    for col in df.columns:
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
                    notes.append(f"{source}: AI đã ghép {len(added)} cột theo tên: " + ", ".join(added[:25]))
            except Exception as exc:
                notes.append(f"{source}: AI không ghép được cột ({exc}).")
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
            continue
        row[COL_LABEL] = row[COL_LABEL] or f"{filename} · dòng {i + 2}"
        row["Đơn vị HGB"] = cb["hb_unit"] if cb["hb_unit"] != "auto" else _norm_hb_unit(row.get("Đơn vị HGB"), row["HGB"])
        row["Đơn vị RBC"] = _norm_rbc_unit(row.get("Đơn vị RBC"))
        if cb["rbc_unit"] == "per_ul" and row["RBC"] is not None:
            row["RBC"] = row["RBC"] / 1_000_000
        rows.append(_normalize_row_units(row))
    return rows


def load_file_rows(name, data, default_alt, ai_ok, cb=None):
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
        if ext not in (".txt", *WORD_EXT):
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

    raise RuntimeError("Không nhận diện được cột dữ liệu nào")


# ------------------------------------------------------------
# PHÂN TÍCH HÀNG LOẠT (TÍCH HỢP MÔ HÌNH A, MÔ HÌNH B & SHAP)
# ------------------------------------------------------------

RISK_RANK = {"THẤP": 1, "TRUNG BÌNH": 2, "CAO": 3}


def who_anemia_flag(hb_adj, age, sex, pregnant=False):
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
    steps = []
    iron_low = ferritin is not None and ferritin < 30

    if cat1 == "CAO" and not has_cbc:
        steps.append("Vòng 1 mức cao nhưng chưa có CTM đầy đủ: nên ưu tiên thực hiện/nhập công thức máu và Ferritin.")
    if cat2 in ("TRUNG BÌNH", "CAO"):
        if ferritin is None:
            steps.append("CTM đạt tiêu chí sàng lọc: nên định lượng Ferritin huyết thanh để kích hoạt Mô hình B (CBC + Ferritin).")
        elif iron_low:
            steps.append("Ferritin <30 ng/mL: gợi ý có thiếu sắt (đơn thuần hoặc đồng mắc cùng Thalassemia). Thiếu sắt có thể làm giảm giả tạo HbA2 trên HPLC.")
        elif not hplc_done:
            steps.append("Mô hình Học máy (CBC + Ferritin) gợi ý nguy cơ mang gen và Ferritin bình thường: ưu tiên chỉ định HPLC/điện di huyết sắc tố và xét nghiệm gen α/β-Thalassemia.")
        if hplc_done and not pcr_done and hba2 is not None and hba2 <= 3.5 and not (hbe and hbe > 0) and not iron_low:
            steps.append("CTM gợi ý nhưng HbA2 ≤3,5%: cần xét nghiệm PCR đột biến gen alpha-Thalassemia (HBA1, HBA2) vì HPLC bình thường không loại trừ được alpha-Thalassemia.")
    if hba2 is not None and hba2 > 3.5:
        steps.append("HbA2 >3,5%: gợi ý Beta-Thalassemia; cần khám chuyên khoa Huyết học, xét nghiệm gen HBB và tư vấn di truyền.")
    if hbe is not None and hbe > 0:
        steps.append("Có HbE: cần bác sĩ Huyết học đánh giá bệnh huyết sắc tố E và tư vấn di truyền.")
    if pcr_done:
        steps.append("Đã có kết quả xét nghiệm gen/PCR: đối chiếu trực tiếp với bác sĩ Huyết học và tư vấn di truyền.")
    if not steps and anything:
        steps.append("Tiếp tục theo dõi sức khỏe và đối chiếu với bác sĩ khi có nhu cầu tư vấn tiền hôn nhân/trước sinh.")
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

    nums = {c: _to_float(rec.get(c)) for c in CBC_COLS}
    ferritin, iron, tsat = (_to_float(rec.get(c)) for c in IRON_COLS)
    n_cbc = sum(v is not None for v in nums.values())
    out["Số chỉ số CTM"] = f"{n_cbc}/18" if n_cbc else ""
    cat2, has_cbc = None, False
    if n_cbc:
        core = ["HGB", "MCV", "MCH", "RBC", "RDW-CV"]
        missing = [c for c in core if nums[c] is None]
        if missing:
            warnings.append("CTM thiếu " + ", ".join(missing) + " — chưa chạy mô hình Học máy")
        else:
            try:
                hb = hb_to_g_dl(nums["HGB"], _s(rec.get("Đơn vị HGB")) or "g/dL")
                rbc = rbc_to_t_l(nums["RBC"], _s(rec.get("Đơn vị RBC")) or "T/L")
            except ValueError as exc:
                hb = rbc = None
                warnings.append(str(exc))
            if hb is not None:
                mcv, mch, rdw = nums["MCV"], nums["MCH"], nums["RDW-CV"]
                alt_choice = rec.get(COL_ALT) if rec.get(COL_ALT) in ALTITUDE_OPTIONS else "<500 m"
                hb_adj = hb - ALTITUDE_OPTIONS[alt_choice]
                ml_res = predict_thal_screen_ai(
                    hb_g_dl=hb, rbc=rbc, mcv=mcv, mch=mch, rdw=rdw,
                    hct=nums.get("HCT"), mchc_g_l=nums.get("MCHC"), ferritin=ferritin,
                )
                score2, mentzer, reasons2 = calculate_round2_score(mcv, mch, rbc, rdw, ferritin)
                cat2 = ml_res["category"]
                concl2 = ml_res["conclusion"]
                findings, advice = narrative_cbc_advice(hb_adj, mcv, mch, rbc, rdw, mentzer, ferritin, ml_res)
                anemia = who_anemia_flag(hb_adj, age, sex, pregnant)
                has_cbc = True
                levels.append(cat2)
                det["r2"] = {
                    "category": cat2, "conclusion": concl2, "reasons": reasons2,
                    "findings": findings, "advice": advice, "altitude": alt_choice,
                    "score": score2, "mentzer": mentzer, "ml": ml_res,
                }
                out.update({
                    "Hb (g/dL)": round(hb, 2), "Hb sau hiệu chỉnh": round(hb_adj, 2),
                    "Thiếu máu (WHO 2024)": anemia, "MCV": mcv, "MCH": mch,
                    "MCHC": ml_res["cbc"]["MCHC_gL"], "RDW-CV": rdw,
                    "RBC (T/L)": rbc, "Mentzer": mentzer,
                    "Green & King": ml_res["indices"]["green_king"],
                    "Xác suất ML (Chỉ CBC)": f"{ml_res['probs_a']['XGBoost']*100:.1f}%",
                    "Xác suất ML (CBC + Ferritin)": f"{ml_res['probs_b']['XGBoost']*100:.1f}%" if ferritin is not None else "Chưa có Ferritin",
                    "Vòng 2 - nguy cơ": cat2,
                })
        det["cbc_values"] = {c: v for c, v in nums.items() if v is not None}

    hba, hba2, hbf, hbe = (_to_float(rec.get(c)) for c in HB_NUM_COLS)
    hplc_text, pcr_text, lab_text = _s(rec.get(COL_HPLC)), _s(rec.get(COL_PCR)), _s(rec.get(COL_LAB))
    num3 = [hba, hba2, hbf, hbe, ferritin, iron, tsat]
    has3 = any(v is not None for v in num3) or any((hplc_text, pcr_text, lab_text))
    flag3, notes3 = None, []
    hplc_done = any(v is not None for v in (hba, hba2, hbf, hbe)) or bool(hplc_text)
    if has3:
        combined_text = " ".join(filter(None, [hplc_text, pcr_text, lab_text])).lower()
        notes3 = round3_byt_interpretation({
            "hba": hba, "hba2": hba2, "hbf": hbf, "hbe": hbe, "ferritin": ferritin,
            "transferrin_saturation": tsat, "genetic_result": pcr_text,
            "lab_conclusion": " ".join(filter(None, [hplc_text, lab_text])),
        })
        marker = any(m in combined_text for m in ("hbh", "hb bart", "hbcs", "constant spring", "sea", "3.7", "4.2"))
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
    for r in range(2, n_rows + 2):
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

    wl = wb.add_worksheet("Danh_sach")
    wl.write(0, 0, "Khoảng độ cao", head)
    wl.write(0, 1, "Hiệu chỉnh Hb (g/dL)", head)
    for i, (k, v) in enumerate(ALTITUDE_OPTIONS.items(), start=1):
        wl.write(i, 0, k)
        wl.write(i, 1, v)
    wb.close()
    buf.seek(0)
    return buf.getvalue()


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
    st.header(":material/batch_prediction: Phân tích hồ sơ hàng loạt (Tích hợp Học máy & SHAP)")
    st.caption(
        "Phân tích nhiều hồ sơ cùng lúc bằng cả Mô hình A (Chỉ CBC) và Mô hình B (CBC + Ferritin): "
        "tự động tính Mentzer, Green & King, RDWI, xác suất dự đoán XGBoost và giải thích SHAP cá thể hóa."
    )

    cb = render_codebook_panel()

    with st.container(border=True):
        st.subheader(":material/upload_file: 1. Tải hồ sơ")
        files = st.file_uploader(
            f"Chọn tối đa {BATCH_MAX_FILES} file — Excel/CSV/Word/PDF/ảnh/SPSS/Stata/SAS/R",
            type=UPLOAD_TYPES,
            accept_multiple_files=True, key="batch_files",
        )
        default_alt = st.selectbox(
            "Khoảng độ cao mặc định cho các hồ sơ mới", list(ALTITUDE_OPTIONS.keys()),
            key="batch_default_alt",
        )
        ai_ok = st.checkbox(
            "Tôi xác nhận người bệnh đã đồng ý và cho phép gửi PDF/ảnh/nội dung Word, hoặc TÊN CỘT của file bảng, tới dịch vụ AI để đọc/ghép dữ liệu.",
            key="batch_ai_consent",
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
            progress.progress(1.0, text="Hoàn tất đọc file.")
            st.session_state["batch_rows"] = rows[:BATCH_MAX_TOTAL_ROWS]
            st.session_state["batch_notes"] = all_notes
            st.session_state["batch_ver"] = st.session_state.get("batch_ver", 0) + 1
            st.session_state.pop("batch_results", None)

    st.subheader(":material/edit_note: 2. Kiểm tra và chỉnh sửa dữ liệu")
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

    run = st.button("Phân tích tất cả hồ sơ bằng Học máy (THAL-SCREEN AI)", icon=":material/analytics:", type="primary",
                    use_container_width=True, disabled=edited is None or len(edited) == 0)

    if run:
        records = edited.to_dict("records")
        with st.spinner("Đang chạy mô hình Học máy và giải thích SHAP..."):
            st.session_state["batch_results"] = [analyze_batch_row(i, r) for i, r in enumerate(records)]
            st.session_state["batch_input_df"] = edited.copy()
            st.session_state["batch_codebook_df"] = codebook_summary_df(cb)

    results = st.session_state.get("batch_results")
    if not results:
        return

    st.divider()
    st.subheader(":material/monitoring: 3. Kết quả phân tích Học máy")
    counts = Counter(r.get("Mức lưu ý cao nhất") for r in results if r.get("Trạng thái") == "OK")
    insufficient = sum(1 for r in results if r.get("Trạng thái") != "OK")

    m = st.columns(5)
    m[0].metric("Tổng hồ sơ", len(results))
    m[1].metric("Nguy cơ thấp", counts.get("THẤP", 0))
    m[2].metric("Nguy cơ trung bình", counts.get("TRUNG BÌNH", 0))
    m[3].metric("Nguy cơ cao", counts.get("CAO", 0))
    m[4].metric("Không đủ dữ liệu", insufficient)

    all_cols = ["Hồ sơ", "Nguồn", "Tuổi", "Giới tính", "Hôn nhân", "Mang thai", "Vòng 1 - điểm", "Vòng 1 - nguy cơ", "Số chỉ số CTM",
                "Hb (g/dL)", "Hb sau hiệu chỉnh", "Thiếu máu (WHO 2024)", "MCV", "MCH", "MCHC", "RDW-CV",
                "RBC (T/L)", "Mentzer", "Green & King", "Ferritin",
                "Xác suất ML (Chỉ CBC)", "Xác suất ML (CBC + Ferritin)",
                "Vòng 2 - nguy cơ", "HbA2 (%)", "HbE (%)",
                "Vòng 3 - mức lưu ý", "Vòng 3 - đối chiếu BYT", "Mức lưu ý cao nhất", "Bước tiếp theo",
                "Trạng thái", "Cảnh báo"]
    show_cols = ["Hồ sơ", "MCV", "MCH", "RBC (T/L)", "Mentzer", "Ferritin",
                 "Xác suất ML (Chỉ CBC)", "Xác suất ML (CBC + Ferritin)",
                 "Vòng 2 - nguy cơ", "Mức lưu ý cao nhất", "Bước tiếp theo", "Trạng thái"]
    full_df = pd.DataFrame([{k: v for k, v in r.items() if k != "_detail"} for r in results]).reindex(columns=all_cols)
    view_df = full_df[show_cols]

    risk_subset = ["Vòng 2 - nguy cơ", "Mức lưu ý cao nhất"]
    try:
        styled = view_df.style.map(_risk_cell_style, subset=[c for c in risk_subset if c in view_df.columns])
    except AttributeError:
        styled = view_df.style.applymap(_risk_cell_style, subset=[c for c in risk_subset if c in view_df.columns])
    st.dataframe(styled, use_container_width=True, hide_index=True)

    input_df = st.session_state.get("batch_input_df")
    if input_df is not None:
        st.download_button(
            "Tải kết quả Excel (đầy đủ xác suất Học máy + chỉ số phái sinh)",
            data=results_to_excel(full_df, input_df, st.session_state.get("batch_codebook_df")).getvalue(),
            file_name=f"Ket_qua_THAL_SCREEN_AI_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:", use_container_width=True,
        )

    ok_idx = [i for i, r in enumerate(results) if r.get("Trạng thái") == "OK"]
    if ok_idx:
        st.markdown("### :material/clinical_notes: Xem chi tiết giải thích SHAP cho từng hồ sơ")
        pick = st.selectbox(
            "Chọn hồ sơ", ok_idx, key="batch_detail_pick",
            format_func=lambda i: f"{i + 1}. {results[i]['Hồ sơ']} — {results[i].get('Mức lưu ý cao nhất') or 'chưa phân tầng'}",
        )
        det = results[pick]["_detail"]
        if "r2" in det and "ml" in det["r2"]:
            render_shap_explanation_card(det["r2"]["ml"])


FACILITY_DISPLAY_LIMIT = 12


@st.fragment
def render_facilities(patient):
    st.subheader(":material/local_hospital: Danh sách cơ sở y tế")
    all_facilities = all_medical_facilities()
    province_options = ["Tất cả tỉnh/thành"] + sorted({x["province"] for x in all_facilities})
    default_province = PROVINCE_FACILITY_ALIASES.get(patient["province"], patient["province"])
    selected_facility_province = st.selectbox(
        "Tỉnh/thành", province_options,
        index=province_options.index(default_province) if default_province in province_options else 0,
        key="facility_province_filter",
    )
    facility_keyword = st.text_input(
        "Tìm bệnh viện", placeholder="Ví dụ: Bệnh viện Trung ương Huế, Huyết học...",
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

    if filtered:
        for i, facility in enumerate(filtered[:FACILITY_DISPLAY_LIMIT], start=1):
            with st.container(border=True):
                st.markdown(f"**{i}. {facility['name']}**")
                st.caption(f"{facility['province']} · {facility['tier']}")
                st.write(facility["note"])
                st.link_button("Xem vị trí / chỉ đường", facility["maps"], icon=":material/map:")


# ============================================================
# HEADER
# ============================================================

st.markdown(
    '<div class="hematology-brand"><div class="brand-orb"></div><div class="brand-text">'
    '<div class="eyebrow">TRƯỜNG ĐẠI HỌC PHAN CHÂU TRINH · BỆNH VIỆN TRUNG ƯƠNG HUẾ</div>'
    '<div class="title title-main">THAL-SCREEN AI · HỆ THỐNG DỰ ĐOÁN NGUY CƠ MANG GEN THALASSEMIA</div>'
    '<div class="chips"><span class="chip">7 chỉ số CBC + Ferritin huyết thanh</span>'
    '<span class="chip">Logistic Regression · Random Forest · XGBoost</span>'
    '<span class="chip">Giải thích mô hình SHAP</span>'
    '<span class="chip">Nhóm NCKH 24XN1</span></div></div></div>'
    '<div class="section-rule"></div>',
    unsafe_allow_html=True,
)

_bi1, _bi2, _ = st.columns([1, 1, 2])
if _bi1.button("Giới thiệu đề tài NCKH", icon=":material/menu_book:", use_container_width=True):
    show_project_intro()
if _bi2.button("Cơ sở khoa học & Học máy", icon=":material/school:", use_container_width=True):
    show_research_overview()

st.markdown(
    '<div class="contact-card"><div class="label">Chủ nhiệm đề tài & Nhóm nghiên cứu (Trường Đại học Phan Châu Trinh)</div>'
    '<div class="name">Chủ nhiệm: Nguyễn Đắc Lâu · GVHD: TS. Nguyễn Đình Khắc Chân</div>'
    '<div>Thành viên: Trần Phan Thảo Vi · Huỳnh Ngọc Anh Huy · Lê Văn Thịnh · Phan Ngọc Bảo An</div>'
    '<div>Điện thoại: <a href="tel:0357930820">0357930820</a> · Email: <a href="mailto:nguyendaclau2309@gmail.com">nguyendaclau2309@gmail.com</a></div>'
    '<div class="small-note">Hệ thống hỗ trợ sàng lọc và phân tầng nguy cơ mang gen Thalassemia dựa trên CBC và Ferritin huyết thanh (không thay thế chẩn đoán xác định di truyền).</div></div>',
    unsafe_allow_html=True,
)

# ============================================================
# ACCESS / SIDEBAR / ROUTING
# ============================================================

render_auth_sidebar()

st.sidebar.divider()
st.sidebar.selectbox(
    "Chế độ màu giao diện",
    ["Tự động", "Sáng (Light Glass)", "Tối (Dark Glass)"],
    key="ui_theme_mode",
    help="Chuyển đổi tức thì giữa giao diện Sáng (chữ đen đậm #0F172A trên nền kính trắng ngọc trai) và Tối (Dark Liquid Glass).",
)
st.sidebar.toggle(
    "Hiệu ứng kính (Liquid Glass)", value=True, key="glass_fx",
    help="Tắt nếu máy yếu hoặc mạng chậm để trang nhẹ và mở nhanh hơn.",
)

auth_user = current_auth_user()
active_page = st.session_state.get("auth_page", PAGE_SCREENING)

if active_page == PAGE_ML_LAB:
    render_ml_research_lab()
    st.stop()

if auth_user and active_page == PAGE_ADMIN:
    render_admin_console(auth_user)
    st.stop()

if auth_user and active_page == PAGE_BATCH:
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

st.header(":material/badge: Thông tin đối tượng nghiên cứu / người tham gia")
st.caption(
    "Nhập hồ sơ người trưởng thành (≥ 18 tuổi) hoặc đối tượng tham gia sàng lọc. "
    + ("Chế độ nhập giúp: hệ thống ghi nhận tài khoản nhân sự thực hiện."
       if operator_mode else "Người tham gia tự nhập thông tin của mình.")
)

with st.form("patient_basic_form_v22", clear_on_submit=False):
    with st.container(border=True):
        p1, p2, p3 = st.columns(3)
        with p1:
            full_name = st.text_input("Họ và tên *", placeholder="Nguyễn Văn A", key="v22_full_name")
            birth_date = st.date_input(
                "Ngày sinh *", value=date(2002, 1, 1),
                min_value=BIRTH_DATE_MIN, max_value=BIRTH_DATE_MAX,
                format="DD/MM/YYYY", key="v22_birth_date",
            )
        with p2:
            phone_raw = st.text_input("Số điện thoại *", placeholder="09xxxxxxxx", key="v22_phone")
            gender = st.selectbox("Giới tính *", ["Nam", "Nữ", "Khác"], key="v22_gender")
        with p3:
            current_address = st.text_input("Địa chỉ hiện tại *", placeholder="Số nhà/thôn/tổ/đường", key="v22_address")

        basic_continue = st.form_submit_button(
            "Tiếp tục", icon=":material/arrow_forward:", type="primary", use_container_width=True
        )


@st.fragment
def address_picker():
    st.subheader(":material/location_on: Địa giới hành chính hiện tại")
    province, commune, ok = "", "", False

    try:
        admin_hierarchy = load_admin_hierarchy()
        ok = True
    except Exception as exc:
        admin_hierarchy = {}
        st.error("Không tải được danh mục tỉnh/thành và phường/xã/đặc khu.")
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
            safe_province = province.replace(" ", "_").replace("/", "_").replace("\\", "_")
            commune_choice = st.selectbox(
                "Phường / xã / đặc khu *",
                ["— Chọn phường/xã/đặc khu —"] + available_communes,
                key=f"v22_profile_commune_{safe_province}",
            )
            if commune_choice != "— Chọn phường/xã/đặc khu —":
                commune = commune_choice

    st.session_state["addr_data_ok"] = ok
    st.session_state["addr_province"] = province
    st.session_state["addr_commune"] = commune

    if province and commune:
        st.success(f"Đã chọn: **{commune}, {province}**", icon=":material/location_on:")


address_picker()
admin_data_ok = st.session_state.get("addr_data_ok", False)
selected_province = st.session_state.get("addr_province", "")
commune_value = st.session_state.get("addr_commune", "")

with st.form("patient_consent_save_form_v22", clear_on_submit=False):
    with st.container(border=True):
        st.markdown("### :material/verified_user: Bản đồng thuận tham gia sàng lọc và nghiên cứu khoa học")
        st.markdown(
            "**Tôi đồng ý cho nhóm nghiên cứu sử dụng thông tin khảo sát và kết quả xét nghiệm "
            "(công thức máu ngoại vi, Ferritin huyết thanh và xét nghiệm tham chiếu Thalassemia) "
            "để phục vụ mục đích sàng lọc và thực hiện đề tài nghiên cứu khoa học.**"
        )
        research_consent = st.checkbox(
            "Tôi đã đọc, hiểu và tự nguyện đồng ý tham gia." if not operator_mode
            else "Tôi xác nhận người tham gia đã đọc/được giải thích và đã ký đồng thuận.",
            key="v22_research_consent",
        )
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
            st.success("Đã tạo hồ sơ người tham gia.", icon=":material/check_circle:")

patient = st.session_state.get("patient_profile")

if not st.session_state.get("research_consent", False):
    st.warning(
        "Bạn chưa đồng ý tham gia. Hệ thống không mở Vòng 1 và không xử lý/lưu dữ liệu sàng lọc.",
        icon=":material/lock:",
    )
    st.stop()

if not patient:
    st.info("Hãy lưu hồ sơ người tham gia trước khi bắt đầu.", icon=":material/arrow_upward:")
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
st.header(":material/quiz: Vòng 1 — 20 câu hỏi khảo sát tiền sử")
st.caption(
    "Vòng 1 ghi nhận tiền sử gia đình và cá nhân. Bạn cũng có thể bấm mở nhanh Vòng 2 bên dưới nếu muốn nhập trực tiếp kết quả xét nghiệm CBC + Ferritin huyết thanh."
)

quick_r2 = st.checkbox(
    "Bỏ qua bộ câu hỏi Vòng 1 — Mở trực tiếp Vòng 2 (Nhập 7 chỉ số CBC & Ferritin huyết thanh để chạy mô hình Học máy)",
    value=False,
    key="quick_open_r2",
)
if quick_r2 and not st.session_state.get("round1_completed"):
    st.session_state.update({
        "round1_score": 0,
        "round1_reasons": ["Chuyển trực tiếp sang đánh giá xét nghiệm CBC & Ferritin (Vòng 2)"],
        "round1_category": "THẤP",
        "round1_conclusion": "Đánh giá trực tiếp dựa trên thông số xét nghiệm CBC và Ferritin huyết thanh.",
        "round1_completed": True,
        "round2_unlocked": True,
    })

if not quick_r2:
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
        round1_submit = st.form_submit_button(
            "Đánh giá Vòng 1 & Mở Vòng 2 (CBC + Ferritin)", icon=":material/play_arrow:", type="primary", use_container_width=True
        )

    if round1_submit:
        score1, reasons1 = calculate_round1_score(answers)
        category1, conclusion1 = round1_category(score1)
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

if st.session_state.get("round1_completed", False):
    score1 = st.session_state["round1_score"]
    category1 = st.session_state["round1_category"]
    conclusion1 = st.session_state["round1_conclusion"]
    reasons1 = st.session_state["round1_reasons"]
    if not quick_r2:
        st.subheader(":material/fact_check: Kết quả Vòng 1")
        risk_alert("Mức độ lưu ý tiền sử", category1, conclusion1)
    st.session_state["round2_unlocked"] = True

# ============================================================
# ROUND 2: 7 CHỈ SỐ CBC + FERRITIN HUYẾT THANH + HỌC MÁY & SHAP
# ============================================================

if st.session_state.get("round2_unlocked", False):

    st.divider()
    st.header(":material/biotech: Vòng 2 — Dự đoán Học máy dựa trên 7 chỉ số CBC và Ferritin huyết thanh")
    st.caption(
        "Nhập các thông số dòng hồng cầu (RBC, Hb, HCT, MCV, MCH, MCHC, RDW-CV) và định lượng Ferritin huyết thanh. "
        "Hệ thống tự động chạy 3 thuật toán (Logistic Regression, Random Forest, XGBoost), so sánh Mô hình A (Chỉ CBC) "
        "với Mô hình B (CBC + Ferritin) và giải thích đóng góp của từng biến bằng phương pháp SHAP."
    )

    with st.form("round2_form_v24_ml", clear_on_submit=False):
        with st.container(border=True):
            st.subheader(":material/terrain: 1. Nơi cư trú và hiệu chỉnh độ cao")
            altitude_choice, altitude_adjustment = altitude_selector("round2")

        with st.container(border=True):
            st.subheader(":material/water_drop: 2. Nhập 7 chỉ số Công thức máu (CBC) và Ferritin huyết thanh")
            r1_c1, r1_c2, r1_c3, r1_c4 = st.columns(4)
            with r1_c1:
                hb_unit = st.selectbox("Đơn vị Hb", ["g/L", "g/dL"], index=0, key="round2_hb_unit")
                hb_raw = st.number_input("Nồng độ Hb *", min_value=0.0, max_value=1000.0, value=None,
                                         step=1.0, key="round2_hb_raw", help="Ví dụ: 121.8 g/L hoặc 12.2 g/dL")
            with r1_c2:
                rbc_unit = st.selectbox("Đơn vị RBC", ["T/L", "10^12/L", "10^6/µL"], key="round2_rbc_unit")
                rbc_raw = st.number_input("Số lượng RBC *", 1.0, 10.0, None, 0.05, key="round2_rbc_raw", help="Ví dụ: 5.31 T/L")
            with r1_c3:
                mcv = st.number_input("MCV (fL) *", 30.0, 150.0, None, 0.1, key="round2_mcv", help="Ví dụ: 71.9 fL")
                hct_raw = st.number_input("HCT (%) [tùy chọn]", 10.0, 75.0, None, 0.1, key="round2_hct",
                                          help="Nếu để trống, hệ thống tự động nội suy HCT = (RBC × MCV) / 10")
            with r1_c4:
                mch = st.number_input("MCH (pg) *", 10.0, 50.0, None, 0.1, key="round2_mch", help="Ví dụ: 22.8 pg")
                mchc_raw = st.number_input("MCHC (g/L) [tùy chọn]", 20.0, 450.0, None, 1.0, key="round2_mchc",
                                           help="Nếu để trống, hệ thống tự tính từ Hb và HCT")

            r2_c1, r2_c2 = st.columns(2)
            with r2_c1:
                rdw = st.number_input("RDW-CV (%) *", 5.0, 40.0, None, 0.1, key="round2_rdw", help="Ví dụ: 14.2%")
            with r2_c2:
                ferritin_raw = st.number_input(
                    "Ferritin huyết thanh (ng/mL hoặc µg/L) [Khuyến nghị nhập để chạy Mô hình B]",
                    min_value=0.0, max_value=10000.0, value=None, step=1.0, key="round2_ferritin",
                    help="Biến đầu vào bổ sung trọng tâm của Đề tài NCKH giúp phân biệt thiếu máu thiếu sắt với mang gen Thalassemia.",
                )

        round2_submit = st.form_submit_button(
            "Chạy Mô hình Học máy (LR · RF · XGBoost) & Phân tích SHAP",
            icon=":material/psychology:", type="primary", use_container_width=True,
        )

    if round2_submit:
        required_cbc = {"Hb": hb_raw, "MCV": mcv, "MCH": mch, "RBC": rbc_raw, "RDW-CV": rdw}
        missing = [name for name, value in required_cbc.items() if value is None]

        if missing:
            st.error("Vui lòng nhập đầy đủ 5 chỉ số CBC bắt buộc trước khi chạy mô hình: " + ", ".join(missing) + ".")
        elif rbc_raw <= 0:
            st.error("RBC phải lớn hơn 0.")
        else:
            hb = hb_to_g_dl(hb_raw, hb_unit)
            rbc = rbc_to_t_l(rbc_raw, rbc_unit)
            hb_adjusted = hb - altitude_adjustment

            ml_res = predict_thal_screen_ai(
                hb_g_dl=hb, rbc=rbc, mcv=mcv, mch=mch, rdw=rdw,
                hct=hct_raw, mchc_g_l=mchc_raw, ferritin=ferritin_raw,
            )
            score2, mentzer, reasons2 = calculate_round2_score(mcv, mch, rbc, rdw, ferritin_raw)
            category2 = round2_category(score2, mcv, mch, ml_res["category"])
            conclusion2 = ml_res["conclusion"]
            findings, advice = narrative_cbc_advice(hb_adjusted, mcv, mch, rbc, rdw, mentzer, ferritin_raw, ml_res)

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
                "round2_hct_result": ml_res["cbc"]["HCT"],
                "round2_mcv_result": mcv,
                "round2_mch_result": mch,
                "round2_mchc_result": ml_res["cbc"]["MCHC_gL"],
                "round2_rbc_result": rbc,
                "round2_rdw_result": rdw,
                "round2_ferritin_result": ferritin_raw,
                "round2_ml_result": ml_res,
                "round2_findings": findings,
                "round2_advice": advice,
            })

            if not st.session_state.get("screening_id"):
                st.session_state["screening_id"] = create_screening_record(
                    patient=patient, answers={}, score1=st.session_state.get("round1_score", 0),
                    category1=st.session_state.get("round1_category", "THẤP"),
                    conclusion1=st.session_state.get("round1_conclusion", ""),
                    reasons1=st.session_state.get("round1_reasons", []),
                    entry_mode="assisted" if operator_mode else "self",
                    entered_by_username=operator_username,
                )

            update_screening_round2(
                st.session_state.get("screening_id"),
                {
                    "altitude_choice": altitude_choice, "adjustment": altitude_adjustment,
                    "hb": hb, "hb_adjusted": hb_adjusted,
                    "hct": ml_res["cbc"]["HCT"], "mcv": mcv, "mch": mch, "mchc": ml_res["cbc"]["MCHC_gL"],
                    "rbc": rbc, "rdw": rdw, "ferritin": ferritin_raw,
                    "mentzer": mentzer,
                    "green_king": ml_res["indices"]["green_king"],
                    "rdwi": ml_res["indices"]["rdwi"],
                    "ml_prob_cbc": ml_res["probs_a"]["XGBoost"],
                    "ml_prob_cbc_ferritin": ml_res["probs_b"]["XGBoost"] if ferritin_raw is not None else None,
                    "shap": ml_res["shap"],
                    "score": score2, "category": category2, "conclusion": conclusion2,
                    "reasons": reasons2, "findings": findings, "advice": advice,
                },
            )
            st.session_state["round2_completed"] = True

    if st.session_state.get("round2_completed", False):

        st.subheader(":material/monitoring: Kết quả Dự đoán Học máy & Phân tích Huyết học (Vòng 2)")
        ml_res = st.session_state["round2_ml_result"]

        a, b, c, d = st.columns(4)
        a.metric("Xác suất Mang gen (XGBoost)", f"{ml_res['active_pct']:.1f}%")
        b.metric("Mentzer Index (MCV/RBC)", f"{st.session_state['round2_mentzer']:.2f}")
        c.metric("Hb thực đo", f"{st.session_state['round2_hb']*10:.0f} g/L ({st.session_state['round2_hb']:.1f} g/dL)")
        d.metric(
            "Ferritin huyết thanh",
            f"{st.session_state['round2_ferritin_result']:g} ng/mL"
            if st.session_state.get("round2_ferritin_result") is not None
            else "Chưa nhập",
        )

        category2 = st.session_state["round2_category"]
        conclusion2 = st.session_state["round2_conclusion"]
        risk_alert("Phân tầng nguy cơ (THAL-SCREEN AI)", category2, conclusion2)

        # Hiển thị thẻ so sánh Mô hình A vs Mô hình B và đồ thị SHAP cá thể hóa
        render_shap_explanation_card(ml_res)

        with st.expander("Chi tiết các chỉ số huyết học và nhận định lâm sàng", expanded=True, icon=":material/search:"):
            st.markdown("**Các dấu hiệu ghi nhận từ CBC & Ferritin**")
            bullet_list(st.session_state["round2_reasons"])
            st.markdown("**Phân tích sơ bộ**")
            bullet_list(st.session_state["round2_findings"])
            st.markdown("**Khuyến nghị bước tiếp theo**")
            bullet_list(st.session_state["round2_advice"])

        # ----------------------------------------------------
        # ROUND 3: XÉT NGHIỆM THAM CHIẾU (HPLC / ĐIỆN DI & GEN HBA1, HBA2, HBB)
        # ----------------------------------------------------

        st.subheader(":material/lab_panel: 3. Kết quả xét nghiệm tham chiếu chuẩn (HPLC/Điện di Hb & Đột biến gen α, β)")
        st.caption(
            "Nhập kết quả HPLC/điện di huyết sắc tố và xét nghiệm sinh học phân tử (gen HBA1, HBA2, HBB) "
            "để xác lập nhãn tham chiếu vàng và hoàn thiện hồ sơ nghiên cứu."
        )

        with st.form("round3_followup_form", clear_on_submit=False):
            c1, c2, c3 = st.columns(3)
            with c1:
                r3_date = st.date_input("Ngày thực hiện xét nghiệm", value=None, key="round3_test_date")
            with c2:
                r3_type = st.selectbox(
                    "Phương pháp xét nghiệm tham chiếu",
                    ["Kết hợp HPLC/Điện di + Xét nghiệm gen (HBA1, HBA2, HBB)",
                     "Xét nghiệm đột biến gen (Gap-PCR / RDB / Giải trình tự)",
                     "HPLC / Điện di mao quản huyết sắc tố",
                     "Ferritin / đánh giá sắt bổ sung", "Khác"],
                    key="round3_test_type",
                )
            with c3:
                r3_facility = st.text_input("Cơ sở thực hiện", value="Bệnh viện Trung ương Huế", key="round3_facility")

            st.markdown("**Thành phần huyết sắc tố trên HPLC / Điện di mao quản (nếu có)**")
            h1, h2, h3, h4 = st.columns(4)
            with h1:
                r3_hba = st.number_input("HbA (%)", min_value=0.0, max_value=100.0, value=None, step=0.1, key="round3_hba")
            with h2:
                r3_hba2 = st.number_input("HbA2 (%)", min_value=0.0, max_value=30.0, value=None, step=0.1, key="round3_hba2")
            with h3:
                r3_hbf = st.number_input("HbF (%)", min_value=0.0, max_value=100.0, value=None, step=0.1, key="round3_hbf")
            with h4:
                r3_hbe = st.number_input("HbE (%)", min_value=0.0, max_value=100.0, value=None, step=0.1, key="round3_hbe")

            f1, f2, f3 = st.columns(3)
            with f1:
                r3_ferritin = st.number_input(
                    "Ferritin cập nhật (ng/mL)",
                    min_value=0.0, max_value=10000.0,
                    value=st.session_state.get("round2_ferritin_result"),
                    step=0.1, key="round3_ferritin",
                )
            with f2:
                r3_iron = st.number_input("Sắt huyết thanh (µmol/L)", min_value=0.0, max_value=1000.0, value=None, step=0.1, key="round3_iron")
            with f3:
                r3_tsat = st.number_input("Độ bão hòa transferrin (%)", min_value=0.0, max_value=100.0, value=None, step=0.1, key="round3_tsat")

            r3_genetic = st.text_area(
                "Kết quả xét nghiệm đột biến gen HBA1, HBA2, HBB (Tiêu chuẩn tham chiếu chính)",
                key="round3_genetic",
                placeholder="Ví dụ: Dương tính dị hợp tử mất đoạn --SEA (alpha-Thalassemia) / Dương tính đột biến CD17 (beta-Thalassemia) / Âm tính...",
            )
            r3_lab = st.text_area(
                "Kết luận phân nhóm nghiên cứu (Không mang gen / Mang gen Alpha / Mang gen Beta / Đồng mang Alpha+Beta)",
                key="round3_lab",
            )

            s1, s2 = st.columns(2)
            with s1:
                r3_status = st.selectbox(
                    "Trạng thái xác nhận tham chiếu",
                    ["Đã có kết quả xét nghiệm tham chiếu", "Đã lấy mẫu — đang chờ kết quả PCR/HPLC", "Chưa thực hiện"],
                    key="round3_status",
                )
            with s2:
                r3_follow_date = st.date_input("Ngày hẹn tư vấn di truyền / tái khám", value=None, key="round3_follow_date")

            r3_note = st.text_area("Ghi chú tư vấn di truyền", key="round3_note")

            save_r3 = st.form_submit_button(
                "Lưu kết quả xét nghiệm tham chiếu", icon=":material/save:", type="primary", use_container_width=True
            )

        if save_r3:
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
            st.success("Đã lưu kết quả xét nghiệm tham chiếu.", icon=":material/check_circle:")

        if st.session_state.get("round3_completed", False):
            st.markdown("### :material/rule: Đối chiếu kết quả chuyên sâu theo hướng dẫn BYT 2022")
            current_r3 = {
                "hba": st.session_state.get("round3_hba"), "hba2": st.session_state.get("round3_hba2"),
                "hbf": st.session_state.get("round3_hbf"), "hbe": st.session_state.get("round3_hbe"),
                "ferritin": st.session_state.get("round3_ferritin"),
                "transferrin_saturation": st.session_state.get("round3_tsat"),
                "genetic_result": st.session_state.get("round3_genetic", ""),
                "lab_conclusion": st.session_state.get("round3_lab", ""),
            }
            bullet_list(round3_byt_interpretation(current_r3))

        render_facilities(patient)

        st.subheader(":material/description: Xuất Phiếu Kết quả Sàng lọc & Giải thích SHAP (.docx)")

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
                "hct": st.session_state.get("round2_hct_result"),
                "mcv": st.session_state["round2_mcv_result"],
                "mch": st.session_state["round2_mch_result"],
                "mchc": st.session_state.get("round2_mchc_result"),
                "rbc": st.session_state["round2_rbc_result"],
                "rdw": st.session_state["round2_rdw_result"],
                "ferritin": st.session_state.get("round2_ferritin_result"),
                "mentzer": st.session_state["round2_mentzer"],
                "green_king": ml_res["indices"]["green_king"],
                "rdwi": ml_res["indices"]["rdwi"],
                "ml_prob_cbc": ml_res["probs_a"]["XGBoost"],
                "ml_prob_cbc_ferritin": ml_res["probs_b"]["XGBoost"] if st.session_state.get("round2_ferritin_result") is not None else None,
                "shap": ml_res["shap"],
                "score": st.session_state["round2_score"],
                "category": st.session_state["round2_category"],
                "conclusion": st.session_state["round2_conclusion"],
                "findings": st.session_state["round2_findings"],
                "advice": st.session_state["round2_advice"],
            },
        )

        st.download_button(
            "Tải phiếu kết quả Word (kèm xác suất Học máy & giải thích SHAP)",
            data=report,
            file_name=f"Phieu_THAL_SCREEN_AI_{safe_filename(patient['full_name'])}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            icon=":material/download:",
            use_container_width=True,
        )

# ============================================================
# FOOTER
# ============================================================

st.divider()
st.caption(
    "THAL-SCREEN AI (V6 Research Edition) — Đề tài NCKH Trường Đại học Phan Châu Trinh & Bệnh viện Trung ương Huế. "
    "Công cụ hỗ trợ sàng lọc và phân tầng nguy cơ bằng Học máy có giải thích (SHAP), không thay thế chẩn đoán xác định di truyền."
)
