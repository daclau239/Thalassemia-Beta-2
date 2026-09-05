
import io
import math
import re
import sqlite3
from datetime import date, datetime, timedelta

import requests
import streamlit as st
from docx import Document

# ============================================================
# THALASSEMIA SCREENING - ĐẮC LÂU DỄ THƯƠNG
# ============================================================
#
# Ý tưởng:
#   1. Hồ sơ bệnh nhân
#   2. Vòng 1: 20 câu hỏi nguy cơ
#   3. Chỉ nguy cơ CAO -> mở Vòng 2
#   4. Vòng 2:
#        - người dùng chọn Tỉnh/Thành phố
#        - người dùng chọn trực tiếp Phường/Xã/Đặc khu
#        - hệ thống lấy tọa độ của chính lựa chọn đó
#        - Google Elevation lấy độ cao
#        - nhập CBC
#        - tính Mentzer + CBC screening score
#        - hiệu chỉnh Hb theo độ cao (prototype)
#        - đề xuất 3–5 cơ sở y tế gần nhất
#   5. Nguy cơ thấp:
#        - hiển thị bảng theo dõi sức khỏe
#        - có khu vực nhập CBC nếu về sau có bất thường
#
# ĐỊA GIỚI:
#   Việt Nam hiện có 34 đơn vị cấp tỉnh và 3.321 đơn vị cấp xã;
#   cấp huyện đã được bãi bỏ từ 01/07/2025.
#   Ứng dụng tải hierarchy.json từ Open Admin Data (CC-BY-4.0)
#   khi chạy, nên không cần hard-code 3.321 xã/phường trong code.
#
# GOOGLE:
#   .streamlit/secrets.toml
#   [google]
#   maps_api_key = "YOUR_GOOGLE_MAPS_API_KEY"
#
# Cần bật:
#   - Geocoding API
#   - Elevation API
#   - Places API (New)
#
# CÀI:
#   pip install streamlit python-docx requests
#
# LƯU Ý:
#   - Risk score chỉ là PROTOTYPE, chưa validation trên quần thể Việt Nam.
#   - Không dùng để tự chẩn đoán hay điều trị.
#   - SQLite là giải pháp lưu hồ sơ cho prototype. Streamlit Cloud có thể
#     reset filesystem khi app được rebuild/redeploy; nếu cần lưu lâu dài,
#     nên thay DB bằng Supabase/PostgreSQL/Google Sheets có kiểm soát truy cập.
# ============================================================


# ============================================================
# CONFIG
# ============================================================

st.set_page_config(
    page_title="Hệ thống Sàng lọc Thalassemia",
    page_icon="🩸",
    layout="wide",
)

DB_PATH = "thalassemia_patients.db"

ADMIN_URL = (
    "https://raw.githubusercontent.com/"
    "open-admin-data/vietnam-administrative-divisions/"
    "refs/heads/main/data/hierarchy.json"
)

ROUND1_MAX_SCORE = 20
ROUND1_HIGH_THRESHOLD = 8
FOLLOWUP_DAYS = 30


# ============================================================
# GOOGLE KEY
# ============================================================

def get_google_key():
    try:
        return st.secrets["google"]["maps_api_key"]
    except Exception:
        pass

    try:
        return st.secrets["GOOGLE_MAPS_API_KEY"]
    except Exception:
        return ""


GOOGLE_API_KEY = get_google_key()


# ============================================================
# DATABASE
# ============================================================

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
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )

    conn.commit()
    return conn


def normalize_phone(value):
    digits = re.sub(
        r"\D",
        "",
        value or "",
    )

    if (
        digits.startswith("84")
        and len(digits) == 11
    ):
        digits = "0" + digits[2:]

    return digits


def valid_vietnam_phone(phone):
    return bool(
        re.fullmatch(
            r"0\d{9}",
            phone,
        )
    )


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
    """
    Một số điện thoại = một hồ sơ.
    Cùng số nhập lại -> UPDATE, lần nhập cuối cùng thắng.
    """

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
                now,
                profile["phone"],
            ),
        )

        action = "updated"

    else:

        conn.execute(
            """
            INSERT INTO patient_profiles
            (
                phone,
                full_name,
                birth_date,
                gender,
                current_address,
                province,
                commune,
                created_at,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                profile["phone"],
                profile["full_name"],
                profile["birth_date"],
                profile["gender"],
                profile["current_address"],
                profile["province"],
                profile["commune"],
                now,
                now,
            ),
        )

        action = "inserted"

    conn.commit()
    conn.close()

    return action


# ============================================================
# UTILITIES
# ============================================================

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
    text = (text or "").strip()

    text = re.sub(
        r"[^0-9A-Za-zÀ-ỹĐđ _-]",
        "_",
        text,
    )

    return (
        text.strip(" _")
        or "nguoi_sang_loc"
    )


def clear_survey_results():
    for key in list(
        st.session_state.keys()
    ):
        if (
            key.startswith("round1_")
            or key.startswith("round2_")
            or key.startswith("location_")
            or key.startswith("low_cbc_")
        ):
            del st.session_state[key]


# ============================================================
# ADMIN DATA — 34 PROVINCES + 3,321 COMMUNES
# ============================================================

@st.cache_data(ttl=86400)
def load_vietnam_admin():
    """
    Tải danh mục hiện hành từ Open Admin Data.
    Cấu trúc dữ liệu:
      {
        "_attribution": ...,
        "data": [
          {
            "name": {"local": "..."},
            "ward": [
              {"name": {"local": "..."}}
            ]
          }
        ]
      }

    Nếu không tải được, dùng danh sách fallback nhỏ để app không crash.
    """

    fallback = {
        "Đà Nẵng": [
            "Hòa Phong",
            "Hòa Phú",
            "Hòa Sơn",
            "Hòa Tiến",
        ],
        "Hà Nội": [
            "Ba Đình",
            "Ngọc Hà",
            "Giảng Võ",
        ],
        "Hồ Chí Minh": [
            "Bến Nghé",
            "Bến Thành",
            "Tân Định",
        ],
    }

    try:

        response = requests.get(
            ADMIN_URL,
            timeout=15,
        )

        response.raise_for_status()

        raw = response.json()

        result = {}

        for province in raw.get(
            "data",
            [],
        ):

            province_name = (
                province.get(
                    "name",
                    {},
                ).get(
                    "local"
                )
            )

            if not province_name:
                continue

            communes = []

            for ward in province.get(
                "ward",
                [],
            ):

                ward_name = (
                    ward.get(
                        "name",
                        {},
                    ).get(
                        "local"
                    )
                )

                if ward_name:
                    communes.append(
                        ward_name
                    )

            result[
                province_name
            ] = sorted(
                set(communes),
                key=lambda x: x.lower(),
            )

        if (
            len(result) >= 30
            and sum(
                len(v)
                for v in result.values()
            ) >= 3000
        ):
            return result, True

        return fallback, False

    except Exception:
        return fallback, False


ADMIN_DATA, ADMIN_DATA_OK = (
    load_vietnam_admin()
)


def provinces():
    return sorted(
        ADMIN_DATA.keys(),
        key=lambda x: x.lower(),
    )


def communes_for(province):
    return ADMIN_DATA.get(
        province,
        [],
    )


# ============================================================
# GOOGLE GEOCODING
# ============================================================

@st.cache_data(ttl=86400)
def google_geocode(
    address,
    api_key,
):
    if not api_key:
        return None

    if not address.strip():
        return None

    url = (
        "https://maps.googleapis.com/maps/api/geocode/json"
    )

    params = {
        "address": address,
        "key": api_key,
        "language": "vi",
        "region": "vn",
    }

    try:

        response = requests.get(
            url,
            params=params,
            timeout=12,
        )

        response.raise_for_status()

        data = response.json()

        if (
            data.get("status") != "OK"
            or not data.get("results")
        ):
            return None

        result = data["results"][0]

        location = (
            result["geometry"][
                "location"
            ]
        )

        return {
            "formatted_address": result.get(
                "formatted_address",
                address,
            ),
            "lat": float(
                location["lat"]
            ),
            "lng": float(
                location["lng"]
            ),
            "place_id": result.get(
                "place_id",
                "",
            ),
        }

    except Exception:
        return None


# ============================================================
# GOOGLE ELEVATION
# ============================================================

@st.cache_data(ttl=86400)
def google_elevation(
    lat,
    lng,
    api_key,
):
    if not api_key:
        return None

    url = (
        "https://maps.googleapis.com/maps/api/elevation/json"
    )

    params = {
        "locations": f"{lat},{lng}",
        "key": api_key,
    }

    try:

        response = requests.get(
            url,
            params=params,
            timeout=12,
        )

        response.raise_for_status()

        data = response.json()

        if (
            data.get("status") != "OK"
            or not data.get("results")
        ):
            return None

        item = data["results"][0]

        return {
            "elevation": float(
                item["elevation"]
            ),
            "resolution": float(
                item.get(
                    "resolution",
                    0,
                )
            ),
        }

    except Exception:
        return None


# ============================================================
# WHO 2024 ALTITUDE ADJUSTMENT
# ============================================================

def hb_altitude_adjustment(
    elevation_m,
):
    """
    WHO 2024 equation used as a prototype:
      adjustment (g/L) =
          0.0056384 * elevation
          + 0.0000003 * elevation^2

    Convert g/L -> g/dL by /10.
    """

    if (
        elevation_m is None
        or elevation_m <= 0
    ):
        return 0.0

    adjustment_g_l = (
        0.0056384 * elevation_m
        + 0.0000003
        * elevation_m**2
    )

    return adjustment_g_l / 10.0


def altitude_band(
    elevation_m,
):
    if elevation_m is None:
        return "Chưa xác định"

    if elevation_m < 500:
        return "<500 m"

    if elevation_m < 1000:
        return "500–999 m"

    if elevation_m < 1500:
        return "1.000–1.499 m"

    if elevation_m < 2000:
        return "1.500–1.999 m"

    if elevation_m < 2500:
        return "2.000–2.499 m"

    return "≥2.500 m"


# ============================================================
# ROUND 1 SCORE
# ============================================================

def calculate_round1_score(
    answers,
):
    """
    PROTOTYPE.
    Q19-Q20 không cộng điểm vì chỉ là khả năng tiếp cận y tế.
    """

    score = 0
    reasons = []

    items = {
        "q1": (
            3,
            "Có người thân/dòng họ mắc Thalassemia",
        ),
        "q2": (
            3,
            "Có người thân/dòng họ mang gen Thalassemia/hemoglobinopathy",
        ),
        "q3": (
            1,
            "Cha/mẹ từng xét nghiệm Thalassemia/hemoglobinopathy",
        ),
        "q4": (
            2,
            "Anh/chị/em từng thiếu máu hoặc hồng cầu nhỏ",
        ),
        "q5": (
            2,
            "Gia đình có trẻ từng truyền máu nhiều lần/định kỳ",
        ),
        "q6": (
            1,
            "Từng được thông báo thiếu máu",
        ),
        "q7": (
            2,
            "Từng được thông báo MCV thấp/hồng cầu nhỏ",
        ),
        "q8": (
            1,
            "Từng được thông báo MCH thấp/hồng cầu nhược sắc",
        ),
        "q10": (
            2,
            "Từng được chẩn đoán HbE/hemoglobinopathy khác",
        ),
        "q11": (
            1,
            "Bản thân từng truyền máu nhiều lần/định kỳ",
        ),
        "q12": (
            2,
            "Thiếu máu kéo dài từ nhỏ/tuổi thiếu niên",
        ),
        "q13": (
            1,
            "Có mệt mỏi/giảm sức hoạt động",
        ),
        "q14": (
            1,
            "Có hoa mắt/chóng mặt không rõ nguyên nhân",
        ),
        "q15": (
            1,
            "Da/niêm nhợt",
        ),
        "q16": (
            1,
            "Vàng da/vàng mắt không rõ nguyên nhân",
        ),
        "q17": (
            2,
            "Từng được ghi nhận lách to/gan lách to",
        ),
        "q18": (
            1,
            "Có tiền sử/biến chứng bệnh huyết học mạn",
        ),
    }

    for key, (
        weight,
        label,
    ) in items.items():

        if answers.get(
            key
        ) == "Có":

            score += weight
            reasons.append(
                label
            )

    q9 = answers.get(
        "q9"
    )

    if q9 == "Đã nghi ngờ":

        score += 2

        reasons.append(
            "Từng có kết quả nghi ngờ Thalassemia/hemoglobinopathy"
        )

    elif q9 == "Đã xác định mang gen":

        score += 4

        reasons.append(
            "Từng được xác định mang gen"
        )

    return (
        min(
            score,
            ROUND1_MAX_SCORE,
        ),
        reasons,
    )


def round1_category(
    score,
):
    if score >= ROUND1_HIGH_THRESHOLD:
        return (
            "CAO",
            "Có đủ yếu tố sàng lọc ban đầu để chuyển sang Vòng 2.",
        )

    if score >= 4:
        return (
            "TRUNG BÌNH",
            "Có một số yếu tố đáng lưu ý nhưng chưa đạt ngưỡng "
            "mở Vòng 2 trong prototype.",
        )

    return (
        "THẤP",
        "Chưa ghi nhận nhiều yếu tố nguy cơ qua bộ câu hỏi ban đầu.",
    )


# ============================================================
# ROUND 2 CBC
# ============================================================

def calculate_round2_score(
    mcv,
    mch,
    rbc,
    rdw,
):
    score = 0
    reasons = []

    if mcv < 70:
        score += 3
        reasons.append(
            "MCV rất thấp (<70 fL)"
        )
    elif mcv < 75:
        score += 2
        reasons.append(
            "MCV giảm rõ (70–74,9 fL)"
        )
    elif mcv < 80:
        score += 1
        reasons.append(
            "MCV giảm (75–79,9 fL)"
        )

    if mch < 24:
        score += 2
        reasons.append(
            "MCH thấp (<24 pg)"
        )
    elif mch < 27:
        score += 1
        reasons.append(
            "MCH giảm (24–26,9 pg)"
        )

    if mcv < 80:

        if rbc >= 5.5:
            score += 2
            reasons.append(
                "RBC tương đối cao khi MCV thấp"
            )

        elif rbc >= 5.0:
            score += 1
            reasons.append(
                "RBC tương đối cao khi MCV thấp"
            )

    mentzer = (
        mcv / rbc
        if rbc > 0
        else None
    )

    if (
        mentzer is not None
        and mcv < 80
    ):

        if mentzer < 13:
            score += 2
            reasons.append(
                "Mentzer Index <13"
            )

        elif mentzer < 14:
            score += 1
            reasons.append(
                "Mentzer Index 13–13,9"
            )

    if rdw > 15:
        reasons.append(
            "RDW tăng — cần lưu ý thiếu sắt/nguồn microcytosis khác"
        )

    return (
        score,
        mentzer,
        reasons,
    )


def round2_category(
    score,
    mcv,
):
    if (
        mcv >= 80
        and score <= 2
    ):
        return (
            "THẤP",
            "CBC hiện tại chưa cho thấy mẫu hình hồng cầu nhỏ rõ.",
        )

    if score <= 3:
        return (
            "THẤP",
            "Nguy cơ sàng lọc từ CBC hiện tại thấp; không loại trừ "
            "hoàn toàn Thalassemia.",
        )

    if score <= 6:
        return (
            "TRUNG BÌNH",
            "Có đặc điểm hồng cầu nhỏ/nhược sắc. Nên đánh giá "
            "tình trạng thiếu sắt và các nguyên nhân khác.",
        )

    if score <= 9:
        return (
            "CAO",
            "Mẫu hình CBC gợi ý cần đánh giá hemoglobinopathy "
            "bằng HPLC/điện di Hb.",
        )

    return (
        "RẤT CAO",
        "Có nhiều dấu hiệu sàng lọc đáng chú ý; cần xét nghiệm xác nhận.",
    )


# ============================================================
# PLACES
# ============================================================

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
            "places.userRatingCount,"
            "places.primaryType"
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

        r = requests.post(
            url,
            headers=headers,
            json=payload,
            timeout=15,
        )

        r.raise_for_status()

        return r.json().get(
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
    origin_lat,
    origin_lng,
):
    rows = []

    for place in places:

        loc = place.get(
            "location",
            {},
        )

        lat = loc.get(
            "latitude"
        )
        lng = loc.get(
            "longitude"
        )

        if (
            lat is None
            or lng is None
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
                    origin_lat,
                    origin_lng,
                    lat,
                    lng,
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
        key=lambda x: x[
            "distance"
        ]
    )

    return rows[:5]


# ============================================================
# LOW RISK PANEL
# ============================================================

def render_low_risk_panel():

    st.success(
        "🟢 **NGUY CƠ SÀNG LỌC BAN ĐẦU: THẤP**"
    )

    st.markdown(
        """
### 📅 Theo dõi sức khỏe

Kết quả Vòng 1 hiện chưa cho thấy nhiều yếu tố nguy cơ rõ ràng.
Bạn nên tiếp tục **theo dõi sức khỏe và khám định kỳ theo hướng dẫn
của cơ sở y tế**.

Trong prototype, hệ thống đặt một mốc nhắc xem xét lại khoảng **30 ngày**.
Đây là chức năng theo dõi của ứng dụng, **không phải chỉ định bắt buộc
mọi người nguy cơ thấp phải khám hàng tháng**.

### 🩸 Khi nào nên nhập CBC lại?

Nếu sau này bạn có:
- mệt mỏi kéo dài;
- da/niêm nhợt;
- hoa mắt/chóng mặt;
- vàng da/vàng mắt;
- hoặc phiếu CBC có bất thường,

hãy mang kết quả đến nhân viên y tế và có thể nhập các chỉ số vào ô
**“Kiểm tra CBC bất thường”** bên dưới để hệ thống sàng lọc lại.
"""
    )

    next_date = (
        date.today()
        + timedelta(
            days=FOLLOWUP_DAYS
        )
    )

    st.info(
        f"🗓️ **Mốc nhắc prototype:** "
        f"{next_date.strftime('%d/%m/%Y')}"
    )

    with st.expander(
        "🩸 Có CBC bất thường? Nhập vào để hệ thống xem xét",
        expanded=False,
    ):

        c1, c2, c3, c4 = st.columns(4)

        with c1:
            hb = st.number_input(
                "Hb (g/dL)",
                3.0,
                25.0,
                13.0,
                0.1,
                key="low_cbc_hb",
            )

        with c2:
            mcv = st.number_input(
                "MCV (fL)",
                30.0,
                150.0,
                85.0,
                0.1,
                key="low_cbc_mcv",
            )

        with c3:
            mch = st.number_input(
                "MCH (pg)",
                10.0,
                50.0,
                29.0,
                0.1,
                key="low_cbc_mch",
            )

        with c4:
            rbc = st.number_input(
                "RBC (T/L)",
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
            "🔎 XEM XÉT CBC",
            key="review_low_cbc",
        ):

            if rbc <= 0:

                st.error(
                    "RBC phải lớn hơn 0."
                )

            else:

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
                    )
                )

                st.metric(
                    "Mentzer Index",
                    f"{mentzer:.2f}",
                )

                if category == "THẤP":
                    st.success(
                        f"🟢 **CBC: {category}** — {conclusion}"
                    )
                elif category == "TRUNG BÌNH":
                    st.warning(
                        f"🟡 **CBC: {category}** — {conclusion}"
                    )
                else:
                    st.error(
                        f"🟠 **CBC: {category}** — {conclusion}"
                    )

                if reasons:
                    for reason in reasons:
                        st.write(
                            f"• {reason}"
                        )


# ============================================================
# WORD
# ============================================================

def create_word(
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
        "Công cụ hỗ trợ sàng lọc. Không thay thế chẩn đoán "
        "hoặc chỉ định của nhân viên y tế."
    )

    doc.add_heading(
        "I. THÔNG TIN BỆNH NHÂN",
        level=2,
    )

    items = [
        ("Họ và tên", patient["full_name"]),
        ("Ngày sinh", patient["birth_date"]),
        ("Tuổi", str(patient["age"])),
        ("Giới tính", patient["gender"]),
        ("Số điện thoại", patient["phone"]),
        ("Địa chỉ hiện tại", patient["current_address"]),
        ("Tỉnh/thành phố", patient["province"]),
        ("Phường/xã/đặc khu", patient["commune"]),
    ]

    for label, value in items:
        doc.add_paragraph(
            f"{label}: {value}"
        )

    doc.add_heading(
        "II. VÒNG 1 — 20 CÂU HỎI",
        level=2,
    )

    doc.add_paragraph(
        f"Điểm: {r1_score}/{ROUND1_MAX_SCORE}"
    )

    doc.add_paragraph(
        f"Mức nguy cơ: {r1_category}"
    )

    if r1_reasons:
        for reason in r1_reasons:
            doc.add_paragraph(
                f"- {reason}"
            )

    if r2:

        doc.add_heading(
            "III. VÒNG 2 — CBC + ĐỘ CAO",
            level=2,
        )

        doc.add_paragraph(
            f"Địa điểm Google: {r2['geo_address']}"
        )

        if r2["elevation"] is not None:
            doc.add_paragraph(
                f"Độ cao: {r2['elevation']:.0f} m"
            )

        doc.add_paragraph(
            f"Hb thực đo: {r2['hb']:.1f} g/dL"
        )

        doc.add_paragraph(
            f"Hiệu chỉnh Hb: -{r2['adjustment']:.2f} g/dL"
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

        doc.add_heading(
            "IV. KẾT QUẢ VÒNG 2",
            level=2,
        )

        doc.add_paragraph(
            f"Điểm CBC prototype: {r2['score']}"
        )

        doc.add_paragraph(
            f"Mức nguy cơ: {r2['category']}"
        )

        doc.add_paragraph(
            f"Nhận định: {r2['conclusion']}"
        )

    else:

        doc.add_heading(
            "III. THEO DÕI",
            level=2,
        )

        doc.add_paragraph(
            "Vòng 2 chưa được mở. Tiếp tục theo dõi sức khỏe "
            "và đánh giá lại nếu xuất hiện bất thường/CBC bất thường."
        )

    output = io.BytesIO()

    doc.save(output)

    output.seek(0)

    return output


# ============================================================
# HEADER
# ============================================================

st.title(
    "🩸 HỆ THỐNG SÀNG LỌC VÀ PHÂN TẦNG NGUY CƠ THALASSEMIA"
)

st.write(
    "Hồ sơ bệnh nhân → 20 câu hỏi Vòng 1 → chỉ nguy cơ cao mới mở "
    "Vòng 2 → chọn địa điểm → độ cao + CBC → đánh giá → gợi ý cơ sở y tế."
)

with st.expander(
    "ℹ️ Nguyên tắc hệ thống",
    expanded=False,
):

    st.write(
        "Hệ thống được thiết kế như một công cụ hỗ trợ sàng lọc "
        "và điều hướng, không phải công cụ chẩn đoán."
    )

    st.write(
        "Vòng 1 tập trung vào tiền sử/nguy cơ. Vòng 2 mới sử dụng CBC "
        "và thông tin độ cao nơi cư trú."
    )


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header(
        "⚙️ Trạng thái"
    )

    if GOOGLE_API_KEY:
        st.success(
            "Google Maps API: đã cấu hình"
        )
    else:
        st.warning(
            "Google Maps API: chưa cấu hình"
        )

    if ADMIN_DATA_OK:
        st.success(
            "Địa giới: tải đầy đủ dữ liệu"
        )
    else:
        st.warning(
            "Địa giới: đang dùng dữ liệu fallback"
        )

    st.divider()

    st.write(
        "👤 Hồ sơ bệnh nhân\n"
        "↓\n"
        "🟦 Vòng 1 — 20 câu\n"
        "↓\n"
        "🔴 Nguy cơ cao?\n"
        "↓\n"
        "🟧 Vòng 2 — chọn xã/phường\n"
        "↓\n"
        "⛰️ Độ cao + CBC\n"
        "↓\n"
        "🏥 3–5 cơ sở gần"
    )


# ============================================================
# PATIENT PROFILE
# ============================================================

st.header(
    "👤 THÔNG TIN BỆNH NHÂN"
)

st.write(
    "Vui lòng nhập đầy đủ hồ sơ trước khi bắt đầu Vòng 1."
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
            help=(
                "Mỗi số điện thoại chỉ có một hồ sơ. "
                "Nhập lại cùng số sẽ ghi nhận thông tin mới nhất."
            ),
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
            placeholder="Số nhà/thôn/tổ/đường...",
        )

    st.subheader(
        "📍 Địa giới hành chính hiện tại"
    )

    st.caption(
        "Bệnh nhân chọn trực tiếp từ bảng. "
        "Hệ thống sử dụng xã/phường đã chọn để lấy tọa độ và độ cao."
    )

    province_options = provinces()

    selected_province = st.selectbox(
        "Tỉnh / Thành phố *",
        province_options,
        key="profile_province_select",
    )

    commune_options = communes_for(
        selected_province
    )

    if commune_options:

        selected_commune = st.selectbox(
            "Phường / Xã / Đặc khu *",
            [
                "— Chọn phường/xã —"
            ]
            + commune_options,
            key="profile_commune_select",
        )

        commune_value = (
            ""
            if selected_commune
            == "— Chọn phường/xã —"
            else selected_commune
        )

    else:

        commune_value = ""

        st.error(
            "Không có danh sách phường/xã của tỉnh này. "
            "Hãy kiểm tra dữ liệu địa giới."
        )

    if commune_value:
        st.success(
            f"📍 Đã chọn: **{commune_value}, {selected_province}**"
        )


phone = normalize_phone(
    phone_raw
)

if (
    phone
    and valid_vietnam_phone(phone)
    and phone_exists(phone)
):

    st.warning(
        "📌 Số điện thoại này đã tồn tại. "
        "Khi lưu, **lần nhập sau cùng sẽ ghi đè thông tin cũ**."
    )


if st.button(
    "💾 LƯU / CẬP NHẬT HỒ SƠ",
    type="secondary",
):

    if not full_name.strip():

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
        }

        action = upsert_patient(
            profile
        )

        profile["age"] = calculate_age(
            birth_date
        )

        st.session_state[
            "patient_profile"
        ] = profile

        clear_survey_results()

        if action == "updated":

            st.success(
                "✅ Đã cập nhật hồ sơ bằng **lần nhập sau cùng**. "
                "Số điện thoại vẫn chỉ có một hồ sơ."
            )

        else:

            st.success(
                "✅ Đã tạo hồ sơ bệnh nhân."
            )


patient = st.session_state.get(
    "patient_profile"
)

if not patient:

    st.info(
        "👆 Lưu hồ sơ bệnh nhân để bắt đầu Vòng 1."
    )

    st.stop()


# ============================================================
# PROFILE CHIP
# ============================================================

st.success(
    f"✅ **{patient['full_name']}** · "
    f"{patient['age']} tuổi · "
    f"{patient['phone']} · "
    f"{patient['commune']}, {patient['province']}"
)


# ============================================================
# ROUND 1
# ============================================================

st.divider()

st.header(
    "🟦 VÒNG 1 — 20 CÂU HỎI SÀNG LỌC BAN ĐẦU"
)

st.info(
    "Vòng 1 chưa yêu cầu CBC. Q19–Q20 chỉ đánh giá khả năng tiếp cận "
    "xét nghiệm và **không được cộng vào điểm nguy cơ**."
)


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
        "5. Gia đình có trẻ từng phải truyền máu nhiều lần hoặc truyền máu định kỳ không?",
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
        "10. Bạn từng được chẩn đoán HbE hoặc một hemoglobinopathy khác chưa?",
        ["Không", "Có", "Không biết"],
        horizontal=True,
    )

    q11 = st.radio(
        "11. Bản thân bạn từng truyền máu nhiều lần hoặc truyền máu định kỳ chưa?",
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
        "D. Khả năng tiếp cận xét nghiệm"
    )

    q19 = st.radio(
        "19. Bạn hiện có kết quả CBC trong vòng 6–12 tháng gần đây không?",
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
        "ℹ️ **Q19 và Q20 = 0 điểm.** "
        "Hai câu này không phải yếu tố nguy cơ Thalassemia."
    )


if st.button(
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

    score, reasons = (
        calculate_round1_score(
            answers
        )
    )

    category, conclusion = (
        round1_category(
            score
        )
    )

    st.session_state[
        "round1_score"
    ] = score

    st.session_state[
        "round1_reasons"
    ] = reasons

    st.session_state[
        "round1_category"
    ] = category

    st.session_state[
        "round1_conclusion"
    ] = conclusion

    st.session_state[
        "round1_completed"
    ] = True

    # Reset all V2 data.
    for key in list(
        st.session_state.keys()
    ):
        if key.startswith(
            "round2_"
        ) or key.startswith(
            "location_"
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

    score = st.session_state[
        "round1_score"
    ]

    category = st.session_state[
        "round1_category"
    ]

    conclusion = st.session_state[
        "round1_conclusion"
    ]

    reasons = st.session_state[
        "round1_reasons"
    ]

    c1, c2 = st.columns(2)

    with c1:
        st.metric(
            "Điểm Vòng 1",
            f"{score}/{ROUND1_MAX_SCORE}",
        )

    with c2:
        st.metric(
            "Ngưỡng mở Vòng 2",
            f"≥ {ROUND1_HIGH_THRESHOLD}",
        )

    if category == "CAO":

        st.error(
            f"🔴 **NGUY CƠ: CAO**\n\n"
            f"{conclusion}"
        )

        if reasons:
            with st.expander(
                "🔎 Các yếu tố đáng chú ý",
                expanded=True,
            ):
                for item in reasons:
                    st.write(
                        f"• {item}"
                    )

        st.session_state[
            "round2_unlocked"
        ] = True

        st.success(
            "✅ Vòng 2 đã được mở."
        )

    elif category == "TRUNG BÌNH":

        st.warning(
            f"🟡 **NGUY CƠ: TRUNG BÌNH**\n\n"
            f"{conclusion}"
        )

        if reasons:
            with st.expander(
                "🔎 Các yếu tố đáng chú ý",
                expanded=False,
            ):
                for item in reasons:
                    st.write(
                        f"• {item}"
                    )

        st.info(
            "Prototype chưa tự động mở Vòng 2 ở mức trung bình. "
            "Nhân viên y tế có thể cân nhắc đánh giá thêm."
        )

    else:

        render_low_risk_panel()


# ============================================================
# ROUND 2
# ============================================================

if st.session_state.get(
    "round2_unlocked",
    False,
):

    st.divider()

    st.header(
        "🟧 VÒNG 2 — ĐỊA ĐIỂM + ĐỘ CAO + CBC"
    )

    # --------------------------------------------------------
    # 1 LOCATION
    # --------------------------------------------------------

    with st.container(border=True):

        st.subheader(
            "1. Xác nhận vị trí cư trú"
        )

        st.write(
            f"**Tỉnh/thành:** {patient['province']}"
        )

        st.write(
            f"**Phường/xã/đặc khu:** {patient['commune']}"
        )

        st.write(
            f"**Địa chỉ chi tiết:** {patient['current_address']}"
        )

        location_query = ", ".join(
            part
            for part in [
                patient["current_address"],
                patient["commune"],
                patient["province"],
                "Việt Nam",
            ]
            if part
        )

        if st.button(
            "📍 XÁC ĐỊNH TỌA ĐỘ + ĐỘ CAO",
            type="secondary",
        ):

            if not GOOGLE_API_KEY:

                st.error(
                    "Chưa cấu hình Google Maps API key."
                )

            else:

                geo = google_geocode(
                    location_query,
                    GOOGLE_API_KEY,
                )

                if not geo:

                    st.error(
                        "Không xác định được vị trí từ địa chỉ đã chọn. "
                        "Hãy kiểm tra lại địa chỉ chi tiết."
                    )

                else:

                    elevation_result = (
                        google_elevation(
                            geo["lat"],
                            geo["lng"],
                            GOOGLE_API_KEY,
                        )
                    )

                    st.session_state[
                        "location_geo"
                    ] = geo

                    if elevation_result:

                        st.session_state[
                            "location_elevation"
                        ] = elevation_result[
                            "elevation"
                        ]

                        st.session_state[
                            "location_resolution"
                        ] = elevation_result[
                            "resolution"
                        ]

                    else:

                        st.session_state[
                            "location_elevation"
                        ] = None

        geo = st.session_state.get(
            "location_geo"
        )

        elevation = st.session_state.get(
            "location_elevation"
        )

        resolution = st.session_state.get(
            "location_resolution"
        )

        if geo:

            st.success(
                f"📍 **Google xác định:** "
                f"{geo['formatted_address']}"
            )

            g1, g2, g3 = st.columns(3)

            with g1:
                st.metric(
                    "Latitude",
                    f"{geo['lat']:.6f}",
                )

            with g2:
                st.metric(
                    "Longitude",
                    f"{geo['lng']:.6f}",
                )

            with g3:

                if elevation is not None:

                    st.metric(
                        "Độ cao",
                        f"{elevation:.0f} m",
                    )

                else:

                    st.metric(
                        "Độ cao",
                        "Chưa có",
                    )

            if elevation is not None:

                adjustment = (
                    hb_altitude_adjustment(
                        elevation
                    )
                )

                st.info(
                    f"⛰️ **Độ cao:** "
                    f"{elevation:.0f} m\n\n"
                    f"📊 **Phân tầng:** "
                    f"{altitude_band(elevation)}\n\n"
                    f"📐 **Hiệu chỉnh Hb dự kiến:** "
                    f"-{adjustment:.2f} g/dL"
                )

                if resolution:

                    st.caption(
                        f"Độ phân giải dữ liệu elevation: "
                        f"{resolution:.0f} m."
                    )

                if elevation >= 2500:

                    st.warning(
                        "Độ cao ≥2.500 m: cần thận trọng khi áp dụng "
                        "hiệu chỉnh Hb trong nghiên cứu chính thức."
                    )

    # --------------------------------------------------------
    # 2 CBC
    # --------------------------------------------------------

    with st.container(border=True):

        st.subheader(
            "2. Công thức máu (CBC)"
        )

        c1, c2, c3, c4, c5 = st.columns(5)

        with c1:
            hb = st.number_input(
                "Hb (g/dL)",
                3.0,
                25.0,
                13.0,
                0.1,
                key="round2_hb_input",
            )

        with c2:
            mcv = st.number_input(
                "MCV (fL)",
                30.0,
                150.0,
                85.0,
                0.1,
                key="round2_mcv_input",
            )

        with c3:
            mch = st.number_input(
                "MCH (pg)",
                10.0,
                50.0,
                29.0,
                0.1,
                key="round2_mch_input",
            )

        with c4:
            rbc = st.number_input(
                "RBC (T/L)",
                1.0,
                10.0,
                4.8,
                0.1,
                key="round2_rbc_input",
            )

        with c5:
            rdw = st.number_input(
                "RDW-CV (%)",
                5.0,
                40.0,
                13.0,
                0.1,
                key="round2_rdw_input",
            )

    # --------------------------------------------------------
    # 3 ANALYZE
    # --------------------------------------------------------

    if st.button(
        "🩸 PHÂN TÍCH VÒNG 2",
        type="primary",
        use_container_width=True,
    ):

        if rbc <= 0:

            st.error(
                "RBC phải lớn hơn 0."
            )

        elif st.session_state.get(
            "location_geo"
        ) is None:

            st.error(
                "Hãy xác định tọa độ + độ cao ở mục 1 trước."
            )

        else:

            elevation = st.session_state.get(
                "location_elevation"
            )

            if elevation is None:

                adjustment = 0.0
                hb_adjusted = hb

            else:

                adjustment = (
                    hb_altitude_adjustment(
                        elevation
                    )
                )

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
                "round2_adjustment"
            ] = adjustment

            st.session_state[
                "round2_mcv"
            ] = mcv

            st.session_state[
                "round2_mch"
            ] = mch

            st.session_state[
                "round2_rbc"
            ] = rbc

            st.session_state[
                "round2_rdw"
            ] = rdw

            st.session_state[
                "round2_geo_address"
            ] = (
                st.session_state[
                    "location_geo"
                ][
                    "formatted_address"
                ]
            )

            st.session_state[
                "round2_completed"
            ] = True

    # --------------------------------------------------------
    # 4 RESULT
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
                "Hb hiệu chỉnh",
                f"{st.session_state['round2_hb_adjusted']:.1f} g/dL",
            )

        category2 = st.session_state[
            "round2_category"
        ]

        conclusion2 = st.session_state[
            "round2_conclusion"
        ]

        reasons2 = st.session_state[
            "round2_reasons"
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

        elif category2 == "CAO":

            st.error(
                f"🟠 **NGUY CƠ VÒNG 2: CAO**\n\n"
                f"{conclusion2}"
            )

        else:

            st.error(
                f"🔴 **NGUY CƠ VÒNG 2: RẤT CAO**\n\n"
                f"{conclusion2}"
            )

        if reasons2:

            with st.expander(
                "🔎 Các yếu tố từ CBC",
                expanded=True,
            ):
                for reason in reasons2:
                    st.write(
                        f"• {reason}"
                    )

        st.subheader(
            "🧭 Khuyến nghị"
        )

        if category2 == "THẤP":

            st.info(
                "Chưa ghi nhận mẫu hình CBC mạnh gợi ý Thalassemia. "
                "Nếu vẫn có tiền sử đặc biệt hoặc chỉ định lâm sàng, "
                "cần nhân viên y tế đánh giá thêm."
            )

        elif category2 == "TRUNG BÌNH":

            st.warning(
                "Ưu tiên đánh giá thiếu sắt và các nguyên nhân khác gây "
                "hồng cầu nhỏ. Có thể cân nhắc Ferritin; nếu vẫn nghi ngờ, "
                "HPLC/điện di hemoglobin."
            )

        else:

            st.error(
                "Nên được đánh giá tại cơ sở có khả năng thực hiện "
                "HPLC/điện di hemoglobin; xét nghiệm phân tử có thể được "
                "cân nhắc theo chỉ định chuyên môn."
            )

        # ----------------------------------------------------
        # FACILITIES
        # ----------------------------------------------------

        st.subheader(
            "🏥 CƠ SỞ Y TẾ GẦN NHẤT"
        )

        geo = st.session_state.get(
            "location_geo"
        )

        if geo is not None and GOOGLE_API_KEY:

            facilities_raw = nearby_medical(
                geo["lat"],
                geo["lng"],
                GOOGLE_API_KEY,
            )

            facilities = rank_facilities(
                facilities_raw,
                geo["lat"],
                geo["lng"],
            )

            if not facilities:

                st.info(
                    "Không tìm thấy cơ sở trong bán kính 50 km "
                    "theo truy vấn hiện tại."
                )

            else:

                st.write(
                    "Hệ thống ưu tiên tối đa **5 cơ sở gần nhất**. "
                    "Khoảng cách bên dưới là khoảng cách địa lý ước tính."
                )

                for index, facility in enumerate(
                    facilities,
                    start=1,
                ):

                    with st.container(
                        border=True
                    ):

                        st.markdown(
                            f"### {index}. {facility['name']}"
                        )

                        st.write(
                            f"📍 {facility['address']}"
                        )

                        st.write(
                            f"📏 **{facility['distance']:.1f} km**"
                        )

                        if facility[
                            "rating"
                        ] is not None:

                            rating = (
                                f"⭐ {facility['rating']:.1f}"
                            )

                            if facility[
                                "rating_count"
                            ]:

                                rating += (
                                    f" "
                                    f"({facility['rating_count']:,} đánh giá)"
                                )

                            st.write(
                                rating
                            )

                        if facility[
                            "maps_uri"
                        ]:

                            st.link_button(
                                "🗺️ Mở Google Maps",
                                facility[
                                    "maps_uri"
                                ],
                            )

                        st.caption(
                            "Cần xác nhận trước khả năng thực hiện "
                            "Ferritin/HPLC/điện di Hb/xét nghiệm gen."
                        )

        else:

            st.info(
                "Chưa có dữ liệu Google Maps để tìm cơ sở gần."
            )

        # ----------------------------------------------------
        # WORD
        # ----------------------------------------------------

        st.subheader(
            "📄 PHIẾU KẾT QUẢ"
        )

        report = create_word(
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
                "geo_address": st.session_state[
                    "round2_geo_address"
                ],
                "elevation": st.session_state.get(
                    "location_elevation"
                ),
                "hb": st.session_state[
                    "round2_hb"
                ],
                "adjustment": st.session_state[
                    "round2_adjustment"
                ],
                "hb_adjusted": st.session_state[
                    "round2_hb_adjusted"
                ],
                "mcv": st.session_state[
                    "round2_mcv"
                ],
                "mch": st.session_state[
                    "round2_mch"
                ],
                "rbc": st.session_state[
                    "round2_rbc"
                ],
                "rdw": st.session_state[
                    "round2_rdw"
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
    "THALASSEMIA SCREENING V4.0 — PROTOTYPE NGHIÊN CỨU. "
    "Risk score chưa được validation trên quần thể người Việt Nam."
)
