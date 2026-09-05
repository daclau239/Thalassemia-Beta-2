import io
import math
import re
import sqlite3
from pathlib import Path
from datetime import date, datetime, timedelta

import requests
import streamlit as st
from docx import Document

# ============================================================
# THALASSEMIA SCREENING V3.3
# ============================================================
# FLOW:
#   HỒ SƠ BỆNH NHÂN
#       ↓
#   VÒNG 1 — 20 câu hỏi
#       ↓
#   ┌───────────────┬────────────────┐
#   ↓               ↓                ↓
# THẤP          TRUNG BÌNH           CAO
#   ↓               ↓                ↓
# Theo dõi        Theo dõi /       MỞ VÒNG 2
# định kỳ         nhân viên y tế       ↓
# + nếu CBC      có thể cân nhắc     Xã/phường
# bất thường     CBC                  ↓
# nhập CBC                           Google Geocoding
#                                       ↓
#                                   Elevation
#                                       ↓
#                                      CBC
#                                       ↓
#                                Hb hiệu chỉnh + Mentzer
#                                       ↓
#                               Gợi ý 3–5 cơ sở y tế
#
# HỒ SƠ:
#   - Số điện thoại là khóa duy nhất.
#   - Nếu nhập lại cùng số, thông tin mới nhất ghi đè hồ sơ cũ.
#
# LƯU Ý:
#   - Điểm nguy cơ là PROTOTYPE, chưa validation trên quần thể Việt Nam.
#   - Không dùng kết quả để chẩn đoán hoặc tự điều trị.
#   - Câu 19–20 về khả năng tiếp cận xét nghiệm KHÔNG cộng điểm nguy cơ.
#   - Khuyến nghị "theo dõi hàng tháng" trong giao diện chỉ là lịch nhắc
#     prototype, không phải chỉ định y khoa bắt buộc cho mọi người nguy cơ thấp.
#   - Công thức hiệu chỉnh Hb theo độ cao cần được xác nhận theo guideline
#     và thiết kế nghiên cứu chính thức trước khi triển khai lâm sàng.
#
# CÀI:
#   pip install streamlit python-docx requests
#
# .streamlit/secrets.toml:
# [google]
# maps_api_key = "YOUR_GOOGLE_MAPS_API_KEY"
#
# GOOGLE CLOUD:
#   Geocoding API
#   Elevation API
#   Places API (New)
# ============================================================

st.set_page_config(
    page_title="Hệ thống Sàng lọc Thalassemia",
    page_icon="🩸",
    layout="wide",
)

DB_PATH = "thalassemia_patients.db"
ROUND1_MAX_SCORE = 20
ROUND1_HIGH_THRESHOLD = 8
FOLLOWUP_DAYS = 30


# ============================================================
# GOOGLE
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
# DATABASE — PHONE UNIQUE + LAST ENTRY WINS
# ============================================================

def db():
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
            district TEXT,
            commune TEXT NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )
    conn.commit()
    return conn


def normalize_phone(value):
    digits = re.sub(r"\D", "", value or "")
    if digits.startswith("84") and len(digits) == 11:
        digits = "0" + digits[2:]
    return digits


def valid_vn_phone(phone):
    return bool(re.fullmatch(r"0\d{9}", phone))


def patient_exists(phone):
    conn = db()
    row = conn.execute(
        "SELECT 1 FROM patient_profiles WHERE phone = ?",
        (phone,),
    ).fetchone()
    conn.close()
    return row is not None


def upsert_patient(profile):
    conn = db()
    now = datetime.now().isoformat(timespec="seconds")

    exists = patient_exists(profile["phone"])

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
                district = ?,
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
                profile["district"],
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
                district,
                commune,
                created_at,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                profile["phone"],
                profile["full_name"],
                profile["birth_date"],
                profile["gender"],
                profile["current_address"],
                profile["province"],
                profile["district"],
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
# UTILITY
# ============================================================

def safe_filename(text):
    text = (text or "").strip()
    text = re.sub(
        r"[^0-9A-Za-zÀ-ỹĐđ _-]",
        "_",
        text,
    )
    return text.strip(" _") or "nguoi_sang_loc"


def calculate_age(birth_date):
    if birth_date is None:
        return None

    today = date.today()
    return (
        today.year
        - birth_date.year
        - (
            (today.month, today.day)
            < (birth_date.month, birth_date.day)
        )
    )


def reset_round2():
    for key in list(st.session_state.keys()):
        if (
            key.startswith("round2_")
            or key.startswith("location_")
            or key.startswith("medical_")
        ):
            del st.session_state[key]


# ============================================================
# VÒNG 1 — 20 QUESTIONS
# ============================================================

def round1_score(answers):
    """
    PROTOTYPE.
    Q1-Q18 = yếu tố nguy cơ.
    Q19-Q20 = tiếp cận y tế, KHÔNG cộng điểm.
    """

    score = 0
    reasons = []

    weights = {
        "q1": 3,
        "q2": 3,
        "q3": 1,
        "q4": 2,
        "q5": 2,
        "q6": 1,
        "q7": 2,
        "q8": 1,
        "q9_suspected": 2,
        "q9_carrier": 4,
        "q10": 2,
        "q11": 1,
        "q12": 2,
        "q13": 1,
        "q14": 1,
        "q15": 1,
        "q16": 1,
        "q17": 2,
        "q18": 1,
    }

    labels = {
        "q1": "Có người thân/dòng họ mắc Thalassemia",
        "q2": "Có người thân/dòng họ mang gen Thalassemia/hemoglobinopathy",
        "q3": "Cha/mẹ từng xét nghiệm Thalassemia/hemoglobinopathy",
        "q4": "Anh/chị/em từng thiếu máu hoặc hồng cầu nhỏ",
        "q5": "Gia đình có trẻ từng truyền máu nhiều lần/định kỳ",
        "q6": "Từng được thông báo thiếu máu",
        "q7": "Từng được thông báo MCV thấp/hồng cầu nhỏ",
        "q8": "Từng được thông báo MCH thấp/hồng cầu nhược sắc",
        "q10": "Từng được chẩn đoán HbE/hemoglobinopathy khác",
        "q11": "Bản thân từng truyền máu nhiều lần/định kỳ",
        "q12": "Thiếu máu kéo dài từ nhỏ/tuổi thiếu niên",
        "q13": "Mệt mỏi/giảm sức hoạt động",
        "q14": "Hoa mắt/chóng mặt không rõ nguyên nhân",
        "q15": "Da/niêm nhợt",
        "q16": "Vàng da/vàng mắt không rõ nguyên nhân",
        "q17": "Từng được ghi nhận lách to/gan lách to",
        "q18": "Biến chứng liên quan bệnh huyết học mạn",
    }

    for key in [
        "q1", "q2", "q3", "q4", "q5",
        "q6", "q7", "q8", "q10",
        "q11", "q12", "q13", "q14",
        "q15", "q16", "q17", "q18",
    ]:
        if answers.get(key) == "Có":
            score += weights[key]
            reasons.append(labels[key])

    if answers.get("q9") == "Đã nghi ngờ":
        score += weights["q9_suspected"]
        reasons.append(
            "Từng có kết quả nghi ngờ Thalassemia/hemoglobinopathy"
        )
    elif answers.get("q9") == "Đã xác định mang gen":
        score += weights["q9_carrier"]
        reasons.append(
            "Từng được xác định mang gen"
        )

    return min(score, ROUND1_MAX_SCORE), reasons


def round1_category(score):
    if score >= ROUND1_HIGH_THRESHOLD:
        return (
            "CAO",
            "Có đủ yếu tố sàng lọc ban đầu để chuyển sang Vòng 2.",
        )
    if score >= 4:
        return (
            "TRUNG BÌNH",
            "Có một số yếu tố đáng lưu ý nhưng chưa đạt ngưỡng mở Vòng 2 "
            "trong prototype hiện tại.",
        )
    return (
        "THẤP",
        "Chưa ghi nhận nhiều yếu tố nguy cơ qua bộ câu hỏi ban đầu.",
    )


# ============================================================
# GOOGLE GEOCODING
# ============================================================

@st.cache_data(ttl=86400)
def google_geocode(address, api_key):
    if not api_key or not address.strip():
        return None

    url = "https://maps.googleapis.com/maps/api/geocode/json"
    params = {
        "address": address,
        "key": api_key,
        "language": "vi",
        "region": "vn",
    }

    try:
        r = requests.get(
            url,
            params=params,
            timeout=12,
        )
        r.raise_for_status()
        data = r.json()

        if (
            data.get("status") != "OK"
            or not data.get("results")
        ):
            return None

        result = data["results"][0]
        location = result["geometry"]["location"]

        components = {}
        for comp in result.get(
            "address_components",
            [],
        ):
            for comp_type in comp.get(
                "types",
                [],
            ):
                components.setdefault(
                    comp_type,
                    comp.get(
                        "long_name",
                        "",
                    ),
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
            "components": components,
        }

    except Exception:
        return None



# ============================================================
# GOOGLE PLACES AUTOCOMPLETE — CHỌN XÃ/PHƯỜNG
# ============================================================

@st.cache_data(ttl=3600)
def google_region_predictions(
    input_text,
    api_key,
):
    """
    Tìm các địa danh/vùng hành chính để người dùng CHỌN.
    Dùng Places API (New) Autocomplete với type collection (regions).
    """
    if not api_key or not input_text.strip():
        return []

    url = "https://places.googleapis.com/v1/places:autocomplete"

    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": api_key,
        "X-Goog-FieldMask": (
            "suggestions.placePrediction.placeId,"
            "suggestions.placePrediction.text.text,"
            "suggestions.placePrediction.structuredFormat.mainText.text,"
            "suggestions.placePrediction.structuredFormat.secondaryText.text"
        ),
    }

    payload = {
        "input": input_text,
        "includedPrimaryTypes": ["(regions)"],
        "includedRegionCodes": ["vn"],
        "languageCode": "vi",
        "regionCode": "vn",
    }

    try:
        response = requests.post(
            url,
            headers=headers,
            json=payload,
            timeout=12,
        )
        response.raise_for_status()
        data = response.json()

        suggestions = []

        for item in data.get("suggestions", []):
            prediction = item.get(
                "placePrediction"
            )

            if not prediction:
                continue

            text_obj = prediction.get(
                "text",
                {}
            )

            display = text_obj.get(
                "text",
                ""
            )

            main_text = prediction.get(
                "structuredFormat",
                {},
            ).get(
                "mainText",
                {},
            ).get(
                "text",
                "",
            )

            secondary_text = prediction.get(
                "structuredFormat",
                {},
            ).get(
                "secondaryText",
                {},
            ).get(
                "text",
                "",
            )

            if display:
                suggestions.append(
                    {
                        "display": display,
                        "main": main_text or display,
                        "secondary": secondary_text,
                        "place_id": prediction.get(
                            "placeId",
                            "",
                        ),
                    }
                )

        return suggestions[:10]

    except Exception:
        return []


def google_place_geocode(place_id, api_key):
    """
    Dùng Geocoding API với place_id đã được người dùng chọn.
    Điều này giúp lấy đúng tọa độ của lựa chọn đó.
    """
    if not api_key or not place_id:
        return None

    url = (
        "https://maps.googleapis.com/maps/api/geocode/json"
    )

    params = {
        "place_id": place_id,
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
        location = result["geometry"]["location"]

        components = {}

        for component in result.get(
            "address_components",
            [],
        ):
            for component_type in component.get(
                "types",
                [],
            ):
                components.setdefault(
                    component_type,
                    component.get(
                        "long_name",
                        "",
                    ),
                )

        return {
            "formatted_address": result.get(
                "formatted_address",
                "",
            ),
            "lat": float(
                location["lat"]
            ),
            "lng": float(
                location["lng"]
            ),
            "place_id": result.get(
                "place_id",
                place_id,
            ),
            "components": components,
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

    url = "https://maps.googleapis.com/maps/api/elevation/json"
    params = {
        "locations": f"{lat},{lng}",
        "key": api_key,
    }

    try:
        r = requests.get(
            url,
            params=params,
            timeout=12,
        )
        r.raise_for_status()
        data = r.json()

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
# WHO 2024 — Hb ALTITUDE ADJUSTMENT
# ============================================================

def hb_altitude_adjustment(elevation_m):
    if (
        elevation_m is None
        or elevation_m <= 0
    ):
        return 0.0

    adjustment_g_l = (
        0.0056384 * elevation_m
        + 0.0000003 * elevation_m**2
    )

    return adjustment_g_l / 10.0


def altitude_band(elevation_m):
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
# ROUND 2 CBC
# ============================================================

def round2_score(
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
# MEDICAL FACILITIES
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
    earth = 6371.0

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
        * earth
        * math.asin(
            math.sqrt(a)
        )
    )


def rank_facilities(
    places,
    origin_lat,
    origin_lng,
):
    result = []

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

        result.append(
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

    result.sort(
        key=lambda x: x["distance"]
    )

    return result[:5]


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

Hiện tại bộ câu hỏi chưa ghi nhận nhiều yếu tố nguy cơ rõ ràng.
Bạn nên tiếp tục theo dõi tình trạng sức khỏe và khám sức khỏe
định kỳ theo hướng dẫn của cơ sở y tế.

Trong **prototype**, hệ thống đặt một mốc nhắc xem xét lại sau
khoảng **30 ngày**. Đây là chức năng theo dõi của ứng dụng, **không
phải chỉ định bắt buộc rằng mọi người nguy cơ thấp phải khám hàng tháng**.

### 🩸 Khi nào nên kiểm tra lại?

Nếu trong thời gian theo dõi xuất hiện bất thường như mệt mỏi kéo dài,
da/niêm nhợt, chóng mặt, vàng da/vàng mắt, hoặc đã có kết quả CBC
bất thường, hãy đưa kết quả cho nhân viên y tế và có thể nhập CBC
vào phần bên dưới để hệ thống **sàng lọc lại**, thay vì tự kết luận bệnh.
"""
    )

    next_date = date.today() + timedelta(
        days=FOLLOWUP_DAYS
    )

    st.info(
        f"🗓️ **Mốc nhắc prototype:** "
        f"{next_date.strftime('%d/%m/%Y')}"
    )

    with st.expander(
        "🩸 Nhập CBC nếu lần xét nghiệm sau có bất thường",
        expanded=False,
    ):

        st.caption(
            "Phần này dành cho người Vòng 1 nguy cơ thấp nhưng sau đó "
            "có CBC bất thường. Không dùng riêng CBC để chẩn đoán."
        )

        c1, c2, c3, c4 = st.columns(4)

        with c1:
            hb = st.number_input(
                "Hb (g/dL)",
                3.0,
                25.0,
                13.0,
                0.1,
                key="low_hb",
            )

        with c2:
            mcv = st.number_input(
                "MCV (fL)",
                30.0,
                150.0,
                85.0,
                0.1,
                key="low_mcv",
            )

        with c3:
            mch = st.number_input(
                "MCH (pg)",
                10.0,
                50.0,
                29.0,
                0.1,
                key="low_mch",
            )

        with c4:
            rbc = st.number_input(
                "RBC (T/L)",
                1.0,
                10.0,
                4.8,
                0.1,
                key="low_rbc",
            )

        rdw = st.number_input(
            "RDW-CV (%)",
            5.0,
            40.0,
            13.0,
            0.1,
            key="low_rdw",
        )

        if st.button(
            "🔎 XEM XÉT CBC",
            key="low_cbc_review",
        ):
            if rbc <= 0:
                st.error(
                    "RBC phải lớn hơn 0."
                )
            else:
                score, mentzer, reasons = (
                    round2_score(
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
        "I. THÔNG TIN NGƯỜI ĐƯỢC SÀNG LỌC",
        level=2,
    )

    fields = [
        ("Họ và tên", patient["full_name"]),
        ("Ngày sinh", patient["birth_date"]),
        ("Tuổi", str(patient["age"])),
        ("Giới tính", patient["gender"]),
        ("Số điện thoại", patient["phone"]),
        ("Nơi ở hiện tại", patient["current_address"]),
        ("Tỉnh/thành phố", patient["province"]),
        ("Quận/huyện/thị xã", patient["district"] or "—"),
        ("Xã/phường/thị trấn", patient["commune"]),
    ]

    for label, value in fields:
        doc.add_paragraph(
            f"{label}: {value}"
        )

    doc.add_heading(
        "II. VÒNG 1 — BỘ 20 CÂU HỎI",
        level=2,
    )

    doc.add_paragraph(
        f"Điểm Vòng 1: {r1_score}/{ROUND1_MAX_SCORE}"
    )

    doc.add_paragraph(
        f"Mức nguy cơ: {r1_category}"
    )

    if r1_reasons:
        doc.add_paragraph(
            "Các yếu tố đáng chú ý:"
        )
        for item in r1_reasons:
            doc.add_paragraph(
                f"- {item}"
            )

    if r2 is not None:
        doc.add_heading(
            "III. VÒNG 2 — CBC + ĐỘ CAO",
            level=2,
        )

        doc.add_paragraph(
            f"Địa điểm Google xác định: {r2['geo_address']}"
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
            "Kết quả Vòng 1 chưa mở Vòng 2. "
            "Theo dõi sức khỏe và đánh giá lại nếu xuất hiện "
            "triệu chứng hoặc CBC bất thường."
        )

    doc.add_heading(
        "V. LƯU Ý",
        level=2,
    )

    doc.add_paragraph(
        "Điểm số trong prototype chưa được validation trên "
        "quần thể người Việt Nam."
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
    "Hồ sơ bệnh nhân → Vòng 1 gồm 20 câu hỏi → "
    "chỉ nguy cơ cao mới mở Vòng 2 → CBC + độ cao → "
    "đánh giá và gợi ý cơ sở y tế."
)

with st.expander(
    "ℹ️ Nguyên tắc",
    expanded=False,
):
    st.write(
        "Mục tiêu là hỗ trợ sàng lọc ban đầu tại tuyến cơ sở, "
        "đặc biệt cho người ở khu vực khó tiếp cận xét nghiệm chuyên sâu."
    )
    st.write(
        "Risk score là prototype nghiên cứu, chưa được validation "
        "và không dùng để tự chẩn đoán."
    )


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.header("⚙️ Hệ thống")

    if GOOGLE_API_KEY:
        st.success(
            "Google Maps API: đã cấu hình"
        )
    else:
        st.warning(
            "Google Maps API: chưa cấu hình"
        )

    st.divider()

    if Path(ADMIN_DATA_PATH).exists():
        st.success(
            "Địa giới: data/vietnam_admin.json"
        )
    else:
        st.warning(
            "Địa giới: đang dùng dữ liệu mẫu. "
            "Hãy bổ sung data/vietnam_admin.json trước khi triển khai."
        )

    st.divider()

    st.write(
        "👤 Hồ sơ + chọn địa giới\n"
        "↓\n"
        "🟦 Vòng 1 — 20 câu\n"
        "↓\n"
        "🔴 Nguy cơ cao?\n"
        "↓\n"
        "🟧 Vòng 2 — tọa độ + độ cao + CBC\n"
        "↓\n"
        "🏥 Cơ sở y tế"
    )



# ============================================================
# DỮ LIỆU ĐỊA GIỚI HÀNH CHÍNH
# ============================================================

ADMIN_DATA_PATH = "data/vietnam_admin.json"


@st.cache_data(ttl=86400)
def load_admin_data():
    """
    Cấu trúc JSON:
    {
      "Tỉnh/Thành phố": {
        "Quận/Huyện": ["Xã/Phường/Đặc khu", ...]
      }
    }

    Dùng file JSON riêng để dễ cập nhật theo danh mục hành chính hiện hành.
    Nếu chưa có file, app dùng dữ liệu mẫu nhỏ để chạy thử.
    """
    import json

    file_path = Path(ADMIN_DATA_PATH)

    if not file_path.exists():
        return {
            "Đà Nẵng": {
                "Hòa Vang": [
                    "Hòa Phong",
                    "Hòa Phú",
                    "Hòa Sơn",
                    "Hòa Tiến",
                ],
                "Ngũ Hành Sơn": [
                    "Mỹ An",
                    "Khuê Mỹ",
                    "Hòa Hải",
                    "Hòa Quý",
                ],
            }
        }

    try:
        with file_path.open("r", encoding="utf-8") as f:
            data = json.load(f)

        if not isinstance(data, dict):
            raise ValueError("Dữ liệu địa giới phải là JSON object.")

        return data
    except Exception as exc:
        st.warning(
            f"Không đọc được {ADMIN_DATA_PATH}: {exc}. "
            "Ứng dụng đang dùng dữ liệu mẫu."
        )
        return {
            "Đà Nẵng": {
                "Hòa Vang": [
                    "Hòa Phong",
                    "Hòa Phú",
                    "Hòa Sơn",
                    "Hòa Tiến",
                ]
            }
        }


ADMIN_DATA = load_admin_data()


def admin_provinces():
    return sorted(
        ADMIN_DATA.keys(),
        key=lambda x: x.lower(),
    )


def admin_districts(province):
    return sorted(
        ADMIN_DATA.get(province, {}).keys(),
        key=lambda x: x.lower(),
    )


def admin_communes(province, district):
    return sorted(
        ADMIN_DATA.get(province, {}).get(
            district,
            [],
        ),
        key=lambda x: x.lower(),
    )


# ============================================================
# ============================================================
# PATIENT PROFILE
# ============================================================

st.header(
    "👤 THÔNG TIN NGƯỜI ĐƯỢC SÀNG LỌC"
)

st.write(
    "Nhập hồ sơ bệnh nhân trước khi bắt đầu Vòng 1. "
    "Số điện thoại là mã hồ sơ duy nhất."
)

with st.container(border=True):

    c1, c2, c3 = st.columns(3)

    with c1:
        full_name = st.text_input(
            "Họ và tên *",
            placeholder="Nguyễn Văn A",
        )

        birth_date = st.date_input(
            "Ngày sinh *",
            value=date(2000, 1, 1),
            min_value=date(1900, 1, 1),
            max_value=date.today(),
            format="DD/MM/YYYY",
        )

    with c2:
        phone_raw = st.text_input(
            "Số điện thoại *",
            placeholder="09xxxxxxxx",
            help=(
                "Mỗi số điện thoại chỉ có một hồ sơ. "
                "Nhập lại cùng số sẽ ghi nhận lần nhập mới nhất."
            ),
        )

        gender = st.selectbox(
            "Giới tính *",
            ["Nam", "Nữ", "Khác"],
        )

    with c3:
        current_address = st.text_input(
            "Địa chỉ chi tiết hiện tại *",
            placeholder="Thôn/tổ/số nhà/đường...",
        )

    st.subheader(
        "📍 Địa giới hành chính hiện tại"
    )

    st.caption(
        "Bệnh nhân **chọn trực tiếp từ danh sách**, không cần tự gõ "
        "tên xã/phường. Danh sách được liên kết theo cấp địa giới."
    )

    provinces = admin_provinces()

    selected_province = st.selectbox(
        "Tỉnh / thành phố *",
        provinces,
        key="profile_province",
    )

    districts = admin_districts(
        selected_province
    )

    district_options = (
        ["— Chọn quận/huyện —"]
        + districts
    )

    selected_district = st.selectbox(
        "Quận / huyện / thị xã",
        district_options,
        key="profile_district",
    )

    district_value = (
        ""
        if selected_district
        == "— Chọn quận/huyện —"
        else selected_district
    )

    communes = admin_communes(
        selected_province,
        district_value,
    )

    commune_options = (
        ["— Chọn phường/xã/đặc khu —"]
        + communes
    )

    selected_commune = st.selectbox(
        "Phường / xã / đặc khu *",
        commune_options,
        key="profile_commune",
    )

    commune_value = (
        ""
        if selected_commune
        == "— Chọn phường/xã/đặc khu —"
        else selected_commune
    )

    if (
        selected_province
        and commune_value
    ):
        st.success(
            "📍 Địa điểm hành chính đã chọn: "
            f"**{commune_value}, "
            f"{district_value + ', ' if district_value else ''}"
            f"{selected_province}**"
        )


phone = normalize_phone(phone_raw)

if phone and valid_vn_phone(phone):
    if patient_exists(phone):
        st.warning(
            "📌 **Số điện thoại này đã có hồ sơ.** "
            "Khi lưu, thông tin mới sẽ ghi đè hồ sơ cũ "
            "và được ghi nhận là lần nhập sau cùng."
        )


if st.button(
    "💾 LƯU / CẬP NHẬT HỒ SƠ",
    type="secondary",
):

    if not full_name.strip():
        st.error(
            "Vui lòng nhập họ và tên."
        )

    elif not valid_vn_phone(phone):
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
            "district": district_value,
            "commune": commune_value,
        }

        action = upsert_patient(profile)

        st.session_state[
            "patient_profile"
        ] = profile

        st.session_state[
            "patient_age"
        ] = calculate_age(
            birth_date
        )

        # Lưu hồ sơ mới -> xóa kết quả sàng lọc cũ.
        for key in list(
            st.session_state.keys()
        ):
            if (
                key.startswith("round1_")
                or key.startswith("round2_")
                or key.startswith("location_")
                or key.startswith("medical_")
            ):
                del st.session_state[key]

        if action == "updated":
            st.success(
                "✅ Hồ sơ đã được cập nhật bằng **lần nhập sau cùng**. "
                "Số điện thoại vẫn chỉ có một hồ sơ."
            )
        else:
            st.success(
                "✅ Đã tạo hồ sơ bệnh nhân."
            )

patient = st.session_state.get(
    "patient_profile"
)


# ROUND 1
# ============================================================

st.divider()

st.header(
    "🟦 VÒNG 1 — 20 CÂU HỎI SÀNG LỌC BAN ĐẦU"
)

st.info(
    "Vòng 1 chưa yêu cầu CBC. Câu 19–20 chỉ nhằm đánh giá "
    "khả năng tiếp cận xét nghiệm và **không được cộng vào điểm nguy cơ**."
)

# -------------------------
# A
# -------------------------

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


# -------------------------
# B
# -------------------------

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


# -------------------------
# C
# -------------------------

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


# -------------------------
# D — ACCESS, ZERO SCORE
# -------------------------

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
        "Hai câu này không phản ánh nguy cơ mắc Thalassemia. "
        "Chúng chỉ giúp hệ thống hiểu khả năng tiếp cận y tế để hỗ trợ "
        "điều hướng ở các bước sau."
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

    score, reasons = round1_score(
        answers
    )

    category, conclusion = round1_category(
        score
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

    reset_round2()


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

    m1, m2 = st.columns(2)

    with m1:
        st.metric(
            "Điểm Vòng 1",
            f"{score}/{ROUND1_MAX_SCORE}",
        )

    with m2:
        st.metric(
            "Ngưỡng mở Vòng 2",
            f"≥ {ROUND1_HIGH_THRESHOLD}",
        )

    if category == "CAO":

        st.error(
            f"🔴 **NGUY CƠ VÒNG 1: CAO**\n\n"
            f"{conclusion}"
        )

        if reasons:
            with st.expander(
                "🔎 Các yếu tố đáng chú ý",
                expanded=True,
            ):
                for reason in reasons:
                    st.write(
                        f"• {reason}"
                    )

        st.session_state[
            "round2_unlocked"
        ] = True

        st.success(
            "✅ **Vòng 2 đã được mở.** "
            "Bây giờ mới cần thực hiện đánh giá CBC và vị trí cư trú."
        )

    elif category == "TRUNG BÌNH":

        st.warning(
            f"🟡 **NGUY CƠ VÒNG 1: TRUNG BÌNH**\n\n"
            f"{conclusion}"
        )

        if reasons:
            with st.expander(
                "🔎 Các yếu tố đáng chú ý",
                expanded=False,
            ):
                for reason in reasons:
                    st.write(
                        f"• {reason}"
                    )

        st.info(
            "Prototype chưa tự động mở Vòng 2 ở mức trung bình. "
            "Nhân viên y tế có thể cân nhắc đánh giá thêm tùy trường hợp."
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
        "🟧 VÒNG 2 — THÔNG SỐ HUYẾT HỌC + ĐỘ CAO"
    )

    st.info(
        "Vòng 2 chỉ mở sau khi Vòng 1 đạt ngưỡng nguy cơ cao."
    )

    # --------------------------------------------------------
    # PATIENT SUMMARY
    # --------------------------------------------------------

    with st.container(border=True):

        st.subheader(
            "1. Hồ sơ bệnh nhân"
        )

        info1, info2, info3 = st.columns(3)

        with info1:
            st.write(
                f"**Họ và tên:** {patient['full_name']}"
            )
            st.write(
                f"**Tuổi:** {patient_age}"
            )

        with info2:
            st.write(
                f"**Số điện thoại:** {patient['phone']}"
            )
            st.write(
                f"**Giới tính:** {patient['gender']}"
            )

        with info3:
            st.write(
                f"**Xã/phường:** {patient['commune']}"
            )
            st.write(
                f"**Tỉnh/thành:** {patient['province']}"
            )

    # --------------------------------------------------------
    # LOCATION — NGƯỜI DÙNG PHẢI CHỌN XÃ/PHƯỜNG
    # --------------------------------------------------------

    with st.container(border=True):

        st.subheader(
            "2. Chọn nơi đang sinh sống"
        )

        st.write(
            "Người dùng phải **chọn xã/phường/đặc khu từ danh sách gợi ý**. "
            "Hệ thống không tự đoán xã/phường từ một chuỗi địa chỉ mơ hồ."
        )

        l1, l2 = st.columns(2)

        with l1:
            province_input = st.text_input(
                "Tỉnh / thành phố *",
                value=patient["province"],
                key="location_province_input",
            )

            district_input = st.text_input(
                "Quận / huyện / thị xã",
                value=patient["district"],
                key="location_district_input",
            )

        with l2:
            commune_search = st.text_input(
                "Tìm xã / phường / đặc khu *",
                value=patient["commune"],
                key="commune_search_input",
                placeholder="Ví dụ: Hòa Phong",
                help=(
                    "Nhập tên hoặc một phần tên xã/phường, sau đó "
                    "bấm 'Tìm xã/phường' để chọn."
                ),
            )

            extra_address = st.text_input(
                "Thôn/tổ/đường/địa chỉ chi tiết (nếu có)",
                value=patient["current_address"],
                key="location_extra_address",
            )

        if st.button(
            "🔎 TÌM XÃ / PHƯỜNG ĐỂ CHỌN",
            key="search_commune_button",
            type="secondary",
        ):

            if not GOOGLE_API_KEY:
                st.error(
                    "Chưa cấu hình Google Maps API key."
                )

            elif not province_input.strip():
                st.error(
                    "Vui lòng nhập tỉnh/thành phố trước."
                )

            elif not commune_search.strip():
                st.error(
                    "Vui lòng nhập từ khóa xã/phường."
                )

            else:

                region_query = ", ".join(
                    part
                    for part in [
                        commune_search.strip(),
                        district_input.strip(),
                        province_input.strip(),
                        "Việt Nam",
                    ]
                    if part
                )

                predictions = google_region_predictions(
                    region_query,
                    GOOGLE_API_KEY,
                )

                if not predictions:
                    st.warning(
                        "Không tìm thấy xã/phường phù hợp. "
                        "Hãy thử tên đầy đủ hoặc bỏ bớt địa chỉ chi tiết."
                    )
                    st.session_state[
                        "location_predictions"
                    ] = []
                else:
                    st.session_state[
                        "location_predictions"
                    ] = predictions

        predictions = st.session_state.get(
            "location_predictions",
            [],
        )

        if predictions:

            st.markdown(
                "### 📍 Chọn đúng xã/phường"
            )

            option_map = {
                (
                    f"{item['main']}"
                    + (
                        f" — {item['secondary']}"
                        if item["secondary"]
                        else ""
                    )
                    + f" | {item['display']}"
                ): item
                for item in predictions
            }

            selected_label = st.selectbox(
                "Danh sách xã/phường tìm được",
                list(option_map.keys()),
                key="selected_commune_prediction",
            )

            selected_prediction = option_map[
                selected_label
            ]

            if st.button(
                "✅ XÁC NHẬN XÃ/PHƯỜNG NÀY",
                key="confirm_commune_button",
                type="primary",
            ):

                selected_geo = google_place_geocode(
                    selected_prediction["place_id"],
                    GOOGLE_API_KEY,
                )

                if not selected_geo:

                    st.error(
                        "Không lấy được tọa độ của xã/phường đã chọn."
                    )

                else:

                    st.session_state[
                        "location_geo"
                    ] = selected_geo

                    st.session_state[
                        "location_selected_commune"
                    ] = selected_prediction[
                        "display"
                    ]

                    # Lấy độ cao chính xác theo tọa độ của lựa chọn.
                    elevation_result = google_elevation(
                        selected_geo["lat"],
                        selected_geo["lng"],
                        GOOGLE_API_KEY,
                    )

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

                        st.session_state[
                            "location_resolution"
                        ] = None

                    st.success(
                        "✅ Đã xác nhận địa điểm cư trú."
                    )

        geo = st.session_state.get(
            "location_geo"
        )

        elevation = st.session_state.get(
            "location_elevation"
        )

        resolution = st.session_state.get(
            "location_resolution"
        )

        selected_commune = st.session_state.get(
            "location_selected_commune"
        )

        if selected_commune:
            st.info(
                f"**Xã/phường đã chọn:** {selected_commune}"
            )

        if geo:

            st.success(
                f"📍 **Địa điểm được xác định:** "
                f"{geo['formatted_address']}"
            )

            g1, g2, g3 = st.columns(3)

            with g1:
                st.metric(
                    "Vĩ độ",
                    f"{geo['lat']:.6f}",
                )

            with g2:
                st.metric(
                    "Kinh độ",
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
                    f"⛰️ **Phân tầng độ cao:** "
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
                        "hiệu chỉnh Hb trong phiên bản nghiên cứu chính thức."
                    )

                st.caption(
                    "Độ cao được lấy từ **xã/phường mà người dùng đã chọn**, "
                    "sau đó truy vấn theo tọa độ của lựa chọn đó."
                )

        else:
            st.warning(
                "Chưa có xã/phường được xác nhận. "
                "Hãy tìm và chọn một địa điểm trước khi phân tích CBC."
            )

    # --------------------------------------------------------
    # CBC
    # --------------------------------------------------------

    with st.container(border=True):

        st.subheader(
            "3. Công thức máu (CBC)"
        )

        b1, b2, b3, b4, b5 = st.columns(5)

        with b1:
            hb = st.number_input(
                "Hb (g/dL)",
                3.0,
                25.0,
                13.0,
                0.1,
            )

        with b2:
            mcv = st.number_input(
                "MCV (fL)",
                30.0,
                150.0,
                85.0,
                0.1,
            )

        with b3:
            mch = st.number_input(
                "MCH (pg)",
                10.0,
                50.0,
                29.0,
                0.1,
            )

        with b4:
            rbc = st.number_input(
                "RBC (T/L)",
                1.0,
                10.0,
                4.8,
                0.1,
            )

        with b5:
            rdw = st.number_input(
                "RDW-CV (%)",
                5.0,
                40.0,
                13.0,
                0.1,
            )

    # --------------------------------------------------------
    # ANALYZE ROUND 2
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
        elif st.session_state.get("location_geo") is None:
            st.error(
                "Vui lòng tìm và **chọn xác nhận một xã/phường** "
                "ở Vòng 2 trước khi phân tích CBC."
            )
        else:

            elevation = st.session_state.get(
                "location_elevation"
            )

            if elevation is None:

                adjustment = 0.0
                hb_adjusted = hb

                st.warning(
                    "Chưa lấy được độ cao. Hb được giữ nguyên và không hiệu chỉnh."
                )

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
                round2_score(
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
                "round2_score_value"
            ] = score2

            st.session_state[
                "round2_mentzer_value"
            ] = mentzer

            st.session_state[
                "round2_category_value"
            ] = category2

            st.session_state[
                "round2_conclusion_value"
            ] = conclusion2

            st.session_state[
                "round2_reasons_value"
            ] = reasons2

            st.session_state[
                "round2_hb_value"
            ] = hb

            st.session_state[
                "round2_hb_adjusted_value"
            ] = hb_adjusted

            st.session_state[
                "round2_adjustment_value"
            ] = adjustment

            st.session_state[
                "round2_mcv_value"
            ] = mcv

            st.session_state[
                "round2_mch_value"
            ] = mch

            st.session_state[
                "round2_rbc_value"
            ] = rbc

            st.session_state[
                "round2_rdw_value"
            ] = rdw

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

        r1, r2, r3, r4 = st.columns(4)

        with r1:
            st.metric(
                "Mentzer Index",
                f"{st.session_state['round2_mentzer_value']:.2f}",
            )

        with r2:
            st.metric(
                "Điểm CBC",
                str(
                    st.session_state[
                        "round2_score_value"
                    ]
                ),
            )

        with r3:
            st.metric(
                "Hb thực đo",
                f"{st.session_state['round2_hb_value']:.1f} g/dL",
            )

        with r4:
            st.metric(
                "Hb hiệu chỉnh",
                f"{st.session_state['round2_hb_adjusted_value']:.1f} g/dL",
            )

        category2 = st.session_state[
            "round2_category_value"
        ]

        conclusion2 = st.session_state[
            "round2_conclusion_value"
        ]

        reasons2 = st.session_state[
            "round2_reasons_value"
        ]

        if category2 == "THẤP":
            st.success(
                f"🟢 **NGUY CƠ VÒNG 2: {category2}**\n\n"
                f"{conclusion2}"
            )
        elif category2 == "TRUNG BÌNH":
            st.warning(
                f"🟡 **NGUY CƠ VÒNG 2: {category2}**\n\n"
                f"{conclusion2}"
            )
        elif category2 == "CAO":
            st.error(
                f"🟠 **NGUY CƠ VÒNG 2: {category2}**\n\n"
                f"{conclusion2}"
            )
        else:
            st.error(
                f"🔴 **NGUY CƠ VÒNG 2: {category2}**\n\n"
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

        # Hb interpretation
        hb_adj = st.session_state[
            "round2_hb_adjusted_value"
        ]

        if patient_age is not None:

            if (
                patient_age >= 15
                and patient["gender"] in [
                    "Nam",
                    "Nữ",
                ]
            ):

                cutoff = (
                    13.0
                    if patient["gender"] == "Nam"
                    else 12.0
                )

                if hb_adj < cutoff:
                    st.info(
                        f"🩸 Hb sau hiệu chỉnh thấp hơn "
                        f"ngưỡng {cutoff:.1f} g/dL đang dùng cho nhóm này."
                    )
                else:
                    st.info(
                        f"🩸 Hb sau hiệu chỉnh chưa thấp hơn "
                        f"ngưỡng {cutoff:.1f} g/dL đang dùng cho nhóm này."
                    )

        # Next step
        st.subheader(
            "🧭 Khuyến nghị bước tiếp theo"
        )

        if category2 == "THẤP":

            st.info(
                "CBC hiện tại chưa cho thấy mẫu hình mạnh gợi ý Thalassemia. "
                "Nếu vẫn có tiền sử hoặc chỉ định lâm sàng, cần được nhân viên y tế "
                "đánh giá thêm."
            )

        elif category2 == "TRUNG BÌNH":

            st.warning(
                "Nên đánh giá nguyên nhân hồng cầu nhỏ, đặc biệt tình trạng thiếu sắt "
                "và cân nhắc xét nghiệm phù hợp. Nếu vẫn nghi ngờ, có thể cần "
                "HPLC/điện di hemoglobin."
            )

        else:

            st.error(
                "Nên được đánh giá tại cơ sở có khả năng thực hiện "
                "HPLC/điện di hemoglobin. Xét nghiệm phân tử có thể được "
                "cân nhắc theo chỉ định chuyên môn."
            )

        # ----------------------------------------------------
        # MEDICAL FACILITIES
        # ----------------------------------------------------

        st.subheader(
            "🏥 CƠ SỞ Y TẾ GẦN NƠI Ở HIỆN TẠI"
        )

        if geo is None:

            st.warning(
                "Hãy xác định tọa độ/độ cao trước khi tìm cơ sở."
            )

        elif not GOOGLE_API_KEY:

            st.warning(
                "Chưa cấu hình Google Places API."
            )

        else:

            places = nearby_medical(
                geo["lat"],
                geo["lng"],
                GOOGLE_API_KEY,
            )

            candidates = rank_facilities(
                places,
                geo["lat"],
                geo["lng"],
            )

            if not candidates:

                st.info(
                    "Không tìm thấy cơ sở trong phạm vi 50 km "
                    "theo truy vấn hiện tại."
                )

            else:

                st.write(
                    "Hệ thống ưu tiên **3–5 cơ sở gần nhất**. "
                    "Thông tin Google Maps không tự xác nhận "
                    "cơ sở có HPLC/điện di/gen."
                )

                for index, facility in enumerate(
                    candidates,
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
                            f"📏 Khoảng cách địa lý ước tính: "
                            f"**{facility['distance']:.1f} km**"
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
                            "Nên xác nhận trước về Ferritin, HPLC/"
                            "điện di Hb và xét nghiệm gen."
                        )

        # ----------------------------------------------------
        # WORD
        # ----------------------------------------------------

        st.subheader(
            "📄 PHIẾU KẾT QUẢ"
        )

        geo_address = (
            geo["formatted_address"]
            if geo
            else patient["current_address"]
        )

        r2_data = {
            "geo_address": geo_address,
            "elevation": elevation,
            "hb": st.session_state[
                "round2_hb_value"
            ],
            "adjustment": st.session_state[
                "round2_adjustment_value"
            ],
            "hb_adjusted": st.session_state[
                "round2_hb_adjusted_value"
            ],
            "mcv": st.session_state[
                "round2_mcv_value"
            ],
            "mch": st.session_state[
                "round2_mch_value"
            ],
            "rbc": st.session_state[
                "round2_rbc_value"
            ],
            "rdw": st.session_state[
                "round2_rdw_value"
            ],
            "mentzer": st.session_state[
                "round2_mentzer_value"
            ],
            "score": st.session_state[
                "round2_score_value"
            ],
            "category": st.session_state[
                "round2_category_value"
            ],
            "conclusion": st.session_state[
                "round2_conclusion_value"
            ],
        }

        patient_for_word = {
            **patient,
            "age": patient_age,
        }

        report = create_word(
            patient=patient_for_word,
            r1_score=st.session_state[
                "round1_score"
            ],
            r1_category=st.session_state[
                "round1_category"
            ],
            r1_reasons=st.session_state[
                "round1_reasons"
            ],
            r2=r2_data,
        )

        st.download_button(
            label="📥 TẢI PHIẾU WORD",
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
    "THALASSEMIA SCREENING V3.3 — PROTOTYPE NGHIÊN CỨU. "
    "Các trọng số/ngưỡng cần được validation trên dữ liệu người Việt Nam "
    "trước khi dùng trong nghiên cứu lâm sàng hoặc triển khai thực tế."
)
