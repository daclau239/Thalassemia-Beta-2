import io
import math
from datetime import date

import requests
import streamlit as st
from docx import Document

# ============================================================
# THALASSEMIA SCREENING V3.1
#
# VÒNG 1:
#   20 câu hỏi sàng lọc nguy cơ ban đầu.
#   Chưa yêu cầu CBC.
#
# VÒNG 2:
#   Chỉ mở khi Vòng 1 đạt ngưỡng NGUY CƠ CAO.
#   Nhập nơi cư trú -> Geocoding -> Elevation -> CBC
#   -> Hb hiệu chỉnh độ cao -> Mentzer/CBC score
#   -> đề xuất 3–5 cơ sở y tế gần/phù hợp.
#
# LƯU Ý:
#   - Risk score là PROTOTYPE, chưa validation trên quần thể Việt Nam.
#   - Không dùng để chẩn đoán hoặc tự điều trị.
#   - WHO 2024: hiệu chỉnh Hb theo độ cao bằng:
#       adjustment (g/L) = 0.0056384*elevation
#                             + 0.0000003*elevation^2
#     và trừ adjustment khỏi Hb quan sát.
#   - Google dùng để xác định địa điểm/độ cao/tìm cơ sở.
#
# Cài:
#   pip install streamlit python-docx requests
#
# .streamlit/secrets.toml:
#
# [google]
# maps_api_key = "YOUR_GOOGLE_MAPS_API_KEY"
#
# Google Cloud cần bật:
#   - Geocoding API
#   - Elevation API
#   - Places API (New)
# ============================================================

st.set_page_config(
    page_title="Hệ thống Sàng lọc Thalassemia",
    page_icon="🩸",
    layout="wide",
)

# ============================================================
# HÀM TIỆN ÍCH
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

ROUND1_HIGH_THRESHOLD = 8
ROUND1_MAX_SCORE = 20


def safe_filename(text):
    text = text.strip()
    cleaned = "".join(
        c if c.isalnum() or c in " _-" else "_"
        for c in text
    )
    return cleaned or "nguoi_sang_loc"


def reset_round2():
    keys = [
        "round2_unlocked",
        "geo",
        "elevation",
        "elevation_resolution",
        "round2_score",
        "round2_category",
        "round2_mentzer",
        "round2_conclusion",
        "round2_reasons",
        "round2_hb",
        "round2_hb_adjusted",
        "round2_altitude_adjustment",
        "round2_mcv",
        "round2_mch",
        "round2_rbc",
        "round2_rdw",
        "round2_dia_diem",
    ]

    for key in keys:
        st.session_state.pop(key, None)


# ============================================================
# VÒNG 1 — 20 CÂU HỎI
# ============================================================

def calculate_round1_score(answers):
    """
    PROTOTYPE.
    Điểm chỉ dùng để quyết định có mở Vòng 2 hay không.
    Không phải thang điểm chẩn đoán đã được validation.
    """

    score = 0
    reasons = []

    # Q1 — gia đình chẩn đoán Thalassemia
    if answers["q1"] == "Có":
        score += 3
        reasons.append("Có người thân/dòng họ mắc Thalassemia")

    # Q2 — gia đình mang gen
    if answers["q2"] == "Có":
        score += 3
        reasons.append("Có người thân/dòng họ mang gen Thalassemia/hemoglobinopathy")

    # Q3 — cha/mẹ từng xét nghiệm
    if answers["q3"] == "Có":
        score += 1
        reasons.append("Cha/mẹ từng được xét nghiệm Thalassemia/hemoglobinopathy")

    # Q4 — anh chị em có thiếu máu/hồng cầu nhỏ
    if answers["q4"] == "Có":
        score += 2
        reasons.append("Anh/chị/em từng được chẩn đoán thiếu máu hoặc hồng cầu nhỏ")

    # Q5 — trẻ truyền máu định kỳ
    if answers["q5"] == "Có":
        score += 2
        reasons.append("Gia đình có trẻ từng truyền máu nhiều lần/định kỳ")

    # Q6 — bản thân thiếu máu
    if answers["q6"] == "Có":
        score += 1
        reasons.append("Từng được thông báo thiếu máu")

    # Q7 — MCV thấp
    if answers["q7"] == "Có":
        score += 2
        reasons.append("Từng được thông báo MCV thấp/hồng cầu nhỏ")

    # Q8 — MCH thấp
    if answers["q8"] == "Có":
        score += 1
        reasons.append("Từng được thông báo MCH thấp/hồng cầu nhược sắc")

    # Q9 — xét nghiệm thal
    if answers["q9"] == "Đã nghi ngờ":
        score += 2
        reasons.append("Từng có kết quả nghi ngờ Thalassemia/hemoglobinopathy")
    elif answers["q9"] == "Đã xác định mang gen":
        score += 4
        reasons.append("Từng được xác định mang gen")

    # Q10 — HbE/hemoglobinopathy
    if answers["q10"] == "Có":
        score += 2
        reasons.append("Từng được chẩn đoán HbE hoặc hemoglobinopathy khác")

    # Q11 — truyền máu
    if answers["q11"] == "Có":
        score += 1
        reasons.append("Bản thân từng truyền máu nhiều lần/định kỳ")

    # Q12 — thiếu máu kéo dài từ nhỏ
    if answers["q12"] == "Có":
        score += 2
        reasons.append("Có tiền sử thiếu máu kéo dài từ nhỏ/tuổi thiếu niên")

    # Q13 — mệt mỏi
    if answers["q13"] == "Có":
        score += 1
        reasons.append("Có triệu chứng mệt mỏi/giảm sức hoạt động")

    # Q14 — chóng mặt
    if answers["q14"] == "Có":
        score += 1
        reasons.append("Có triệu chứng hoa mắt/chóng mặt không rõ nguyên nhân")

    # Q15 — da niêm nhợt
    if answers["q15"] == "Có":
        score += 1
        reasons.append("Từng được nhận xét da/niêm nhợt")

    # Q16 — vàng da
    if answers["q16"] == "Có":
        score += 1
        reasons.append("Từng có vàng da/vàng mắt không rõ nguyên nhân")

    # Q17 — lách to
    if answers["q17"] == "Có":
        score += 2
        reasons.append("Từng được ghi nhận lách to/gan lách to")

    # Q18 — chậm phát triển / biến chứng (câu thăm dò)
    if answers["q18"] == "Có":
        score += 1
        reasons.append("Từng được bác sĩ lưu ý biến chứng liên quan bệnh huyết học mạn")

    # Q19/Q20 KHÔNG cộng điểm nguy cơ bệnh.
    # Chúng được dùng trong Vòng 2 để hiểu khả năng tiếp cận y tế.

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
            "Có một số yếu tố đáng lưu ý nhưng chưa đạt ngưỡng mở "
            "Vòng 2 trong prototype hiện tại.",
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
        response = requests.get(
            url,
            params=params,
            timeout=12,
        )
        response.raise_for_status()

        data = response.json()

        if data.get("status") != "OK" or not data.get("results"):
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
                    component.get("long_name", ""),
                )

        return {
            "formatted_address": result.get(
                "formatted_address",
                address,
            ),
            "lat": float(location["lat"]),
            "lng": float(location["lng"]),
            "place_id": result.get(
                "place_id",
                "",
            ),
            "components": components,
        }

    except Exception:
        return None


# ============================================================
# GOOGLE ELEVATION
# ============================================================

@st.cache_data(ttl=86400)
def google_elevation(lat, lng, api_key):
    if not api_key:
        return None

    url = "https://maps.googleapis.com/maps/api/elevation/json"

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

        if data.get("status") != "OK" or not data.get(
            "results"
        ):
            return None

        result = data["results"][0]

        return {
            "elevation": float(
                result["elevation"]
            ),
            "resolution": float(
                result.get("resolution", 0)
            ),
        }

    except Exception:
        return None


# ============================================================
# WHO 2024 — HIỆU CHỈNH HB THEO ĐỘ CAO
# ============================================================

def who_2024_hb_adjustment_g_dl(elevation_m):
    """
    WHO 2024:
      adjustment (g/L) =
          0.0056384 * elevation
          + 0.0000003 * elevation^2

    Hb adjusted = Hb observed - adjustment

    1 g/dL = 10 g/L.
    """

    if elevation_m is None or elevation_m <= 0:
        return 0.0

    adjustment_g_l = (
        0.0056384 * elevation_m
        + 0.0000003 * (elevation_m ** 2)
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
# VÒNG 2 — CBC
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
        reasons.append("MCV rất thấp (<70 fL)")
    elif mcv < 75:
        score += 2
        reasons.append("MCV giảm rõ (70–74,9 fL)")
    elif mcv < 80:
        score += 1
        reasons.append("MCV giảm (75–79,9 fL)")

    if mch < 24:
        score += 2
        reasons.append("MCH thấp (<24 pg)")
    elif mch < 27:
        score += 1
        reasons.append("MCH giảm (24–26,9 pg)")

    if mcv < 80:
        if rbc >= 5.5:
            score += 2
            reasons.append("RBC tương đối cao khi MCV thấp")
        elif rbc >= 5.0:
            score += 1
            reasons.append("RBC tương đối cao khi MCV thấp")

    mentzer = (
        mcv / rbc
        if rbc > 0
        else None
    )

    if mentzer is not None and mcv < 80:
        if mentzer < 13:
            score += 2
            reasons.append("Mentzer Index <13")
        elif mentzer < 14:
            score += 1
            reasons.append("Mentzer Index 13–13,9")

    if rdw > 15:
        reasons.append(
            "RDW tăng — cần lưu ý thiếu sắt/nguồn microcytosis khác"
        )

    return score, mentzer, reasons


def round2_category(score, mcv):
    if mcv >= 80 and score <= 2:
        return (
            "THẤP",
            "CBC hiện tại chưa cho thấy mẫu hình hồng cầu nhỏ rõ.",
        )

    if score <= 3:
        return (
            "THẤP",
            "Nguy cơ sàng lọc từ CBC hiện tại thấp; không loại trừ hoàn toàn "
            "Thalassemia.",
        )

    if score <= 6:
        return (
            "TRUNG BÌNH",
            "Có đặc điểm hồng cầu nhỏ/nhược sắc. Nên đánh giá tình trạng thiếu sắt "
            "và các nguyên nhân khác.",
        )

    if score <= 9:
        return (
            "CAO",
            "Mẫu hình CBC gợi ý cần đánh giá hemoglobinopathy bằng HPLC/điện di Hb.",
        )

    return (
        "RẤT CAO",
        "Có nhiều dấu hiệu sàng lọc đáng chú ý; cần xét nghiệm xác nhận.",
    )


# ============================================================
# GOOGLE PLACES NEW — CƠ SỞ Y TẾ
# ============================================================

@st.cache_data(ttl=3600)
def google_nearby_medical(lat, lng, api_key):
    if not api_key:
        return []

    url = "https://places.googleapis.com/v1/places:searchNearby"

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
        response = requests.post(
            url,
            headers=headers,
            json=payload,
            timeout=15,
        )
        response.raise_for_status()

        data = response.json()

        return data.get("places", [])

    except Exception:
        return []


def haversine_km(
    lat1,
    lon1,
    lat2,
    lon2,
):
    earth_radius = 6371.0

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
        * earth_radius
        * math.asin(math.sqrt(a))
    )


def rank_medical_facilities(
    places,
    origin_lat,
    origin_lng,
):
    candidates = []

    for place in places:
        location = place.get(
            "location",
            {},
        )

        lat2 = location.get(
            "latitude"
        )
        lng2 = location.get(
            "longitude"
        )

        if lat2 is None or lng2 is None:
            continue

        distance = haversine_km(
            origin_lat,
            origin_lng,
            lat2,
            lng2,
        )

        candidates.append(
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
                "distance": distance,
                "rating": place.get(
                    "rating"
                ),
                "rating_count": place.get(
                    "userRatingCount"
                ),
                "maps_uri": place.get(
                    "googleMapsUri"
                ),
                "primary_type": place.get(
                    "primaryType",
                    "",
                ),
            }
        )

    # Ưu tiên gần trước; điểm đánh giá chỉ là thông tin phụ.
    candidates.sort(
        key=lambda item: item["distance"]
    )

    return candidates[:5]


# ============================================================
# ANEMIA CONTEXT
# ============================================================

def anemia_interpretation(
    hb_adjusted,
    sex,
    age_years,
):
    """
    Chỉ dành cho diễn giải thiếu máu, không phải score Thalassemia.
    Người >=15 tuổi, không mang thai:
      nam <13 g/dL
      nữ <12 g/dL
    """

    if age_years is None or age_years < 15:
        return (
            "Phiên bản hiện tại chưa tự động áp ngưỡng Hb người trưởng thành "
            "cho trẻ <15 tuổi."
        )

    if sex == "Nam":
        cutoff = 13.0
    elif sex == "Nữ":
        cutoff = 12.0
    else:
        return (
            "Cần đánh giá ngưỡng Hb theo nhóm đối tượng cụ thể."
        )

    if hb_adjusted < cutoff:
        return (
            f"Hb sau hiệu chỉnh {hb_adjusted:.1f} g/dL thấp hơn "
            f"ngưỡng {cutoff:.1f} g/dL đang sử dụng."
        )

    return (
        f"Hb sau hiệu chỉnh {hb_adjusted:.1f} g/dL chưa thấp hơn "
        f"ngưỡng {cutoff:.1f} g/dL đang sử dụng."
    )


# ============================================================
# WORD REPORT
# ============================================================

def create_word_report(
    ho_ten,
    ngay_sinh,
    gioi_tinh,
    round1_score,
    round1_category_text,
    round1_reasons,
    dia_diem,
    elevation,
    hb,
    hb_adjusted,
    altitude_adjustment,
    mcv,
    mch,
    rbc,
    rdw,
    mentzer,
    round2_score,
    round2_category_text,
    round2_reasons,
    conclusion,
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
        "Đây là công cụ hỗ trợ sàng lọc. Kết quả không thay thế "
        "chẩn đoán hoặc chỉ định của nhân viên y tế."
    )

    doc.add_heading(
        "I. Thông tin hành chính",
        level=2,
    )
    doc.add_paragraph(
        f"Họ và tên: {ho_ten}"
    )
    doc.add_paragraph(
        f"Ngày sinh: {ngay_sinh}"
    )
    doc.add_paragraph(
        f"Giới tính: {gioi_tinh}"
    )

    doc.add_heading(
        "II. Vòng 1 — Sàng lọc bằng bộ câu hỏi",
        level=2,
    )
    doc.add_paragraph(
        f"Điểm Vòng 1: {round1_score}/{ROUND1_MAX_SCORE}"
    )
    doc.add_paragraph(
        f"Mức nguy cơ: {round1_category_text}"
    )

    if round1_reasons:
        doc.add_paragraph(
            "Các yếu tố đáng chú ý:"
        )
        for reason in round1_reasons:
            doc.add_paragraph(
                f"- {reason}"
            )

    doc.add_heading(
        "III. Vòng 2 — CBC và vị trí cư trú",
        level=2,
    )

    doc.add_paragraph(
        f"Nơi cư trú được xác định: {dia_diem}"
    )

    if elevation is not None:
        doc.add_paragraph(
            f"Độ cao ước tính: {elevation:.0f} m"
        )
    else:
        doc.add_paragraph(
            "Độ cao: chưa xác định"
        )

    doc.add_paragraph(
        f"Hb thực đo: {hb:.1f} g/dL"
    )
    doc.add_paragraph(
        f"Hiệu chỉnh Hb theo độ cao: "
        f"-{altitude_adjustment:.2f} g/dL"
    )
    doc.add_paragraph(
        f"Hb sau hiệu chỉnh: "
        f"{hb_adjusted:.1f} g/dL"
    )

    doc.add_paragraph(
        f"MCV: {mcv:.1f} fL"
    )
    doc.add_paragraph(
        f"MCH: {mch:.1f} pg"
    )
    doc.add_paragraph(
        f"RBC: {rbc:.2f} T/L"
    )
    doc.add_paragraph(
        f"RDW-CV: {rdw:.1f}%"
    )
    doc.add_paragraph(
        f"Mentzer Index: {mentzer:.2f}"
    )

    doc.add_heading(
        "IV. Kết quả Vòng 2",
        level=2,
    )
    doc.add_paragraph(
        f"Điểm CBC prototype: {round2_score}"
    )
    doc.add_paragraph(
        f"Mức nguy cơ Vòng 2: {round2_category_text}"
    )

    if round2_reasons:
        for reason in round2_reasons:
            doc.add_paragraph(
                f"- {reason}"
            )

    doc.add_paragraph(
        f"Nhận định: {conclusion}"
    )

    doc.add_heading(
        "V. Khuyến nghị",
        level=2,
    )
    doc.add_paragraph(
        "Tùy kết quả sàng lọc, có thể cân nhắc Ferritin/đánh giá sắt, "
        "HPLC hoặc điện di hemoglobin và xét nghiệm phân tử theo chỉ định."
    )

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer


# ============================================================
# HEADER
# ============================================================

st.title(
    "🩸 HỆ THỐNG SÀNG LỌC VÀ PHÂN TẦNG NGUY CƠ THALASSEMIA"
)

st.write(
    "Mục tiêu: sàng lọc ban đầu tại tuyến cơ sở, "
    "chỉ chuyển sang đánh giá CBC chuyên sâu khi Vòng 1 "
    "xác định nguy cơ cao, sau đó hỗ trợ điều hướng tới "
    "cơ sở y tế phù hợp gần nơi cư trú."
)

with st.expander(
    "ℹ️ Lưu ý về bản chất của hệ thống",
    expanded=False,
):
    st.write(
        "Hệ thống là prototype nghiên cứu và không được dùng để "
        "khẳng định người bệnh có hoặc không có Thalassemia."
    )
    st.write(
        "Risk score cần được validation trên dữ liệu người Việt Nam "
        "trước khi sử dụng trong nghiên cứu lâm sàng."
    )

# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.header("⚙️ Trạng thái hệ thống")

    if GOOGLE_API_KEY:
        st.success(
            "Google Maps API đã được cấu hình."
        )
    else:
        st.warning(
            "Chưa cấu hình Google Maps API. "
            "Bạn vẫn có thể chạy Vòng 1 và nhập CBC ở Vòng 2, "
            "nhưng không tự động lấy được độ cao/cơ sở y tế."
        )

    st.divider()

    st.write(
        "**Luồng:**\n\n"
        "Vòng 1 → nguy cơ cao → Vòng 2 → CBC + độ cao → "
        "phân tầng → cơ sở y tế"
    )


# ============================================================
# VÒNG 1
# ============================================================

st.header(
    "🟦 VÒNG 1 — BỘ CÂU HỎI SÀNG LỌC BAN ĐẦU"
)

st.info(
    "Vòng 1 **không yêu cầu CBC**. Người dùng trả lời bộ câu hỏi trước. "
    "Chỉ khi đạt ngưỡng nguy cơ cao, Vòng 2 mới xuất hiện."
)

# ------------------------------------------------------------
# A. GIA ĐÌNH
# ------------------------------------------------------------

with st.container(border=True):
    st.subheader(
        "A. Tiền sử gia đình"
    )

    q1 = st.radio(
        "1. Trong gia đình/dòng họ có người từng được chẩn đoán Thalassemia không?",
        ["Không", "Có", "Không biết"],
    )

    q2 = st.radio(
        "2. Trong gia đình/dòng họ có người từng được thông báo mang gen Thalassemia/hemoglobinopathy không?",
        ["Không", "Có", "Không biết"],
    )

    q3 = st.radio(
        "3. Cha hoặc mẹ bạn có từng được xét nghiệm Thalassemia/hemoglobinopathy không?",
        ["Không", "Có", "Không biết"],
    )

    q4 = st.radio(
        "4. Anh/chị/em ruột của bạn có từng được chẩn đoán thiếu máu hoặc hồng cầu nhỏ không?",
        ["Không", "Có", "Không biết"],
    )

    q5 = st.radio(
        "5. Gia đình có trẻ từng phải truyền máu nhiều lần hoặc định kỳ không?",
        ["Không", "Có", "Không biết"],
    )


# ------------------------------------------------------------
# B. BẢN THÂN
# ------------------------------------------------------------

with st.container(border=True):
    st.subheader(
        "B. Tiền sử bản thân"
    )

    q6 = st.radio(
        "6. Bạn từng được nhân viên y tế thông báo bị thiếu máu chưa?",
        ["Không", "Có", "Không biết"],
    )

    q7 = st.radio(
        "7. Bạn từng được thông báo MCV thấp/hồng cầu nhỏ chưa?",
        ["Không", "Có", "Không biết"],
    )

    q8 = st.radio(
        "8. Bạn từng được thông báo MCH thấp/hồng cầu nhược sắc chưa?",
        ["Không", "Có", "Không biết"],
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
    )

    q11 = st.radio(
        "11. Bản thân bạn từng truyền máu nhiều lần hoặc truyền máu định kỳ chưa?",
        ["Không", "Có", "Không biết"],
    )

    q12 = st.radio(
        "12. Bạn có tiền sử thiếu máu kéo dài từ nhỏ hoặc từ tuổi thiếu niên không?",
        ["Không", "Có", "Không biết"],
    )


# ------------------------------------------------------------
# C. DẤU HIỆU HỖ TRỢ
# ------------------------------------------------------------

with st.container(border=True):
    st.subheader(
        "C. Dấu hiệu và tiền sử lâm sàng hỗ trợ"
    )

    q13 = st.radio(
        "13. Bạn có thường xuyên mệt mỏi hoặc giảm khả năng hoạt động không?",
        ["Không", "Có", "Không biết"],
    )

    q14 = st.radio(
        "14. Bạn có thường xuyên hoa mắt/chóng mặt không rõ nguyên nhân không?",
        ["Không", "Có", "Không biết"],
    )

    q15 = st.radio(
        "15. Bạn từng được nhận xét da hoặc niêm mạc nhợt hơn bình thường chưa?",
        ["Không", "Có", "Không biết"],
    )

    q16 = st.radio(
        "16. Bạn từng có vàng da/vàng mắt không rõ nguyên nhân chưa?",
        ["Không", "Có", "Không biết"],
    )

    q17 = st.radio(
        "17. Bạn từng được bác sĩ ghi nhận lách to hoặc gan lách to chưa?",
        ["Không", "Có", "Không biết"],
    )

    q18 = st.radio(
        "18. Bạn từng được bác sĩ lưu ý có biến chứng liên quan bệnh huyết học mạn chưa?",
        ["Không", "Có", "Không biết"],
    )


# ------------------------------------------------------------
# D. TIẾP CẬN XÉT NGHIỆM — KHÔNG CỘNG ĐIỂM
# ------------------------------------------------------------

with st.container(border=True):
    st.subheader(
        "D. Khả năng tiếp cận xét nghiệm"
    )

    q19 = st.radio(
        "19. Bạn hiện có kết quả CBC trong vòng 6–12 tháng gần đây không?",
        ["Không", "Có", "Không biết"],
    )

    q20 = st.radio(
        "20. Bạn có gặp khó khăn khi đi đến cơ sở có xét nghiệm chuyên sâu "
        "do khoảng cách, chi phí hoặc thời gian di chuyển không?",
        ["Không", "Có", "Không biết"],
    )

    st.caption(
        "Câu 19–20 không được cộng vào nguy cơ Thalassemia. "
        "Chúng phục vụ bước điều hướng và tiếp cận y tế."
    )


st.divider()

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

    st.session_state[
        "round1_category"
    ] = category1

    st.session_state[
        "round1_conclusion"
    ] = conclusion1

    # Tính lại Vòng 1 -> xóa toàn bộ Vòng 2 cũ.
    reset_round2()


# ============================================================
# KẾT QUẢ VÒNG 1
# ============================================================

if "round1_score" in st.session_state:

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
        st.metric(
            "Điểm Vòng 1",
            f"{score1}/{ROUND1_MAX_SCORE}",
        )

    with a2:
        st.metric(
            "Ngưỡng mở Vòng 2",
            f"≥{ROUND1_HIGH_THRESHOLD}",
        )

    if category1 == "CAO":

        st.error(
            f"🔴 **NGUY CƠ VÒNG 1: CAO**\n\n"
            f"{conclusion1}"
        )

        if reasons1:
            with st.expander(
                "🔎 Các yếu tố đáng chú ý",
                expanded=True,
            ):
                for reason in reasons1:
                    st.write(
                        f"• {reason}"
                    )

        st.success(
            "✅ Vòng 2 đã được mở. "
            "Bây giờ mới cần nhập địa điểm cư trú và CBC."
        )

        st.session_state[
            "round2_unlocked"
        ] = True

    elif category1 == "TRUNG BÌNH":

        st.warning(
            f"🟡 **NGUY CƠ VÒNG 1: TRUNG BÌNH**\n\n"
            f"{conclusion1}"
        )

        if reasons1:
            with st.expander(
                "🔎 Các yếu tố đáng chú ý",
                expanded=False,
            ):
                for reason in reasons1:
                    st.write(
                        f"• {reason}"
                    )

        st.info(
            "Prototype hiện tại chưa tự động mở Vòng 2 "
            "ở mức nguy cơ trung bình."
        )

    else:

        st.success(
            f"🟢 **NGUY CƠ VÒNG 1: THẤP**\n\n"
            f"{conclusion1}"
        )

        if reasons1:
            with st.expander(
                "🔎 Các yếu tố đáng chú ý",
                expanded=False,
            ):
                for reason in reasons1:
                    st.write(
                        f"• {reason}"
                    )

        st.info(
            "Chưa cần chuyển sang Vòng 2 trong prototype."
        )


# ============================================================
# VÒNG 2 — CHỈ CHẠY KHI VÒNG 1 CAO
# ============================================================

if st.session_state.get(
    "round2_unlocked",
    False,
):

    st.divider()

    st.header(
        "🟧 VÒNG 2 — CBC + ĐỊA ĐIỂM + ĐỘ CAO"
    )

    st.info(
        "Vòng 2 chỉ xuất hiện sau khi Vòng 1 đạt ngưỡng nguy cơ cao."
    )

    # --------------------------------------------------------
    # 1. THÔNG TIN
    # --------------------------------------------------------

    with st.container(border=True):

        st.subheader(
            "1. Thông tin người được sàng lọc"
        )

        c1, c2, c3 = st.columns(3)

        with c1:
            ho_ten = st.text_input(
                "Họ và tên",
                value="Nguyễn Văn A",
            )

        with c2:
            ngay_sinh = st.date_input(
                "Ngày sinh",
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

        with c3:
            gioi_tinh = st.selectbox(
                "Giới tính",
                ["Nam", "Nữ", "Khác"],
            )

    # --------------------------------------------------------
    # 2. ĐỊA ĐIỂM
    # --------------------------------------------------------

    with st.container(border=True):

        st.subheader(
            "2. Chọn nơi đang sinh sống"
        )

        st.write(
            "Người dùng chọn **xã/phường/thị trấn + huyện/quận + tỉnh/thành**. "
            "Hệ thống tự xác định tọa độ và độ cao địa hình."
        )

        loc1, loc2 = st.columns(2)

        with loc1:
            tinh = st.text_input(
                "Tỉnh / thành phố",
                value="Đà Nẵng",
            )

            huyen = st.text_input(
                "Huyện / quận / thị xã",
                value="",
            )

        with loc2:
            xa = st.text_input(
                "Xã / phường / thị trấn",
                value="",
            )

            dia_chi_bo_sung = st.text_input(
                "Địa chỉ bổ sung (nếu cần)",
                value="",
            )

        if st.button(
            "📍 XÁC ĐỊNH XÃ + TỌA ĐỘ + ĐỘ CAO",
            type="secondary",
        ):

            if not GOOGLE_API_KEY:
                st.error(
                    "Chưa cấu hình GOOGLE_MAPS_API_KEY."
                )
            else:

                parts = []

                if dia_chi_bo_sung.strip():
                    parts.append(
                        dia_chi_bo_sung.strip()
                    )

                if xa.strip():
                    parts.append(
                        xa.strip()
                    )

                if huyen.strip():
                    parts.append(
                        huyen.strip()
                    )

                if tinh.strip():
                    parts.append(
                        tinh.strip()
                    )

                parts.append(
                    "Việt Nam"
                )

                query = ", ".join(
                    parts
                )

                geo = google_geocode(
                    query,
                    GOOGLE_API_KEY,
                )

                if not geo:
                    st.error(
                        "Không xác định được địa điểm. "
                        "Hãy nhập rõ xã/phường + huyện/quận + tỉnh/thành."
                    )

                else:
                    elev = google_elevation(
                        geo["lat"],
                        geo["lng"],
                        GOOGLE_API_KEY,
                    )

                    st.session_state[
                        "geo"
                    ] = geo

                    if elev:
                        st.session_state[
                            "elevation"
                        ] = elev[
                            "elevation"
                        ]

                        st.session_state[
                            "elevation_resolution"
                        ] = elev[
                            "resolution"
                        ]

                    else:
                        st.session_state[
                            "elevation"
                        ] = None

                        st.session_state[
                            "elevation_resolution"
                        ] = None

        geo = st.session_state.get(
            "geo"
        )

        elevation = st.session_state.get(
            "elevation"
        )

        elevation_resolution = (
            st.session_state.get(
                "elevation_resolution"
            )
        )

        if geo:

            st.success(
                f"📍 **Địa điểm:** "
                f"{geo['formatted_address']}"
            )

            g1, g2 = st.columns(2)

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

            if elevation is not None:

                adjustment = (
                    who_2024_hb_adjustment_g_dl(
                        elevation
                    )
                )

                st.info(
                    f"⛰️ **Độ cao:** "
                    f"{elevation:.0f} m\n\n"
                    f"📐 **Hiệu chỉnh Hb dự kiến:** "
                    f"-{adjustment:.2f} g/dL\n\n"
                    f"📊 **Phân tầng độ cao:** "
                    f"{altitude_band(elevation)}"
                )

                if elevation_resolution:
                    st.caption(
                        f"Resolution dữ liệu địa hình: "
                        f"{elevation_resolution:.0f} m."
                    )

                if elevation >= 2500:
                    st.warning(
                        "Độ cao ≥2.500 m: cần thận trọng khi diễn giải Hb. "
                        "WHO lưu ý mức độ không chắc chắn cao hơn ở độ cao lớn."
                    )

                st.caption(
                    "Độ cao được lấy tự động từ tọa độ địa điểm; "
                    "người dùng không cần tự nhập độ cao."
                )

            else:
                st.warning(
                    "Đã xác định địa điểm nhưng chưa lấy được độ cao."
                )

    # --------------------------------------------------------
    # 3. CBC
    # --------------------------------------------------------

    with st.container(border=True):

        st.subheader(
            "3. Nhập chỉ số CBC"
        )

        b1, b2, b3, b4, b5 = st.columns(5)

        with b1:
            hb = st.number_input(
                "Hb (g/dL)",
                min_value=3.0,
                max_value=25.0,
                value=13.0,
                step=0.1,
            )

        with b2:
            mcv = st.number_input(
                "MCV (fL)",
                min_value=30.0,
                max_value=150.0,
                value=85.0,
                step=0.1,
            )

        with b3:
            mch = st.number_input(
                "MCH (pg)",
                min_value=10.0,
                max_value=50.0,
                value=29.0,
                step=0.1,
            )

        with b4:
            rbc = st.number_input(
                "RBC (T/L)",
                min_value=1.0,
                max_value=10.0,
                value=4.8,
                step=0.1,
            )

        with b5:
            rdw = st.number_input(
                "RDW-CV (%)",
                min_value=5.0,
                max_value=40.0,
                value=13.0,
                step=0.1,
            )

        st.caption(
            "CBC là dữ liệu đầu vào quan trọng của Vòng 2. "
            "MCH/RDW được dùng hỗ trợ diễn giải, không dùng đơn độc để chẩn đoán."
        )

    # --------------------------------------------------------
    # 4. ANALYSIS
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
            st.stop()

        elevation = st.session_state.get(
            "elevation"
        )

        if elevation is None:
            altitude_adjustment = 0.0
            hb_adjusted = hb

            st.warning(
                "Chưa lấy được độ cao; Hb giữ nguyên và không hiệu chỉnh."
            )

        else:
            altitude_adjustment = (
                who_2024_hb_adjustment_g_dl(
                    elevation
                )
            )

            hb_adjusted = (
                hb - altitude_adjustment
            )

        score2, mentzer, reasons2 = (
            calculate_round2_score(
                mcv=mcv,
                mch=mch,
                rbc=rbc,
                rdw=rdw,
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
            "round2_category"
        ] = category2

        st.session_state[
            "round2_mentzer"
        ] = mentzer

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
            "round2_altitude_adjustment"
        ] = altitude_adjustment

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
            "round2_dia_diem"
        ] = (
            geo[
                "formatted_address"
            ]
            if geo
            else "Chưa xác định"
        )

    # --------------------------------------------------------
    # 5. DISPLAY RESULT
    # --------------------------------------------------------

    if "round2_score" in st.session_state:

        st.subheader(
            "📊 KẾT QUẢ VÒNG 2"
        )

        rr1, rr2, rr3, rr4 = st.columns(4)

        with rr1:
            st.metric(
                "Mentzer Index",
                f"{st.session_state['round2_mentzer']:.2f}",
            )

        with rr2:
            st.metric(
                "Điểm CBC",
                str(
                    st.session_state[
                        "round2_score"
                    ]
                ),
            )

        with rr3:
            st.metric(
                "Hb thực đo",
                f"{st.session_state['round2_hb']:.1f} g/dL",
            )

        with rr4:
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

        reasons2 = st.session_state[
            "round2_reasons"
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

        # Hb context
        if ngay_sinh:

            today = date.today()

            age_years = (
                today.year
                - ngay_sinh.year
                - (
                    (
                        today.month,
                        today.day,
                    )
                    < (
                        ngay_sinh.month,
                        ngay_sinh.day,
                    )
                )
            )

        else:
            age_years = None

        st.info(
            anemia_interpretation(
                st.session_state[
                    "round2_hb_adjusted"
                ],
                gioi_tinh,
                age_years,
            )
        )

        # ----------------------------------------------------
        # 6. NEXT STEP
        # ----------------------------------------------------

        st.subheader(
            "🧭 Bước tiếp theo"
        )

        if category2 == "THẤP":

            st.info(
                "Chưa ghi nhận mẫu hình CBC mạnh gợi ý Thalassemia. "
                "Nếu có tiền sử hoặc chỉ định lâm sàng đặc biệt, vẫn cần đánh giá."
            )

        elif category2 == "TRUNG BÌNH":

            st.warning(
                "Ưu tiên đánh giá nguyên nhân hồng cầu nhỏ, đặc biệt tình trạng "
                "thiếu sắt; cân nhắc Ferritin. Nếu vẫn nghi ngờ, có thể đánh giá "
                "HPLC/điện di Hb."
            )

        else:

            st.error(
                "Nên được đánh giá tại cơ sở có khả năng thực hiện "
                "HPLC/điện di Hb. Xét nghiệm phân tử có thể được cân nhắc "
                "theo chỉ định chuyên môn."
            )

        # ----------------------------------------------------
        # 7. CƠ SỞ Y TẾ
        # ----------------------------------------------------

        st.subheader(
            "🏥 GỢI Ý CƠ SỞ Y TẾ GẦN NƠI CƯ TRÚ"
        )

        if not geo:

            st.warning(
                "Chưa có tọa độ. Hãy xác định vị trí ở mục 2."
            )

        elif not GOOGLE_API_KEY:

            st.warning(
                "Chưa có Google Places API."
            )

        else:

            places = google_nearby_medical(
                geo["lat"],
                geo["lng"],
                GOOGLE_API_KEY,
            )

            candidates = rank_medical_facilities(
                places=places,
                origin_lat=geo["lat"],
                origin_lng=geo["lng"],
            )

            if not candidates:

                st.info(
                    "Không tìm thấy cơ sở y tế trong bán kính 50 km "
                    "theo truy vấn hiện tại."
                )

            else:

                st.write(
                    "Hệ thống ưu tiên tối đa **5 cơ sở gần nhất**. "
                    "Thông tin trên Google không tự xác nhận khả năng HPLC/điện di/gen."
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

                        if facility["rating"] is not None:

                            rating_text = (
                                f"⭐ {facility['rating']:.1f}"
                            )

                            if facility[
                                "rating_count"
                            ]:
                                rating_text += (
                                    f" "
                                    f"({facility['rating_count']:,} đánh giá)"
                                )

                            st.write(
                                rating_text
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
                            "Nên gọi xác nhận trước về dịch vụ "
                            "HPLC/điện di Hb, Ferritin hoặc xét nghiệm gen."
                        )

        # ----------------------------------------------------
        # 8. WORD
        # ----------------------------------------------------

        st.subheader(
            "📄 XUẤT PHIẾU"
        )

        report = create_word_report(
            ho_ten=ho_ten,
            ngay_sinh=str(
                ngay_sinh
            ),
            gioi_tinh=gioi_tinh,
            round1_score=st.session_state[
                "round1_score"
            ],
            round1_category_text=st.session_state[
                "round1_category"
            ],
            round1_reasons=st.session_state[
                "round1_reasons"
            ],
            dia_diem=st.session_state.get(
                "round2_dia_diem",
                "Chưa xác định",
            ),
            elevation=st.session_state.get(
                "elevation"
            ),
            hb=st.session_state[
                "round2_hb"
            ],
            hb_adjusted=st.session_state[
                "round2_hb_adjusted"
            ],
            altitude_adjustment=st.session_state[
                "round2_altitude_adjustment"
            ],
            mcv=st.session_state[
                "round2_mcv"
            ],
            mch=st.session_state[
                "round2_mch"
            ],
            rbc=st.session_state[
                "round2_rbc"
            ],
            rdw=st.session_state[
                "round2_rdw"
            ],
            mentzer=st.session_state[
                "round2_mentzer"
            ],
            round2_score=st.session_state[
                "round2_score"
            ],
            round2_category_text=st.session_state[
                "round2_category"
            ],
            round2_reasons=st.session_state[
                "round2_reasons"
            ],
            conclusion=st.session_state[
                "round2_conclusion"
            ],
        )

        st.download_button(
            label="📥 Tải phiếu kết quả Word",
            data=report,
            file_name=(
                "Phieu_Thalassemia_"
                f"{safe_filename(ho_ten)}.docx"
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
    "THALASSEMIA SCREENING V3.1 — PROTOTYPE NGHIÊN CỨU. "
    "Các ngưỡng điểm cần được validation trên dữ liệu người Việt Nam "
    "trước khi sử dụng trong nghiên cứu/triển khai thực tế."
)
