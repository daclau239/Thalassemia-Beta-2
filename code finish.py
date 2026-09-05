import io
import math
import requests
import streamlit as st
from docx import Document

# ============================================================
# THALASSEMIA SCREENING V2
# CBC + tiền sử + hiệu chỉnh Hb theo độ cao + gợi ý cơ sở y tế
#
# Yêu cầu:
#   pip install streamlit python-docx requests
#
# API:
#   Đặt GOOGLE_MAPS_API_KEY trong .streamlit/secrets.toml
#
# Ví dụ:
#   [google]
#   maps_api_key = "YOUR_GOOGLE_MAPS_API_KEY"
#
# App này là công cụ SÀNG LỌC, không phải công cụ chẩn đoán.
# ============================================================

st.set_page_config(
    page_title="Hệ thống Sàng lọc Thalassemia",
    page_icon="🩸",
    layout="wide",
)

st.title("🩸 HỆ THỐNG SÀNG LỌC VÀ PHÂN TẦNG NGUY CƠ THALASSEMIA")
st.caption(
    "Công cụ hỗ trợ sàng lọc ban đầu dựa trên tiền sử, CBC và một số chỉ số "
    "tính toán. Kết quả không thay thế chẩn đoán của bác sĩ."
)

# ------------------------------------------------------------
# CẤU HÌNH GOOGLE
# ------------------------------------------------------------
def get_google_key():
    try:
        return st.secrets["google"]["maps_api_key"]
    except Exception:
        try:
            return st.secrets["GOOGLE_MAPS_API_KEY"]
        except Exception:
            return ""

GOOGLE_API_KEY = get_google_key()


# ------------------------------------------------------------
# HÀM GOOGLE: GEOCODING
# ------------------------------------------------------------
@st.cache_data(ttl=86400)
def google_geocode(address, api_key):
    """Chuyển địa điểm người dùng chọn thành tọa độ."""
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
        response = requests.get(url, params=params, timeout=10)
        response.raise_for_status()
        data = response.json()

        if data.get("status") != "OK" or not data.get("results"):
            return None

        result = data["results"][0]
        location = result["geometry"]["location"]

        return {
            "formatted_address": result.get("formatted_address", address),
            "lat": location["lat"],
            "lng": location["lng"],
        }
    except Exception:
        return None


# ------------------------------------------------------------
# HÀM GOOGLE: ELEVATION
# ------------------------------------------------------------
@st.cache_data(ttl=86400)
def google_elevation(lat, lng, api_key):
    """Lấy độ cao địa hình tại tọa độ bằng Google Elevation API."""
    if not api_key:
        return None

    url = "https://maps.googleapis.com/maps/api/elevation/json"
    params = {
        "locations": f"{lat},{lng}",
        "key": api_key,
    }

    try:
        response = requests.get(url, params=params, timeout=10)
        response.raise_for_status()
        data = response.json()

        if data.get("status") != "OK" or not data.get("results"):
            return None

        return float(data["results"][0]["elevation"])
    except Exception:
        return None


# ------------------------------------------------------------
# HÀM GOOGLE PLACES (NEW): TÌM CƠ SỞ Y TẾ GẦN
# ------------------------------------------------------------
@st.cache_data(ttl=3600)
def google_nearby_medical(lat, lng, api_key, radius=50000):
    """
    Tìm cơ sở y tế gần tọa độ.
    Google Places API (New) - Nearby Search.
    Chỉ dùng để GỢI Ý địa điểm; cần kiểm tra lại dịch vụ thực tế
    trước khi người bệnh đi.
    """
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
            "medical_clinic",
            "medical_center",
        ],
        "maxResultCount": 20,
        "rankPreference": "DISTANCE",
        "locationRestriction": {
            "circle": {
                "center": {
                    "latitude": lat,
                    "longitude": lng,
                },
                "radius": float(radius),
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


# ------------------------------------------------------------
# TÍNH KHOẢNG CÁCH HAVERSINE
# ------------------------------------------------------------
def haversine_km(lat1, lon1, lat2, lon2):
    r = 6371.0

    p1 = math.radians(lat1)
    p2 = math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)

    a = (
        math.sin(dp / 2) ** 2
        + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    )

    return 2 * r * math.asin(math.sqrt(a))


# ------------------------------------------------------------
# WHO 2024: HIỆU CHỈNH Hb THEO ĐỘ CAO
# ------------------------------------------------------------
def who_2024_altitude_adjustment_g_dl(elevation_m):
    """
    WHO 2024 common adjustment.
    Trả về lượng Hb (g/dL) cần TRỪ khỏi Hb quan sát
    để diễn giải theo điều kiện gần mực nước biển.

    Bảng WHO 2024 theo các khoảng 500 m.
    """
    if elevation_m is None or elevation_m < 500:
        return 0.0

    if elevation_m < 1000:
        return 0.4
    elif elevation_m < 1500:
        return 0.8
    elif elevation_m < 2000:
        return 1.1
    elif elevation_m < 2500:
        return 1.4
    elif elevation_m < 3000:
        return 1.8
    elif elevation_m < 3500:
        return 2.1
    elif elevation_m < 4000:
        return 2.5
    elif elevation_m < 4500:
        return 2.9
    elif elevation_m < 5000:
        return 3.3
    else:
        # WHO 2024 lưu ý dữ liệu >2500 m hạn chế hơn và cần thận trọng.
        # Không tự ngoại suy vô hạn.
        return 3.3


# ------------------------------------------------------------
# RISK SCORE — PROTOTYPE
# ------------------------------------------------------------
def calculate_screening_score(mcv, mch, rbc, hb, rdw, family_history,
                              personal_microcytosis, prior_thalassemia_test):
    """
    Đây là SCORE PROTOTYPE để nghiên cứu/giao diện.
    Không phải thang điểm lâm sàng đã được validation.
    """

    score = 0
    reasons = []

    # MCV
    if mcv < 70:
        score += 3
        reasons.append("MCV rất thấp (<70 fL)")
    elif mcv < 75:
        score += 2
        reasons.append("MCV giảm (70–74,9 fL)")
    elif mcv < 80:
        score += 1
        reasons.append("MCV giảm nhẹ (75–79,9 fL)")

    # MCH
    if mch < 24:
        score += 2
        reasons.append("MCH thấp (<24 pg)")
    elif mch < 27:
        score += 1
        reasons.append("MCH giảm (24–26,9 pg)")

    # RBC: chỉ cộng điểm khi microcytosis đi kèm RBC tương đối cao
    if mcv < 80:
        if rbc >= 5.5:
            score += 2
            reasons.append("RBC tương đối cao khi MCV thấp")
        elif rbc >= 5.0:
            score += 1
            reasons.append("RBC tương đối cao khi MCV thấp")

    # Mentzer
    mentzer = mcv / rbc if rbc > 0 else None

    if mentzer is not None and mcv < 80:
        if mentzer < 13:
            score += 2
            reasons.append("Mentzer Index <13")
        elif mentzer < 14:
            score += 1
            reasons.append("Mentzer Index 13–13,9")

    # RDW chỉ là biến hỗ trợ
    if rdw > 15:
        reasons.append("RDW tăng — cần lưu ý thiếu sắt/nguồn microcytosis khác")

    # Family history
    if family_history == "Có":
        score += 2
        reasons.append("Có tiền sử gia đình/dòng họ liên quan Thalassemia")
    elif family_history == "Không biết":
        score += 0

    # Previous microcytosis
    if personal_microcytosis == "Có":
        score += 1
        reasons.append("Từng được thông báo có hồng cầu nhỏ")

    # Prior confirmed carrier/test
    if prior_thalassemia_test == "Đã xác định mang gen":
        score += 4
        reasons.append("Đã từng được xác định mang gen")
    elif prior_thalassemia_test == "Đã nghi ngờ":
        score += 2
        reasons.append("Từng có kết quả nghi ngờ")

    # Hb KHÔNG cộng điểm Thalassemia trực tiếp.
    # Hb dùng chủ yếu để đánh giá thiếu máu và bối cảnh lâm sàng.

    return score, mentzer, reasons


def risk_category(score, mcv):
    if mcv >= 80 and score <= 2:
        return (
            "THẤP",
            "Các thông tin hiện có chưa cho thấy mẫu hình hồng cầu nhỏ rõ ràng.",
        )

    if score <= 3:
        return (
            "THẤP",
            "Nguy cơ sàng lọc hiện tại thấp; kết quả không loại trừ hoàn toàn "
            "Thalassemia.",
        )

    if score <= 6:
        return (
            "TRUNG BÌNH",
            "Có một số đặc điểm gợi ý Thalassemia/hemoglobinopathy. "
            "Nên đánh giá thêm nguyên nhân hồng cầu nhỏ, đặc biệt tình trạng thiếu sắt.",
        )

    if score <= 9:
        return (
            "CAO",
            "Mẫu hình CBC/tiền sử gợi ý cần đánh giá chuyên sâu về "
            "Thalassemia/hemoglobinopathy.",
        )

    return (
        "RẤT CAO",
        "Có nhiều yếu tố sàng lọc đáng chú ý. Cần xét nghiệm xác nhận; "
        "không dùng điểm số này để khẳng định chẩn đoán.",
    )


# ------------------------------------------------------------
# ANEMIA CONTEXT
# ------------------------------------------------------------
def anemia_context(hb_adjusted, sex, age_years):
    """
    Cutoff người trưởng thành không mang thai:
    nam <13 g/dL; nữ <12 g/dL.
    Đây chỉ là phần cảnh báo thiếu máu, không phải chẩn đoán Thalassemia.
    """
    if age_years < 15:
        return (
            "Chưa tự động phân loại thiếu máu bằng ngưỡng người trưởng thành "
            "trong phiên bản này."
        )

    if sex == "Nam":
        cutoff = 13.0
    elif sex == "Nữ":
        cutoff = 12.0
    else:
        return "Cần đánh giá ngưỡng Hb theo giới/đối tượng cụ thể."

    if hb_adjusted < cutoff:
        return (
            f"Hb sau hiệu chỉnh ({hb_adjusted:.1f} g/dL) thấp hơn "
            f"ngưỡng {cutoff:.1f} g/dL đang sử dụng cho nhóm này."
        )

    return (
        f"Hb sau hiệu chỉnh ({hb_adjusted:.1f} g/dL) chưa thấp hơn "
        f"ngưỡng {cutoff:.1f} g/dL đang sử dụng cho nhóm này."
    )


# ------------------------------------------------------------
# WORD
# ------------------------------------------------------------
def tao_file_word(
    ho_ten,
    ngay_sinh,
    gioi_tinh,
    dia_diem,
    elevation,
    hb,
    hb_adjusted,
    mcv,
    mch,
    rbc,
    rdw,
    mentzer,
    score,
    category,
    conclusion,
):
    doc = Document()

    doc.add_heading("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", level=3)
    doc.add_heading(
        "PHIẾU SÀNG LỌC VÀ PHÂN TẦNG NGUY CƠ THALASSEMIA",
        level=1,
    )

    doc.add_paragraph(
        "Lưu ý: Đây là công cụ hỗ trợ sàng lọc, không thay thế chẩn đoán "
        "hoặc chỉ định điều trị của bác sĩ."
    )

    doc.add_heading("I. Thông tin hành chính", level=2)
    doc.add_paragraph(f"Họ và tên: {ho_ten}")
    doc.add_paragraph(f"Ngày sinh: {ngay_sinh}")
    doc.add_paragraph(f"Giới tính: {gioi_tinh}")
    doc.add_paragraph(f"Nơi đang sinh sống: {dia_diem}")
    doc.add_paragraph(
        f"Độ cao ước tính từ dữ liệu địa điểm: "
        f"{elevation:.0f} m so với mực nước biển"
        if elevation is not None
        else "Độ cao: chưa xác định"
    )

    doc.add_heading("II. Chỉ số CBC", level=2)
    doc.add_paragraph(f"Hb đo được: {hb:.1f} g/dL")
    doc.add_paragraph(f"Hb sau hiệu chỉnh độ cao: {hb_adjusted:.1f} g/dL")
    doc.add_paragraph(f"MCV: {mcv:.1f} fL")
    doc.add_paragraph(f"MCH: {mch:.1f} pg")
    doc.add_paragraph(f"RBC: {rbc:.2f} T/L")
    doc.add_paragraph(f"RDW-CV: {rdw:.1f}%")
    doc.add_paragraph(f"Mentzer Index: {mentzer:.2f}")

    doc.add_heading("III. Phân tầng nguy cơ sàng lọc", level=2)
    doc.add_paragraph(f"Điểm prototype: {score}")
    doc.add_paragraph(f"Mức nguy cơ sàng lọc: {category}")
    doc.add_paragraph(f"Nhận định: {conclusion}")

    doc.add_heading("IV. Khuyến nghị", level=2)
    doc.add_paragraph(
        "Nếu kết quả sàng lọc gợi ý nguy cơ trung bình/cao, cần cân nhắc "
        "đánh giá tình trạng sắt và xét nghiệm huyết sắc tố (HPLC/điện di Hb) "
        "theo chỉ định chuyên môn. Trường hợp cần thiết có thể làm xét nghiệm "
        "phân tử để xác định loại Thalassemia."
    )

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ============================================================
# GIAO DIỆN
# ============================================================

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Thông tin người được sàng lọc")

    ho_ten = st.text_input("Họ và tên", value="Nguyễn Văn A")

    ngay_sinh = st.date_input(
        "Ngày sinh",
        value=None,
    )

    gioi_tinh = st.selectbox(
        "Giới tính",
        ["Nam", "Nữ", "Khác"],
    )

with col2:
    st.subheader("2. Tiền sử gia đình & bản thân")

    family_history = st.radio(
        "Gia đình/dòng họ có người mắc hoặc mang gen Thalassemia không?",
        ["Không", "Có", "Không biết"],
    )

    personal_microcytosis = st.radio(
        "Bạn từng được thông báo có hồng cầu nhỏ/MCV thấp chưa?",
        ["Không", "Có", "Không biết"],
    )

    prior_thalassemia_test = st.selectbox(
        "Bạn từng xét nghiệm Thalassemia/hemoglobinopathy chưa?",
        [
            "Chưa xét nghiệm",
            "Đã xét nghiệm, bình thường",
            "Đã nghi ngờ",
            "Đã xác định mang gen",
            "Không nhớ",
        ],
    )

st.divider()

# ------------------------------------------------------------
# ĐỊA ĐIỂM
# ------------------------------------------------------------
st.subheader("3. Nơi đang sinh sống và độ cao")

st.write(
    "Chọn địa điểm gần nhất với nơi người được sàng lọc đang cư trú. "
    "Hệ thống dùng địa điểm này để ước tính độ cao và hỗ trợ đề xuất "
    "cơ sở y tế gần phù hợp."
)

dia_diem = st.text_input(
    "Địa chỉ / xã / huyện / tỉnh",
    value="Đà Nẵng, Việt Nam",
)

if st.button("📍 Xác định vị trí và độ cao", type="secondary"):
    if not GOOGLE_API_KEY:
        st.error(
            "Chưa cấu hình GOOGLE_MAPS_API_KEY. Hãy thêm API key vào "
            ".streamlit/secrets.toml."
        )
    else:
        geo = google_geocode(dia_diem, GOOGLE_API_KEY)

        if not geo:
            st.error(
                "Không xác định được địa điểm. Hãy nhập rõ hơn, ví dụ: "
                "xã, huyện, tỉnh."
            )
        else:
            elevation = google_elevation(
                geo["lat"],
                geo["lng"],
                GOOGLE_API_KEY,
            )

            st.session_state["geo"] = geo
            st.session_state["elevation"] = elevation

geo = st.session_state.get("geo")
elevation = st.session_state.get("elevation")

if geo:
    st.success(f"📍 {geo['formatted_address']}")
    st.write(
        f"**Tọa độ:** {geo['lat']:.6f}, {geo['lng']:.6f}"
    )

if elevation is not None:
    st.info(
        f"⛰️ **Độ cao ước tính:** {elevation:.0f} m so với mực nước biển"
    )

st.divider()

# ------------------------------------------------------------
# CBC
# ------------------------------------------------------------
st.subheader("4. Nhập chỉ số xét nghiệm CBC")

c1, c2, c3, c4, c5 = st.columns(5)

with c1:
    hb = st.number_input(
        "Hb (g/dL)",
        min_value=3.0,
        max_value=25.0,
        value=13.0,
        step=0.1,
    )

with c2:
    mcv = st.number_input(
        "MCV (fL)",
        min_value=30.0,
        max_value=150.0,
        value=85.0,
        step=0.1,
    )

with c3:
    mch = st.number_input(
        "MCH (pg)",
        min_value=10.0,
        max_value=50.0,
        value=29.0,
        step=0.1,
    )

with c4:
    rbc = st.number_input(
        "RBC (T/L)",
        min_value=1.0,
        max_value=10.0,
        value=4.8,
        step=0.1,
    )

with c5:
    rdw = st.number_input(
        "RDW-CV (%)",
        min_value=5.0,
        max_value=40.0,
        value=13.0,
        step=0.1,
    )

st.caption(
    "Nếu phiếu CBC không có RDW, không nên tự ước tính; có thể nhập giá trị "
    "trên phiếu xét nghiệm nếu có."
)

st.divider()

# ------------------------------------------------------------
# NÚT ĐÁNH GIÁ
# ------------------------------------------------------------
if st.button(
    "🩸 PHÂN TÍCH NGUY CƠ VÀ ĐỀ XUẤT CƠ SỞ Y TẾ",
    type="primary",
    use_container_width=True,
):

    if rbc <= 0:
        st.error("RBC phải lớn hơn 0.")
        st.stop()

    # Hb altitude adjustment
    if elevation is None:
        altitude_adjustment = 0.0
        hb_adjusted = hb

        st.warning(
            "Chưa xác định được độ cao. Hb được giữ nguyên và KHÔNG hiệu chỉnh."
        )
    else:
        altitude_adjustment = who_2024_altitude_adjustment_g_dl(elevation)
        hb_adjusted = hb - altitude_adjustment

    # Tuổi
    age_years = None
    if ngay_sinh is not None:
        from datetime import date
        today = date.today()
        age_years = today.year - ngay_sinh.year - (
            (today.month, today.day) < (ngay_sinh.month, ngay_sinh.day)
        )

    score, mentzer, reasons = calculate_screening_score(
        mcv=mcv,
        mch=mch,
        rbc=rbc,
        hb=hb,
        rdw=rdw,
        family_history=family_history,
        personal_microcytosis=personal_microcytosis,
        prior_thalassemia_test=prior_thalassemia_test,
    )

    category, conclusion = risk_category(score, mcv)
    anemia_note = anemia_context(
        hb_adjusted,
        gioi_tinh,
        age_years if age_years is not None else 18,
    )

    # --------------------------------------------------------
    # KẾT QUẢ
    # --------------------------------------------------------
    st.subheader("5. Kết quả sàng lọc")

    r1, r2, r3, r4 = st.columns(4)

    with r1:
        st.metric("Mentzer Index", f"{mentzer:.2f}")

    with r2:
        st.metric("Điểm sàng lọc", str(score))

    with r3:
        st.metric("Hb thực đo", f"{hb:.1f} g/dL")

    with r4:
        st.metric("Hb sau hiệu chỉnh", f"{hb_adjusted:.1f} g/dL")

    if category == "THẤP":
        st.success(f"🟢 **Nguy cơ sàng lọc: {category}**")
    elif category == "TRUNG BÌNH":
        st.warning(f"🟡 **Nguy cơ sàng lọc: {category}**")
    elif category == "CAO":
        st.error(f"🟠 **Nguy cơ sàng lọc: {category}**")
    else:
        st.error(f"🔴 **Nguy cơ sàng lọc: {category}**")

    st.write(conclusion)

    with st.expander("🔎 Các yếu tố đóng góp vào điểm", expanded=True):
        if reasons:
            for reason in reasons:
                st.write(f"• {reason}")
        else:
            st.write("Không có yếu tố cộng điểm đáng chú ý trong prototype.")

    st.info(
        f"**Diễn giải Hb:** {anemia_note}\n\n"
        f"Hiệu chỉnh độ cao theo bảng WHO 2024 đang áp dụng: "
        f"−{altitude_adjustment:.1f} g/dL."
    )

    # --------------------------------------------------------
    # HƯỚNG XỬ TRÍ SÀNG LỌC
    # --------------------------------------------------------
    st.subheader("6. Bước tiếp theo được đề xuất")

    if category == "THẤP":
        st.success(
            "Hiện chưa có chỉ dấu CBC/tiền sử đủ mạnh để ưu tiên chuyển tuyến "
            "vì Thalassemia. Nếu có triệu chứng hoặc tiền sử đặc biệt, vẫn cần "
            "đánh giá lâm sàng."
        )

    elif category == "TRUNG BÌNH":
        st.warning(
            "Nên đánh giá nguyên nhân hồng cầu nhỏ, đặc biệt tình trạng thiếu sắt "
            "(ví dụ Ferritin theo chỉ định). Nếu không giải thích được bằng thiếu sắt "
            "hoặc vẫn nghi ngờ, cân nhắc HPLC/điện di Hb."
        )

    else:
        st.error(
            "Nên được cơ sở có khả năng đánh giá hemoglobinopathy tiếp nhận. "
            "HPLC/điện di Hb và xét nghiệm phân tử có thể được chỉ định tùy trường hợp."
        )

    # --------------------------------------------------------
    # TÌM CƠ SỞ GẦN
    # --------------------------------------------------------
    if category in ["CAO", "RẤT CAO", "TRUNG BÌNH"]:

        st.subheader("7. Cơ sở y tế gợi ý gần nơi cư trú")

        if not geo:
            st.warning(
                "Chưa có tọa độ. Hãy xác định vị trí ở bước 3 để hệ thống "
                "đề xuất cơ sở gần nhất."
            )

        elif not GOOGLE_API_KEY:
            st.warning(
                "Chưa cấu hình Google Places API nên chưa thể tự động lấy "
                "danh sách cơ sở y tế."
            )

        else:
            places = google_nearby_medical(
                geo["lat"],
                geo["lng"],
                GOOGLE_API_KEY,
                radius=50000,
            )

            if not places:
                st.warning(
                    "Không tìm thấy cơ sở phù hợp trong phạm vi tìm kiếm. "
                    "Có thể mở rộng bán kính hoặc tìm kiếm thủ công."
                )
            else:
                # Tính khoảng cách và lấy tối đa 5 cơ sở
                candidates = []

                for place in places:
                    loc = place.get("location", {})
                    lat2 = loc.get("latitude")
                    lng2 = loc.get("longitude")

                    if lat2 is None or lng2 is None:
                        continue

                    distance = haversine_km(
                        geo["lat"],
                        geo["lng"],
                        lat2,
                        lng2,
                    )

                    candidates.append(
                        {
                            "name": place.get("displayName", {}).get(
                                "text", "Cơ sở y tế"
                            ),
                            "address": place.get(
                                "formattedAddress",
                                "Chưa có địa chỉ",
                            ),
                            "distance": distance,
                            "rating": place.get("rating"),
                            "rating_count": place.get("userRatingCount"),
                            "maps_uri": place.get("googleMapsUri"),
                        }
                    )

                candidates.sort(key=lambda x: x["distance"])
                candidates = candidates[:5]

                for i, place in enumerate(candidates, start=1):
                    with st.container(border=True):
                        st.markdown(f"### {i}. {place['name']}")
                        st.write(f"📍 {place['address']}")
                        st.write(
                            f"📏 **Khoảng cách đường chim bay:** "
                            f"{place['distance']:.1f} km"
                        )

                        if place["rating"] is not None:
                            rating_text = f"⭐ {place['rating']:.1f}"
                            if place["rating_count"]:
                                rating_text += (
                                    f" ({place['rating_count']:,} đánh giá)"
                                )
                            st.write(rating_text)

                        if place["maps_uri"]:
                            st.link_button(
                                "🗺️ Mở trên Google Maps",
                                place["maps_uri"],
                            )

                st.caption(
                    "Danh sách được xếp ưu tiên theo khoảng cách địa lý. "
                    "Thông tin Google Maps không xác nhận rằng cơ sở có HPLC, "
                    "điện di Hb hoặc xét nghiệm gen. Nhân viên y tế/người bệnh "
                    "cần xác nhận dịch vụ trước khi đến."
                )

    # --------------------------------------------------------
    # WORD
    # --------------------------------------------------------
    st.subheader("8. Xuất phiếu")

    file_word = tao_file_word(
        ho_ten=ho_ten,
        ngay_sinh=str(ngay_sinh),
        gioi_tinh=gioi_tinh,
        dia_diem=geo["formatted_address"] if geo else dia_diem,
        elevation=elevation,
        hb=hb,
        hb_adjusted=hb_adjusted,
        mcv=mcv,
        mch=mch,
        rbc=rbc,
        rdw=rdw,
        mentzer=mentzer,
        score=score,
        category=category,
        conclusion=conclusion,
    )

    safe_name = "".join(
        c if c.isalnum() or c in " _-" else "_"
        for c in ho_ten
    ).strip() or "nguoi_sang_loc"

    st.download_button(
        label="📥 Tải phiếu kết quả Word",
        data=file_word,
        file_name=f"Phieu_Sang_Loc_Thalassemia_{safe_name}.docx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        ),
    )

st.divider()

st.caption(
    "Phiên bản prototype nghiên cứu. Điểm nguy cơ chưa được validation trên "
    "quần thể người Việt Nam và không được sử dụng như tiêu chuẩn chẩn đoán."
)
