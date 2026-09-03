import base64
import io
import os
import json
import sqlite3
import uuid
from datetime import date, datetime

import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

# Mật khẩu quản trị viên
ADMIN_PASSWORD = "admin123"

# 1. Cấu hình trang
st.set_page_config(
    page_title="Hệ thống Sàng lọc & Tư vấn Di truyền Thalassemia",
    page_icon="🩸",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 2. Giao diện CSS
CSS = """<style>
.stApp { background: linear-gradient(180deg, #E6F4F8 0%, #F4F8FB 100%); }
.block-container { padding-top: 1.2rem; padding-bottom: 3.5rem; max-width: 1180px; }
header, footer, div[data-testid="stDecoration"] { display: none; }
.hero { background: linear-gradient(135deg, #012A4A 0%, #014F86 50%, #0077B6 100%);
        color: #fff; border-radius: 18px; padding: 22px 26px; margin-bottom: 16px; }
.hero h1 { color: #fff !important; font-size: 24px !important; margin: 0 0 6px 0 !important; }
.hero p { color: #D9F3FF; font-size: 14px; margin: 0; }
.info-box, .warning-box, .danger-box, .success-box {
    padding: 12px 15px; border-radius: 12px; margin: 10px 0; line-height: 1.6; font-size: 14px; }
.info-box { background: #E7F6FB; border-left: 5px solid #0077B6; }
.warning-box { background: #FFF6E8; border-left: 5px solid #F4A261; }
.danger-box { background: #FDECEC; border-left: 5px solid #D62828; }
.success-box { background: #E9F7EF; border-left: 5px solid #2A9D8F; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


def box(kind, text):
    st.markdown(f'<div class="{kind}">{text}</div>', unsafe_allow_html=True)


# 3. Danh mục dữ liệu chuẩn
VUNG_MIEN = [
    "Chọn vùng/miền",
    "Đông Bắc",
    "Tây Bắc",
    "Đồng bằng sông Hồng",
    "Bắc Trung Bộ",
    "Trung Trung Bộ",
    "Nam Trung Bộ",
    "Tây Nguyên",
    "Đồng bằng sông Cửu Long",
]

TINH_THEO_VUNG = {
    "Đông Bắc": [
        "Hà Giang",
        "Cao Bằng",
        "Bắc Kạn - Thái Nguyên",
        "Tuyên Quang",
        "Lạng Sơn",
        "Bắc Giang",
    ],
    "Tây Bắc": ["Điện Biên - Lai Châu", "Sơn La", "Hòa Bình", "Yên Bái"],
    "Đồng bằng sông Hồng": [
        "Hà Nội",
        "Hải Phòng",
        "Vĩnh Phúc - Phú Thọ",
        "Bắc Ninh - Hưng Yên",
        "Quảng Ninh",
        "Hải Dương",
        "Thái Bình",
        "Nam Định - Ninh Bình",
        "Thanh Hóa",
    ],
    "Bắc Trung Bộ": [
        "Nghệ An",
        "Hà Tĩnh",
        "Quảng Bình",
        "Quảng Trị - Thừa Thiên Huế",
    ],
    "Trung Trung Bộ": [
        "Đà Nẵng",
        "Quảng Nam - Quảng Ngãi",
        "Bình Định",
        "Phú Yên",
    ],
    "Nam Trung Bộ": ["Khánh Hòa", "Ninh Thuận - Bình Thuận"],
    "Tây Nguyên": ["Kon Tum", "Gia Lai", "Đắc Lắc", "Đắc Nông - Lâm Đồng"],
    "Đồng bằng sông Cửu Long": [
        "Tiền Giang - Vĩnh Long",
        "Bến Tre - Trà Vinh",
        "Đồng Tháp - An Giang",
        "Kiên Giang - Hậu Giang",
        "Cần Thơ",
        "Sóc Trăng - Bạc Liêu - Cà Mau",
    ],
}

TAT_CA_TINH = ["Chọn Tỉnh/Thành phố"]
for ds in TINH_THEO_VUNG.values():
    TAT_CA_TINH.extend(ds)

DAN_TOC = [
    "Chọn dân tộc",
    "Stiêng",
    "Ê Đê",
    "Gia Rai",
    "Ba Na",
    "Xơ Đăng",
    "Cơ Ho",
    "Hrê",
    "Chăm",
    "Khơ Me",
    "Thái",
    "Mường",
    "Tày",
    "Nùng",
    "Dao",
    "Sán Chay",
    "Kinh",
    "Hoa",
    "Dân tộc khác",
]
DIEM_DAN_TOC_VN = {
    "Stiêng": 3.0,
    "Ê Đê": 3.0,
    "Gia Rai": 3.0,
    "Ba Na": 2.5,
    "Xơ Đăng": 2.5,
    "Cơ Ho": 2.5,
    "Hrê": 2.5,
    "Chăm": 2.0,
    "Khơ Me": 2.0,
    "Thái": 2.0,
    "Mường": 2.0,
    "Tày": 1.5,
    "Nùng": 1.5,
    "Dao": 1.5,
    "Sán Chay": 1.5,
    "Kinh": 0.5,
    "Hoa": 0.5,
    "Dân tộc khác": 0.5,
}

# 4. Danh mục câu hỏi Vòng 1 phân theo 4 Nhóm chuẩn lâm sàng
CAU_HOI_NHOM_A = [
    (
        "q_a1",
        "Gia đình/dòng họ có người đã chẩn đoán Thalassemia hoặc cần truyền máu định kỳ?",
    ),
    (
        "q_a2",
        "Gia đình có người bị lách to, biến dạng xương mặt, da vàng hoặc sạm đen bất thường?",
    ),
    (
        "q_a3",
        "Tiền sử gia đình có ghi nhận trường hợp sảy thai liên tiếp, thai chết lưu hoặc phù thai?",
    ),
]

CAU_HOI_NHOM_B = [
    (
        "q_b1",
        "Bản thân từng có kết quả xét nghiệm ghi nhận thiếu máu nhược sắc, hồng cầu nhỏ?",
    ),
    (
        "q_b2",
        "Từng làm Điện di Hb / Xét nghiệm Gen có nghi ngờ mang gen hoặc biến thể Hb (HbE, HbCS...)?",
    ),
    (
        "q_b3",
        "Bản thân từng bị từ chối hiến máu do nồng độ Huyết sắc tố (Hb) thấp?",
    ),
]

CAU_HOI_NHOM_C = [
    (
        "q_c1",
        "Bản thân hoặc bố/mẹ thuộc các dân tộc thiểu số có tỷ lệ mang gen cao (Thái, Mường, Tày, Nùng, Ê Đê...)?",
    ),
    (
        "q_c2",
        "Vợ/Chồng hoặc bạn đời thuộc cùng dòng họ, cùng dân tộc thiểu số hoặc cùng thôn/bản?",
    ),
]

CAU_HOI_NHOM_D = [
    (
        "q_d1",
        "Thường xuyên mệt mỏi, hoa mắt, chóng mặt, giảm khả năng tập trung kéo dài?",
    ),
    (
        "q_d2",
        "Da xanh xao, niêm mạc nhợt nhạt không rõ nguyên nhân dù ăn uống bình thường?",
    ),
]


# 5. Phân tích Vòng 2
def analyze_round2(mcv, mch, hb, rbc, rdw, gioitinh):
    if mcv < 80.0:
        morphology = "Microcytic (Hồng cầu nhỏ)"
    elif mcv > 100.0:
        morphology = "Macrocytic (Hồng cầu to)"
    else:
        morphology = "Normocytic (Kích thước bình thường)"

    chromic = "Nhược sắc (MCH < 27 pg)" if mch < 27.0 else "Đẳng sắc"
    mentzer = (mcv / rbc) if (rbc > 0 and mcv > 0) else 0.0

    differential = []
    if mcv < 85.0 or mch < 28.0:
        if 0 < mentzer < 13.0:
            differential.append(
                "Mẫu hình huyết học **gợi ý nghiêng nhiều về Thalassemia** (Mentzer Index < 13)."
            )
        elif mentzer >= 13.0:
            differential.append(
                "Mẫu hình huyết học **gợi ý nghiêng nhiều về Thiếu máu thiếu sắt** (Mentzer Index ≥ 13)."
            )
        else:
            differential.append(
                "Có biểu hiện hồng cầu nhỏ nhược sắc, cần làm thêm Ferritin serum."
            )

        if rdw > 15.0:
            differential.append(
                "RDW > 15%: Hồng cầu kích thước không đều (thường gặp ở thiếu sắt tiến triển)."
            )
        else:
            differential.append(
                "RDW bình thường/tăng nhẹ: Thường gặp ở người mang gen Thalassemia."
            )
    elif mcv > 100.0:
        differential.append(
            "Gợi ý nguyên nhân khác: Thiếu B12, Acid Folic hoặc bệnh lý gan."
        )
    else:
        differential.append(
            "Các chỉ số thể tích và Hb hồng cầu trong giới hạn bình thường."
        )

    return morphology, chromic, mentzer, differential


# 6. Cơ sở dữ liệu SQLite
DATA_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "thalalassemia_v3.db"
)


def _db_connect():
    conn = sqlite3.connect(DATA_FILE)
    conn.execute(
        "CREATE TABLE IF NOT EXISTS ho_so (HoSoID TEXT PRIMARY KEY, data TEXT NOT NULL, updated_at TEXT NOT NULL)"
    )
    conn.commit()
    return conn


def ghi_du_lieu(record_data):
    try:
        conn = _db_connect()
        now = datetime.now().isoformat(timespec="seconds")
        ho_so_id = str(record_data.get("HoSoID") or uuid.uuid4().hex)
        record_data["HoSoID"] = ho_so_id
        conn.execute(
            "INSERT OR REPLACE INTO ho_so (HoSoID, data, updated_at) VALUES (?, ?, ?)",
            (ho_so_id, json.dumps(record_data, ensure_ascii=False), now),
        )
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Lỗi ghi dữ liệu: {e}")
        return False


# 7. Khởi tạo State
def init_state():
    defaults = {
        "ho_so_id": "",
        "hoten": "",
        "gioitinh": "Nữ",
        "ngaysinh": date(2000, 1, 1),
        "dantoc": "Chọn dân tộc",
        "vung_o": "Chọn vùng/miền",
        "tinh_o": "Chọn Tỉnh/Thành phố",
        "s1_score": 0.0,
        "mcv": 0.0,
        "mch": 0.0,
        "hb": 0.0,
        "rbc": 0.0,
        "rdw": 0.0,
        "dien_di": "Chưa thực hiện",
        "gen_test": "Chưa thực hiện",
        "bo_mang_gen": "Chưa xác định",
        "me_mang_gen": "Chưa xác định",
        "loai_gen": "β-Thalassemia",
    }
    all_q = (
        CAU_HOI_NHOM_A
        + CAU_HOI_NHOM_B
        + CAU_HOI_NHOM_C
        + CAU_HOI_NHOM_D
    )
    for q_id, _ in all_q:
        defaults[q_id] = "Không"
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# 8. Giao diện chính
def main():
    init_state()
    ss = st.session_state

    st.markdown(
        """
    <div class="hero">
        <h1>🩸 HỆ THỐNG SÀNG LỌC & TƯ VẤN DI TRUYỀN THALASSEMIA</h1>
        <p>Quy trình sàng lọc 3 vòng chuẩn hóa • Phân nhóm yếu tố nguy cơ • Mô hình tư vấn Mendel</p>
    </div>
    """,
        unsafe_allow_html=True,
    )

    # Thông tin hành chính
    st.subheader("📋 Thông tin chung (Tuổi, Giới, Dân tộc, Khu vực)")
    c1, c2, c3 = st.columns(3)
    with c1:
        ss.hoten = st.text_input("Họ và tên:", value=ss.hoten)
        ss.gioitinh = st.selectbox(
            "Giới tính:",
            ["Nữ", "Nam"],
            index=0 if ss.gioitinh == "Nữ" else 1,
        )
    with c2:
        ss.ngaysinh = st.date_input(
            "Ngày sinh:", value=ss.ngaysinh, format="DD/MM/YYYY"
        )
        ss.dantoc = st.selectbox(
            "Dân tộc:",
            DAN_TOC,
            index=DAN_TOC.index(ss.dantoc) if ss.dantoc in DAN_TOC else 0,
        )
    with c3:
        ss.vung_o = st.selectbox(
            "Vùng/Miền:",
            VUNG_MIEN,
            index=VUNG_MIEN.index(ss.vung_o) if ss.vung_o in VUNG_MIEN else 0,
        )
        tinh_ds = (
            TAT_CA_TINH
            if ss.vung_o == "Chọn vùng/miền"
            else ["Chọn Tỉnh/Thành phố"] + TINH_THEO_VUNG.get(ss.vung_o, [])
        )
        ss.tinh_o = st.selectbox(
            "Tỉnh/Thành phố:",
            tinh_ds,
            index=tinh_ds.index(ss.tinh_o) if ss.tinh_o in tinh_ds else 0,
        )

    st.markdown("---")

    # ------------------ VÒNG 1 ------------------
    st.subheader("🟢 VÒNG 1: Đánh giá yếu tố nguy cơ")
    st.caption("*Sắp xếp theo các nhóm chỉ số lâm sàng:*")

    score_v1 = DIEM_DAN_TOC_VN.get(ss.dantoc, 0.5)

    def render_group(title, questions, weight=1.0):
        nonlocal score_v1
        st.markdown(f"**{title}**")
        for q_id, q_text in questions:
            cq, ca = st.columns([3.5, 1])
            with cq:
                st.write(f"• {q_text}")
            with ca:
                if (
                    st.radio(
                        q_id,
                        ["Có", "Không"],
                        key=q_id,
                        horizontal=True,
                        label_visibility="collapsed",
                    )
                    == "Có"
                ):
                    score_v1 += weight

    render_group("🔹 Nhóm A — Tiền sử gia đình & Thai sản", CAU_HOI_NHOM_A, 1.5)
    render_group("🔹 Nhóm B — Kết quả xét nghiệm cũ", CAU_HOI_NHOM_B, 1.5)
    render_group("🔹 Nhóm C — Yếu tố dịch tễ & Hôn nhân", CAU_HOI_NHOM_C, 1.0)
    render_group("🔹 Nhóm D — Triệu chứng không đặc hiệu", CAU_HOI_NHOM_D, 0.5)

    ss.s1_score = score_v1

    st.markdown("**Kết quả Vòng 1:**")
    if ss.s1_score >= 4.0:
        box(
            "danger-box",
            f"🔴 <b>KẾT QUẢ: CÓ YẾU TỐ NGUY CƠ ĐÁNG CHÚ Ý ({ss.s1_score:.1f} điểm)</b><br>Ghi nhận các yếu tố nguy cơ cao từ tiền sử gia đình, xét nghiệm cũ hoặc dịch tễ.<br>👉 <b>Khuyến cáo:</b> Bắt buộc thực hiện Vòng 2 (Tổng phân tích tế bào máu).",
        )
    elif ss.s1_score >= 2.0:
        box(
            "warning-box",
            f"🟡 <b>KẾT QUẢ: CÓ YẾU TỐ NGUY CƠ ({ss.s1_score:.1f} điểm)</b><br>Ghi nhận một số yếu tố tiền sử nhẹ hoặc triệu chứng không đặc hiệu.<br>👉 <b>Khuyến cáo:</b> Khuyên thực hiện xét nghiệm công thức máu kiểm tra MCV, MCH, Hb.",
        )
    else:
        box(
            "success-box",
            f"🟢 <b>KẾT QUẢ: YẾU TỐ NGUY CƠ THẤP ({ss.s1_score:.1f} điểm)</b><br>Chưa phát hiện các yếu tố tiền sử dịch tễ đáng lo ngại.",
        )

    st.markdown("---")

    # ------------------ VÒNG 2 ------------------
    st.subheader("🔬 VÒNG 2: Phân tích công thức máu")
    st.caption("*Nhập các chỉ số từ kết quả xét nghiệm máu:*")

    c_hb, c_rbc, c_mcv, c_mch, c_rdw = st.columns(5)
    with c_hb:
        ss.hb = st.number_input("Hb (g/dL):", value=ss.hb, step=0.1)
    with c_rbc:
        ss.rbc = st.number_input("RBC (M/uL):", value=ss.rbc, step=0.01)
    with c_mcv:
        ss.mcv = st.number_input("MCV (fL):", value=ss.mcv, step=0.1)
    with c_mch:
        ss.mch = st.number_input("MCH (pg):", value=ss.mch, step=0.1)
    with c_rdw:
        ss.rdw = st.number_input("RDW (%):", value=ss.rdw, step=0.1)

    if ss.mcv > 0 and ss.rbc > 0:
        morphology, chromic, mentzer, diff_list = analyze_round2(
            ss.mcv, ss.mch, ss.hb, ss.rbc, ss.rdw, ss.gioitinh
        )

        st.markdown("**Kết quả phân tích Vòng 2:**")
        st.write(f"• **Phân loại hồng cầu:** {morphology} | {chromic}")
        st.write(
            f"• **Mentzer Index (MCV / RBC):** {mentzer:.2f} "
            + ("(< 13)" if mentzer < 13 else "(≥ 13)")
        )

        st.markdown("**Phân biệt định hướng:**")
        for d in diff_list:
            st.write(f"- {d}")

        if ss.mcv < 85.0 or ss.mch < 28.0:
            box(
                "warning-box",
                "📊 <b>Kết quả Vòng 2: MẪU HÌNH HUYẾT HỌC GỢI Ý THALASSEMIA</b><br><i>(Lưu ý: Đây là định hướng huyết học sàng lọc, không phải kết luận 'Bạn bị Thalassemia').</i><br>👉 <b>Khuyến cáo:</b> Cần thực hiện tiếp Vòng 3 (Xét nghiệm chuyên sâu).",
            )
        else:
            box(
                "success-box",
                "🟢 <b>Kết quả Vòng 2: MẪU HÌNH HUYẾT HỌC TRONG GIỚI HẠN BÌNH THƯỜNG</b>",
            )

    st.markdown("---")

    # ------------------ VÒNG 3 ------------------
    st.subheader("🧬 VÒNG 3: Xét nghiệm chuyên sâu")
    st.caption("*Hb electrophoresis / HPLC & Xét nghiệm Gen α-globin / β-globin*")

    c_v3_1, c_v3_2 = st.columns(2)
    with c_v3_1:
        ss.dien_di = st.selectbox(
            "Kết quả Điện di Hb (HbA2, HbF, HbE...):",
            [
                "Chưa thực hiện",
                "HbA2 / HbF bình thường (HbA2 2.0-3.5%)",
                "HbA2 tăng (> 3.5%)",
                "Xuất hiện băng HbE",
                "Xuất hiện băng HbH / Bart's",
            ],
        )
    with c_v3_2:
        ss.gen_test = st.selectbox(
            "Kết quả Gen α-globin / β-globin:",
            [
                "Chưa thực hiện",
                "Không phát hiện đột biến",
                "Phát hiện 1 đột biến dị hợp (Mang gen ẩn)",
                "Phát hiện đột biến đồng hợp / Tạp dị hợp (Bệnh thể nặng/trung bình)",
            ],
        )

    st.markdown("**Kết luận Vòng 3:**")
    if (
        ss.gen_test
        == "Phát hiện đột biến đồng hợp / Tạp dị hợp (Bệnh thể nặng/trung bình)"
    ):
        box(
            "danger-box",
            "🔴 <b>ĐÃ XÁC ĐỊNH BIẾN THỂ GEN / THỂ BỆNH</b><br>Cần theo dõi và quản lý bởi bác sĩ chuyên khoa Huyết học.",
        )
    elif (
        ss.gen_test == "Phát hiện 1 đột biến dị hợp (Mang gen ẩn)"
        or "tăng" in ss.dien_di
        or "Xuất hiện" in ss.dien_di
    ):
        box(
            "warning-box",
            "🟡 <b>KẾT QUẢ GỢI Ý NGƯỜI MANG GEN (TRAIT / CARRIER)</b><br>Người mang gen ẩn khỏe mạnh bình thường nhưng cần tư vấn di truyền trước khi sinh con.",
        )
    elif (
        ss.gen_test == "Không phát hiện đột biến"
        and "bình thường" in ss.dien_di
    ):
        box(
            "success-box",
            "🟢 <b>KHÔNG GHI NHẬN BẤT THƯỜNG TRONG XÉT NGHIỆM ĐÃ THỰC HIỆN</b>",
        )
    else:
        box(
            "info-box",
            "🟠 <b>CẦN ĐÁNH GIÁ CHUYÊN KHOA</b> (Chưa đủ thông tin khẳng định hoặc xét nghiệm đang thực hiện).",
        )

    st.markdown("---")

    # ------------------ TƯ VẤN DI TRUYỀN ------------------
    st.subheader("💡 TƯ VẤN DI TRUYỀN (Theo mô hình Mendel)")
    st.caption("*Tư vấn nguy cơ di truyền cho thế hệ sau*")

    c_parent1, c_parent2, c_gen_type = st.columns(3)
    with c_parent1:
        ss.bo_mang_gen = st.selectbox(
            "Tình trạng Bố:",
            ["Chưa xác định", "Bình thường (Không mang gen)", "Người mang gen (Trait)"],
            key="sb_bo",
        )
    with c_parent2:
        ss.me_mang_gen = st.selectbox(
            "Tình trạng Mẹ:",
            ["Chưa xác định", "Bình thường (Không mang gen)", "Người mang gen (Trait)"],
            key="sb_me",
        )
    with c_gen_type:
        ss.loai_gen = st.selectbox(
            "Bệnh lý xét đến:", ["β-Thalassemia", "α-Thalassemia"]
        )

    if (
        ss.bo_mang_gen == "Người mang gen (Trait)"
        and ss.me_mang_gen == "Người mang gen (Trait)"
    ):
        box(
            "danger-box",
            f"<b>🧬 TƯ VẤN NGUY CƠ DI TRUYỀN ({ss.loai_gen}):</b><br>"
            f"Trong <b>mỗi lần mang thai</b>, nếu cả hai bố mẹ đều là người mang biến thể {ss.loai_gen} phù hợp, nguy cơ theo mô hình di truyền Mendel là:<br>"
            f"• <b>25%</b> nguy cơ con mắc bệnh thể nặng.<br>"
            f"• <b>50%</b> nguy cơ con là người mang gen (khỏe mạnh).<br>"
            f"• <b>25%</b> nguy cơ con hoàn toàn không mang gen bệnh.<br>"
            f"👉 <i>Khuyên thực hiện chẩn đoán trước sinh (chọc ối / sinh thiết gai nhau) khi mang thai.</i>",
        )
    elif (
        ss.bo_mang_gen == "Người mang gen (Trait)"
        or ss.me_mang_gen == "Người mang gen (Trait)"
    ) and (
        ss.bo_mang_gen == "Bình thường (Không mang gen)"
        or ss.me_mang_gen == "Bình thường (Không mang gen)"
    ):
        box(
            "info-box",
            f"<b>🧬 TƯ VẤN NGUY CƠ DI TRUYỀN ({ss.loai_gen}):</b><br>"
            f"Trong <b>mỗi lần mang thai</b>, khi một trong hai bố mẹ là người mang biến thể {ss.loai_gen}, nguy cơ theo mô hình di truyền Mendel là:<br>"
            f"• <b>50%</b> nguy cơ con là người mang gen (khỏe mạnh).<br>"
            f"• <b>50%</b> nguy cơ con hoàn toàn không mang gen bệnh.<br>"
            f"• <b>0%</b> nguy cơ con mắc bệnh thể nặng.",
        )
    else:
        st.info(
            "Vui lòng chọn đầy đủ tình trạng mang gen của Bố và Mẹ để hệ thống đưa ra mô hình tư vấn Mendel chính xác."
        )

    if st.button("💾 Lưu dữ liệu hồ sơ"):
        data = {
            "HoTen": ss.hoten,
            "GioiTinh": ss.gioitinh,
            "NgaySinh": ss.ngaysinh.strftime("%d/%m/%Y"),
            "DanToc": ss.dantoc,
            "TinhO": ss.tinh_o,
            "DiemV1": ss.s1_score,
            "MCV": ss.mcv,
            "MCH": ss.mch,
            "Hb": ss.hb,
            "RBC": ss.rbc,
            "RDW": ss.rdw,
            "DienDi": ss.dien_di,
            "GenTest": ss.gen_test,
            "BoMangGen": ss.bo_mang_gen,
            "MeMangGen": ss.me_mang_gen,
        }
        if ghi_du_lieu(data):
            st.success("Lưu dữ liệu hồ sơ thành công!")


if __name__ == "__main__":
    main()
