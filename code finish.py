import base64
import io
import json
import os
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

DO_CAO_SONG = [
    "Dưới 1000m (Đồng bằng / Trung du)",
    "1000m - 1499m (Núi vừa)",
    "1500m - 1999m (Núi cao - Đà Lạt, Sa Pa...)",
    "Từ 2000m trở lên (Núi rất cao)",
]

MOI_TRUONG_LAM_VIEC = [
    "Văn phòng / Học tập trong nhà",
    "Lao động ngoài trời / Nông - Lâm - Ngư nghiệp",
    "Nhà máy / Tiếp xúc hóa chất, kim loại nặng, chì",
    "Môi trường hầm mỏ / Thiếu oxy kéo dài",
]

# 4. Câu hỏi Vòng 1
CAU_HOI_NHOM_A = [
    (
        "q_a1",
        "Trong gia đình hoặc dòng họ, đã có ai từng được chẩn đoán mắc bệnh Thalassemia hoặc phải đi truyền máu định kỳ dài hạn chưa?",
    ),
    (
        "q_a2",
        "Trong họ hàng có ai ghi nhận các dấu hiệu bất thường như: lách to, biến dạng xương mặt hoặc da sạm xám/vàng da kéo dài không rõ nguyên nhân?",
    ),
    (
        "q_a3",
        "Tiền sử thai sản trong gia đình hoặc bản thân từng có ghi nhận các trường hợp sảy thai liên tiếp, thai chết lưu không rõ nguyên nhân, hoặc thai nhi bị phù thai?",
    ),
]

CAU_HOI_NHOM_B = [
    (
        "q_b1",
        "Trong các lần khám sức khỏe trước đây, bản thân bạn đã bao giờ được thông báo hoặc ghi nhận chỉ số thiếu máu, hồng cầu nhỏ nhược sắc chưa?",
    ),
    (
        "q_b2",
        "Bản thân đã từng làm xét nghiệm Điện di Hb hoặc Xét nghiệm Gen và nhận kết quả nghi ngờ/xác định mang gen ẩn Thalassemia hoặc các biến thể Hb?",
    ),
    (
        "q_b3",
        "Bạn đã từng đi hiến máu nhân đạo nhưng bị bác sĩ từ chối tiếp nhận vì lý do nồng độ Huyết sắc tố (Hb) quá thấp chưa?",
    ),
]

CAU_HOI_NHOM_C = [
    (
        "q_c1",
        "Bản thân bạn hoặc bố/mẹ thuộc các dân tộc thiểu số tại Việt Nam có tỷ lệ mang gen Thalassemia cao (Thái, Mường, Tày, Nùng, Ê Đê...)?",
    ),
    (
        "q_c2",
        "Vợ/chồng hoặc bạn đời dự định kết hôn của bạn có cùng dòng họ, cùng dân tộc thiểu số, hoặc sinh sống cùng trong một thôn/bản/xã có tính chất khép kín?",
    ),
]

CAU_HOI_NHOM_D = [
    (
        "q_d1",
        "Bạn có thường xuyên xuất hiện cảm giác mệt mỏi mạn tính, hoa mắt, chóng mặt, thể lực suy giảm kéo dài không?",
    ),
    (
        "q_d2",
        "Bản thân bạn có nhận thấy da xanh xao, niêm mạc mắt/môi nhợt nhạt dai dẳng dù ăn uống đầy đủ dinh dưỡng không?",
    ),
]


# 5. Hàm xuất file Word từng Vòng
def export_docx_vong(vong_num, title, ss, content_dict):
    doc = Document()
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(f"KẾT QUẢ SÀNG LỌC THALASSEMIA - VÒNG {vong_num}")
    run_title.font.size = Pt(16)
    run_title.font.bold = True
    run_title.font.name = "Arial"

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = p_sub.add_run(f"Hạng mục: {title}")
    run_sub.font.size = Pt(13)
    run_sub.font.italic = True
    run_sub.font.name = "Arial"

    doc.add_paragraph("--------------------------------------------------------------------------------")

    p_info = doc.add_paragraph()
    p_info.add_run("1. THÔNG TIN HÀNH CHÍNH & DỊCH TỄ:\n").bold = True
    p_info.add_run(f"• Họ và tên: {ss.hoten if ss.hoten else 'N/A'}\n")
    p_info.add_run(f"• Giới tính: {ss.gioitinh} | Ngày sinh: {ss.ngaysinh.strftime('%d/%m/%Y')}\n")
    p_info.add_run(f"• Dân tộc: {ss.dantoc} | Tỉnh/Thành: {ss.tinh_o}\n")
    p_info.add_run(f"• Độ cao sinh sống: {ss.do_cao}\n")
    p_info.add_run(f"• Môi trường làm việc: {ss.moi_truong}\n")

    doc.add_paragraph("--------------------------------------------------------------------------------")

    p_detail = doc.add_paragraph()
    p_detail.add_run(f"2. KẾT QUẢ ĐÁNH GIÁ VÒNG {vong_num}:\n").bold = True
    for key, val in content_dict.items():
        p_detail.add_run(f"• {key}: {val}\n")

    p_footer = doc.add_paragraph()
    p_footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_footer.add_run(f"\nNgày xuất báo cáo: {datetime.now().strftime('%d/%m/%Y %H:%M')}\n").italic = True

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


# 6. Database SQLite
DATA_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "thalalassemia_v5.db"
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


# 7. Phân tích Vòng 2 (Tính cả Hb hiệu chỉnh theo Độ cao)
def analyze_round2(mcv, mch, hb_raw, rbc, rdw, gioitinh, do_cao):
    # Tính toán mức trừ Hb do độ cao
    hb_adj_val = 0.0
    if "1000m - 1499m" in do_cao:
        hb_adj_val = 0.2
    elif "1500m - 1999m" in do_cao:
        hb_adj_val = 0.5
    elif "Từ 2000m trở lên" in do_cao:
        hb_adj_val = 1.2

    hb_eff = max(0.0, hb_raw - hb_adj_val)

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
            differential.append("Cần kiểm tra thêm Ferritin huyết thanh.")

        if rdw > 15.0:
            differential.append(
                "RDW > 15%: Kích thước hồng cầu không đều (gợi ý thiếu sắt tiến triển)."
            )
        else:
            differential.append(
                "RDW bình thường/tăng nhẹ: Thường gặp trong thể mang gen Thalassemia."
            )
    elif mcv > 100.0:
        differential.append(
            "Gợi ý nguyên nhân khác: Thiếu Vitamin B12, Folic Acid hoặc bệnh lý gan."
        )
    else:
        differential.append(
            "Các chỉ số thể tích và Hb hồng cầu nằm trong giới hạn bình thường."
        )

    return morphology, chromic, mentzer, differential, hb_eff, hb_adj_val


# 8. Init State
def init_state():
    defaults = {
        "ho_so_id": "",
        "hoten": "",
        "gioitinh": "Nữ",
        "ngaysinh": date(2000, 1, 1),
        "dantoc": "Chọn dân tộc",
        "vung_o": "Chọn vùng/miền",
        "tinh_o": "Chọn Tỉnh/Thành phố",
        "do_cao": DO_CAO_SONG[0],
        "moi_truong": MOI_TRUONG_LAM_VIEC[0],
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


# 9. Main App
def main():
    init_state()
    ss = st.session_state

    st.markdown(
        """
    <div class="hero">
        <h1>🩸 HỆ THỐNG SÀNG LỌC & TƯ VẤN DI TRUYỀN THALASSEMIA</h1>
        <p>Quản lý hồ sơ theo từng Vòng • Tự động hiệu chỉnh Hb theo độ cao • Xuất Word từng công đoạn</p>
    </div>
    """,
        unsafe_allow_html=True,
    )

    # Thông tin hành chính
    st.subheader(
        "📋 Thông tin chung (Tuổi, Giới, Dân tộc, Địa lý & Môi trường)"
    )

    r1_1, r1_2, r1_3 = st.columns(3)
    with r1_1:
        ss.hoten = st.text_input("Họ và tên:", value=ss.hoten)
        ss.gioitinh = st.selectbox(
            "Giới tính:",
            ["Nữ", "Nam"],
            index=0 if ss.gioitinh == "Nữ" else 1,
        )
    with r1_2:
        ss.ngaysinh = st.date_input(
            "Ngày sinh:", value=ss.ngaysinh, format="DD/MM/YYYY"
        )
        ss.dantoc = st.selectbox(
            "Dân tộc:",
            DAN_TOC,
            index=DAN_TOC.index(ss.dantoc) if ss.dantoc in DAN_TOC else 0,
        )
    with r1_3:
        ss.vung_o = st.selectbox(
            "Vùng/Miền cư trú:",
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

    r2_1, r2_2 = st.columns(2)
    with r2_1:
        ss.do_cao = st.selectbox(
            "⛰️ Độ cao khu vực sinh sống (so với mực nước biển):",
            DO_CAO_SONG,
            index=DO_CAO_SONG.index(ss.do_cao)
            if ss.do_cao in DO_CAO_SONG
            else 0,
        )
    with r2_2:
        ss.moi_truong = st.selectbox(
            "🏭 Môi trường làm việc / Học tập chính:",
            MOI_TRUONG_LAM_VIEC,
            index=MOI_TRUONG_LAM_VIEC.index(ss.moi_truong)
            if ss.moi_truong in MOI_TRUONG_LAM_VIEC
            else 0,
        )

    st.markdown("---")

    # ------------------ VÒNG 1 ------------------
    st.subheader("🟢 VÒNG 1: Đánh giá yếu tố nguy cơ")

    score_v1 = DIEM_DAN_TOC_VN.get(ss.dantoc, 0.5)

    def render_group(title, questions, weight=1.0):
        nonlocal score_v1
        st.markdown(f"### {title}")
        for q_id, q_text in questions:
            cq, ca = st.columns([3.8, 1.2])
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

    render_group("📌 Nhóm A — Tiền sử gia đình & Thai sản", CAU_HOI_NHOM_A, 1.5)
    render_group("📌 Nhóm B — Kết quả xét nghiệm cũ & Tiền sử hiến máu", CAU_HOI_NHOM_B, 1.5)
    render_group("📌 Nhóm C — Yếu tố dịch tễ dân tộc & Hôn nhân", CAU_HOI_NHOM_C, 1.0)
    render_group("📌 Nhóm D — Triệu chứng lâm sàng không đặc hiệu", CAU_HOI_NHOM_D, 0.5)

    ss.s1_score = score_v1

    v1_label = "Nguy cơ thấp"
    if ss.s1_score >= 4.0:
        v1_label = "Có yếu tố nguy cơ đáng chú ý"
        box(
            "danger-box",
            f"🔴 <b>KẾT QUẢ VÒNG 1: {v1_label.upper()} ({ss.s1_score:.1f} điểm)</b><br>👉 Khuyên thực hiện Vòng 2 (Tổng phân tích tế bào máu).",
        )
    elif ss.s1_score >= 2.0:
        v1_label = "Có yếu tố nguy cơ"
        box(
            "warning-box",
            f"🟡 <b>KẾT QUẢ VÒNG 1: {v1_label.upper()} ({ss.s1_score:.1f} điểm)</b><br>👉 Khuyên xét nghiệm công thức máu kiểm tra MCV, MCH, Hb.",
        )
    else:
        box(
            "success-box",
            f"🟢 <b>KẾT QUẢ VÒNG 1: YẾU TỐ NGUY CƠ THẤP ({ss.s1_score:.1f} điểm)</b>",
        )

    # Báo cáo Vòng 1
    col_v1_save, col_v1_exp = st.columns(2)
    with col_v1_save:
        if st.button("💾 Lưu kết quả Vòng 1 vào CSDL"):
            data_v1 = {
                "HoTen": ss.hoten,
                "GioiTinh": ss.gioitinh,
                "NgaySinh": ss.ngaysinh.strftime("%d/%m/%Y"),
                "Vong1_Diem": ss.s1_score,
                "Vong1_KetLuan": v1_label,
            }
            if ghi_du_lieu(data_v1):
                st.success("Đã lưu dữ liệu Vòng 1!")
    with col_v1_exp:
        docx_v1 = export_docx_vong(
            1,
            "Đánh giá yếu tố nguy cơ",
            ss,
            {"Tổng điểm nguy cơ": f"{ss.s1_score:.1f}", "Xếp loại nguy cơ": v1_label},
        )
        st.download_button(
            label="📄 Xuất Báo Cáo Word (Vòng 1)",
            data=docx_v1,
            file_name=f"Vong1_SangLoc_{ss.hoten if ss.hoten else 'BenhNhan'}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    st.markdown("---")

    # ------------------ VÒNG 2 ------------------
    st.subheader("🔬 VÒNG 2: Phân tích công thức máu")
    c_hb, c_rbc, c_mcv, c_mch, c_rdw = st.columns(5)
    with c_hb:
        ss.hb = st.number_input("Hb thực tế (g/dL):", value=ss.hb, step=0.1)
    with c_rbc:
        ss.rbc = st.number_input("RBC (M/uL):", value=ss.rbc, step=0.01)
    with c_mcv:
        ss.mcv = st.number_input("MCV (fL):", value=ss.mcv, step=0.1)
    with c_mch:
        ss.mch = st.number_input("MCH (pg):", value=ss.mch, step=0.1)
    with c_rdw:
        ss.rdw = st.number_input("RDW (%):", value=ss.rdw, step=0.1)

    v2_label = "Chưa hoàn thành"
    diff_str = "Chưa có dữ liệu"

    if ss.mcv > 0 and ss.rbc > 0:
        morphology, chromic, mentzer, diff_list, hb_eff, hb_adj_val = analyze_round2(
            ss.mcv, ss.mch, ss.hb, ss.rbc, ss.rdw, ss.gioitinh, ss.do_cao
        )
        diff_str = " | ".join(diff_list)

        # Hiển thị thông tin điều chỉnh Hb theo độ cao
        if hb_adj_val > 0:
            st.info(
                f"⛰️ **Điều chỉnh Hb sinh lý theo độ cao:** Khu vực **{ss.do_cao}** làm tăng Hb sinh lý. "
                f"Chỉ số Hb đo được là **{ss.hb} g/dL**, sau khi trừ hiệu chỉnh (-{hb_adj_val} g/dL) để về chuẩn mực nước biển là **{hb_eff:.1f} g/dL**."
            )
        else:
            st.caption("ℹ️ Độ cao sinh sống < 1000m: Giữ nguyên chỉ số Hb đo thực tế.")

        st.write(f"• **Phân loại hình thái:** {morphology} | {chromic}")
        st.write(f"• **Mentzer Index:** {mentzer:.2f}")

        if ss.mcv < 85.0 or ss.mch < 28.0:
            v2_label = "Mẫu hình huyết học gợi ý Thalassemia"
            box(
                "warning-box",
                f"📊 <b>Kết quả Vòng 2: {v2_label.upper()}</b><br><i>(Lưu ý: Đây là định hướng huyết học sàng lọc, không phải kết luận 'Bạn bị Thalassemia').</i><br>👉 <b>Khuyến cáo:</b> Thực hiện Vòng 3.",
            )
        else:
            v2_label = "Mẫu hình huyết học trong giới hạn bình thường"
            box(
                "success-box",
                f"🟢 <b>Kết quả Vòng 2: {v2_label.upper()}</b>",
            )

        col_v2_save, col_v2_exp = st.columns(2)
        with col_v2_save:
            if st.button("💾 Lưu kết quả Vòng 2 vào CSDL"):
                data_v2 = {
                    "HoTen": ss.hoten,
                    "MCV": ss.mcv,
                    "MCH": ss.mch,
                    "Hb_ThucTe": ss.hb,
                    "Hb_HieuChinh": hb_eff,
                    "RBC": ss.rbc,
                    "RDW": ss.rdw,
                    "Mentzer": mentzer,
                    "Vong2_KetLuan": v2_label,
                }
                if ghi_du_lieu(data_v2):
                    st.success("Đã lưu dữ liệu Vòng 2!")
        with col_v2_exp:
            docx_v2 = export_docx_vong(
                2,
                "Phân tích công thức máu",
                ss,
                {
                    "Chỉ số CBC thực tế": f"Hb: {ss.hb} g/dL | RBC: {ss.rbc} | MCV: {ss.mcv} | MCH: {ss.mch} | RDW: {ss.rdw}",
                    "Hb hiệu chỉnh (theo độ cao)": f"{hb_eff:.1f} g/dL (đã trừ {hb_adj_val} g/dL)",
                    "Mentzer Index": f"{mentzer:.2f}",
                    "Định hướng": diff_str,
                    "Kết luận Vòng 2": v2_label,
                },
            )
            st.download_button(
                label="📄 Xuất Báo Cáo Word (Vòng 2)",
                data=docx_v2,
                file_name=f"Vong2_CongThucMau_{ss.hoten if ss.hoten else 'BenhNhan'}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

    st.markdown("---")

    # ------------------ VÒNG 3 ------------------
    st.subheader("🧬 VÒNG 3: Xét nghiệm chuyên sâu")
    c_v3_1, c_v3_2 = st.columns(2)
    with c_v3_1:
        ss.dien_di = st.selectbox(
            "Kết quả Điện di Hb:",
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

    v3_label = "Chưa kết luận"
    if (
        ss.gen_test
        == "Phát hiện đột biến đồng hợp / Tạp dị hợp (Bệnh thể nặng/trung bình)"
    ):
        v3_label = "Đã xác định biến thể gen / Thể bệnh"
        box(
            "danger-box",
            f"🔴 <b>{v3_label.upper()}</b><br>Cần quản lý bởi bác sĩ chuyên khoa Huyết học.",
        )
    elif (
        ss.gen_test == "Phát hiện 1 đột biến dị hợp (Mang gen ẩn)"
        or "tăng" in ss.dien_di
        or "Xuất hiện" in ss.dien_di
    ):
        v3_label = "Kết quả gợi ý người mang gen (Trait / Carrier)"
        box(
            "warning-box",
            f"🟡 <b>{v3_label.upper()}</b><br>Cần tư vấn di truyền trước khi sinh con.",
        )
    elif (
        ss.gen_test == "Không phát hiện đột biến"
        and "bình thường" in ss.dien_di
    ):
        v3_label = "Không ghi nhận bất thường trong xét nghiệm đã thực hiện"
        box("success-box", f"🟢 <b>{v3_label.upper()}</b>")
    else:
        v3_label = "Cần đánh giá chuyên khoa"
        box("info-box", f"🟠 <b>{v3_label.upper()}</b>")

    col_v3_save, col_v3_exp = st.columns(2)
    with col_v3_save:
        if st.button("💾 Lưu kết quả Vòng 3 vào CSDL"):
            data_v3 = {
                "HoTen": ss.hoten,
                "DienDi": ss.dien_di,
                "GenTest": ss.gen_test,
                "Vong3_KetLuan": v3_label,
            }
            if ghi_du_lieu(data_v3):
                st.success("Đã lưu dữ liệu Vòng 3!")
    with col_v3_exp:
        docx_v3 = export_docx_vong(
            3,
            "Xét nghiệm chuyên sâu & Gen",
            ss,
            {
                "Điện di Hb": ss.dien_di,
                "Xét nghiệm Gen": ss.gen_test,
                "Kết luận Vòng 3": v3_label,
            },
        )
        st.download_button(
            label="📄 Xuất Báo Cáo Word (Vòng 3)",
            data=docx_v3,
            file_name=f"Vong3_ChuyenSau_{ss.hoten if ss.hoten else 'BenhNhan'}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    st.markdown("---")

    # ------------------ TƯ VẤN DI TRUYỀN ------------------
    st.subheader("💡 TƯ VẤN DI TRUYỀN (Theo mô hình Mendel)")
    c_p1, c_p2, c_gt = st.columns(3)
    with c_p1:
        ss.bo_mang_gen = st.selectbox(
            "Tình trạng Bố:",
            ["Chưa xác định", "Bình thường (Không mang gen)", "Người mang gen (Trait)"],
            key="sb_bo",
        )
    with c_p2:
        ss.me_mang_gen = st.selectbox(
            "Tình trạng Mẹ:",
            ["Chưa xác định", "Bình thường (Không mang gen)", "Người mang gen (Trait)"],
            key="sb_me",
        )
    with c_gt:
        ss.loai_gen = st.selectbox(
            "Bệnh lý xét đến:", ["β-Thalassemia", "α-Thalassemia"]
        )

    tu_van_str = "Chưa đủ dữ liệu tư vấn"
    if (
        ss.bo_mang_gen == "Người mang gen (Trait)"
        and ss.me_mang_gen == "Người mang gen (Trait)"
    ):
        tu_van_str = (
            f"Trong mỗi lần mang thai, nguy cơ theo mô hình Mendel là: "
            f"25% bệnh thể nặng, 50% mang gen (khỏe mạnh), 25% hoàn toàn không mang gen."
        )
        box(
            "danger-box",
            f"<b>🧬 TƯ VẤN NGUY CƠ DI TRUYỀN ({ss.loai_gen}):</b><br>"
            f"Trong <b>mỗi lần mang thai</b>, nếu cả hai bố mẹ đều là người mang biến thể {ss.loai_gen} phù hợp, nguy cơ theo mô hình di truyền Mendel là:<br>"
            f"• <b>25%</b> nguy cơ con mắc bệnh thể nặng.<br>"
            f"• <b>50%</b> nguy cơ con là người mang gen (khỏe mạnh).<br>"
            f"• <b>25%</b> nguy cơ con hoàn toàn không mang gen bệnh.<br>"
            f"👉 <i>Khuyên thực hiện chẩn đoán trước sinh (chọc ối / sinh thiết gai nhau).</i>",
        )
    elif (
        ss.bo_mang_gen == "Người mang gen (Trait)"
        or ss.me_mang_gen == "Người mang gen (Trait)"
    ) and (
        ss.bo_mang_gen == "Bình thường (Không mang gen)"
        or ss.me_mang_gen == "Bình thường (Không mang gen)"
    ):
        tu_van_str = (
            f"Trong mỗi lần mang thai, nguy cơ theo mô hình Mendel là: "
            f"50% mang gen (khỏe mạnh), 50% hoàn toàn không mang gen, 0% thể nặng."
        )
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

    col_tv_save, col_tv_exp = st.columns(2)
    with col_tv_save:
        if st.button("💾 Lưu tư vấn di truyền vào CSDL"):
            data_tv = {
                "HoTen": ss.hoten,
                "BoMangGen": ss.bo_mang_gen,
                "MeMangGen": ss.me_mang_gen,
                "LoaiGen": ss.loai_gen,
                "NoiDungTuVan": tu_van_str,
            }
            if ghi_du_lieu(data_tv):
                st.success("Đã lưu dữ liệu Tư vấn di truyền!")
    with col_tv_exp:
        docx_tv = export_docx_vong(
            4,
            "Tư vấn di truyền Mendel",
            ss,
            {
                "Tình trạng Bố": ss.bo_mang_gen,
                "Tình trạng Mẹ": ss.me_mang_gen,
                "Loại gen": ss.loai_gen,
                "Dự báo nguy cơ con": tu_van_str,
            },
        )
        st.download_button(
            label="📄 Xuất Báo Cáo Word (Tư Vấn Di Truyền)",
            data=docx_tv,
            file_name=f"TuVanDiTruyen_{ss.hoten if ss.hoten else 'BenhNhan'}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )


if __name__ == "__main__":
    main()
