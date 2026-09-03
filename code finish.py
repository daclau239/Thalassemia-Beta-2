import base64
import io
import os
import json
from pathlib import Path
from datetime import date, datetime
import uuid
import sqlite3

import streamlit as st
import pandas as pd

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# Mật khẩu quản trị viên
ADMIN_PASSWORD = "admin123"

# 1. Cấu hình trang
st.set_page_config(
    page_title="Sàng lọc Thalassemia & Bệnh lý Huyết học",
    page_icon="🩸",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Giao diện CSS
CSS = """<style>
.stApp { background: linear-gradient(180deg, #E6F4F8 0%, #F4F8FB 100%); }
.block-container { padding-top: 1.2rem; padding-bottom: 3.5rem; max-width: 1180px; }
header, footer, div[data-testid="stDecoration"] { display: none; }
.hero { background: linear-gradient(135deg, #012A4A 0%, #014F86 50%, #0077B6 100%);
        color: #fff; border-radius: 18px; padding: 22px 26px; margin-bottom: 16px; }
.hero h1 { color: #fff !important; font-size: 26px !important; margin: 0 0 6px 0 !important; }
.hero p { color: #D9F3FF; font-size: 15px; margin: 0; }
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

# 3. Danh mục dữ liệu chuẩn người Việt
VUNG_MIEN = [
    "Chọn vùng/miền", "Đông Bắc", "Tây Bắc", "Đồng bằng sông Hồng", "Bắc Trung Bộ",
    "Trung Trung Bộ", "Nam Trung Bộ", "Tây Nguyên", "Đồng bằng sông Cửu Long", "Đồng bằng Sông Cửu Long"
]

TINH_THEO_VUNG = {
    "Đông Bắc": ["Hà Giang", "Cao Bằng", "Bắc Kạn - Thái Nguyên", "Tuyên Quang", "Lạng Sơn", "Bắc Giang"],
    "Tây Bắc": ["Điện Biên - Lai Châu", "Sơn La", "Hòa Bình", "Yên Bái"],
    "Đồng bằng sông Hồng": ["Hà Nội", "Hải Phòng", "Vĩnh Phúc - Phú Thọ", "Bắc Ninh - Hưng Yên", "Quảng Ninh", "Hải Dương", "Thái Bình", "Nam Định - Ninh Bình", "Thanh Hóa"],
    "Bắc Trung Bộ": ["Nghệ An", "Hà Tĩnh", "Quảng Bình", "Quảng Trị - Thừa Thiên Huế"],
    "Trung Trung Bộ": ["Đà Nẵng", "Quảng Nam - Quảng Ngãi", "Bình Định", "Phú Yên"],
    "Nam Trung Bộ": ["Khánh Hòa", "Ninh Thuận - Bình Thuận"],
    "Tây Nguyên": ["Kon Tum", "Gia Lai", "Đắc Lắc", "Đắc Nông - Lâm Đồng"],
    "Đồng bằng sông Cửu Long": ["Tiền Giang - Vĩnh Long", "Bến Tre - Trà Vinh", "Đồng Tháp - An Giang", "Kiên Giang - Hậu Giang", "Cần Thơ", "Sóc Trăng - Bạc Liêu - Cà Mau"],
}

TAT_CA_TINH = ["Chọn Tỉnh/Thành phố"]
for ds in TINH_THEO_VUNG.values(): TAT_CA_TINH.extend(ds)

DAN_TOC = [
    "Chọn dân tộc", "Stiêng", "Ê Đê", "Gia Rai", "Ba Na", "Xơ Đăng", "Cơ Ho", "Hrê",
    "Chăm", "Khơ Me", "Thái", "Mường", "Tày", "Nùng", "Dao", "Sán Chay", "Kinh", "Hoa", "Dân tộc khác"
]
DIEM_DAN_TOC_VN = {
    "Stiêng": 3.0, "Ê Đê": 3.0, "Gia Rai": 3.0, "Ba Na": 2.5, "Xơ Đăng": 2.5, "Cơ Ho": 2.5, "Hrê": 2.5,
    "Chăm": 2.0, "Khơ Me": 2.0, "Thái": 2.0, "Mường": 2.0, "Tày": 1.5, "Nùng": 1.5,
    "Dao": 1.5, "Sán Chay": 1.5, "Kinh": 0.5, "Hoa": 0.5, "Dân tộc khác": 0.5,
}

ALTITUDE_OPTIONS = [
    "Dưới 1.000m (Đồng bằng / Trung duy / Núi thấp)", "1.000m – 1.499m (-0.8 g/dL)",
    "1.500m – 1.999m (-1.1 g/dL)", "2.000m – 2.499m (-1.4 g/dL)", "2.500m – 2.999m (-1.8 g/dL)",
    "3.000m – 3.499m (-2.1 g/dL)", "3.500m – 3.999m (-2.5 g/dL)", "4.000m – 4.499m (-2.9 g/dL)", "Từ 4.500m trở lên (-3.3 g/dL)"
]
ALTITUDE_CORRECTION_MAP = {
    "Dưới 1.000m (Đồng bằng / Trung duy / Núi thấp)": 0.0, "1.000m – 1.499m (-0.8 g/dL)": 0.8,
    "1.500m – 1.999m (-1.1 g/dL)": 1.1, "2.000m – 2.499m (-1.4 g/dL)": 1.4, "2.500m – 2.999m (-1.8 g/dL)": 1.8,
    "3.000m – 3.499m (-2.1 g/dL)": 2.1, "3.500m – 3.999m (-2.5 g/dL)": 2.5, "4.000m – 4.499m (-2.9 g/dL)": 2.9, "Từ 4.500m trở lên (-3.3 g/dL)": 3.3
}

CAU_HOI = [
    "1. Bản thân hoặc người thân trong gia đình đã từng được chẩn đoán mắc bệnh Thalassemia / Thiếu máu truyền máu chưa?",
    "2. Gia đình/dòng họ có ai bị biến dạng xương mặt, lách to, da vàng hoặc da sạm đen bất thường không?",
    "3. Bản thân đã từng có tiền sử xét nghiệm ghi nhận thiếu máu nhược sắc, hồng cầu nhỏ hoặc nghi ngờ mang gen chưa?",
    "4. Vợ/Chồng hoặc người chuẩn bị kết hôn có thuộc cùng dòng họ, cùng dân tộc ít người hoặc cùng thôn/bản không?",
    "5. Vợ/Chồng hoặc bạn đời đã từng được xét nghiệm sàng lọc Thalassemia chưa?",
    "6. Đã từng bị từ chối hiến máu do nồng độ Huyết sắc tố (Hb) thấp hoặc hồng cầu quá nhỏ chưa?",
    "7. Bản thân thường xuyên có biểu hiện mệt mỏi, hoa mắt, chóng mặt, da xanh xao kéo dài không?",
    "8. Đã từng thực hiện xét nghiệm Điện di Huyết sắc tố (Hb electrophoresis) hoặc xét nghiệm Gen Thalassemia chưa?",
    "9. Kết quả xét nghiệm trước đây có ghi nhận biến thể Hb (như HbE, HbCS) hoặc mang gen dị hợp tử không?",
    "10. Đã từng được bác sĩ chuyên khoa huyết học tư vấn về nguy cơ sinh con mắc bệnh Thalassemia thể nặng chưa?",
    "11. Trong tiền sử gia đình có ghi nhận trường hợp sảy thai liên tiếp, thai chết lưu không rõ nguyên nhân hoặc phù thai không?",
    "12. Nơi sinh sống hiện tại hoặc quê quán thuộc khu vực miền núi, vùng sâu vùng xa có tỷ lệ mang gen cao không?",
    "13. Bố và Mẹ đẻ có phải là người cùng một dân tộc thiểu số không?",
    "14. Bố hoặc Mẹ đẻ đã từng được xác định mang gen bệnh Thalassemia chưa?",
    "15. Bạn có nhu cầu tư vấn di truyền trước hôn nhân hoặc trước khi sinh con không?",
]

BENH_VIEN = {
    "Hà Nội": [{"ten": "Viện Huyết học - Truyền máu Trung ương", "diachi": "Phố Phạm Văn Bạch, Cầu Giấy, Hà Nội", "dt": "024 3782 1895"}],
    "Thành phố Hồ Chí Minh": [{"ten": "Bệnh viện Truyền máu Huyết học TP.HCM", "diachi": "1 Trần Hữu Trang, Q. Tân Bình, TP.HCM", "dt": "028 3839 7535"}],
    "Thừa Thiên Huế": [{"ten": "Bệnh viện Trung ương Huế", "diachi": "16 Lê Lợi, TP. Huế", "dt": "0234 3822 325"}],
    "Đà Nẵng": [{"ten": "Bệnh viện Đà Nẵng", "diachi": "124 Hải Phòng, Q. Hải Châu, Đà Nẵng", "dt": "0236 3821 118"}],
    "Cần Thơ": [{"ten": "Bệnh viện Đa khoa Trung ương Cần Thơ", "diachi": "315 Nguyễn Văn Linh, Q. Ninh Kiều, Cần Thơ", "dt": "0292 3820 071"}],
    "Hải Phòng": [{"ten": "Bệnh viện Hữu nghị Việt Tiệp Hải Phòng", "diachi": "1 Nhà Thương, Cát Dài, Lê Chân, Hải Phòng", "dt": "0225 3700 436"}]
}

def lay_benh_vien_theo_tinh(tinh_ten):
    if not tinh_ten or tinh_ten == "Chọn Tỉnh/Thành phố": return None
    if tinh_ten in BENH_VIEN: return BENH_VIEN[tinh_ten][0]
    return {"ten": f"Bệnh viện Đa khoa Tỉnh/TP {tinh_ten}", "diachi": f"Trung tâm Tỉnh/TP {tinh_ten}", "dt": "Liên hệ 115"}

# 4. Phân tích chỉ số máu (Vòng 2)
def phan_tich_chiso_huyet_hoc(mcv, mch, hb_hieuchinh, rbc, rdw, gioitinh):
    goi_y_list = []
    hb_cut = 12.0 if gioitinh == "Nữ" else 13.0
    co_thieu_mau = (hb_hieuchinh > 0 and hb_hieuchinh < hb_cut)

    if mcv > 95.0 or mch > 32.0:
        goi_y_list.append("🔴 **Hồng cầu to / Ưu sắc (MCV > 95 fL, MCH > 32 pg):** Nghi ngờ **Thiếu Vitamin B12** hoặc **Thiếu Acid Folic (Folate)**, bệnh lý gan hoặc lạm dụng rượu.")
    elif 0 < mcv < 85.0 or 0 < mch < 28.0:
        if co_thieu_mau:
            goi_y_list.append("🟡 **Thiếu máu Hồng cầu nhỏ Nhược sắc (MCV < 85 fL, MCH < 28 pg):** Nghi ngờ **Thalassemia (Mang gen/Mắc bệnh)** hoặc **Thiếu máu Thiếu sắt**.")
        else:
            goi_y_list.append("🟡 **Hồng cầu nhỏ Nhược sắc không thiếu máu:** Dấu hiệu nghi ngờ cao **Người mang gen Thalassemia thể ẩn (Carrier)**.")

    if rdw > 15.0:
        goi_y_list.append("🟠 **Kích thước hồng cầu không đều (RDW > 15%):** Thường gặp trong **Thiếu máu thiếu sắt tiến triển** hoặc phối hợp thiếu máu.")
        
    if rbc >= (4.9 if gioitinh == "Nữ" else 5.4) and (0 < mcv < 85.0):
        goi_y_list.append("🟢 **Số lượng Hồng cầu (RBC) bảo tồn/tăng cao kèm MCV giảm:** Dấu hiệu nghiêng nhiều về **Thalassemia** hơn là Thiếu sắt.")

    if not goi_y_list and mcv > 0:
        goi_y_list.append("✅ Các chỉ số thể tích và hàm lượng Huyết sắc tố hồng cầu nằm trong giới hạn bình thường.")
        
    return goi_y_list

def score_round2(mcv, mch, hb_tho, rbc, rdw, gioitinh, do_cao_option):
    giam_hb = ALTITUDE_CORRECTION_MAP.get(do_cao_option, 0.0)
    hb_hieuchinh = hb_tho - giam_hb  
    hb_cut = 12.0 if gioitinh == "Nữ" else 13.0
    
    score_mcv = 4 if (0 < mcv < 85.0) else 0
    score_mch = 3 if (0 < mch < 28.0) else 0
    score_hb = 2 if (0 < hb_hieuchinh < hb_cut) else 0
    score_rbc = 2 if rbc >= (4.9 if gioitinh == "Nữ" else 5.4) else 0
    score_rdw = 2 if (11.0 <= rdw <= 15.0 and 0 < mcv < 85.0) else 0

    mentzer_idx = (mcv / rbc) if (rbc > 0 and mcv > 0) else 0.0
    mentzer_str = "Chưa xác định"
    if mentzer_idx > 0:
        if mentzer_idx < 13.0:
            mentzer_str = f"{mentzer_idx:.2f} (< 13: Nghi ngờ nghiêng về THALASSEMIA)"
        else:
            mentzer_str = f"{mentzer_idx:.2f} (≥ 13: Nghi ngờ nghiêng về THIẾU MÁU THIẾU SẮT)"

    str_hb_display = f"{hb_hieuchinh:.2f} g/dL (đã trừ {giam_hb:.1f} g/dL)" if giam_hb > 0 else f"{hb_hieuchinh:.2f} g/dL"
    ref_rbc = "4.0–4.9 M/uL" if gioitinh == "Nữ" else "4.2–5.4 M/uL"
    ref_hb = f"< 12.0 g/dL ({gioitinh})" if gioitinh == "Nữ" else f"< 13.0 g/dL ({gioitinh})"

    rows = [
        ("MCV (Thể tích hồng cầu)", f"{mcv:.2f} fL", "85.0 – 95.0 fL", score_mcv),
        ("MCH (Hb trung bình HC)", f"{mch:.2f} pg", "28.0 – 32.0 pg", score_mch),
        ("Hemoglobin (Hb hiệu chỉnh)", str_hb_display, ref_hb, score_hb),
        ("RBC (Số lượng hồng cầu)", f"{rbc:.2f} M/uL", ref_rbc, score_rbc),
        ("RDW (Độ phân bố HC)", f"{rdw:.2f} %", "11.0 – 15.0 %", score_rdw),
    ]
    
    goi_y_benh = phan_tich_chiso_huyet_hoc(mcv, mch, hb_hieuchinh, rbc, rdw, gioitinh)
    return sum(r[3] for r in rows), rows, hb_hieuchinh, mentzer_str, goi_y_benh

# 5. Lưu trữ SQLite
DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "thalalassemia_ho_so.db")

def _db_connect():
    conn = sqlite3.connect(DATA_FILE)
    conn.execute("CREATE TABLE IF NOT EXISTS ho_so (HoSoID TEXT PRIMARY KEY, data TEXT NOT NULL, updated_at TEXT NOT NULL)")
    conn.commit()
    return conn

def doc_du_lieu_luu_tru():
    try:
        conn = _db_connect()
        rows = conn.execute("SELECT data FROM ho_so ORDER BY updated_at DESC").fetchall()
        conn.close()
        return [json.loads(row[0]) for row in rows]
    except Exception: return []

def ghi_du_lieu_duy_nhat(record_data):
    try:
        conn = _db_connect()
        now = datetime.now().isoformat(timespec="seconds")
        ho_so_id = str(record_data.get("HoSoID") or uuid.uuid4().hex)
        record_data["HoSoID"] = ho_so_id
        conn.execute("INSERT OR REPLACE INTO ho_so (HoSoID, data, updated_at) VALUES (?, ?, ?)",
                     (ho_so_id, json.dumps(record_data, ensure_ascii=False), now))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Lỗi ghi dữ liệu: {e}")
        return False

# 6. Trạng thái Session
def init_state():
    defaults = {
        "page": "home", "is_admin": False, "ho_so_id": "", "s1": 0.0, "s2": 0, "r2_detail": None,
        "gioitinh": "Nữ", "hoten": "", "ngaysinh": date(2000, 1, 1), "dantoc": "Chọn dân tộc",
        "vung_o": "Chọn vùng/miền", "vung_lamviec": "Chọn vùng/miền", "tinh_o": "Chọn Tỉnh/Thành phố",
        "tinh_lamviec": "Chọn Tỉnh/Thành phố", "sdt": "", "do_cao": ALTITUDE_OPTIONS[0],
        "mcv": 0.0, "mch": 0.0, "hb": 0.0, "rbc": 0.0, "rdw": 0.0, "mentzer_str": "", "goi_y_benh": [],
        "vong1_da_luu": False, "vong2_da_luu": False, "vong3_da_luu": False,
        "dien_di_select": "Chưa thực hiện", "gen_select": "Chưa thực hiện", "ketluan_v3": "", "ghichu_v3": ""
    }
    for i in range(1, 16): defaults[f"q{i}"] = "Không"
    for k, v in defaults.items():
        if k not in st.session_state: st.session_state[k] = v

def luu_thong_tin(vong=None):
    ss = st.session_state
    if not ss.hoten.strip():
        st.warning("Vui lòng nhập Họ tên!")
        return False
    if not ss.ho_so_id: ss.ho_so_id = uuid.uuid4().hex

    if vong == 1: ss.vong1_da_luu = True
    elif vong == 2: ss.vong2_da_luu = True
    elif vong == 3: ss.vong3_da_luu = True

    data = {
        "HoSoID": ss.ho_so_id, "ThoiGian": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "HoTen": ss.hoten, "GioiTinh": ss.gioitinh, "NgaySinh": ss.ngaysinh.strftime("%d/%m/%Y"),
        "DanToc": ss.dantoc, "SDT": ss.sdt, 
        "TinhO": ss.tinh_o, "VungO": ss.vung_o,
        "TinhLamViec": ss.tinh_lamviec, "VungLamViec": ss.vung_lamviec, 
        "DoCaoSinhSong": ss.do_cao,
        "DiemV1": ss.s1, "CauTraLoiV1": [{"stt": i, "cau_hoi": CAU_HOI[i-1], "tra_loi": ss.get(f"q{i}")} for i in range(1, 16)],
        "MCV": ss.mcv, "MCH": ss.mch, "Hb": ss.hb, "RBC": ss.rbc, "RDW": ss.rdw,
        "MentzerIndex": ss.mentzer_str, "GoiYBenhLy": ss.goi_y_benh, "DiemV2": ss.s2, "ChiTietV2": ss.r2_detail,
        "DienDiHb": ss.dien_di_select, "XetNghiemGen": ss.gen_select,
        "KetLuanV3": ss.ketluan_v3, "GhiChuV3": ss.ghichu_v3,
        "Vong1DaLuu": ss.vong1_da_luu, "Vong2DaLuu": ss.vong2_da_luu, "Vong3DaLuu": ss.vong3_da_luu
    }
    return ghi_du_lieu_duy_nhat(data)

# 7. Xuất Báo Cáo Word
def tao_phieu_word(data_dict):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(1.5)

    def set_cell_text(cell, text, bold=False):
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(2)
        r = p.add_run(str(text))
        r.bold = bold
        r.font.name = "Arial"
        r.font.size = Pt(9.5)

    def add_table(headers, rows):
        table = doc.add_table(rows=1, cols=len(headers))
        table.style = "Table Grid"
        for i, h in enumerate(headers): set_cell_text(table.rows[0].cells[i], h, True)
        for row_idx, row in enumerate(rows):
            row_cells = table.add_row().cells
            for i, value in enumerate(row): set_cell_text(row_cells[i], value)
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("PHIẾU KẾT QUẢ SÀNG LỌC THALASSEMIA & BỆNH LÝ HUYẾT HỌC\n")
    r.bold = True
    r.font.size = Pt(13)

    add_table(["Thông tin cá nhân", "Giá trị khai báo"], [
        ["Họ và tên", data_dict.get("HoTen", "")],
        ["Giới tính / Ngày sinh", f"{data_dict.get('GioiTinh', '')} - {data_dict.get('NgaySinh', '')}"],
        ["Dân tộc / SĐT", f"{data_dict.get('DanToc', '')} - {data_dict.get('SDT', '')}"],
        ["Nơi sinh sống", f"{data_dict.get('TinhO', '')} ({data_dict.get('VungO', '')})"],
        ["Nơi làm việc / Học tập", f"{data_dict.get('TinhLamViec', '')} ({data_dict.get('VungLamViec', '')})"],
    ])

    if data_dict.get("Vong1DaLuu"):
        p1 = doc.add_paragraph()
        p1.add_run(f"1. Điểm Tiền sử & Dịch tễ (Vòng 1): {data_dict.get('DiemV1', 0.0)} điểm\n").bold = True

    if data_dict.get("Vong2DaLuu"):
        p2 = doc.add_paragraph()
        p2.add_run("2. Kết quả Vòng 2 (Công thức máu & Phân tích định hướng)\n").bold = True
        add_table(["Thông số", "Giá trị", "Tham chiếu (Việt Nam)", "Điểm"], data_dict.get("ChiTietV2", []))
        
        p_m = doc.add_paragraph()
        p_m.add_run(f"• Chỉ số Mentzer Index (MCV/RBC): {data_dict.get('MentzerIndex', '')}\n").bold = True
        
        p_g = doc.add_paragraph()
        p_g.add_run("• Gợi ý bệnh lý / triệu chứng từ chỉ số máu:\n").bold = True
        for g in data_dict.get("GoiYBenhLy", []):
            p_g.add_run(f"  - {g.replace('**', '')}\n")

    if data_dict.get("Vong3DaLuu"):
        p3 = doc.add_paragraph()
        p3.add_run("\n3. Kết quả Vòng 3 (Chuyên sâu Điện di & Gen)\n").bold = True
        add_table(["Hạng mục Vòng 3", "Kết quả ghi nhận / Tích chọn"], [
            ["Điện di Hemoglobin (Hb)", data_dict.get("DienDiHb", "Chưa chọn")],
            ["Xét nghiệm Gen Thalassemia", data_dict.get("XetNghiemGen", "Chưa chọn")],
            ["Kết luận chuyên môn Vòng 3", data_dict.get("KetLuanV3", "")],
            ["Ghi chú / Hướng xử trí", data_dict.get("GhiChuV3", "")]
        ])

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# 8. Màn hình Báo cáo Admin
def render_admin():
    st.subheader("📊 Trang Quản Trị & Báo Cáo Thống Kê")
    ds = doc_du_lieu_luu_tru()
    if not ds:
        st.info("Chưa có dữ liệu.")
        return
    
    st.metric("Tổng số hồ sơ trong hệ thống", len(ds))
    df = pd.DataFrame(ds)
    st.dataframe(df[["HoSoID", "ThoiGian", "HoTen", "GioiTinh", "DanToc", "TinhO", "TinhLamViec", "DiemV1", "DiemV2", "MentzerIndex"]], use_container_width=True)
    
    csv_data = df.to_csv(index=False).encode('utf-8-sig')
    st.download_button("📥 Tải file Báo cáo CSV toàn bộ hồ sơ", csv_data, f"BAO_CAO_{datetime.now().strftime('%Y%m%d')}.csv", "text/csv")

# 9. Giao diện Chính
def render_main():
    init_state()
    ss = st.session_state

    st.markdown("""
    <div class="hero">
        <h1>🩸 SÀNG LỌC THALASSEMIA & BỆNH LÝ HUYẾT HỌC</h1>
        <p>Phân tích chỉ số máu nâng cao • Chuẩn hóa khoảng tham chiếu Việt Nam • Đánh giá rủi ro theo từng vòng</p>
    </div>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.title("⚙️ Hệ thống")
        if not ss.is_admin:
            pwd = st.text_input("Mật khẩu Admin:", type="password")
            if st.button("Đăng nhập Admin"):
                if pwd == ADMIN_PASSWORD:
                    ss.is_admin = True
                    st.rerun()
                else: st.error("Sai mật khẩu!")
        else:
            st.success("👨‍⚕️ Cán bộ Y tế")
            vmode = st.radio("Chế độ:", ["Chẩn đoán", "Báo cáo / Quản trị"])
            if st.button("Đăng xuất"):
                ss.is_admin = False
                st.rerun()

    if ss.is_admin and 'vmode' in locals() and vmode == "Báo cáo / Quản trị":
        render_admin()
        return

    # Khai báo thông tin hành chính
    st.subheader("📋 1. Thông tin cá nhân & Địa bàn sinh sống, làm việc")
    c1, c2 = st.columns(2)
    with c1:
        ss.hoten = st.text_input("Họ và tên (*):", value=ss.hoten)
        c1a, c1b = st.columns(2)
        with c1a: ss.gioitinh = st.selectbox("Giới tính:", ["Nữ", "Nam"], index=0 if ss.gioitinh=="Nữ" else 1)
        with c1b: ss.ngaysinh = st.date_input("Ngày sinh:", value=ss.ngaysinh, format="DD/MM/YYYY")
        ss.dantoc = st.selectbox("Dân tộc:", DAN_TOC, index=DAN_TOC.index(ss.dantoc) if ss.dantoc in DAN_TOC else 0)
        ss.sdt = st.text_input("Số điện thoại:", value=ss.sdt)

    with c2:
        c2a, c2b = st.columns(2)
        with c2a: ss.vung_o = st.selectbox("Vùng sinh sống:", VUNG_MIEN, index=VUNG_MIEN.index(ss.vung_o) if ss.vung_o in VUNG_MIEN else 0)
        with c2b:
            tinh_ds_o = TAT_CA_TINH if ss.vung_o == "Chọn vùng/miền" else ["Chọn Tỉnh/Thành phố"] + TINH_THEO_VUNG.get(ss.vung_o, [])
            ss.tinh_o = st.selectbox("Tỉnh/Thành phố sinh sống:", tinh_ds_o, index=tinh_ds_o.index(ss.tinh_o) if ss.tinh_o in tinh_ds_o else 0)
        
        c2c, c2d = st.columns(2)
        with c2c: ss.vung_lamviec = st.selectbox("Vùng làm việc / Học tập:", VUNG_MIEN, index=VUNG_MIEN.index(ss.vung_lamviec) if ss.vung_lamviec in VUNG_MIEN else 0)
        with c2d:
            tinh_ds_lv = TAT_CA_TINH if ss.vung_lamviec == "Chọn vùng/miền" else ["Chọn Tỉnh/Thành phố"] + TINH_THEO_VUNG.get(ss.vung_lamviec, [])
            ss.tinh_lamviec = st.selectbox("Tỉnh/TP làm việc / Học tập:", tinh_ds_lv, index=tinh_ds_lv.index(ss.tinh_lamviec) if ss.tinh_lamviec in tinh_ds_lv else 0)
            
        ss.do_cao = st.selectbox("🏔️ Độ cao nơi sinh sống (Trừ Hb WHO):", ALTITUDE_OPTIONS, index=ALTITUDE_OPTIONS.index(ss.do_cao) if ss.do_cao in ALTITUDE_OPTIONS else 0)

    bv_o = lay_benh_vien_theo_tinh(ss.tinh_o)
    bv_lv = lay_benh_vien_theo_tinh(ss.tinh_lamviec)

    if bv_o or bv_lv:
        st.markdown("---")
        st.markdown("### 🏥 Cơ sở y tế gợi ý gần khu vực của bạn:")
        if bv_o:
            st.markdown(f"🏡 **Gần nơi sinh sống ({ss.tinh_o}):** {bv_o['ten']} - 📍 {bv_o['diachi']} (Hotline: {bv_o['dt']})")
        if bv_lv and ss.tinh_lamviec != ss.tinh_o:
            st.markdown(f"🏢 **Gần nơi làm việc ({ss.tinh_lamviec}):** {bv_lv['ten']} - 📍 {bv_lv['diachi']} (Hotline: {bv_lv['dt']})")

    st.markdown("---")

    tab1, tab2, tab3 = st.tabs(["VÒNG 1: Tiền sử & Dịch tễ", "VÒNG 2: Công thức máu & Gợi ý bệnh lý", "VÒNG 3: Điện di & Gen (Chuyên sâu)"])

    # ------------------ VÒNG 1 ------------------
    with tab1:
        st.caption("📝 *Trả lời câu hỏi tiền sử bản thân và gia đình:*")
        score_v1 = DIEM_DAN_TOC_VN.get(ss.dantoc, 0.5)
        for idx, q_text in enumerate(CAU_HOI, 1):
            cq, ca = st.columns([3.5, 1])
            with cq: st.write(q_text)
            with ca:
                if st.radio(f"q_{idx}", ["Có", "Không"], key=f"q{idx}", horizontal=True, label_visibility="collapsed") == "Có":
                    score_v1 += 1.0
        ss.s1 = score_v1

        st.markdown("---")
        st.markdown("### 📊 Đánh giá & Gợi ý Vòng 1:")
        st.metric("Tổng điểm yếu tố nguy cơ Vòng 1", f"{ss.s1:.1f} điểm")

        if ss.s1 >= 4.0:
            box("danger-box", f"🚨 <b>Mức độ nguy cơ: CAO ({ss.s1:.1f} điểm)</b><br>• Tiền sử gia đình hoặc dân tộc có tỷ lệ mang gen Thalassemia cao.<br>👉 <b>Gợi ý:</b> Bắt buộc thực hiện tiếp xét nghiệm Tổng phân tích tế bào máu (Vòng 2) và chủ động tham vấn y tế.")
        elif ss.s1 >= 2.0:
            box("warning-box", f"⚠️ <b>Mức độ nguy cơ: TRUNG BÌNH ({ss.s1:.1f} điểm)</b><br>• Đã ghi nhận yếu tố tiền sử hoặc triệu chứng nghi ngờ nhẹ.<br>👉 <b>Gợi ý:</b> Khuyên làm xét nghiệm Công thức máu (Vòng 2) để kiểm tra các chỉ số MCV, MCH, Hb.")
        else:
            box("success-box", f"✅ <b>Mức độ nguy cơ: THẤP ({ss.s1:.1f} điểm)</b><br>• Chưa phát hiện yếu tố tiền sử dịch tễ đáng ngại.<br>👉 <b>Gợi ý:</b> Vẫn nên làm xét nghiệm công thức máu định kỳ hoặc khám sức khỏe tiền hôn nhân.")

        if st.button("💾 Lưu Vòng 1"):
            if luu_thong_tin(vong=1): st.success("Đã lưu Vòng 1 thành công!")

    # ------------------ VÒNG 2 ------------------
    with tab2:
        st.caption("🔬 *Nhập các chỉ số từ kết quả xét nghiệm Tổng phân tích tế bào máu ngoại vi:*")
        c_mcv, c_mch, c_hb, c_rbc, c_rdw = st.columns(5)
        with c_mcv: ss.mcv = st.number_input("MCV (fL):", value=ss.mcv, step=0.1)
        with c_mch: ss.mch = st.number_input("MCH (pg):", value=ss.mch, step=0.1)
        with c_hb: ss.hb = st.number_input("Hb (g/dL):", value=ss.hb, step=0.1)
        with c_rbc: ss.rbc = st.number_input("RBC (M/uL):", value=ss.rbc, step=0.01)
        with c_rdw: ss.rdw = st.number_input("RDW (%):", value=ss.rdw, step=0.1)

        s2, r2_rows, hb_h, mentzer_str, goi_y_benh = score_round2(ss.mcv, ss.mch, ss.hb, ss.rbc, ss.rdw, ss.gioitinh, ss.do_cao)
        ss.s2, ss.r2_detail, ss.mentzer_str, ss.goi_y_benh = s2, r2_rows, mentzer_str, goi_y_benh

        if ss.mcv > 0:
            st.markdown("---")
            st.markdown("### 📊 Đánh giá & Gợi ý Vòng 2:")
            
            c_score1, c_score2 = st.columns(2)
            with c_score1: st.metric("Điểm chỉ số máu Vòng 2", f"{ss.s2} / 13 điểm")
            with c_score2: box("info-box", f"📊 <b>Chỉ số Mentzer Index (MCV/RBC):</b> {mentzer_str}")

            if ss.s2 >= 6:
                box("danger-box", "🚨 <b>Mức độ nghi ngờ: RẤT CAO</b><br>Các chỉ số thể hiện rõ tình trạng Hồng cầu nhỏ Nhược sắc. Cần làm ngay Vòng 3 (Điện di Hb & Xét nghiệm Gen).")
            elif ss.s2 >= 3:
                box("warning-box", "⚠️ <b>Mức độ nghi ngờ: TRUNG BÌNH</b><br>Có bất thường nhẹ về kích thước hoặc hàm lượng huyết sắc tố hồng cầu.")
            else:
                box("success-box", "✅ <b>Mức độ nghi ngờ: THẤP</b><br>Các chỉ số dòng hồng cầu nằm trong giới hạn bình thường.")

            st.markdown("**Gợi ý nghi ngờ bệnh lý & Triệu chứng:**")
            for item in goi_y_benh:
                st.markdown(f"- {item}")

        if st.button("💾 Lưu Vòng 2"):
            if luu_thong_tin(vong=2): st.success("Đã lưu Vòng 2 thành công!")

    # ------------------ VÒNG 3 ------------------
    with tab3:
        st.markdown("### 🧬 Vòng 3: Xét nghiệm Điện di Hb & Phân tích Gen")
        st.caption("*Tích chọn các kết quả cận lâm sàng chuyên sâu:*")

        ss.dien_di_select = st.radio(
            "1. Kết quả Điện di Huyết sắc tố (Hb):",
            [
                "Chưa thực hiện",
                "Bình thường (HbA ≥ 96%, HbA2: 2.0-3.5%, HbF < 1%)",
                "Gợi ý β-Thalassemia thể ẩn (HbA2 > 3.5% hoặc HbF: 2-10%)",
                "Gợi ý Bệnh HbE / β-Thal-HbE (Xuất hiện băng HbE)",
                "Gợi ý α-Thalassemia / HbH (Xuất hiện băng HbH/Bart's)",
                "Bất thường khác"
            ]
        )

        ss.gen_select = st.radio(
            "2. Kết quả Xét nghiệm Gen (PCR / DNA Sequencing):",
            [
                "Chưa thực hiện",
                "Không phát hiện đột biến",
                "Đột biến α-Globin mất đoạn (--SEA, -α3.7, -α4.2...)",
                "Đột biến α-Globin không mất đoạn (HbCS, HbQS...)",
                "Đột biến β-Globin thể dị hợp (CD41/42, IVS1-1, CD17...)",
                "Đột biến β-Globin thể đồng hợp / Tạp dị hợp"
            ]
        )

        ss.ketluan_v3 = st.text_area("3. Kết luận chuyên môn Vòng 3:", value=ss.ketluan_v3)
        ss.ghichu_v3 = st.text_area("4. Ghi chú / Hướng xử trí tiếp theo:", value=ss.ghichu_v3)

        st.markdown("---")
        st.markdown("### 📊 Kết luận & Khuyến cáo di truyền Vòng 3:")
        
        co_dot_bien = "Đột biến" in ss.gen_select or "Gợi ý" in ss.dien_di_select
        if co_dot_bien:
            box("danger-box", "🔴 **XÁC ĐỊNH MANG GEN / MẮC BỆNH THALASSEMIA**<br>• Cần được tư vấn di truyền trước hôn nhân hoặc trước khi mang thai.<br>• Nếu bạn đời cũng mang gen, cần thực hiện chẩn đoán trước sinh (chọc ối/sinh thiết gai nhau).")
        elif ss.gen_select == "Không phát hiện đột biến" and "Bình thường" in ss.dien_di_select:
            box("success-box", "🟢 **KẾT QUẢ AN TOÀN**<br>• Không phát hiện các đột biến gen Thalassemia phổ biến. Rủi ro di truyền cho thế hệ sau rất thấp.")
        else:
            box("info-box", "🔵 **ĐANG CHỜ / CHƯA HOÀN THÀNH XÉT NGHIỆM VÒNG 3**<br>• Vui lòng hoàn thành xét nghiệm Điện di Hb và Gen để nhận được kết luận chính xác nhất từ bác sĩ chuyên khoa.")

        if st.button("💾 Lưu Vòng 3"):
            if luu_thong_tin(vong=3): st.success("Đã lưu Vòng 3 thành công!")

    st.markdown("---")
    
    st.markdown("🔗 **Cổng thông tin & Trang tư vấn Sàng lọc Tiền hôn nhân uy tín:**")
    st.markdown("""
    * 🩸 [Viện Huyết học - Truyền máu Trung ương (Tư vấn Thalassemia)](https://vienhuyethoc.vn)
    * 🏥 [Bệnh viện Từ Dũ - Khám & Tư vấn Sức khỏe Tiền hôn nhân](https://tudu.com.vn)
    * 👶 [Bệnh viện Nhi Đồng 1 - Chuyên khoa Huyết học](https://nhidong.org.vn)
    """)

    word_bytes = tao_phieu_word({
        "HoTen": ss.hoten, "GioiTinh": ss.gioitinh, "NgaySinh": ss.ngaysinh.strftime("%d/%m/%Y"),
        "DanToc": ss.dantoc, "SDT": ss.sdt, "TinhO": ss.tinh_o, "VungO": ss.vung_o,
        "TinhLamViec": ss.tinh_lamviec, "VungLamViec": ss.vung_lamviec,
        "DiemV1": ss.s1, "DiemV2": ss.s2,
        "ChiTietV2": ss.r2_detail, "MentzerIndex": ss.mentzer_str, "GoiYBenhLy": ss.goi_y_benh,
        "DienDiHb": ss.dien_di_select, "XetNghiemGen": ss.gen_select,
        "KetLuanV3": ss.ketluan_v3, "GhiChuV3": ss.ghichu_v3,
        "Vong1DaLuu": ss.vong1_da_luu, "Vong2DaLuu": ss.vong2_da_luu, "Vong3DaLuu": ss.vong3_da_luu
    })
    
    st.download_button("📄 Tải Phiếu Kết Quả Chuẩn Word (.docx)", word_bytes, f"PHIEU_SANG_LOC_{ss.hoten.replace(' ', '_')}.docx",
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

if __name__ == "__main__":
    render_main()
