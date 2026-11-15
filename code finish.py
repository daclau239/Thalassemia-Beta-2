"""thal_core.py — lõi tính toán của đề tài:
"Xây dựng và đánh giá mô hình học máy dự đoán nguy cơ mang gen Thalassemia dựa trên chỉ số
huyết học (CBC) và Ferritin huyết thanh ở người trưởng thành" (BV Trung ương Huế).

Không phụ thuộc Streamlit → dễ kiểm thử. Gồm:
  1. Cỡ mẫu (công thức đề cương + độ nhạy/độ đặc hiệu + mô hình dự đoán Riley)
  2. Tiêu chuẩn chọn/loại trừ, gán nhãn kết cục theo xét nghiệm tham chiếu
  3. Thống kê mô tả/so sánh (Mục tiêu 1)
  4. Học máy LR/RF/XGBoost, đánh giá, so sánh CBC vs CBC+Ferritin, SHAP (Mục tiêu 2)
  5. Đóng gói mô hình có chữ ký HMAC, dự đoán cho một người
  6. Nhập dữ liệu từ file (mã hóa, quy ước mã, đơn vị), AI đọc phiếu
"""
from __future__ import annotations

import base64
import hashlib
import hmac
import io
import json
import math
import numbers
import os
import pickle
import re
import tempfile
import unicodedata
import warnings
from dataclasses import dataclass, asdict
from datetime import date

import numpy as np
import pandas as pd

# =====================================================================
# 0. LƯỢC ĐỒ DỮ LIỆU (theo mục 2.4 đề cương)
# =====================================================================
CBC7 = ["rbc", "hgb", "hct", "mcv", "mch", "mchc", "rdw"]          # biến đầu vào CBC của mô hình lý luận
EXTRA_FEATURES = ["age", "sex_male", "wbc", "plt"]                   # biến tùy chọn (Bảng 2)
LAB_COLS = CBC7 + ["wbc", "plt", "ferritin"]

VAR_LABEL = {
    "rbc": "RBC (T/L)", "hgb": "HGB (g/L)", "hct": "HCT (%)", "mcv": "MCV (fL)", "mch": "MCH (pg)",
    "mchc": "MCHC (g/L)", "rdw": "RDW-CV (%)", "wbc": "WBC (G/L)", "plt": "PLT (G/L)",
    "ferritin": "Ferritin (ng/mL)", "age": "Tuổi (năm)", "sex_male": "Giới tính nam (1/0)",
    "crp": "CRP (mg/L)", "hba2": "HbA2 (%)", "hbf": "HbF (%)", "hbe": "HbE (%)",
}
VAR_SHORT = {k: v.split(" (")[0] for k, v in VAR_LABEL.items()}
VAR_DECIMALS = {"rbc": 2, "hgb": 1, "hct": 1, "mcv": 1, "mch": 1, "mchc": 1, "rdw": 1, "wbc": 2, "plt": 0,
                "ferritin": 1, "age": 1, "crp": 1, "hba2": 2, "hbf": 2, "hbe": 2}
PLAUSIBLE = {"rbc": (1.0, 9.0), "hgb": (30, 250), "hct": (10, 70), "mcv": (40, 130), "mch": (10, 45),
             "mchc": (250, 400), "rdw": (8, 35), "wbc": (0.3, 150), "plt": (5, 1500),
             "ferritin": (0.5, 20000), "age": (0, 120)}

PHASES = ["Hồi cứu", "Tiến cứu"]
SEXES = ["Nam", "Nữ"]
YN3 = ["Không", "Có", "Không rõ"]
GENE_STATUS = ["Chưa làm", "Không phát hiện biến thể", "Alpha-thalassemia", "Beta-thalassemia",
               "Đồng mang alpha và beta"]
HPLC_STATUS = ["Chưa làm", "Bình thường", "Gợi ý beta-thalassemia", "HbE / hemoglobinopathy khác",
               "Bất thường khác"]
CONSENT_TYPES = ["Hồi cứu — theo quyết định của Hội đồng đạo đức", "Tiến cứu — đã đồng thuận tham gia"]
GROUP_ORDER = ["Không mang gen", "Alpha", "Beta", "Đồng mang"]

STUDY_COLUMNS = [
    "study_code", "phase", "sample_date", "age", "sex",
    "rbc", "hgb", "hct", "mcv", "mch", "mchc", "rdw", "wbc", "plt",
    "ferritin", "ferritin_date", "crp", "inflammation", "transfusion", "transfusion_days",
    "gene_status", "gene_detail", "gene_date", "hplc_status", "hba2", "hbf", "hbe", "hplc_date",
    "analyzer", "ferritin_system", "consent_type", "note",
]
NUMERIC_COLS = ["age", "rbc", "hgb", "hct", "mcv", "mch", "mchc", "rdw", "wbc", "plt", "ferritin", "crp",
                "transfusion_days", "hba2", "hbf", "hbe"]
DATE_COLS = ["sample_date", "ferritin_date", "gene_date", "hplc_date"]
COLUMN_VI = {
    "study_code": "Mã nghiên cứu", "phase": "Giai đoạn", "sample_date": "Ngày XN CBC", "age": "Tuổi", "sex": "Giới tính",
    "rbc": "RBC (T/L)", "hgb": "HGB (g/L)", "hct": "HCT (%)", "mcv": "MCV (fL)", "mch": "MCH (pg)",
    "mchc": "MCHC (g/L)", "rdw": "RDW-CV (%)", "wbc": "WBC (G/L)", "plt": "PLT (G/L)",
    "ferritin": "Ferritin (ng/mL)", "ferritin_date": "Ngày XN Ferritin", "crp": "CRP (mg/L)",
    "inflammation": "Tình trạng viêm", "transfusion": "Tiền sử truyền máu",
    "transfusion_days": "Số ngày từ lần truyền gần nhất", "gene_status": "Kết quả gen (HBA1/HBA2/HBB)",
    "gene_detail": "Chi tiết biến thể", "gene_date": "Ngày XN gen", "hplc_status": "HPLC/điện di",
    "hba2": "HbA2 (%)", "hbf": "HbF (%)", "hbe": "HbE (%)", "hplc_date": "Ngày HPLC/điện di",
    "analyzer": "Máy huyết học", "ferritin_system": "Hệ thống Ferritin", "consent_type": "Hình thức đồng thuận",
    "note": "Ghi chú",
}


# =====================================================================
# 1. CỠ MẪU
# =====================================================================
def z_value(alpha: float = 0.05) -> float:
    """Z(1-α/2). Với α=0,05 dùng 1,96 đúng như đề cương."""
    if abs(alpha - 0.05) < 1e-12:
        return 1.96
    from scipy.stats import norm
    return float(norm.ppf(1 - alpha / 2))


def n_single_proportion(p: float, E: float, alpha: float = 0.05) -> dict:
    """n = Z²·p·(1−p)/E²  (mục 2.2.5 đề cương: p = 0,1326; E = 0,05 → 177)."""
    z = z_value(alpha)
    raw = z * z * p * (1 - p) / (E * E)
    return {"z": z, "numerator": z * z * p * (1 - p), "denominator": E * E, "n_raw": raw, "n": math.ceil(raw - 1e-9)}


def n_sens_spec(se: float, sp: float, prevalence: float, E: float, alpha: float = 0.05) -> dict:
    """Cỡ mẫu để ước lượng độ nhạy/độ đặc hiệu với sai số tuyệt đối E (thay p bằng Se hoặc Sp)."""
    z = z_value(alpha)
    n_cases = math.ceil(z * z * se * (1 - se) / (E * E) - 1e-9)
    n_ctrl = math.ceil(z * z * sp * (1 - sp) / (E * E) - 1e-9)
    total_from_cases = math.ceil(n_cases / prevalence)
    total_from_ctrl = math.ceil(n_ctrl / (1 - prevalence))
    return {"n_cases": n_cases, "n_controls": n_ctrl, "total_from_cases": total_from_cases,
            "total_from_controls": total_from_ctrl, "total": max(total_from_cases, total_from_ctrl)}


def n_prediction_model(P: int, prevalence: float, shrinkage: float = 0.9, r2_fraction: float = 0.15,
                       r2_cs: float | None = None, E: float = 0.05, alpha: float = 0.05, epp: float = 10) -> dict:
    """Cỡ mẫu cho nghiên cứu XÂY DỰNG mô hình dự đoán nhị phân (Riley và cs., BMJ 2020):
      (i)  hệ số co rút ≥ S:   n = P / ((S−1)·ln(1 − R²cs/S))
      (iii) độ chính xác của nguy cơ nền: n = Z²·p(1−p)/E²   (chính là công thức của đề cương)
      và quy tắc tham khảo: số biến cố trên mỗi tham số (EPP)."""
    p = prevalence
    max_r2 = 1 - math.exp(2 * (p * math.log(p) + (1 - p) * math.log(1 - p)))
    r2 = r2_cs if r2_cs is not None else r2_fraction * max_r2
    n1 = P / ((shrinkage - 1) * math.log(1 - r2 / shrinkage))
    n3 = n_single_proportion(p, E, alpha)["n_raw"]
    n_epp = epp * P / p
    n_final = max(n1, n3, n_epp)
    return {"max_r2cs": max_r2, "r2cs": r2, "n_shrinkage": math.ceil(n1), "n_precision": math.ceil(n3),
            "n_epp": math.ceil(n_epp), "n": math.ceil(n_final), "events": math.ceil(n_final * p)}


# =====================================================================
# 2. TIÊU CHUẨN CHỌN / LOẠI TRỪ VÀ KẾT CỤC
# =====================================================================
def _num(s: pd.Series) -> pd.Series:
    if s.dtype == object or str(s.dtype).startswith(("str", "string")):
        s = s.astype(str).str.replace(",", ".", regex=False).replace({"nan": None, "None": None, "": None, "<NA>": None})
    return pd.to_numeric(s, errors="coerce")


def _date(s: pd.Series) -> pd.Series:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return pd.to_datetime(s, errors="coerce", dayfirst=True, format="mixed")


def prepare(df: pd.DataFrame) -> pd.DataFrame:
    """Chuẩn hóa kiểu dữ liệu, thêm các cột dẫn xuất (sex_male)."""
    d = df.copy()
    for c in STUDY_COLUMNS:
        if c not in d.columns:
            d[c] = np.nan
    for c in NUMERIC_COLS:
        d[c] = _num(d[c])
    for c in DATE_COLS:
        d[c] = _date(d[c])
    d["sex_male"] = d["sex"].map({"Nam": 1.0, "Nữ": 0.0})
    return d.reset_index(drop=True)


def derive_outcome(d: pd.DataFrame, scope: str = "alpha_beta"):
    """Gán nhãn kết cục theo xét nghiệm tham chiếu (mục 2.2.4 và 2.4.1).

    scope = "alpha_beta": kết cục = mang alpha và/hoặc beta → CẦN xét nghiệm gen (HBA1/HBA2/HBB).
        HPLC/điện di bình thường KHÔNG đủ để gán nhãn âm tính (không loại trừ alpha-thalassemia).
    scope = "beta_only": kết cục = mang beta; cho phép HPLC/điện di làm tham chiếu khi chưa có gen
        (HPLC bình thường → âm tính với beta; gợi ý beta → dương tính).
    Trả về (y, nguồn tham chiếu, nhóm gen) — y là NaN nếu không xác định được."""
    n = len(d)
    y = np.full(n, np.nan)
    src = np.full(n, "", dtype=object)
    grp = np.full(n, None, dtype=object)
    gs = d["gene_status"].to_numpy(dtype=object)
    hs = d["hplc_status"].to_numpy(dtype=object)
    for i in range(n):
        g, h = gs[i], hs[i]
        if g == "Không phát hiện biến thể":
            y[i], src[i], grp[i] = 0, "Gen", "Không mang gen"
        elif g == "Alpha-thalassemia":
            y[i], src[i], grp[i] = (1 if scope == "alpha_beta" else 0), "Gen", "Alpha"
        elif g == "Beta-thalassemia":
            y[i], src[i], grp[i] = 1, "Gen", "Beta"
        elif g == "Đồng mang alpha và beta":
            y[i], src[i], grp[i] = 1, "Gen", "Đồng mang"
        elif scope == "beta_only":
            if h == "Gợi ý beta-thalassemia":
                y[i], src[i] = 1, "HPLC/điện di"
            elif h == "Bình thường":
                y[i], src[i] = 0, "HPLC/điện di"
    return pd.Series(y, index=d.index), pd.Series(src, index=d.index), pd.Series(grp, index=d.index)


@dataclass
class Criteria:
    min_age: float = 18
    require_ferritin: bool = True
    require_reference: bool = True
    scope: str = "alpha_beta"
    link_days: int | None = 30
    transfusion_days: int | None = 90
    check_ranges: bool = True
    required_cbc: tuple = tuple(CBC7)


def apply_criteria(df: pd.DataFrame, c: Criteria):
    """Áp dụng tiêu chuẩn chọn mẫu/loại trừ tuần tự (mục 2.2.2–2.2.3).
    Trả về (hồ sơ đạt, bảng sơ đồ luồng, bảng lý do loại trừ, ghi chú)."""
    d = prepare(df)
    d["_reason"] = ""
    d["y"], d["ref_source"], d["nhom_gen"] = derive_outcome(d, c.scope)
    flow = [{"Bước": "Tổng số hồ sơ", "Loại": 0, "Còn lại": len(d)}]

    def exclude(mask, reason):
        m = pd.Series(mask, index=d.index).fillna(False).astype(bool) & (d["_reason"] == "")
        d.loc[m, "_reason"] = reason
        flow.append({"Bước": reason, "Loại": int(m.sum()), "Còn lại": int((d["_reason"] == "").sum())})

    notes = []
    exclude(d["age"].isna(), "Không xác định tuổi")
    exclude(d["age"] < c.min_age, f"Tuổi < {c.min_age:g}")
    exclude(d[list(c.required_cbc)].isna().any(axis=1), "Thiếu chỉ số CBC bắt buộc")
    if c.require_ferritin:
        exclude(d["ferritin"].isna(), "Thiếu Ferritin huyết thanh")
    if c.check_ranges:
        bad = pd.Series(False, index=d.index)
        for k, (lo, hi) in PLAUSIBLE.items():
            if k in d.columns:
                bad |= (d[k] < lo) | (d[k] > hi)
        exclude(bad, "Giá trị xét nghiệm không hợp lý")
    if c.transfusion_days is not None:
        tr = (d["transfusion"] == "Có") & (d["transfusion_days"].isna() | (d["transfusion_days"] <= c.transfusion_days))
        exclude(tr, f"Truyền máu ≤ {c.transfusion_days} ngày hoặc không rõ thời điểm")
    if c.link_days is not None:
        ref_date = d["gene_date"].where(d["gene_date"].notna(), d["hplc_date"])
        d1 = (d["ferritin_date"] - d["sample_date"]).abs().dt.days
        d2 = (ref_date - d["sample_date"]).abs().dt.days
        exclude((d1 > c.link_days) | (d2 > c.link_days), f"Các xét nghiệm cách nhau > {c.link_days} ngày")
        unverified = int((d["_reason"] == "").sum() - ((d["sample_date"].notna() & d["ferritin_date"].notna()
                                                        & ref_date.notna()) & (d["_reason"] == "")).sum())
        if unverified:
            notes.append(f"{unverified} hồ sơ thiếu ngày xét nghiệm nên chưa kiểm tra được khoảng cách thời gian giữa CBC, Ferritin và xét nghiệm tham chiếu.")
    if c.require_reference:
        exclude(d["y"].isna(), "Không có kết quả tham chiếu xác định được kết cục")
        if c.scope == "alpha_beta":
            n_hplc_only = int(((d["gene_status"] == "Chưa làm") | d["gene_status"].isna()).sum())
            if n_hplc_only:
                notes.append("Phạm vi alpha và/hoặc beta: hồ sơ chỉ có HPLC/điện di (chưa có xét nghiệm gen) không được gán nhãn — "
                             "HPLC bình thường không loại trừ alpha-thalassemia.")
    inc = d[d["_reason"] == ""].drop(columns=["_reason"]).reset_index(drop=True)
    excluded = d[d["_reason"] != ""][["study_code", "_reason"]].rename(columns={"study_code": "Mã nghiên cứu", "_reason": "Lý do loại trừ"})
    return inc, pd.DataFrame(flow), excluded.reset_index(drop=True), notes


# =====================================================================
# 3. THỐNG KÊ MÔ TẢ VÀ SO SÁNH (MỤC TIÊU 1)
# =====================================================================
def fmt_num(x, d: int = 2) -> str:
    if x is None or (isinstance(x, float) and (math.isnan(x) or math.isinf(x))):
        return "—"
    return f"{x:.{d}f}".replace(".", ",")


def fmt_p(p) -> str:
    if p is None or (isinstance(p, float) and math.isnan(p)):
        return "—"
    return "<0,001" if p < 0.001 else fmt_num(p, 3)


def fmt_pct(x, d: int = 1) -> str:
    return fmt_num(x * 100 if abs(x) <= 1 else x, d)


def is_normal(x) -> bool:
    from scipy import stats
    x = np.asarray(pd.Series(x).dropna(), dtype=float)
    if len(x) < 3 or np.ptp(x) == 0:
        return False
    p = stats.shapiro(x).pvalue if len(x) <= 5000 else stats.normaltest(x).pvalue
    return bool(p >= 0.05)


def summarize(x, normal: bool, d: int = 2) -> str:
    x = pd.Series(x).dropna().astype(float)
    if len(x) == 0:
        return "—"
    if normal:
        return f"{fmt_num(x.mean(), d)} ± {fmt_num(x.std(ddof=1) if len(x) > 1 else 0.0, d)}"
    q1, med, q3 = x.quantile([0.25, 0.5, 0.75])
    return f"{fmt_num(med, d)} ({fmt_num(q1, d)}–{fmt_num(q3, d)})"


def compare_two(d: pd.DataFrame, ycol: str, variables: list, names=("Không mang gen", "Mang gen")) -> pd.DataFrame:
    """Bảng 3.4: so sánh biến định lượng giữa hai nhóm (t độc lập / Mann–Whitney)."""
    from scipy import stats
    g0, g1 = d[d[ycol] == 0], d[d[ycol] == 1]
    rows = []
    for v in variables:
        a, b = g0[v].dropna().astype(float), g1[v].dropna().astype(float)
        dec = VAR_DECIMALS.get(v, 2)
        normal = is_normal(a) and is_normal(b)
        p, test = np.nan, "—"
        if len(a) >= 2 and len(b) >= 2:
            if normal:
                eq = stats.levene(a, b).pvalue >= 0.05
                p = stats.ttest_ind(a, b, equal_var=eq).pvalue
                test = "t độc lập" if eq else "t (Welch)"
            else:
                p = stats.mannwhitneyu(a, b, alternative="two-sided").pvalue
                test = "Mann–Whitney"
        rows.append({"Chỉ số": VAR_LABEL.get(v, v), f"{names[0]} (n = {len(g0)})": summarize(a, normal, dec),
                     f"{names[1]} (n = {len(g1)})": summarize(b, normal, dec),
                     "Trình bày": "TB ± ĐLC" if normal else "Trung vị (Q1–Q3)", "Kiểm định": test,
                     "p": fmt_p(p), "_p": p})
    return pd.DataFrame(rows)


def compare_groups(d: pd.DataFrame, gcol: str, variables: list, groups: list) -> pd.DataFrame:
    """Bảng 3.5: so sánh ≥3 nhóm (ANOVA nếu chuẩn và đồng nhất phương sai, ngược lại Kruskal–Wallis)."""
    from scipy import stats
    rows = []
    sub = d[d[gcol].isin(groups)]
    for v in variables:
        samples = [sub.loc[sub[gcol] == g, v].dropna().astype(float) for g in groups]
        dec = VAR_DECIMALS.get(v, 2)
        normal = all(is_normal(s) for s in samples)
        p, test = np.nan, "—"
        valid = [s for s in samples if len(s) >= 2]
        if len(valid) >= 2:
            if normal and stats.levene(*valid).pvalue >= 0.05:
                p, test = stats.f_oneway(*valid).pvalue, "ANOVA 1 yếu tố"
            else:
                p, test = stats.kruskal(*valid).pvalue, "Kruskal–Wallis"
        row = {"Chỉ số": VAR_LABEL.get(v, v)}
        for g, s in zip(groups, samples):
            row[f"{g} (n = {len(s)})"] = summarize(s, normal, dec)
        row.update({"Trình bày": "TB ± ĐLC" if normal else "Trung vị (Q1–Q3)", "Kiểm định": test, "p": fmt_p(p), "_p": p})
        rows.append(row)
    return pd.DataFrame(rows)


def compare_categorical(d: pd.DataFrame, ycol: str, col: str, order: list | None = None) -> pd.DataFrame:
    """Biến định tính theo nhóm (χ² hoặc Fisher khi bảng 2×2 có tần số kỳ vọng < 5)."""
    from scipy import stats
    sub = d[[ycol, col]].dropna()
    cats = order or sorted(sub[col].unique(), key=str)
    tab = pd.crosstab(sub[col], sub[ycol]).reindex(index=cats, columns=[0, 1], fill_value=0)
    n0, n1 = int(tab[0].sum()), int(tab[1].sum())
    p, test = np.nan, "—"
    if tab.values.sum() > 0 and (tab.sum(axis=1) > 0).sum() >= 2 and n0 > 0 and n1 > 0:
        t = tab[(tab.sum(axis=1) > 0)]
        chi2, p_chi, _, exp = stats.chi2_contingency(t.values)
        if t.shape == (2, 2) and (exp < 5).any():
            p, test = stats.fisher_exact(t.values)[1], "Fisher"
        else:
            p, test = p_chi, "Chi-square"
    rows = [{col: c, f"Không mang gen (n = {n0})": f"{int(tab.loc[c, 0])} ({fmt_pct(tab.loc[c, 0] / n0 if n0 else 0)}%)",
             f"Mang gen (n = {n1})": f"{int(tab.loc[c, 1])} ({fmt_pct(tab.loc[c, 1] / n1 if n1 else 0)}%)"} for c in cats]
    out = pd.DataFrame(rows)
    out["Kiểm định"], out["p"] = test, fmt_p(p)
    return out


def age_group(age: float) -> str:
    if pd.isna(age):
        return "Không rõ"
    return "18–29 tuổi" if age < 30 else "30–39 tuổi" if age < 40 else "40–49 tuổi" if age < 50 else "≥ 50 tuổi"


AGE_GROUPS = ["18–29 tuổi", "30–39 tuổi", "40–49 tuổi", "≥ 50 tuổi"]


def distribution_table(series: pd.Series, order: list, title: str) -> pd.DataFrame:
    """Bảng 3.1–3.3: số lượng (n) và tỷ lệ (%)."""
    s = series.dropna()
    counts = s.value_counts()
    total = int(counts.sum())
    rows = [{title: c, "Số lượng (n)": int(counts.get(c, 0)),
             "Tỷ lệ (%)": fmt_num(counts.get(c, 0) / total * 100 if total else 0, 1)} for c in order]
    rows.append({title: "Tổng", "Số lượng (n)": total, "Tỷ lệ (%)": "100,0"})
    return pd.DataFrame(rows)


# =====================================================================
# 4. HỌC MÁY (MỤC TIÊU 2)
# =====================================================================
ALGOS = ["Logistic Regression", "Random Forest", "XGBoost"]


def library_status() -> dict:
    out = {}
    for name in ("sklearn", "xgboost", "shap", "matplotlib"):
        try:
            __import__(name)
            out[name] = True
        except Exception:
            out[name] = False
    return out


def feature_sets(extras: list | None = None, include_ferritin_in_b: bool = True) -> dict:
    extras = extras or []
    return {"A": CBC7 + extras, "B": CBC7 + (["ferritin"] if include_ferritin_in_b else []) + extras}


def make_model(name: str, seed: int, n_jobs: int, balanced: bool, y_dev=None):
    from sklearn.ensemble import GradientBoostingClassifier, RandomForestClassifier
    from sklearn.impute import SimpleImputer
    from sklearn.linear_model import LogisticRegression
    from sklearn.pipeline import Pipeline
    from sklearn.preprocessing import StandardScaler

    def pipe(clf, scale):
        steps = [("imp", SimpleImputer(strategy="median"))]
        if scale:
            steps.append(("sc", StandardScaler()))
        return Pipeline(steps + [("clf", clf)])

    cw = "balanced" if balanced else None
    if name == "Logistic Regression":
        return pipe(LogisticRegression(max_iter=5000, class_weight=cw), True), {"clf__C": [0.01, 0.1, 1, 10, 100]}, name
    if name == "Random Forest":
        return (pipe(RandomForestClassifier(n_estimators=300, random_state=seed, n_jobs=n_jobs, class_weight=cw), False),
                {"clf__max_depth": [3, 5, None], "clf__min_samples_leaf": [1, 3, 5]}, name)
    try:
        from xgboost import XGBClassifier
        clf = XGBClassifier(n_estimators=200, learning_rate=0.05, max_depth=3, subsample=0.9, eval_metric="logloss",
                            random_state=seed, n_jobs=n_jobs, tree_method="hist", verbosity=0)
        grid = {"clf__max_depth": [2, 3, 4], "clf__n_estimators": [100, 200], "clf__learning_rate": [0.05, 0.1]}
        if balanced and y_dev is not None and np.sum(y_dev) > 0:
            grid["clf__scale_pos_weight"] = [float((len(y_dev) - np.sum(y_dev)) / np.sum(y_dev))]
        return pipe(clf, False), grid, "XGBoost"
    except Exception:  # chưa cài xgboost → dùng Gradient Boosting của scikit-learn và ghi rõ
        clf = GradientBoostingClassifier(random_state=seed)
        return (pipe(clf, False), {"clf__max_depth": [2, 3], "clf__n_estimators": [100, 200], "clf__learning_rate": [0.05, 0.1]},
                "XGBoost (thay bằng GradientBoosting — chưa cài xgboost)")


def metrics_at(y, p, thr: float) -> dict:
    from sklearn.metrics import average_precision_score, brier_score_loss, roc_auc_score
    y = np.asarray(y, dtype=int)
    p = np.asarray(p, dtype=float)
    pred = (p >= thr).astype(int)
    tp = int(((pred == 1) & (y == 1)).sum()); tn = int(((pred == 0) & (y == 0)).sum())
    fp = int(((pred == 1) & (y == 0)).sum()); fn = int(((pred == 0) & (y == 1)).sum())
    div = lambda a, b: a / b if b else float("nan")
    both = len(np.unique(y)) == 2
    return {"Accuracy": div(tp + tn, len(y)), "Sensitivity": div(tp, tp + fn), "Specificity": div(tn, tn + fp),
            "PPV": div(tp, tp + fp), "NPV": div(tn, tn + fn),
            "ROC-AUC": roc_auc_score(y, p) if both else float("nan"),
            "PR-AUC": average_precision_score(y, p) if both else float("nan"),
            "Brier": brier_score_loss(y, p), "TP": tp, "FP": fp, "TN": tn, "FN": fn}


METRIC_KEYS = ["Accuracy", "Sensitivity", "Specificity", "PPV", "NPV", "ROC-AUC", "PR-AUC", "Brier"]


def bootstrap_ci(y, p, thr: float, n_boot: int = 500, seed: int = 42) -> dict:
    rng = np.random.default_rng(seed)
    y = np.asarray(y); p = np.asarray(p)
    rows = []
    for _ in range(n_boot):
        idx = rng.integers(0, len(y), len(y))
        if y[idx].min() == y[idx].max():
            continue
        m = metrics_at(y[idx], p[idx], thr)
        rows.append([m[k] for k in METRIC_KEYS])
    if not rows:
        return {k: (float("nan"), float("nan")) for k in METRIC_KEYS}
    arr = np.array(rows, dtype=float)
    lo, hi = np.nanpercentile(arr, 2.5, axis=0), np.nanpercentile(arr, 97.5, axis=0)
    return {k: (float(lo[i]), float(hi[i])) for i, k in enumerate(METRIC_KEYS)}


def choose_threshold(y, p, target_sens: float) -> float:
    """Chọn ngưỡng trên tập PHÁT TRIỂN (dự đoán chéo): độ đặc hiệu cao nhất mà độ nhạy ≥ mục tiêu; sau đó cố định."""
    from sklearn.metrics import roc_curve
    fpr, tpr, thr = roc_curve(y, p)
    ok = np.where(tpr >= target_sens)[0]
    t = float(thr[ok[0]]) if len(ok) else float(np.min(p))
    return min(max(t, 0.0), 1.0) if np.isfinite(t) else 1.0


def calibration_info(y, p, n_bins: int | None = None) -> dict:
    from sklearn.calibration import calibration_curve
    from sklearn.linear_model import LogisticRegression
    y = np.asarray(y); p = np.clip(np.asarray(p, dtype=float), 1e-6, 1 - 1e-6)
    nb = n_bins or int(min(10, max(3, len(y) // 20)))
    try:
        frac, mean_p = calibration_curve(y, p, n_bins=nb, strategy="quantile")
    except Exception:
        frac, mean_p = np.array([]), np.array([])
    slope = intercept = float("nan")
    if len(np.unique(y)) == 2:
        lr = LogisticRegression(C=1e6, max_iter=1000).fit(np.log(p / (1 - p)).reshape(-1, 1), y)
        slope, intercept = float(lr.coef_[0][0]), float(lr.intercept_[0])
    return {"mean_pred": mean_p, "frac_pos": frac, "slope": slope, "intercept": intercept}


def run_experiment(d: pd.DataFrame, sets: dict, algos: list, test_size: float = 0.3, seed: int = 42, cv: int = 5,
                   target_sens: float = 0.90, n_boot: int = 500, split: str = "random", balanced: bool = False,
                   n_jobs: int = 1, progress=None) -> dict:
    """Huấn luyện/đánh giá các thuật toán cho từng bộ biến (A: CBC; B: CBC + Ferritin) trên CÙNG một lần chia dữ liệu.
    Siêu tham số tối ưu bằng kiểm định chéo trong tập phát triển; ngưỡng cố định trước khi đánh giá trên tập kiểm định."""
    from sklearn.base import clone
    from sklearn.metrics import precision_recall_curve, roc_curve
    from sklearn.model_selection import GridSearchCV, StratifiedKFold, cross_val_predict, train_test_split

    d = d.reset_index(drop=True)
    y = d["y"].astype(int).to_numpy()
    warns = []
    if split == "temporal":
        dev_idx = np.where(d["phase"] == "Hồi cứu")[0]
        test_idx = np.where(d["phase"] == "Tiến cứu")[0]
        if len(dev_idx) == 0 or len(test_idx) == 0:
            raise ValueError("Chia theo thời gian cần cả hồ sơ Hồi cứu (phát triển) và Tiến cứu (kiểm định).")
    else:
        dev_idx, test_idx = train_test_split(np.arange(len(d)), test_size=test_size, stratify=y, random_state=seed)
    ydev, ytest = y[dev_idx], y[test_idx]
    for nm, yy in (("phát triển", ydev), ("kiểm định", ytest)):
        if len(np.unique(yy)) < 2:
            raise ValueError(f"Tập {nm} chỉ có một lớp kết cục — cần cả người mang gen và không mang gen.")
    if min(ydev.sum(), len(ydev) - ydev.sum()) < 10:
        warns.append("Tập phát triển có rất ít ca ở một nhóm (<10): kết quả chỉ mang tính thăm dò.")
    if ytest.sum() < 20:
        warns.append("Tập kiểm định có <20 người mang gen: khoảng tin cậy rất rộng, diễn giải thận trọng.")
    k = int(max(2, min(cv, ydev.sum(), len(ydev) - ydev.sum())))
    skf = StratifiedKFold(n_splits=k, shuffle=True, random_state=seed)

    out = {"meta": {"n_total": len(d), "n_dev": len(dev_idx), "n_test": len(test_idx), "events_dev": int(ydev.sum()),
                    "events_test": int(ytest.sum()), "split": split, "seed": seed, "cv": k, "target_sens": target_sens,
                    "balanced": balanced, "sets": sets, "prevalence_test": float(ytest.mean())},
           "models": {}, "warnings": warns, "y_test": ytest, "dev_idx": dev_idx, "test_idx": test_idx}
    total_steps = len(sets) * len(algos)
    step = 0
    for sname, feats in sets.items():
        Xdev, Xtest = d.loc[dev_idx, feats], d.loc[test_idx, feats]
        for algo in algos:
            est, grid, label = make_model(algo, seed, n_jobs, balanced, ydev)
            gs = GridSearchCV(est, grid, cv=skf, scoring="roc_auc", n_jobs=n_jobs, refit=True)
            gs.fit(Xdev, ydev)
            best = gs.best_estimator_
            oof = cross_val_predict(clone(best), Xdev, ydev, cv=skf, method="predict_proba")[:, 1]
            thr = choose_threshold(ydev, oof, target_sens)
            ptest = best.predict_proba(Xtest)[:, 1]
            fpr, tpr, _ = roc_curve(ytest, ptest)
            prec, rec, _ = precision_recall_curve(ytest, ptest)
            out["models"][f"{sname}|{algo}"] = {
                "set": sname, "algo": algo, "label": label, "features": feats, "pipeline": best,
                "best_params": gs.best_params_, "cv_auc": float(gs.best_score_), "threshold": thr,
                "metrics": metrics_at(ytest, ptest, thr), "ci": bootstrap_ci(ytest, ptest, thr, n_boot, seed),
                "p_test": ptest, "roc": (fpr, tpr), "pr": (rec, prec), "calib": calibration_info(ytest, ptest),
                "train_ranges": {f: (float(Xdev[f].min()), float(Xdev[f].max())) for f in feats},
            }
            step += 1
            if progress:
                progress(step / total_steps, f"{label} — bộ biến {sname}")
    out["test_index"] = test_idx
    return out


def best_key(res: dict, set_name: str = "B") -> str | None:
    cands = [(k, m["metrics"]["ROC-AUC"]) for k, m in res["models"].items() if m["set"] == set_name]
    if not cands:
        cands = [(k, m["metrics"]["ROC-AUC"]) for k, m in res["models"].items()]
    return max(cands, key=lambda t: (-1 if np.isnan(t[1]) else t[1]))[0] if cands else None


def performance_table(res: dict, set_name: str) -> pd.DataFrame:
    """Bảng 3.6 / 3.7."""
    rows = []
    for k, m in res["models"].items():
        if m["set"] != set_name:
            continue
        mt, ci = m["metrics"], m["ci"]
        row = {"Thuật toán": m["label"]}
        for key in ["Accuracy", "Sensitivity", "Specificity", "PPV", "NPV"]:
            row[key + " (%)"] = fmt_num(mt[key] * 100, 1)
        row["ROC-AUC"] = fmt_num(mt["ROC-AUC"], 2)
        row["ROC-AUC (KTC 95%)"] = f"{fmt_num(ci['ROC-AUC'][0], 2)}–{fmt_num(ci['ROC-AUC'][1], 2)}"
        row["PR-AUC"] = fmt_num(mt["PR-AUC"], 2)
        row["Brier"] = fmt_num(mt["Brier"], 3)
        row["Ngưỡng"] = fmt_num(m["threshold"], 3)
        rows.append(row)
    return pd.DataFrame(rows)


def compare_added_value(res: dict, algo: str, n_boot: int = 500, seed: int = 42) -> pd.DataFrame:
    """Bảng 3.8: giá trị bổ sung của Ferritin — mô hình A (CBC) so với B (CBC + Ferritin) trên CÙNG tập kiểm định,
    chênh lệch kèm KTC 95% và p từ bootstrap ghép cặp."""
    a, b = res["models"].get(f"A|{algo}"), res["models"].get(f"B|{algo}")
    if not a or not b:
        return pd.DataFrame()
    y = res["y_test"]
    rng = np.random.default_rng(seed)
    diffs = []
    for _ in range(n_boot):
        idx = rng.integers(0, len(y), len(y))
        if y[idx].min() == y[idx].max():
            continue
        ma = metrics_at(y[idx], a["p_test"][idx], a["threshold"])
        mb = metrics_at(y[idx], b["p_test"][idx], b["threshold"])
        diffs.append([mb[k] - ma[k] for k in METRIC_KEYS])
    diffs = np.array(diffs, dtype=float)
    rows = []
    for i, k in enumerate(METRIC_KEYS):
        va, vb = a["metrics"][k], b["metrics"][k]
        col = diffs[:, i] if len(diffs) else np.array([np.nan])
        lo, hi = np.nanpercentile(col, [2.5, 97.5])
        p = float(min(1.0, 2 * min(np.nanmean(col <= 0), np.nanmean(col >= 0))))
        scale, dec = (100, 1) if k in ("Accuracy", "Sensitivity", "Specificity", "PPV", "NPV") else (1, 2 if k != "Brier" else 3)
        unit = " (%)" if scale == 100 else ""
        rows.append({"Chỉ số": k + unit, "Mô hình CBC": fmt_num(va * scale, dec), "Mô hình CBC + Ferritin": fmt_num(vb * scale, dec),
                     "Chênh lệch": ("+" if vb - va >= 0 else "−") + fmt_num(abs(vb - va) * scale, dec),
                     "KTC 95% chênh lệch": f"{fmt_num(lo * scale, dec)} đến {fmt_num(hi * scale, dec)}", "p (bootstrap)": fmt_p(p)})
    return pd.DataFrame(rows)


def confusion_table(m: dict) -> pd.DataFrame:
    """Bảng 3.9."""
    mt = m["metrics"]
    return pd.DataFrame({"Kết quả dự đoán": ["Nguy cơ mang gen", "Không mang gen", "Tổng"],
                         "Thực tế mang gen": [mt["TP"], mt["FN"], mt["TP"] + mt["FN"]],
                         "Thực tế không mang gen": [mt["FP"], mt["TN"], mt["FP"] + mt["TN"]],
                         "Tổng": [mt["TP"] + mt["FP"], mt["FN"] + mt["TN"], mt["TP"] + mt["FP"] + mt["FN"] + mt["TN"]]})


def feature_importance(res: dict, key: str, d: pd.DataFrame, max_rows: int = 400) -> pd.DataFrame:
    """Bảng 3.10: mức đóng góp tương đối của biến. Dùng SHAP nếu cài được, nếu không dùng permutation importance."""
    m = res["models"][key]
    feats, pipe = m["features"], m["pipeline"]
    idx = res["test_idx"]
    X = d.loc[idx, feats].reset_index(drop=True)
    y = res["y_test"]
    method = None
    imp = None
    try:
        import shap
        Xt = pipe[:-1].transform(X.iloc[:max_rows])
        clf = pipe[-1]
        if m["algo"] == "Logistic Regression":
            sv = shap.LinearExplainer(clf, pipe[:-1].transform(d.loc[res["dev_idx"], feats])).shap_values(Xt)
        else:
            sv = shap.TreeExplainer(clf).shap_values(Xt)
        if isinstance(sv, list):
            sv = sv[-1]
        sv = np.asarray(sv)
        if sv.ndim == 3:
            sv = sv[:, :, -1]
        imp = np.abs(sv).mean(axis=0)
        method = "SHAP (trị tuyệt đối trung bình)"
    except Exception:
        from sklearn.inspection import permutation_importance
        r = permutation_importance(pipe, X, y, scoring="roc_auc", n_repeats=20, random_state=res["meta"]["seed"])
        imp = np.clip(r.importances_mean, 0, None)
        method = "Permutation importance (ΔROC-AUC) — chưa cài SHAP"
    total = float(np.sum(imp)) or 1.0
    df = pd.DataFrame({"Biến đầu vào": [VAR_SHORT.get(f, f) for f in feats], "Mức độ đóng góp tương đối (%)": imp / total * 100})
    df = df.sort_values("Mức độ đóng góp tương đối (%)", ascending=False).reset_index(drop=True)
    df.insert(0, "Xếp hạng", np.arange(1, len(df) + 1))
    df["Mức độ đóng góp tương đối (%)"] = df["Mức độ đóng góp tương đối (%)"].round(1)
    df["Phương pháp"] = method
    return df


# ---------- Đóng gói mô hình có chữ ký ----------
def pack_model(res: dict, key: str, sign_key: str, extra_meta: dict | None = None) -> tuple[bytes, dict]:
    import sklearn
    m = res["models"][key]
    meta = {"key": key, "set": m["set"], "algo": m["label"], "features": m["features"], "threshold": m["threshold"],
            "target_sens": res["meta"]["target_sens"], "metrics": {k: (None if np.isnan(v) else float(v)) for k, v in m["metrics"].items() if k in METRIC_KEYS},
            "n_dev": res["meta"]["n_dev"], "n_test": res["meta"]["n_test"], "events_dev": res["meta"]["events_dev"],
            "split": res["meta"]["split"], "train_ranges": m["train_ranges"], "sklearn": sklearn.__version__,
            "validation": "Kiểm định độc lập theo thời gian (tiến cứu)" if res["meta"]["split"] == "temporal" else "Đánh giá thăm dò (chia ngẫu nhiên)"}
    meta.update(extra_meta or {})
    blob = pickle.dumps({"pipeline": m["pipeline"], "meta": meta}, protocol=4)
    sig = hmac.new(sign_key.encode(), blob, hashlib.sha256).digest()
    return sig + blob, meta


def unpack_model(data: bytes, sign_key: str) -> dict:
    """Chỉ nạp pickle khi chữ ký HMAC khớp (chống mô hình bị thay thế/độc hại)."""
    data = bytes(data)
    sig, blob = data[:32], data[32:]
    if not hmac.compare_digest(sig, hmac.new(sign_key.encode(), blob, hashlib.sha256).digest()):
        raise ValueError("Chữ ký mô hình không hợp lệ — từ chối nạp.")
    return pickle.loads(blob)


def predict_one(bundle: dict, values: dict) -> dict:
    meta, pipe = bundle["meta"], bundle["pipeline"]
    row = {}
    for f in meta["features"]:
        v = values.get(f)
        row[f] = np.nan if v is None else float(v)
    X = pd.DataFrame([row])[meta["features"]]
    prob = float(pipe.predict_proba(X)[0, 1])
    outside = [f for f, (lo, hi) in meta.get("train_ranges", {}).items()
               if f in row and not np.isnan(row[f]) and (row[f] < lo or row[f] > hi)]
    return {"prob": prob, "high": prob >= meta["threshold"], "threshold": meta["threshold"], "outside": outside,
            "missing": [f for f, v in row.items() if np.isnan(v)]}


# =====================================================================
# 5. HÌNH VẼ (matplotlib)
# =====================================================================
def _style(ax, fg):
    for s in ax.spines.values():
        s.set_color(fg)
    ax.tick_params(colors=fg)
    ax.xaxis.label.set_color(fg); ax.yaxis.label.set_color(fg); ax.title.set_color(fg)
    ax.grid(alpha=0.2, color=fg)


def _new_fig(w=5.2, h=4.2, fg="#1d1d1f"):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots(figsize=(w, h))
    fig.patch.set_alpha(0); ax.set_facecolor("none")
    return fig, ax


def fig_curves(res: dict, kind: str, set_name: str, fg="#1d1d1f"):
    fig, ax = _new_fig(fg=fg)
    for k, m in res["models"].items():
        if m["set"] != set_name:
            continue
        if kind == "roc":
            ax.plot(*m["roc"], label=f"{m['label'].split(' (')[0]} (AUC {fmt_num(m['metrics']['ROC-AUC'], 2)})")
        elif kind == "pr":
            ax.plot(*m["pr"], label=f"{m['label'].split(' (')[0]} (AP {fmt_num(m['metrics']['PR-AUC'], 2)})")
        else:
            c = m["calib"]
            if len(c["mean_pred"]):
                ax.plot(c["mean_pred"], c["frac_pos"], marker="o", label=m["label"].split(" (")[0])
    if kind == "roc":
        ax.plot([0, 1], [0, 1], "--", color=fg, alpha=0.4); ax.set_xlabel("1 − Specificity"); ax.set_ylabel("Sensitivity"); ax.set_title(f"Đường cong ROC — bộ biến {set_name}")
    elif kind == "pr":
        ax.set_xlabel("Recall (Sensitivity)"); ax.set_ylabel("Precision (PPV)"); ax.set_title(f"Đường cong Precision–Recall — bộ biến {set_name}")
    else:
        ax.plot([0, 1], [0, 1], "--", color=fg, alpha=0.4); ax.set_xlabel("Xác suất dự đoán"); ax.set_ylabel("Tỷ lệ mang gen quan sát"); ax.set_title(f"Hiệu chỉnh (calibration) — bộ biến {set_name}")
    _style(ax, fg)
    leg = ax.legend(frameon=False, fontsize=8)
    for t in leg.get_texts():
        t.set_color(fg)
    fig.tight_layout()
    return fig


def fig_boxplots(d: pd.DataFrame, variables: list, ycol: str = "y", fg="#1d1d1f", ncols: int = 4):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    n = len(variables)
    nrows = math.ceil(n / ncols)
    fig, axes = plt.subplots(nrows, ncols, figsize=(3.0 * ncols, 3.0 * nrows))
    fig.patch.set_alpha(0)
    axes = np.atleast_1d(axes).ravel()
    for ax, v in zip(axes, variables):
        data = [d.loc[d[ycol] == 0, v].dropna(), d.loc[d[ycol] == 1, v].dropna()]
        ax.set_facecolor("none")
        bp = ax.boxplot(data, tick_labels=["Không", "Mang gen"], patch_artist=True, widths=0.55)
        for patch, col in zip(bp["boxes"], ["#4F7CFF", "#C0394F"]):
            patch.set_facecolor(col); patch.set_alpha(0.55)
        for el in ("whiskers", "caps", "medians"):
            for line in bp[el]:
                line.set_color(fg)
        ax.set_title(VAR_SHORT.get(v, v), fontsize=10)
        _style(ax, fg)
    for ax in axes[n:]:
        ax.axis("off")
    fig.tight_layout()
    return fig


def fig_confusion(m: dict, fg="#1d1d1f"):
    fig, ax = _new_fig(4.2, 3.6, fg)
    mt = m["metrics"]
    mat = np.array([[mt["TP"], mt["FP"]], [mt["FN"], mt["TN"]]])
    ax.imshow(mat, cmap="Blues", alpha=0.6)
    for (i, j), v in np.ndenumerate(mat):
        ax.text(j, i, str(v), ha="center", va="center", fontsize=16, color="#1d1d1f", weight="bold")
    ax.set_xticks([0, 1], ["Thực tế\nmang gen", "Thực tế\nkhông mang gen"]); ax.set_yticks([0, 1], ["Dự đoán\nnguy cơ", "Dự đoán\nkhông"])
    _style(ax, fg); ax.grid(False)
    fig.tight_layout()
    return fig


def fig_bar(df: pd.DataFrame, label_col: str, value_col: str, title: str, fg="#1d1d1f"):
    fig, ax = _new_fig(5.2, 3.4, fg)
    d = df.iloc[::-1]
    ax.barh(d[label_col], d[value_col], color="#C0394F", alpha=0.8)
    ax.set_title(title); ax.set_xlabel("%")
    _style(ax, fg)
    fig.tight_layout()
    return fig


# =====================================================================
# 6. NHẬP DỮ LIỆU TỪ FILE (mã hóa định danh, quy ước mã, đơn vị)
# =====================================================================
def _norm(text) -> str:
    text = str(text or "").replace("đ", "d").replace("Đ", "D")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]", "", text.lower())


def _s(v):
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    t = str(v).strip()
    return t if t and t.lower() not in {"nan", "none", "nat"} else None


def _f(v):
    t = _s(v)
    if t is None:
        return None
    try:
        return float(str(t).replace(",", "."))
    except ValueError:
        return None


def code_key(v):
    t = _s(v)
    if t is None:
        return None
    if isinstance(v, bool):
        return str(int(v))
    if isinstance(v, numbers.Number):
        f = float(v)
        return str(int(f)) if f.is_integer() else str(f)
    try:
        f = float(t.replace(",", "."))
        return str(int(f)) if f.is_integer() else str(f)
    except ValueError:
        return t.lower()


def _is_numeric_key(k) -> bool:
    try:
        float(k)
        return True
    except (TypeError, ValueError):
        return False


def pseudonymize(raw, key: str) -> str:
    """Mã hóa mã bệnh nhân gốc thành mã nghiên cứu không đảo ngược (HMAC-SHA256)."""
    return "TL-" + hmac.new(key.encode(), str(raw).strip().encode(), hashlib.sha256).hexdigest()[:8].upper()


# --- chuẩn hóa giá trị chữ ---
def norm_sex(v):
    n = _norm(_s(v))
    return "Nam" if n in {"nam", "male", "m", "trai"} else "Nữ" if n in {"nu", "female", "f", "gai"} else None


def norm_yn(v):
    n = _norm(_s(v))
    if n in {"khongro", "khongbiet", "kb", "unknown"}:
        return "Không rõ"
    if n in {"co", "c", "yes", "y", "x", "true", "dung", "duongtinh"}:
        return "Có"
    if n in {"khong", "k", "no", "n", "false", "sai", "amtinh"}:
        return "Không"
    return None


def norm_gene(v):
    n = _norm(_s(v))
    if not n:
        return None
    if "chua" in n or "khongxn" in n:
        return "Chưa làm"
    if "dongmang" in n or "alphabeta" in n or "alphavabeta" in n or ("alpha" in n and "beta" in n):
        return "Đồng mang alpha và beta"
    if "khongphathien" in n or "khongmang" in n or "amtinh" in n or "binhthuong" in n or n in {"normal", "negative", "neg"}:
        return "Không phát hiện biến thể"
    if "alpha" in n or "hba1" in n or "hba2gen" in n:
        return "Alpha-thalassemia"
    if "beta" in n or "hbb" in n:
        return "Beta-thalassemia"
    return None


def norm_hplc(v):
    n = _norm(_s(v))
    if not n:
        return None
    if "chua" in n:
        return "Chưa làm"
    if "binhthuong" in n or n == "normal":
        return "Bình thường"
    if "beta" in n or "hba2tang" in n:
        return "Gợi ý beta-thalassemia"
    if "hbe" in n or "hbc" in n or "hbs" in n or "hemoglobinopathy" in n:
        return "HbE / hemoglobinopathy khác"
    if "batthuong" in n or "abnormal" in n:
        return "Bất thường khác"
    return None


def norm_phase(v):
    n = _norm(_s(v))
    return "Hồi cứu" if "hoicuu" in n or "retro" in n else "Tiến cứu" if "tiencuu" in n or "prospective" in n else None


IMPORT_ALIASES = {
    "study_code": {"manghiencuu", "mann", "studycode", "studyid", "mahoso"},
    "id_source": {"mabenhnhan", "mabn", "mabenhan", "mavaovien", "patientid", "idbenhnhan", "id", "sohoso"},
    "phase": {"giaidoan", "phase", "nhom"},
    "sample_date": {"ngayxn", "ngayxncbc", "ngaylaymau", "ngayxetnghiem", "sampledate", "ngaycbc"},
    "age": {"tuoi", "age", "tuoinam"},
    "birth": {"ngaysinh", "namsinh", "dob", "birthdate", "birthyear"},
    "sex": {"gioitinh", "gioi", "sex", "gender", "gt"},
    "rbc": {"rbc", "hongcau", "slhc", "rbctl"}, "hgb": {"hgb", "hb", "hemoglobin", "haemoglobin", "huyetsacto", "hbgl", "hbgdl"},
    "hct": {"hct", "hematocrit", "haematocrit"}, "mcv": {"mcv", "mcvfl"}, "mch": {"mch", "mchpg"}, "mchc": {"mchc", "mchcgl"},
    "rdw": {"rdw", "rdwcv"}, "wbc": {"wbc", "bachcau", "slbc"}, "plt": {"plt", "tieucau", "platelet", "platelets"},
    "ferritin": {"ferritin", "ferritinngml"}, "ferritin_date": {"ngayferritin", "ngayxnferritin"},
    "crp": {"crp", "crpmgl"}, "inflammation": {"tinhtrangviem", "viem", "inflammation"},
    "transfusion": {"tiensutruyenmau", "truyenmau", "transfusion"}, "transfusion_days": {"songaytruyenmau", "songaytulantruyengannhat", "transfusiondays"},
    "gene_status": {"ketquagen", "gen", "pcr", "ketquapcr", "ketquaditruyen", "xetnghiemgen", "genstatus", "ketquagenhba1hba2hbb"},
    "gene_detail": {"chitietbienthe", "bienthe", "genedetail", "kieugen"}, "gene_date": {"ngaygen", "ngayxngen", "ngaypcr"},
    "hplc_status": {"hplc", "diendi", "ketluanhplc", "hplcdiendi", "ketquahplc"},
    "hba2": {"hba2"}, "hbf": {"hbf"}, "hbe": {"hbe"}, "hplc_date": {"ngayhplc", "ngaydiendi"},
    "analyzer": {"mayhuyethoc", "mayxn", "analyzer"}, "ferritin_system": {"hethongferritin", "mayferritin"},
    "consent_type": {"hinhthucdongthuan", "dongthuan"}, "note": {"ghichu", "note", "notes"},
}
IDENTIFIER_ALIASES = {"hoten", "hovaten", "ten", "name", "tenbenhnhan", "sodienthoai", "sdt", "dienthoai", "phone", "diachi",
                      "address", "cccd", "cmnd", "bhyt", "sobhyt", "email"}

CB_YN_PRESETS = {
    "1 = Có, 2 = Không, 9 = Không rõ": {"1": "Có", "2": "Không", "9": "Không rõ"},
    "1 = Có, 0 = Không, 9 = Không rõ": {"1": "Có", "0": "Không", "9": "Không rõ"},
    "0 = Không, 1 = Có, 2 = Không rõ": {"0": "Không", "1": "Có", "2": "Không rõ"},
}
CB_SEX_PRESETS = {
    "1 = Nam, 2 = Nữ": {"1": "Nam", "2": "Nữ"}, "0 = Nữ, 1 = Nam": {"0": "Nữ", "1": "Nam"},
    "0 = Nam, 1 = Nữ": {"0": "Nam", "1": "Nữ"}, "1 = Nữ, 2 = Nam": {"1": "Nữ", "2": "Nam"},
}


def default_codebook() -> dict:
    return {
        "sex": dict(CB_SEX_PRESETS["1 = Nam, 2 = Nữ"]), "yn": dict(CB_YN_PRESETS["1 = Có, 2 = Không, 9 = Không rõ"]),
        "gene": {"0": "Không phát hiện biến thể", "1": "Alpha-thalassemia", "2": "Beta-thalassemia", "3": "Đồng mang alpha và beta", "9": "Chưa làm"},
        "hplc": {"0": "Bình thường", "1": "Gợi ý beta-thalassemia", "2": "HbE / hemoglobinopathy khác", "3": "Bất thường khác", "9": "Chưa làm"},
        "phase": {"1": "Hồi cứu", "2": "Tiến cứu"},
        "num_missing": set(), "manual": {},
    }


def decode_cat(value, table: dict, text_fn, unmapped: dict, group: str):
    """Giá trị số → tra bảng quy ước (mã lạ → bỏ trống + ghi nhận); giá trị chữ → nhận diện tự động."""
    k = code_key(value)
    if k is None:
        return None
    if k in table:
        return table[k]
    if _is_numeric_key(k):
        unmapped.setdefault(group, {}).setdefault(k, 0)
        unmapped[group][k] += 1
        return None
    return text_fn(value)


def to_standard_units(rec: dict) -> dict:
    """Quy đổi về đơn vị của đề cương: HGB g/L, MCHC g/L, HCT %, RBC T/L, WBC/PLT G/L."""
    notes = []
    def conv(key, cond, fn, text):
        v = rec.get(key)
        if v is not None and not (isinstance(v, float) and math.isnan(v)) and cond(v):
            rec[key] = fn(v); notes.append(text)
    conv("hgb", lambda v: v < 30, lambda v: v * 10, "HGB g/dL→g/L")
    conv("mchc", lambda v: v < 100, lambda v: v * 10, "MCHC g/dL→g/L")
    conv("hct", lambda v: 0 < v < 1, lambda v: v * 100, "HCT tỷ lệ→%")
    conv("rbc", lambda v: v > 1000, lambda v: v / 1_000_000, "RBC /µL→T/L")
    conv("wbc", lambda v: v > 500, lambda v: v / 1000, "WBC /µL→G/L")
    conv("plt", lambda v: v > 5000, lambda v: v / 1000, "PLT /µL→G/L")
    return rec, notes


def _decode_csv(data: bytes) -> pd.DataFrame:
    for enc in ("utf-8-sig", "utf-16", "cp1258", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(data), sep=None, engine="python", encoding=enc)
        except UnicodeError:
            continue
    raise ValueError("Không đọc được bảng văn bản (sai mã hóa ký tự).")


def _docx_tables(data: bytes):
    from docx import Document
    out = []
    for ti, table in enumerate(Document(io.BytesIO(data)).tables, 1):
        grid = [[c.text.strip() for c in row.cells] for row in table.rows]
        if len(grid) < 2:
            continue
        seen, header = {}, []
        for i, h in enumerate(grid[0]):
            h = h or f"cot_{i}"
            k = seen.get(h, 0); seen[h] = k + 1
            header.append(h if k == 0 else f"{h}_{k}")
        out.append({"sheet": f"Bảng Word {ti}", "df": pd.DataFrame(grid[1:], columns=header), "labels": {}})
    return out


def docx_text(data: bytes) -> str:
    from docx import Document
    doc = Document(io.BytesIO(data))
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            cells = []
            for c in row.cells:
                t = c.text.strip()
                if t and (not cells or cells[-1] != t):
                    cells.append(t)
            if cells:
                parts.append(" | ".join(cells))
    return "\n".join(parts)


def read_tables(name: str, data: bytes):
    """Đọc Excel/CSV/JSON/Parquet/Word-bảng/Stata/SAS/SPSS/R → danh sách {sheet, df, labels}."""
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
        if ext == ".docx":
            return _docx_tables(data)
        if ext == ".dta":
            with pd.read_stata(bio, iterator=True) as r:
                df = r.read()
            return [{"sheet": "", "df": df, "labels": {}}]
        if ext in (".sas7bdat", ".xpt"):
            fmt = "xport" if ext == ".xpt" else "sas7bdat"
            try:
                df = pd.read_sas(io.BytesIO(data), format=fmt, encoding="utf-8")
            except UnicodeDecodeError:
                df = pd.read_sas(io.BytesIO(data), format=fmt, encoding="latin-1")
            return [{"sheet": "", "df": df, "labels": {}}]
        if ext in (".sav", ".zsav", ".por", ".rds", ".rdata", ".rda"):
            with tempfile.TemporaryDirectory() as tmp:  # ghi tạm rồi xóa ngay
                path = os.path.join(tmp, "data" + ext)
                with open(path, "wb") as fh:
                    fh.write(data)
                if ext in (".sav", ".zsav", ".por"):
                    import pyreadstat
                    df, _ = (pyreadstat.read_por(path) if ext == ".por" else pyreadstat.read_sav(path, apply_value_formats=True))
                    return [{"sheet": "", "df": df, "labels": {}}]
                import pyreadr
                return [{"sheet": str(k or ""), "df": v, "labels": {}} for k, v in pyreadr.read_r(path).items() if hasattr(v, "columns")]
    except ImportError as exc:
        raise RuntimeError(f"Thiếu thư viện để đọc định dạng {ext}: {exc.name or exc}. Hãy thêm vào requirements.txt.") from exc
    raise RuntimeError("Định dạng không được hỗ trợ")


def _age_from_birth(v):
    try:
        if v is None or pd.isna(v):
            return None
        if isinstance(v, numbers.Number) and not isinstance(v, bool):
            y = int(v)
            a = date.today().year - y if 1900 <= y <= date.today().year else None
        else:
            ts = pd.to_datetime(v, errors="coerce", dayfirst=True)
            if pd.isna(ts):
                return None
            t0 = date.today()
            a = t0.year - ts.year - ((t0.month, t0.day) < (ts.month, ts.day))
        return float(a) if a is not None and 0 <= a <= 120 else None
    except Exception:
        return None


def records_from_dataframe(df: pd.DataFrame, source: str, cb: dict, default_phase: str, pseudo_key: str, auto_start: int = 1):
    """Biến một bảng bất kỳ thành các bản ghi theo lược đồ nghiên cứu.
    - Cột định danh (họ tên, SĐT, địa chỉ, CCCD…) bị BỎ QUA; mã bệnh nhân gốc được mã hóa HMAC thành mã nghiên cứu.
    - Giá trị mã số đọc theo bảng quy ước `cb`; mã lạ để trống và báo lại."""
    mapping, ignored = {}, []
    manual = cb.get("manual") or {}
    for col in df.columns:
        key = _norm(col)
        if key in IDENTIFIER_ALIASES:
            ignored.append(str(col))
            continue
        tgt = manual.get(key)
        if tgt in STUDY_COLUMNS + ["id_source", "birth"] and tgt not in mapping:
            mapping[tgt] = col
            continue
        for target, aliases in IMPORT_ALIASES.items():
            if target not in mapping and (key in aliases or key == _norm(target)):
                mapping[target] = col
                break
    notes = []
    if ignored:
        notes.append(f"{source}: đã BỎ QUA {len(ignored)} cột định danh ({', '.join(ignored[:6])}).")
    data_targets = [t for t in mapping if t in LAB_COLS + ["gene_status", "hplc_status"]]
    if not data_targets:
        return [], notes + [f"{source}: không nhận diện được cột xét nghiệm nào (CBC/Ferritin/kết quả gen)."]
    unmapped, recs, n_auto = {}, [], auto_start
    for _, r in df.iterrows():
        rec = {c: None for c in STUDY_COLUMNS}
        for tgt, col in mapping.items():
            v = r[col]
            if tgt in NUMERIC_COLS:
                f = _f(v)
                rec[tgt] = None if f is not None and f in cb["num_missing"] else f
            elif tgt == "sex":
                rec[tgt] = decode_cat(v, cb["sex"], norm_sex, unmapped, "Giới tính")
            elif tgt in ("inflammation", "transfusion"):
                rec[tgt] = decode_cat(v, cb["yn"], norm_yn, unmapped, "Có/Không")
            elif tgt == "gene_status":
                rec[tgt] = decode_cat(v, cb["gene"], norm_gene, unmapped, "Kết quả gen")
            elif tgt == "hplc_status":
                rec[tgt] = decode_cat(v, cb["hplc"], norm_hplc, unmapped, "HPLC/điện di")
            elif tgt == "phase":
                rec[tgt] = decode_cat(v, cb["phase"], norm_phase, unmapped, "Giai đoạn")
            elif tgt in DATE_COLS:
                ts = pd.to_datetime(v, errors="coerce", dayfirst=True) if _s(v) else pd.NaT
                rec[tgt] = None if pd.isna(ts) else ts.strftime("%Y-%m-%d")
            elif tgt in STUDY_COLUMNS:
                rec[tgt] = _s(v)
        if rec["age"] is None and "birth" in mapping:
            rec["age"] = _age_from_birth(r[mapping["birth"]])
        if not any(rec[t] is not None for t in data_targets):
            continue
        if not rec.get("study_code"):
            raw = r[mapping["id_source"]] if "id_source" in mapping else None
            if _s(raw):
                rec["study_code"] = pseudonymize(raw, pseudo_key)
            else:
                rec["study_code"] = f"TL-{n_auto:05d}"; n_auto += 1
        rec["phase"] = rec.get("phase") or default_phase
        rec, conv_notes = to_standard_units(rec)
        recs.append(rec)
    for group, counter in unmapped.items():
        ex = ", ".join(f"{k} (×{n})" for k, n in sorted(counter.items(), key=lambda t: -t[1])[:6])
        notes.append(f"{source}: {group} có mã KHÔNG nằm trong bảng quy ước → để trống: {ex}.")
    return recs, notes


def validate_records(df: pd.DataFrame):
    """Kiểm tra nhanh trước khi lưu: mã trùng trong file, giá trị ngoài khoảng hợp lý."""
    issues = []
    dup = df["study_code"][df["study_code"].duplicated(keep=False)].unique().tolist()
    if dup:
        issues.append(f"{len(dup)} mã nghiên cứu bị trùng trong file (chỉ lưu bản ghi đầu tiên): {', '.join(map(str, dup[:5]))}…")
    d = prepare(df)
    for k, (lo, hi) in PLAUSIBLE.items():
        n = int(((d[k] < lo) | (d[k] > hi)).sum())
        if n:
            issues.append(f"{n} giá trị {VAR_LABEL.get(k, k)} ngoài khoảng hợp lý ({lo}–{hi}) — kiểm tra đơn vị/số liệu gốc.")
    return issues


# ---------- AI đọc phiếu xét nghiệm (tùy chọn) ----------
AI_PROMPT = """Bạn là công cụ trích xuất số liệu từ phiếu xét nghiệm (CBC/công thức máu, Ferritin, HPLC/điện di huyết sắc tố, xét nghiệm gen HBA1/HBA2/HBB).
QUY TẮC: KHÔNG trích xuất/nhắc lại họ tên, ngày sinh, SĐT, địa chỉ, mã bệnh nhân. Chỉ lấy KẾT QUẢ của người bệnh (không lấy khoảng tham chiếu). Không thấy thì null; tuyệt đối không đoán hay tự tính.
Ghi đơn vị đúng như trên phiếu. Ngày theo dạng YYYY-MM-DD. gene_status: "Không phát hiện biến thể" | "Alpha-thalassemia" | "Beta-thalassemia" | "Đồng mang alpha và beta" | null. hplc_status: "Bình thường" | "Gợi ý beta-thalassemia" | "HbE / hemoglobinopathy khác" | "Bất thường khác" | null.
Chỉ trả về MỘT đối tượng JSON, không thêm chữ nào khác:
{"age_years":null,"sex":null,"sample_date":null,"rbc":null,"rbc_unit":null,"hgb":null,"hgb_unit":null,"hct":null,"mcv":null,"mch":null,"mchc":null,"mchc_unit":null,"rdw":null,"wbc":null,"plt":null,
 "ferritin":null,"ferritin_date":null,"crp":null,"hba2":null,"hbf":null,"hbe":null,"hplc_status":null,"hplc_date":null,"gene_status":null,"gene_detail":null,"gene_date":null,"analyzer":null,"legible":true,"notes":""}"""


def ai_request(user_content: list, api_key: str, model: str, max_tokens: int = 1500) -> dict:
    import requests
    if not api_key:
        raise RuntimeError("Chưa cấu hình ANTHROPIC_API_KEY trong Streamlit Secrets.")
    try:
        resp = requests.post("https://api.anthropic.com/v1/messages",
                             headers={"x-api-key": api_key, "anthropic-version": "2023-06-01", "content-type": "application/json"},
                             json={"model": model, "max_tokens": max_tokens, "temperature": 0,
                                   "messages": [{"role": "user", "content": user_content}]}, timeout=120)
    except requests.RequestException as exc:
        raise RuntimeError(f"Không kết nối được dịch vụ AI: {type(exc).__name__}") from exc
    if resp.status_code != 200:
        try:
            detail = resp.json().get("error", {}).get("message", "")
        except Exception:
            detail = ""
        raise RuntimeError(f"Dịch vụ AI trả lỗi {resp.status_code}. {detail[:200]}")
    text = "".join(b.get("text", "") for b in resp.json().get("content", []) if b.get("type") == "text")
    m = re.search(r"\{.*\}", text, flags=re.DOTALL)
    if not m:
        raise RuntimeError("AI không trả về dữ liệu đúng định dạng.")
    return json.loads(m.group(0))


def ai_extract_file(data: bytes, mime: str, api_key: str, model: str) -> dict:
    kind = "document" if mime == "application/pdf" else "image"
    return ai_request([{"type": kind, "source": {"type": "base64", "media_type": mime,
                                                 "data": base64.standard_b64encode(data).decode("ascii")}},
                       {"type": "text", "text": AI_PROMPT}], api_key, model)


def ai_extract_text(text: str, api_key: str, model: str) -> dict:
    return ai_request([{"type": "text", "text": 'NỘI DUNG TÀI LIỆU:\n"""\n' + text[:60000] + '\n"""\n\n' + AI_PROMPT}], api_key, model)


def case_from_ai(ai: dict) -> dict:
    """Chuyển kết quả AI thành bản ghi theo lược đồ (đã quy đổi đơn vị)."""
    rec = {c: None for c in STUDY_COLUMNS}
    for k in ["rbc", "hgb", "hct", "mcv", "mch", "mchc", "rdw", "wbc", "plt", "ferritin", "crp", "hba2", "hbf", "hbe"]:
        rec[k] = _f(ai.get(k))
    rec["age"] = _f(ai.get("age_years"))
    rec["sex"] = norm_sex(ai.get("sex"))
    rec["gene_status"] = norm_gene(ai.get("gene_status")) if ai.get("gene_status") else None
    rec["hplc_status"] = norm_hplc(ai.get("hplc_status")) if ai.get("hplc_status") else None
    rec["gene_detail"] = (_s(ai.get("gene_detail")) or "")[:300] or None
    for k in ("sample_date", "ferritin_date", "gene_date", "hplc_date"):
        ts = pd.to_datetime(ai.get(k), errors="coerce") if _s(ai.get(k)) else pd.NaT
        rec[k] = None if pd.isna(ts) else ts.strftime("%Y-%m-%d")
    rec["analyzer"] = _s(ai.get("analyzer"))
    rec, _ = to_standard_units(rec)
    return rec
