# -*- coding: utf-8 -*-
"""
Streamlit app: Thẩm định phương án kinh doanh/ sử dụng vốn (pasdv.docx)
"""
import io
import os
import re
import math
import json
import zipfile
import datetime as dt
from typing import Dict, Any, Tuple, Optional

import numpy as np
import pandas as pd
import streamlit as st

# Docx parsing
try:
    from docx import Document
except Exception:
    Document = None

# Gemini
try:
    import google.generativeai as genai
except Exception:
    genai = None

# Plotly cho biểu đồ
try:
    import plotly.graph_objects as go
    import plotly.express as px
except Exception:
    go = None
    px = None

st.set_page_config(page_title="PASDV - Thẩm định phương án", page_icon="💼", layout="wide")


# ========================== Helpers ==========================
FIELD_DEFAULTS = {
    "ten_khach_hang": "",
    "cccd": "",
    "noi_cu_tru": "",
    "so_dien_thoai": "",
    "muc_dich_vay": "",
    "tong_nhu_cau_von": 0.0,
    "von_doi_ung": 0.0,
    "so_tien_vay": 0.0,
    "lai_suat_nam": 10.0,
    "thoi_gian_vay_thang": 12,
    "ky_han_tra": "Tháng",
    "thu_nhap_thang": 0.0,
    "gia_tri_tsdb": 0.0,
    "tong_no_hien_tai": 0.0,
    "loi_nhuan_rong_nam": 0.0,
    "tong_von_dau_tu": 0.0,
}

def vnd_to_float(s: str) -> float:
    """Chuyển chuỗi tiền tệ VN về float (hỗ trợ dấu . ngăn cách, , thập phân)."""
    if s is None:
        return 0.0
    s = str(s)
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." not in s:
        s = s.replace(".", "")
        s = s.replace(",", ".")
    else:
        s = s.replace(".", "")

    s = s.replace("đ", "").replace("VND", "").replace("vnđ", "").replace("₫", "").replace(" ", "")
    s = re.sub(r"[^\d\.\-]", "", s)
    try:
        return float(s) if s else 0.0
    except Exception:
        return 0.0

def format_vnd(amount: float) -> str:
    """Định dạng tiền VND: 1.234.567"""
    try:
        return f"{float(amount):,.0f}".replace(",", ".")
    except Exception:
        return "0"

def format_vnd_float(amount: float) -> str:
    """Định dạng số thập phân kiểu VN: 1.234.567,89"""
    try:
        s = f"{float(amount):,.2f}"
        s = s.replace(",", "_").replace(".", ",").replace("_", ".")
        return s
    except Exception:
        return "0,00"

def percent_to_float(s: str) -> float:
    """Chuyển đổi chuỗi phần trăm sang số float; chấp nhận '8,5' hoặc '8.5'."""
    if s is None:
        return 0.0
    s = str(s).replace(",", ".")
    m = re.search(r"(\d+(?:\.\d+)?)", s)
    return float(m.group(1)) if m else 0.0

def vn_money_input(label: str, value: float, key: Optional[str] = None, help: Optional[str] = None) -> float:
    """Ô nhập tiền tệ kiểu VN: hiển thị 1.234.567 và parse lại về float."""
    raw = st.text_input(label, value=format_vnd(value), key=key, help=help)
    return float(vnd_to_float(raw))

def vn_percent_input(label: str, value: float, key: Optional[str] = None, help: Optional[str] = None) -> float:
    """Ô nhập phần trăm linh hoạt: cho phép nhập '8,5' hoặc '8.5'."""
    shown = f"{float(value):.2f}".replace(".", ",")
    raw = st.text_input(label, value=shown, key=key, help=help)
    return percent_to_float(raw)

def extract_from_docx(file_bytes: bytes) -> Dict[str, Any]:
    """Đọc .docx PASDV và trích xuất thông tin theo cấu trúc thực tế."""
    data = FIELD_DEFAULTS.copy()
    if Document is None:
        return data

    bio = io.BytesIO(file_bytes)
    doc = Document(bio)
    full_text = "\n".join([p.text for p in doc.paragraphs])

    lines = [line.strip() for line in full_text.split('\n') if line.strip()]
    full_text = "\n".join(lines)

    # === 1. THÔNG TIN KHÁCH HÀNG ===
    ten_pattern1 = r"(?:\d+\.\s*)?Họ\s+và\s+tên\s*[:：]\s*([A-ZÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬĐÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴ][a-zàáảãạăằắẳẵặâầấẨẫậđèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵA-ZÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬĐÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴ\s]+)"
    m = re.search(ten_pattern1, full_text, flags=re.IGNORECASE)
    if m:
        data["ten_khach_hang"] = m.group(1).strip()
    else:
        ten_pattern2 = r"(?:Ông|Bà)\s*\((?:bà|ông)\)\s*[:：]\s*([A-ZÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬĐÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴ][a-zàáảãạăằắẳẵặâầấẨẫậđèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵA-ZÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬĐÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴ\s]+)"
        m = re.search(ten_pattern2, full_text, flags=re.IGNORECASE)
        if m:
            data["ten_khach_hang"] = m.group(1).strip()
    
    cccd_pattern = r"(?:CMND|CCCD)(?:\/(?:CCCD|CMND))?(?:\/hộ\s*chiếu)?\s*[:：]\s*(\d{9,12})"
    m = re.search(cccd_pattern, full_text, flags=re.IGNORECASE)
    if m:
        data["cccd"] = m.group(1).strip()

    noi_cu_tru_pattern = r"Nơi\s*cư\s*trú\s*[:：]\s*([^\n]+?)(?=\n|Số\s*điện\s*thoại|$)"
    m = re.search(noi_cu_tru_pattern, full_text, flags=re.IGNORECASE | re.DOTALL)
    if m:
        data["noi_cu_tru"] = m.group(1).strip()

    sdt_pattern = r"Số\s*điện\s*thoại\s*[:：]\s*(0\d{9,10})"
    m = re.search(sdt_pattern, full_text, flags=re.IGNORECASE)
    if m:
        data["so_dien_thoai"] = m.group(1).strip()

    # === 2. PHƯƠNG ÁN SỬ DỤNG VỐN ===
    muc_dich_pattern1 = r"Mục\s*đích\s*vay\s*[:：]\s*([^\n]+)"
    m = re.search(muc_dich_pattern1, full_text, flags=re.IGNORECASE)
    if m:
        data["muc_dich_vay"] = m.group(1).strip()
    else:
        muc_dich_pattern2 = r"Vốn\s*vay\s*Agribank.*?[:：].*?(?:Thực\s*hiện|Sử\s*dụng\s*vào)\s*([^\n]+)"
        m = re.search(muc_dich_pattern2, full_text, flags=re.IGNORECASE | re.DOTALL)
        if m:
            data["muc_dich_vay"] = m.group(1).strip()[:200]

    tnc_pattern = r"(?:Tổng\s*nhu\s*cầu\s*vốn|1\.\s*Tổng\s*nhu\s*cầu\s*vốn)\s*[:：]\s*([\d\.,]+)"
    m = re.search(tnc_pattern, full_text, flags=re.IGNORECASE)
    if m:
        data["tong_nhu_cau_von"] = vnd_to_float(m.group(1))

    von_du_pattern = r"Vốn\s*đối\s*ứng\s*(?:tham\s*gia)?[^\d]*([\d\.,]+)\s*đồng"
    m = re.search(von_du_pattern, full_text, flags=re.IGNORECASE)
    if m:
        data["von_doi_ung"] = vnd_to_float(m.group(1))

    so_tien_vay_pattern = r"Vốn\s*vay\s*Agribank\s*(?:số\s*tiền)?[:\s]*([\d\.,]+)\s*đồng"
    m = re.search(so_tien_vay_pattern, full_text, flags=re.IGNORECASE)
    if m:
        data["so_tien_vay"] = vnd_to_float(m.group(1))

    thoi_han_pattern = r"Thời\s*hạn\s*vay\s*[:：]\s*(\d+)\s*tháng"
    m = re.search(thoi_han_pattern, full_text, flags=re.IGNORECASE)
    if m:
        data["thoi_gian_vay_thang"] = int(m.group(1))

    lai_suat_pattern = r"Lãi\s*suất\s*[:：]\s*([\d\.,]+)\s*%"
    m = re.search(lai_suat_pattern, full_text, flags=re.IGNORECASE)
    if m:
        data["lai_suat_nam"] = percent_to_float(m.group(1))

    # === 3. NGUỒN TRẢ NỢ & THU NHẬP ===
    thu_nhap_du_an_pattern = r"Từ\s*nguồn\s*thu\s*của\s*dự\s*án[^\d]*([\d\.,]+)\s*đồng\s*/\s*tháng"
    m = re.search(thu_nhap_du_an_pattern, full_text, flags=re.IGNORECASE)
    thu_nhap_du_an = 0.0
    if m:
        thu_nhap_du_an = vnd_to_float(m.group(1))

    thu_nhap_luong_pattern = r"Thu\s*nhập\s*từ\s*lương\s*[:：]\s*([\d\.,]+)\s*đồng\s*/\s*tháng"
    m = re.search(thu_nhap_luong_pattern, full_text, flags=re.IGNORECASE)
    thu_nhap_luong = 0.0
    if m:
        thu_nhap_luong = vnd_to_float(m.group(1))

    tong_thu_nhap_pattern = r"Tổng\s*thu\s*nhập\s*(?:ổn\s*định)?\s*(?:hàng\s*)?tháng\s*[:：]\s*([\d\.,]+)\s*đồng"
    m = re.search(tong_thu_nhap_pattern, full_text, flags=re.IGNORECASE)
    if m:
        data["thu_nhap_thang"] = vnd_to_float(m.group(1))
    else:
        data["thu_nhap_thang"] = thu_nhap_luong + thu_nhap_du_an

    # === 4. TÀI SẢN BẢO ĐẢM ===
    tsdb_pattern1 = r"Tài\s*sản\s*1[^\n]*Giá\s*trị\s*[:：]\s*([\d\.,]+)\s*đồng"
    m = re.search(tsdb_pattern1, full_text, flags=re.IGNORECASE | re.DOTALL)
    if m:
        data["gia_tri_tsdb"] = vnd_to_float(m.group(1))
    else:
        tsdb_pattern2 = r"Giá\s*trị\s*nhà\s*dự\s*kiến\s*mua\s*[:：]\s*([\d\.,]+)\s*đồng"
        m = re.search(tsdb_pattern2, full_text, flags=re.IGNORECASE)
        if m:
            data["gia_tri_tsdb"] = vnd_to_float(m.group(1))

    # === 5. THÔNG TIN BỔ SUNG ===
    loi_nhuan_pattern = r"Lợi\s*nhuận\s*(?:ròng)?\s*(?:năm)?[^\d]*([\d\.,]+)\s*đồng"
    m = re.search(loi_nhuan_pattern, full_text, flags=re.IGNORECASE)
    if m:
        data["loi_nhuan_rong_nam"] = vnd_to_float(m.group(1))
    elif thu_nhap_du_an > 0:
        data["loi_nhuan_rong_nam"] = thu_nhap_du_an * 12

    if data["tong_nhu_cau_von"] == 0 and (data["von_doi_ung"] + data["so_tien_vay"] > 0):
        data["tong_nhu_cau_von"] = data["von_doi_ung"] + data["so_tien_vay"]

    if data["tong_von_dau_tu"] == 0:
        data["tong_von_dau_tu"] = data["tong_nhu_cau_von"]

    if data["gia_tri_tsdb"] == 0 and data["tong_nhu_cau_von"] > 0:
        data["gia_tri_tsdb"] = data["tong_nhu_cau_von"]

    return data


def annuity_payment(principal: float, annual_rate_pct: float, months: int) -> float:
    r = annual_rate_pct / 100.0 / 12.0
    if months <= 0:
        return 0.0
    if r == 0:
        return principal / months
    pmt = principal * r * (1 + r) ** months / ((1 + r) ** months - 1)
    return pmt


def build_amortization(principal: float, annual_rate_pct: float, months: int, start_date: Optional[dt.date] = None) -> pd.DataFrame:
    if start_date is None:
        start_date = dt.date.today()
    r = annual_rate_pct / 100.0 / 12.0
    pmt = annuity_payment(principal, annual_rate_pct, months)

    schedule = []
    balance = principal
    for i in range(1, months + 1):
        interest = balance * r
        principal_pay = pmt - interest
        balance = max(0.0, balance - principal_pay)
        pay_date = start_date + dt.timedelta(days=30 * i)
        schedule.append({
            "Kỳ": i,
            "Ngày thanh toán": pay_date.strftime("%d/%m/%Y"),
            "Tiền lãi": round(interest, 0),
            "Tiền gốc": round(principal_pay, 0),
            "Tổng phải trả": round(pmt, 0),
            "Dư nợ còn lại": round(balance, 0),
        })
    df = pd.DataFrame(schedule)
    return df

def style_schedule_table(df: pd.DataFrame) -> pd.DataFrame:
    """Tô màu bảng kế hoạch trả nợ"""
    def color_row(row):
        if row['Kỳ'] % 2 == 0:
            return ['background-color: #f0f8ff'] * len(row)
        else:
            return ['background-color: #ffffff'] * len(row)

    styled = df.style.apply(color_row, axis=1)
    styled = styled.format({
        'Tiền lãi': lambda x: format_vnd(x),
        'Tiền gốc': lambda x: format_vnd(x),
        'Tổng phải trả': lambda x: format_vnd(x),
        'Dư nợ còn lại': lambda x: format_vnd(x)
    })
    styled = styled.set_properties(**{
        'text-align': 'right',
        'font-size': '14px'
    }, subset=['Tiền lãi', 'Tiền gốc', 'Tổng phải trả', 'Dư nợ còn lại'])
    styled = styled.set_properties(**{
        'text-align': 'center'
    }, subset=['Kỳ', 'Ngày thanh toán'])

    return styled


def compute_metrics(d: Dict[str, Any]) -> Dict[str, Any]:
    res = {}
    pmt = annuity_payment(d.get("so_tien_vay", 0.0), d.get("lai_suat_nam", 0.0), d.get("thoi_gian_vay_thang", 0))
    thu_nhap_thang = max(1e-9, d.get("thu_nhap_thang", 0.0))
    res["PMT_thang"] = pmt
    res["DSR"] = pmt / thu_nhap_thang if thu_nhap_thang > 0 else np.nan
    tong_nhu_cau = d.get("tong_nhu_cau_von", 0.0)
    von_doi_ung = d.get("von_doi_ung", 0.0)
    so_tien_vay = d.get("so_tien_vay", 0.0)
    gia_tri_tsdb = d.get("gia_tri_tsdb", 0.0)
    tong_no_hien_tai = d.get("tong_no_hien_tai", 0.0)
    loi_nhuan_rong_nam = d.get("loi_nhuan_rong_nam", 0.0)
    tong_von_dau_tu = d.get("tong_von_dau_tu", 0.0)

    res["E_over_C"] = (von_doi_ung / tong_nhu_cau) if tong_nhu_cau > 0 else np.nan
    res["LTV"] = (so_tien_vay / gia_tri_tsdb) if gia_tri_tsdb > 0 else np.nan
    thu_nhap_nam = thu_nhap_thang * 12.0
    res["Debt_over_Income"] = (tong_no_hien_tai + so_tien_vay) / max(1e-9, thu_nhap_nam)
    res["ROI"] = (loi_nhuan_rong_nam / max(1e-9, tong_von_dau_tu)) if tong_von_dau_tu > 0 else np.nan
    res["CFR"] = ((thu_nhap_thang - pmt) / thu_nhap_thang) if thu_nhap_thang > 0 else np.nan
    res["Coverage"] = (gia_tri_tsdb / max(1e-9, so_tien_vay)) if so_tien_vay > 0 else np.nan
    res["Phuong_an_hop_ly"] = math.isclose(tong_nhu_cau, von_doi_ung + so_tien_vay, rel_tol=0.02, abs_tol=1_000_000)

    score = 0.0
    if not np.isnan(res["DSR"]):
        score += max(0.0, 1.0 - min(1.0, res["DSR"])) * 0.25
    if not np.isnan(res["LTV"]):
        score += max(0.0, 1.0 - min(1.0, res["LTV"])) * 0.25
    if not np.isnan(res["E_over_C"]):
        score += min(1.0, res["E_over_C"] / 0.3) * 0.2
    if not np.isnan(res["CFR"]):
        score += max(0.0, min(1.0, (res["CFR"]))) * 0.2
    if not np.isnan(res["Coverage"]):
        score += min(1.0, (res["Coverage"] / 1.5)) * 0.1
    res["Score_AI_demo"] = round(score, 3)
    return res

def create_metrics_chart(metrics: Dict[str, Any]):
    """Tạo biểu đồ trực quan cho các chỉ tiêu tài chính chính"""
    if go is None or px is None:
        st.warning("Thư viện Plotly chưa được cài đặt. Không thể vẽ biểu đồ.")
        return

    df_metrics = pd.DataFrame({
        "Chỉ tiêu": ["DSR", "LTV", "E/C", "Coverage", "CFR"],
        "Giá trị": [
            metrics.get("DSR", np.nan),
            metrics.get("LTV", np.nan),
            metrics.get("E_over_C", np.nan),
            metrics.get("Coverage", np.nan),
            metrics.get("CFR", np.nan),
        ],
        "Ngưỡng tham chiếu": [0.8, 0.8, 0.2, 1.2, 0.0]
    })
    df_metrics = df_metrics.dropna(subset=['Giá trị']).reset_index(drop=True)

    if df_metrics.empty:
        st.info("Không có đủ dữ liệu để vẽ biểu đồ chỉ tiêu tài chính.")
        return

    def get_color(row):
        metric = row['Chỉ tiêu']
        value = row['Giá trị']
        ref = row['Ngưỡng tham chiếu']
        if metric in ["DSR", "LTV"]:
            return "green" if value <= ref else "red"
        elif metric in ["E/C", "Coverage", "CFR"]:
            return "green" if value >= ref else "red"
        return "gray"

    df_metrics['Màu'] = df_metrics.apply(get_color, axis=1)
    df_metrics['Giá trị (%)'] = df_metrics['Giá trị'] * 100

    fig = px.bar(
        df_metrics,
        x="Chỉ tiêu",
        y="Giá trị (%)",
        color="Màu",
        color_discrete_map={"green": "#28a745", "red": "#dc3545", "gray": "#6c757d"},
        text=df_metrics['Giá trị (%)'].apply(lambda x: f"{x:,.1f}%"),
        title="Biểu đồ Chỉ tiêu Tài chính (CADAP)",
        labels={"Giá trị (%)": "Giá trị (%)", "Chỉ tiêu": "Chỉ tiêu"},
    )

    for index, row in df_metrics.iterrows():
        metric = row['Chỉ tiêu']
        ref_value = row['Ngưỡng tham chiếu'] * 100
        color = "#ffc107" if ref_value > 0 else "#007bff"

        if metric in ["DSR", "LTV"]:
            fig.add_shape(
                type="line",
                x0=index - 0.4, x1=index + 0.4, y0=ref_value, y1=ref_value,
                line=dict(color=color, width=2, dash="dash"),
                xref="x", yref="y",
                name=f"Ngưỡng {metric}"
            )
            fig.add_annotation(
                x=index, y=ref_value + 3,
                text=f"Max {ref_value:g}%", showarrow=False,
                font=dict(color=color, size=10),
            )
        elif metric in ["E/C", "Coverage"]:
            fig.add_shape(
                type="line",
                x0=index - 0.4, x1=index + 0.4, y0=ref_value, y1=ref_value,
                line=dict(color=color, width=2, dash="dash"),
                xref="x", yref="y",
                name=f"Ngưỡng {metric}"
            )
            fig.add_annotation(
                x=index, y=ref_value - 3,
                text=f"Min {ref_value:g}%", showarrow=False,
                font=dict(color=color, size=10),
            )

    fig.update_layout(
        showlegend=False,
        yaxis_title="Giá trị (%)",
        xaxis_title="Chỉ tiêu",
        hovermode="x unified"
    )

    st.plotly_chart(fig, use_container_width=True)


def gemini_analyze(d: Dict[str, Any], metrics: Dict[str, Any], model_name: str, api_key: str) -> str:
    if genai is None:
        return "Thư viện google-generativeai chưa được cài. Vui lòng thêm 'google-generativeai' vào requirements.txt."
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(model_name)

        d_formatted = {k: format_vnd(v) if isinstance(v, (int, float)) and k != 'lai_suat_nam' else v for k, v in d.items()}
        metrics_formatted = {
            k: (f"{v*100:,.1f}%"
                if k not in ["PMT_thang", "Debt_over_Income", "Score_AI_demo"] and not np.isnan(v)
                else format_vnd(v) if k == "PMT_thang"
                else f"{v:,.2f}")
            for k, v in metrics.items()
        }

        prompt = f"""
Bạn là chuyên viên tín dụng. Phân tích hồ sơ vay sau (JSON) và đưa ra đề xuất "Cho vay" / "Cho vay có điều kiện" / "Không cho vay" kèm giải thích ngắn gọn (<=200 từ).
JSON đầu vào:
Khách hàng & phương án: {json.dumps(d_formatted, ensure_ascii=False)}
Chỉ tiêu tính toán: {json.dumps(metrics_formatted, ensure_ascii=False)}
Ngưỡng tham chiếu:
- DSR ≤ 0.8; LTV ≤ 0.8; E/C ≥ 0.2; CFR > 0; Coverage > 1.2.
- Nếu thông tin thiếu, hãy nêu giả định rõ ràng.
"""
        resp = model.generate_content(prompt)
        return resp.text or "(Không có nội dung từ Gemini)"
    except Exception as e:
        return f"
