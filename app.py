import streamlit as st
import pandas as pd
import os
import io
import requests
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Cấu hình tiêu đề trang web hiển thị trên trình duyệt
st.set_page_config(page_title="Hệ Thống In Tem Tự Động", page_icon="🏷️", layout="centered")

@st.cache_resource
def tai_va_dang_ky_font():
    """Tải font Arial từ Internet về máy chủ web để không bị lỗi font Tiếng Việt"""
    font_dir = "fonts"
    if not os.path.exists(font_dir):
        os.makedirs(font_dir)
        
    font_reg_path = os.path.join(font_dir, "arial.ttf")
    font_bold_path = os.path.join(font_dir, "arialbd.ttf")
    
    # Tải file font nếu chưa tồn tại trên hosting
    if not os.path.exists(font_reg_path):
        r = requests.get("https://github.com")
        with open(font_reg_path, "wb") as f:
            f.write(r.content)
            
    if not os.path.exists(font_bold_path):
        r = requests.get("https://github.com")
        with open(font_bold_path, "wb") as f:
            f.write(r.content)
            
    try:
        pdfmetrics.registerFont(TTFont('Arial-VN', font_reg_path))
        pdfmetrics.registerFont(TTFont('Arial-VN-Bold', font_bold_path))
        return 'Arial-VN-Bold', 'Arial-VN'
    except Exception:
        return 'Helvetica-Bold', 'Helvetica'

def ngat_dong_tu_dong_theo_chieu_rong(txt, c, font_name, font_size, max_width_mm):
    txt = str(txt).strip().upper()
    max_width_points = max_width_mm * mm
    if c.stringWidth(txt, font_name, font_size) <= max_width_points:
        return txt, ""
    words = txt.split()
    line1 = ""
    for i, word in enumerate(words):
        test_line = f"{line1} {word}".strip()
        if c.stringWidth(test_line, font_name, font_size) <= max_width_points:
            line1 = test_line
        else:
            return line1, " ".join(words[i:])
    mid = len(txt) // 2
    return txt[:mid], txt[mid:]

def ngat_dong_chu_dong_theo_dau_gach(txt):
    txt = str(txt).strip().upper()
    if '-' in txt:
        parts = txt.split('-', 1)
        line1 = f"{parts[0].strip()} -"
        line2 = parts[1].strip()
        return line1, line2
    return txt, ""

def xu_ly_in_tem_web(file_excel, font_bold, font_reg):
    # Đọc file Excel từ bộ nhớ đệm (buffer) của trình duyệt
    df = pd.read_excel(file_excel, header=0).fillna('')
    
    # Tạo luồng lưu trữ file PDF trong bộ nhớ (Memory Buffer) thay vì lưu vào ổ cứng
    pdf_buffer = io.BytesIO()
    
    width_mm = 100 * mm
    height_mm = 50 * mm
    c = canvas.Canvas(pdf_buffer, pagesize=(width_mm, height_mm))
    
    so_tem = 0
    for index, row in df.iterrows():
        cot_A = str(row.get('DỰ ÁN', '')).strip()
        cot_B = str(row.get('KÝ HIỆU CỬA', '')).strip()
        cot_C = str(row.get('ĐỢT', '')).strip()
        cot_D = str(row.get('TỔ', '')).strip()
        cot_E = str(row.get('W(mm)', '')).strip()
        cot_F = str(row.get('H(mm)', '')).strip()
        cot_G = str(row.get('SỐ LƯỢNG CÁNH', '')).strip()
        cot_I = str(row.get('KB', '')).strip()
        
        if not cot_A and not cot_B and not cot_I:
            continue

        c.setFont(font_bold, 17)
        c.drawCentredString(57 * mm, 40 * mm, cot_A.upper())
        
        c.setFont(font_bold, 18)
        line1_B, line2_B = ngat_dong_tu_dong_theo_chieu_rong(cot_B, c, font_bold, 18, max_width_mm=75)
        if line2_B:
            c.drawCentredString(57 * mm, 33 * mm, line1_B)
            c.drawCentredString(57 * mm, 27 * mm, line2_B)
        else:
            c.drawCentredString(57 * mm, 30 * mm, line1_B)
        
        le_trai_moi = 16
        c.setFont(font_reg, 8.5)
        line1_C, line2_C = ngat_dong_chu_dong_theo_dau_gach(cot_C)
        if line2_C:
            c.drawString(le_trai_moi * mm, 25 * mm, line1_C)
            c.drawString(le_trai_moi * mm, 21 * mm, line2_C)
        else:
            c.drawString(le_trai_moi * mm, 23 * mm, line1_C)
            
        c.drawString(le_trai_moi * mm, 12 * mm, cot_D)
        
        c.setFont(font_reg, 11)
        c.drawRightString(94 * mm, 24 * mm, cot_G)
        
        c.setFont(font_reg, 12.5)
        vi_tri_khung_do = 45
        c.drawString(vi_tri_khung_do * mm, 18 * mm, cot_E)  
        
        do_rong_chu_do_mm = c.stringWidth(cot_E, font_reg, 12.5) / mm
        vi_tri_khung_lam = vi_tri_khung_do + do_rong_chu_do_mm + 5
        c.drawString(vi_tri_khung_lam * mm, 18 * mm, cot_F)
        
        do_dai_chuoi = len(cot_I)
        if do_dai_chuoi > 18:
            c.setFont(font_bold, 13)
        elif do_dai_chuoi > 14:
            c.setFont(font_bold, 15)
        else:
            c.setFont(font_bold, 18)
        c.drawCentredString(60 * mm, 4.5 * mm, cot_I)
        
        # --- VẼ KÝ TỰ VECTO "C-bmw" SIÊU NHỎ ---
        c.saveState()
        c.setStrokeColorRGB(0.72, 0.72, 0.72)
        c.setLineWidth(0.12)
        c.setDash(0.3, 0.3)
        x_base, y_base, h, w, sp = 2*mm, 2.2*mm, 1.1*mm, 0.6*mm, 0.3*mm
        c.line(x_base + w, y_base + h, x_base, y_base + h)
        c.line(x_base, y_base + h, x_base, y_base)
        c.line(x_base, y_base, x_base + w, y_base)
        x_base += w + sp
        c.line(x_base, y_base + h/2, x_base + w*0.8, y_base + h/2)
        x_base += w*0.8 + sp
        c.line(x_base, y_base + h, x_base, y_base)
        c.line(x_base, y_base + h/2, x_base + w, y_base + h/2)
        c.line(x_base + w, y_base + h/2, x_base + w, y_base)
        c.line(x_base + w, y_base, x_base, y_base)
        x_base += w + sp
        c.line(x_base, y_base + h/2, x_base, y_base)
        c.line(x_base, y_base + h/2, x_base + w*0.5, y_base + h/2)
        c.line(x_base + w*0.5, y_base + h/2, x_base + w*0.5, y_base)
        c.line(x_base + w*0.5, y_base + h/2, x_base + w, y_base + h/2)
        c.line(x_base + w, y_base + h/2, x_base + w, y_base)
        x_base += w + sp
        c.line(x_base, y_base + h/2, x_base + w*0.25, y_base)
        c.line(x_base + w*0.25, y_base, x_base + w*0.5, y_base + h/2)
        c.line(x_base + w*0.5, y_base + h/2, x_base + w*0.75, y_base)
        c.line(x_base + w*0.75, y_base, x_base + w, y_base + h/2)
        c.restoreState()
        
        c.showPage()
        so_tem += 1
        
    c.save()
    pdf_buffer.seek(0)
    return pdf_buffer, so_tem

# --- GIAO DIỆN WEB hiển thị trên Streamlit ---
st.title("🏷️ Hệ Thống Xuất Tem Tự Động")
st.write("Kéo thả file dữ liệu Excel vào khung bên dưới để tự động tạo file PDF in tem chuẩn kích thước 100x50 mm.")

font_bold, font_reg = tai_va_dang_ky_font()

# Khung upload file Excel
uploaded_file = st.file_uploader("Chọn file dữ liệu Excel (.xlsx, .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    with st.spinner("Đang xử lý dữ liệu và tạo file tem..."):
        try:
            pdf_data, tong_so_tem = xu_ly_in_tem_web(uploaded_file, font_bold, font_reg)
            
            st.success(f"🎉 Đã xử lý thành công! Tổng số: **{tong_so_tem}** tem.")
            
            # Nút tải file PDF xuống trình duyệt
            st.download_button(
                label="📥 TẢI FILE PDF IN TEM VỀ MÁY",
                data=pdf_data,
                file_name=f"ket_qua_in_tem_{uploaded_file.name.split('.')[0]}.pdf",
                mime="application/pdf"
            )
        except Exception as e:
            st.error(f"❌ Có lỗi xảy ra khi xử lý file. Chi tiết: {str(e)}")
