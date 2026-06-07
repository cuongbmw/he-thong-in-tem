import streamlit as st
import pandas as pd
import os
import io
import requests
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pdf2image import convert_from_bytes

st.set_page_config(page_title="Hệ Thống Quản Lý In Ấn", page_icon="🛠️", layout="centered")

@st.cache_resource
def tai_va_dang_ky_font():
    font_dir = "fonts"
    if not os.path.exists(font_dir): os.makedirs(font_dir)
    font_reg_path = os.path.join(font_dir, "arial.ttf")
    font_bold_path = os.path.join(font_dir, "arialbd.ttf")
    if not os.path.exists(font_reg_path):
        r = requests.get("https://github.com")
        with open(font_reg_path, "wb") as f: f.write(r.content)
    if not os.path.exists(font_bold_path):
        r = requests.get("https://github.com")
        with open(font_bold_path, "wb") as f: f.write(r.content)
    try:
        pdfmetrics.registerFont(TTFont('Arial-VN', font_reg_path))
        pdfmetrics.registerFont(TTFont('Arial-VN-Bold', font_bold_path))
        return 'Arial-VN-Bold', 'Arial-VN'
    except Exception: return 'Helvetica-Bold', 'Helvetica'

def ngat_dong_tu_dong_theo_chieu_rong(txt, c, font_name, font_size, max_width_mm):
    txt = str(txt).strip().upper()
    max_width_points = max_width_mm * mm
    if c.stringWidth(txt, font_name, font_size) <= max_width_points: return txt, ""
    words = txt.split()
    line1 = ""
    for i, word in enumerate(words):
        test_line = f"{line1} {word}".strip()
        if c.stringWidth(test_line, font_name, font_size) <= max_width_points: line1 = test_line
        else: return line1, " ".join(words[i:])
    mid = len(txt) // 2
    return txt[:mid], txt[mid:]

def ngat_dong_chu_dong_theo_dau_gach(txt):
    txt = str(txt).strip().upper()
    if '-' in txt:
        parts = txt.split('-', 1)
        return f"{parts[0].strip()} -", parts[1].strip()
    return txt, ""

def xu_ly_in_tem_web(file_excel, font_bold, font_reg):
    df = pd.read_excel(file_excel, header=0).fillna('')
    pdf_buffer = io.BytesIO()
    width_mm, height_mm = 100 * mm, 50 * mm
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
        if not cot_A and not cot_B and not cot_I: continue
        c.setFont(font_bold, 17)
        c.drawCentredString(57 * mm, 40 * mm, cot_A.upper())
        c.setFont(font_bold, 18)
        line1_B, line2_B = ngat_dong_tu_dong_theo_chieu_rong(cot_B, c, font_bold, 18, max_width_mm=75)
        if line2_B:
            c.drawCentredString(57 * mm, 33 * mm, line1_B)
            c.drawCentredString(57 * mm, 27 * mm, line2_B)
        else: c.drawCentredString(57 * mm, 30 * mm, line1_B)
        le_trai_moi = 16
        c.setFont(font_reg, 8.5)
        line1_C, line2_C = ngat_dong_chu_dong_theo_dau_gach(cot_C)
        if line2_C:
            c.drawString(le_trai_moi * mm, 25 * mm, line1_C)
            c.drawString(le_trai_moi * mm, 21 * mm, line2_C)
        else: c.drawString(le_trai_moi * mm, 23 * mm, line1_C)
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
        if do_dai_chuoi > 18: c.setFont(font_bold, 13)
        elif do_dai_chuoi > 14: c.setFont(font_bold, 15)
        else: c.setFont(font_bold, 18)
        c.drawCentredString(60 * mm, 4.5 * mm, cot_I)
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

# --- GIAO DIỆN CHÍNH ---
st.title("🛠️ Trung Tâm Xử Lý Bản In")

tab1, tab2 = st.tabs(["🏷️ Xuất Tem Từ Excel", "🖼️ Chuyển PDF Sang Ảnh"])

with tab1:
    st.header("Tạo Tem Tự Động")
    font_bold, font_reg = tai_va_dang_ky_font()
    uploaded_excel = st.file_uploader("Kéo thả file Excel", type=["xlsx", "xls"], key="excel_up")
    if uploaded_excel is not None:
        with st.spinner("Đang xử lý..."):
            try:
                pdf_data, tong_so_tem = xu_ly_in_tem_web(uploaded_excel, font_bold, font_reg)
                st.success(f"🎉 Đã xử lý thành công {tong_so_tem} tem.")
                st.download_button(label="📥 TẢI FILE PDF IN TEM", data=pdf_data, file_name="tem_in.pdf", mime="application/pdf")
            except Exception as e: st.error(f"Lỗi: {str(e)}")

with tab2:
    st.header("Chuyển Đổi PDF Sang Ảnh")
    uploaded_pdf = st.file_uploader("Kéo thả file PDF cần chuyển đổi", type=["pdf"], key="pdf_up")
    
    if uploaded_pdf is not None:
        pdf_bytes = uploaded_pdf.read()
        
        try:
            test_imgs = convert_from_bytes(pdf_bytes, dpi=72)
            tong_so_trang = len(test_imgs)
            st.info(f"📄 Tìm thấy tổng cộng: **{tong_so_trang}** trang trong file PDF.")
        except Exception as e:
            st.error(f"Không thể đọc file PDF. Lỗi: {str(e)}")
            tong_so_trang = 0
            
        if tong_so_trang > 0:
            col1, col2 = st.columns(2)
            with col1:
                dinh_dang = st.selectbox("Định dạng ảnh đầu ra", ["PNG", "JPEG"])
            with col2:
                cac_muc_dpi = [100, 200, 300, 400, 500, 600]
                muc_dpi = st.select_slider("Độ phân giải (DPI)", options=cac_muc_dpi, value=300)
            
            che_do_chon = st.radio("Chế độ xuất ảnh:", ["Xuất toàn bộ các trang", "Chỉ xuất một vài trang cụ thể"])
            danh_sach_trang_can_xuat = list(range(1, tong_so_trang + 1))
            
            if che_do_chon == "Chỉ xuất một vài trang cụ thể":
                nhap_trang = st.text_input(f"Nhập số trang muốn xuất (Ví dụ: 1 hoặc 1,3,5). Giới hạn từ 1 đến {tong_so_trang}:", value="1")
                try:
                    danh_sach_trang_can_xuat = [int(p.strip()) for p in nhap_trang.split(",") if p.strip().isdigit()]
                    danh_sach_trang_can_xuat = [p for p in danh_sach_trang_can_xuat if 1 <= p <= tong_so_trang]
                except:
                    st.warning("Định dạng số trang nhập vào chưa chuẩn, hệ thống mặc định chọn trang 1.")
                    danh_sach_trang_can_xuat = [1]
            
            if len(danh_sach_trang_can_xuat) == 0:
                st.warning("Vui lòng nhập ít nhất một số trang hợp lệ.")
            elif st.button("🚀 BẮT ĐẦU CHUYỂN ĐỔI"):
                with st.spinner("Đang trích xuất ảnh chất lượng cao..."):
                    try:
                        images = convert_from_bytes(
                            pdf_bytes, 
                            dpi=muc_dpi,
                            first_page=min(danh_sach_trang_can_xuat),
                            last_page=max(danh_sach_trang_can_xuat)
                        )
                        
                        st.success("📸 Đã xử lý xong danh sách trang yêu cầu!")
                        
                                                for idx_trang in danh_sach_trang_can_xuat:
                            idx_chuan = idx_trang - min(danh_sach_trang_can_xuat)
                            if idx_chuan < len(images):
                                img = images[idx_chuan]
                                img_buffer = io.BytesIO()
                                img.save(img_buffer, format=dinh_dang)
                                img_buffer.seek(0)
                                
                                # Gom gọn tham số thành biến ngắn để tránh gãy dòng chat
                                cap = f"Trang {idx_trang} ({muc_dpi} DPI)"
                                lbl = f"📥 Tải ảnh Trang {idx_trang} ({dinh_dang})"
                                ext = dinh_dang.lower()
                                f_name = f"trang_{idx_trang}_{muc_dpi}dpi.{ext}"
                                m_type = f"image/{ext}"
                                
                                st.image(img, caption=cap, use_container_width=True)
                                st.download_button(
                                    label=lbl,
                                    data=img_buffer,
                                    file_name=f_name,
                                    mime=m_type
                                )
                    except Exception as img_err:
                        st.error(f"Lỗi trích xuất hình ảnh: {str(img_err)}")

