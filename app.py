import os
import urllib.request
import streamlit as st
import pandas as pd
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

def dang_ky_font_tieng_viet():
    """Nhúng font Arial của Windows để in tiếng Việt không bị lỗi font"""
    font_path_regular = "C:\\Windows\\Fonts\\arial.ttf"
    font_path_bold = "C:\\Windows\\Fonts\\arialbd.ttf"
    
    # Đường dẫn lưu font tạm thời nếu chạy trên môi trường Linux (Streamlit Cloud)
    cloud_font_reg = "/tmp/arial.ttf"
    cloud_font_bold = "/tmp/arialbd.ttf"
    
    if os.path.exists(font_path_regular) and os.path.exists(font_path_bold):
        try:
            pdfmetrics.registerFont(TTFont('Arial-VN', font_path_regular))
            pdfmetrics.registerFont(TTFont('Arial-VN-Bold', font_path_bold))
            return 'Arial-VN-Bold', 'Arial-VN'
        except Exception:
            pass

    # Tự động tải font từ Internet nếu không tìm thấy font cục bộ trong hệ thống
    try:
        if not os.path.exists(cloud_font_reg):
            urllib.request.urlretrieve("https://github.com", cloud_font_reg)
        if not os.path.exists(cloud_font_bold):
            urllib.request.urlretrieve("https://github.com", cloud_font_bold)
            
        pdfmetrics.registerFont(TTFont('Arial-VN', cloud_font_reg))
        pdfmetrics.registerFont(TTFont('Arial-VN-Bold', cloud_font_bold))
        return 'Arial-VN-Bold', 'Arial-VN'
    except Exception:
        return 'Helvetica-Bold', 'Helvetica'

def ngat_dong_tu_dong_theo_chieu_rong(txt, c, font_name, font_size, max_width_mm):
    """
    Tự động ngắt dòng thông minh dựa trên độ rộng thực tế của chữ.
    Chỉ ngắt dòng khi chữ đi hết khuôn (vượt quá chiều rộng tối đa).
    """
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
    """
    Ép ngắt dòng chủ động tại dấu gạch ngang (-) đầu tiên.
    """
    txt = str(txt).strip().upper()
    
    if '-' in txt:
        parts = txt.split('-', 1)
        line1 = f"{parts[0].strip()} -"
        line2 = parts[1].strip()
        return line1, line2
        
    return txt, ""

def xu_ly_in_tem_chuan_so_do(file_excel):
    """Giữ nguyên toàn bộ logic đọc data excel và vẽ canvas của bạn"""
    try:
        df = pd.read_excel(file_excel, header=0).fillna('')
        
        # Tạo file đầu ra tại thư mục tạm của hệ thống Cloud
        file_pdf_output = "/tmp/output_tem_in.pdf"
        
        width_mm = 100 * mm
        height_mm = 50 * mm
        
        c = canvas.Canvas(file_pdf_output, pagesize=(width_mm, height_mm))
        font_bold, font_reg = dang_ky_font_tieng_viet()
        so_tem_da_in = 0
        
        for index, row in df.iterrows():
            cot_A = str(row.get('DỰ ÁN', '')).strip()
            cot_B = str(row.get('KÝ HIỆU CỬA', '')).strip()
            cot_C = str(row.get('ĐỢT', '')).strip()
            cot_D = str(row.get('TỔ', '')).strip()
            cot_E = str(row.get('W(mm)', '')).strip()
            cot_F = str(row.get('H(mm)', '')).strip()
            cot_G = str(row.get('SỐ LƯỢNG CÁNH', '')).strip()
            cot_I = str(row.get('KB', '')).strip()
            
            try:
                cot_J = str(row.iloc[9]).strip() if len(row) > 9 else ""
                if cot_J.endswith('.0'):
                    cot_J = cot_J[:-2]
            except Exception:
                cot_J = ""
            
            if not cot_A and not cot_B and not cot_I:
                continue

            # --- BẮT ĐẦU VẼ CHỮ LÊN PHÔI TEM ---
            # Dòng 1: TIÊU ĐỀ DỰ ÁN
            c.setFont(font_bold, 17)
            c.drawCentredString(57 * mm, 43 * mm, cot_A.upper())
            
            # Dòng 2: KÝ HIỆU CỬA
            c.saveState() 
            kich_thuoc_font_goc = 18
            vung_xanh_max_width_mm = 84
            max_width_points = vung_xanh_max_width_mm * mm
            chuoi_ky_hieu = cot_B.upper().strip()
            
            while kich_thuoc_font_goc > 9:
                do_rong_thuc_te = c.stringWidth(chuoi_ky_hieu, font_bold, kich_thuoc_font_goc)
                if do_rong_thuc_te <= max_width_points:
                    break
                kich_thuoc_font_goc -= 1.5 
            
            c.setFont(font_bold, kich_thuoc_font_goc)
            if kich_thuoc_font_goc >= 16:
                toa_do_y = 30.5 * mm
            elif kich_thuoc_font_goc >= 12:
                toa_do_y = 32.0 * mm
            else:
                toa_do_y = 33.5 * mm
                
            c.drawCentredString(57 * mm, toa_do_y, chuoi_ky_hieu)
            c.restoreState()

            # KHU VỰC TRÁI: ĐỢT
            le_trai_moi = 16
            c.setFont(font_reg, 8.5)
            line1_C, line2_C = ngat_dong_chu_dong_theo_dau_gach(cot_C)
            if line2_C:
                c.drawString(le_trai_moi * mm, 25 * mm, line1_C)
                c.drawString(le_trai_moi * mm, 21 * mm, line2_C)
            else:
                c.drawString(le_trai_moi * mm, 23 * mm, line1_C)
                
            # TỔ
            c.drawString(le_trai_moi * mm, 12 * mm, cot_D)
            
            # KHU VỰC PHẢI TRÊN: SỐ LƯỢNG CÁNH
            c.setFont(font_reg, 11)
            c.drawRightString(94 * mm, 24 * mm, cot_G)
            
            # KHU VỰC GIỮA: THÔNG SỐ KÍCH THƯỚC W & H
            c.setFont(font_reg, 12.5)
            vi_tri_khung_do = 45
            c.drawString(vi_tri_khung_do * mm, 18 * mm, cot_E)  
            
            do_rong_chu_do_mm = c.stringWidth(cot_E, font_reg, 12.5) / mm
            vi_tri_khung_lam = vi_tri_khung_do + do_rong_chu_do_mm + 5
            c.drawString(vi_tri_khung_lam * mm, 18 * mm, cot_F)
            
            # DÒNG CUỐI CÙNG: MÃ CHÂN TEM
            do_dai_chuoi = len(cot_I)
            if do_dai_chuoi > 18:
                c.setFont(font_bold, 13)
            elif do_dai_chuoi > 14:
                c.setFont(font_bold, 15)
            else:
                c.setFont(font_bold, 18)
            c.drawCentredString(60 * mm, 4.5 * mm, cot_I)
            
            # VẼ Ô TRÒN CHỨA SỐ THỨ TỰ (CỘT J)
            if cot_J:
                c.saveState()
                toa_do_x_tron = 96 * mm 
                toa_do_y_tron = 4.2 * mm
                ban_kinh = 2.2 * mm
                
                c.setLineWidth(0.5)
                c.circle(toa_do_x_tron, toa_do_y_tron, ban_kinh, stroke=1, fill=0)
                c.setFont(font_bold, 6)
                c.drawCentredString(toa_do_x_tron, toa_do_y_tron - 0.7 * mm, str(cot_J))
                c.restoreState()

            # VẼ VECTO "C-bmw"
            c.saveState()
            c.setStrokeColorRGB(0.72, 0.72, 0.72)      
            c.setLineWidth(0.12)                      
            c.setDash(0.3, 0.3)                       
            
            x_base = 2 * mm
            y_base = 2.2 * mm
            h = 1.1 * mm                              
            w = 0.6 * mm                              
            sp = 0.3 * mm                             
            
            # Vẽ chữ 'C'
            c.line(x_base + w, y_base + h, x_base, y_base + h)
            c.line(x_base, y_base + h, x_base, y_base)
            c.line(x_base, y_base, x_base + w, y_base)
            
            # Vẽ dấu gạch ngang '-'
            x_base += w + sp
            c.line(x_base, y_base + h/2, x_base + w*0.8, y_base + h/2)
            
            # Vẽ chữ 'b'
            x_base += w*0.8 + sp
            c.line(x_base, y_base + h, x_base, y_base)
            c.line(x_base, y_base + h/2, x_base + w, y_base + h/2)
            c.line(x_base + w, y_base + h/2, x_base + w, y_base)
            c.line(x_base + w, y_base, x_base, y_base)
            
            # Vẽ chữ 'm'
            x_base += w + sp
            c.line(x_base, y_base + h/2, x_base, y_base)
            c.line(x_base, y_base + h/2, x_base + w*0.5, y_base + h/2)
            c.line(x_base + w*0.5, y_base + h/2, x_base + w*0.5, y_base)
            c.line(x_base + w*0.5, y_base + h/2, x_base + w, y_base + h/2)
            c.line(x_base + w, y_base + h/2, x_base + w, y_base)
            
            # Vẽ chữ 'w'
            x_base += w + sp
            c.line(x_base, y_base + h/2, x_base + w*0.25, y_base)
            c.line(x_base + w*0.25, y_base, x_base + w*0.5, y_base + h/2)
            c.line(x_base + w*0.5, y_base + h/2, x_base + w*0.75, y_base)
            c.line(x_base + w*0.75, y_base, x_base + w, y_base + h/2)
            
            c.restoreState()
            
            c.showPage()
            so_tem_da_in += 1
            
        c.save()
        return file_pdf_output, so_tem_da_in
    except Exception as e:
        st.error(f"Lỗi hệ thống: Không thể xử lý dữ liệu. Chi tiết: {str(e)}")
        return None, 0
# --- GIAO DIỆN HIỂN THỊ WEB STREAMLIT ---
st.set_page_config(page_title="Hệ thống In Tem Tự Động", page_icon="🏷️")
st.title("🏷️ Ứng dụng xuất dữ liệu tem chuẩn")
st.write("Vui lòng tải file Excel lên để chuyển đổi thành file PDF tem in.")

# Widget tải file của Streamlit thay cho filedialog của tkinter
uploaded_file = st.file_uploader("Chọn file Excel để xuất dữ liệu lên tem", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.success(f"Đã tải lên file: {uploaded_file.name}")
    
    # Nút bấm kích hoạt xử lý dữ liệu
    if st.button("Bắt đầu xử lý và tạo Tem PDF", type="primary"):
        with st.spinner("Hệ thống đang xử lý dữ liệu, vui lòng đợi..."):
            path_pdf_result, tong_so_tem = xu_ly_in_tem_chuan_so_do(uploaded_file)
            
            if path_pdf_result and tong_so_tem > 0:
                st.success(f"Đã tạo thành công file in tem! Tổng số: {tong_so_tem} tem.")
                
                # Đọc file PDF vừa tạo ra từ bộ nhớ tạm để chuẩn bị tải xuống
                with open(path_pdf_result, "rb") as f:
                    pdf_bytes = f.read()
                
                # Đổi đuôi tên file gốc từ .xlsx sang .pdf để người dùng tải về đồng bộ
                ten_file_goc = os.path.splitext(uploaded_file.name)[0]
                
                # Widget cho người dùng bấm tải file về máy thay cho messagebox thành công
                st.download_button(
                    label="📥 Nhấn vào đây để tải file PDF Tem về máy",
                    data=pdf_bytes,
                    file_name=f"{ten_file_goc}.pdf",
                    mime="application/pdf"
                )

