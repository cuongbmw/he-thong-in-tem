import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

def dang_ky_font_tieng_viet():
    """Nhúng font Arial của Windows để in tiếng Việt không bị lỗi font"""
    font_path_regular = "C:\\Windows\\Fonts\\arial.ttf"
    font_path_bold = "C:\\Windows\\Fonts\\arialbd.ttf"
    
    if os.path.exists(font_path_regular) and os.path.exists(font_path_bold):
        try:
            pdfmetrics.registerFont(TTFont('Arial-VN', font_path_regular))
            pdfmetrics.registerFont(TTFont('Arial-VN-Bold', font_path_bold))
            return 'Arial-VN-Bold', 'Arial-VN'
        except Exception:
            return 'Helvetica-Bold', 'Helvetica'
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
    try:
        df = pd.read_excel(file_excel, header=0).fillna('')
        
        thu_muc = os.path.dirname(file_excel)
        ten_file_base = os.path.splitext(os.path.basename(file_excel))[0]
            
        file_pdf_output = os.path.join(thu_muc, f"{ten_file_base}.pdf")
        dem = 1
        while os.path.exists(file_pdf_output):
            file_pdf_output = os.path.join(thu_muc, f"{ten_file_base} ({dem}).pdf")
            dem += 1
        
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
            
            # --- LẤY DỮ LIỆU CỘT J (Cột số thứ tự) ---
            # Sử dụng iloc[9] (Index 9 tương đương với Cột J: A=0, B=1... J=9)
            try:
                cot_J = str(row.iloc[9]).strip() if len(row) > 9 else ""
                # Xóa đuôi '.0' nếu pandas tự động ép số nguyên thành số thực
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
            
            # =================================================================
            # SỬA ĐỔI: DÒNG 2 - KÝ HIỆU CỬA (TỰ ĐỘNG THU NHỎ & ĐẨY LÊN TRÊN)
            # =================================================================
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
            # =================================================================

            # KHU VỰC TRÁI: ĐỢT
            le_trai_moi = 16
            c.setFont(font_reg, 8.5)
            line1_C, line2_C = ngat_dong_chu_dong_theo_dau_gach(cot_C)
            if line2_C:
                c.drawString(le_trai_moi * mm, 25 * mm, line1_C)
                c.drawString(le_trai_moi * mm, 21 * mm, line2_C)
            else:
                c.drawString(le_trai_moi * mm, 23 * mm, line1_C)
                
            # TỔ nằm dưới Đợt ổn định
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
            
            # -----------------------------------------------------------------
            # PHẦN THÊM MỚI: VẼ Ô TRÒN CHỨA SỐ THỨ TỰ (CỘT J) - ĐÃ CÂN CHỈNH LÙI VÀO TRONG
            # -----------------------------------------------------------------
            if cot_J:
                c.saveState()
                
                # Lùi X sang trái từ 102mm về 96mm, tăng Y lên từ 3.5mm thành 5mm để không bị lố viền
                toa_do_x_tron = 96 * mm 
                toa_do_y_tron = 4.2 * mm
                ban_kinh = 2.2 * mm   # Giữ nguyên kích thước nhỏ gọn
                
                # Vẽ viền vòng tròn
                c.setLineWidth(0.5)
                c.circle(toa_do_x_tron, toa_do_y_tron, ban_kinh, stroke=1, fill=0)
                
                # Giữ nguyên cỡ chữ lọt lòng
                c.setFont(font_bold, 6)
                
                # Căn chỉnh tâm chữ theo tọa độ Y mới (trừ đi 0.7mm)
                c.drawCentredString(toa_do_x_tron, toa_do_y_tron - 0.7 * mm, str(cot_J))
                
                c.restoreState()


            # -----------------------------------------------------------------

            # -----------------------------------------------------------------
            # CẬP NHẬT: THU NHỎ TỶ LỆ KÝ TỰ VECTO "C-bmw"
            # -----------------------------------------------------------------
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
            # -----------------------------------------------------------------
            
            c.showPage()
            so_tem_da_in += 1
            
        c.save()
        thong_bao = f"Đã tạo thành công file in tem!\nTổng số: {so_tem_da_in} tem.\nLưu tại: {file_pdf_output}"
        messagebox.showinfo("Thành công", thong_bao)
        
    except Exception as e:
        messagebox.showerror("Lỗi hệ thống", f"Không thể xử lý dữ liệu in ấn.\nChi tiết lỗi: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    
    duong_dan_file = filedialog.askopenfilename(
        title="Chọn file Excel để xuất dữ liệu lên tem",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    
    if duong_dan_file:
        xu_ly_in_tem_chuan_so_do(duong_dan_file)