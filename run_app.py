import os
import sys
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Cấu hình giao diện theo hệ thống máy tính (Sáng/Tối)
ctk.set_appearance_mode("System")  
ctk.set_default_color_theme("blue") 

class AppInTemOffline(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Hệ Thống Xuất Dữ Liệu Tem Chuẩn - Offline v1.0")
        self.geometry("550x350")
        self.resizable(False, False)
        
        self.selected_file_path = None
        self.init_ui()

    def init_ui(self):
        # Tiêu đề chính giao diện
        self.title_label = ctk.CTkLabel(self, text="🏷️ ỨNG DỤNG IN TEM CHUẨN SỐ ĐỒ", font=ctk.CTkFont(size=20, weight="bold"))
        self.title_label.pack(pady=25)

        # Khung chọn file Excel
        self.file_frame = ctk.CTkFrame(self)
        self.file_frame.pack(pady=10, padx=20, fill="x")

        self.btn_browse = ctk.CTkButton(self.file_frame, text="Chọn file Excel (.xlsx, .xls)", command=self.browse_file, width=180, fg_color="#1f538d")
        self.btn_browse.pack(side="left", padx=10, pady=10)

        self.lbl_file_name = ctk.CTkLabel(self.file_frame, text="Chưa chọn file dữ liệu nào...", text_color="gray", anchor="w")
        self.lbl_file_name.pack(side="left", padx=10, fill="x", expand=True)

        # Cấu hình chuẩn hóa chữ nghiêng slant="italic" của CustomTkinter
        self.status_label = ctk.CTkLabel(
            self, 
            text="Hệ thống chạy Offline 100% dựa trên phần cứng thiết bị", 
            font=ctk.CTkFont(size=12, slant="italic"), 
            text_color="green"
        )
        self.status_label.pack(pady=15)

        # Nút bấm tiến hành xử lý xuất bản in
        self.btn_process = ctk.CTkButton(self, text="🚀 BẮT ĐẦU XỬ LÝ VÀ XUẤT PDF", font=ctk.CTkFont(size=15, weight="bold"), height=45, fg_color="#2fa572", hover_color="#107c41", command=self.process_data)
        self.btn_process.pack(pady=20, fill="x", padx=40)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Chọn file Excel để xuất dữ liệu lên tem",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if file_path:
            self.selected_file_path = file_path
            self.lbl_file_name.configure(text=os.path.basename(file_path), text_color="white" if ctk.get_appearance_mode() == "Dark" else "black")

    def dang_ky_font_tieng_viet(self):
        """Đọc trực tiếp từ kho font hệ thống Windows không cần Internet"""
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

    def ngat_dong_chu_dong_theo_dau_gach(self, txt):
        txt = str(txt).strip().upper()
        if '-' in txt:
            parts = txt.split('-', 1)
            return f"{parts[0].strip()} -", parts[1].strip()
        return txt, ""
    def process_data(self):
        if not self.selected_file_path:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file dữ liệu Excel trước khi nhấn xuất tem!")
            return

        try:
            # Nhận diện Engine đọc file không phụ thuộc internet
            ten_file = self.selected_file_path.lower()
            if ten_file.endswith('.xls'):
                df = pd.read_excel(self.selected_file_path, header=0, engine='xlrd').fillna('')
            else:
                df = pd.read_excel(self.selected_file_path, header=0, engine='openpyxl').fillna('')

            # Định dạng vị trí lưu file PDF kế bên file Excel cũ
            thu_muc = os.path.dirname(self.selected_file_path)
            ten_file_base = os.path.splitext(os.path.basename(self.selected_file_path))[0]
            file_pdf_output = os.path.join(thu_muc, f"{ten_file_base}.pdf")
            
            dem = 1
            while os.path.exists(file_pdf_output):
                file_pdf_output = os.path.join(thu_muc, f"{ten_file_base} ({dem}).pdf")
                dem += 1

            # Tiến hành dựng Canvas bản in
            c = canvas.Canvas(file_pdf_output, pagesize=(100 * mm, 50 * mm))
            font_bold, font_reg = self.dang_ky_font_tieng_viet()
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

                # --- VẼ CHỮ LÊN PHÔI TEM ---
                c.setFont(font_bold, 17)
                c.drawCentredString(57 * mm, 43 * mm, cot_A.upper())
                
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

                le_trai_moi = 16
                c.setFont(font_reg, 8.5)
                line1_C, line2_C = self.ngat_dong_chu_dong_theo_dau_gach(cot_C)
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
                
                # -----------------------------------------------------------------
                # ĐOẠN CẤY MỚI 1: VẼ Ô TRÒN CHỨA SỐ THỨ TỰ (CỘT J)
                # -----------------------------------------------------------------
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

                # -----------------------------------------------------------------
                # ĐOẠN CẤY MỚI 2: VẼ KÝ TỰ VECTO "C-bmw"
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
                
                c.showPage()
                so_tem_da_in += 1
                
            c.save()
            thong_bao = f"Đã tạo thành công file in tem!\nTổng số: {so_tem_da_in} tem.\nLưu tại: {file_pdf_output}"
            messagebox.showinfo("Thành công", thong_bao)
            
        except Exception as e:
            messagebox.showerror("Lỗi hệ thống", f"Không thể xử lý dữ liệu in ấn.\nChi tiết lỗi: {str(e)}")

if __name__ == "__main__":
    app = AppInTemOffline()
    app.mainloop()
