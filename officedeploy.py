import os
import subprocess
import urllib.request
import threading
import tkinter as tk
from tkinter import ttk, messagebox

# ==============================================================================
# TỪ ĐIỂN DỮ LIỆU CỦA MICROSOFT (ĐÃ BỔ SUNG HOME & BUSINESS, HOME & STUDENT)
# ==============================================================================
tu_dien_phien_ban = {
    # Dòng ProPlus
    "2016_ProPlus": "ProPlusRetail",
    "2019_ProPlus": "ProPlus2019Retail",
    "2021_ProPlus": "ProPlus2021Retail",
    "2024_ProPlus": "ProPlus2024Retail",
    "365_ProPlus": "O365ProPlusRetail",
    
    # Dòng Standard
    "2016_Standard": "StandardRetail",
    "2019_Standard": "Standard2019Retail",
    "2021_Standard": "Standard2021Retail",
    "2024_Standard": "Standard2024Retail",
    
    # Dòng Home & Business
    "2016_HomeBusiness": "HomeBusinessRetail",
    "2019_HomeBusiness": "HomeBusiness2019Retail",
    "2021_HomeBusiness": "HomeBusiness2021Retail",
    "2024_HomeBusiness": "HomeBusiness2024Retail",
    "365_HomeBusiness": "O365BusinessRetail",
    
    # Dòng Home & Student (Lưu ý: 2024 Microsoft đổi mã thành Home2024)
    "2016_HomeStudent": "HomeStudentRetail",
    "2019_HomeStudent": "HomeStudent2019Retail",
    "2021_HomeStudent": "HomeStudent2021Retail",
    "2024_HomeStudent": "Home2024Retail",
    "365_HomeStudent": "O365HomePremRetail"
}

tu_dien_ung_dung = {
    "Access": "Access",
    "Excel": "Excel",
    "Word": "Word",
    "PowerPoint": "PowerPoint",
    "Outlook": "Outlook",
    "Publisher": "Publisher",
    "OneNote": "OneNote",
    "Skype": "Lync",
    "Teams": "Teams"
}

# ==============================================================================
# HÀM HỖ TRỢ GIAO DIỆN
# ==============================================================================
def tao_nut_bam_mau(khung_chua, chu_hien_thi, mau_nen, mau_chu="white", hanh_dong=None):
    nut_bam = tk.Button(
        khung_chua, text=chu_hien_thi, bg=mau_nen, fg=mau_chu,
        font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2",
        command=hanh_dong, width=15, pady=5
    )
    return nut_bam

# ==============================================================================
# LỚP GIAO DIỆN VÀ XỬ LÝ CHÍNH
# ==============================================================================
class UngDungCaiDatOffice:
    def __init__(self, cua_so_chinh):
        self.cua_so = cua_so_chinh
        self.cua_so.title("Cài đặt Microsoft Office - Trực tiếp từ Server CDN")
        self.cua_so.geometry("600x650")
        self.cua_so.configure(bg="white")
        self.cua_so.resizable(False, False)

        self.phong_chu_thuong = ("Segoe UI", 9)
        self.phong_chu_dam = ("Segoe UI", 9, "bold")

        self.xay_dung_giao_dien()

    def xay_dung_giao_dien(self):
        # --- PHẦN TIÊU ĐỀ KHUNG TRÊN ---
        khung_tieu_de = tk.Frame(self.cua_so, bg="#E64A19", height=60)
        khung_tieu_de.pack(fill="x", side="top")
        
        tk.Label(khung_tieu_de, text="🏢 CÀI ĐẶT MICROSOFT OFFICE RETAIL", font=("Segoe UI", 14, "bold"), bg="#E64A19", fg="white").place(x=15, y=5)
        tk.Label(khung_tieu_de, text="Tải trực tiếp Max Speed từ Server Microsoft (CDN)", font=("Segoe UI", 9), bg="#E64A19", fg="white").place(x=40, y=32)
        tk.Label(khung_tieu_de, text="🌐 ONLINE", font=("Segoe UI", 10, "bold"), bg="#E64A19", fg="white").place(x=500, y=15)

        # --- KHUNG TAB CHUYỂN ĐỔI ---
        phong_cach = ttk.Style()
        phong_cach.configure("TNotebook", background="white")
        phong_cach.configure("TFrame", background="white")
        
        hop_tab = ttk.Notebook(self.cua_so)
        hop_tab.pack(fill="both", expand=True, padx=10, pady=10)

        tab_cai_dat = ttk.Frame(hop_tab)
        hop_tab.add(tab_cai_dat, text="  ⚙️ Cài đặt Office  ")

        # 1. KHUNG CHỌN PHIÊN BẢN
        khung_phien_ban = ttk.LabelFrame(tab_cai_dat, text=" 📄 Phiên bản Office ")
        khung_phien_ban.pack(fill="x", padx=10, pady=5, ipady=5)
        
        self.bien_phien_ban = tk.StringVar(value="2024")
        danh_sach_ban = ["Office 2016", "Office 2019", "Office 2021", "Office 2024", "Office 365"]
        for vi_tri, ten_ban in enumerate(danh_sach_ban):
            tk.Radiobutton(khung_phien_ban, text=ten_ban, variable=self.bien_phien_ban, value=ten_ban.split()[-1], bg="white").grid(row=0, column=vi_tri, padx=5, pady=5)
            
        tk.Label(khung_phien_ban, text="Phiên bản con:", bg="white").grid(row=1, column=0, columnspan=2, sticky="e", padx=5)
        
        # BỔ SUNG CÁC BẢN CON MỚI VÀO COMBOBOX
        self.hop_chon_ban_con = ttk.Combobox(khung_phien_ban, values=["ProPlus", "Standard", "Home & Business", "Home & Student"], state="readonly", width=40)
        self.hop_chon_ban_con.current(0)
        self.hop_chon_ban_con.grid(row=1, column=2, columnspan=3, sticky="w", pady=5)

        # 2. KHUNG KIẾN TRÚC VÀ NGÔN NGỮ
        khung_ngang = tk.Frame(tab_cai_dat, bg="white")
        khung_ngang.pack(fill="x", padx=10, pady=5)
        
        khung_kien_truc = ttk.LabelFrame(khung_ngang, text=" ⚙️ Kiến trúc ")
        khung_kien_truc.pack(side="left", fill="both", expand=True, padx=(0, 5))
        self.bien_kien_truc = tk.StringVar(value="64")
        tk.Radiobutton(khung_kien_truc, text="64-bit", variable=self.bien_kien_truc, value="64", bg="white").pack(side="left", padx=15, pady=5)
        tk.Radiobutton(khung_kien_truc, text="32-bit", variable=self.bien_kien_truc, value="32", bg="white").pack(side="right", padx=15, pady=5)

        khung_ngon_ngu = ttk.LabelFrame(khung_ngang, text=" 🌐 Ngôn ngữ ")
        khung_ngon_ngu.pack(side="right", fill="both", expand=True, padx=(5, 0))
        self.hop_chon_ngon_ngu = ttk.Combobox(khung_ngon_ngu, values=["English (US) - en-us", "Vietnamese - vi-vn"], state="readonly")
        self.hop_chon_ngon_ngu.current(0)
        self.hop_chon_ngon_ngu.pack(fill="x", padx=10, pady=8)

        # 3. KHUNG CHỌN ỨNG DỤNG (TÍCH CHỌN)
        khung_ung_dung = ttk.LabelFrame(tab_cai_dat, text=" ☑️ Chọn ứng dụng cần cài ")
        khung_ung_dung.pack(fill="x", padx=10, pady=5)
        
        danh_sach_app = [
            ("Access", 0, 0), ("Excel", 0, 1), ("Word", 0, 2),
            ("PowerPoint", 1, 0), ("Outlook", 1, 1), ("Publisher", 1, 2),
            ("OneNote", 2, 0), ("Skype", 2, 1), ("Teams", 2, 2)
        ]
        
        self.cac_bien_ung_dung = {}
        for ten_app, hang, cot in danh_sach_app:
            bien_tich = tk.BooleanVar(value=True)
            self.cac_bien_ung_dung[ten_app] = bien_tich
            tk.Checkbutton(khung_ung_dung, text=f" {ten_app}", variable=bien_tich, bg="white").grid(row=hang, column=cot, sticky="w", padx=25, pady=5)

        # 4. KHUNG CÁC NÚT BẤM (GRID MÀU SẮC)
        khung_nut_bam = tk.Frame(tab_cai_dat, bg="white")
        khung_nut_bam.pack(fill="x", padx=10, pady=10)
        
        tao_nut_bam_mau(khung_nut_bam, "⚙ CÀI ĐẶT", "#0288D1", hanh_dong=self.khoi_dong_luong_cai_dat).grid(row=0, column=0, padx=2, pady=2)
        tao_nut_bam_mau(khung_nut_bam, "🗑 Thoát", "#FFCDD2", mau_chu="#C62828", hanh_dong=self.cua_so.quit).grid(row=0, column=1, padx=2, pady=2)

        # 5. THANH TRẠNG THÁI DƯỚI CÙNG
        khung_trang_thai = tk.Frame(tab_cai_dat, bg="#E0F7FA")
        khung_trang_thai.pack(fill="x", padx=10, pady=20)
        
        self.nhan_trang_thai = tk.Label(khung_trang_thai, text="✅ Sẵn sàng kết nối tới máy chủ Microsoft...", font=self.phong_chu_dam, fg="#2E7D32", bg="#E0F7FA")
        self.nhan_trang_thai.pack(anchor="w", padx=10, pady=(10, 10))

    # ==========================================================================
    # CÁC HÀM XỬ LÝ LÕI BACKEND
    # ==========================================================================
    def tao_file_cauhinh_xml(self, phien_ban_tong_hop, kien_truc, ngon_ngu, danh_sach_app_chon):
        ma_san_pham = tu_dien_phien_ban.get(phien_ban_tong_hop, "ProPlus2024Retail") # Mặc định lấy ProPlus2024 nếu lỗi
        ma_ngon_ngu = "vi-vn" if "Vietnamese" in ngon_ngu else "en-us"
        
        noi_dung_xml = f"""<Configuration>
  <Add OfficeClientEdition="{kien_truc}" Channel="Current">
    <Product ID="{ma_san_pham}">
      <Language ID="{ma_ngon_ngu}" />\n"""

        for ten_app, ma_app in tu_dien_ung_dung.items():
            if ten_app not in danh_sach_app_chon:
                noi_dung_xml += f'      <ExcludeApp ID="{ma_app}" />\n'

        noi_dung_xml += """    </Product>
  </Add>
  <Updates Enabled="TRUE" Channel="Current" />
  <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>"""

        duong_dan_file_xml = os.path.join(os.getcwd(), "CauHinhCaiDat.xml")
        with open(duong_dan_file_xml, "w", encoding="utf-8") as file_xml:
            file_xml.write(noi_dung_xml)
            
        return duong_dan_file_xml

    def chuan_bi_cong_cu_odt(self):
        duong_dan_file_setup = os.path.join(os.getcwd(), "setup.exe")
        if not os.path.exists(duong_dan_file_setup):
            self.cap_nhat_trang_thai("⏳ Đang tải công cụ Office Deployment Tool từ Microsoft...")
            link_tai_odt = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A5D4A7E/officedeploymenttool_18230-20045.exe"
            file_tam_thoi = os.path.join(os.getcwd(), "cong_cu_tam.exe")
            
            try:
                urllib.request.urlretrieve(link_tai_odt, file_tam_thoi)
                self.cap_nhat_trang_thai("⏳ Đang giải nén công cụ cài đặt...")
                subprocess.run([file_tam_thoi, "/extract:" + os.getcwd(), "/quiet"], creationflags=subprocess.CREATE_NO_WINDOW)
                os.remove(file_tam_thoi)
            except Exception as loi_mang:
                messagebox.showerror("Lỗi hệ thống", f"Không thể lấy công cụ từ Microsoft: {loi_mang}")
                return None
                
        return duong_dan_file_setup

    def khoi_dong_luong_cai_dat(self):
        luong_xu_ly = threading.Thread(target=self.tien_trinh_cai_dat_ngam, daemon=True)
        luong_xu_ly.start()

    def tien_trinh_cai_dat_ngam(self):
        nam_phien_ban = self.bien_phien_ban.get()
        
        # Xử lý chuỗi hiển thị: "Home & Business" -> "HomeBusiness" để khớp với Từ điển
        ban_con_hien_thi = self.hop_chon_ban_con.get()
        ban_con_dinh_dang = ban_con_hien_thi.replace(" & ", "").replace(" ", "")
        
        phien_ban_tong_hop = f"{nam_phien_ban}_{ban_con_dinh_dang}"
        
        kien_truc_chon = self.bien_kien_truc.get()
        ngon_ngu_chon = self.hop_chon_ngon_ngu.get()

        danh_sach_app_duoc_chon = []
        for ten_app, bien_tich in self.cac_bien_ung_dung.items():
            if bien_tich.get() == True:
                danh_sach_app_duoc_chon.append(ten_app)

        self.cap_nhat_trang_thai("⏳ Đang thiết lập cấu hình XML...")
        duong_dan_xml = self.tao_file_cauhinh_xml(phien_ban_tong_hop, kien_truc_chon, ngon_ngu_chon, danh_sach_app_duoc_chon)
        
        duong_dan_setup = self.chuan_bi_cong_cu_odt()
        
        if duong_dan_setup and os.path.exists(duong_dan_setup):
            self.cap_nhat_trang_thai("🚀 Đang mở trình cài đặt của Microsoft. Băng thông Max Speed!")
            try:
                subprocess.Popen([duong_dan_setup, "/configure", duong_dan_xml])
            except Exception as loi_thuc_thi:
                messagebox.showerror("Lỗi chạy ứng dụng", f"Không thể bắt đầu cài đặt: {loi_thuc_thi}")
                self.cap_nhat_trang_thai("❌ Đã xảy ra lỗi.")
        else:
            self.cap_nhat_trang_thai("❌ Lỗi: Thiếu công cụ setup.exe!")

    def cap_nhat_trang_thai(self, noi_dung_thong_bao):
        self.cua_so.after(0, lambda: self.nhan_trang_thai.config(text=noi_dung_thong_bao))

# ==============================================================================
# KHỞI CHẠY CHƯƠNG TRÌNH
# ==============================================================================
if __name__ == "__main__":
    cua_so_giao_dien = tk.Tk()
    ung_dung = UngDungCaiDatOffice(cua_so_giao_dien)
    cua_so_giao_dien.mainloop()