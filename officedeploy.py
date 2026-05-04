import os
import subprocess
import urllib.request
import threading
import tkinter as tk
from tkinter import ttk, messagebox

# ==============================================================================
# TỪ ĐIỂN DỮ LIỆU CỦA MICROSOFT
# ==============================================================================
tu_dien_phien_ban = {
    "2016_ProPlus": "ProPlusRetail",
    "2019_ProPlus": "ProPlus2019Retail",
    "2021_ProPlus": "ProPlus2021Retail",
    "2024_ProPlus": "ProPlus2024Retail",
    "365_ProPlus": "O365ProPlusRetail",
    "2016_Standard": "StandardRetail",
    "2019_Standard": "Standard2019Retail",
    "2021_Standard": "Standard2021Retail",
    "2024_Standard": "Standard2024Retail",
    "2016_HomeBusiness": "HomeBusinessRetail",
    "2019_HomeBusiness": "HomeBusiness2019Retail",
    "2021_HomeBusiness": "HomeBusiness2021Retail",
    "2024_HomeBusiness": "HomeBusiness2024Retail",
    "365_HomeBusiness": "O365BusinessRetail",
    "2016_HomeStudent": "HomeStudentRetail",
    "2019_HomeStudent": "HomeStudent2019Retail",
    "2021_HomeStudent": "HomeStudent2021Retail",
    "2024_HomeStudent": "Home2024Retail",
    "365_HomeStudent": "O365HomePremRetail"
}

tu_dien_ung_dung = {
    "Access": "Access", "Excel": "Excel", "Word": "Word",
    "PowerPoint": "PowerPoint", "Outlook": "Outlook", "Publisher": "Publisher",
    "OneNote": "OneNote", "Skype": "Lync", "Teams": "Teams"
}

# ==============================================================================
# HÀM HỖ TRỢ GIAO DIỆN
# ==============================================================================
def tao_nut_bam_mau(khung_chua, chu_hien_thi, mau_nen, mau_chu="white", hanh_dong=None):
    nut_bam = tk.Button(
        khung_chua, text=chu_hien_thi, bg=mau_nen, fg=mau_chu,
        font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2",
        command=hanh_dong, width=18, pady=6
    )
    return nut_bam

# ==============================================================================
# LỚP GIAO DIỆN VÀ XỬ LÝ CHÍNH
# ==============================================================================
class UngDungCaiDatOffice:
    def __init__(self, cua_so_chinh):
        self.cua_so = cua_so_chinh
        self.cua_so.title("Cài đặt Microsoft Office - Trực tiếp từ Server CDN")
        self.cua_so.geometry("620x680")
        self.cua_so.resizable(False, False)

        self.phong_chu_thuong = ("Segoe UI", 9)
        self.phong_chu_dam = ("Segoe UI", 9, "bold")

        self.xay_dung_giao_dien()

    def xay_dung_giao_dien(self):
        # --- TIÊU ĐỀ HEADER ---
        khung_tieu_de = tk.Frame(self.cua_so, bg="#E64A19", height=65)
        khung_tieu_de.pack(fill="x", side="top")
        
        tk.Label(khung_tieu_de, text="🏢 CÀI ĐẶT MICROSOFT OFFICE RETAIL", font=("Segoe UI", 14, "bold"), bg="#E64A19", fg="white").place(x=15, y=8)
        tk.Label(khung_tieu_de, text="Tải trực tiếp Max Speed từ Server Microsoft (CDN)", font=("Segoe UI", 9), bg="#E64A19", fg="white").place(x=40, y=35)
        tk.Label(khung_tieu_de, text="🌐 ONLINE", font=("Segoe UI", 10, "bold"), bg="#E64A19", fg="white").place(x=520, y=18)

        # --- KHUNG CHỨA TAB ---
        hop_tab = ttk.Notebook(self.cua_so)
        hop_tab.pack(fill="both", expand=True, padx=12, pady=12)

        tab_cai_dat = ttk.Frame(hop_tab)
        tab_go_cai_dat = ttk.Frame(hop_tab)

        hop_tab.add(tab_cai_dat, text="  ⚙️ Cài đặt Office  ")
        hop_tab.add(tab_go_cai_dat, text="  🗑️ Gỡ Cài Đặt & Dọn Dẹp  ")

        self.xay_dung_tab_cai_dat(tab_cai_dat)
        self.xay_dung_tab_go_cai_dat(tab_go_cai_dat)

        # --- THANH TRẠNG THÁI CHUNG ---
        khung_trang_thai = tk.Frame(self.cua_so, bg="#E0F7FA")
        khung_trang_thai.pack(fill="x", padx=12, pady=(0, 12))
        
        self.nhan_trang_thai = tk.Label(khung_trang_thai, text="✅ Sẵn sàng kết nối tới máy chủ Microsoft...", font=self.phong_chu_dam, fg="#2E7D32", bg="#E0F7FA")
        self.nhan_trang_thai.pack(anchor="w", padx=10, pady=10)

    # --------------------------------------------------------------------------
    # TAB 1: CÀI ĐẶT OFFICE
    # --------------------------------------------------------------------------
    def xay_dung_tab_cai_dat(self, tab):
        # 1. KHUNG PHIÊN BẢN
        khung_phien_ban = ttk.LabelFrame(tab, text=" 📄 Phiên bản Office ")
        khung_phien_ban.pack(fill="x", padx=10, pady=10, ipady=5)
        
        self.bien_phien_ban = tk.StringVar(value="2024")
        danh_sach_ban = ["Office 2016", "Office 2019", "Office 2021", "Office 2024", "Office 365"]
        
        khung_rb_phien_ban = ttk.Frame(khung_phien_ban)
        khung_rb_phien_ban.pack(fill="x", padx=10, pady=5)
        for vi_tri, ten_ban in enumerate(danh_sach_ban):
            ttk.Radiobutton(khung_rb_phien_ban, text=ten_ban, variable=self.bien_phien_ban, value=ten_ban.split()[-1]).grid(row=0, column=vi_tri, padx=8)
            
        khung_cb_phien_ban = ttk.Frame(khung_phien_ban)
        khung_cb_phien_ban.pack(fill="x", padx=10, pady=5)
        ttk.Label(khung_cb_phien_ban, text="Phiên bản con:").pack(side="left", padx=(0, 10))
        self.hop_chon_ban_con = ttk.Combobox(khung_cb_phien_ban, values=["ProPlus", "Standard", "Home & Business", "Home & Student"], state="readonly", width=35)
        self.hop_chon_ban_con.current(0)
        self.hop_chon_ban_con.pack(side="left")

        # 2. KHUNG KIẾN TRÚC & NGÔN NGỮ
        khung_ngang = ttk.Frame(tab)
        khung_ngang.pack(fill="x", padx=10, pady=5)
        
        khung_kien_truc = ttk.LabelFrame(khung_ngang, text=" ⚙️ Kiến trúc ")
        khung_kien_truc.pack(side="left", fill="both", expand=True, padx=(0, 5))
        self.bien_kien_truc = tk.StringVar(value="64")
        ttk.Radiobutton(khung_kien_truc, text="64-bit", variable=self.bien_kien_truc, value="64").pack(side="left", padx=15, pady=10)
        ttk.Radiobutton(khung_kien_truc, text="32-bit", variable=self.bien_kien_truc, value="32").pack(side="right", padx=15, pady=10)

        khung_ngon_ngu = ttk.LabelFrame(khung_ngang, text=" 🌐 Ngôn ngữ ")
        khung_ngon_ngu.pack(side="right", fill="both", expand=True, padx=(5, 0))
        self.hop_chon_ngon_ngu = ttk.Combobox(khung_ngon_ngu, values=["English (US) - en-us", "Vietnamese - vi-vn"], state="readonly")
        self.hop_chon_ngon_ngu.current(0)
        self.hop_chon_ngon_ngu.pack(fill="x", padx=10, pady=10)

        # 3. KHUNG ỨNG DỤNG
        khung_ung_dung = ttk.LabelFrame(tab, text=" ☑️ Chọn ứng dụng cần cài ")
        khung_ung_dung.pack(fill="x", padx=10, pady=10)
        
        danh_sach_app = [
            ("Access", 0, 0), ("Excel", 0, 1), ("Word", 0, 2),
            ("PowerPoint", 1, 0), ("Outlook", 1, 1), ("Publisher", 1, 2),
            ("OneNote", 2, 0), ("Skype", 2, 1), ("Teams", 2, 2)
        ]
        
        self.cac_bien_ung_dung = {}
        for ten_app, hang, cot in danh_sach_app:
            bien_tich = tk.BooleanVar(value=True)
            self.cac_bien_ung_dung[ten_app] = bien_tich
            ttk.Checkbutton(khung_ung_dung, text=f" {ten_app}", variable=bien_tich).grid(row=hang, column=cot, sticky="w", padx=25, pady=8)

        # 4. NÚT BẤM CÀI ĐẶT
        khung_nut_bam = ttk.Frame(tab)
        khung_nut_bam.pack(fill="x", padx=10, pady=10)
        tao_nut_bam_mau(khung_nut_bam, "⚙ BẮT ĐẦU CÀI ĐẶT", "#0288D1", hanh_dong=self.khoi_dong_luong_cai_dat).pack(side="right", padx=5)

    # --------------------------------------------------------------------------
    # TAB 2: GỠ CÀI ĐẶT & DỌN DẸP
    # --------------------------------------------------------------------------
    def xay_dung_tab_go_cai_dat(self, tab):
        # 1. KHUNG GỠ OFFICE
        khung_go_office = ttk.LabelFrame(tab, text=" 🗑️ Gỡ Cài Đặt Office Toàn Diện ")
        khung_go_office.pack(fill="x", padx=10, pady=15)
        
        ttk.Label(khung_go_office, text="Thao tác này sẽ gọi trình gỡ cài đặt gốc của Microsoft để xóa sạch\ntoàn bộ các phiên bản Office (Word, Excel...) đang có trên máy tính.", justify="left").pack(anchor="w", padx=15, pady=10)
        
        khung_nut_go = ttk.Frame(khung_go_office)
        khung_nut_go.pack(fill="x", padx=10, pady=5)
        tao_nut_bam_mau(khung_nut_go, "🗑 BẮT ĐẦU GỠ OFFICE", "#D32F2F", hanh_dong=self.khoi_dong_luong_go_office).pack(side="right", padx=5, pady=5)

        # 2. KHUNG GỠ CRACK / RESET BẢN QUYỀN
        khung_go_crack = ttk.LabelFrame(tab, text=" 🧹 Gỡ Crack / Reset Bản Quyền (KMS) ")
        khung_go_crack.pack(fill="x", padx=10, pady=10)
        
        ttk.Label(khung_go_crack, text="Dọn dẹp các máy chủ KMS ảo, làm sạch Registry kích hoạt lậu và\nđưa trạng thái bản quyền Office về như máy mới (Rearm).", justify="left").pack(anchor="w", padx=15, pady=10)
        
        khung_nut_crack = ttk.Frame(khung_go_crack)
        khung_nut_crack.pack(fill="x", padx=10, pady=5)
        tao_nut_bam_mau(khung_nut_crack, "🧹 DỌN DẸP BẢN QUYỀN", "#F57C00", hanh_dong=self.tien_trinh_go_crack_ngam).pack(side="right", padx=5, pady=5)

    # ==========================================================================
    # CÁC HÀM XỬ LÝ LÕI BACKEND
    # ==========================================================================
    def chuan_bi_cong_cu_odt(self):
        duong_dan_file_setup = os.path.join(os.getcwd(), "setup.exe")
        if not os.path.exists(duong_dan_file_setup):
            self.cap_nhat_trang_thai("⏳ Đang tải công cụ Office Deployment Tool từ Microsoft...")
            link_tai_odt = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A5D4A7E/officedeploymenttool_18230-20045.exe"
            file_tam_thoi = os.path.join(os.getcwd(), "cong_cu_tam.exe")
            try:
                urllib.request.urlretrieve(link_tai_odt, file_tam_thoi)
                self.cap_nhat_trang_thai("⏳ Đang giải nén công cụ...")
                subprocess.run([file_tam_thoi, "/extract:" + os.getcwd(), "/quiet"], creationflags=subprocess.CREATE_NO_WINDOW)
                os.remove(file_tam_thoi)
            except Exception as loi_mang:
                messagebox.showerror("Lỗi mạng", f"Không thể lấy công cụ từ Microsoft: {loi_mang}")
                return None
        return duong_dan_file_setup

    # ---------- LOGIC CÀI ĐẶT ----------
    def tao_file_xml_cai_dat(self, phien_ban_tong_hop, kien_truc, ngon_ngu, danh_sach_app_chon):
        ma_san_pham = tu_dien_phien_ban.get(phien_ban_tong_hop, "ProPlus2024Retail")
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

        duong_dan = os.path.join(os.getcwd(), "CauHinhCaiDat.xml")
        with open(duong_dan, "w", encoding="utf-8") as f: f.write(noi_dung_xml)
        return duong_dan

    def khoi_dong_luong_cai_dat(self):
        threading.Thread(target=self.tien_trinh_cai_dat_ngam, daemon=True).start()

    def tien_trinh_cai_dat_ngam(self):
        nam_phien_ban = self.bien_phien_ban.get()
        ban_con_dinh_dang = self.hop_chon_ban_con.get().replace(" & ", "").replace(" ", "")
        phien_ban_tong_hop = f"{nam_phien_ban}_{ban_con_dinh_dang}"
        kien_truc_chon = self.bien_kien_truc.get()
        ngon_ngu_chon = self.hop_chon_ngon_ngu.get()

        danh_sach_app_duoc_chon = [ten for ten, bien in self.cac_bien_ung_dung.items() if bien.get()]

        self.cap_nhat_trang_thai("⏳ Đang thiết lập cấu hình cài đặt...")
        duong_dan_xml = self.tao_file_xml_cai_dat(phien_ban_tong_hop, kien_truc_chon, ngon_ngu_chon, danh_sach_app_duoc_chon)
        duong_dan_setup = self.chuan_bi_cong_cu_odt()
        
        if duong_dan_setup and os.path.exists(duong_dan_setup):
            self.cap_nhat_trang_thai("🚀 Đang chạy trình cài đặt Microsoft...")
            try:
                subprocess.Popen([duong_dan_setup, "/configure", duong_dan_xml])
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể bắt đầu cài đặt: {e}")
                self.cap_nhat_trang_thai("❌ Đã xảy ra lỗi.")
        else:
            self.cap_nhat_trang_thai("❌ Lỗi: Thiếu công cụ setup.exe!")

    # ---------- LOGIC GỠ CÀI ĐẶT ----------
    def khoi_dong_luong_go_office(self):
        hoi_dap = messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn gỡ toàn bộ Office khỏi máy tính không?")
        if hoi_dap:
            threading.Thread(target=self.tien_trinh_go_office_ngam, daemon=True).start()

    def tien_trinh_go_office_ngam(self):
        self.cap_nhat_trang_thai("⏳ Đang thiết lập cấu hình gỡ cài đặt...")
        noi_dung_xml = """<Configuration>\n  <Remove All="True" />\n  <Display Level="Full" AcceptEULA="TRUE" />\n</Configuration>"""
        duong_dan_xml = os.path.join(os.getcwd(), "CauHinhGoCaiDat.xml")
        
        with open(duong_dan_xml, "w", encoding="utf-8") as f: f.write(noi_dung_xml)
        duong_dan_setup = self.chuan_bi_cong_cu_odt()
        
        if duong_dan_setup and os.path.exists(duong_dan_setup):
            self.cap_nhat_trang_thai("🚀 Đang chạy trình gỡ cài đặt của Microsoft...")
            try:
                subprocess.Popen([duong_dan_setup, "/configure", duong_dan_xml])
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể bắt đầu gỡ cài đặt: {e}")
                self.cap_nhat_trang_thai("❌ Đã xảy ra lỗi.")

    # ---------- LOGIC GỠ CRACK / KMS ----------
    def tien_trinh_go_crack_ngam(self):
        hoi_dap = messagebox.askyesno("Xác nhận", "Bạn có muốn xóa thông tin máy chủ KMS và Reset trạng thái bản quyền Office không?")
        if not hoi_dap: return
        
        self.cap_nhat_trang_thai("⏳ Đang tìm kiếm và dọn dẹp hệ thống KMS/Crack...")
        
        # Đường dẫn mặc định của script quản lý bản quyền Office
        cac_thu_muc = [
            os.environ.get("ProgramFiles", "C:\\Program Files") + "\\Microsoft Office\\Office16",
            os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)") + "\\Microsoft Office\\Office16"
        ]
        
        file_ospp = None
        for thu_muc in cac_thu_muc:
            if os.path.exists(os.path.join(thu_muc, "ospp.vbs")):
                file_ospp = os.path.join(thu_muc, "ospp.vbs")
                break
                
        if file_ospp:
            try:
                # 1. Gỡ bỏ địa chỉ máy chủ KMS ảo
                subprocess.run(["cscript", "//nologo", file_ospp, "/remhst"], creationflags=subprocess.CREATE_NO_WINDOW)
                # 2. Đặt lại trạng thái bản quyền gốc (Rearm)
                subprocess.run(["cscript", "//nologo", file_ospp, "/rearm"], creationflags=subprocess.CREATE_NO_WINDOW)
                
                self.cap_nhat_trang_thai("✅ Đã dọn dẹp sạch bản quyền KMS cũ!")
                messagebox.showinfo("Thành công", "Đã xóa máy chủ KMS ảo và Reset trạng thái bản quyền thành công.\nVui lòng khởi động lại máy tính để thay đổi có hiệu lực triệt để.")
            except Exception as e:
                self.cap_nhat_trang_thai("❌ Lỗi trong quá trình dọn dẹp.")
        else:
            self.cap_nhat_trang_thai("⚠️ Không tìm thấy file quản lý bản quyền gốc. Có thể Office đã bị gỡ.")
            messagebox.showwarning("Thông báo", "Máy tính hiện không có Office 16/19/21/24/365 hoặc file bản quyền hệ thống đã bị xóa.")

    def cap_nhat_trang_thai(self, noi_dung_thong_bao):
        self.cua_so.after(0, lambda: self.nhan_trang_thai.config(text=noi_dung_thong_bao))

# ==============================================================================
# KHỞI CHẠY CHƯƠNG TRÌNH
# ==============================================================================
if __name__ == "__main__":
    cua_so_giao_dien = tk.Tk()
    ung_dung = UngDungCaiDatOffice(cua_so_giao_dien)
    cua_so_giao_dien.mainloop()