import os
import re
import subprocess
import urllib.request
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox

# ==============================================================================
# TỪ ĐIỂN DỮ LIỆU CỦA MICROSOFT
# ==============================================================================
tu_dien_phien_ban = {
    "2016_ProPlus": "ProPlusRetail", "2019_ProPlus": "ProPlus2019Retail",
    "2021_ProPlus": "ProPlus2021Retail", "2024_ProPlus": "ProPlus2024Retail",
    "365_ProPlus": "O365ProPlusRetail", "2016_Standard": "StandardRetail",
    "2019_Standard": "Standard2019Retail", "2021_Standard": "Standard2021Retail",
    "2024_Standard": "Standard2024Retail", "2016_HomeBusiness": "HomeBusinessRetail",
    "2019_HomeBusiness": "HomeBusiness2019Retail", "2021_HomeBusiness": "HomeBusiness2021Retail",
    "2024_HomeBusiness": "HomeBusiness2024Retail", "365_Business": "O365BusinessRetail",
    "2016_HomeStudent": "HomeStudentRetail", "2019_HomeStudent": "HomeStudent2019Retail",
    "2021_HomeStudent": "HomeStudent2021Retail", "2024_HomeStudent": "Home2024Retail",
    "365_HomePremium": "O365HomePremRetail"
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
        command=hanh_dong, width=20, pady=6
    )
    return nut_bam

# ==============================================================================
# LỚP GIAO DIỆN VÀ XỬ LÝ CHÍNH
# ==============================================================================
class UngDungCaiDatOffice:
    def __init__(self, cua_so_chinh):
        self.cua_so = cua_so_chinh
        self.cua_so.title("Cài đặt Microsoft Office - Trực tiếp từ Server CDN (V5.3)")
        self.cua_so.geometry("640x660")
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
        tk.Label(khung_tieu_de, text="🌐 ONLINE", font=("Segoe UI", 10, "bold"), bg="#E64A19", fg="white").place(x=540, y=18)

        # --- KHUNG CHỨA TAB ---
        hop_tab = ttk.Notebook(self.cua_so)
        hop_tab.pack(fill="both", expand=True, padx=12, pady=12)

        tab_cai_dat = ttk.Frame(hop_tab)
        tab_go_cai_dat = ttk.Frame(hop_tab)

        hop_tab.add(tab_cai_dat, text="  ⚙️ Cài đặt & Kích hoạt  ")
        hop_tab.add(tab_go_cai_dat, text="  🗑️ Gỡ Cài Đặt & Dọn Dẹp  ")

        self.xay_dung_tab_cai_dat(tab_cai_dat)
        self.xay_dung_tab_go_cai_dat(tab_go_cai_dat)

        # --- THANH TRẠNG THÁI CHUNG ---
        khung_trang_thai = tk.Frame(self.cua_so, bg="#E0F7FA")
        khung_trang_thai.pack(fill="x", padx=12, pady=(0, 12))
        
        self.nhan_trang_thai = tk.Label(khung_trang_thai, text="✅ Sẵn sàng kết nối tới hệ thống...", font=self.phong_chu_dam, fg="#2E7D32", bg="#E0F7FA")
        self.nhan_trang_thai.pack(anchor="w", padx=10, pady=(10, 5))
        
        # --- THANH TIẾN TRÌNH (PROGRESS BAR) ---
        self.thanh_tien_do = ttk.Progressbar(khung_trang_thai, mode='indeterminate')
        self.thanh_tien_do.pack(fill="x", padx=12, pady=(0, 10))

    # --------------------------------------------------------------------------
    # TAB 1: CÀI ĐẶT & KÍCH HOẠT OFFICE
    # --------------------------------------------------------------------------
    def xay_dung_tab_cai_dat(self, tab):
        khung_phien_ban = ttk.LabelFrame(tab, text=" 📄 Phiên bản Office ")
        khung_phien_ban.pack(fill="x", padx=10, pady=8, ipady=3)
        
        self.bien_phien_ban = tk.StringVar(value="2024")
        danh_sach_ban = ["Office 2016", "Office 2019", "Office 2021", "Office 2024", "Office 365"]
        
        khung_rb_phien_ban = ttk.Frame(khung_phien_ban)
        khung_rb_phien_ban.pack(fill="x", padx=10, pady=2)
        
        self.bien_phien_ban.trace_add("write", self.cap_nhat_hop_chon_ban_con)
        for vi_tri, ten_ban in enumerate(danh_sach_ban):
            ttk.Radiobutton(khung_rb_phien_ban, text=ten_ban, variable=self.bien_phien_ban, value=ten_ban.split()[-1]).grid(row=0, column=vi_tri, padx=6)
            
        khung_cb_phien_ban = ttk.Frame(khung_phien_ban)
        khung_cb_phien_ban.pack(fill="x", padx=10, pady=5)
        ttk.Label(khung_cb_phien_ban, text="Phiên bản con:").pack(side="left", padx=(0, 10))
        self.hop_chon_ban_con = ttk.Combobox(khung_cb_phien_ban, state="readonly", width=35)
        self.cap_nhat_hop_chon_ban_con()
        self.hop_chon_ban_con.pack(side="left")

        khung_ngang = ttk.Frame(tab)
        khung_ngang.pack(fill="x", padx=10, pady=5)
        
        khung_kien_truc = ttk.LabelFrame(khung_ngang, text=" ⚙️ Kiến trúc ")
        khung_kien_truc.pack(side="left", fill="both", expand=True, padx=(0, 5))
        self.bien_kien_truc = tk.StringVar(value="64")
        ttk.Radiobutton(khung_kien_truc, text="64-bit", variable=self.bien_kien_truc, value="64").pack(side="left", padx=15, pady=5)
        ttk.Radiobutton(khung_kien_truc, text="32-bit", variable=self.bien_kien_truc, value="32").pack(side="right", padx=15, pady=5)

        khung_ngon_ngu = ttk.LabelFrame(khung_ngang, text=" 🌐 Ngôn ngữ ")
        khung_ngon_ngu.pack(side="right", fill="both", expand=True, padx=(5, 0))
        self.hop_chon_ngon_ngu = ttk.Combobox(khung_ngon_ngu, values=["English (US) - en-us", "Vietnamese - vi-vn"], state="readonly")
        self.hop_chon_ngon_ngu.current(0)
        self.hop_chon_ngon_ngu.pack(fill="x", padx=10, pady=6)

        khung_ung_dung = ttk.LabelFrame(tab, text=" ☑️ Chọn ứng dụng cần cài ")
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
            ttk.Checkbutton(khung_ung_dung, text=f" {ten_app}", variable=bien_tich).grid(row=hang, column=cot, sticky="w", padx=25, pady=4)

        khung_tuy_chon = ttk.LabelFrame(tab, text=" 🔧 Tùy chọn tự động sau cài đặt ")
        khung_tuy_chon.pack(fill="x", padx=10, pady=5)
        
        self.bien_tu_dong_ohook = tk.BooleanVar(value=True)
        ttk.Checkbutton(khung_tuy_chon, text="Tự động Kích hoạt Ohook", variable=self.bien_tu_dong_ohook).pack(side="left", padx=25, pady=6)
        
        self.bien_tao_shortcut = tk.BooleanVar(value=True)
        ttk.Checkbutton(khung_tuy_chon, text="Đưa lối tắt (Word, Excel) ra Desktop", variable=self.bien_tao_shortcut).pack(side="left", padx=25, pady=6)

        khung_nut_bam = ttk.Frame(tab)
        khung_nut_bam.pack(fill="x", padx=10, pady=5)
        tao_nut_bam_mau(khung_nut_bam, "💊 KÍCH HOẠT OHOOK", "#F57C00", hanh_dong=self.khoi_dong_thuoc_ohook).pack(side="left", padx=5)
        tao_nut_bam_mau(khung_nut_bam, "⚙ BẮT ĐẦU CÀI ĐẶT", "#0288D1", hanh_dong=self.khoi_dong_luong_cai_dat).pack(side="right", padx=5)

    # --------------------------------------------------------------------------
    # TAB 2: GỠ CÀI ĐẶT & DỌN DẸP
    # --------------------------------------------------------------------------
    def xay_dung_tab_go_cai_dat(self, tab):
        khung_go_office = ttk.LabelFrame(tab, text=" 🗑️ Gỡ Cài Đặt Office Toàn Diện ")
        khung_go_office.pack(fill="x", padx=10, pady=10)
        ttk.Label(khung_go_office, text="Thao tác này sẽ gọi trình gỡ cài đặt gốc của Microsoft để xóa sạch\ntoàn bộ các phiên bản Office (Word, Excel...) đang có trên máy tính.", justify="left").pack(anchor="w", padx=15, pady=5)
        khung_nut_go = ttk.Frame(khung_go_office)
        khung_nut_go.pack(fill="x", padx=10, pady=5)
        tao_nut_bam_mau(khung_nut_go, "🗑 BẮT ĐẦU GỠ OFFICE", "#D32F2F", hanh_dong=self.khoi_dong_luong_go_office).pack(side="right", padx=5, pady=5)

        khung_go_kms = ttk.LabelFrame(tab, text=" 🧹 Gỡ Crack KMS / Reset Bản Quyền ")
        khung_go_kms.pack(fill="x", padx=10, pady=10)
        ttk.Label(khung_go_kms, text="Dọn dẹp các máy chủ KMS ảo, làm sạch Registry kích hoạt lậu và\nđưa trạng thái bản quyền Office về như máy mới (Rearm).", justify="left").pack(anchor="w", padx=15, pady=5)
        khung_nut_kms = ttk.Frame(khung_go_kms)
        khung_nut_kms.pack(fill="x", padx=10, pady=5)
        tao_nut_bam_mau(khung_nut_kms, "🧹 DỌN DẸP KMS", "#1976D2", hanh_dong=self.tien_trinh_go_kms_ngam).pack(side="right", padx=5, pady=5)

        khung_go_ohook = ttk.LabelFrame(tab, text=" 💊 Gỡ Crack Ohook ")
        khung_go_ohook.pack(fill="x", padx=10, pady=10)
        ttk.Label(khung_go_ohook, text="Gỡ bỏ hoàn toàn Hook kích hoạt bản quyền vĩnh viễn (Ohook)\nvà trả lại các file hệ thống nguyên bản cho ứng dụng Office.", justify="left").pack(anchor="w", padx=15, pady=5)
        khung_nut_ohook = ttk.Frame(khung_go_ohook)
        khung_nut_ohook.pack(fill="x", padx=10, pady=5)
        tao_nut_bam_mau(khung_nut_ohook, "🛡️ GỠ OHOOK TRIỆT ĐỂ", "#7B1FA2", hanh_dong=self.khoi_dong_go_ohook).pack(side="right", padx=5, pady=5)

    # ==========================================================================
    # QUẢN LÝ TIẾN TRÌNH VÀ GIAO DIỆN
    # ==========================================================================
    def dieu_khien_thanh_tien_do(self, bat_dau=True):
        if bat_dau:
            self.cua_so.after(0, lambda: self.thanh_tien_do.start(15))
        else:
            self.cua_so.after(0, lambda: self.thanh_tien_do.stop())

    def cap_nhat_hop_chon_ban_con(self, *args):
        ban_dang_chon = self.bien_phien_ban.get()
        if ban_dang_chon == "365":
            self.hop_chon_ban_con['values'] = ["ProPlus", "Business", "Home Premium"]
        else:
            self.hop_chon_ban_con['values'] = ["ProPlus", "Standard", "Home & Business", "Home & Student"]
        self.hop_chon_ban_con.current(0)

    def chuan_bi_cong_cu_odt(self):
        duong_dan_file_setup = os.path.join(os.getcwd(), "setup.exe")
        
        # Nếu đã tải tool về máy rồi thì bỏ qua không tải lại nữa
        if os.path.exists(duong_dan_file_setup):
            return duong_dan_file_setup

        self.cap_nhat_trang_thai("⏳ Đang tải công cụ Office Deployment Tool từ Microsoft...")
        link_tai_odt = ""
        
        # Đeo mặt nạ trình duyệt để không bị Microsoft chặn
        tieu_de_gia_mao = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'}

        try:
            yeu_cau_quet = urllib.request.Request("https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117", headers=tieu_de_gia_mao)
            trang_tai_ve = urllib.request.urlopen(yeu_cau_quet, timeout=10).read().decode('utf-8')
            
            # Quét tìm đường link file .exe mới nhất
            mau_tim = re.search(r'(https://download\.microsoft\.com/download/[^\s"\'<>]+officedeploymenttool_[^\s"\'<>]+\.exe)', trang_tai_ve, re.IGNORECASE)
            if mau_tim:
                link_tai_odt = mau_tim.group(1)
        except:
            pass

        # Phương án dự phòng (Fallback) nếu trang Microsoft lỗi
        if not link_tai_odt:
            link_tai_odt = "https://raw.githubusercontent.com/tuantran19912512/pythonoffice/main/setup.exe"

        file_tam_thoi = os.path.join(os.getcwd(), "cong_cu_tam.exe")
        try:
            # Tải file về máy
            yeu_cau_tai = urllib.request.Request(link_tai_odt, headers=tieu_de_gia_mao)
            with urllib.request.urlopen(yeu_cau_tai, timeout=30) as phan_hoi, open(file_tam_thoi, 'wb') as file_luu:
                file_luu.write(phan_hoi.read())

            self.cap_nhat_trang_thai("⏳ Đang giải nén công cụ...")
            
            # Xử lý: Giải nén nếu là file gốc của Microsoft, đổi tên nếu là file từ GitHub
            if "officedeploymenttool" in link_tai_odt.lower():
                subprocess.run([file_tam_thoi, "/extract:" + os.getcwd(), "/quiet"], creationflags=subprocess.CREATE_NO_WINDOW)
                os.remove(file_tam_thoi)
            else:
                os.rename(file_tam_thoi, duong_dan_file_setup)

        except Exception as loi_mang:
            messagebox.showerror("Lỗi mạng", f"Không thể lấy công cụ từ Microsoft.\n\nChi tiết: {loi_mang}")
            return None

        return duong_dan_file_setup if os.path.exists(duong_dan_file_setup) else None

    def tao_loi_tat_desktop(self):
        self.cap_nhat_trang_thai("⏳ Đang tìm và đưa Shortcut ra màn hình Desktop...")
        desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
        cac_thu_muc = [
            os.environ.get('ProgramFiles', 'C:\\Program Files') + "\\Microsoft Office\\root\\Office16",
            os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)') + "\\Microsoft Office\\root\\Office16"
        ]
        cac_ung_dung = {"WINWORD.EXE": "Word", "EXCEL.EXE": "Excel", "POWERPNT.EXE": "PowerPoint", "MSACCESS.EXE": "Access", "OUTLOOK.EXE": "Outlook"}
        
        for thu_muc in cac_thu_muc:
            if os.path.exists(thu_muc):
                for file_exe, ten_app in cac_ung_dung.items():
                    muc_tieu = os.path.join(thu_muc, file_exe)
                    if os.path.exists(muc_tieu):
                        duong_dan_lnk = os.path.join(desktop, f"{ten_app}.lnk")
                        ma_vbs = f'Set ws = CreateObject("WScript.Shell")\nSet link = ws.CreateShortcut("{duong_dan_lnk}")\nlink.TargetPath = "{muc_tieu}"\nlink.Save'
                        file_vbs = os.path.join(os.environ['TEMP'], 'tao_shortcut.vbs')
                        with open(file_vbs, 'w', encoding='utf-8') as f: f.write(ma_vbs)
                        subprocess.run(['cscript', '//nologo', file_vbs], creationflags=subprocess.CREATE_NO_WINDOW)
                break

    # ---------- LOGIC CÀI ĐẶT ----------
    def tao_file_xml_cai_dat(self, phien_ban_tong_hop, kien_truc, ngon_ngu, danh_sach_app_chon):
        ma_san_pham = tu_dien_phien_ban.get(phien_ban_tong_hop, "ProPlus2024Retail")
        ma_ngon_ngu = "vi-vn" if "Vietnamese" in ngon_ngu else "en-us"
        
        noi_dung_xml = f"""<Configuration>\n  <Add OfficeClientEdition="{kien_truc}" Channel="Current">\n    <Product ID="{ma_san_pham}">\n      <Language ID="{ma_ngon_ngu}" />\n"""
        for ten_app, ma_app in tu_dien_ung_dung.items():
            if ten_app not in danh_sach_app_chon:
                noi_dung_xml += f'      <ExcludeApp ID="{ma_app}" />\n'
        noi_dung_xml += """    </Product>\n  </Add>\n  <Updates Enabled="TRUE" Channel="Current" />\n  <Display Level="Full" AcceptEULA="TRUE" />\n</Configuration>"""

        duong_dan = os.path.join(os.getcwd(), "CauHinhCaiDat.xml")
        with open(duong_dan, "w", encoding="utf-8") as f: f.write(noi_dung_xml)
        return duong_dan

    def khoi_dong_luong_cai_dat(self):
        threading.Thread(target=self.tien_trinh_cai_dat_ngam, daemon=True).start()

    def tien_trinh_cai_dat_ngam(self):
        self.dieu_khien_thanh_tien_do(True)
        try:
            nam_phien_ban = self.bien_phien_ban.get()
            ban_con_dinh_dang = self.hop_chon_ban_con.get().replace(" & ", "").replace(" ", "")
            phien_ban_tong_hop = f"{nam_phien_ban}_{ban_con_dinh_dang}"
            kien_truc_chon = self.bien_kien_truc.get()
            ngon_ngu_chon = self.hop_chon_ngon_ngu.get()
            danh_sach_app_duoc_chon = [ten for ten, bien in self.cac_bien_ung_dung.items() if bien.get()]

            self.cap_nhat_trang_thai("⏳ Đang thiết lập cấu hình cài đặt XML...")
            duong_dan_xml = self.tao_file_xml_cai_dat(phien_ban_tong_hop, kien_truc_chon, ngon_ngu_chon, danh_sach_app_duoc_chon)
            duong_dan_setup = self.chuan_bi_cong_cu_odt()
            
            if duong_dan_setup and os.path.exists(duong_dan_setup):
                self.cap_nhat_trang_thai("🚀 Đang chạy trình cài đặt. Vui lòng đợi đến khi bảng màu cam tắt hẳn...")
                tien_trinh_cai_dat = subprocess.Popen([duong_dan_setup, "/configure", duong_dan_xml])
                tien_trinh_cai_dat.wait()
                
                if self.bien_tao_shortcut.get():
                    self.tao_loi_tat_desktop()
                    
                if self.bien_tu_dong_ohook.get():
                    self.cap_nhat_trang_thai("⏳ Đang tự động tiến hành kích hoạt bản quyền Ohook...")
                    self.tien_trinh_gist_ngam("/Ohook")
                
                self.cap_nhat_trang_thai("✅ HOÀN TẤT: Cài đặt và thiết lập Office thành công!")
                messagebox.showinfo("Hoàn tất", "Đã hoàn tất quá trình cài đặt Office và các tùy chọn bổ sung thành công.")
            else:
                self.cap_nhat_trang_thai("❌ Lỗi: Thiếu công cụ setup.exe!")
        except Exception as e:
            self.cap_nhat_trang_thai("❌ Đã xảy ra lỗi trong lúc cài đặt.")
        finally:
            self.dieu_khien_thanh_tien_do(False)

    # ---------- LOGIC GỠ CÀI ĐẶT ----------
    def khoi_dong_luong_go_office(self):
        if messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn gỡ toàn bộ Office khỏi máy tính không?"):
            threading.Thread(target=self.tien_trinh_go_office_ngam, daemon=True).start()

    def tien_trinh_go_office_ngam(self):
        self.dieu_khien_thanh_tien_do(True)
        try:
            self.cap_nhat_trang_thai("⏳ Đang thiết lập cấu hình gỡ cài đặt...")
            noi_dung_xml = """<Configuration>\n  <Remove All="True" />\n  <Display Level="Full" AcceptEULA="TRUE" />\n</Configuration>"""
            duong_dan_xml = os.path.join(os.getcwd(), "CauHinhGoCaiDat.xml")
            with open(duong_dan_xml, "w", encoding="utf-8") as f: f.write(noi_dung_xml)
            
            duong_dan_setup = self.chuan_bi_cong_cu_odt()
            if duong_dan_setup and os.path.exists(duong_dan_setup):
                self.cap_nhat_trang_thai("🚀 Đang chạy trình gỡ cài đặt của Microsoft...")
                tien_trinh = subprocess.Popen([duong_dan_setup, "/configure", duong_dan_xml])
                tien_trinh.wait()
                self.cap_nhat_trang_thai("✅ Đã gọi lệnh gỡ cài đặt xong!")
        finally:
            self.dieu_khien_thanh_tien_do(False)

    # ---------- LOGIC GỠ CRACK KMS ----------
    def tien_trinh_go_kms_ngam(self):
        if not messagebox.askyesno("Xác nhận", "Bạn có muốn xóa thông tin máy chủ KMS và Reset trạng thái bản quyền Office không?"): return
        self.dieu_khien_thanh_tien_do(True)
        try:
            self.cap_nhat_trang_thai("⏳ Đang tìm kiếm và dọn dẹp hệ thống KMS...")
            cac_thu_muc = [
                os.environ.get("ProgramFiles", "C:\\Program Files") + "\\Microsoft Office\\Office16",
                os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)") + "\\Microsoft Office\\Office16"
            ]
            file_ospp = next((os.path.join(tm, "ospp.vbs") for tm in cac_thu_muc if os.path.exists(os.path.join(tm, "ospp.vbs"))), None)
            if file_ospp:
                subprocess.run(["cscript", "//nologo", file_ospp, "/remhst"], creationflags=subprocess.CREATE_NO_WINDOW)
                subprocess.run(["cscript", "//nologo", file_ospp, "/rearm"], creationflags=subprocess.CREATE_NO_WINDOW)
                self.cap_nhat_trang_thai("✅ Đã dọn dẹp sạch bản quyền KMS cũ!")
                messagebox.showinfo("Thành công", "Đã xóa KMS ảo và Reset trạng thái bản quyền.\nVui lòng khởi động lại máy tính.")
            else:
                self.cap_nhat_trang_thai("⚠️ Không tìm thấy file hệ thống Office. Có thể Office đã bị gỡ.")
        finally:
            self.dieu_khien_thanh_tien_do(False)

    # ---------- LOGIC XỬ LÝ OHOOK (GIST) ----------
    def tien_trinh_gist_ngam(self, tham_so):
        url_gist = f"https://gist.githubusercontent.com/tuantran19912512/81329d670436ea8492b73bd5889ad444/raw/Ohook.cmd?t={time.time()}"
        file_tam = os.path.join(os.environ['TEMP'], "Ohook_Script.cmd")
        try:
            noi_dung = urllib.request.urlopen(url_gist).read().decode('utf-8')
            noi_dung = noi_dung.replace("\r\n", "\n").replace("\n", "\r\n") + "\r\n\r\n"
            with open(file_tam, 'w', encoding='utf-8') as f: f.write(noi_dung)
            tien_trinh = subprocess.Popen(["cmd.exe", "/c", file_tam, tham_so], creationflags=subprocess.CREATE_NO_WINDOW)
            tien_trinh.wait()
        except Exception:
            pass
        finally:
            if os.path.exists(file_tam):
                try: os.remove(file_tam)
                except: pass

    def tai_va_chay_gist_giao_dien(self, tham_so, loi_nhan_thanh_cong):
        self.dieu_khien_thanh_tien_do(True)
        try:
            self.tien_trinh_gist_ngam(tham_so)
            self.cap_nhat_trang_thai(loi_nhan_thanh_cong)
            messagebox.showinfo("Thành công", loi_nhan_thanh_cong[2:])
        except Exception:
            self.cap_nhat_trang_thai("❌ Lỗi kết nối mạng hoặc lỗi thực thi Gist.")
        finally:
            self.dieu_khien_thanh_tien_do(False)

    def khoi_dong_thuoc_ohook(self):
        self.cap_nhat_trang_thai("⏳ Đang tải và kích hoạt Ohook Silent...")
        threading.Thread(target=self.tai_va_chay_gist_giao_dien, args=("/Ohook", "✅ Đã KÍCH HOẠT thành công bản quyền Ohook!"), daemon=True).start()

    def khoi_dong_go_ohook(self):
        if not messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn gỡ bỏ hoàn toàn Crack Ohook không?"): return
        self.cap_nhat_trang_thai("⏳ Đang tải cấu hình gỡ Ohook...")
        threading.Thread(target=self.tai_va_chay_gist_giao_dien, args=("/OhookUninstall", "✅ Đã GỠ BỎ Ohook và khôi phục file gốc thành công!"), daemon=True).start()

    def cap_nhat_trang_thai(self, noi_dung_thong_bao):
        self.cua_so.after(0, lambda: self.nhan_trang_thai.config(text=noi_dung_thong_bao))

# ==============================================================================
# KHỞI CHẠY CHƯƠNG TRÌNH
# ==============================================================================
if __name__ == "__main__":
    cua_so_giao_dien = tk.Tk()
    ung_dung = UngDungCaiDatOffice(cua_so_giao_dien)
    cua_so_giao_dien.mainloop()