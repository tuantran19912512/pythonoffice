import os
import re
import subprocess
import urllib.request
import threading
import time
import queue
import socket
import glob
import shutil
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

socket.setdefaulttimeout(10)

# ==============================================================================
# CẤU HÌNH DỮ LIỆU MICROSOFT
# ==============================================================================
TUDIEN_PHIENBAN = {
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

TUDIEN_UNGDUNG = {
    "Access": "Access", "Excel": "Excel", "Word": "Word",
    "PowerPoint": "PowerPoint", "Outlook": "Outlook", "Publisher": "Publisher",
    "OneNote": "OneNote", "Skype": "Lync", "Teams": "Teams"
}

# ==============================================================================
# LÕI 1: ĐỘNG CƠ TẢI ĐA LUỒNG GHI TRỰC TIẾP (KIẾN TRÚC IDM)
# ==============================================================================
class DongCoTaiXuong:
    def __init__(self, duong_dan_mang, duong_dan_luu, so_luong_luong=16):
        self.duong_dan_mang = duong_dan_mang
        self.duong_dan_luu = duong_dan_luu
        self.so_luong_luong = so_luong_luong
        self.tong_dung_luong = 0
        self.dung_luong_da_tai = 0
        self.trang_thai_loi = False
        self.trang_thai_huy = False
        self.tieu_de_mang = {'User-Agent': 'Mozilla/5.0'}
        self.khoa_luong = threading.Lock()

    def kiem_tra_dung_luong(self):
        try:
            yeu_cau = urllib.request.Request(self.duong_dan_mang, method='HEAD', headers=self.tieu_de_mang)
            with urllib.request.urlopen(yeu_cau, timeout=10) as phan_hoi:
                self.tong_dung_luong = int(phan_hoi.headers.get('Content-Length', 0))
            return self.tong_dung_luong > 0
        except Exception:
            return False

    def luong_thuc_thi_tai(self, hang_doi_viec):
        while not hang_doi_viec.empty() and not self.trang_thai_loi and not self.trang_thai_huy:
            try:
                diem_bat_dau, diem_ket_thuc, chi_so = hang_doi_viec.get_nowait()
            except queue.Empty:
                break
                
            file_danh_dau = f"{self.duong_dan_luu}.done{chi_so}"
            if os.path.exists(file_danh_dau):
                hang_doi_viec.task_done()
                continue

            thanh_cong = False
            for _ in range(5): 
                if self.trang_thai_huy: break
                luong_da_tai_cuc_nay = 0
                try:
                    yeu_cau = urllib.request.Request(self.duong_dan_mang, headers={'Range': f'bytes={diem_bat_dau}-{diem_ket_thuc}', **self.tieu_de_mang})
                    with urllib.request.urlopen(yeu_cau, timeout=10) as phan_hoi:
                        with open(self.duong_dan_luu, 'r+b') as file_dich:
                            file_dich.seek(diem_bat_dau)
                            while True:
                                if self.trang_thai_huy: break
                                khoi_du_lieu = phan_hoi.read(262144) 
                                if not khoi_du_lieu: break
                                file_dich.write(khoi_du_lieu)
                                luong_da_tai_cuc_nay += len(khoi_du_lieu)
                                with self.khoa_luong:
                                    self.dung_luong_da_tai += len(khoi_du_lieu)
                    
                    if luong_da_tai_cuc_nay == (diem_ket_thuc - diem_bat_dau + 1):
                        with open(file_danh_dau, 'w') as f: f.write('OK')
                        thanh_cong = True
                        break
                    else:
                        with self.khoa_luong: self.dung_luong_da_tai -= luong_da_tai_cuc_nay
                except Exception:
                    with self.khoa_luong: self.dung_luong_da_tai -= luong_da_tai_cuc_nay
                    time.sleep(1)
                    
            if not thanh_cong and not self.trang_thai_huy:
                self.trang_thai_loi = True
            hang_doi_viec.task_done()

    def khoi_chay_dong_co(self):
        for rac in glob.glob(f"{self.duong_dan_luu}.phan*"):
            try: os.remove(rac)
            except: pass

        hang_doi_viec = queue.Queue()
        kich_thuoc_cuc = 50 * 1024 * 1024
        so_luong_cuc = self.tong_dung_luong // kich_thuoc_cuc
        if self.tong_dung_luong % kich_thuoc_cuc != 0:
            so_luong_cuc += 1

        if not os.path.exists(self.duong_dan_luu) or os.path.getsize(self.duong_dan_luu) != self.tong_dung_luong:
            with open(self.duong_dan_luu, "wb") as f:
                f.truncate(self.tong_dung_luong)
            for f in glob.glob(f"{self.duong_dan_luu}.done*"):
                os.remove(f)

        self.dung_luong_da_tai = 0
        
        for i in range(so_luong_cuc):
            diem_bat_dau = i * kich_thuoc_cuc
            diem_ket_thuc = min(diem_bat_dau + kich_thuoc_cuc - 1, self.tong_dung_luong - 1)
            
            if os.path.exists(f"{self.duong_dan_luu}.done{i}"):
                self.dung_luong_da_tai += (diem_ket_thuc - diem_bat_dau + 1)
            else:
                hang_doi_viec.put((diem_bat_dau, diem_ket_thuc, i))

        danh_sach_luong = []
        for _ in range(self.so_luong_luong):
            luong_moi = threading.Thread(target=self.luong_thuc_thi_tai, args=(hang_doi_viec,))
            luong_moi.start()
            danh_sach_luong.append(luong_moi)

        for luong in danh_sach_luong:
            luong.join()

        if self.trang_thai_huy or self.trang_thai_loi:
            return False

        self.don_dep_file_rac(so_luong_cuc)
        return True

    def don_dep_file_rac(self, so_luong_cuc):
        for i in range(so_luong_cuc):
            file_rac = f"{self.duong_dan_luu}.done{i}"
            if os.path.exists(file_rac):
                try: os.remove(file_rac)
                except: pass

# ==============================================================================
# LÕI 2: TIỆN ÍCH XỬ LÝ HỆ THỐNG VÀ REGISTRY (VIETTOOLBOX CORE)
# ==============================================================================
class TienIchHeThong:
    @staticmethod
    def tao_loi_tat_man_hinh(ham_cap_nhat_trang_thai):
        ham_cap_nhat_trang_thai("⏳ Đang tìm và đưa Shortcut ra màn hình Desktop...")
        thu_muc_man_hinh = os.path.join(os.environ['USERPROFILE'], 'Desktop')
        danh_sach_thu_muc = [
            os.environ.get('ProgramFiles', 'C:\\Program Files') + "\\Microsoft Office\\root\\Office16",
            os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)') + "\\Microsoft Office\\root\\Office16"
        ]
        # Thêm Project (WINPROJ.EXE) và Visio (VISIO.EXE) vào danh sách làm Shortcut
        danh_sach_ung_dung = {"WINWORD.EXE": "Word", "EXCEL.EXE": "Excel", "POWERPNT.EXE": "PowerPoint", "MSACCESS.EXE": "Access", "OUTLOOK.EXE": "Outlook", "WINPROJ.EXE": "Project", "VISIO.EXE": "Visio"}
        
        for thu_muc in danh_sach_thu_muc:
            if os.path.exists(thu_muc):
                for ten_file_exe, ten_hien_thi in danh_sach_ung_dung.items():
                    duong_dan_exe = os.path.join(thu_muc, ten_file_exe)
                    if os.path.exists(duong_dan_exe):
                        duong_dan_shortcut = os.path.join(thu_muc_man_hinh, f"{ten_hien_thi}.lnk")
                        ma_lenh_vbs = f'Set ws = CreateObject("WScript.Shell")\nSet link = ws.CreateShortcut("{duong_dan_shortcut}")\nlink.TargetPath = "{duong_dan_exe}"\nlink.Save'
                        file_vbs = os.path.join(os.environ['TEMP'], 'tao_shortcut.vbs')
                        with open(file_vbs, 'w', encoding='utf-8') as f:
                            f.write(ma_lenh_vbs)
                        subprocess.run(['cscript', '//nologo', file_vbs], creationflags=subprocess.CREATE_NO_WINDOW)
                break

    @staticmethod
    def doi_tien_trinh_ket_thuc(ten_tien_trinh, ham_cap_nhat_trang_thai):
        ham_cap_nhat_trang_thai(f"⏳ Đang theo dõi tiến trình hệ thống ({ten_tien_trinh})...")
        while True:
            ket_qua_kiem_tra = subprocess.run(f'tasklist /FI "IMAGENAME eq {ten_tien_trinh}"', shell=True, capture_output=True, text=True)
            if ten_tien_trinh.lower() not in ket_qua_kiem_tra.stdout.lower():
                break
            time.sleep(2)

    @staticmethod
    def don_dep_registry_chuyen_sau(ham_cap_nhat_trang_thai):
        ham_cap_nhat_trang_thai("🧹 Đang kích hoạt Deep Clean: Tiêu diệt tiến trình ngầm...")
        os.system('taskkill /F /IM ClickToRunSvc.exe >nul 2>&1')
        os.system('taskkill /F /IM OfficeClickToRun.exe >nul 2>&1')
        os.system('taskkill /F /IM WINWORD.EXE >nul 2>&1')
        os.system('taskkill /F /IM EXCEL.EXE >nul 2>&1')
        
        ham_cap_nhat_trang_thai("🧹 Đang Deep Clean: Quét sạch Registry rác...")
        danh_sach_khoa = [
            r"HKCU\Software\Microsoft\Office",
            r"HKLM\SOFTWARE\Microsoft\Office",
            r"HKLM\SOFTWARE\WOW6432Node\Microsoft\Office",
            r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail - vi-vn"
        ]
        for khoa_reg in danh_sach_khoa:
            os.system(f'reg delete "{khoa_reg}" /f >nul 2>&1')
            
        ham_cap_nhat_trang_thai("🧹 Đang Deep Clean: Xóa tàn dư trong ổ C:...")
        thu_muc_rac = [
            os.path.join(os.environ.get('ProgramFiles', 'C:\\Program Files'), "Microsoft Office"),
            os.path.join(os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)'), "Microsoft Office"),
            os.path.join(os.environ.get('ProgramData', 'C:\\ProgramData'), "Microsoft\\Office")
        ]
        for thu_muc in thu_muc_rac:
            if os.path.exists(thu_muc):
                try: shutil.rmtree(thu_muc, ignore_errors=True)
                except: pass

    @staticmethod
    def kich_hoat_ohook_ngam(tham_so):
        duong_dan_mang = f"https://gist.githubusercontent.com/tuantran19912512/81329d670436ea8492b73bd5889ad444/raw/Ohook.cmd?t={time.time()}"
        file_tam = os.path.join(os.environ['TEMP'], "O.cmd")
        try:
            with open(file_tam, 'w', encoding='utf-8') as f:
                f.write(urllib.request.urlopen(duong_dan_mang).read().decode('utf-8').replace("\n", "\r\n"))
            subprocess.run(["cmd.exe", "/c", file_tam, tham_so], creationflags=subprocess.CREATE_NO_WINDOW)
        except: pass
        finally:
            if os.path.exists(file_tam):
                try: os.remove(file_tam)
                except: pass

# ==============================================================================
# LÕI 3: GIAO DIỆN CHÍNH (VIETTOOLBOX GUI)
# ==============================================================================
class TrienKhaiOffice(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("VietToolbox - Triển khai Microsoft Office (V10.6 All-In-One)")
        self.geometry("640x760")
        self.resizable(False, False)
        self.phong_chu_dam = ("Segoe UI", 9, "bold")
        self.thu_muc_lam_viec = tk.StringVar(value=os.getcwd())
        self.tien_trinh_tai_mang = None
        self.xay_dung_bo_khung_giao_dien()

    def tao_nut_bam(self, khung_chua, chu_hien_thi, mau_nen, hanh_dong=None, do_rong=20):
        return tk.Button(khung_chua, text=chu_hien_thi, bg=mau_nen, fg="white", font=self.phong_chu_dam, relief="flat", cursor="hand2", command=hanh_dong, width=do_rong, pady=6)

    def cap_nhat_trang_thai(self, tin_nhan):
        self.after(0, lambda: self.nhan_trang_thai.config(text=tin_nhan))

    def xay_dung_bo_khung_giao_dien(self):
        khung_tieu_de = tk.Frame(self, bg="#1565C0", height=65)
        khung_tieu_de.pack(fill="x", side="top")
        tk.Label(khung_tieu_de, text="VIETTOOLBOX - BỘ CÀI ĐẶT MICROSOFT OFFICE", font=("Segoe UI", 13, "bold"), bg="#1565C0", fg="white").place(x=15, y=10)
        tk.Label(khung_tieu_de, text="Bản quyền công nghệ thuộc về hệ sinh thái VietToolbox", font=("Segoe UI", 8), bg="#1565C0", fg="#BBDEFB").place(x=18, y=35)

        hop_dieu_huong = ttk.Notebook(self)
        hop_dieu_huong.pack(fill="both", expand=True, padx=12, pady=12)

        tab_cai_dat = ttk.Frame(hop_dieu_huong)
        tab_go_cai_dat = ttk.Frame(hop_dieu_huong)
        tab_toi_uu = ttk.Frame(hop_dieu_huong)
        
        hop_dieu_huong.add(tab_cai_dat, text="  ⚙️ Triển Khai  ")
        hop_dieu_huong.add(tab_go_cai_dat, text="  🗑️ Gỡ Cài Đặt  ")
        hop_dieu_huong.add(tab_toi_uu, text="  🚀 Tối Ưu Hóa & Fix Lỗi  ")

        self.thiet_ke_tab_cai_dat(tab_cai_dat)
        self.thiet_ke_tab_go_cai_dat(tab_go_cai_dat)
        self.thiet_ke_tab_toi_uu(tab_toi_uu)

        khung_trang_thai = tk.Frame(self, bg="#E3F2FD")
        khung_trang_thai.pack(fill="x", padx=12, pady=(0, 12))
        self.nhan_trang_thai = tk.Label(khung_trang_thai, text="✅ Hệ thống VietToolbox sẵn sàng nhận lệnh...", font=self.phong_chu_dam, fg="#1565C0", bg="#E3F2FD")
        self.nhan_trang_thai.pack(anchor="w", padx=10, pady=(10, 5))
        self.thanh_tien_do = ttk.Progressbar(khung_trang_thai, mode='determinate')
        self.thanh_tien_do.pack(fill="x", padx=12, pady=(0, 10))

    def thiet_ke_tab_cai_dat(self, tab):
        khung_phien_ban = ttk.LabelFrame(tab, text=" 📄 Chọn Phiên Bản Office ")
        khung_phien_ban.pack(fill="x", padx=10, pady=8, ipady=3)
        self.bien_nam_phien_ban = tk.StringVar(value="2024")
        danh_sach_nam = ["Office 2016", "Office 2019", "Office 2021", "Office 2024", "Office 365"]
        khung_chon_nam = ttk.Frame(khung_phien_ban)
        khung_chon_nam.pack(fill="x", padx=10, pady=2)
        self.bien_nam_phien_ban.trace_add("write", self.cap_nhat_danh_sach_ban_con)
        for vi_tri, ten_nam in enumerate(danh_sach_nam):
            ttk.Radiobutton(khung_chon_nam, text=ten_nam, variable=self.bien_nam_phien_ban, value=ten_nam.split()[-1]).grid(row=0, column=vi_tri, padx=6)
            
        khung_chon_chi_tiet = ttk.Frame(khung_phien_ban)
        khung_chon_chi_tiet.pack(fill="x", padx=10, pady=5)
        ttk.Label(khung_chon_chi_tiet, text="Loại giấy phép:").pack(side="left", padx=(0, 10))
        self.danh_sach_tha_xuong = ttk.Combobox(khung_chon_chi_tiet, state="readonly", width=35)
        self.cap_nhat_danh_sach_ban_con()
        self.danh_sach_tha_xuong.pack(side="left")

        khung_ngang_thong_so = ttk.Frame(tab)
        khung_ngang_thong_so.pack(fill="x", padx=10, pady=5)
        khung_kien_truc = ttk.LabelFrame(khung_ngang_thong_so, text=" ⚙️ Kiến trúc HĐH ")
        khung_kien_truc.pack(side="left", fill="both", expand=True, padx=(0, 5))
        self.bien_nen_tang = tk.StringVar(value="64")
        ttk.Radiobutton(khung_kien_truc, text="64-bit", variable=self.bien_nen_tang, value="64").pack(side="left", padx=15, pady=5)
        ttk.Radiobutton(khung_kien_truc, text="32-bit", variable=self.bien_nen_tang, value="32").pack(side="right", padx=15, pady=5)

        khung_ngon_ngu = ttk.LabelFrame(khung_ngang_thong_so, text=" 🌐 Ngôn ngữ ")
        khung_ngon_ngu.pack(side="right", fill="both", expand=True, padx=(5, 0))
        self.hop_chon_ngon_ngu = ttk.Combobox(khung_ngon_ngu, values=["English (US) - en-US", "Vietnamese - vi-VN"], state="readonly")
        self.hop_chon_ngon_ngu.current(0)
        self.hop_chon_ngon_ngu.pack(fill="x", padx=10, pady=6)

        khung_ung_dung = ttk.LabelFrame(tab, text=" ☑️ Thành phần cài đặt ")
        khung_ung_dung.pack(fill="x", padx=10, pady=5)
        mang_ung_dung = [("Access",0,0),("Excel",0,1),("Word",0,2),("PowerPoint",1,0),("Outlook",1,1),("Publisher",1,2),("OneNote",2,0),("Skype",2,1),("Teams",2,2)]
        self.tu_dien_tich_chon = {}
        for ten_ung_dung, hang, cot in mang_ung_dung:
            bien_tich = tk.BooleanVar(value=True)
            self.tu_dien_tich_chon[ten_ung_dung] = bien_tich
            ttk.Checkbutton(khung_ung_dung, text=f" {ten_ung_dung}", variable=bien_tich).grid(row=hang, column=cot, sticky="w", padx=25, pady=4)

        # TÍNH NĂNG MỚI: SẢN PHẨM BỔ SUNG (PROJECT & VISIO)
        khung_sp_phu = ttk.LabelFrame(tab, text=" ➕ Sản phẩm bổ sung (Pro) ")
        khung_sp_phu.pack(fill="x", padx=10, pady=5)
        self.bien_project = tk.BooleanVar(value=False)
        self.bien_visio = tk.BooleanVar(value=False)
        ttk.Checkbutton(khung_sp_phu, text=" Microsoft Project Pro", variable=self.bien_project).pack(side="left", padx=25, pady=4)
        ttk.Checkbutton(khung_sp_phu, text=" Microsoft Visio Pro", variable=self.bien_visio).pack(side="left", padx=25, pady=4)

        khung_luu_tru = ttk.LabelFrame(tab, text=" 📂 Thư mục làm việc & Lưu trữ lõi C2R ")
        khung_luu_tru.pack(fill="x", padx=10, pady=5)
        khung_con_luu_tru = ttk.Frame(khung_luu_tru)
        khung_con_luu_tru.pack(fill="x", padx=10, pady=8)
        ttk.Entry(khung_con_luu_tru, textvariable=self.thu_muc_lam_viec, state="readonly").pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.tao_nut_bam(khung_con_luu_tru, "Đổi Thư Mục", "#757575", hanh_dong=self.hanh_dong_chon_thu_muc, do_rong=12).pack(side="right")

        khung_tuy_chon_phu = ttk.LabelFrame(tab, text=" 🔧 Hành động tự động sau khi cài ")
        khung_tuy_chon_phu.pack(fill="x", padx=10, pady=5)
        self.bien_tu_dong_crack = tk.BooleanVar(value=True)
        ttk.Checkbutton(khung_tuy_chon_phu, text="Tự động nhúng thuốc Ohook", variable=self.bien_tu_dong_crack).pack(side="left", padx=25, pady=6)
        self.bien_tu_dong_shortcut = tk.BooleanVar(value=True)
        ttk.Checkbutton(khung_tuy_chon_phu, text="Gắn lối tắt ra Desktop", variable=self.bien_tu_dong_shortcut).pack(side="left", padx=25, pady=6)

        khung_dieu_khien = ttk.Frame(tab)
        khung_dieu_khien.pack(fill="x", padx=10, pady=5)
        self.tao_nut_bam(khung_dieu_khien, "💊 BƠM THUỐC OHOOK", "#F57C00", hanh_dong=self.hanh_dong_bom_thuoc).pack(side="left", padx=5)
        self.nut_bat_dau_cai = self.tao_nut_bam(khung_dieu_khien, "🚀 BẮT ĐẦU TRIỂN KHAI", "#1565C0", hanh_dong=self.hanh_dong_cai_dat)
        self.nut_bat_dau_cai.pack(side="right", padx=5)

    def thiet_ke_tab_go_cai_dat(self, tab):
        khung_go_office = ttk.LabelFrame(tab, text=" 🗑️ VietToolbox Deep Clean (Gỡ Cài Đặt + Xóa Sạch Registry) ")
        khung_go_office.pack(fill="x", padx=10, pady=7)
        ttk.Label(khung_go_office, text="Chức năng này gọi bộ gỡ chuẩn của Microsoft, sau đó sử dụng thuật toán\ncủa VietToolbox để rà soát và tiêu diệt mọi file rác, khóa Registry tồn đọng.", justify="left").pack(anchor="w", padx=15, pady=3)
        khung_nut_go = ttk.Frame(khung_go_office)
        khung_nut_go.pack(fill="x", padx=10, pady=5)
        self.tao_nut_bam(khung_nut_go, "🗑 BẮT ĐẦU GỠ & CLEAN", "#D32F2F", hanh_dong=self.hanh_dong_go_office).pack(side="right", padx=5, pady=2)

        khung_go_kms = ttk.LabelFrame(tab, text=" 🧹 Dọn dẹp máy chủ KMS ảo ")
        khung_go_kms.pack(fill="x", padx=10, pady=7)
        ttk.Label(khung_go_kms, text="Bóc tách và xóa bỏ mọi cấu hình kích hoạt lậu KMS cũ, đưa trạng thái\nbản quyền Office về nguyên bản (Rearm).", justify="left").pack(anchor="w", padx=15, pady=3)
        khung_nut_kms = ttk.Frame(khung_go_kms)
        khung_nut_kms.pack(fill="x", padx=10, pady=5)
        self.tao_nut_bam(khung_nut_kms, "🧹 DỌN SẠCH KMS", "#1976D2", hanh_dong=self.hanh_dong_go_kms).pack(side="right", padx=5, pady=2)

        khung_go_ohook = ttk.LabelFrame(tab, text=" 💊 Gỡ Crack Ohook ")
        khung_go_ohook.pack(fill="x", padx=10, pady=7)
        ttk.Label(khung_go_ohook, text="Xóa sạch mã nhúng Ohook và hoàn nguyên cấu trúc thư mục bảo mật.", justify="left").pack(anchor="w", padx=15, pady=3)
        khung_nut_ohook = ttk.Frame(khung_go_ohook)
        khung_nut_ohook.pack(fill="x", padx=10, pady=5)
        self.tao_nut_bam(khung_nut_ohook, "🛡️ RÚT THUỐC OHOOK", "#7B1FA2", hanh_dong=self.hanh_dong_rut_thuoc_ohook).pack(side="right", padx=5, pady=2)

        khung_cuu_ho = ttk.LabelFrame(tab, text=" 🆘 CỨU HỘ KHẨN CẤP: LỖI KẸT BỘ CÀI (ZOMBIE) ")
        khung_cuu_ho.pack(fill="x", padx=10, pady=7)
        ttk.Label(khung_cuu_ho, text="Chỉ dùng khi bị mất Registry không thể gỡ qua Control Panel, cài mới thì báo lỗi.\nTính năng này sẽ xóa ép buộc dịch vụ ClickToRunSvc lõi.", justify="left").pack(anchor="w", padx=15, pady=3)
        khung_nut_cuu_ho = ttk.Frame(khung_cuu_ho)
        khung_nut_cuu_ho.pack(fill="x", padx=10, pady=5)
        self.tao_nut_bam(khung_nut_cuu_ho, "🆘 XÓA ÉP BUỘC C2R", "#E65100", hanh_dong=self.hanh_dong_cuu_ho_zombie).pack(side="right", padx=5, pady=2)

    def thiet_ke_tab_toi_uu(self, tab):
        khung_nghi_dinh = ttk.LabelFrame(tab, text=" 📜 Cấu hình chuẩn nhà nước (Nghị định 30) ")
        khung_nghi_dinh.pack(fill="x", padx=10, pady=7)
        ttk.Label(khung_nghi_dinh, text="Tự động can thiệp vào Normal.dotm để cài đặt mặc định cho Word:\n- Font: Times New Roman, Size 14.\n- Căn lề: Trên 2cm, Dưới 2cm, Trái 3cm, Phải 2cm.", justify="left").pack(anchor="w", padx=15, pady=5)
        self.bien_nghi_dinh = tk.BooleanVar(value=True)
        ttk.Checkbutton(khung_nghi_dinh, text="Bật cấu hình chuẩn Nghị định 30/2020", variable=self.bien_nghi_dinh).pack(anchor="w", padx=25, pady=(0,5))

        khung_fix_loi = ttk.LabelFrame(tab, text=" 🔧 Fix lỗi & Tăng tốc độ mở file ")
        khung_fix_loi.pack(fill="x", padx=10, pady=7)
        
        self.bien_tat_welcome = tk.BooleanVar(value=True)
        ttk.Checkbutton(khung_fix_loi, text="Tắt màn hình Welcome (Vào thẳng trang trắng khi mở Word/Excel)", variable=self.bien_tat_welcome).pack(anchor="w", padx=25, pady=4)
        
        self.bien_tat_protect = tk.BooleanVar(value=True)
        ttk.Checkbutton(khung_fix_loi, text="Tắt Protected View (Mở file Zalo/Internet KHÔNG bị thanh vàng báo lỗi)", variable=self.bien_tat_protect).pack(anchor="w", padx=25, pady=4)
        
        self.bien_tat_hw_accel = tk.BooleanVar(value=True)
        ttk.Checkbutton(khung_fix_loi, text="Tắt Hardware Acceleration (Chống đen màn hình, giật lag trên máy cũ)", variable=self.bien_tat_hw_accel).pack(anchor="w", padx=25, pady=4)

        self.bien_bat_autosave = tk.BooleanVar(value=True)
        ttk.Checkbutton(khung_fix_loi, text="Bật AutoSave siêu tốc (Tự động lưu file mỗi 3 phút chống mất điện)", variable=self.bien_bat_autosave).pack(anchor="w", padx=25, pady=4)

        khung_nut_toi_uu = ttk.Frame(tab)
        khung_nut_toi_uu.pack(fill="x", padx=10, pady=15)
        self.tao_nut_bam(khung_nut_toi_uu, "🚀 THỰC THI TỐI ƯU OFFICE", "#00838F", hanh_dong=self.hanh_dong_thuc_thi_toi_uu).pack(side="right", padx=5)

    def cap_nhat_danh_sach_ban_con(self, *args):
        if self.bien_nam_phien_ban.get() == "365":
            self.danh_sach_tha_xuong['values'] = ["ProPlus", "Business", "Home Premium"]
        else:
            self.danh_sach_tha_xuong['values'] = ["ProPlus", "Standard", "Home & Business", "Home & Student"]
        self.danh_sach_tha_xuong.current(0)

    def hanh_dong_chon_thu_muc(self):
        duong_dan_moi = filedialog.askdirectory(initialdir=self.thu_muc_lam_viec.get(), title="Chọn thư mục tải bản cài")
        if duong_dan_moi:
            if os.path.exists(duong_dan_moi) and os.access(duong_dan_moi, os.W_OK):
                self.thu_muc_lam_viec.set(duong_dan_moi)
            else:
                messagebox.showerror("Lỗi Truy Cập", "VietToolbox không có quyền ghi vào thư mục này.")

    def chuan_bi_setup_exe(self, thu_muc_luu_tru=None):
        duong_dan_setup = os.path.join(thu_muc_luu_tru or os.environ['TEMP'], "setup.exe")
            
        if os.path.exists(duong_dan_setup) and os.path.getsize(duong_dan_setup) > 2000000:
            self.cap_nhat_trang_thai("♻️ Đã tìm thấy lõi Setup.exe, VietToolbox đang tái sử dụng...")
            return duong_dan_setup

        self.cap_nhat_trang_thai("⏳ VietToolbox đang móc nối lấy lõi Setup.exe trực tiếp từ MS WSUS...")
        link_tai_truc_tiep = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
        try:
            yeu_cau = urllib.request.Request(link_tai_truc_tiep, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(yeu_cau, timeout=30) as phan_hoi, open(duong_dan_setup, 'wb') as file_luu:
                file_luu.write(phan_hoi.read())
            if os.path.exists(duong_dan_setup) and os.path.getsize(duong_dan_setup) > 1000000:
                return duong_dan_setup
            return None
        except:
            return None

    def hanh_dong_huy_tai(self):
        if messagebox.askyesno("Xác nhận hủy", "Hủy bỏ toàn bộ quá trình triển khai?"):
            if self.tien_trinh_tai_mang:
                self.tien_trinh_tai_mang.trang_thai_huy = True
            self.cap_nhat_trang_thai("⏳ Đang dọn dẹp các luồng tải dang dở...")

    def khoi_phuc_nut_cai_dat(self):
        self.after(0, lambda: self.nut_bat_dau_cai.config(text="🚀 BẮT ĐẦU TRIỂN KHAI", bg="#1565C0", command=self.hanh_dong_cai_dat))

    def hanh_dong_cai_dat(self):
        self.nut_bat_dau_cai.config(text="🛑 HỦY TRIỂN KHAI", bg="#D32F2F", command=self.hanh_dong_huy_tai)
        threading.Thread(target=self.luong_xu_ly_cai_dat_chinh, daemon=True).start()

    def luong_xu_ly_cai_dat_chinh(self):
        thu_muc_goc = os.path.normpath(self.thu_muc_lam_viec.get())
        ma_nam = self.bien_nam_phien_ban.get()
        ma_loai = self.danh_sach_tha_xuong.get().replace(" & ", "").replace(" ", "")
        ma_san_pham = TUDIEN_PHIENBAN.get(f"{ma_nam}_{ma_loai}", "ProPlus2024Retail")
        nhi_phan_so = "64" if self.bien_nen_tang.get() == "64" else "32"
        nhi_phan_chu = "x64" if nhi_phan_so == "64" else "x86"
        ma_ngon_ngu = "vi-VN" if "Vietnamese" in self.hop_chon_ngon_ngu.get() else "en-US"
        ma_quoc_gia = "1066" if ma_ngon_ngu == "vi-VN" else "1033"
        
        self.cap_nhat_trang_thai("🔍 VietToolbox đang giải mã Cab để lấy phiên bản mới nhất...")
        duong_dan_goc = "https://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60"
        phien_ban_moi_nhat = None
        try:
            thu_muc_tam_cab = os.path.join(os.environ['TEMP'], "VTCabTemp")
            if os.path.exists(thu_muc_tam_cab):
                shutil.rmtree(thu_muc_tam_cab)
            os.makedirs(thu_muc_tam_cab, exist_ok=True)
            file_cab = os.path.join(thu_muc_tam_cab, f"v{nhi_phan_so}.cab")
            yeu_cau = urllib.request.Request(f"{duong_dan_goc}/Office/Data/v{nhi_phan_so}.cab", headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(yeu_cau, timeout=15) as phan_hoi, open(file_cab, 'wb') as f:
                f.write(phan_hoi.read())
            subprocess.run(["expand.exe", file_cab, "-F:*.xml", thu_muc_tam_cab], creationflags=subprocess.CREATE_NO_WINDOW)
            for file_xml in glob.glob(os.path.join(thu_muc_tam_cab, "*.xml")):
                noi_dung_xml = open(file_xml, 'r', encoding='utf-8', errors='ignore').read()
                ket_qua_tim = re.search(r'Version="(\d{2}\.\d+\.\d+\.\d+)"', noi_dung_xml)
                if ket_qua_tim:
                    phien_ban_moi_nhat = ket_qua_tim.group(1)
                    break
            if not phien_ban_moi_nhat:
                raise Exception("Lỗi đọc XML")
        except:
            self.cap_nhat_trang_thai("❌ Lỗi: VietToolbox không thể kết nối tới máy chủ Microsoft.")
            self.khoi_phuc_nut_cai_dat()
            return

        thu_muc_data = os.path.join(thu_muc_goc, "Office", "Data")
        thu_muc_version = os.path.join(thu_muc_data, phien_ban_moi_nhat)
        os.makedirs(thu_muc_version, exist_ok=True)

        danh_sach_link_tai = [
            (f"{duong_dan_goc}/Office/Data/v{nhi_phan_so}.cab", os.path.join(thu_muc_data, f"v{nhi_phan_so}.cab"), "Danh mục gốc"),
            (f"{duong_dan_goc}/Office/Data/v{nhi_phan_so}_{phien_ban_moi_nhat}.cab", os.path.join(thu_muc_data, f"v{nhi_phan_so}_{phien_ban_moi_nhat}.cab"), "Danh mục phiên bản"),
            (f"{duong_dan_goc}/Office/Data/{phien_ban_moi_nhat}/i{nhi_phan_so}0.cab", os.path.join(thu_muc_version, f"i{nhi_phan_so}0.cab"), "Mục lục lõi"),
            (f"{duong_dan_goc}/Office/Data/{phien_ban_moi_nhat}/s{nhi_phan_so}0.cab", os.path.join(thu_muc_version, f"s{nhi_phan_so}0.cab"), "Mục lục phụ"),
            (f"{duong_dan_goc}/Office/Data/{phien_ban_moi_nhat}/i{nhi_phan_so}{ma_quoc_gia}.cab", os.path.join(thu_muc_version, f"i{nhi_phan_so}{ma_quoc_gia}.cab"), "Mục lục ngôn ngữ 1"),
            (f"{duong_dan_goc}/Office/Data/{phien_ban_moi_nhat}/s{nhi_phan_so}{ma_quoc_gia}.cab", os.path.join(thu_muc_version, f"s{nhi_phan_so}{ma_quoc_gia}.cab"), "Mục lục ngôn ngữ 2"),
            (f"{duong_dan_goc}/Office/Data/{phien_ban_moi_nhat}/stream.{nhi_phan_chu}.x-none.dat", os.path.join(thu_muc_version, f"stream.{nhi_phan_chu}.x-none.dat"), "Dữ liệu cấu trúc lớn"),
            (f"{duong_dan_goc}/Office/Data/{phien_ban_moi_nhat}/stream.{nhi_phan_chu}.{ma_ngon_ngu}.dat", os.path.join(thu_muc_version, f"stream.{nhi_phan_chu}.{ma_ngon_ngu}.dat"), "Dữ liệu Ngôn ngữ")
        ]

        for link_mang, duong_dan_luu, ten_hien_thi in danh_sach_link_tai:
            if self.tien_trinh_tai_mang and self.tien_trinh_tai_mang.trang_thai_huy:
                break
            yeu_cau_kiem_tra = urllib.request.Request(link_mang, method='HEAD', headers={'User-Agent': 'Mozilla/5.0'})
            try:
                dung_luong_thuc = int(urllib.request.urlopen(yeu_cau_kiem_tra, timeout=10).headers.get('Content-Length', 0))
            except:
                dung_luong_thuc = 0
            if os.path.exists(duong_dan_luu) and os.path.getsize(duong_dan_luu) == dung_luong_thuc:
                continue

            self.tien_trinh_tai_mang = DongCoTaiXuong(link_mang, duong_dan_luu, 16)
            if self.tien_trinh_tai_mang.kiem_tra_dung_luong():
                luong_tai_chinh = threading.Thread(target=self.tien_trinh_tai_mang.khoi_chay_dong_co)
                luong_tai_chinh.start()
                thoi_gian_truoc = time.time()
                dung_luong_truoc = 0
                while luong_tai_chinh.is_alive():
                    time.sleep(0.5)
                    thoi_gian_troi_qua = time.time() - thoi_gian_truoc
                    if thoi_gian_troi_qua > 0:
                        toc_do_mang = ((self.tien_trinh_tai_mang.dung_luong_da_tai - dung_luong_truoc) / 0.5) / 1048576
                        phan_tram = (self.tien_trinh_tai_mang.dung_luong_da_tai / self.tien_trinh_tai_mang.tong_dung_luong) * 100
                        self.after(0, lambda p=phan_tram: self.thanh_tien_do.config(value=p))
                        self.cap_nhat_trang_thai(f"⬇️ {ten_hien_thi}: {phan_tram:.1f}% | Tốc độ: {toc_do_mang:.1f} MB/s")
                    dung_luong_truoc = self.tien_trinh_tai_mang.dung_luong_da_tai
                    thoi_gian_truoc = time.time()

                if self.tien_trinh_tai_mang.trang_thai_huy:
                    self.cap_nhat_trang_thai("❌ Đã dừng luồng tải thành công.")
                    self.after(0, lambda: self.thanh_tien_do.config(value=0))
                    self.khoi_phuc_nut_cai_dat()
                    return
                if self.tien_trinh_tai_mang.trang_thai_loi:
                    self.cap_nhat_trang_thai(f"❌ Mạng rớt khi kéo {ten_hien_thi}.")
                    self.khoi_phuc_nut_cai_dat()
                    return

        self.after(0, lambda: self.thanh_tien_do.config(mode='indeterminate', value=0))
        self.after(0, lambda: self.thanh_tien_do.start(15))
        self.cap_nhat_trang_thai("🚀 Lõi VietToolbox đang nạp cấu trúc vào hệ thống. Vui lòng đợi...")
        
        duong_dan_setup = self.chuan_bi_setup_exe(thu_muc_goc)
        if not duong_dan_setup:
            self.cap_nhat_trang_thai("❌ Lỗi: Không thể móc nối với lõi cài đặt của Microsoft!")
            self.khoi_phuc_nut_cai_dat()
            return

        danh_sach_chon = [ten for ten, bien in self.tu_dien_tich_chon.items() if bien.get()]
        ma_lenh_xml = f"""<Configuration>\n  <Add SourcePath="{thu_muc_goc}" OfficeClientEdition="{nhi_phan_so}" Channel="Current" Version="{phien_ban_moi_nhat}" AllowCdnFallback="True">\n    <Product ID="{ma_san_pham}">\n      <Language ID="{ma_ngon_ngu}" />\n"""
        
        for ten_ung_dung, id_ung_dung in TUDIEN_UNGDUNG.items():
            if ten_ung_dung not in danh_sach_chon:
                ma_lenh_xml += f'      <ExcludeApp ID="{id_ung_dung}" />\n'
        ma_lenh_xml += "    </Product>\n"

        # TỰ ĐỘNG GEN MÃ XML CHO PROJECT VÀ VISIO THEO NĂM BẢN QUYỀN
        if self.bien_project.get():
            id_project = f"ProjectPro{ma_nam}Retail" if ma_nam not in ["365", "2016"] else "ProjectProRetail"
            ma_lenh_xml += f'    <Product ID="{id_project}">\n      <Language ID="{ma_ngon_ngu}" />\n    </Product>\n'
            
        if self.bien_visio.get():
            id_visio = f"VisioPro{ma_nam}Retail" if ma_nam not in ["365", "2016"] else "VisioProRetail"
            ma_lenh_xml += f'    <Product ID="{id_visio}">\n      <Language ID="{ma_ngon_ngu}" />\n    </Product>\n'

        ma_lenh_xml += """  </Add>\n  <Updates Enabled="TRUE" />\n  <Display Level="Full" AcceptEULA="TRUE" />\n</Configuration>"""
        
        file_cau_hinh_xml = os.path.join(thu_muc_goc, "C2R_Config.xml")
        with open(file_cau_hinh_xml, "w", encoding="utf-8") as f:
            f.write(ma_lenh_xml)
        
        self.cap_nhat_trang_thai("🚀 Đang khởi động lõi cài đặt Microsoft C2R. Bác chờ bảng Office hiện lên nhé...")
        tien_trinh_cai_dat = subprocess.Popen([duong_dan_setup, "/configure", file_cau_hinh_xml], cwd=thu_muc_goc)
        tien_trinh_cai_dat.wait()
        
        TienIchHeThong.doi_tien_trinh_ket_thuc("setup.exe", self.cap_nhat_trang_thai)
        ma_tra_ve = tien_trinh_cai_dat.returncode
        
        if os.path.exists(file_cau_hinh_xml):
            os.remove(file_cau_hinh_xml)
        if os.path.exists(duong_dan_setup):
            os.remove(duong_dan_setup)
            
        self.after(0, lambda: self.thanh_tien_do.stop())
        
        if ma_tra_ve == 0:
            if messagebox.askyesno("Dọn dẹp", "Cài đặt thành công!\nBạn có muốn XÓA thư mục 2.9GB tải về để tiết kiệm ổ cứng không?"):
                try: shutil.rmtree(os.path.join(thu_muc_goc, "Office"))
                except: pass
            
            if self.bien_tu_dong_shortcut.get():
                TienIchHeThong.tao_loi_tat_man_hinh(self.cap_nhat_trang_thai)
            if self.bien_tu_dong_crack.get():
                self.cap_nhat_trang_thai("⏳ VietToolbox đang bơm thuốc Ohook...")
                TienIchHeThong.kich_hoat_ohook_ngam("/Ohook")
            self.cap_nhat_trang_thai("✅ HOÀN TẤT: Cấu trúc C2R đã được tích hợp vững chắc!")
            messagebox.showinfo("Hoàn Tất Chuyên Nghiệp", "Mọi quy trình đã hoàn tất theo chuẩn hệ sinh thái VietToolbox. Chúc bác một ngày làm việc năng suất!")
        else:
            self.cap_nhat_trang_thai(f"❌ Lỗi: Cài đặt thất bại (Mã thoát: {ma_tra_ve})")
            messagebox.showerror("Báo Cáo Lỗi", f"Tiến trình bị gián đoạn đột ngột!\nMã lỗi: {ma_tra_ve}\nNguyên nhân: Đường dẫn có thể chứa ký tự lạ, file tải dở bị lỗi, hoặc bác chưa gỡ bản cũ.")
        self.khoi_phuc_nut_cai_dat()

    def hanh_dong_go_office(self):
        if messagebox.askyesno("Cảnh Báo Chuyên Sâu", "Hành động này sẽ:\n1. Gỡ hoàn toàn bộ Office đang có.\n2. VietToolbox sẽ Deep Clean xóa sạch Registry và File tàn dư.\n\nBác có chắc chắn muốn làm phẳng hệ thống không?"): 
            threading.Thread(target=self.luong_xu_ly_go_office, daemon=True).start()

    def luong_xu_ly_go_office(self):
        self.after(0, lambda: self.thanh_tien_do.config(mode='indeterminate', value=0))
        self.after(0, lambda: self.thanh_tien_do.start(15))
        try:
            self.cap_nhat_trang_thai("⏳ VietToolbox đang dựng kịch bản nhổ rễ Office...")
            lenh_go_bo = """<Configuration>\n  <Remove All="True" />\n  <Display Level="Full" AcceptEULA="TRUE" />\n</Configuration>"""
            file_cau_hinh_go = os.path.join(os.environ['TEMP'], "VTCauHinhGo.xml")
            with open(file_cau_hinh_go, "w", encoding="utf-8") as f:
                f.write(lenh_go_bo)
            
            duong_dan_setup = self.chuan_bi_setup_exe()
            if duong_dan_setup:
                self.cap_nhat_trang_thai("🚀 Đang khởi động máy xúc Microsoft C2R. Vui lòng chờ bảng cam tắt...")
                tien_trinh_go = subprocess.Popen([duong_dan_setup, "/configure", file_cau_hinh_go])
                tien_trinh_go.wait()
                
                TienIchHeThong.doi_tien_trinh_ket_thuc("setup.exe", self.cap_nhat_trang_thai)
                TienIchHeThong.don_dep_registry_chuyen_sau(self.cap_nhat_trang_thai)
                
                self.cap_nhat_trang_thai("✅ Đã dọn dẹp hệ thống sạch sẽ không tì vết!")
                messagebox.showinfo("Thành công", "Tiến trình Uninstall & Deep Clean của VietToolbox đã xong.\nMáy tính đã được giải phóng hoàn toàn khỏi Office!")
            else:
                self.cap_nhat_trang_thai("❌ Lỗi: Không kéo được công cụ gỡ cài đặt từ mạng.")
        finally:
            self.after(0, lambda: self.thanh_tien_do.stop())

    def hanh_dong_go_kms(self):
        if messagebox.askyesno("Xác nhận Tẩy Rửa", "Tiến hành xóa KMS ảo và Reset trạng thái bản quyền?"):
            threading.Thread(target=self.luong_xu_ly_go_kms, daemon=True).start()

    def luong_xu_ly_go_kms(self):
        self.after(0, lambda: self.thanh_tien_do.start(15))
        try:
            self.cap_nhat_trang_thai("⏳ Đang cào sạch hệ thống giả lập KMS...")
            danh_sach_thu_muc = [os.environ.get("ProgramFiles", "C:\\Program Files") + "\\Microsoft Office\\Office16", os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)") + "\\Microsoft Office\\Office16"]
            file_vbs_kms = next((os.path.join(tm, "ospp.vbs") for tm in danh_sach_thu_muc if os.path.exists(os.path.join(tm, "ospp.vbs"))), None)
            if file_vbs_kms:
                subprocess.run(["cscript", "//nologo", file_vbs_kms, "/remhst"], creationflags=subprocess.CREATE_NO_WINDOW)
                subprocess.run(["cscript", "//nologo", file_vbs_kms, "/rearm"], creationflags=subprocess.CREATE_NO_WINDOW)
                self.cap_nhat_trang_thai("✅ Đã trục xuất hoàn toàn mầm mống KMS cũ!")
                messagebox.showinfo("Thành công", "Đã xóa KMS ảo và Reset chứng chỉ bản quyền.\nĐề nghị bác khởi động lại máy tính.")
            else:
                self.cap_nhat_trang_thai("⚠️ Không dò thấy cấu trúc tệp tin Office để xử lý.")
        finally:
            self.after(0, lambda: self.thanh_tien_do.stop())

    def luong_xu_ly_thuoc_chung(self, tham_so_lenh, tin_nhan_thanh_cong):
        self.after(0, lambda: self.thanh_tien_do.start(15))
        try:
            TienIchHeThong.kich_hoat_ohook_ngam(tham_so_lenh)
            self.cap_nhat_trang_thai(tin_nhan_thanh_cong)
            messagebox.showinfo("Phản hồi hệ thống", tin_nhan_thanh_cong[2:])
        finally:
            self.after(0, lambda: self.thanh_tien_do.stop())

    def hanh_dong_bom_thuoc(self):
        self.cap_nhat_trang_thai("⏳ Đang ép xung Kích hoạt Ohook vĩnh viễn...")
        threading.Thread(target=self.luong_xu_ly_thuoc_chung, args=("/Ohook", "✅ Đã bơm thuốc Ohook thành công vào cấu trúc Office!"), daemon=True).start()

    def hanh_dong_rut_thuoc_ohook(self):
        if not messagebox.askyesno("Cảnh báo Gỡ Thuốc", "Xác nhận gỡ bỏ hoàn toàn Hook bản quyền Ohook?"):
            return
        self.cap_nhat_trang_thai("⏳ Đang thanh lọc Ohook khỏi hệ thống...")
        threading.Thread(target=self.luong_xu_ly_thuoc_chung, args=("/OhookUninstall", "✅ Đã rút thuốc và khôi phục sự trong sạch cho Office!"), daemon=True).start()

    def hanh_dong_cuu_ho_zombie(self):
        if messagebox.askyesno("Cảnh báo Cứu Hộ", "Tính năng này sẽ xóa cưỡng bức dịch vụ ClickToRunSvc và toàn bộ lõi thư mục.\n\nSử dụng để trị lỗi 'mất Registry không thể gỡ/cài lại'. Tiếp tục?"):
            threading.Thread(target=self.luong_xu_ly_cuu_ho, daemon=True).start()

    def luong_xu_ly_cuu_ho(self):
        self.after(0, lambda: self.thanh_tien_do.config(mode='indeterminate', value=0))
        self.after(0, lambda: self.thanh_tien_do.start(15))
        try:
            self.cap_nhat_trang_thai("⏳ Đang tiêu diệt tiến trình C2R ngầm...")
            os.system('taskkill /F /IM ClickToRunSvc.exe >nul 2>&1')
            os.system('taskkill /F /IM OfficeClickToRun.exe >nul 2>&1')
            os.system('taskkill /F /IM setup.exe >nul 2>&1')
            
            self.cap_nhat_trang_thai("⏳ Đang nhổ rễ dịch vụ hệ thống ClickToRunSvc...")
            os.system('sc delete ClickToRunSvc >nul 2>&1')
            os.system('schtasks /Delete /TN "Microsoft\\Office" /F >nul 2>&1')
            
            self.cap_nhat_trang_thai("⏳ Đang xóa cưỡng bức thư mục gốc...")
            danh_sach_thu_muc_rac = [
                os.path.join(os.environ.get('ProgramFiles', 'C:\\Program Files'), "Microsoft Office"),
                os.path.join(os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)'), "Microsoft Office"),
                os.path.join(os.environ.get('ProgramFiles', 'C:\\Program Files'), "Common Files\\Microsoft Shared\\ClickToRun"),
                os.path.join(os.environ.get('ProgramData', 'C:\\ProgramData'), "Microsoft\\Office")
            ]
            for thu_muc in danh_sach_thu_muc_rac:
                if os.path.exists(thu_muc):
                    try: shutil.rmtree(thu_muc, ignore_errors=True)
                    except: pass
                    
            self.cap_nhat_trang_thai("✅ Cứu hộ thành công! Hệ thống đã sẵn sàng cài mới.")
            messagebox.showinfo("Hoàn Tất Xóa Ép Buộc", "Đã diệt gọn tàn dư lõi C2R.\nBây giờ bác có thể quay lại tab 'Triển Khai & Kích Hoạt' để cài mới bình thường!")
        finally:
            self.after(0, lambda: self.thanh_tien_do.stop())

    # ==========================================================================
    # LUỒNG THỰC THI TỐI ƯU HÓA OFFICE THÔNG MINH
    # ==========================================================================
    def hanh_dong_thuc_thi_toi_uu(self):
        threading.Thread(target=self.luong_xu_ly_toi_uu, daemon=True).start()

    def luong_xu_ly_toi_uu(self):
        self.after(0, lambda: self.thanh_tien_do.config(mode='indeterminate', value=0))
        self.after(0, lambda: self.thanh_tien_do.start(15))
        try:
            # 1. XỬ LÝ NGHỊ ĐỊNH 30 QUA VBSCRIPT VÀO NORMAL.DOTM
            if self.bien_nghi_dinh.get():
                self.cap_nhat_trang_thai("⏳ Đang nạp cấu hình Nghị định 30 vào lõi Template Word...")
                ma_lenh_vbs = """On Error Resume Next
Set objWord = CreateObject("Word.Application")
objWord.Visible = False
Set objDoc = objWord.Documents.Add()
Set objTemplate = objWord.NormalTemplate
objTemplate.OpenAsDocument
Set activeDoc = objWord.ActiveDocument
activeDoc.Styles("Normal").Font.Name = "Times New Roman"
activeDoc.Styles("Normal").Font.Size = 14
activeDoc.PageSetup.TopMargin = 56.7
activeDoc.PageSetup.BottomMargin = 56.7
activeDoc.PageSetup.LeftMargin = 85.05
activeDoc.PageSetup.RightMargin = 56.7
activeDoc.Save
activeDoc.Close
objDoc.Close False
objWord.Quit
"""
                file_vbs_nd30 = os.path.join(os.environ['TEMP'], 'toi_uu_nd30.vbs')
                with open(file_vbs_nd30, 'w', encoding='utf-8') as f:
                    f.write(ma_lenh_vbs)
                subprocess.run(['cscript', '//nologo', file_vbs_nd30], creationflags=subprocess.CREATE_NO_WINDOW)

            # 2. TẮT MÀN HÌNH WELCOME
            if self.bien_tat_welcome.get():
                self.cap_nhat_trang_thai("⏳ Đang ép Office khởi động thẳng vào trang trắng...")
                danh_sach_app = ["Word", "Excel", "PowerPoint"]
                for app in danh_sach_app:
                    os.system(f'reg add "HKCU\\Software\\Microsoft\\Office\\16.0\\{app}\\Options" /v DisableBootToOfficeStart /t REG_DWORD /d 1 /f >nul 2>&1')

            # 3. TẮT PROTECTED VIEW (MỞ FILE ZALO KHÔNG CẢNH BÁO)
            if self.bien_tat_protect.get():
                self.cap_nhat_trang_thai("⏳ Đang tiêu diệt rào cản Protected View...")
                danh_sach_app = ["Word", "Excel", "PowerPoint"]
                for app in danh_sach_app:
                    khoa_reg = f'"HKCU\\Software\\Microsoft\\Office\\16.0\\{app}\\Security\\ProtectedView"'
                    os.system(f'reg add {khoa_reg} /v DisableAttachementsInPV /t REG_DWORD /d 1 /f >nul 2>&1')
                    os.system(f'reg add {khoa_reg} /v DisableInternetFilesInPV /t REG_DWORD /d 1 /f >nul 2>&1')
                    os.system(f'reg add {khoa_reg} /v DisableUnsafeLocationsInPV /t REG_DWORD /d 1 /f >nul 2>&1')

            # 4. TẮT HARDWARE ACCELERATION
            if self.bien_tat_hw_accel.get():
                self.cap_nhat_trang_thai("⏳ Đang vô hiệu hóa Hardware Acceleration chống đen màn hình...")
                os.system('reg add "HKCU\\Software\\Microsoft\\Office\\16.0\\Common\\Graphics" /v DisableHardwareAcceleration /t REG_DWORD /d 1 /f >nul 2>&1')

            # 5. BẬT AUTOSAVE 3 PHÚT
            if self.bien_bat_autosave.get():
                self.cap_nhat_trang_thai("⏳ Đang siết chu kỳ AutoSave xuống 3 phút...")
                os.system('reg add "HKCU\\Software\\Microsoft\\Office\\16.0\\Word\\Options" /v AutoSaveInterval /t REG_DWORD /d 3 /f >nul 2>&1')
                os.system('reg add "HKCU\\Software\\Microsoft\\Office\\16.0\\Excel\\Options" /v AutoSaveInterval /t REG_DWORD /d 3 /f >nul 2>&1')

            self.cap_nhat_trang_thai("✅ TỐI ƯU HOÀN TẤT: Office đã được buff sức mạnh tối đa!")
            messagebox.showinfo("VietToolbox Optimize", "Quá trình ép xung phần mềm đã hoàn tất!\n- Căn lề chuẩn nhà nước.\n- Tắt bảng vàng báo lỗi Zalo.\n- Khởi động một phát ăn ngay.\n\nBác có thể mở Word lên để test ngay nhé!")
        finally:
            self.after(0, lambda: self.thanh_tien_do.stop())

if __name__ == "__main__":
    phan_mem_viettoolbox = TrienKhaiOffice()
    phan_mem_viettoolbox.mainloop()