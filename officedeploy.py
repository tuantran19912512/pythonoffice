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
from tkinter import ttk, messagebox
from tkinter import filedialog

socket.setdefaulttimeout(10)

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
# ĐỘNG CƠ TẢI MICRO-RESUMING (16 LUỒNG CHỐNG BÓP BĂNG THÔNG)
# ==============================================================================
class DongCoTaiDaLuong:
    def __init__(self, url, file_luu, so_luong=16):
        self.url = url
        self.file_luu = file_luu
        self.so_luong = so_luong
        self.tong_dung_luong = 0
        self.da_tai = 0
        self.loi = False
        self.huy_tai = False
        self.tieu_de = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'}
        self.khoa = threading.Lock()

    def lay_dung_luong(self):
        try:
            req = urllib.request.Request(self.url, method='HEAD', headers=self.tieu_de)
            with urllib.request.urlopen(req, timeout=10) as res:
                self.tong_dung_luong = int(res.headers.get('Content-Length', 0))
            return self.tong_dung_luong > 0
        except: return False

    def tho_tai(self, hang_doi):
        while not hang_doi.empty() and not self.loi and not self.huy_tai:
            try: start_goc, end, idx = hang_doi.get_nowait()
            except queue.Empty: break
                
            part_file = f"{self.file_luu}.part{idx}"
            start_hien_tai = start_goc
            
            if os.path.exists(part_file):
                dung_luong_da_tai = os.path.getsize(part_file)
                if dung_luong_da_tai == (end - start_goc + 1):
                    with self.khoa: self.da_tai += dung_luong_da_tai
                    hang_doi.task_done()
                    continue
                else:
                    start_hien_tai += dung_luong_da_tai
                    with self.khoa: self.da_tai += dung_luong_da_tai

            so_lan_thu = 0
            thanh_cong = False
            while so_lan_thu < 10 and not thanh_cong and not self.loi and not self.huy_tai:
                try:
                    req = urllib.request.Request(self.url, headers={'Range': f'bytes={start_hien_tai}-{end}', **self.tieu_de})
                    with urllib.request.urlopen(req, timeout=10) as res, open(part_file, 'ab') as f:
                        while True:
                            if self.huy_tai: break
                            chunk = res.read(16384) 
                            if not chunk: break
                            f.write(chunk)
                            with self.khoa: self.da_tai += len(chunk)
                            start_hien_tai += len(chunk)
                    thanh_cong = True
                except Exception:
                    so_lan_thu += 1
                    time.sleep(1)
                    
            if not thanh_cong and not self.huy_tai: self.loi = True
            hang_doi.task_done()

    def chay(self):
        hang_doi = queue.Queue()
        kieu_chunk = 20 * 1024 * 1024
        so_chunk = self.tong_dung_luong // kieu_chunk
        if self.tong_dung_luong % kieu_chunk != 0: so_chunk += 1

        for i in range(so_chunk):
            start = i * kieu_chunk
            end = min(start + kieu_chunk - 1, self.tong_dung_luong - 1)
            hang_doi.put((start, end, i))

        luong_list = []
        for _ in range(self.so_luong):
            t = threading.Thread(target=self.tho_tai, args=(hang_doi,))
            t.start()
            luong_list.append(t)

        for t in luong_list: t.join()

        if self.huy_tai:
            self.don_rac(so_chunk)
            return False
        if self.loi: return False

        with open(self.file_luu, 'wb') as outfile:
            for i in range(so_chunk):
                part_file = f"{self.file_luu}.part{i}"
                if os.path.exists(part_file):
                    with open(part_file, 'rb') as infile: outfile.write(infile.read())
                    os.remove(part_file)
        return True

    def don_rac(self, so_chunk):
        for i in range(so_chunk):
            p = f"{self.file_luu}.part{i}"
            if os.path.exists(p):
                try: os.remove(p)
                except: pass

# ==============================================================================
# LỚP GIAO DIỆN VÀ XỬ LÝ CHÍNH
# ==============================================================================
class UngDungCaiDatOffice:
    def __init__(self, cua_so_chinh):
        self.cua_so = cua_so_chinh
        self.cua_so.title("Cài đặt Microsoft Office - Max Speed (V8 Ultimate C2R)")
        self.cua_so.geometry("640x720")
        self.cua_so.resizable(False, False)
        self.phong_chu_dam = ("Segoe UI", 9, "bold")
        self.thu_muc_luu_file = tk.StringVar(value=os.getcwd())
        self.may_tai_hien_tai = None
        self.xay_dung_giao_dien()

    def tao_nut_bam_mau(self, khung_chua, chu_hien_thi, mau_nen, mau_chu="white", hanh_dong=None, width=20):
        return tk.Button(khung_chua, text=chu_hien_thi, bg=mau_nen, fg=mau_chu, font=self.phong_chu_dam, relief="flat", cursor="hand2", command=hanh_dong, width=width, pady=6)

    def xay_dung_giao_dien(self):
        khung_tieu_de = tk.Frame(self.cua_so, bg="#E64A19", height=65)
        khung_tieu_de.pack(fill="x", side="top")
        tk.Label(khung_tieu_de, text="🏢 CÀI ĐẶT MICROSOFT OFFICE RETAIL", font=("Segoe UI", 14, "bold"), bg="#E64A19", fg="white").place(x=15, y=8)
        tk.Label(khung_tieu_de, text="Kiến trúc tải C2R thông minh (~2.9GB) Max Speed", font=("Segoe UI", 9), bg="#E64A19", fg="white").place(x=40, y=35)

        hop_tab = ttk.Notebook(self.cua_so)
        hop_tab.pack(fill="both", expand=True, padx=12, pady=12)

        tab_cai_dat = ttk.Frame(hop_tab)
        tab_go_cai_dat = ttk.Frame(hop_tab)
        hop_tab.add(tab_cai_dat, text="  ⚙️ Cài đặt & Kích hoạt  ")
        hop_tab.add(tab_go_cai_dat, text="  🗑️ Gỡ Cài Đặt & Dọn Dẹp  ")

        self.xay_dung_tab_cai_dat(tab_cai_dat)
        self.xay_dung_tab_go_cai_dat(tab_go_cai_dat)

        khung_trang_thai = tk.Frame(self.cua_so, bg="#E0F7FA")
        khung_trang_thai.pack(fill="x", padx=12, pady=(0, 12))
        self.nhan_trang_thai = tk.Label(khung_trang_thai, text="✅ Sẵn sàng kết nối tới hệ thống...", font=self.phong_chu_dam, fg="#2E7D32", bg="#E0F7FA")
        self.nhan_trang_thai.pack(anchor="w", padx=10, pady=(10, 5))
        self.thanh_tien_do = ttk.Progressbar(khung_trang_thai, mode='determinate')
        self.thanh_tien_do.pack(fill="x", padx=12, pady=(0, 10))

    def xay_dung_tab_cai_dat(self, tab):
        khung_phien_ban = ttk.LabelFrame(tab, text=" 📄 Phiên bản Office ")
        khung_phien_ban.pack(fill="x", padx=10, pady=8, ipady=3)
        self.bien_phien_ban = tk.StringVar(value="2024")
        danh_sach_ban = ["Office 2016", "Office 2019", "Office 2021", "Office 2024", "Office 365"]
        khung_rb = ttk.Frame(khung_phien_ban)
        khung_rb.pack(fill="x", padx=10, pady=2)
        self.bien_phien_ban.trace_add("write", self.cap_nhat_hop_chon)
        for vi_tri, ten_ban in enumerate(danh_sach_ban):
            ttk.Radiobutton(khung_rb, text=ten_ban, variable=self.bien_phien_ban, value=ten_ban.split()[-1]).grid(row=0, column=vi_tri, padx=6)
            
        khung_cb = ttk.Frame(khung_phien_ban)
        khung_cb.pack(fill="x", padx=10, pady=5)
        ttk.Label(khung_cb, text="Phiên bản con:").pack(side="left", padx=(0, 10))
        self.hop_chon_ban_con = ttk.Combobox(khung_cb, state="readonly", width=35)
        self.cap_nhat_hop_chon()
        self.hop_chon_ban_con.pack(side="left")

        khung_ngang = ttk.Frame(tab)
        khung_ngang.pack(fill="x", padx=10, pady=5)
        khung_kt = ttk.LabelFrame(khung_ngang, text=" ⚙️ Kiến trúc ")
        khung_kt.pack(side="left", fill="both", expand=True, padx=(0, 5))
        self.bien_kien_truc = tk.StringVar(value="64")
        ttk.Radiobutton(khung_kt, text="64-bit", variable=self.bien_kien_truc, value="64").pack(side="left", padx=15, pady=5)
        ttk.Radiobutton(khung_kt, text="32-bit", variable=self.bien_kien_truc, value="32").pack(side="right", padx=15, pady=5)

        khung_nn = ttk.LabelFrame(khung_ngang, text=" 🌐 Ngôn ngữ ")
        khung_nn.pack(side="right", fill="both", expand=True, padx=(5, 0))
        self.hop_chon_ngon_ngu = ttk.Combobox(khung_nn, values=["English (US) - en-US", "Vietnamese - vi-VN"], state="readonly")
        self.hop_chon_ngon_ngu.current(0)
        self.hop_chon_ngon_ngu.pack(fill="x", padx=10, pady=6)

        khung_ud = ttk.LabelFrame(tab, text=" ☑️ Chọn ứng dụng cần cài ")
        khung_ud.pack(fill="x", padx=10, pady=5)
        danh_sach_app = [("Access",0,0),("Excel",0,1),("Word",0,2),("PowerPoint",1,0),("Outlook",1,1),("Publisher",1,2),("OneNote",2,0),("Skype",2,1),("Teams",2,2)]
        self.cac_bien_ung_dung = {}
        for ten, r, c in danh_sach_app:
            var = tk.BooleanVar(value=True)
            self.cac_bien_ung_dung[ten] = var
            ttk.Checkbutton(khung_ud, text=f" {ten}", variable=var).grid(row=r, column=c, sticky="w", padx=25, pady=4)

        khung_luu = ttk.LabelFrame(tab, text=" 📂 Thư mục chứa cấu trúc cài đặt C2R ")
        khung_luu.pack(fill="x", padx=10, pady=5)
        khung_con_luu = ttk.Frame(khung_luu)
        khung_con_luu.pack(fill="x", padx=10, pady=8)
        ttk.Entry(khung_con_luu, textvariable=self.thu_muc_luu_file, state="readonly").pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.tao_nut_bam_mau(khung_con_luu, "Đổi Thư Mục", "#757575", width=12, hanh_dong=self.chon_thu_muc).pack(side="right")

        khung_tc = ttk.LabelFrame(tab, text=" 🔧 Tùy chọn tự động sau cài đặt ")
        khung_tc.pack(fill="x", padx=10, pady=5)
        self.bien_tu_dong_ohook = tk.BooleanVar(value=True)
        ttk.Checkbutton(khung_tc, text="Tự động Kích hoạt Ohook", variable=self.bien_tu_dong_ohook).pack(side="left", padx=25, pady=6)
        self.bien_tao_shortcut = tk.BooleanVar(value=True)
        ttk.Checkbutton(khung_tc, text="Đưa lối tắt ra Desktop", variable=self.bien_tao_shortcut).pack(side="left", padx=25, pady=6)

        khung_btn = ttk.Frame(tab)
        khung_btn.pack(fill="x", padx=10, pady=5)
        self.tao_nut_bam_mau(khung_btn, "💊 KÍCH HOẠT OHOOK", "#F57C00", hanh_dong=self.khoi_dong_thuoc).pack(side="left", padx=5)
        self.nut_cai_dat = self.tao_nut_bam_mau(khung_btn, "🚀 BẮT ĐẦU CÀI ĐẶT", "#0288D1", hanh_dong=self.khoi_dong_cai_dat)
        self.nut_cai_dat.pack(side="right", padx=5)

    def xay_dung_tab_go_cai_dat(self, tab):
        khung_go_office = ttk.LabelFrame(tab, text=" 🗑️ Gỡ Cài Đặt Office Toàn Diện ")
        khung_go_office.pack(fill="x", padx=10, pady=10)
        ttk.Label(khung_go_office, text="Thao tác này sẽ gọi trình gỡ cài đặt gốc của Microsoft để xóa sạch\ntoàn bộ các phiên bản Office (Word, Excel...) đang có trên máy tính.", justify="left").pack(anchor="w", padx=15, pady=5)
        khung_nut_go = ttk.Frame(khung_go_office)
        khung_nut_go.pack(fill="x", padx=10, pady=5)
        self.tao_nut_bam_mau(khung_nut_go, "🗑 BẮT ĐẦU GỠ OFFICE", "#D32F2F", hanh_dong=self.khoi_dong_go_office).pack(side="right", padx=5, pady=5)

        khung_go_kms = ttk.LabelFrame(tab, text=" 🧹 Gỡ Crack KMS / Reset Bản Quyền ")
        khung_go_kms.pack(fill="x", padx=10, pady=10)
        ttk.Label(khung_go_kms, text="Dọn dẹp các máy chủ KMS ảo, làm sạch Registry kích hoạt lậu và\nđưa trạng thái bản quyền Office về như máy mới (Rearm).", justify="left").pack(anchor="w", padx=15, pady=5)
        khung_nut_kms = ttk.Frame(khung_go_kms)
        khung_nut_kms.pack(fill="x", padx=10, pady=5)
        self.tao_nut_bam_mau(khung_nut_kms, "🧹 DỌN DẸP KMS", "#1976D2", hanh_dong=self.khoi_dong_go_kms).pack(side="right", padx=5, pady=5)

        khung_go_ohook = ttk.LabelFrame(tab, text=" 💊 Gỡ Crack Ohook ")
        khung_go_ohook.pack(fill="x", padx=10, pady=10)
        ttk.Label(khung_go_ohook, text="Gỡ bỏ hoàn toàn Hook kích hoạt bản quyền vĩnh viễn (Ohook)\nvà trả lại các file hệ thống nguyên bản cho ứng dụng Office.", justify="left").pack(anchor="w", padx=15, pady=5)
        khung_nut_ohook = ttk.Frame(khung_go_ohook)
        khung_nut_ohook.pack(fill="x", padx=10, pady=5)
        self.tao_nut_bam_mau(khung_nut_ohook, "🛡️ GỠ OHOOK TRIỆT ĐỂ", "#7B1FA2", hanh_dong=self.khoi_dong_go_ohook_xoa).pack(side="right", padx=5, pady=5)

    def cap_nhat_hop_chon(self, *args):
        if self.bien_phien_ban.get() == "365":
            self.hop_chon_ban_con['values'] = ["ProPlus", "Business", "Home Premium"]
        else:
            self.hop_chon_ban_con['values'] = ["ProPlus", "Standard", "Home & Business", "Home & Student"]
        self.hop_chon_ban_con.current(0)

    def chon_thu_muc(self):
        thu_muc_moi = filedialog.askdirectory(initialdir=self.thu_muc_luu_file.get(), title="Chọn thư mục tải bản cài")
        if thu_muc_moi:
            if os.path.exists(thu_muc_moi) and os.access(thu_muc_moi, os.W_OK):
                self.thu_muc_luu_file.set(thu_muc_moi)
            else:
                messagebox.showerror("Lỗi Truy Cập", "Thư mục không hợp lệ hoặc không có quyền ghi.")

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

    def huy_tien_trinh(self):
        if messagebox.askyesno("Xác nhận hủy", "Bạn có chắc chắn muốn ngừng tải và hủy bỏ quá trình cài đặt không?"):
            if self.may_tai_hien_tai: self.may_tai_hien_tai.huy_tai = True
            self.cap_nhat_trang_thai("⏳ Đang dọn dẹp kết nối và xóa các file tạm...")

    def khoi_dong_cai_dat(self):
        self.nut_cai_dat.config(text="🛑 HỦY TẢI & CÀI ĐẶT", bg="#D32F2F", command=self.huy_tien_trinh)
        threading.Thread(target=self.tien_trinh_cai_dat_thong_minh, daemon=True).start()

    def phuc_hoi_nut_cai_dat(self):
        self.cua_so.after(0, lambda: self.nut_cai_dat.config(text="🚀 BẮT ĐẦU CÀI ĐẶT", bg="#0288D1", command=self.khoi_dong_cai_dat))

    def tien_trinh_cai_dat_thong_minh(self):
        nam_pb = self.bien_phien_ban.get()
        ban_con = self.hop_chon_ban_con.get().replace(" & ", "").replace(" ", "")
        ma_san_pham = tu_dien_phien_ban.get(f"{nam_pb}_{ban_con}", "ProPlus2024Retail")
        
        kien_truc_so = "64" if self.bien_kien_truc.get() == "64" else "32"
        kien_truc_chu = "x64" if kien_truc_so == "64" else "x86"
        ngon_ngu = "vi-VN" if "Vietnamese" in self.hop_chon_ngon_ngu.get() else "en-US"
        
        ma_lcid = "1066" if ngon_ngu == "vi-VN" else "1033"
        thu_muc_goc = self.thu_muc_luu_file.get()
        
        self.cap_nhat_trang_thai("🔍 Đang bung tệp Danh mục (.cab) để dò mã Phiên bản...")
        base_url = "https://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60"
        version_hien_tai = None
        try:
            tm_cab = os.path.join(os.environ['TEMP'], "OfficeCabTemp")
            if os.path.exists(tm_cab): shutil.rmtree(tm_cab)
            os.makedirs(tm_cab, exist_ok=True)
            
            cab_file = os.path.join(tm_cab, f"v{kien_truc_so}.cab")
            req = urllib.request.Request(f"{base_url}/Office/Data/v{kien_truc_so}.cab", headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=15) as res, open(cab_file, 'wb') as f:
                f.write(res.read())
                
            subprocess.run(["expand.exe", cab_file, "-F:*.xml", tm_cab], creationflags=subprocess.CREATE_NO_WINDOW)
            
            for f_xml in glob.glob(os.path.join(tm_cab, "*.xml")):
                noi_dung = open(f_xml, 'r', encoding='utf-8', errors='ignore').read()
                match = re.search(r'Version="(\d{2}\.\d+\.\d+\.\d+)"', noi_dung)
                if match:
                    version_hien_tai = match.group(1)
                    break
                    
            if not version_hien_tai: raise Exception("Không tìm thấy Version XML")
        except Exception as e:
            self.cap_nhat_trang_thai("❌ Lỗi: Không thể phân tích Version từ Microsoft.")
            self.phuc_hoi_nut_cai_dat()
            return

        tm_data = os.path.join(thu_muc_goc, "Office", "Data")
        tm_version = os.path.join(tm_data, version_hien_tai)
        os.makedirs(tm_version, exist_ok=True)

        danh_sach_tai = [
            (f"{base_url}/Office/Data/v{kien_truc_so}.cab", os.path.join(tm_data, f"v{kien_truc_so}.cab"), "Danh mục gốc"),
            (f"{base_url}/Office/Data/v{kien_truc_so}_{version_hien_tai}.cab", os.path.join(tm_data, f"v{kien_truc_so}_{version_hien_tai}.cab"), "Danh mục phiên bản"),
            (f"{base_url}/Office/Data/{version_hien_tai}/i{kien_truc_so}0.cab", os.path.join(tm_version, f"i{kien_truc_so}0.cab"), "Mục lục lõi (i0)"),
            (f"{base_url}/Office/Data/{version_hien_tai}/s{kien_truc_so}0.cab", os.path.join(tm_version, f"s{kien_truc_so}0.cab"), "Mục lục lõi (s0)"),
            (f"{base_url}/Office/Data/{version_hien_tai}/i{kien_truc_so}{ma_lcid}.cab", os.path.join(tm_version, f"i{kien_truc_so}{ma_lcid}.cab"), "Mục lục ngôn ngữ (iLang)"),
            (f"{base_url}/Office/Data/{version_hien_tai}/s{kien_truc_so}{ma_lcid}.cab", os.path.join(tm_version, f"s{kien_truc_so}{ma_lcid}.cab"), "Mục lục ngôn ngữ (sLang)"),
            (f"{base_url}/Office/Data/{version_hien_tai}/stream.{kien_truc_chu}.x-none.dat", os.path.join(tm_version, f"stream.{kien_truc_chu}.x-none.dat"), "Dữ liệu Office siêu nặng"),
            (f"{base_url}/Office/Data/{version_hien_tai}/stream.{kien_truc_chu}.{ngon_ngu}.dat", os.path.join(tm_version, f"stream.{kien_truc_chu}.{ngon_ngu}.dat"), "Dữ liệu Ngôn ngữ")
        ]

        for url_tai, duong_dan_luu, ten_goi in danh_sach_tai:
            if self.may_tai_hien_tai and self.may_tai_hien_tai.huy_tai: break
            
            req_check = urllib.request.Request(url_tai, method='HEAD', headers={'User-Agent': 'Mozilla/5.0'})
            try: dung_luong_that = int(urllib.request.urlopen(req_check, timeout=10).headers.get('Content-Length', 0))
            except: dung_luong_that = 0
            if os.path.exists(duong_dan_luu) and os.path.getsize(duong_dan_luu) == dung_luong_that:
                continue

            self.may_tai_hien_tai = DongCoTaiDaLuong(url_tai, duong_dan_luu, 16)
            if self.may_tai_hien_tai.lay_dung_luong():
                luong_tai = threading.Thread(target=self.may_tai_hien_tai.chay)
                luong_tai.start()
                
                thoi_gian_bat_dau = time.time()
                lan_truoc_dl = 0
                
                while luong_tai.is_alive():
                    time.sleep(0.5)
                    giay_troi = time.time() - thoi_gian_bat_dau
                    if giay_troi > 0:
                        toc_do_mb = ((self.may_tai_hien_tai.da_tai - lan_truoc_dl) / 0.5) / 1048576
                        phan_tram = (self.may_tai_hien_tai.da_tai / self.may_tai_hien_tai.tong_dung_luong) * 100
                        self.cua_so.after(0, lambda p=phan_tram: self.thanh_tien_do.config(value=p))
                        self.cap_nhat_trang_thai(f"⬇️ Đang kéo {ten_goi}: {phan_tram:.1f}% | Tốc độ: {toc_do_mb:.1f} MB/s")
                    lan_truoc_dl = self.may_tai_hien_tai.da_tai
                    thoi_gian_bat_dau = time.time()

                if self.may_tai_hien_tai.huy_tai:
                    self.cap_nhat_trang_thai("❌ Đã hủy tải xuống thành công.")
                    self.cua_so.after(0, lambda: self.thanh_tien_do.config(value=0))
                    self.phuc_hoi_nut_cai_dat()
                    return
                if self.may_tai_hien_tai.loi:
                    self.cap_nhat_trang_thai(f"❌ Lỗi mạng khi tải {ten_goi}.")
                    self.phuc_hoi_nut_cai_dat()
                    return

        self.cua_so.after(0, lambda: self.thanh_tien_do.config(mode='indeterminate', value=0))
        self.cua_so.after(0, lambda: self.thanh_tien_do.start(15))
        self.cap_nhat_trang_thai("🚀 Đang khởi động trình cài đặt nội bộ. Vui lòng đợi bảng cam tắt...")
        
        duong_dan_setup = self.chuan_bi_cong_cu_odt_cho_go()
        if not duong_dan_setup:
            self.cap_nhat_trang_thai("❌ Lỗi: Không thể lấy bộ cài ODT mới nhất từ Microsoft!")
            self.phuc_hoi_nut_cai_dat()
            return

        app_chon = [t for t, v in self.cac_bien_ung_dung.items() if v.get()]
        
        xml_code = f"""<Configuration>\n  <Add SourcePath="{thu_muc_goc}" OfficeClientEdition="{kien_truc_so}" Channel="Current" Version="{version_hien_tai}" AllowCdnFallback="True">\n    <Product ID="{ma_san_pham}">\n      <Language ID="{ngon_ngu}" />\n"""
        for t, m in tu_dien_ung_dung.items():
            if t not in app_chon: xml_code += f'      <ExcludeApp ID="{m}" />\n'
        xml_code += """    </Product>\n  </Add>\n  <Updates Enabled="TRUE" />\n  <Display Level="Full" AcceptEULA="TRUE" />\n</Configuration>"""
        
        file_xml = os.path.join(thu_muc_goc, "C2R_Config.xml")
        with open(file_xml, "w", encoding="utf-8") as f: f.write(xml_code)
        
        subprocess.Popen([duong_dan_setup, "/configure", file_xml]).wait()
        
        if os.path.exists(file_xml): os.remove(file_xml)
        if messagebox.askyesno("Dọn dẹp", "Cài đặt thành công!\nBạn có muốn XÓA thư mục 2.9GB tải về để tiết kiệm ổ cứng không?"):
            try: shutil.rmtree(os.path.join(thu_muc_goc, "Office"))
            except: pass
        
        if self.bien_tao_shortcut.get(): self.tao_loi_tat_desktop()
        if self.bien_tu_dong_ohook.get():
            self.cap_nhat_trang_thai("⏳ Đang tự động Kích hoạt Ohook...")
            self.chay_gist_ngam("/Ohook")
            
        self.cua_so.after(0, lambda: self.thanh_tien_do.stop())
        self.cap_nhat_trang_thai("✅ HOÀN TẤT: Đã cài đặt C2R siêu tốc độ thành công!")
        self.phuc_hoi_nut_cai_dat()
        messagebox.showinfo("Thành công", "Lỗi 30029-2016 đã bị triệt tiêu hoàn toàn!")

    # BẢN VÁ LỖI CỰC MẠNH: TẢI TRỰC TIẾP ODT TỪ MÁY CHỦ MICROSOFT THAY VÌ GITHUB
    def chuan_bi_cong_cu_odt_cho_go(self):
        duong_dan_file_setup = os.path.join(os.environ['TEMP'], "setup.exe")
        
        # Luôn ưu tiên dùng file mới nhất trong Temp nếu có
        if os.path.exists(duong_dan_file_setup): 
            try: os.remove(duong_dan_file_setup) # Xóa file cũ đi để kéo bản mới nhất cho chắc
            except: pass

        self.cap_nhat_trang_thai("⏳ Đang tải công cụ Setup chuẩn 2024 từ Server Microsoft...")
        # Link tải file cài đặt ODT (Office Deployment Tool) bản quyền mới nhất
        link_odt = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A8D4A7E/officedeploymenttool_39329-33611.exe"
        file_odt = os.path.join(os.environ['TEMP'], "odt_installer.exe")
        try:
            req = urllib.request.Request(link_odt, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=30) as phan_hoi, open(file_odt, 'wb') as f: 
                f.write(phan_hoi.read())
            
            self.cap_nhat_trang_thai("⏳ Đang trích xuất lõi cài đặt Setup.exe...")
            # Gọi lệnh giải nén ngầm (không hiện cửa sổ) để móc thằng setup.exe ra
            thu_muc_temp = os.environ['TEMP']
            subprocess.run([file_odt, f"/extract:{thu_muc_temp}", "/quiet"], creationflags=subprocess.CREATE_NO_WINDOW)
            
            if os.path.exists(duong_dan_file_setup):
                return duong_dan_file_setup
            return None
        except Exception as e: 
            return None

    def khoi_dong_go_office(self):
        if messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn gỡ toàn bộ Office khỏi máy tính không?"): threading.Thread(target=self.tien_trinh_go_office, daemon=True).start()

    def tien_trinh_go_office(self):
        self.cua_so.after(0, lambda: self.thanh_tien_do.config(mode='indeterminate', value=0))
        self.cua_so.after(0, lambda: self.thanh_tien_do.start(15))
        try:
            self.cap_nhat_trang_thai("⏳ Đang thiết lập cấu hình gỡ cài đặt...")
            noi_dung_xml = """<Configuration>\n  <Remove All="True" />\n  <Display Level="Full" AcceptEULA="TRUE" />\n</Configuration>"""
            duong_dan_xml = os.path.join(os.environ['TEMP'], "CauHinhGo.xml")
            with open(duong_dan_xml, "w", encoding="utf-8") as f: f.write(noi_dung_xml)
            duong_dan_setup = self.chuan_bi_cong_cu_odt_cho_go()
            if duong_dan_setup:
                self.cap_nhat_trang_thai("🚀 Đang chạy trình gỡ cài đặt của Microsoft...")
                subprocess.Popen([duong_dan_setup, "/configure", duong_dan_xml]).wait()
                self.cap_nhat_trang_thai("✅ Đã gọi lệnh gỡ cài đặt xong!")
        finally: self.cua_so.after(0, lambda: self.thanh_tien_do.stop())

    def khoi_dong_go_kms(self):
        if messagebox.askyesno("Xác nhận", "Xóa KMS ảo và Reset trạng thái bản quyền?"): threading.Thread(target=self.tien_trinh_go_kms, daemon=True).start()

    def tien_trinh_go_kms(self):
        self.cua_so.after(0, lambda: self.thanh_tien_do.start(15))
        try:
            self.cap_nhat_trang_thai("⏳ Đang dọn dẹp hệ thống KMS...")
            cac_thu_muc = [os.environ.get("ProgramFiles", "C:\\Program Files") + "\\Microsoft Office\\Office16", os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)") + "\\Microsoft Office\\Office16"]
            file_ospp = next((os.path.join(tm, "ospp.vbs") for tm in cac_thu_muc if os.path.exists(os.path.join(tm, "ospp.vbs"))), None)
            if file_ospp:
                subprocess.run(["cscript", "//nologo", file_ospp, "/remhst"], creationflags=subprocess.CREATE_NO_WINDOW)
                subprocess.run(["cscript", "//nologo", file_ospp, "/rearm"], creationflags=subprocess.CREATE_NO_WINDOW)
                self.cap_nhat_trang_thai("✅ Đã dọn dẹp sạch bản quyền KMS cũ!")
                messagebox.showinfo("Thành công", "Đã xóa KMS ảo và Reset trạng thái bản quyền.\nVui lòng khởi động lại máy tính.")
            else: self.cap_nhat_trang_thai("⚠️ Không tìm thấy file hệ thống Office.")
        finally: self.cua_so.after(0, lambda: self.thanh_tien_do.stop())

    def chay_gist_ngam(self, tham_so):
        url = f"https://gist.githubusercontent.com/tuantran19912512/81329d670436ea8492b73bd5889ad444/raw/Ohook.cmd?t={time.time()}"
        tmp = os.path.join(os.environ['TEMP'], "O.cmd")
        try:
            with open(tmp, 'w', encoding='utf-8') as f: f.write(urllib.request.urlopen(url).read().decode('utf-8').replace("\n", "\r\n"))
            subprocess.run(["cmd.exe", "/c", tmp, tham_so], creationflags=subprocess.CREATE_NO_WINDOW)
        except: pass
        finally:
            if os.path.exists(tmp):
                try: os.remove(tmp)
                except: pass

    def tien_trinh_gist_giao_dien(self, tham_so, loi_nhan):
        self.cua_so.after(0, lambda: self.thanh_tien_do.start(15))
        try:
            self.chay_gist_ngam(tham_so)
            self.cap_nhat_trang_thai(loi_nhan)
            messagebox.showinfo("Thành công", loi_nhan[2:])
        finally: self.cua_so.after(0, lambda: self.thanh_tien_do.stop())

    def khoi_dong_thuoc(self):
        self.cap_nhat_trang_thai("⏳ Đang Kích hoạt Ohook Silent...")
        threading.Thread(target=self.tien_trinh_gist_giao_dien, args=("/Ohook", "✅ Đã KÍCH HOẠT thành công bản quyền Ohook!"), daemon=True).start()

    def khoi_dong_go_ohook_xoa(self):
        if not messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn gỡ bỏ hoàn toàn Crack Ohook không?"): return
        self.cap_nhat_trang_thai("⏳ Đang gỡ Ohook...")
        threading.Thread(target=self.tien_trinh_gist_giao_dien, args=("/OhookUninstall", "✅ Đã GỠ BỎ Ohook và khôi phục file gốc thành công!"), daemon=True).start()

    def cap_nhat_trang_thai(self, msg):
        self.cua_so.after(0, lambda: self.nhan_trang_thai.config(text=msg))

if __name__ == "__main__":
    cua_so = tk.Tk()
    app = UngDungCaiDatOffice(cua_so)
    cua_so.mainloop()