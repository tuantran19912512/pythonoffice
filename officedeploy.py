# ==============================================================================
# BỘ CÀI OFFICE - GOOGLE DRIVE V600 (PHIÊN BẢN PYTHON - ĐỘC LẬP)
# ==============================================================================

import os
import sys
import time
import base64
import subprocess
import threading
import ctypes
import csv
import urllib.request
import urllib.error
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ==============================================================================
# [MÔ-ĐUN 1] QUYỀN QUẢN TRỊ VIÊN
# ==============================================================================
def kiem_tra_quyen_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not kiem_tra_quyen_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    sys.exit()

# ==============================================================================
# [MÔ-ĐUN 2] BIẾN ĐỒNG BỘ TOÀN CỤC
# ==============================================================================
tu_khoa_api = ["QUl6YVN5Q2V0SVlWVzRsQmlULTd3TzdNQUJoWlNVQ0dKR1puQTM0", "QUl6YVN5Q3VKUkJaTDZnUU8tdVZOMWVvdHhmMlppTXNtYy1sandR", "QUl6YVN5QlRhVmRQdmlLaUJyR0JUVk0tUlRiVW51QUdFUzRWck1v"]
du_lieu_office = []
trang_thai_app = {}

bien_nhat_ky = ""
bien_trang_thai = "Đang nạp..."
bien_lenh = "CHO"
bien_file_hien_tai = ""
bien_thu_muc_giai_nen = ""

# Thông số tải file
thong_so_phan_tram = 0
thong_so_toc_do = "0 MB/s"
thong_so_thong_tin = "0/0 MB"
thong_so_thoi_gian = "--:--"

# ==============================================================================
# [MÔ-ĐUN 3] ĐỘNG CƠ TẢI FILE (XỬ LÝ 403 & AUTO RESUME)
# ==============================================================================
def tai_file_mang(duong_dan_link, duong_dan_luu):
    global thong_so_phan_tram, thong_so_toc_do, thong_so_thong_tin, thong_so_thoi_gian, bien_lenh
    so_lan_thu = 5
    
    for lan in range(so_lan_thu):
        if bien_lenh == "DUNG": return -1
        
        dung_luong_cu = 0
        if os.path.exists(duong_dan_luu):
            dung_luong_cu = os.path.getsize(duong_dan_luu)
            
        yeu_cau = urllib.request.Request(duong_dan_link, headers={"User-Agent": "Mozilla/5.0"})
        if dung_luong_cu > 0:
            yeu_cau.add_header("Range", f"bytes={dung_luong_cu}-")
            
        try:
            with urllib.request.urlopen(yeu_cau, timeout=18000) as phan_hoi:
                tong_dung_luong = int(phan_hoi.headers.get("Content-Length", -1))
                if tong_dung_luong > 0 and dung_luong_cu > 0:
                    tong_dung_luong += dung_luong_cu
                elif tong_dung_luong <= 0:
                    tong_dung_luong = -1
                    
                che_do_ghi = 'ab' if dung_luong_cu > 0 and phan_hoi.status == 206 else 'wb'
                if che_do_ghi == 'wb':
                    dung_luong_cu = 0
                    
                thoi_gian_bat_dau = time.time()
                da_tai = dung_luong_cu
                
                with open(duong_dan_luu, che_do_ghi) as file_luu:
                    while True:
                        if bien_lenh == "DUNG": return -1
                        bo_nho_dem = phan_hoi.read(4194304) # 4MB Buffer
                        if not bo_nho_dem:
                            break
                            
                        file_luu.write(bo_nho_dem)
                        da_tai += len(bo_nho_dem)
                        
                        if tong_dung_luong > 0:
                            thong_so_phan_tram = int((da_tai * 100) / tong_dung_luong)
                            thoi_gian_qua = time.time() - thoi_gian_bat_dau
                            if thoi_gian_qua > 0:
                                byte_tren_giay = (da_tai - dung_luong_cu) / thoi_gian_qua
                                if byte_tren_giay > 0:
                                    thong_so_toc_do = f"{byte_tren_giay / 1048576.0:.2f} MB/s"
                                    thong_so_thong_tin = f"{da_tai / 1048576.0:.2f} / {tong_dung_luong / 1048576.0:.2f} MB"
                                    giay_con_lai = (tong_dung_luong - da_tai) / byte_tren_giay
                                    phut, giay = divmod(int(giay_con_lai), 60)
                                    thong_so_thoi_gian = f"{phut:02}:{giay:02}"
            return 200
        except urllib.error.HTTPError as loi_http:
            if loi_http.code == 403: return 403
            if loi_http.code == 416: # Range không hợp lệ
                try: os.remove(duong_dan_luu)
                except: pass
                continue
        except Exception:
            if bien_lenh == "DUNG": return -1
            time.sleep(3)
            
    return 500

# ==============================================================================
# [MÔ-ĐUN 4] LUỒNG XỬ LÝ CHÍNH
# ==============================================================================
def them_nhat_ky(tin_nhan):
    global bien_nhat_ky
    thoi_gian_thuc = time.strftime("%H:%M:%S")
    bien_nhat_ky += f"[{thoi_gian_thuc}] {tin_nhan}\n"

def giai_ma_b64(chuoi_ma):
    return base64.b64decode(chuoi_ma).decode('utf-8')

def tao_loi_tat_desktop():
    them_nhat_ky("📌 Đang tìm và đưa Shortcut ra Desktop...")
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    cac_thu_muc_office = [
        os.path.join(os.environ.get('ProgramFiles', 'C:\\Program Files'), 'Microsoft Office', 'root', 'Office16'),
        os.path.join(os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)'), 'Microsoft Office', 'root', 'Office16')
    ]
    cac_ung_dung = {"WINWORD.EXE": "Word", "EXCEL.EXE": "Excel", "POWERPNT.EXE": "PowerPoint", "MSACCESS.EXE": "Access", "OUTLOOK.EXE": "Outlook"}
    da_tao = 0
    
    for thu_muc in cac_thu_muc_office:
        if os.path.exists(thu_muc):
            for file_exe, ten_app in cac_ung_dung.items():
                muc_tieu = os.path.join(thu_muc, file_exe)
                if os.path.exists(muc_tieu):
                    duong_dan_lnk = os.path.join(desktop, f"{ten_app}.lnk")
                    ma_vbs = f'Set ws = CreateObject("WScript.Shell")\nSet link = ws.CreateShortcut("{duong_dan_lnk}")\nlink.TargetPath = "{muc_tieu}"\nlink.Save'
                    file_vbs = os.path.join(os.environ['TEMP'], 'tao_shortcut.vbs')
                    with open(file_vbs, 'w', encoding='utf-8') as f: f.write(ma_vbs)
                    subprocess.run(['cscript', '//nologo', file_vbs], creationflags=subprocess.CREATE_NO_WINDOW)
                    da_tao += 1
            if da_tao > 0: break

def kich_ban_xu_ly(danh_sach_chon, thu_muc_luu, co_thuoc, co_giu_file, co_loi_tat):
    global bien_trang_thai, bien_lenh, bien_file_hien_tai, bien_thu_muc_giai_nen
    
    try:
        # 1. Kiểm tra 7-Zip
        may_giai_nen = ""
        thu_muc_7z = [os.path.join(os.environ.get('ProgramFiles', ''), '7-Zip', '7z.exe'), os.path.join(os.environ.get('ProgramFiles(x86)', ''), '7-Zip', '7z.exe')]
        for duong_dan in thu_muc_7z:
            if os.path.exists(duong_dan):
                may_giai_nen = duong_dan
                break
                
        if not may_giai_nen:
            them_nhat_ky("📦 Máy chưa có 7-Zip. Đang tải và cài đặt 7-Zip...")
            file_7z = os.path.join(os.environ['TEMP'], '7z_setup.exe')
            if tai_file_mang("https://www.7-zip.org/a/7z2408-x64.exe", file_7z) == 200:
                subprocess.run([file_7z, "/S"])
                for duong_dan in thu_muc_7z:
                    if os.path.exists(duong_dan):
                        may_giai_nen = duong_dan
                        break

        chi_so_khoa = 0
        for phan_tu in danh_sach_chon:
            if bien_lenh == "DUNG": break
            duoi_file = ".iso" if ".img" in phan_tu['ID'] or ".iso" in phan_tu['ID'] else ".zip"
            ten_file_sach = "".join([c if c.isalnum() else "_" for c in phan_tu['Ten']])
            file_luu = os.path.join(thu_muc_luu, ten_file_sach + duoi_file)
            bien_file_hien_tai = file_luu
            
            trang_thai_app[phan_tu['ID']] = {"STT": "🚀 Đang tải", "PCT": "0%", "SPD": "--", "DL": "--"}
            them_nhat_ky(f"📡 [GOOGLE DRIVE]: {phan_tu['Ten']}")
            
            # Vòng lặp chống lỗi 403
            thanh_cong = False
            so_lan_thu = 0
            while not thanh_cong and so_lan_thu < len(tu_khoa_api) and bien_lenh != "DUNG":
                duong_dan_mang = f"https://www.googleapis.com/drive/v3/files/{phan_tu['ID']}?alt=media&key={giai_ma_b64(tu_khoa_api[chi_so_khoa])}&acknowledgeAbuse=true"
                ket_qua_tai = tai_file_mang(duong_dan_mang, file_luu)
                
                if ket_qua_tai == 200:
                    thanh_cong = True
                elif ket_qua_tai == 403:
                    chi_so_khoa = (chi_so_khoa + 1) % len(tu_khoa_api)
                    them_nhat_ky(f"⚠️ Bị giới hạn tải! Đang đổi sang API Key dự phòng số {chi_so_khoa + 1}...")
                    so_lan_thu += 1
                else:
                    break
                    
            if thanh_cong:
                trang_thai_app[phan_tu['ID']] = {"STT": "📦 Đang cài đặt", "PCT": "100%", "SPD": "Hoàn thành", "DL": "Hoàn thành"}
                
                # 2. Giải nén & Cài đặt
                thu_muc_giai_nen = file_luu + "_GiaiNen"
                bien_thu_muc_giai_nen = thu_muc_giai_nen
                
                try:
                    them_nhat_ky("💿 Đang giải nén file...")
                    tien_trinh_giai_nen = subprocess.Popen([may_giai_nen, "x", file_luu, f"-o{thu_muc_giai_nen}", "-pAdmin@2512", "-y"], creationflags=subprocess.CREATE_NO_WINDOW)
                    while tien_trinh_giai_nen.poll() is None:
                        if bien_lenh == "DUNG": tien_trinh_giai_nen.kill(); break
                        time.sleep(0.5)
                        
                    if bien_lenh != "DUNG":
                        file_chay = None
                        for goc, cac_thu_muc, cac_file in os.walk(thu_muc_giai_nen):
                            for ten_file in cac_file:
                                if ten_file.endswith(".bat") or ten_file.lower() == "setup.exe":
                                    file_chay = os.path.join(goc, ten_file)
                                    break
                            if file_chay: break
                            
                        if file_chay:
                            them_nhat_ky("🛠 Chạy bộ cài đặt ngầm...")
                            thu_muc_lam_viec = os.path.dirname(file_chay)
                            tien_trinh_cai = subprocess.Popen([file_chay], cwd=thu_muc_lam_viec)
                            while tien_trinh_cai.poll() is None:
                                if bien_lenh == "DUNG": tien_trinh_cai.kill(); break
                                time.sleep(0.5)
                        else:
                            them_nhat_ky("⚠️ Không tìm thấy file cài đặt trong thư mục giải nén.")
                except Exception as e:
                    them_nhat_ky(f"⚠️ Lỗi khi giải nén hoặc chạy cài đặt: {e}")

                if bien_lenh == "DUNG": break
                
                # 3. Kích hoạt Ohook
                if co_thuoc:
                    them_nhat_ky("==========================================")
                    them_nhat_ky(">>> KÍCH HOẠT OFFICE OHOOK (CHẠY NGẦM) <<<")
                    them_nhat_ky("==========================================")
                    url_gist = f"https://gist.githubusercontent.com/tuantran19912512/81329d670436ea8492b73bd5889ad444/raw/Ohook.cmd?t={time.time()}"
                    file_tam = os.path.join(os.environ['TEMP'], "Ohook_Activation.cmd")
                    
                    them_nhat_ky("-> Kiểm tra Internet (Ping 8.8.8.8)...")
                    if subprocess.run(["ping", "-n", "1", "8.8.8.8"], creationflags=subprocess.CREATE_NO_WINDOW).returncode == 0:
                        try:
                            them_nhat_ky("-> Đang tải file Ohook từ Gist...")
                            noi_dung = urllib.request.urlopen(url_gist).read().decode('utf-8')
                            noi_dung = noi_dung.replace("\r\n", "\n").replace("\n", "\r\n") + "\r\n\r\n"
                            with open(file_tam, 'w', encoding='utf-8') as f: f.write(noi_dung)
                            
                            them_nhat_ky("-> Chạy Ohook Silent...")
                            tien_trinh_thuoc = subprocess.Popen(["cmd.exe", "/c", file_tam, "/Ohook"], creationflags=subprocess.CREATE_NO_WINDOW)
                            while tien_trinh_thuoc.poll() is None:
                                if bien_lenh == "DUNG": tien_trinh_thuoc.kill(); break
                                time.sleep(0.5)
                                
                            if bien_lenh != "DUNG": them_nhat_ky("   + Đã kích hoạt bản quyền xong.")
                        except:
                            them_nhat_ky("!!! LỖI: Tải hoặc chạy file Gist thất bại.")
                        finally:
                            if os.path.exists(file_tam): os.remove(file_tam)
                    else:
                        them_nhat_ky("!!! LỖI: Mất mạng, không thể kích hoạt.")

                # 4. Tạo lối tắt & Dọn dẹp
                if co_loi_tat and bien_lenh != "DUNG": tao_loi_tat_desktop()
                if bien_lenh != "DUNG":
                    subprocess.run(["cmd", "/c", "rmdir", "/s", "/q", thu_muc_giai_nen], creationflags=subprocess.CREATE_NO_WINDOW)
                    if not co_giu_file:
                        try: os.remove(file_luu); them_nhat_ky("🧹 Đã xóa file nén nguồn.")
                        except: pass
                        
                trang_thai_app[phan_tu['ID']] = {"STT": "✅ Hoàn Tất", "PCT": "", "SPD": "", "DL": ""}
                them_nhat_ky(f"🎉 XONG: {phan_tu['Ten']}")
                
            else:
                if bien_lenh == "DUNG":
                    try: 
                        if os.path.exists(file_luu): os.remove(file_luu)
                    except: pass
                    trang_thai_app[phan_tu['ID']] = {"STT": "🛑 Đã Hủy", "PCT": "", "SPD": "", "DL": ""}
                    them_nhat_ky("🛑 Đã hủy và dọn rác.")
                else:
                    trang_thai_app[phan_tu['ID']] = {"STT": "❌ Lỗi Tải", "PCT": "", "SPD": "", "DL": ""}
                    them_nhat_ky("❌ Tải thất bại. Vui lòng kiểm tra dung lượng Google Drive hoặc mạng.")

    except Exception as e:
        them_nhat_ky(f"❌ LỖI HỆ THỐNG: {str(e)}")

    if bien_lenh == "DUNG": bien_trang_thai = "🛑 ĐÃ HỦY VÀ XÓA SẠCH RÁC"
    else: bien_trang_thai = "✅ HOÀN TẤT TOÀN BỘ YÊU CẦU"

# ==============================================================================
# [MÔ-ĐUN 5] GIAO DIỆN CHÍNH (TKINTER)
# ==============================================================================
cua_so = tk.Tk()
cua_so.title("OFFICE DEPLOY - GOOGLE DRIVE V600 (PYTHON EDITION)")
cua_so.geometry("950x750")
cua_so.configure(bg="#F4F6F8")
cua_so.eval('tk::PlaceWindow . center')

phong_chu_tieu_de = ("Segoe UI", 20, "bold")
phong_chu_phu = ("Segoe UI", 11)
phong_chu_dam = ("Segoe UI", 11, "bold")

khung_tieu_de = tk.Frame(cua_so, bg="#F4F6F8")
khung_tieu_de.pack(fill="x", padx=15, pady=10)
tk.Label(khung_tieu_de, text="MÁY CHỦ GOOGLE DRIVE - V600 MASTER", font=phong_chu_tieu_de, fg="#0277BD", bg="#F4F6F8").pack(anchor="w")
tk.Label(khung_tieu_de, text="☁ Xoay API Key | Tự Giải Nén (Admin@2512) | Ohook Silent | Dọn Rác Sạch", font=phong_chu_phu, fg="#555555", bg="#F4F6F8").pack(anchor="w")

cot_danh_sach = ("TEN", "TRANGTHAI", "TIENDO", "TOCDO", "DUNGLUONG")
danh_sach = ttk.Treeview(cua_so, columns=cot_danh_sach, show="headings", height=10)
danh_sach.heading("TEN", text="BẢN CÀI OFFICE (GOOGLE DRIVE)")
danh_sach.heading("TRANGTHAI", text="TRẠNG THÁI")
danh_sach.heading("TIENDO", text="TIẾN ĐỘ")
danh_sach.heading("TOCDO", text="TỐC ĐỘ")
danh_sach.heading("DUNGLUONG", text="DUNG LƯỢNG")

danh_sach.column("TEN", width=450)
danh_sach.column("TRANGTHAI", width=140)
danh_sach.column("TIENDO", width=70)
danh_sach.column("TOCDO", width=90)
danh_sach.column("DUNGLUONG", width=120)
danh_sach.pack(fill="both", expand=True, padx=15, pady=5)

khung_nhat_ky = tk.LabelFrame(cua_so, text="NHẬT KÝ HỆ THỐNG", font=phong_chu_dam, bg="#F4F6F8")
khung_nhat_ky.pack(fill="x", padx=15, pady=5)
hop_nhat_ky = tk.Text(khung_nhat_ky, height=8, bg="#1E1E1E", fg="#00E676", font=("Consolas", 10))
hop_nhat_ky.pack(fill="both", expand=True, padx=5, pady=5)

khung_thu_muc = tk.Frame(cua_so, bg="#F4F6F8")
khung_thu_muc.pack(fill="x", padx=15, pady=5)
hop_thu_muc = tk.Entry(khung_thu_muc, font=phong_chu_phu)
hop_thu_muc.pack(side="left", fill="x", expand=True)

def chon_thu_muc():
    thu_muc = filedialog.askdirectory()
    if thu_muc:
        hop_thu_muc.delete(0, tk.END)
        hop_thu_muc.insert(0, thu_muc)

def mo_thu_muc():
    thu_muc = hop_thu_muc.get()
    if os.path.exists(thu_muc):
        os.startfile(thu_muc)

tk.Button(khung_thu_muc, text="📂 CHỌN", font=phong_chu_dam, command=chon_thu_muc, width=10).pack(side="left", padx=5)
tk.Button(khung_thu_muc, text="MỞ FOLDER", font=phong_chu_dam, bg="#E3F2FD", command=mo_thu_muc, width=12).pack(side="left")

khung_tuy_chon = tk.Frame(cua_so, bg="#F4F6F8")
khung_tuy_chon.pack(fill="x", padx=15, pady=10)
bien_thuoc = tk.BooleanVar(value=False)
bien_loi_tat = tk.BooleanVar(value=True)
bien_giu_file = tk.BooleanVar(value=False)

tk.Checkbutton(khung_tuy_chon, text="💊 Kích hoạt Ohook (Gist Silent)", variable=bien_thuoc, font=phong_chu_dam, fg="#D84315", bg="#F4F6F8").pack(side="left", expand=True)
tk.Checkbutton(khung_tuy_chon, text="📌 Đưa Shortcut ra Desktop", variable=bien_loi_tat, font=phong_chu_dam, fg="#1565C0", bg="#F4F6F8").pack(side="left", expand=True)
tk.Checkbutton(khung_tuy_chon, text="💾 Giữ lại file nén (.zip) sau khi cài", variable=bien_giu_file, font=phong_chu_dam, bg="#F4F6F8").pack(side="left", expand=True)

khung_dieu_khien = tk.Frame(cua_so, bg="#E3F2FD", pady=10)
khung_dieu_khien.pack(fill="x", side="bottom")

khung_tien_do = tk.Frame(khung_dieu_khien, bg="#E3F2FD")
khung_tien_do.pack(side="left", fill="x", expand=True, padx=15)
nhan_trang_thai = tk.Label(khung_tien_do, text="Sẵn sàng...", font=phong_chu_dam, fg="#0277BD", bg="#E3F2FD")
nhan_trang_thai.pack(anchor="w")
thanh_tien_do = ttk.Progressbar(khung_tien_do, orient="horizontal", mode="determinate")
thanh_tien_do.pack(fill="x", pady=2)

khung_thong_so = tk.Frame(khung_dieu_khien, bg="#E3F2FD")
khung_thong_so.pack(side="left", padx=10)
nhan_toc_do = tk.Label(khung_thong_so, text="0 MB/s", font=phong_chu_dam, fg="#D84315", bg="#E3F2FD", width=10)
nhan_toc_do.pack(side="left")
nhan_thoi_gian = tk.Label(khung_thong_so, text="ETA: --:--", font=phong_chu_dam, fg="#2E7D32", bg="#E3F2FD", width=12)
nhan_thoi_gian.pack(side="left")
nhan_thong_tin = tk.Label(khung_thong_so, text="0/0 MB", font=phong_chu_phu, fg="#666666", bg="#E3F2FD", width=15)
nhan_thong_tin.pack(side="left")

def su_kien_bat_dau():
    muc_chon = danh_sach.selection()
    if not muc_chon:
        messagebox.showwarning("Cảnh báo", "Vui lòng chọn bản cài Office Google Drive!")
        return
        
    ds_chon = []
    for m in muc_chon:
        gia_tri = danh_sach.item(m, "values")
        for item in du_lieu_office:
            if item['Ten'] == gia_tri[0]:
                ds_chon.append(item)
                break
                
    global bien_lenh, bien_trang_thai, thong_so_phan_tram
    bien_lenh = "CHAY"
    bien_trang_thai = "Đang kết nối API Google..."
    thong_so_phan_tram = 0
    
    nut_bat_dau.config(state="disabled")
    nut_huy.config(state="normal")
    
    luong_xu_ly = threading.Thread(target=kich_ban_xu_ly, args=(ds_chon, hop_thu_muc.get(), bien_thuoc.get(), bien_giu_file.get(), bien_loi_tat.get()), daemon=True)
    luong_xu_ly.start()

def su_kien_huy():
    global bien_lenh, bien_trang_thai
    bien_lenh = "DUNG"
    bien_trang_thai = "🛑 Đang cắt mạng & dọn rác..."
    nut_huy.config(state="disabled")
    
    def don_rac_ngam():
        time.sleep(2)
        subprocess.run(["taskkill", "/F", "/IM", "7z.exe"], creationflags=subprocess.CREATE_NO_WINDOW, stderr=subprocess.DEVNULL)
        subprocess.run(["taskkill", "/F", "/IM", "setup.exe"], creationflags=subprocess.CREATE_NO_WINDOW, stderr=subprocess.DEVNULL)
        try: 
            if bien_thu_muc_giai_nen and os.path.exists(bien_thu_muc_giai_nen):
                subprocess.run(["cmd", "/c", "rmdir", "/s", "/q", bien_thu_muc_giai_nen], creationflags=subprocess.CREATE_NO_WINDOW)
        except: pass
        try:
            if bien_file_hien_tai and os.path.exists(bien_file_hien_tai): os.remove(bien_file_hien_tai)
        except: pass

    threading.Thread(target=don_rac_ngam, daemon=True).start()

nut_huy = tk.Button(khung_dieu_khien, text="🛑 HỦY BỎ", font=phong_chu_dam, bg="#FFCDD2", fg="#C62828", state="disabled", command=su_kien_huy)
nut_huy.pack(side="left", padx=5, ipady=5, ipadx=10)
nut_bat_dau = tk.Button(khung_dieu_khien, text="🚀 BẮT ĐẦU", font=("Segoe UI", 12, "bold"), bg="#0277BD", fg="white", command=su_kien_bat_dau)
nut_bat_dau.pack(side="left", padx=15, ipady=5, ipadx=15)

# ==============================================================================
# [MÔ-ĐUN 6] ĐỒNG BỘ GIAO DIỆN & TẢI CSV
# ==============================================================================
def cap_nhat_giao_dien():
    nhan_trang_thai.config(text=bien_trang_thai)
    thanh_tien_do["value"] = thong_so_phan_tram
    nhan_toc_do.config(text=thong_so_toc_do)
    nhan_thong_tin.config(text=thong_so_thong_tin)
    nhan_thoi_gian.config(text=f"ETA: {thong_so_thoi_gian}")
    
    van_ban_hien_tai = hop_nhat_ky.get("1.0", tk.END)
    if van_ban_hien_tai != bien_nhat_ky:
        hop_nhat_ky.delete("1.0", tk.END)
        hop_nhat_ky.insert(tk.END, bien_nhat_ky)
        hop_nhat_ky.see(tk.END)
        
    for m in danh_sach.get_children():
        ten_hang = danh_sach.item(m, "values")[0]
        for item in du_lieu_office:
            if item['Ten'] == ten_hang and item['ID'] in trang_thai_app:
                stt = trang_thai_app[item['ID']]
                danh_sach.item(m, values=(item['Ten'], stt["STT"], stt["PCT"], stt["SPD"], stt["DL"]))
                
    if "HOÀN TẤT" in bien_trang_thai or "ĐÃ HỦY" in bien_trang_thai:
        nut_bat_dau.config(state="normal")
        nut_huy.config(state="disabled")
        
    cua_so.after(300, cap_nhat_giao_dien)

def nap_danh_sach():
    global bien_trang_thai
    thu_muc_mac_dinh = "D:\\BoCaiOffice" if os.path.exists("D:\\") else "C:\\BoCaiOffice"
    hop_thu_muc.insert(0, thu_muc_mac_dinh)
    if not os.path.exists(thu_muc_mac_dinh): os.makedirs(thu_muc_mac_dinh)
    
    try:
        url_csv = f"https://raw.githubusercontent.com/tuantran19912512/Windows-tool-box/refs/heads/main/DanhSachOffice.csv?t={time.time()}"
        yeu_cau = urllib.request.urlopen(url_csv, timeout=10)
        noi_dung = yeu_cau.read().decode('utf-8').splitlines()
        trinh_doc = csv.DictReader(noi_dung)
        
        for dong in trinh_doc:
            if "drive" in dong['ID'] or "docs" in dong['ID'] or "http" not in dong['ID']:
                import re
                mau = re.search(r'id=([^&]+)|/d/([^/]+)', dong['ID'])
                if mau:
                    id_goc = mau.group(1) if mau.group(1) else mau.group(2)
                    du_lieu_office.append({"Ten": dong['Name'], "ID": id_goc})
                    danh_sach.insert("", "end", values=(dong['Name'], "Sẵn sàng", "", "", ""))
                    
        bien_trang_thai = "Sẵn sàng"
    except Exception as e:
        bien_trang_thai = "❌ Lỗi mạng"
        them_nhat_ky(f"❌ Lỗi nạp danh sách Google Drive: {e}")

cua_so.after(100, nap_danh_sach)
cua_so.after(300, cap_nhat_giao_dien)
cua_so.mainloop()