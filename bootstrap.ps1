# ==============================================================================
# VIETTOOLBOX BOOTSTRAP V4.1 - GIAO DIỆN CỬA SỔ CÓ DẤU TIẾNG VIỆT
# ==============================================================================

# 1. Ép quyền Admin
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- KHỞI TẠO CỬA SỔ GIAO DIỆN ---
$form = New-Object Windows.Forms.Form
$form.Text = "VietToolbox - Chuẩn bị môi trường"
$form.Size = New-Object Drawing.Size(460, 190)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.TopMost = $true
$form.BackColor = [Drawing.Color]::FromArgb(245, 245, 245)

# Nhãn tiêu đề
$lblTitle = New-Object Windows.Forms.Label
$lblTitle.Text = "HỆ THỐNG ĐANG KIỂM TRA ĐIỀU KIỆN CÀI ĐẶT"
$lblTitle.Location = New-Object Drawing.Point(20, 20)
$lblTitle.Size = New-Object Drawing.Size(420, 25)
$lblTitle.Font = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
$form.Controls.Add($lblTitle)

# Nhãn trạng thái
$lblStatus = New-Object Windows.Forms.Label
$lblStatus.Text = "Vui lòng chờ trong giây lát..."
$lblStatus.Location = New-Object Drawing.Point(20, 50)
$lblStatus.Size = New-Object Drawing.Size(420, 20)
$lblStatus.Font = New-Object Drawing.Font("Segoe UI", 9)
$form.Controls.Add($lblStatus)

# Thanh Progress Bar
$progressBar = New-Object Windows.Forms.ProgressBar
$progressBar.Location = New-Object Drawing.Point(20, 85)
$progressBar.Size = New-Object Drawing.Size(400, 25)
$progressBar.Style = "Continuous"
$form.Controls.Add($progressBar)

# Hiển thị form
$form.Show()
$form.Focus()

# Hàm cập nhật giao diện
function Update-VTUI {
    param([string]$Status, [int]$Value)
    $lblStatus.Text = $Status
    $progressBar.Value = $Value
    $form.Refresh()
}

# --- BẮT ĐẦU QUY TRÌNH XỬ LÝ ---
$ErrorActionPreference = "SilentlyContinue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# 1. Kiểm tra Python
Update-VTUI -Status "Đang kiểm tra môi trường Python..." -Value 10
$pyCmd = Get-Command "python.exe" -ErrorAction SilentlyContinue

if ($pyCmd -and ($pyCmd.Source -notmatch "WindowsApps")) {
    Update-VTUI -Status "Đã tìm thấy Python chuẩn. Đang nạp Tool..." -Value 50
    Start-Sleep -Milliseconds 800
} else {
    Update-VTUI -Status "Không tìm thấy Python. Đang tải bộ cài từ Microsoft..." -Value 20
    $link = "https://www.python.org/ftp/python/3.12.3/python-3.12.3-amd64.exe"
    $path = Join-Path $env:TEMP "py_inst.exe"
    Invoke-WebRequest -Uri $link -OutFile $path -UseBasicParsing
    
    Update-VTUI -Status "Đang cài đặt Python (quá trình này mất khoảng 1 phút)..." -Value 40
    $p = Start-Process -FilePath $path -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1 Include_test=0 Include_tcltk=1" -Wait -PassThru
    
    if ($p.ExitCode -eq 0) {
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
        Update-VTUI -Status "Cài đặt Python thành công!" -Value 60
    }
    if (Test-Path $path) { Remove-Item $path -Force }
}

# 2. Tải Script từ GitHub
Update-VTUI -Status "Đang đồng bộ dữ liệu từ GitHub..." -Value 85
$url = "https://raw.githubusercontent.com/tuantran19912512/pythonoffice/refs/heads/main/officedeploy.py?t=$((Get-Date).Ticks)"
$out = Join-Path $env:TEMP "VT_Office.py"

try {
    Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing
    if (Test-Path $out) {
        Update-VTUI -Status "Mọi thứ đã sẵn sàng! Đang mở giao diện chính..." -Value 100
        Start-Sleep -Milliseconds 1200
        $form.Close()
        
        # Chạy tool chính
        & python.exe $out
    }
} catch {
    $lblStatus.ForeColor = [Drawing.Color]::Red
    Update-VTUI -Status "LỖI: Không thể kết nối đến GitHub!" -Value 0
    Start-Sleep -Seconds 5
    $form.Close()
} finally {
    if (Test-Path $out) { Remove-Item $out -Force }
}