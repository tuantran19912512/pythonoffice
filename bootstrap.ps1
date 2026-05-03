# ==============================================================================
# KỊCH BẢN MỒI TỰ ĐỘNG (BOOTSTRAP SCRIPT) - CÀI PYTHON VÀ CHẠY OFFICE DEPLOY
# ==============================================================================

# Yêu cầu quyền Quản trị viên (Admin) để có thể cài đặt phần mềm
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Vui lòng chạy lệnh này trong cửa sổ PowerShell dưới quyền Administrator!"
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

$ErrorActionPreference = "SilentlyContinue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# 1. KIỂM TRA VÀ CÀI ĐẶT PYTHON
Write-Host "Đang kiểm tra môi trường hệ thống..." -ForegroundColor Cyan
$coPython = $false

if (Get-Command "python" -ErrorAction SilentlyContinue) {
    $coPython = $true
    Write-Host "[Xong] Tuyệt vời! Máy tính đã có sẵn Python." -ForegroundColor Green
} else {
    Write-Host "[Chú ý] Chưa tìm thấy Python. Đang tiến hành tải và cài đặt tự động ngầm..." -ForegroundColor Yellow
    
    # Tải bản Python mới nhất và ổn định (Ví dụ: 3.12.3)
    $linkTaiPython = "https://www.python.org/ftp/python/3.12.3/python-3.12.3-amd64.exe"
    $fileCaiDat = Join-Path $env:TEMP "python_installer.exe"

    try {
        Invoke-WebRequest -Uri $linkTaiPython -OutFile $fileCaiDat -UseBasicParsing
        Write-Host "-> Đang cài đặt Python (có thể mất 1-2 phút). Vui lòng không tắt cửa sổ này..." -ForegroundColor Yellow
        
        # Cài đặt ngầm (Silent) và tự động thêm vào biến môi trường PATH
        $tienTrinh = Start-Process -FilePath $fileCaiDat -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1 Include_test=0" -Wait -PassThru
        
        if ($tienTrinh.ExitCode -eq 0) {
            Write-Host "[Xong] Cài đặt Python thành công!" -ForegroundColor Green
            # Làm mới biến môi trường ngay trong phiên làm việc hiện tại
            $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
            $coPython = $true
        } else {
            Write-Host "[Lỗi] Có sự cố xảy ra khi cài đặt Python (Mã lỗi: $($tienTrinh.ExitCode))." -ForegroundColor Red
        }
    } catch {
        Write-Host "[Lỗi] Không thể tải bộ cài Python. Hãy kiểm tra lại kết nối mạng." -ForegroundColor Red
    } finally {
        if (Test-Path $fileCaiDat) { Remove-Item $fileCaiDat -Force }
    }
}

# 2. TẢI VÀ CHẠY ỨNG DỤNG OFFICE DEPLOY
if ($coPython) {
    Write-Host "Đang nạp ứng dụng Office Deploy từ hệ thống máy chủ..." -ForegroundColor Cyan
    
    # BẠN HÃY THAY ĐƯỜNG DẪN RAW GITHUB CỦA FILE PYTHON VÀO ĐÂY
    $linkScriptPy = "https://raw.githubusercontent.com/tuantran19912512/pythonoffice/refs/heads/main/officedeploy.py"
    $filePyLuu = Join-Path $env:TEMP "OfficeDeploy_Master.py"

    try {
        Invoke-WebRequest -Uri $linkScriptPy -OutFile $filePyLuu -UseBasicParsing
        Write-Host "-> Đang khởi chạy giao diện..." -ForegroundColor Green
        
        # Gọi Python để chạy file vừa tải về (Wait để chờ ứng dụng đóng mới đi tiếp)
        Start-Process -FilePath "python" -ArgumentList "`"$filePyLuu`"" -Wait
        
        Write-Host "[Xong] Đã đóng ứng dụng." -ForegroundColor Cyan
    } catch {
        Write-Host "[Lỗi] Không thể tải ứng dụng. Vui lòng kiểm tra lại đường dẫn GitHub." -ForegroundColor Red
    } finally {
        # Dọn dẹp rác sau khi khách hàng tắt tool
        if (Test-Path $filePyLuu) { Remove-Item $filePyLuu -Force }
    }
} else {
    Write-Host "[Thất bại] Hệ thống không đủ điều kiện để chạy phần mềm." -ForegroundColor Red
}