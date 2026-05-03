# ==============================================================================
# KỊCH BẢN MỒI TỰ ĐỘNG (BOOTSTRAP SCRIPT) - CẬP NHẬT GỠ LỖI
# ==============================================================================

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Vui lòng chạy lệnh này trong cửa sổ PowerShell dưới quyền Administrator!"
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit
}

$ErrorActionPreference = "SilentlyContinue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Write-Host "Đang kiểm tra môi trường hệ thống..." -ForegroundColor Cyan
$coPython = $false

# Cố gắng tìm Python qua biến môi trường hoặc file thực thi
if (Get-Command "python.exe" -ErrorAction SilentlyContinue) {
    $coPython = $true
    Write-Host "[Xong] Đã tìm thấy Python trong hệ thống." -ForegroundColor Green
} else {
    Write-Host "[Chú ý] Chưa có Python. Đang tự động tải và cài đặt (khoảng 1 phút)..." -ForegroundColor Yellow
    $linkTaiPython = "https://www.python.org/ftp/python/3.12.3/python-3.12.3-amd64.exe"
    $fileCaiDat = Join-Path $env:TEMP "python_installer.exe"

    Invoke-WebRequest -Uri $linkTaiPython -OutFile $fileCaiDat -UseBasicParsing
    $tienTrinh = Start-Process -FilePath $fileCaiDat -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1 Include_test=0 Include_tcltk=1" -Wait -PassThru
    
    if ($tienTrinh.ExitCode -eq 0) {
        Write-Host "[Xong] Cài đặt Python thành công!" -ForegroundColor Green
        # Ép Windows nạp lại biến môi trường
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
        $coPython = $true
    } else {
        Write-Host "[Lỗi] Cài đặt Python thất bại. Mã lỗi: $($tienTrinh.ExitCode)" -ForegroundColor Red
    }
}

if ($coPython) {
    Write-Host "Đang nạp Office Deploy từ GitHub..." -ForegroundColor Cyan
    
    # LINK RAW CỦA BẠN ĐÃ ĐƯỢC CẬP NHẬT VÀO ĐÂY
    $linkScriptPy = "https://raw.githubusercontent.com/tuantran19912512/pythonoffice/refs/heads/main/officedeploy.py"
    $filePyLuu = Join-Path $env:TEMP "OfficeDeploy_Master.py"

    Invoke-WebRequest -Uri $linkScriptPy -OutFile $filePyLuu -UseBasicParsing
    
    if (Test-Path $filePyLuu) {
        Write-Host "-> Bắt đầu khởi chạy Tool..." -ForegroundColor Green
        Write-Host "------------------------------------------------------"
        
        # Chạy trực tiếp để BẮT LỖI nếu file Python crash
        try {
            & python.exe $filePyLuu
            if ($LASTEXITCODE -ne 0) {
                Write-Host "`n[LỖI PYTHON] Ứng dụng Python đã bị crash! Vui lòng đọc thông báo lỗi bằng tiếng Anh ở phía trên." -ForegroundColor Red
            }
        } catch {
            Write-Host "`n[LỖI GỌI PYTHON] Không thể gọi lệnh 'python'. Hãy thử tắt PowerShell mở lại, hoặc khởi động lại máy tính!" -ForegroundColor Red
        }
        
        Write-Host "------------------------------------------------------"
    } else {
        Write-Host "[Lỗi] Không thể tải được file Python từ GitHub. Hãy kiểm tra lại link." -ForegroundColor Red
    }
}

Write-Host "Ấn phím Enter để thoát..." -ForegroundColor Cyan
Read-Host