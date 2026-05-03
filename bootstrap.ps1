# ==============================================================================
# KỊCH BẢN MỒI TỰ ĐỘNG (BOOTSTRAP SCRIPT) - CHỐNG LỖI MICROSOFT STORE
# ==============================================================================

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Vui lòng chạy lệnh này trong cửa sổ PowerShell dưới quyền Administrator!"
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit
}

$ErrorActionPreference = "SilentlyContinue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Write-Host "Đang kiểm tra môi trường hệ thống..." -ForegroundColor Cyan
$coPython = $false

# Kiểm tra Python và loại trừ file giả mạo của Microsoft Store (nằm trong thư mục WindowsApps)
$pyCmd = Get-Command "python" -ErrorAction SilentlyContinue
if ($pyCmd -and ($pyCmd.Source -notmatch "WindowsApps")) {
    $coPython = $true
    Write-Host "[Xong] Đã tìm thấy Python chuẩn trong hệ thống." -ForegroundColor Green
} else {
    Write-Host "[Chú ý] Máy chưa có Python (hoặc đang bị kẹt file ảo của Windows)."
    Write-Host "-> Đang tự động tải và cài đặt Python chuẩn (quá trình này mất khoảng 1 phút)..." -ForegroundColor Yellow
    
    $linkTaiPython = "https://www.python.org/ftp/python/3.12.3/python-3.12.3-amd64.exe"
    $fileCaiDat = Join-Path $env:TEMP "python_installer.exe"

    try {
        Invoke-WebRequest -Uri $linkTaiPython -OutFile $fileCaiDat -UseBasicParsing
        # Cài đặt ngầm (Silent), tự động add PATH và cài tcltk (bắt buộc cho giao diện tkinter)
        $tienTrinh = Start-Process -FilePath $fileCaiDat -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1 Include_test=0 Include_tcltk=1" -Wait -PassThru
        
        if ($tienTrinh.ExitCode -eq 0) {
            Write-Host "[Xong] Cài đặt Python thành công!" -ForegroundColor Green
            # Ép Windows nạp lại biến môi trường ngay lập tức
            $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
            $coPython = $true
        } else {
            Write-Host "[Lỗi] Cài đặt Python thất bại. Mã lỗi: $($tienTrinh.ExitCode)" -ForegroundColor Red
        }
    } catch {
        Write-Host "[Lỗi] Không thể tải bộ cài Python từ máy chủ." -ForegroundColor Red
    } finally {
        if (Test-Path $fileCaiDat) { Remove-Item $fileCaiDat -Force }
    }
}

# Tiến hành tải và chạy tool của bạn
if ($coPython) {
    Write-Host "Đang nạp Office Deploy từ GitHub..." -ForegroundColor Cyan
    
    $linkScriptPy = "https://raw.githubusercontent.com/tuantran19912512/pythonoffice/refs/heads/main/officedeploy.py"
    $filePyLuu = Join-Path $env:TEMP "OfficeDeploy_Master.py"

    try {
        Invoke-WebRequest -Uri $linkScriptPy -OutFile $filePyLuu -UseBasicParsing
        
        if (Test-Path $filePyLuu) {
            Write-Host "-> Bắt đầu khởi chạy Tool..." -ForegroundColor Green
            Write-Host "------------------------------------------------------"
            
            # Khởi chạy Python trực tiếp
            python $filePyLuu
            
            Write-Host "------------------------------------------------------"
            Write-Host "[Xong] Đã đóng ứng dụng." -ForegroundColor Cyan
        }
    } catch {
        Write-Host "[Lỗi] Không thể tải được file Python từ GitHub. Hãy kiểm tra lại link." -ForegroundColor Red
    } finally {
        if (Test-Path $filePyLuu) { Remove-Item $filePyLuu -Force }
    }
}

Write-Host "`nQuá trình hoàn tất. Ấn phím Enter để thoát..." -ForegroundColor DarkGray
Read-Host