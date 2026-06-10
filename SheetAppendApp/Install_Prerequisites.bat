@echo off
:: ==========================================
:: SCRIPT CÀI ĐẶT MÔI TRƯỜNG CHO PHẦN MỀM GỘP EXCEL
:: Dành cho mọi phiên bản Windows (từ Win 7 đến Win 11)
:: Tác giả: Trợ lý AI (Google Deepmind)
:: ==========================================
chcp 65001 >nul
title Cài đặt môi trường bắt buộc (Prerequisites)

:: 1. Yêu cầu quyền Quản trị viên (Admin)
::--------------------------------------------------------
fsutil dirty query %systemdrive% >nul
if %errorLevel% NEQ 0 (
    echo Đang yeu cau quyen Quan tri vien ^(Administrator^)...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo ========================================================
echo   KIEM TRA VA CAI DAT .NET 6.0 DESKTOP RUNTIME
echo ========================================================
echo.

:: 2. Kiểm tra Windows 32-bit hay 64-bit
::--------------------------------------------------------
set ARCH=x64
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" (
    if not defined PROCESSOR_ARCHITEW6432 set ARCH=x86
)

:: 3. Kiểm tra xem .NET 6 đã cài chưa
::--------------------------------------------------------
echo Dang kiem tra .NET 6.0 Desktop Runtime...
dotnet --list-runtimes 2>nul | findstr /i "Microsoft.WindowsDesktop.App 6.0" >nul
if %errorLevel% EQU 0 (
    echo [OK] May tinh da duoc cai dat san .NET 6.0 Desktop Runtime!
    echo.
    pause
    exit /b
)

echo [!] May tinh chua co .NET 6.0 Desktop Runtime. Bat dau tai va cai dat tu dong...
echo.

:: 4. Tải xuống trình cài đặt từ Microsoft
::--------------------------------------------------------
set DOWNLOAD_URL=https://aka.ms/dotnet/6.0/windowsdesktop-runtime-win-%ARCH%.exe
set INSTALLER_NAME=%TEMP%\dotnet6_desktop_runtime_%ARCH%.exe

echo Dang tai xuong tu: %DOWNLOAD_URL%
echo Vui long doi trong it phut (tuy thuoc vao toc do mang)...

:: Sử dụng PowerShell WebClient để tải file (hỗ trợ tốt Win 7 SP1 trở lên, bỏ qua lỗi IE)
powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object System.Net.WebClient).DownloadFile('%DOWNLOAD_URL%', '%INSTALLER_NAME%')"

if not exist "%INSTALLER_NAME%" (
    echo [LOI] Khong the tai xuong file cai dat. Vui long kiem tra lai ket noi mang!
    pause
    exit /b
)

:: 5. Cài đặt tự động (Silent Install)
::--------------------------------------------------------
echo.
echo Da tai xong! Dang tien hanh cai dat ngam...
"%INSTALLER_NAME%" /install /quiet /norestart

if %errorLevel% EQU 0 (
    echo.
    echo ========================================================
    echo [THANH CONG] Da cai dat xong .NET 6.0 Desktop Runtime!
    echo Bay gio ban co the mo va su dung phan mem binh thuong.
    echo ========================================================
) else (
    echo.
    echo [CANH BAO] Cai dat co the chua hoan tat hoac can khoi dong lai may.
    echo Ma loi: %errorLevel%
    echo Vui long chay tay file cai dat tai: %INSTALLER_NAME%
)

:: 6. Dọn dẹp
::--------------------------------------------------------
del /f /q "%INSTALLER_NAME%" >nul 2>&1

echo.
pause
