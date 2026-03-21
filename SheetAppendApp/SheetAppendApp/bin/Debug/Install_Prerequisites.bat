@echo off
title Cai dat moi truong cho SheetAppendApp
color 0B

:: 1. Kiểm tra quyền Administrator
net session >nul 2>&1
if %errorLevel% == 0 (
    echo [OK] Da cap quyen Administrator.
) else (
    color 0C
    echo =======================================================
    echo LỖI: BAN CHUA CHAY BẰNG QUYỀN ADMINISTRATOR!
    echo =======================================================
    echo Vui long click chuot phai vao file bat nay
    echo va chon "Run as administrator".
    echo.
    pause
    exit /b 1
)

echo =======================================================
echo     CAI DAT MOI TRUONG CHO PHAN MEM MERGE EXCEL
echo =======================================================
echo.

:: 2. Link tải bản .NET 6.0 Desktop Runtime (x64) chính thức từ Microsoft
set DOTNET_URL=https://aka.ms/dotnet/6.0/windowsdesktop-runtime-win-x64.exe
set INSTALLER_NAME=dotnet_desktop_runtime_6_x64.exe

echo [1/3] Dang tai .NET 6 Desktop Runtime tu Microsoft...
:: Dùng curl (có sẵn trên Win 10/11) để tải file
curl -L -o %INSTALLER_NAME% %DOTNET_URL%

if not exist %INSTALLER_NAME% (
    color 0C
    echo [LỖI] Khong the tai file. Vui long kiem tra lai mang Internet!
    pause
    exit /b 1
)

:: 3. Cài đặt ngầm (Silent Install)
echo [2/3] Dang cai dat .NET 6 Desktop Runtime (Vui long doi 1-2 phut)...
start /wait %INSTALLER_NAME% /install /quiet /norestart

:: 4. Dọn dẹp
echo [3/3] Dang don dep file tam...
del %INSTALLER_NAME%

echo.
color 0A
echo =======================================================
echo HOAN TAT! May tinh da san sang de chay ung dung.
echo =======================================================
echo.

:: 5. Kích hoạt phần mềm để nó tự động thêm Menu Chuột Phải
set EXE_PATH=%~dp0SheetAppendApp.exe
if exist "%EXE_PATH%" (
    echo Dang khoi dong ung dung lan dau de dang ky Menu Chuot Phai...
    start "" "%EXE_PATH%"
) else (
    echo Khong tim thay SheetAppendApp.exe trong cung thu muc.
    echo Ban co the tu mo phan mem bang tay sau.
)

pause