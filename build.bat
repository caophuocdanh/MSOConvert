@echo off
title Build Excel Converter Tool
color 0A
cls

:: ============================================================
:: CẤU HÌNH (Sửa tên file tại đây nếu cần)
:: ============================================================
set "SCRIPT_NAME=xls2xlsx.py"
set "EXE_NAME=xls2xlsx"
:: ============================================================

echo.
echo ========================================================
echo   TOOL DONG GOI PYTHON SANG EXE (AUTO BUILD)
echo ========================================================
echo.

:: Dọn dẹp file cũ
echo.
echo Don dep file tam cu (build folder, spec file)...
if exist "build" rmdir /s /q "build"
if exist "%EXE_NAME%.spec" del /q "%EXE_NAME%.spec"
if exist "__pycache__" rmdir /s /q "__pycache__"
if exist "dist" rmdir /s /q "dist"

:: 1. Kiểm tra xem file Python có tồn tại không
if not exist "%SCRIPT_NAME%" (
    color 0C
    echo [LOI] Khong tim thay file "%SCRIPT_NAME%"!
    echo Vui long doi ten file Python cua ban thanh "%SCRIPT_NAME%"
    echo hoac sua lai ten trong file build.bat nay.
    echo.
    pause
    exit /b
)

:: 2. Cài đặt các thư viện cần thiết (nếu chưa có)
echo [BUOC 1/3] Kiem tra va cai dat thu vien (PyInstaller, PyWin32)...
pip install pyinstaller pywin32 --upgrade --quiet
if %errorlevel% neq 0 (
    color 0C
    echo [LOI] Khong the cai dat thu vien. Kiem tra lai Python/PIP.
    pause
    exit /b
)


:: 3. Tiến hành Build
echo.
echo [BUOC 2/3] Dang dong goi file EXE... (Vui long doi)
echo.

:: Lệnh PyInstaller với các tham số tối ưu cho PyWin32
pyinstaller --noconsole --onefile --clean ^
    --icon=icon.ico ^
    --name "%EXE_NAME%" ^
    --hidden-import="win32com.client" ^
    --hidden-import="pythoncom" ^
    "%SCRIPT_NAME%"

if %errorlevel% neq 0 (
    color 0C
    echo.
    echo [LOI] Qua trinh Build that bai!
    pause
    exit /b
)

:: 4. Dọn dẹp file rác
echo.
echo [BUOC 3/3] Don dep file tam (build folder, spec file)...
if exist "build" rmdir /s /q "build"
if exist "%EXE_NAME%.spec" del /q "%EXE_NAME%.spec"
if exist "__pycache__" rmdir /s /q "__pycache__"

echo.
echo ========================================================
echo   HOAN TAT!
echo   File EXE cua ban nam trong thu muc "dist"
echo ========================================================
echo.
pause