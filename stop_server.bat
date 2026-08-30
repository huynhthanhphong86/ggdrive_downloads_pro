@echo off
title Dung May Chu - Google Drive Downloads Pro
cd /d "%~dp0"

echo ========================================================
echo   DUNG MAY CHU GOOGLE DRIVE DOWNLOADS PRO (PORT 5000)
echo ========================================================
echo.
echo Dang tim va dong cac tien trinh tren cong 5000...

set FOUND=0
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :5000 ^| findstr LISTENING') do (
    set FOUND=1
    echo  -> Dang tat tien trinh (PID: %%a)...
    taskkill /F /PID %%a >nul 2>&1
)

if "%FOUND%"=="1" (
    echo.
    echo ========================================================
    echo   [THANH CONG] Da dung may chu thanh cong!
    echo ========================================================
) else (
    echo.
    echo [THONG BAO] Khong tim thay tien trinh nao dang chay tren cong 5000.
)

echo.
timeout /t 3
