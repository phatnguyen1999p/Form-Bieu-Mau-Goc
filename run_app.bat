@echo off
chcp 65001 >nul
title KHOI DONG HE THONG NHAN BAN BIEU MAU QUATEST 3
echo ======================================================================
echo    HE THONG NHAN BAN VA TAI CHE BIEU MAU HIEN TRUONG QUATEST 3
echo ======================================================================
echo Dang khoi dong giao dien nguoi dung, vui long cho...

python app_gui.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Co loi khi khoi dong giao dien. Vui long kiem tra Python da duoc cai dat.
    pause
)
