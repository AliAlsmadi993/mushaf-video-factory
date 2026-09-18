@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
echo ========================================
echo   مُصحَف - تثبيت المتطلبات
 echo ========================================
where py >nul 2>nul
if %errorlevel% neq 0 (
  echo لم يتم العثور على Python Launcher ^(py^). ثبّت Python 3.10 أو أحدث ثم أعد المحاولة.
  pause
  exit /b 1
)
py -m pip install --upgrade pip
py -m pip install -r requirements.txt
if %errorlevel% neq 0 (
  echo فشل تثبيت بعض المتطلبات.
  pause
  exit /b 1
)
if not exist "خلفيات" mkdir "خلفيات"
if not exist "output" mkdir "output"
if not exist "temp" mkdir "temp"
py build_full_quran_rukus.py
echo.
echo اكتمل التثبيت. شغّل الملف تشغيل_لوحة_التحكم.bat
pause
