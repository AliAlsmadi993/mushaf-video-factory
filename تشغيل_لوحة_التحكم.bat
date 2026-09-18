@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
echo جاري تشغيل لوحة تحكم مُصحَف...
start "" http://127.0.0.1:5055
py server.py
pause
