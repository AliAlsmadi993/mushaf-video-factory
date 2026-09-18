@echo off
chcp 65001 >nul
cd /d "%~dp0"
py generator.py --demo --output output\demo.mp4
pause
