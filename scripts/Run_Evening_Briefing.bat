@echo off
chcp 65001 >nul
title DUC_SmartOps - Chạy Bản Tin Cuối Ca Chiều & Quét ITSM
python "%~dp0..\src\scheduler_runner.py" --mode EVENING
pause
