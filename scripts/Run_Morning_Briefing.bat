@echo off
chcp 65001 >nul
title DUC_SmartOps - Chạy Bản Tin Đầu Ca Sáng
python "%~dp0..\src\scheduler_runner.py" --mode MORNING
pause
