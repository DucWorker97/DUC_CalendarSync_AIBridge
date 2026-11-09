@echo off
chcp 65001 >nul
title DUC_SmartOps - Kiểm Tra Trạng Thái Lịch Windows Task Scheduler
python "%~dp0..\src\scheduler_runner.py" --status-tasks
pause
