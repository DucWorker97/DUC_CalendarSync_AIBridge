@echo off
chcp 65001 >nul
title DUC_SmartOps - Hủy Đăng Ký Windows Task Scheduler
echo ======================================================================
echo 🗑️ Đang hủy đăng ký các tác vụ DUC_SmartOps khỏi Windows Task Scheduler...
echo ======================================================================
python "%~dp0..\src\scheduler_runner.py" --unregister-tasks
echo.
python "%~dp0..\src\scheduler_runner.py" --status-tasks
echo ======================================================================
pause
