@echo off
chcp 65001 >nul
title DUC_SmartOps - Đăng Ký Windows Task Scheduler
echo ======================================================================
echo 📅 Đang đăng ký Lịch Quét Tự Động Hàng Ngày vào Windows Task Scheduler...
echo 🌅 Ca Sáng: 08:00 Thứ 2 - Thứ 7 (Bản tin Sáng)
echo 🌇 Ca Chiều: 17:15 Thứ 2 - Thứ 7 (Bản tin Chiều & Bù đơn ITSM)
echo ======================================================================
python "%~dp0..\src\scheduler_runner.py" --register-tasks
echo.
echo ======================================================================
echo Trạng thái nhiệm vụ hiện tại:
python "%~dp0..\src\scheduler_runner.py" --status-tasks
echo ======================================================================
pause
