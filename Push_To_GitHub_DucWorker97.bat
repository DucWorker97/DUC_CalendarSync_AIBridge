@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion
title Push DUC_CalendarSync_AIBridge to GitHub (Account: DucWorker97)

cd /d "%~dp0"

echo ==============================================================================
echo   🚀 ĐẨY DỰ ÁN MCP LÊN TÀI KHOẢN GITHUB: DucWorker97
echo ==============================================================================
echo.
echo   Gợi ý tên Repository trên tài khoản DucWorker97:
echo     [1] DUC_CalendarSync_AIBridge  (Mặc định - Đồng bộ chuẩn DUC_*_AIBridge)
echo     [2] DUC_ChronoOps_MCP          (Phong cách hiện đại Chrono + SmartOps)
echo     [3] DUC_SmartCalendar_MCP      (Trực diện, ngắn gọn)
echo.
set /p repo_name="Nhập tên Repo trên GitHub DucWorker97 (Bấm Enter để chọn DUC_CalendarSync_AIBridge): "
if "%repo_name%"=="" set repo_name=DUC_CalendarSync_AIBridge

echo.
echo [1/4] Khởi tạo Git Repository & Cấu hình Author = DucWorker97...
if not exist ".git" (
    git init
    git branch -M main
)
git config user.name "DucWorker97"
git config user.email "ducworker97@users.noreply.github.com"

echo [2/4] Đóng gói toàn bộ mã nguồn (git add & commit)...
git add -A
git commit -m "feat(init): Initial release of %repo_name% v1.0.0 (Multi-Platform Calendar Sync & ITSM Smart Reminder MCP)"

echo [3/4] Thiết lập Remote URL: https://github.com/DucWorker97/%repo_name%.git ...
git remote remove origin >nul 2>&1
git remote add origin https://github.com/DucWorker97/%repo_name%.git

echo [4/4] Đẩy lên GitHub (Nếu đã cài gh CLI sẽ tự tạo repo mới nếu chưa có)...
where gh >nul 2>&1
if %errorlevel%==0 (
    gh repo view DucWorker97/%repo_name% >nul 2>&1
    if not !errorlevel!==0 (
        echo Đang tạo Repository mới DucWorker97/%repo_name% qua GitHub CLI...
        gh repo create DucWorker97/%repo_name% --public --source=. --remote=origin --push
        goto done
    )
)

git push -u origin main

:done
echo.
echo ==============================================================================
echo   ✅ HOÀN TẤT! Sau khi đẩy xong trên máy DucWorker97, trên máy chính anh chạy:
echo      cd /d D:\DucWorker\DUC_SmartOps_AIHub\mcp_bridges
echo      git clone https://github.com/DucWorker97/%repo_name%.git
echo ==============================================================================
pause
