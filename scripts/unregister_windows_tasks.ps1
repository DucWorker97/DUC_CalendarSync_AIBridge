# ==============================================================================
# unregister_windows_tasks.ps1 — Hủy Đăng Ký Lịch Khỏi Windows Task Scheduler
# ==============================================================================
# DUC_CalendarSync_AIBridge (v1.3.0) — Tác giả: Nguyễn Duy Đức (DucWorker97)
# ==============================================================================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BridgeRoot = Split-Path -Parent $ScriptDir
$RunnerPy = Join-Path $BridgeRoot "src\scheduler_runner.py"

Write-Host "======================================================================" -ForegroundColor Yellow
Write-Host "🗑️ DUC_SmartOps — Hủy Đăng Ký Windows Task Scheduler" -ForegroundColor Yellow
Write-Host "======================================================================" -ForegroundColor Yellow

$PythonCmd = (Get-Command python -ErrorAction SilentlyContinue).Source
if (-not $PythonCmd) {
    Write-Host "❌ Không tìm thấy Python trong hệ thống PATH!" -ForegroundColor Red
    exit 1
}

& $PythonCmd "$RunnerPy" --unregister-tasks

Write-Host ""
Write-Host "Trạng thái sau khi hủy:" -ForegroundColor Cyan
& $PythonCmd "$RunnerPy" --status-tasks
