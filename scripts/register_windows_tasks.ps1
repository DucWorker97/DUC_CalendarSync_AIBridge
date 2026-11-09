# ==============================================================================
# register_windows_tasks.ps1 — Đăng Ký Lịch Vận Hành Vào Windows Task Scheduler
# ==============================================================================
# DUC_CalendarSync_AIBridge (v1.3.0) — Tác giả: Nguyễn Duy Đức (DucWorker97)
# ==============================================================================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BridgeRoot = Split-Path -Parent $ScriptDir
$RunnerPy = Join-Path $BridgeRoot "src\scheduler_runner.py"

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "📅 DUC_SmartOps — Đăng Ký Windows Task Scheduler (Lịch Quét Tự Động)" -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan

# Tim pythonw.exe
$PythonCmd = (Get-Command python -ErrorAction SilentlyContinue).Source
if (-not $PythonCmd) {
    Write-Host "❌ Không tìm thấy Python trong hệ thống PATH!" -ForegroundColor Red
    exit 1
}

$PyDir = Split-Path -Parent $PythonCmd
$PythonW = Join-Path $PyDir "pythonw.exe"
if (-not (Test-Path $PythonW)) {
    $PythonW = $PythonCmd
}

Write-Host "🐍 PythonW Executable : $PythonW" -ForegroundColor Yellow
Write-Host "📜 Scheduler Runner   : $RunnerPy" -ForegroundColor Yellow
Write-Host ""

& $PythonCmd "$RunnerPy" --register-tasks

Write-Host ""
Write-Host "✅ Hoàn tất kiểm tra trạng thái nhiệm vụ:" -ForegroundColor Green
& $PythonCmd "$RunnerPy" --status-tasks
