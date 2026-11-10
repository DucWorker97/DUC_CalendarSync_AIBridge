# -*- coding: utf-8 -*-
"""
scheduler_runner.py — Trình Thực Thi Lịch Quét & Vận Hành Tự Động Hàng Ngày (v1.3.0)
===================================================================================
Chịu trách nhiệm:
1. Bản tin Đầu ca Sáng (MORNING Briefing — 08:00 Thứ 2 – Thứ 7):
   - Đọc lịch trình 2 chiều từ Google Calendar & Outlook Desktop & Kho .ics.
   - Kiểm tra các đơn ITSM cần bù trong ngày và câu hỏi Mã NV (Quy tắc 2 & 4).
   - Xuất tệp Markdown Bản tin lưu vào exports/briefings/ và Hub logs/daily_briefings/.
2. Bản tin Cuối ca Chiều (EVENING Briefing — 17:30 Thứ 2 – Thứ 7):
   - Quét sổ cái THEO-DOI-BU-TICKET-ITSM.md: tự động đặt lịch nhắc bù đơn Bxx.
   - Đồng bộ tức thì 3 chuỗi lịch định kỳ và lịch nhắc nhở (cal_sync_all).
   - Xuất Bản tin tổng kết ca chiều và kích hoạt cảnh báo chuẩn bị Sao lưu 17:45 (Quy tắc 12).
3. Tích hợp Quản lý Windows Task Scheduler:
   - Tự động đăng ký 2 Tasks (Morning 08:00 & Evening 17:30) vào Windows Task Scheduler.
   - Chạy ngầm hoàn toàn qua pythonw.exe (Windowless, không hiện popup CMD làm gián đoạn người dùng).
===================================================================================
"""

import argparse
import csv
import io
import json
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

_SRC_DIR = Path(__file__).resolve().parent
_REPO_ROOT = _SRC_DIR.parent
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from src.calendar_engine import (
    engine_daily_briefing,
    engine_remind_pending_itsm,
    engine_sync_all,
)
from src.notification_dispatcher import (
    dispatch_daily_briefing,
    load_notification_config,
)

TASK_NAME_MORNING = r"DUC_SmartOps\Morning_Briefing_0800"
TASK_NAME_EVENING = r"DUC_SmartOps\Evening_Briefing_ITSM_1730"
OLD_TASK_NAME_EVENING = r"DUC_SmartOps\Evening_Briefing_ITSM_1715"

DEFAULT_HUB_ROOT = Path(r"D:\DucWorker\DUC_SmartOps_AIHub")


def get_python_executables() -> Dict[str, str]:
    """Xác định đường dẫn python.exe và pythonw.exe."""
    py_exe = Path(sys.executable).resolve()
    pyw_exe = py_exe.with_name("pythonw.exe")
    if not pyw_exe.is_file():
        pyw_exe = py_exe
    return {
        "python": str(py_exe),
        "pythonw": str(pyw_exe),
    }


def _get_briefings_export_dir(repo_root: Optional[Path] = None) -> Path:
    root = repo_root or _REPO_ROOT
    bdir = root / "exports" / "briefings"
    bdir.mkdir(parents=True, exist_ok=True)
    return bdir


def _get_hub_briefings_dir() -> Optional[Path]:
    if DEFAULT_HUB_ROOT.is_dir():
        hdir = DEFAULT_HUB_ROOT / "logs" / "daily_briefings"
        hdir.mkdir(parents=True, exist_ok=True)
        return hdir
    return None


def _log_scheduler_history(entry: Dict[str, Any], repo_root: Optional[Path] = None) -> None:
    root = repo_root or _REPO_ROOT
    log_dir = root / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "scheduler_history.jsonl"
    entry_line = json.dumps(entry, ensure_ascii=False, default=str)
    with log_file.open("a", encoding="utf-8") as f:
        f.write(entry_line + "\n")

    # Mirror vao Hub neu ton tai
    if DEFAULT_HUB_ROOT.is_dir():
        hub_audit = DEFAULT_HUB_ROOT / "logs" / "ops_audit.jsonl"
        try:
            hub_entry = {
                "at": entry.get("timestamp", datetime.now().isoformat(timespec="seconds")),
                "source": "calendar_scheduler",
                "action": entry.get("action", "SCHEDULED_RUN"),
                "ok": entry.get("status") in {"SUCCESS", "DRY_RUN", "ALL_CLEAR_ZERO_PENDING"},
                "details": {
                    "mode": entry.get("mode"),
                    "exported_files": entry.get("exported_files", []),
                },
            }
            with hub_audit.open("a", encoding="utf-8") as f:
                f.write(json.dumps(hub_entry, ensure_ascii=False) + "\n")
        except Exception:
            pass


def execute_morning_routine(
    date_iso: str = "",
    dry_run: bool = False,
    export_md: bool = True,
    repo_root: Optional[Path] = None,
    output_dir: Optional[Path] = None,
    dispatch_channels: Optional[Sequence[str]] = None,
    dispatch_dry_run: bool = True,
) -> Dict[str, Any]:
    """Thực thi Bản tin Đầu ca Sáng (08:00) kèm tùy chọn thông báo đa kênh."""
    now = datetime.fromisoformat(date_iso.strip()) if date_iso.strip() else datetime.now()
    date_tag = now.strftime("%Y%m%d")

    briefing = engine_daily_briefing(
        date_iso=date_iso,
        shift_mode="MORNING",
        include_live_sources=not dry_run,
        base_dir=repo_root or _REPO_ROOT,
    )

    exported_files: List[str] = []
    if export_md:
        md_text = briefing.get("markdown_briefing", "")
        # Luu trong bridge
        target_dir = output_dir or _get_briefings_export_dir(repo_root)
        target_dir.mkdir(parents=True, exist_ok=True)
        fname = f"Briefing_{date_tag}_MORNING.md"
        out_path = target_dir / fname
        out_path.write_text(md_text, encoding="utf-8")
        exported_files.append(str(out_path))

        # Luu trong Hub
        hub_dir = _get_hub_briefings_dir()
        if hub_dir:
            hub_path = hub_dir / fname
            hub_path.write_text(md_text, encoding="utf-8")
            exported_files.append(str(hub_path))

    dispatch_res: Dict[str, Any] = {}
    if dispatch_channels:
        dispatch_res = dispatch_daily_briefing(
            date_iso=date_iso,
            shift_mode="MORNING",
            channels=dispatch_channels,
            dry_run=dispatch_dry_run,
            repo_root=repo_root or _REPO_ROOT,
        )

    result = {
        "status": "DRY_RUN" if dry_run else "SUCCESS",
        "action": "MORNING_BRIEFING",
        "mode": "MORNING",
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "date": now.strftime("%Y-%m-%d"),
        "briefing_title": briefing.get("shift_label"),
        "today_events_count": briefing.get("today_events_count", 0),
        "itsm_pending_count": briefing.get("itsm_pending_count", 0),
        "exported_files": exported_files,
        "dispatch_results": dispatch_res.get("channels", {}),
        "briefing": briefing,
    }
    _log_scheduler_history(result, repo_root=repo_root)
    return result


def execute_evening_routine(
    date_iso: str = "",
    auto_remind_itsm: bool = False,  # 10/11/2026: Duc tat tu tao lich nhac Outlook/Google; bat lai bang --remind-itsm
    dry_run: bool = False,
    export_md: bool = True,
    repo_root: Optional[Path] = None,
    output_dir: Optional[Path] = None,
    dispatch_channels: Optional[Sequence[str]] = None,
    dispatch_dry_run: bool = True,
) -> Dict[str, Any]:
    """Thực thi Bản tin Cuối ca Chiều (17:30) kèm Quét Bù Đơn ITSM, Đồng Bộ Lịch & Thông Báo Đa Kênh."""
    now = datetime.fromisoformat(date_iso.strip()) if date_iso.strip() else datetime.now()
    date_tag = now.strftime("%Y%m%d")

    remind_res: Dict[str, Any] = {}
    if auto_remind_itsm:
        remind_res = engine_remind_pending_itsm(
            dry_run=dry_run,
            base_dir=repo_root or _REPO_ROOT,
        )

    sync_res = engine_sync_all(
        dry_run=dry_run,
        base_dir=repo_root or _REPO_ROOT,
    )

    briefing = engine_daily_briefing(
        date_iso=date_iso,
        shift_mode="EVENING",
        include_live_sources=not dry_run,
        base_dir=repo_root or _REPO_ROOT,
    )

    exported_files: List[str] = []
    if export_md:
        md_text = briefing.get("markdown_briefing", "")
        # Luu trong bridge
        target_dir = output_dir or _get_briefings_export_dir(repo_root)
        target_dir.mkdir(parents=True, exist_ok=True)
        fname = f"Briefing_{date_tag}_EVENING.md"
        out_path = target_dir / fname
        out_path.write_text(md_text, encoding="utf-8")
        exported_files.append(str(out_path))

        # Luu trong Hub
        hub_dir = _get_hub_briefings_dir()
        if hub_dir:
            hub_path = hub_dir / fname
            hub_path.write_text(md_text, encoding="utf-8")
            exported_files.append(str(hub_path))

    dispatch_res: Dict[str, Any] = {}
    if dispatch_channels:
        dispatch_res = dispatch_daily_briefing(
            date_iso=date_iso,
            shift_mode="EVENING",
            channels=dispatch_channels,
            dry_run=dispatch_dry_run,
            repo_root=repo_root or _REPO_ROOT,
        )

    result = {
        "status": "DRY_RUN" if dry_run else "SUCCESS",
        "action": "EVENING_BRIEFING_AND_ITSM_CHECK",
        "mode": "EVENING",
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "date": now.strftime("%Y-%m-%d"),
        "briefing_title": briefing.get("shift_label"),
        "itsm_reminder_result": remind_res.get("status", "SKIPPED"),
        "calendar_sync_result": sync_res.get("status", "SUCCESS"),
        "itsm_pending_count": briefing.get("itsm_pending_count", 0),
        "exported_files": exported_files,
        "dispatch_results": dispatch_res.get("channels", {}),
        "briefing": briefing,
    }
    _log_scheduler_history(result, repo_root=repo_root)
    return result


def execute_sync_routine(dry_run: bool = False, repo_root: Optional[Path] = None) -> Dict[str, Any]:
    """Thực thi Đồng bộ toàn diện (cal_sync_all)."""
    sync_res = engine_sync_all(dry_run=dry_run, base_dir=repo_root or _REPO_ROOT)
    result = {
        "status": "DRY_RUN" if dry_run else "SUCCESS",
        "action": "CALENDAR_SYNC_ALL",
        "mode": "SYNC",
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "sync_details": sync_res,
    }
    _log_scheduler_history(result, repo_root=repo_root)
    return result


def execute_auto_routine(
    date_iso: str = "",
    dry_run: bool = False,
    export_md: bool = True,
    repo_root: Optional[Path] = None,
) -> Dict[str, Any]:
    """Tự động lựa chọn MORNING nếu trước 14:00, ngược lại chọn EVENING."""
    now = datetime.fromisoformat(date_iso.strip()) if date_iso.strip() else datetime.now()
    if now.hour < 14:
        return execute_morning_routine(date_iso=date_iso, dry_run=dry_run, export_md=export_md, repo_root=repo_root)
    else:
        return execute_evening_routine(date_iso=date_iso, auto_remind_itsm=False, dry_run=dry_run, export_md=export_md, repo_root=repo_root)


def _query_single_task(task_name: str) -> Dict[str, Any]:
    """Truy vấn thông tin 1 tác vụ trong Windows Task Scheduler qua CSV."""
    if os.name != "nt":
        return {"exists": False, "status": "NOT_WINDOWS"}
    try:
        proc = subprocess.run(
            ["schtasks", "/Query", "/TN", task_name, "/FO", "CSV"],
            capture_output=True,
            text=True,
            check=False,
        )
        if proc.returncode != 0:
            return {"exists": False, "status": "NOT_FOUND"}
        reader = csv.DictReader(io.StringIO(proc.stdout.strip()))
        rows = list(reader)
        if rows:
            r = rows[0]
            return {
                "exists": True,
                "task_name": r.get("TaskName", task_name),
                "next_run_time": r.get("Next Run Time", ""),
                "status": r.get("Status", "Ready"),
            }
        return {"exists": True, "status": "Ready"}
    except Exception as e:
        return {"exists": False, "error": str(e)}


def get_windows_tasks_status() -> Dict[str, Any]:
    """Kiểm tra trạng thái của các tác vụ DUC_SmartOps trong Windows Task Scheduler."""
    if os.name != "nt":
        return {
            "supported": False,
            "registered": False,
            "message": "Windows Task Scheduler chỉ hoạt động trên hệ điều hành Windows.",
        }

    morning_st = _query_single_task(TASK_NAME_MORNING)
    evening_st = _query_single_task(TASK_NAME_EVENING)
    both_registered = bool(morning_st.get("exists") and evening_st.get("exists"))

    return {
        "supported": True,
        "registered": both_registered,
        "checked_at": datetime.now().isoformat(timespec="seconds"),
        "tasks": {
            "morning_briefing": {
                "task_name": TASK_NAME_MORNING,
                "schedule": "08:00 (Thứ 2 – Thứ 7)",
                "exists": morning_st.get("exists", False),
                "next_run_time": morning_st.get("next_run_time", ""),
                "status": morning_st.get("status", "NOT_FOUND"),
            },
            "evening_briefing": {
                "task_name": TASK_NAME_EVENING,
                "schedule": "17:30 (Thứ 2 – Thứ 7)",
                "exists": evening_st.get("exists", False),
                "next_run_time": evening_st.get("next_run_time", ""),
                "status": evening_st.get("status", "NOT_FOUND"),
            },
        },
    }


def register_windows_tasks(python_exe: Optional[str] = None) -> Dict[str, Any]:
    """
    Đăng ký tự động 2 tác vụ vào Windows Task Scheduler:
    1. DUC_SmartOps\\Morning_Briefing_0800 (08:00 Thứ 2 - Thứ 7)
    2. DUC_SmartOps\\Evening_Briefing_ITSM_1730 (17:30 Thứ 2 - Thứ 7)
    Sử dụng pythonw.exe để chạy ngầm hoàn toàn không hiện popup console.
    Mỗi tác vụ tự động phân phối bản tin Zalo vào nhóm Thông báo chung OLAZ & xuất file Markdown.
    """
    if os.name != "nt":
        return {"status": "ERROR", "message": "Chỉ hỗ trợ trên hệ điều hành Windows."}

    exes = get_python_executables()
    pyw = python_exe or exes["pythonw"]
    runner_script = str(Path(__file__).resolve())

    # Xóa task cũ 17:15 nếu còn tồn tại
    subprocess.run(["schtasks", "/Delete", "/TN", OLD_TASK_NAME_EVENING, "/F"], capture_output=True, text=True, check=False)

    # Task 1: Morning Briefing (08:00) — Bắn Zalo & File Markdown
    tr_m = f'"{pyw}" "{runner_script}" --mode MORNING --dispatch --confirm-send --channels file,zalo'
    cmd_m = [
        "schtasks", "/Create",
        "/TN", TASK_NAME_MORNING,
        "/TR", tr_m,
        "/SC", "WEEKLY",
        "/D", "MON,TUE,WED,THU,FRI,SAT",
        "/ST", "08:00",
        "/F",
    ]
    p_m = subprocess.run(cmd_m, capture_output=True, text=True, check=False)

    # Task 2: Evening Briefing & ITSM (17:30) — Bắn Zalo & File Markdown
    tr_e = f'"{pyw}" "{runner_script}" --mode EVENING --dispatch --confirm-send --channels file,zalo'
    cmd_e = [
        "schtasks", "/Create",
        "/TN", TASK_NAME_EVENING,
        "/TR", tr_e,
        "/SC", "WEEKLY",
        "/D", "MON,TUE,WED,THU,FRI,SAT",
        "/ST", "17:30",
        "/F",
    ]
    p_e = subprocess.run(cmd_e, capture_output=True, text=True, check=False)

    success = (p_m.returncode == 0) and (p_e.returncode == 0)
    return {
        "status": "SUCCESS" if success else "PARTIAL_OR_FAILED",
        "action": "REGISTER_WINDOWS_TASKS",
        "registered": success,
        "morning_task": {
            "task_name": TASK_NAME_MORNING,
            "returncode": p_m.returncode,
            "output": p_m.stdout.strip() or p_m.stderr.strip(),
        },
        "evening_task": {
            "task_name": TASK_NAME_EVENING,
            "returncode": p_e.returncode,
            "output": p_e.stdout.strip() or p_e.stderr.strip(),
        },
        "pythonw_executable": pyw,
        "runner_script": runner_script,
    }


def unregister_windows_tasks() -> Dict[str, Any]:
    """Hủy đăng ký các tác vụ khỏi Windows Task Scheduler."""
    if os.name != "nt":
        return {"status": "ERROR", "message": "Chỉ hỗ trợ trên hệ điều hành Windows."}

    cmd_m = ["schtasks", "/Delete", "/TN", TASK_NAME_MORNING, "/F"]
    p_m = subprocess.run(cmd_m, capture_output=True, text=True, check=False)

    cmd_e = ["schtasks", "/Delete", "/TN", TASK_NAME_EVENING, "/F"]
    p_e = subprocess.run(cmd_e, capture_output=True, text=True, check=False)

    # Xóa cả task 17:15 cũ nếu còn sót
    subprocess.run(["schtasks", "/Delete", "/TN", OLD_TASK_NAME_EVENING, "/F"], capture_output=True, text=True, check=False)

    return {
        "status": "SUCCESS",
        "action": "UNREGISTER_WINDOWS_TASKS",
        "morning_task_deleted": p_m.returncode == 0,
        "evening_task_deleted": p_e.returncode == 0,
    }


def run_daemon_loop(check_interval_seconds: int = 30) -> None:
    """Vòng lặp giám sát thời gian thực tự động kích hoạt 08:00 và 17:30 không tốn CPU."""
    print("=" * 70)
    print("⏰  SmartOps Scheduler Runner — Chế độ Daemon Giám Sát Thời Gian Thực")
    print(f"🔄  Khoảng kiểm tra: {check_interval_seconds}s | Múi giờ: Asia/Ho_Chi_Minh")
    print("🌅  Ca Sáng: 08:00 Thứ 2 – Thứ 7 | 🌇  Ca Chiều: 17:30 Thứ 2 – Thứ 7")
    print("=" * 70)

    last_morning_date = ""
    last_evening_date = ""

    while True:
        try:
            now = datetime.now()
            today_str = now.strftime("%Y-%m-%d")
            is_workday = now.weekday() in range(0, 6)  # Mon-Sat

            # Kiem tra Ca Sang (08:00 - 08:05)
            if is_workday and now.hour == 8 and now.minute < 5 and last_morning_date != today_str:
                print(f"[{datetime.now().isoformat()}] 🌅 Kích hoạt Bản tin Đầu ca Sáng...")
                res_m = execute_morning_routine(dispatch_channels=["file", "zalo"], dispatch_dry_run=False)
                last_morning_date = today_str
                print(f"[{datetime.now().isoformat()}] ✅ Hoàn tất Bản tin Sáng: {res_m.get('status')}")

            # Kiem tra Ca Chieu (17:30 - 17:35)
            if is_workday and now.hour == 17 and 30 <= now.minute < 35 and last_evening_date != today_str:
                print(f"[{datetime.now().isoformat()}] 🌇 Kích hoạt Bản tin Cuối ca Chiều & Bù đơn ITSM...")
                res_e = execute_evening_routine(auto_remind_itsm=False, dispatch_channels=["file", "zalo"], dispatch_dry_run=False)
                last_evening_date = today_str
                print(f"[{datetime.now().isoformat()}] ✅ Hoàn tất Bản tin Chiều: {res_e.get('status')}")

            time.sleep(check_interval_seconds)
        except KeyboardInterrupt:
            print("\n🛑 Scheduler Runner đã dừng bởi người dùng.")
            break
        except Exception as e:
            print(f"⚠️  Lỗi vòng lặp Scheduler: {e}")
            time.sleep(check_interval_seconds)


def main() -> None:
    parser = argparse.ArgumentParser(description="SmartOps Scheduler Runner — Lịch Quét & Vận Hành Tự Động Hàng Ngày")
    parser.add_argument(
        "--mode",
        choices=["MORNING", "EVENING", "SYNC", "AUTO", "DAEMON"],
        default="AUTO",
        help="Chế độ chạy: MORNING (08:00), EVENING (17:15), SYNC (Đồng bộ), AUTO (theo giờ), DAEMON (chạy liên tục)",
    )
    parser.add_argument("--dry-run", action="store_true", help="Chạy thử không ghi vào Google/Outlook thật")
    parser.add_argument("--no-export", action="store_true", help="Không xuất tệp markdown bản tin ra ổ đĩa")
    parser.add_argument("--date", default="", help="Ngày chỉ định dạng YYYY-MM-DD")
    parser.add_argument("--quiet", action="store_true", help="Chỉ xuất JSON, không in tiêu đề trang trí")
    parser.add_argument("--register-tasks", action="store_true", help="Đăng ký 2 tác vụ vào Windows Task Scheduler")
    parser.add_argument("--unregister-tasks", action="store_true", help="Hủy đăng ký 2 tác vụ khỏi Windows Task Scheduler")
    parser.add_argument("--status-tasks", action="store_true", help="Kiểm tra trạng thái Windows Task Scheduler")
    parser.add_argument("--dispatch", action="store_true", help="Bắn bản tin ra các kênh thông báo (WeChat, Zalo)")
    parser.add_argument("--channels", default="file,wechat,zalo", help="Danh sách kênh phân phối (phân tách bởi dấu phẩy, mặc định: file,wechat,zalo)")
    parser.add_argument("--confirm-send", action="store_true", help="Xác nhận gửi thật (mặc định là Preview Dry-Run)")
    parser.add_argument("--remind-itsm", action="store_true",
                        help="(EVENING) Tạo lịch nhắc bù đơn [ITSM-Remind] trên Outlook/Google/.ics — mặc định TẮT (Đức chốt 10/11/2026)")
    args = parser.parse_args()

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    if args.register_tasks:
        res = register_windows_tasks()
        print(json.dumps(res, ensure_ascii=False, indent=2))
        return

    if args.unregister_tasks:
        res = unregister_windows_tasks()
        print(json.dumps(res, ensure_ascii=False, indent=2))
        return

    if args.status_tasks:
        res = get_windows_tasks_status()
        print(json.dumps(res, ensure_ascii=False, indent=2))
        return

    if args.mode == "DAEMON":
        run_daemon_loop()
        return

    export_md = not args.no_export
    dispatch_chs = [c.strip() for c in args.channels.split(",")] if args.dispatch else None
    dispatch_dry = not args.confirm_send

    if args.mode == "MORNING":
        out = execute_morning_routine(
            date_iso=args.date,
            dry_run=args.dry_run,
            export_md=export_md,
            dispatch_channels=dispatch_chs,
            dispatch_dry_run=dispatch_dry,
        )
    elif args.mode == "EVENING":
        out = execute_evening_routine(
            date_iso=args.date,
            auto_remind_itsm=args.remind_itsm,
            dry_run=args.dry_run,
            export_md=export_md,
            dispatch_channels=dispatch_chs,
            dispatch_dry_run=dispatch_dry,
        )
    elif args.mode == "SYNC":
        out = execute_sync_routine(dry_run=args.dry_run)
    else:  # AUTO
        out = execute_auto_routine(date_iso=args.date, dry_run=args.dry_run, export_md=export_md)

    if args.quiet:
        print(json.dumps(out, ensure_ascii=False, default=str))
    else:
        print("=" * 70)
        print(f"🎯  SmartOps Scheduler — Kết quả thực thi ({out.get('action')})")
        print(f"📅  Ngày: {out.get('date')} | Trạng thái: {out.get('status')}")
        if out.get("briefing_title"):
            print(f"📋  Tiêu đề: {out.get('briefing_title')}")
        if out.get("exported_files"):
            print("💾  Tệp đã xuất:")
            for ef in out["exported_files"]:
                print(f"    - {ef}")
        print("=" * 70)
        # In doan trich ban tin
        briefing_md = out.get("briefing", {}).get("markdown_briefing", "")
        if briefing_md:
            print(briefing_md)


if __name__ == "__main__":
    main()
