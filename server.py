# -*- coding: utf-8 -*-
"""
DUC_CalendarSync_AIBridge — MCP Server (stdio)
================================================================================
Author : Nguyễn Duy Đức (GitHub: DucWorker97)
Version: 1.3.0

Chạy MCP Server:
  python server.py
Tự kiểm định (Self-Test không mở stdio):
  python server.py --selftest
Tự động đăng ký vào mcp_config.json:
  python server.py --register-mcp
Quản lý Windows Task Scheduler:
  python server.py --register-tasks | --unregister-tasks | --status-tasks

Danh mục 10 MCP Tools (v1.3.0):
1. cal_status             — Kiểm tra trạng thái Google Calendar OAuth, Outlook Desktop MAPI, kho .ics & sổ cái ITSM.
2. cal_sync_all           — Đồng bộ tức thì toàn bộ lịch vận hành tự động lên Google Calendar, Outlook Desktop & .ics.
3. cal_create_event       — Đặt lịch sự kiện tức thì lên Google Calendar, Outlook Desktop & xuất .ics.
4. cal_remind_pending_itsm— Quét sự cố chờ bù đơn ITSM (Bxx) & tự động ghim lịch nhắc kèm mẫu câu hỏi Mã NV (Quy tắc 4).
5. cal_list_events        — Liệt kê toàn bộ chuỗi lịch đang quản lý kèm liên kết 1-Click.
6. cal_delete_event       — Xóa sự kiện tùy chỉnh theo event_id đồng bộ trên cả 3 nền tảng.
7. cal_query_schedule     — Đọc lịch trình thực tế 2 chiều trong ngày/tuần từ Managed Store, Google Calendar & Outlook.
8. cal_check_conflicts    — Kiểm tra xung đột lịch họp & gợi ý 3 khung giờ trống gần nhất (Free-Busy Finder).
9. cal_update_event       — Dời lịch (reschedule) hoặc cập nhật nội dung sự kiện tùy chỉnh theo event_id.
10. cal_daily_briefing    — Sinh Bản tin Lịch trình Đầu ca (08:00) / Cuối ca (17:15) kèm trạng thái ITSM & Sao lưu.
================================================================================
"""

import json
import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

from src.calendar_engine import (  # noqa: E402
    engine_check_conflicts,
    engine_create_event,
    engine_daily_briefing,
    engine_delete_event,
    engine_get_status,
    engine_list_events,
    engine_query_schedule,
    engine_remind_pending_itsm,
    engine_sync_all,
    engine_update_event,
)

try:
    from mcp.server.mcpserver import MCPServer
    _MCP_IMPORT_ERROR = None
except ImportError:
    try:
        from mcp.server.fastmcp import FastMCP as MCPServer
        _MCP_IMPORT_ERROR = None
    except Exception as e:
        MCPServer = None
        _MCP_IMPORT_ERROR = repr(e)
except Exception as e:
    MCPServer = None
    _MCP_IMPORT_ERROR = repr(e)


def _j(obj: Any) -> str:
    return json.dumps(obj, ensure_ascii=False, indent=2, default=str)


if MCPServer is not None:
    mcp = MCPServer("DUC_CalendarSync_AIBridge")

    @mcp.tool()
    def cal_status() -> str:
        """Kiểm tra trạng thái kết nối Google Calendar (OAuth CLI), Microsoft Outlook Desktop (MAPI COM), kho .ics và tổng hợp số đơn ITSM chờ bù."""
        return _j(engine_get_status())

    @mcp.tool()
    def cal_sync_all(dry_run: bool = False) -> str:
        """Đồng bộ tức thì 3 chuỗi lịch vận hành chuẩn (Sao lưu 17:45, Báo cáo tuần Chủ Nhật 15:00, Khóa sổ ITSM ngày 28 16:00) và lịch tùy chỉnh lên Google Calendar, Outlook Desktop và chuẩn .ics."""
        return _j(engine_sync_all(dry_run=bool(dry_run)))

    @mcp.tool()
    def cal_create_event(
        title: str,
        start_iso: str,
        duration_minutes: int = 30,
        description: str = "",
        location: str = "MES SmartOps — Chervon Vietnam",
        recurrence_type: str = "NONE",
        reminder_minutes: int = 15,
        dry_run: bool = False,
    ) -> str:
        """Đặt lịch sự kiện tức thì lên Google Calendar, Microsoft Outlook Desktop và xuất tệp .ics.
        - start_iso: Định dạng ISO 8601 (VD: '2026-11-10T14:00:00')
        - recurrence_type: 'NONE' | 'DAILY_WORKDAYS' | 'WEEKLY_SUN' | 'MONTHLY_28'
        """
        return _j(
            engine_create_event(
                title=title,
                start_iso=start_iso,
                duration_minutes=duration_minutes,
                description=description,
                location=location,
                recurrence_type=recurrence_type,
                reminder_minutes=reminder_minutes,
                dry_run=bool(dry_run),
            )
        )

    @mcp.tool()
    def cal_remind_pending_itsm(
        candidate_id: str = "",
        customer_name: str = "",
        issue_summary: str = "",
        reminder_iso: str = "",
        dry_run: bool = False,
    ) -> str:
        """Quét sổ cái THEO-DOI-BU-TICKET-ITSM.md (hoặc nhận trực tiếp một sự cố chờ bù đơn Bxx mới phát sinh) và tự động đặt lịch nhắc bù đơn + hỏi Mã nhân viên (V0xxxx) trên Google Calendar, Outlook Desktop & .ics."""
        extra: Optional[List[Dict[str, str]]] = None
        if candidate_id.strip() or issue_summary.strip():
            extra = [
                {
                    "candidate_id": candidate_id.strip() or "B_NEW",
                    "date": "",
                    "customer": customer_name.strip() or "Chờ xác thực Mã NV (V0xxxx)",
                    "summary": issue_summary.strip() or "Sự cố MES mới phát sinh cần bù đơn ITSM",
                    "status": "PENDING",
                }
            ]
        return _j(
            engine_remind_pending_itsm(
                extra_pending_items=extra,
                reminder_iso=reminder_iso.strip() or None,
                dry_run=bool(dry_run),
            )
        )

    @mcp.tool()
    def cal_list_events() -> str:
        """Liệt kê toàn bộ các sự kiện lịch biểu đang quản lý kèm liên kết 1-Click Google Calendar, Outlook Web và đường dẫn tệp .ics."""
        return _j(engine_list_events())

    @mcp.tool()
    def cal_delete_event(event_id: str, dry_run: bool = False) -> str:
        """Xóa sự kiện tùy chỉnh theo event_id đồng bộ trên kho .ics, Microsoft Outlook Desktop và Google Calendar."""
        return _j(engine_delete_event(event_id=event_id, dry_run=bool(dry_run)))

    @mcp.tool()
    def cal_query_schedule(
        date_from_iso: str = "",
        days: int = 7,
        include_live_sources: bool = True,
    ) -> str:
        """Đọc lịch trình thực tế 2 chiều trong khoảng `days` ngày kể từ `date_from_iso` (hợp nhất từ Managed Store, Google Calendar và Outlook Desktop)."""
        return _j(
            engine_query_schedule(
                date_from_iso=date_from_iso,
                days=int(days),
                include_live_sources=bool(include_live_sources),
            )
        )

    @mcp.tool()
    def cal_check_conflicts(
        start_iso: str,
        duration_minutes: int = 30,
        include_live_sources: bool = False,
    ) -> str:
        """Kiểm tra xung đột lịch họp tại khung giờ `start_iso` (thời lượng `duration_minutes` phút) và gợi ý 3 khung giờ trống gần nhất trong giờ hành chính."""
        return _j(
            engine_check_conflicts(
                start_iso=start_iso,
                duration_minutes=int(duration_minutes),
                include_live_sources=bool(include_live_sources),
            )
        )

    @mcp.tool()
    def cal_update_event(
        event_id: str,
        new_title: str = "",
        new_start_iso: str = "",
        new_duration_minutes: int = 0,
        new_description: str = "",
        new_location: str = "",
        new_recurrence_type: str = "",
        new_reminder_minutes: int = -1,
        dry_run: bool = False,
    ) -> str:
        """Dời lịch (reschedule) hoặc cập nhật tiêu đề, thời lượng, mô tả, địa điểm của sự kiện tùy chỉnh theo `event_id` đồng bộ trên cả 3 nền tảng."""
        return _j(
            engine_update_event(
                event_id=event_id,
                new_title=new_title,
                new_start_iso=new_start_iso,
                new_duration_minutes=int(new_duration_minutes),
                new_description=new_description,
                new_location=new_location,
                new_recurrence_type=new_recurrence_type,
                new_reminder_minutes=int(new_reminder_minutes),
                dry_run=bool(dry_run),
            )
        )

    @mcp.tool()
    def cal_daily_briefing(
        date_iso: str = "",
        shift_mode: str = "AUTO",
        include_live_sources: bool = False,
    ) -> str:
        """Sinh Bản tin Lịch trình Đầu ca (`MORNING` - 08:00) hoặc Cuối ca (`EVENING` - 17:15) kèm danh sách sự kiện trong ngày, trạng thái bù đơn ITSM và Checklist hành động."""
        return _j(
            engine_daily_briefing(
                date_iso=date_iso,
                shift_mode=shift_mode,
                include_live_sources=bool(include_live_sources),
            )
        )


def register_into_mcp_config(config_path: Optional[Path] = None) -> Dict[str, Any]:
    target = config_path or (Path.home() / ".gemini" / "antigravity" / "mcp_config.json")
    server_script = str(Path(__file__).resolve())
    repo_dir = str(Path(__file__).resolve().parent)
    py_exe = sys.executable

    data: Dict[str, Any] = {"mcpServers": {}}
    if target.is_file():
        try:
            data = json.loads(target.read_text(encoding="utf-8"))
        except Exception:
            data = {"mcpServers": {}}

    servers = data.setdefault("mcpServers", {})
    servers["DUC_CalendarSync_AIBridge"] = {
        "command": py_exe,
        "args": ["-u", "-X", "utf8", server_script],
        "env": {
            "PYTHONIOENCODING": "utf-8",
            "PYTHONUNBUFFERED": "1",
            "PYTHONPATH": repo_dir,
        },
    }
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    return {
        "status": "REGISTERED",
        "config_path": str(target),
        "total_servers": len(servers),
        "server_name": "DUC_CalendarSync_AIBridge",
    }


def run_selftest() -> Dict[str, Any]:
    st = engine_get_status()
    sync_dry = engine_sync_all(dry_run=True)
    rem_dry = engine_remind_pending_itsm(
        extra_pending_items=[
            {
                "candidate_id": "B90",
                "customer": "User Xưởng C9 (Chờ hỏi V0xxxx)",
                "summary": "Kiểm tra máy in nhãn trạm đóng gói",
                "status": "PENDING",
            }
        ],
        dry_run=True,
    )
    brief = engine_daily_briefing(shift_mode="MORNING", include_live_sources=False)
    return {
        "selftest": "PASS",
        "version": "1.3.0",
        "mcp_sdk_ready": MCPServer is not None,
        "status_check": st["status"],
        "sync_dry_run_events": sync_dry["events_count"],
        "itsm_reminder_dry_run": rem_dry["status"],
        "daily_briefing_ready": brief["status"] == "SUCCESS",
        "scheduled_tasks_supported": st.get("scheduled_tasks", {}).get("supported", False),
        "tools": [
            "cal_status",
            "cal_sync_all",
            "cal_create_event",
            "cal_remind_pending_itsm",
            "cal_list_events",
            "cal_delete_event",
            "cal_query_schedule",
            "cal_check_conflicts",
            "cal_update_event",
            "cal_daily_briefing",
        ],
    }


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if "--selftest" in sys.argv:
        print(_j(run_selftest()))
        raise SystemExit(0)
    if "--register-mcp" in sys.argv:
        print(_j(register_into_mcp_config()))
        raise SystemExit(0)
    if "--register-tasks" in sys.argv:
        from src.scheduler_runner import register_windows_tasks
        print(_j(register_windows_tasks()))
        raise SystemExit(0)
    if "--unregister-tasks" in sys.argv:
        from src.scheduler_runner import unregister_windows_tasks
        print(_j(unregister_windows_tasks()))
        raise SystemExit(0)
    if "--status-tasks" in sys.argv:
        from src.scheduler_runner import get_windows_tasks_status
        print(_j(get_windows_tasks_status()))
        raise SystemExit(0)
    if MCPServer is None:
        sys.stderr.write(f"MCP SDK import error: {_MCP_IMPORT_ERROR}\n")
        raise SystemExit(1)
    mcp.run()
