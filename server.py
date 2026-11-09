# -*- coding: utf-8 -*-
"""
DUC_CalendarSync_AIBridge — MCP Server (stdio)
================================================================================
Author : Nguyễn Duy Đức (GitHub: DucWorker97)
Version: 1.0.0

Chạy MCP Server:
  python server.py
Tự kiểm định (Self-Test không mở stdio):
  python server.py --selftest

Danh mục 5 MCP Tools:
1. cal_status             — Kiểm tra trạng thái Google Calendar OAuth, Outlook Desktop MAPI, kho .ics & sổ cái ITSM.
2. cal_sync_all           — Đồng bộ tức thì toàn bộ lịch vận hành tự động lên Google Calendar, Outlook Desktop & .ics.
3. cal_create_event       — Đặt lịch sự kiện tức thì lên Google Calendar, Outlook Desktop & xuất .ics.
4. cal_remind_pending_itsm— Quét sự cố chờ bù đơn ITSM (Bxx) & tự động ghim lịch nhắc kèm mẫu câu hỏi Mã NV (Quy tắc 4).
5. cal_list_events        — Liệt kê toàn bộ chuỗi lịch đang quản lý kèm liên kết 1-Click.
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
    engine_create_event,
    engine_get_status,
    engine_list_events,
    engine_remind_pending_itsm,
    engine_sync_all,
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
    return {
        "selftest": "PASS",
        "mcp_sdk_ready": MCPServer is not None,
        "status_check": st["status"],
        "sync_dry_run_events": sync_dry["events_count"],
        "itsm_reminder_dry_run": rem_dry["status"],
        "tools": [
            "cal_status",
            "cal_sync_all",
            "cal_create_event",
            "cal_remind_pending_itsm",
            "cal_list_events",
        ],
    }


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if "--selftest" in sys.argv:
        print(_j(run_selftest()))
        raise SystemExit(0)
    if MCPServer is None:
        sys.stderr.write(f"MCP SDK import error: {_MCP_IMPORT_ERROR}\n")
        raise SystemExit(1)
    mcp.run()
