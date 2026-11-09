# -*- coding: utf-8 -*-
"""
test_calendar_mcp.py — Bộ Kiểm Thử Tự Động cho DUC_CalendarSync_AIBridge
========================================================================
Kiểm định:
1. Sinh 3 lịch biểu vận hành tự động & xuất tệp chuẩn quốc tế RFC 5545 (.ics).
2. Quét sổ cái ITSM (THEO-DOI-BU-TICKET-ITSM.md) phát hiện Candidate Bxx đang chờ bù đơn / thiếu Mã NV.
3. Tự động đặt lịch nhắc bù đơn ITSM (schedule_pending_itsm_reminder) kèm mẫu câu hỏi Mã NV (Quy tắc 4).
4. Kiểm thử 5 MCP Tools và chế độ `--selftest` của `server.py`.
========================================================================
"""

import shutil
import uuid
from datetime import datetime
from pathlib import Path
from typing import Iterator

import pytest

from src.calendar_engine import (
    engine_create_event,
    engine_get_status,
    engine_list_events,
    engine_remind_pending_itsm,
    engine_sync_all,
)
from src.itsm_reminder_watcher import scan_pending_itsm_candidates
from server import run_selftest


@pytest.fixture
def tmp_path() -> Iterator[Path]:
    """Local workspace temp directory fixture avoiding Python 3.12 + pytest 0o700 DACL issues on Windows."""
    base = Path(__file__).resolve().parents[1] / ".pytest_tmp"
    base.mkdir(parents=True, exist_ok=True)
    work = base / f"case_{uuid.uuid4().hex[:8]}"
    work.mkdir(parents=True, exist_ok=True)
    try:
        yield work
    finally:
        shutil.rmtree(work, ignore_errors=True)


def test_engine_sync_and_list(tmp_path: Path):
    res = engine_sync_all(
        targets=("ics", "outlook", "gcalendar"),
        dry_run=True,
        base_dir=tmp_path,
        reference_dt=datetime(2026, 11, 9, 10, 0, 0),
    )
    assert res["status"] == "DRY_RUN"
    assert res["events_count"] == 3
    assert (tmp_path / "exports" / "DUC_SmartOps_Master_Schedule.ics").is_file()

    lst = engine_list_events(base_dir=tmp_path)
    assert lst["status"] == "SUCCESS"
    assert lst["count"] == 3


def test_engine_create_custom_event(tmp_path: Path):
    res = engine_create_event(
        title="[MES-Ops] Kiểm tra máy in nhãn Xưởng C9",
        start_iso="2026-11-11T14:30:00",
        duration_minutes=45,
        targets=("ics",),
        dry_run=False,
        base_dir=tmp_path,
    )
    assert res["status"] == "SUCCESS"
    assert res["event"]["title"] == "[MES-Ops] Kiểm tra máy in nhãn Xưởng C9"
    assert "calendar.google.com" in res["google_web_url"]
    assert "outlook.office.com" in res["outlook_web_url"]


def test_itsm_pending_scanner_and_auto_reminder(tmp_path: Path):
    sample_ledger = tmp_path / "THEO-DOI-BU-TICKET-ITSM.md"
    sample_ledger.write_text(
        "# Sổ theo dõi bù đơn\n\n"
        "| Mã | Ngày | Customer | Nội dung | Trạng thái | Ticket |\n"
        "|---|---|---|---|---|---|\n"
        "| **B88** | 05/11/2026 | Hoàng Văn Tùng (V002387) | Lỗi trạm đóng gói | Đã bù | INC194400 |\n"
        "| **B89** | 09/11/2026 | User Xưởng C9 (Chưa có V0xxxx) | Nhờ kiểm tra máy in nhãn | ⏳ Chờ mã NV | — |\n",
        encoding="utf-8",
    )

    scan = scan_pending_itsm_candidates(ledger_path=sample_ledger)
    assert scan["ledger_found"] is True
    assert scan["total_candidates"] == 2
    assert scan["done_count"] == 1
    assert scan["pending_count"] == 1
    assert scan["pending_candidates"][0]["candidate_id"] == "B89"

    rem = engine_remind_pending_itsm(
        ledger_path=sample_ledger,
        targets=("ics",),
        dry_run=False,
        base_dir=tmp_path,
    )
    assert rem["status"] == "SCHEDULED_REMINDER"
    assert rem["scheduled"] is True
    assert rem["pending_count"] == 1
    assert "B89" in rem["event"]["title"]
    assert "V0xxxx" in rem["event"]["description"]


def test_itsm_all_clear_when_zero_pending(tmp_path: Path):
    clean_ledger = tmp_path / "THEO-DOI-BU-TICKET-ITSM.md"
    clean_ledger.write_text(
        "| Mã | Ngày | Customer | Nội dung | Trạng thái | Ticket |\n"
        "|---|---|---|---|---|---|\n"
        "| **B88** | 05/11/2026 | Hoàng Văn Tùng | Lỗi trạm | Đã bù | INC194400 |\n",
        encoding="utf-8",
    )
    rem = engine_remind_pending_itsm(ledger_path=clean_ledger, dry_run=True, base_dir=tmp_path)
    assert rem["status"] == "ALL_CLEAR_ZERO_PENDING"
    assert rem["scheduled"] is False


def test_mcp_server_selftest():
    st = run_selftest()
    assert st["selftest"] == "PASS"
    assert len(st["tools"]) == 5
