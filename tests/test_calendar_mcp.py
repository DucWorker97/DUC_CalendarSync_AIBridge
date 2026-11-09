# -*- coding: utf-8 -*-
"""
test_calendar_mcp.py — Bộ Kiểm Thử Tự Động cho DUC_CalendarSync_AIBridge (v1.1.0)
========================================================================
Kiểm định:
1. Sinh 3 lịch biểu vận hành tự động & xuất tệp chuẩn quốc tế RFC 5545 (.ics).
2. Quét sổ cái ITSM (THEO-DOI-BU-TICKET-ITSM.md) phát hiện Candidate Bxx đang chờ bù đơn / thiếu Mã NV.
3. Tự động đặt lịch nhắc bù đơn ITSM (schedule_pending_itsm_reminder) kèm mẫu câu hỏi Mã NV (Quy tắc 4).
4. Quản lý vòng đời sự kiện (tạo & xóa sự kiện tùy chỉnh, bảo vệ sự kiện hệ thống).
5. Truy vấn lịch trình 2 chiều (engine_query_schedule) & phát hiện xung đột (engine_check_conflicts).
6. Kiểm thử 8 MCP Tools, đăng ký mcp_config.json và chế độ `--selftest` của `server.py`.
========================================================================
"""

import shutil
import uuid
from datetime import datetime
from pathlib import Path
from typing import Iterator

import pytest

from src.calendar_engine import (
    engine_check_conflicts,
    engine_create_event,
    engine_delete_event,
    engine_get_status,
    engine_list_events,
    engine_query_schedule,
    engine_remind_pending_itsm,
    engine_sync_all,
)
from src.itsm_reminder_watcher import scan_pending_itsm_candidates
from server import register_into_mcp_config, run_selftest


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


def test_engine_create_and_delete_custom_event(tmp_path: Path):
    res = engine_create_event(
        title="[MES-Ops] Kiểm tra máy in nhãn Xưởng C9",
        start_iso="2026-11-11T14:30:00",
        duration_minutes=45,
        targets=("ics",),
        dry_run=False,
        base_dir=tmp_path,
    )
    assert res["status"] == "SUCCESS"
    ev_id = res["event"]["event_id"]
    assert res["event"]["title"] == "[MES-Ops] Kiểm tra máy in nhãn Xưởng C9"
    assert "calendar.google.com" in res["google_web_url"]
    assert "outlook.office.com" in res["outlook_web_url"]

    # Bảo vệ sự kiện hệ thống
    prot = engine_delete_event("DAILY_BACKUP_1745", targets=("ics",), base_dir=tmp_path)
    assert prot["status"] == "PROTECTED_SYSTEM_EVENT"

    # Xóa sự kiện tùy chỉnh
    del_res = engine_delete_event(ev_id, targets=("ics",), dry_run=False, base_dir=tmp_path)
    assert del_res["status"] == "SUCCESS"
    assert del_res["found_in_custom_store"] is True
    assert del_res["ics_deleted"] is True


def test_query_schedule_and_conflict_checker(tmp_path: Path):
    q = engine_query_schedule(
        date_from_iso="2026-11-09T08:00:00",
        days=7,
        include_live_sources=False,
        base_dir=tmp_path,
    )
    assert q["status"] == "SUCCESS"
    assert q["total_items"] >= 1

    # Kiểm tra trùng khung giờ sao lưu 17:45 ngày 09/11/2026
    c_busy = engine_check_conflicts(
        start_iso="2026-11-09T17:45:00",
        duration_minutes=15,
        include_live_sources=False,
        base_dir=tmp_path,
    )
    assert c_busy["status"] == "CONFLICT_DETECTED"
    assert c_busy["has_conflict"] is True
    assert len(c_busy["suggested_free_slots"]) == 3

    # Kiểm tra khung giờ trống 10:00 sáng
    c_free = engine_check_conflicts(
        start_iso="2026-11-09T10:00:00",
        duration_minutes=30,
        include_live_sources=False,
        base_dir=tmp_path,
    )
    assert c_free["status"] == "FREE"
    assert c_free["has_conflict"] is False


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


def test_mcp_server_selftest_and_register(tmp_path: Path):
    st = run_selftest()
    assert st["selftest"] == "PASS"
    assert st["version"] == "1.1.0"
    assert len(st["tools"]) == 8

    dummy_cfg = tmp_path / "mcp_config.json"
    reg = register_into_mcp_config(config_path=dummy_cfg)
    assert reg["status"] == "REGISTERED"
    assert dummy_cfg.is_file()
