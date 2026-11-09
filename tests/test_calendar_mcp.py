# -*- coding: utf-8 -*-
"""
test_calendar_mcp.py — Bộ Kiểm Thử Tự Động cho DUC_CalendarSync_AIBridge (v1.3.0)
========================================================================
Kiểm định:
1. Sinh 3 lịch biểu vận hành tự động & xuất tệp chuẩn quốc tế RFC 5545 (.ics).
2. Quét sổ cái ITSM (THEO-DOI-BU-TICKET-ITSM.md) phát hiện Candidate Bxx đang chờ bù đơn / thiếu Mã NV.
3. Tự động đặt lịch nhắc bù đơn ITSM (schedule_pending_itsm_reminder) kèm mẫu câu hỏi Mã NV (Quy tắc 4).
4. Quản lý vòng đời sự kiện (tạo, cập nhật/dời lịch & xóa sự kiện tùy chỉnh, bảo vệ sự kiện hệ thống).
5. Truy vấn lịch trình 2 chiều (engine_query_schedule) & phát hiện xung đột (engine_check_conflicts).
6. Sinh Bản tin Đầu ca / Cuối ca (engine_daily_briefing) ở cả 2 chế độ MORNING & EVENING.
7. Kiểm thử Trình thực thi Lịch quét tự động hàng ngày (scheduler_runner.py) & Windows Task Scheduler.
8. Kiểm thử 10 MCP Tools, đăng ký mcp_config.json và chế độ `--selftest` của `server.py`.
========================================================================
"""

import importlib.util
import shutil
import sys
import uuid
from datetime import datetime
from pathlib import Path
from typing import Iterator

import pytest

_BRIDGE_ROOT = Path(__file__).resolve().parents[1]
if str(_BRIDGE_ROOT) not in sys.path:
    sys.path.insert(0, str(_BRIDGE_ROOT))

from src.calendar_engine import (
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
from src.itsm_reminder_watcher import scan_pending_itsm_candidates
from src.scheduler_runner import (
    execute_auto_routine,
    execute_evening_routine,
    execute_morning_routine,
    execute_sync_routine,
    get_windows_tasks_status,
)

_server_spec = importlib.util.spec_from_file_location(
    "calendar_sync_server", _BRIDGE_ROOT / "server.py"
)
assert _server_spec is not None and _server_spec.loader is not None
_calendar_server_mod = importlib.util.module_from_spec(_server_spec)
_server_spec.loader.exec_module(_calendar_server_mod)
register_into_mcp_config = _calendar_server_mod.register_into_mcp_config
run_selftest = _calendar_server_mod.run_selftest


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


def test_engine_create_update_and_delete_custom_event(tmp_path: Path):
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

    # Cập nhật / Dời lịch sự kiện tùy chỉnh (cal_update_event)
    upd = engine_update_event(
        event_id=ev_id,
        new_title="[MES-Ops] Kiểm tra & Bảo trì máy in nhãn Xưởng C9",
        new_start_iso="2026-11-11T15:30:00",
        new_duration_minutes=60,
        targets=("ics",),
        dry_run=False,
        base_dir=tmp_path,
    )
    assert upd["status"] == "SUCCESS"
    assert upd["previous_start_dt"] == "2026-11-11T14:30:00"
    assert upd["event"]["start_dt"] == "2026-11-11T15:30:00"
    assert upd["event"]["duration_minutes"] == 60
    assert "Bảo trì" in upd["event"]["title"]

    # Không cho phép update/delete sự kiện hệ thống
    prot_upd = engine_update_event("DAILY_BACKUP_1745", new_title="Test", base_dir=tmp_path)
    assert prot_upd["status"] == "PROTECTED_SYSTEM_EVENT"

    prot_del = engine_delete_event("DAILY_BACKUP_1745", targets=("ics",), base_dir=tmp_path)
    assert prot_del["status"] == "PROTECTED_SYSTEM_EVENT"

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


def test_daily_briefing_morning_and_evening(tmp_path: Path):
    sample_ledger = tmp_path / "THEO-DOI-BU-TICKET-ITSM.md"
    sample_ledger.write_text(
        "| Mã | Ngày | Customer | Nội dung | Trạng thái | Ticket |\n"
        "|---|---|---|---|---|---|\n"
        "| **B91** | 09/11/2026 | User Xưởng D2 | Lỗi PDA | ⏳ Chờ mã NV | — |\n",
        encoding="utf-8",
    )

    morning = engine_daily_briefing(
        date_iso="2026-11-09T08:00:00",
        shift_mode="MORNING",
        include_live_sources=False,
        ledger_path=sample_ledger,
        base_dir=tmp_path,
    )
    assert morning["status"] == "SUCCESS"
    assert morning["shift_mode"] == "MORNING"
    assert morning["itsm_pending_count"] == 1
    assert "B91" in morning["markdown_briefing"]
    assert "V0xxxx" in morning["markdown_briefing"]

    evening = engine_daily_briefing(
        date_iso="2026-11-09T17:15:00",
        shift_mode="EVENING",
        include_live_sources=False,
        ledger_path=sample_ledger,
        base_dir=tmp_path,
    )
    assert evening["status"] == "SUCCESS"
    assert evening["shift_mode"] == "EVENING"
    assert "17:45" in evening["markdown_briefing"]


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
    assert st["version"] == "1.4.0"
    assert st["daily_briefing_ready"] is True
    assert st["notification_dispatch_ready"] is True
    assert st["scheduled_tasks_supported"] is True
    assert len(st["tools"]) == 11

    dummy_cfg = tmp_path / "mcp_config.json"
    reg = register_into_mcp_config(config_path=dummy_cfg)
    assert reg["status"] == "REGISTERED"
    assert dummy_cfg.is_file()


def test_scheduler_runner_routines(tmp_path: Path):
    # Test Morning Routine
    res_m = execute_morning_routine(
        date_iso="2026-11-10T08:00:00",
        dry_run=True,
        export_md=True,
        repo_root=tmp_path,
        output_dir=tmp_path / "exports" / "briefings",
    )
    assert res_m["status"] == "DRY_RUN"
    assert res_m["mode"] == "MORNING"
    assert len(res_m["exported_files"]) >= 1
    m_file = Path(res_m["exported_files"][0])
    assert m_file.is_file()
    assert "08:00" in m_file.read_text(encoding="utf-8")

    # Test Evening Routine (with ITSM ledger scan)
    res_e = execute_evening_routine(
        date_iso="2026-11-10T17:30:00",
        auto_remind_itsm=True,
        dry_run=True,
        export_md=True,
        repo_root=tmp_path,
        output_dir=tmp_path / "exports" / "briefings",
    )
    assert res_e["status"] == "DRY_RUN"
    assert res_e["mode"] == "EVENING"
    assert len(res_e["exported_files"]) >= 1
    e_file = Path(res_e["exported_files"][0])
    assert e_file.is_file()
    assert "17:30" in e_file.read_text(encoding="utf-8")

    # Test Sync Routine
    res_s = execute_sync_routine(dry_run=True, repo_root=tmp_path)
    assert res_s["status"] == "DRY_RUN"
    assert res_s["mode"] == "SYNC"

    # Test Auto Routine
    res_a1 = execute_auto_routine(date_iso="2026-11-10T09:30:00", dry_run=True, export_md=False, repo_root=tmp_path)
    assert res_a1["mode"] == "MORNING"
    res_a2 = execute_auto_routine(date_iso="2026-11-10T15:30:00", dry_run=True, export_md=False, repo_root=tmp_path)
    assert res_a2["mode"] == "EVENING"


def test_scheduler_windows_tasks_status():
    st = get_windows_tasks_status()
    assert st["supported"] is True
    assert "morning_briefing" in st["tasks"]
    assert "evening_briefing" in st["tasks"]
    assert st["tasks"]["morning_briefing"]["schedule"] == "08:00 (Thứ 2 – Thứ 7)"
    assert st["tasks"]["evening_briefing"]["schedule"] == "17:30 (Thứ 2 – Thứ 7)"


def test_notification_dispatcher_formats_and_dryrun(tmp_path: Path):
    from src.notification_dispatcher import (
        format_for_wechat_markdown,
        format_for_zalo_text,
        dispatch_daily_briefing,
    )
    sample_brief = {
        "shift_label": "🌅 BẢN TIN ĐẦU CA SÁNG (08:00 — MORNING BRIEFING)",
        "date": "2026-11-10",
        "today_events": [{"title": "Sao lưu đám mây", "start_dt": "2026-11-10T17:45:00", "source": "smartops_managed"}],
        "itsm_pending_count": 0,
        "itsm_pending_candidates": [],
        "action_checklist": ["Kiểm tra kết nối MES", "Rà soát lịch"],
    }

    # WeChat format check
    w_md = format_for_wechat_markdown(sample_brief)
    assert "### 🌅 BẢN TIN ĐẦU CA SÁNG" in w_md
    assert '<font color="info">' in w_md
    assert len(w_md.encode("utf-8")) <= 4000

    # Zalo format check
    z_txt = format_for_zalo_text(sample_brief)
    assert "📌 🌅 BẢN TIN ĐẦU CA SÁNG" in z_txt
    assert "📅 LỊCH TRÌNH TRONG NGÀY:" in z_txt
    assert "✅ CHECKLIST TRỌNG TÂM THEO CA:" in z_txt

    # Dispatch dry-run check
    disp = dispatch_daily_briefing(
        date_iso="2026-11-10T08:00:00",
        shift_mode="MORNING",
        channels=("file", "wechat", "zalo"),
        dry_run=True,
        target_zalo="Thông báo chung OLAZ",
        repo_root=tmp_path,
    )
    assert disp["status"] == "SUCCESS"
    assert "wechat" in disp["channels"]
    assert "zalo" in disp["channels"]
    assert "file" in disp["channels"]
    assert disp["channels"]["wechat"]["status"] in {"PREVIEW_READY", "NOT_CONFIGURED"}
    assert disp["channels"]["zalo"]["status"] == "PREVIEW_READY"


