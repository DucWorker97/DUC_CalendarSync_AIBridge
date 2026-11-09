# -*- coding: utf-8 -*-
"""
itsm_reminder_watcher.py — Tự Động Quét Sự Cố Chờ Bù Đơn (Bxx) & Đặt Lịch Nhắc Thông Minh
=========================================================================================
Thực thi Quy tắc 2 (Universal Ticket Creation) & Quy tắc 4 (Thu thập Mã NV):
1. Quét sổ cái `THEO-DOI-BU-TICKET-ITSM.md` để tìm các Candidate `Bxx` chưa đóng (chưa có mã `INC...`,
   đang ở trạng thái `PENDING`, `CHỜ BÙ`, `CHỜ MÃ NV`, `⏳`, `🟡`).
2. Nếu có Candidate chờ xử lý (hoặc được truyền trực tiếp qua tham số `pending_items`):
   - Tự động sinh sự kiện nhắc việc `[ITSM-Remind] 🎫 Nhắc Bù Đơn Ticket ITSM & Thu Thập Mã NV`
     vào khung giờ hành chính tiếp theo (16:00 chiều nay hoặc 09:00 sáng mai).
   - Đồng bộ thẳng lên Google Calendar, Microsoft Outlook Desktop và tệp `.ics`.
   - Đính kèm sẵn câu mẫu song ngữ VI–ZH (Quy tắc 4) để hỏi Mã nhân viên (`V0xxxx`).
=========================================================================================
"""

import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from src.models import SmartOpsCalendarEvent
from src.ics_exporter import export_events_to_ics, build_google_calendar_url, build_outlook_web_url
from src.outlook_adapter import sync_events_to_outlook
from src.gcalendar_adapter import sync_events_to_gcalendar

DEFAULT_LEDGER_PATHS = [
    Path(r"D:\DucWorker\DUC_SmartOps_AIHub\knowledge_core\01_SoCai_DanhBa\THEO-DOI-BU-TICKET-ITSM.md"),
    Path(r"D:\DucWorker\data amazon quick\03_TU-DIEN-LOI\THEO-DOI-BU-TICKET-ITSM.md"),
]


def locate_itsm_ledger(custom_path: Optional[str] = None) -> Optional[Path]:
    if custom_path:
        p = Path(custom_path)
        if p.is_file():
            return p
    for cand in DEFAULT_LEDGER_PATHS:
        if cand.is_file():
            return cand
    return None


def scan_pending_itsm_candidates(ledger_path: Optional[Path] = None) -> Dict[str, Any]:
    """
    Quét bảng Markdown trong `THEO-DOI-BU-TICKET-ITSM.md` để lọc các Candidate `Bxx`
    đang chờ bù đơn hoặc chờ xác thực Mã nhân viên (`V0xxxx`).
    """
    target_file = ledger_path or locate_itsm_ledger()
    if not target_file or not target_file.is_file():
        return {
            "ledger_found": False,
            "ledger_path": str(target_file) if target_file else None,
            "total_candidates": 0,
            "done_count": 0,
            "pending_count": 0,
            "pending_candidates": [],
        }

    text = target_file.read_text(encoding="utf-8", errors="ignore")
    row_pattern = re.compile(r"^\|\s*\**(B\d+)\**\s*\|(.+)$", re.MULTILINE)

    total_candidates = 0
    done_count = 0
    pending_list: List[Dict[str, str]] = []

    for match in row_pattern.finditer(text):
        b_code = match.group(1).strip()
        rest_cols = [c.strip() for c in match.group(2).split("|")]
        total_candidates += 1
        full_row_str = " | ".join(rest_cols)

        has_inc = bool(re.search(r"INC\d{5,}", full_row_str))
        upper_row = full_row_str.upper()
        is_done = has_inc or ("ĐÃ BÙ" in upper_row) or ("DONE" in upper_row) or ("CLOSED" in upper_row) or ("RESOLVED" in upper_row) or ("HOÀN TẤT" in upper_row)
        is_explicit_pending = any(
            kw in upper_row
            for kw in ["CHỜ BÙ", "CHƯA BÙ", "PENDING", "CHỜ MÃ NV", "THIẾU MÃ NV", "TODO", "⏳", "🟡"]
        )

        if is_done and not is_explicit_pending:
            done_count += 1
        else:
            date_str = rest_cols[0] if len(rest_cols) > 0 else ""
            customer_str = rest_cols[1] if len(rest_cols) > 1 else ""
            issue_str = rest_cols[2] if len(rest_cols) > 2 else full_row_str[:100]
            status_str = rest_cols[-2] if len(rest_cols) >= 2 else "PENDING"
            pending_list.append({
                "candidate_id": b_code,
                "date": date_str,
                "customer": customer_str,
                "summary": issue_str,
                "status": status_str,
            })

    return {
        "ledger_found": True,
        "ledger_path": str(target_file),
        "total_candidates": total_candidates,
        "done_count": done_count,
        "pending_count": len(pending_list),
        "pending_candidates": pending_list,
    }


def _next_reminder_slot(now: datetime) -> datetime:
    """Tính khung giờ nhắc bù đơn gần nhất: 16:00 chiều nay (nếu trước 15:45) hoặc 09:00 sáng ngày làm việc kế tiếp."""
    slot_today = now.replace(hour=16, minute=0, second=0, microsecond=0)
    if now < (slot_today - timedelta(minutes=15)) and now.weekday() < 6:
        return slot_today
    next_day = (now + timedelta(days=1)).replace(hour=9, minute=0, second=0, microsecond=0)
    while next_day.weekday() == 6:  # Chủ nhật -> Thứ 2
        next_day += timedelta(days=1)
    return next_day


def schedule_pending_itsm_reminder(
    ledger_path: Optional[Path] = None,
    extra_pending_items: Optional[List[Dict[str, str]]] = None,
    reminder_iso: Optional[str] = None,
    targets: Sequence[str] = ("ics", "outlook", "gcalendar"),
    ics_dir: Optional[Path] = None,
    dry_run: bool = False,
    reference_dt: Optional[datetime] = None,
) -> Dict[str, Any]:
    """
    Quét sổ cái ITSM (cộng thêm các mục `extra_pending_items` truyền trực tiếp nếu có).
    Nếu có bất kỳ sự cố `Bxx` nào đang chờ bù đơn hoặc thiếu Mã NV (`V0xxxx`), tự động tạo
    lịch nhắc trên Google Calendar, Microsoft Outlook Desktop và tệp `.ics`.
    """
    scan_res = scan_pending_itsm_candidates(ledger_path=ledger_path)
    pending_items = list(scan_res.get("pending_candidates", []))
    if extra_pending_items:
        existing_ids = {item.get("candidate_id") for item in pending_items}
        for ext in extra_pending_items:
            if ext.get("candidate_id") not in existing_ids:
                pending_items.append(ext)

    if not pending_items:
        return {
            "status": "ALL_CLEAR_ZERO_PENDING",
            "message": "100% Candidate Bxx trong sổ cái đã được bù đơn và đóng hoàn tất (0 tồn đọng). Không cần tạo lịch nhắc.",
            "scan_summary": scan_res,
            "scheduled": False,
        }

    now = reference_dt or datetime.now()
    start_dt = datetime.fromisoformat(reminder_iso) if reminder_iso else _next_reminder_slot(now)
    codes_str = ", ".join(item.get("candidate_id", "Bxx") for item in pending_items[:5])
    if len(pending_items) > 5:
        codes_str += f" (+{len(pending_items) - 5})"

    lines_detail = [
        f"Phát hiện {len(pending_items)} sự cố chờ bù đơn ITSM / chờ xác thực Mã nhân viên (Quy tắc 2 & Quy tắc 4):",
    ]
    for idx, it in enumerate(pending_items, 1):
        lines_detail.append(
            f"{idx}. [{it.get('candidate_id', 'Bxx')}] Người báo: {it.get('customer', 'Chưa rõ')} — {it.get('summary', '')}"
        )
    lines_detail.append("")
    lines_detail.append(
        "Mẫu câu hỏi Mã NV (Quy tắc 4): 'Nhờ anh/chị cung cấp thêm Mã nhân viên (V0xxxx / 工号) để em lưu hồ sơ hỗ trợ nhé / 麻烦提供一下工号以便系统建单。'"
    )

    event_id = f"ITSM_REMIND_{start_dt.strftime('%Y%m%d')}"
    ev = SmartOpsCalendarEvent(
        event_id=event_id,
        title=f"[ITSM-Remind] 🎫 Nhắc Bù Đơn Ticket ITSM ({len(pending_items)} đơn: {codes_str})",
        description="\n".join(lines_detail),
        location="BMC Remedy Smart IT & THEO-DOI-BU-TICKET-ITSM.md",
        start_dt=start_dt,
        duration_minutes=30,
        rrule=None,
        reminder_minutes=15,
        category="MES-SmartOps-ITSM-Reminder",
        recurrence_type="NONE",
    )

    out_ics_dir = ics_dir or Path(__file__).resolve().parents[1] / "exports"
    ics_res = export_events_to_ics([ev], out_dir=out_ics_dir, master_filename=f"{event_id}.ics") if "ics" in targets else {}
    outlook_res = sync_events_to_outlook([ev], dry_run=dry_run) if "outlook" in targets else {}
    gcal_res = sync_events_to_gcalendar([ev], ics_dir=out_ics_dir, dry_run=dry_run) if "gcalendar" in targets else {}

    return {
        "status": "DRY_RUN" if dry_run else "SCHEDULED_REMINDER",
        "scheduled": True,
        "pending_count": len(pending_items),
        "pending_candidates": pending_items,
        "event": ev.to_dict(),
        "google_web_url": build_google_calendar_url(ev),
        "outlook_web_url": build_outlook_web_url(ev),
        "ics": ics_res,
        "outlook": outlook_res,
        "gcalendar": gcal_res,
    }
