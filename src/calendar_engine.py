# -*- coding: utf-8 -*-
"""
calendar_engine.py — Bộ Điều Phối Trung Tâm của DUC_CalendarSync_AIBridge (v1.1.0)
=========================================================================
Cung cấp các hàm nghiệp vụ cốt lõi cho MCP Server:
1. cal_status: Kiểm tra trạng thái Google Calendar OAuth, Outlook Desktop MAPI, kho .ics và sổ cái ITSM.
2. cal_sync_all: Đồng bộ toàn bộ lịch vận hành tự động + lịch tùy chỉnh lên cả 3 nền tảng.
3. cal_create_event: Đặt lịch tức thì lên Google Calendar, Outlook Desktop & xuất .ics.
4. cal_remind_pending_itsm: Quét các Candidate Bxx chờ bù đơn / thiếu Mã NV và tự động đặt lịch nhắc.
5. cal_list_events: Liệt kê toàn bộ lịch trình kèm liên kết 1-Click.
6. cal_delete_event: Xóa sự kiện tùy chỉnh theo event_id trên cả 3 nền tảng.
7. cal_query_schedule: Đọc lịch trình thực tế 2 chiều từ Managed Store, Google Calendar & Outlook Desktop.
8. cal_check_conflicts: Kiểm tra xung đột khung giờ & gợi ý 3 khung giờ trống gần nhất (Free-Busy Finder).
=========================================================================
"""

import json
import uuid
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from src.models import SmartOpsCalendarEvent, get_default_smartops_events
from src.ics_exporter import (
    export_events_to_ics,
    build_google_calendar_url,
    build_outlook_web_url,
)
from src.outlook_adapter import (
    delete_event_from_outlook,
    fetch_outlook_events_in_range,
    get_outlook_calendar_status,
    sync_events_to_outlook,
)
from src.gcalendar_adapter import (
    delete_event_from_gcalendar,
    fetch_gcalendar_events_in_range,
    get_gcalendar_status,
    sync_events_to_gcalendar,
)
from src.itsm_reminder_watcher import (
    scan_pending_itsm_candidates,
    schedule_pending_itsm_reminder,
)

REPO_ROOT = Path(__file__).resolve().parents[1]


def _get_ics_dir(base_dir: Optional[Path] = None) -> Path:
    root = base_dir or REPO_ROOT
    return root / "exports"


def _get_custom_events_file(base_dir: Optional[Path] = None) -> Path:
    return _get_ics_dir(base_dir) / "custom_events.json"


def load_custom_events(base_dir: Optional[Path] = None) -> List[SmartOpsCalendarEvent]:
    cfile = _get_custom_events_file(base_dir)
    if not cfile.is_file():
        return []
    try:
        raw_list = json.loads(cfile.read_text(encoding="utf-8"))
        out: List[SmartOpsCalendarEvent] = []
        for item in raw_list:
            out.append(
                SmartOpsCalendarEvent(
                    event_id=item["event_id"],
                    title=item["title"],
                    description=item.get("description", ""),
                    location=item.get("location", "MES SmartOps"),
                    start_dt=datetime.fromisoformat(item["start_dt"]),
                    duration_minutes=int(item.get("duration_minutes", 30)),
                    rrule=item.get("rrule"),
                    reminder_minutes=int(item.get("reminder_minutes", 15)),
                    category=item.get("category", "MES-SmartOps"),
                    recurrence_type=item.get("recurrence_type", "NONE"),
                )
            )
        return out
    except Exception:
        return []


def save_custom_event(event: SmartOpsCalendarEvent, base_dir: Optional[Path] = None) -> None:
    cfile = _get_custom_events_file(base_dir)
    cfile.parent.mkdir(parents=True, exist_ok=True)
    existing = load_custom_events(base_dir)
    filtered = [e for e in existing if e.event_id != event.event_id]
    filtered.append(event)
    cfile.write_text(
        json.dumps([e.to_dict() for e in filtered], indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def get_all_managed_events(
    base_dir: Optional[Path] = None,
    reference_dt: Optional[datetime] = None,
) -> List[SmartOpsCalendarEvent]:
    defaults = get_default_smartops_events(reference_dt=reference_dt)
    customs = load_custom_events(base_dir)
    return defaults + customs


def engine_get_status(
    base_dir: Optional[Path] = None,
    ledger_path: Optional[Path] = None,
) -> Dict[str, Any]:
    root = base_dir or REPO_ROOT
    ics_dir = _get_ics_dir(root)
    events = get_all_managed_events(root)
    outlook_st = get_outlook_calendar_status()
    gcal_st = get_gcalendar_status()
    itsm_scan = scan_pending_itsm_candidates(ledger_path=ledger_path)

    ics_files = sorted(ics_dir.glob("*.ics")) if ics_dir.is_dir() else []

    return {
        "status": "READY",
        "bridge": "DUC_CalendarSync_AIBridge",
        "version": "1.1.0",
        "author": "DucWorker97",
        "checked_at": datetime.now().isoformat(timespec="seconds"),
        "managed_events_count": len(events),
        "platforms": {
            "google_calendar": gcal_st,
            "outlook_desktop": outlook_st,
            "ics_vault": {
                "directory": str(ics_dir),
                "files_count": len(ics_files),
            },
        },
        "itsm_pending_summary": {
            "ledger_found": itsm_scan.get("ledger_found", False),
            "total_candidates": itsm_scan.get("total_candidates", 0),
            "done_count": itsm_scan.get("done_count", 0),
            "pending_count": itsm_scan.get("pending_count", 0),
        },
    }


def engine_sync_all(
    targets: Sequence[str] = ("ics", "outlook", "gcalendar"),
    dry_run: bool = False,
    base_dir: Optional[Path] = None,
    reference_dt: Optional[datetime] = None,
) -> Dict[str, Any]:
    root = base_dir or REPO_ROOT
    ics_dir = _get_ics_dir(root)
    events = get_all_managed_events(root, reference_dt=reference_dt)

    res: Dict[str, Any] = {
        "status": "DRY_RUN" if dry_run else "SUCCESS",
        "synced_at": datetime.now().isoformat(timespec="seconds"),
        "targets": list(targets),
        "events_count": len(events),
    }

    if "ics" in targets:
        res["ics"] = export_events_to_ics(events, out_dir=ics_dir)
    if "outlook" in targets:
        res["outlook"] = sync_events_to_outlook(events, dry_run=dry_run)
    if "gcalendar" in targets:
        res["gcalendar"] = sync_events_to_gcalendar(events, ics_dir=ics_dir, dry_run=dry_run)

    return res


def engine_create_event(
    title: str,
    start_iso: str,
    duration_minutes: int = 30,
    description: str = "",
    location: str = "MES SmartOps — Chervon Vietnam",
    recurrence_type: str = "NONE",
    reminder_minutes: int = 15,
    targets: Sequence[str] = ("ics", "outlook", "gcalendar"),
    dry_run: bool = False,
    base_dir: Optional[Path] = None,
) -> Dict[str, Any]:
    root = base_dir or REPO_ROOT
    start_dt = datetime.fromisoformat(start_iso)
    short_uid = uuid.uuid4().hex[:6].upper()
    event_id = f"CUSTOM_{start_dt.strftime('%Y%m%d_%H%M')}_{short_uid}"

    rrule_map = {
        "NONE": None,
        "DAILY_WORKDAYS": "FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR,SA",
        "WEEKLY_SUN": "FREQ=WEEKLY;BYDAY=SU",
        "MONTHLY_28": "FREQ=MONTHLY;BYMONTHDAY=28",
    }

    ev = SmartOpsCalendarEvent(
        event_id=event_id,
        title=title,
        description=description or f"Sự kiện được đặt lịch tức thì qua DUC_CalendarSync_AIBridge lúc {datetime.now().strftime('%Y-%m-%d %H:%M')}.",
        location=location,
        start_dt=start_dt,
        duration_minutes=int(duration_minutes),
        rrule=rrule_map.get(recurrence_type),
        reminder_minutes=int(reminder_minutes),
        category="MES-SmartOps",
        recurrence_type=recurrence_type,
    )

    if not dry_run:
        save_custom_event(ev, base_dir=root)

    ics_dir = _get_ics_dir(root)
    all_events = get_all_managed_events(root)
    if dry_run:
        all_events.append(ev)

    return {
        "status": "DRY_RUN" if dry_run else "SUCCESS",
        "event": ev.to_dict(),
        "google_web_url": build_google_calendar_url(ev),
        "outlook_web_url": build_outlook_web_url(ev),
        "ics": export_events_to_ics(all_events, out_dir=ics_dir) if "ics" in targets else {},
        "outlook": sync_events_to_outlook([ev], dry_run=dry_run) if "outlook" in targets else {},
        "gcalendar": sync_events_to_gcalendar([ev], ics_dir=ics_dir, dry_run=dry_run) if "gcalendar" in targets else {},
    }


def engine_list_events(
    base_dir: Optional[Path] = None,
    reference_dt: Optional[datetime] = None,
) -> Dict[str, Any]:
    root = base_dir or REPO_ROOT
    ics_dir = _get_ics_dir(root)
    events = get_all_managed_events(root, reference_dt=reference_dt)
    items = []
    for ev in events:
        d = ev.to_dict()
        d["google_web_url"] = build_google_calendar_url(ev)
        d["outlook_web_url"] = build_outlook_web_url(ev)
        d["ics_path"] = str(ics_dir / f"{ev.event_id}.ics")
        items.append(d)
    return {
        "status": "SUCCESS",
        "count": len(items),
        "events": items,
    }


def engine_remind_pending_itsm(
    ledger_path: Optional[Path] = None,
    extra_pending_items: Optional[List[Dict[str, str]]] = None,
    reminder_iso: Optional[str] = None,
    targets: Sequence[str] = ("ics", "outlook", "gcalendar"),
    dry_run: bool = False,
    base_dir: Optional[Path] = None,
) -> Dict[str, Any]:
    root = base_dir or REPO_ROOT
    ics_dir = _get_ics_dir(root)
    return schedule_pending_itsm_reminder(
        ledger_path=ledger_path,
        extra_pending_items=extra_pending_items,
        reminder_iso=reminder_iso,
        targets=targets,
        ics_dir=ics_dir,
        dry_run=dry_run,
    )


def engine_delete_event(
    event_id: str,
    targets: Sequence[str] = ("ics", "outlook", "gcalendar"),
    dry_run: bool = False,
    base_dir: Optional[Path] = None,
) -> Dict[str, Any]:
    root = base_dir or REPO_ROOT
    clean_id = event_id.strip()
    if clean_id in {"DAILY_BACKUP_1745", "WEEKLY_REPORT_SUN_1500", "MONTHLY_ITSM_28_1600"}:
        return {
            "status": "PROTECTED_SYSTEM_EVENT",
            "event_id": clean_id,
            "message": "Chuỗi lịch vận hành cốt lõi được bảo vệ mặc định, không xóa qua lệnh tùy chỉnh.",
        }

    customs = load_custom_events(root)
    target_ev = next((e for e in customs if e.event_id == clean_id), None)
    remaining = [e for e in customs if e.event_id != clean_id]

    if not dry_run and target_ev is not None:
        cfile = _get_custom_events_file(root)
        cfile.parent.mkdir(parents=True, exist_ok=True)
        cfile.write_text(
            json.dumps([e.to_dict() for e in remaining], indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    ics_dir = _get_ics_dir(root)
    ics_file = ics_dir / f"{clean_id}.ics"
    ics_deleted = False
    if "ics" in targets and ics_file.is_file():
        if not dry_run:
            ics_file.unlink(missing_ok=True)
            export_events_to_ics(get_all_managed_events(root), out_dir=ics_dir)
        ics_deleted = True

    outlook_res = (
        delete_event_from_outlook(clean_id, title=target_ev.title if target_ev else "", dry_run=dry_run)
        if "outlook" in targets
        else {}
    )
    gcal_res = (
        delete_event_from_gcalendar(target_ev, dry_run=dry_run)
        if ("gcalendar" in targets and target_ev is not None)
        else {"status": "SKIPPED"}
    )

    return {
        "status": "DRY_RUN" if dry_run else "SUCCESS",
        "event_id": clean_id,
        "found_in_custom_store": target_ev is not None,
        "ics_deleted": ics_deleted,
        "outlook": outlook_res,
        "gcalendar": gcal_res,
    }


def engine_query_schedule(
    date_from_iso: str = "",
    days: int = 7,
    include_live_sources: bool = True,
    base_dir: Optional[Path] = None,
) -> Dict[str, Any]:
    root = base_dir or REPO_ROOT
    start_dt = datetime.fromisoformat(date_from_iso) if date_from_iso.strip() else datetime.now()
    span_days = max(1, min(int(days), 60))
    end_dt = start_dt + timedelta(days=span_days)

    managed = get_all_managed_events(root, reference_dt=start_dt)
    items: List[Dict[str, Any]] = []
    for ev in managed:
        if ev.end_dt >= start_dt and ev.start_dt <= end_dt:
            items.append({
                "source": "smartops_managed",
                "event_id": ev.event_id,
                "title": ev.title,
                "start_dt": ev.start_dt.isoformat(timespec="seconds"),
                "end_dt": ev.end_dt.isoformat(timespec="seconds"),
                "duration_minutes": ev.duration_minutes,
                "location": ev.location,
                "recurrence_type": ev.recurrence_type,
            })

    if include_live_sources:
        win_start_rfc = start_dt.strftime("%Y-%m-%dT00:00:00+07:00")
        win_end_rfc = end_dt.strftime("%Y-%m-%dT23:59:59+07:00")
        items.extend(fetch_gcalendar_events_in_range(win_start_rfc, win_end_rfc))
        items.extend(fetch_outlook_events_in_range(start_dt, end_dt))

    items.sort(key=lambda x: str(x.get("start_dt", "")))
    return {
        "status": "SUCCESS",
        "window_start": start_dt.isoformat(timespec="seconds"),
        "window_end": end_dt.isoformat(timespec="seconds"),
        "days": span_days,
        "total_items": len(items),
        "items": items,
    }


def engine_check_conflicts(
    start_iso: str,
    duration_minutes: int = 30,
    include_live_sources: bool = False,
    base_dir: Optional[Path] = None,
) -> Dict[str, Any]:
    root = base_dir or REPO_ROOT
    req_start = datetime.fromisoformat(start_iso)
    dur = max(5, int(duration_minutes))
    req_end = req_start + timedelta(minutes=dur)

    day_start = req_start.replace(hour=0, minute=0, second=0, microsecond=0)
    managed = get_all_managed_events(root, reference_dt=day_start)

    busy_intervals: List[Dict[str, Any]] = []
    for ev in managed:
        busy_intervals.append({
            "source": "smartops_managed",
            "title": ev.title,
            "start": ev.start_dt,
            "end": ev.end_dt,
        })

    if include_live_sources:
        for ol_ev in fetch_outlook_events_in_range(day_start, day_start + timedelta(days=2)):
            try:
                busy_intervals.append({
                    "source": "outlook_desktop",
                    "title": ol_ev["title"],
                    "start": datetime.fromisoformat(ol_ev["start_dt"]),
                    "end": datetime.fromisoformat(ol_ev["end_dt"]),
                })
            except Exception:
                continue

    conflicts = [
        {
            "source": b["source"],
            "title": b["title"],
            "start_dt": b["start"].isoformat(timespec="seconds"),
            "end_dt": b["end"].isoformat(timespec="seconds"),
        }
        for b in busy_intervals
        if b["end"] > req_start and b["start"] < req_end
    ]

    free_slots: List[Dict[str, str]] = []
    probe = req_start.replace(minute=(req_start.minute // 15) * 15, second=0, microsecond=0)
    if conflicts:
        probe = max(b["end"] for b in busy_intervals if b["end"] > req_start and b["start"] < req_end)
    attempts = 0
    while len(free_slots) < 3 and attempts < 48:
        attempts += 1
        if probe.hour < 8:
            probe = probe.replace(hour=8, minute=0)
        if probe.hour >= 17 and probe.minute > 30:
            probe = (probe + timedelta(days=1)).replace(hour=8, minute=0)
            while probe.weekday() == 6:
                probe += timedelta(days=1)
        cand_end = probe + timedelta(minutes=dur)
        has_overlap = any(b["end"] > probe and b["start"] < cand_end for b in busy_intervals)
        if not has_overlap:
            free_slots.append({
                "start_iso": probe.isoformat(timespec="seconds"),
                "end_iso": cand_end.isoformat(timespec="seconds"),
            })
        probe += timedelta(minutes=30)

    return {
        "status": "CONFLICT_DETECTED" if conflicts else "FREE",
        "has_conflict": bool(conflicts),
        "requested_start": req_start.isoformat(timespec="seconds"),
        "requested_end": req_end.isoformat(timespec="seconds"),
        "conflicts_count": len(conflicts),
        "conflicts": conflicts,
        "suggested_free_slots": free_slots,
    }
