# -*- coding: utf-8 -*-
"""
calendar_engine.py — Bộ Điều Phối Trung Tâm của DUC_CalendarSync_AIBridge
=========================================================================
Cung cấp các hàm nghiệp vụ cốt lõi cho MCP Server:
1. cal_status: Kiểm tra trạng thái Google Calendar OAuth, Outlook Desktop MAPI, kho .ics và sổ cái ITSM.
2. cal_sync_all: Đồng bộ toàn bộ lịch vận hành tự động + lịch tùy chỉnh lên cả 3 nền tảng.
3. cal_create_event: Đặt lịch tức thì lên Google Calendar, Outlook Desktop & xuất .ics.
4. cal_remind_pending_itsm: Quét các Candidate Bxx chờ bù đơn / thiếu Mã NV và tự động đặt lịch nhắc.
5. cal_list_events: Liệt kê toàn bộ lịch trình kèm liên kết 1-Click.
=========================================================================
"""

import json
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from src.models import SmartOpsCalendarEvent, get_default_smartops_events
from src.ics_exporter import (
    export_events_to_ics,
    build_google_calendar_url,
    build_outlook_web_url,
)
from src.outlook_adapter import get_outlook_calendar_status, sync_events_to_outlook
from src.gcalendar_adapter import get_gcalendar_status, sync_events_to_gcalendar
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
        "version": "1.0.0",
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
