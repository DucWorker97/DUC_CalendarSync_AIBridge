# -*- coding: utf-8 -*-
"""
ics_exporter.py — Bộ Xuất Chuẩn Quốc Tế iCalendar RFC 5545 (.ics) & Deep-Link URLs
===================================================================================
Tương thích 100% với Google Calendar, Microsoft Outlook, Apple Calendar, Thunderbird.
===================================================================================
"""

import hashlib
import urllib.parse
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List

from src.models import SmartOpsCalendarEvent


def _escape_ics_text(val: str) -> str:
    if not val:
        return ""
    return (
        val.replace("\\", "\\\\")
        .replace(";", "\\;")
        .replace(",", "\\,")
        .replace("\r\n", "\\n")
        .replace("\n", "\\n")
    )


def _format_local_dt(dt: datetime) -> str:
    return dt.strftime("%Y%m%dT%H%M%S")


def _format_utc_dt(dt: datetime) -> str:
    vn_tz = timezone(timedelta(hours=7))
    if dt.tzinfo is None:
        dt_aware = dt.replace(tzinfo=vn_tz)
    else:
        dt_aware = dt
    utc_dt = dt_aware.astimezone(timezone.utc)
    return utc_dt.strftime("%Y%m%dT%H%M%SZ")


def render_vevent(ev: SmartOpsCalendarEvent, dtstamp: str) -> List[str]:
    full_desc = f"{ev.description}\n\n{ev.sync_tag}"
    lines = [
        "BEGIN:VEVENT",
        f"UID:{ev.uid}",
        f"DTSTAMP:{dtstamp}",
        f"DTSTART;TZID=Asia/Ho_Chi_Minh:{_format_local_dt(ev.start_dt)}",
        f"DTEND;TZID=Asia/Ho_Chi_Minh:{_format_local_dt(ev.end_dt)}",
        f"SUMMARY:{_escape_ics_text(ev.title)}",
        f"DESCRIPTION:{_escape_ics_text(full_desc)}",
        f"LOCATION:{_escape_ics_text(ev.location)}",
        f"CATEGORIES:{_escape_ics_text(ev.category)}",
        "STATUS:CONFIRMED",
        "TRANSP:OPAQUE",
    ]
    if ev.rrule:
        lines.append(f"RRULE:{ev.rrule}")

    if ev.reminder_minutes and ev.reminder_minutes > 0:
        lines.extend([
            "BEGIN:VALARM",
            f"TRIGGER:-PT{int(ev.reminder_minutes)}M",
            "ACTION:DISPLAY",
            f"DESCRIPTION:Nhắc lịch: {_escape_ics_text(ev.title)}",
            "END:VALARM",
        ])

    lines.append("END:VEVENT")
    return lines


def render_vcalendar(events: List[SmartOpsCalendarEvent]) -> str:
    dtstamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//DucWorker97//DUC_CalendarSync_AIBridge v1.0//VI",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "X-WR-CALNAME:DUC SmartOps — Lịch Vận Hành Tự Động MES/IT",
        "X-WR-TIMEZONE:Asia/Ho_Chi_Minh",
        "BEGIN:VTIMEZONE",
        "TZID:Asia/Ho_Chi_Minh",
        "X-LIC-LOCATION:Asia/Ho_Chi_Minh",
        "BEGIN:STANDARD",
        "TZOFFSETFROM:+0700",
        "TZOFFSETTO:+0700",
        "TZNAME:+07",
        "DTSTART:19700101T000000",
        "END:STANDARD",
        "END:VTIMEZONE",
    ]
    for ev in events:
        lines.extend(render_vevent(ev, dtstamp))
    lines.append("END:VCALENDAR")
    return "\r\n".join(lines) + "\r\n"


def build_google_calendar_url(ev: SmartOpsCalendarEvent) -> str:
    start_utc = _format_utc_dt(ev.start_dt)
    end_utc = _format_utc_dt(ev.end_dt)
    params = {
        "action": "TEMPLATE",
        "text": ev.title,
        "dates": f"{start_utc}/{end_utc}",
        "details": f"{ev.description}\n\n{ev.sync_tag}",
        "location": ev.location,
        "ctz": "Asia/Ho_Chi_Minh",
    }
    if ev.rrule:
        params["recur"] = f"RRULE:{ev.rrule}"
    return "https://calendar.google.com/calendar/render?" + urllib.parse.urlencode(params, quote_via=urllib.parse.quote)


def build_outlook_web_url(ev: SmartOpsCalendarEvent) -> str:
    start_iso = ev.start_dt.strftime("%Y-%m-%dT%H:%M:%S+07:00")
    end_iso = ev.end_dt.strftime("%Y-%m-%dT%H:%M:%S+07:00")
    params = {
        "path": "/calendar/action/compose",
        "rru": "addevent",
        "subject": ev.title,
        "startdt": start_iso,
        "enddt": end_iso,
        "body": f"{ev.description}\n\n{ev.sync_tag}",
        "location": ev.location,
    }
    return "https://outlook.office.com/calendar/0/deeplink/compose?" + urllib.parse.urlencode(params, quote_via=urllib.parse.quote)


def export_events_to_ics(
    events: List[SmartOpsCalendarEvent],
    out_dir: Path,
    master_filename: str = "DUC_SmartOps_Master_Schedule.ics",
) -> Dict[str, Any]:
    out_dir.mkdir(parents=True, exist_ok=True)
    exported_files: List[Dict[str, Any]] = []

    for ev in events:
        ics_text = render_vcalendar([ev])
        fname = f"{ev.event_id}.ics"
        fpath = out_dir / fname
        fpath.write_text(ics_text, encoding="utf-8", newline="")
        md5_hash = hashlib.md5(fpath.read_bytes()).hexdigest()
        exported_files.append({
            "event_id": ev.event_id,
            "filename": fname,
            "path": str(fpath),
            "size_bytes": fpath.stat().st_size,
            "md5": md5_hash,
            "google_web_url": build_google_calendar_url(ev),
            "outlook_web_url": build_outlook_web_url(ev),
        })

    master_text = render_vcalendar(events)
    master_path = out_dir / master_filename
    master_path.write_text(master_text, encoding="utf-8", newline="")
    master_md5 = hashlib.md5(master_path.read_bytes()).hexdigest()

    return {
        "status": "SUCCESS",
        "master_ics_path": str(master_path),
        "master_ics_filename": master_filename,
        "master_md5": master_md5,
        "events_count": len(events),
        "files": exported_files,
    }
