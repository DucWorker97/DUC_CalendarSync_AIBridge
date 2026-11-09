# -*- coding: utf-8 -*-
"""
outlook_adapter.py — Cầu nối Đồng bộ Trực tiếp Microsoft Outlook Desktop Calendar (COM MAPI)
===========================================================================================
Tự động tạo hoặc cập nhật (Idempotent) sự kiện trên Microsoft Outlook Desktop.
Đảm bảo `MeetingStatus = 0` (olNonMeeting) để chỉ nhắc lịch cá nhân, không gửi email ra ngoài.
===========================================================================================
"""

from datetime import datetime
from typing import Any, Dict, List, Optional

from src.models import SmartOpsCalendarEvent
from src.ics_exporter import build_outlook_web_url

OL_FOLDER_CALENDAR = 9
OL_APPOINTMENT_ITEM = 1
OL_NON_MEETING = 0
OL_BUSY = 2

OL_RECURS_WEEKLY = 1
OL_RECURS_MONTHLY = 2

MASK_WEEKDAYS_MON_TO_SAT = 2 + 4 + 8 + 16 + 32 + 64  # 126
MASK_SUNDAY = 1


def _get_active_outlook() -> Any:
    """Kết nối non-blocking tới tiến trình Outlook Desktop đang mở qua GetActiveObject,
    tránh gọi Dispatch mù quáng làm treo tiến trình nền khi Outlook chưa khởi động.
    """
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    return win32com.client.GetActiveObject("Outlook.Application")


def get_outlook_calendar_status() -> Dict[str, Any]:
    try:
        outlook = _get_active_outlook()
        ns = outlook.GetNamespace("MAPI")
        cal_folder = ns.GetDefaultFolder(OL_FOLDER_CALENDAR)
        items = cal_folder.Items
        total_count = items.Count

        smartops_count = 0
        max_scan = min(total_count, 300)
        for idx in range(total_count, max(0, total_count - max_scan), -1):
            try:
                it = items.Item(idx)
                body = str(getattr(it, "Body", "") or "")
                subj = str(getattr(it, "Subject", "") or "")
                if "[SMARTOPS-ID:" in body or subj.startswith("[MES-Ops]") or subj.startswith("[ITSM-Remind]"):
                    smartops_count += 1
            except Exception:
                continue

        return {
            "available": True,
            "connected": True,
            "folder_name": str(cal_folder.Name),
            "total_items": total_count,
            "smartops_synced_count": smartops_count,
        }
    except Exception as e:
        return {
            "available": False,
            "connected": False,
            "error": str(e),
        }


def _find_existing_appointment(cal_folder: Any, sync_tag: str, title: str) -> Optional[Any]:
    try:
        items = cal_folder.Items
        total = items.Count
        max_scan = min(total, 400)
        for idx in range(total, max(0, total - max_scan), -1):
            try:
                it = items.Item(idx)
                body = str(getattr(it, "Body", "") or "")
                subj = str(getattr(it, "Subject", "") or "")
                if sync_tag in body or subj == title:
                    return it
            except Exception:
                continue
    except Exception:
        pass
    return None


def _apply_recurrence(appt: Any, ev: SmartOpsCalendarEvent) -> None:
    if not ev.rrule or ev.recurrence_type == "NONE":
        return
    try:
        pat = appt.GetRecurrencePattern()
        if ev.recurrence_type == "DAILY_WORKDAYS":
            pat.RecurrenceType = OL_RECURS_WEEKLY
            pat.Interval = 1
            pat.DayOfWeekMask = MASK_WEEKDAYS_MON_TO_SAT
            pat.NoEndDate = True
        elif ev.recurrence_type == "WEEKLY_SUN":
            pat.RecurrenceType = OL_RECURS_WEEKLY
            pat.Interval = 1
            pat.DayOfWeekMask = MASK_SUNDAY
            pat.NoEndDate = True
        elif ev.recurrence_type == "MONTHLY_28":
            pat.RecurrenceType = OL_RECURS_MONTHLY
            pat.Interval = 1
            pat.DayOfMonth = 28
            pat.NoEndDate = True
    except Exception:
        pass


def sync_events_to_outlook(
    events: List[SmartOpsCalendarEvent],
    dry_run: bool = False,
) -> Dict[str, Any]:
    if dry_run:
        return {
            "status": "DRY_RUN",
            "available": True,
            "synced_count": len(events),
            "items": [
                {
                    "event_id": ev.event_id,
                    "title": ev.title,
                    "action": "WOULD_UPSERT_OUTLOOK",
                    "outlook_web_url": build_outlook_web_url(ev),
                }
                for ev in events
            ],
        }

    try:
        outlook = _get_active_outlook()
        ns = outlook.GetNamespace("MAPI")
        cal_folder = ns.GetDefaultFolder(OL_FOLDER_CALENDAR)
    except Exception as e:
        return {
            "status": "FALLBACK_WEB_LINK",
            "available": False,
            "error": str(e),
            "items": [
                {
                    "event_id": ev.event_id,
                    "title": ev.title,
                    "action": "OUTLOOK_WEB_DEEPLINK_READY",
                    "outlook_web_url": build_outlook_web_url(ev),
                }
                for ev in events
            ],
        }

    results: List[Dict[str, Any]] = []
    created_count = 0
    updated_count = 0

    for ev in events:
        try:
            existing = _find_existing_appointment(cal_folder, ev.sync_tag, ev.title)
            if existing is not None:
                appt = existing
                action = "UPDATED_EXISTING"
                updated_count += 1
            else:
                appt = outlook.CreateItem(OL_APPOINTMENT_ITEM)
                action = "CREATED_NEW"
                created_count += 1

            appt.Subject = ev.title
            appt.Body = f"{ev.description}\r\n\r\n{ev.sync_tag}\r\nCập nhật tự động lúc: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            appt.Location = ev.location
            appt.MeetingStatus = OL_NON_MEETING
            appt.BusyStatus = OL_BUSY
            appt.Categories = ev.category

            if ev.reminder_minutes and ev.reminder_minutes > 0:
                appt.ReminderSet = True
                appt.ReminderMinutesBeforeStart = int(ev.reminder_minutes)
            else:
                appt.ReminderSet = False

            if action == "CREATED_NEW" or not getattr(appt, "IsRecurring", False):
                appt.Start = ev.start_dt.strftime("%Y-%m-%d %H:%M:%S")
                appt.Duration = int(ev.duration_minutes)
                _apply_recurrence(appt, ev)

            appt.Save()
            results.append({
                "event_id": ev.event_id,
                "title": ev.title,
                "action": action,
                "start_dt": ev.start_dt.isoformat(timespec="seconds"),
                "recurrence": ev.recurrence_type,
                "outlook_web_url": build_outlook_web_url(ev),
            })
        except Exception as item_err:
            results.append({
                "event_id": ev.event_id,
                "title": ev.title,
                "action": "ERROR_FALLBACK_LINK",
                "error": str(item_err),
                "outlook_web_url": build_outlook_web_url(ev),
            })

    return {
        "status": "SUCCESS",
        "available": True,
        "folder_name": str(cal_folder.Name),
        "created_count": created_count,
        "updated_count": updated_count,
        "synced_count": created_count + updated_count,
        "items": results,
    }
