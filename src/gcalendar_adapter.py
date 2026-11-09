# -*- coding: utf-8 -*-
"""
gcalendar_adapter.py — Cầu nối Đồng bộ Google Calendar (gcalendar OAuth CLI + Web URL + ICS)
===========================================================================================
Tự động đẩy sự kiện trực tiếp lên Google Calendar qua `gcalendar` CLI (với chống trùng lặp
Idempotent) và luôn cung cấp liên kết 1-Click Google Calendar Web URL kèm RRULE.
===========================================================================================
"""

import json
import shutil
import subprocess
from datetime import timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional

from src.models import SmartOpsCalendarEvent
from src.ics_exporter import build_google_calendar_url


def locate_gcalendar_cli() -> Optional[str]:
    """Chỉ tìm kiếm file thực thi nhị phân (.exe) thực sự đã cài đặt,
    tránh gọi các wrapper .cmd chưa tải binary gây treo subprocess.
    """
    in_path_exe = shutil.which("gcalendar.exe")
    if in_path_exe and Path(in_path_exe).is_file():
        return in_path_exe

    plugin_bin_dir = Path.home() / ".gemini" / "antigravity" / "plugin_data" / "gcalendar" / "bin"
    if plugin_bin_dir.is_dir():
        exes = sorted(plugin_bin_dir.glob("gcalendar*.exe"), reverse=True)
        if exes:
            return str(exes[0])

    local_exe = Path.home() / "AppData" / "Local" / "gcalendar" / "gcalendar.exe"
    if local_exe.is_file():
        return str(local_exe)

    return None


def _find_existing_gcal_event(cli_path: str, ev: SmartOpsCalendarEvent) -> Optional[Dict[str, Any]]:
    try:
        win_start = (ev.start_dt - timedelta(days=1)).strftime("%Y-%m-%dT00:00:00+07:00")
        win_end = (ev.start_dt + timedelta(days=32)).strftime("%Y-%m-%dT23:59:59+07:00")
        query_term = "ITSM-Remind" if "[ITSM-Remind]" in ev.title else "MES-Ops"
        proc = subprocess.run(
            [
                cli_path, "readonly", "search", query_term,
                "--start", win_start,
                "--end", win_end,
                "--max", "50",
                "--json",
            ],
            capture_output=True,
            text=True,
            timeout=10,
            encoding="utf-8",
            errors="ignore",
        )
        if proc.returncode != 0 or not proc.stdout:
            return None
        data = json.loads(proc.stdout)
        events_list = data if isinstance(data, list) else data.get("items", data.get("events", []))
        if not isinstance(events_list, list):
            return None
        for item in events_list:
            desc = str(item.get("description", ""))
            summary = str(item.get("summary", ""))
            if ev.sync_tag in desc or ev.event_id in desc or summary == ev.title:
                return {
                    "id": item.get("recurringEventId") or item.get("id", ""),
                    "htmlLink": item.get("htmlLink", ""),
                    "summary": summary,
                }
    except Exception:
        pass
    return None


def get_gcalendar_status() -> Dict[str, Any]:
    cli_path = locate_gcalendar_cli()
    if not cli_path:
        return {
            "cli_installed": False,
            "authenticated": False,
            "mode": "WEB_LINK_AND_ICS",
            "message": "Sử dụng chế độ 1-Click Google Calendar Web URL & tệp .ics chuẩn RFC 5545.",
        }

    try:
        ver_proc = subprocess.run(
            [cli_path, "--version"],
            capture_output=True,
            text=True,
            timeout=8,
            encoding="utf-8",
            errors="ignore",
        )
        version_str = (ver_proc.stdout or ver_proc.stderr or "").strip().splitlines()[0] if (ver_proc.stdout or ver_proc.stderr) else "unknown"

        auth_proc = subprocess.run(
            [cli_path, "readonly", "calendars", "--json"],
            capture_output=True,
            text=True,
            timeout=8,
            encoding="utf-8",
            errors="ignore",
        )
        combined_out = ((auth_proc.stdout or "") + "\n" + (auth_proc.stderr or "")).lower()
        is_authed = (
            auth_proc.returncode == 0
            and "needs authentication" not in combined_out
            and "no credentials found" not in combined_out
            and ("[" in combined_out or "{" in combined_out)
        )
        return {
            "cli_installed": True,
            "cli_path": cli_path,
            "version": version_str,
            "authenticated": is_authed,
            "mode": "DIRECT_OAUTH_SYNC" if is_authed else "WEB_LINK_AND_ICS_READY",
            "message": (
                "Đã kết nối OAuth Google Calendar trực tiếp."
                if is_authed
                else "Gói .ics & 1-Click Google Calendar URL đã sẵn sàng."
            ),
        }
    except Exception as e:
        return {
            "cli_installed": True,
            "cli_path": cli_path,
            "authenticated": False,
            "mode": "WEB_LINK_AND_ICS_READY",
            "error": str(e),
        }


def sync_events_to_gcalendar(
    events: List[SmartOpsCalendarEvent],
    ics_dir: Optional[Path] = None,
    dry_run: bool = False,
) -> Dict[str, Any]:
    status_info = get_gcalendar_status()
    cli_path = status_info.get("cli_path")
    is_authed = bool(status_info.get("authenticated", False))

    items: List[Dict[str, Any]] = []
    direct_synced = 0

    for ev in events:
        web_url = build_google_calendar_url(ev)
        ics_file = str(ics_dir / f"{ev.event_id}.ics") if ics_dir else None
        item_res: Dict[str, Any] = {
            "event_id": ev.event_id,
            "title": ev.title,
            "google_web_url": web_url,
            "ics_file": ics_file,
            "action": "READY_WEB_AND_ICS",
        }

        if not dry_run and is_authed and cli_path:
            try:
                existing = _find_existing_gcal_event(cli_path, ev)
                if existing:
                    item_res["action"] = "ALREADY_SYNCED_ON_GCAL"
                    item_res["gcal_event_id"] = existing.get("id", "")
                    item_res["gcal_html_link"] = existing.get("htmlLink", "")
                    direct_synced += 1
                    items.append(item_res)
                    continue

                start_rfc = ev.start_dt.strftime("%Y-%m-%dT%H:%M:%S+07:00")
                end_rfc = ev.end_dt.strftime("%Y-%m-%dT%H:%M:%S+07:00")
                clean_desc = " | ".join(line.strip() for line in ev.description.splitlines() if line.strip())
                full_desc = f"{clean_desc} | Địa điểm: {ev.location} | {ev.sync_tag}"

                if ev.rrule:
                    cmd = [
                        cli_path, "mutate", "recurring", ev.title,
                        "--start", start_rfc,
                        "--end", end_rfc,
                        "--rrule", f"RRULE:{ev.rrule}",
                        "--description", full_desc,
                        "--send-updates", "none",
                    ]
                else:
                    cmd = [
                        cli_path, "mutate", "create",
                        "--summary", ev.title,
                        "--start", start_rfc,
                        "--end", end_rfc,
                        "--description", full_desc,
                        "--location", ev.location,
                        "--send-updates", "none",
                    ]

                proc = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    timeout=15,
                    encoding="utf-8",
                    errors="ignore",
                )
                if proc.returncode == 0:
                    item_res["action"] = "SYNCED_VIA_OAUTH_CLI"
                    item_res["cli_output"] = (proc.stdout or "").strip()
                    direct_synced += 1
                else:
                    item_res["action"] = "READY_WEB_AND_ICS"
                    item_res["cli_note"] = (proc.stderr or proc.stdout or "").strip()[:200]
            except Exception as e:
                item_res["action"] = "READY_WEB_AND_ICS"
                item_res["cli_note"] = str(e)

        items.append(item_res)

    return {
        "status": "SUCCESS",
        "mode": status_info.get("mode", "WEB_LINK_AND_ICS_READY"),
        "cli_authenticated": is_authed,
        "direct_synced_count": direct_synced,
        "total_events": len(items),
        "items": items,
    }
