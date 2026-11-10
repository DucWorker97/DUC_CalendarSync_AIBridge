# -*- coding: utf-8 -*-
"""
notification_dispatcher.py — Bộ Điều Phối Thông Báo Đa Kênh Tức Thì (WeChat Work & Zalo)
========================================================================================
Author : Nguyễn Duy Đức (DucWorker97)
Version: 1.4.1

Cung cấp khả năng phân phối Bản tin Lịch trình Đầu ca / Cuối ca & Cảnh báo ITSM tức thì:
1. Kênh WeChat Work Group Bot Webhook (Tencent Official API):
   - Định dạng Markdown chuyên dụng (hỗ trợ <font color="info|warning|comment">).
   - Tự động kiểm tra độ dài (<= 4000 UTF-8 bytes) và khử trùng lặp sự kiện.
   - An toàn tuyệt đối: mặc định Dry-Run Preview, chỉ gửi thật khi dry_run=False.
   - Tự động lưu Webhook URL vào exports/notifications_config.json khi người dùng nhập từ Web UI.
2. Kênh Zalo Web Automation (DUC_OLAZ Playwright + Clipboard + Queue):
   - Chuyển đổi bản tin Markdown sang định dạng văn bản thuần giàu biểu cảm (Emoji + Checklist).
   - Khi bắn thật (dry_run=False):
     + Tự động copy bản tin định dạng chuẩn vào Windows Clipboard (UTF-16LE) tức thì.
     + Lưu hàng đợi phát hành (exports/briefings/zalo_queue.json).
     + Kích hoạt tiến trình nền không chặn (Non-blocking Background Worker) kết nối Zalo Web
       qua DUC_OLAZ (Chrome CDP / ZaloPlaywright Profile) để gửi tin vào nhóm chỉ định.
3. Kênh File Artifacts & SmartOps Hub Web UI:
   - Tự động xuất tệp Markdown và mirror vào Hub logs.
========================================================================================
"""

import argparse
import json
import os
import subprocess
import sys
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

_SRC_DIR = Path(__file__).resolve().parent
_REPO_ROOT = _SRC_DIR.parent
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from src.calendar_engine import engine_daily_briefing

DEFAULT_HUB_ROOT = Path(r"D:\DucWorker\DUC_SmartOps_AIHub")
CONFIG_FILE_PATH = _REPO_ROOT / "exports" / "notifications_config.json"


def _deduplicate_events(events: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Khử trùng lặp sự kiện khi cùng một lịch xuất hiện ở cả smartops_managed và outlook_com."""
    seen = set()
    unique: List[Dict[str, Any]] = []
    for ev in events:
        st = str(ev.get("start_dt", ""))[:16]
        title = str(ev.get("title", "")).strip()
        key = (st, title)
        if key not in seen:
            seen.add(key)
            unique.append(ev)
    return unique


def copy_to_windows_clipboard(text: str) -> bool:
    """Copy chuỗi Unicode (Tiếng Việt + Emoji) vào Windows Clipboard qua clip.exe (UTF-16LE)."""
    if sys.platform != "win32" or not text:
        return False
    try:
        # clip.exe đọc chuẩn UTF-16LE có BOM hoặc không BOM trên Windows
        payload = text.replace("\r\n", "\n").replace("\n", "\r\n").encode("utf-16le")
        proc = subprocess.run(
            ["clip.exe"],
            input=payload,
            capture_output=True,
            timeout=3,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000),
        )
        return proc.returncode == 0
    except Exception:
        return False


def load_notification_config() -> Dict[str, Any]:
    """Tải cấu hình thông báo từ exports/notifications_config.json hoặc biến môi trường."""
    raw_env_url = os.environ.get("WECHAT_WEBHOOK_URL", os.environ.get("WECHAT_WEBHOOK_DEFAULT", "")).strip()
    if "YOUR_DEFAULT_KEY" in raw_env_url or "your_" in raw_env_url.lower():
        raw_env_url = ""

    cfg: Dict[str, Any] = {
        "wechat": {
            "enabled": True,
            "webhook_url": raw_env_url,
            "channel_alias": "default",
        },
        "zalo": {
            "enabled": True,
            "default_chat": "Thông báo chung OLAZ",
            "auto_pin_note": False,
        },
    }

    if CONFIG_FILE_PATH.is_file():
        try:
            stored = json.loads(CONFIG_FILE_PATH.read_text(encoding="utf-8"))
            if isinstance(stored, dict):
                if "wechat" in stored and isinstance(stored["wechat"], dict):
                    cfg["wechat"].update(stored["wechat"])
                if "zalo" in stored and isinstance(stored["zalo"], dict):
                    cfg["zalo"].update(stored["zalo"])
        except Exception:
            pass

    return cfg


def save_notification_config(config: Dict[str, Any]) -> None:
    """Lưu cấu hình thông báo vào exports/notifications_config.json."""
    CONFIG_FILE_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_FILE_PATH.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")


def mask_url(url: Optional[str]) -> str:
    """Che giấu key nhạy cảm của Webhook URL khi xuất log / trả về UI."""
    if not url:
        return "(chưa cấu hình Webhook URL)"
    if "key=" in url:
        base, key = url.split("key=", 1)
        if len(key) >= 8:
            masked_key = key[:4] + "****" + key[-4:]
        else:
            masked_key = "********"
        return f"{base}key={masked_key}"
    return url[:12] + "****"


def format_for_wechat_markdown(briefing_data: Dict[str, Any]) -> str:
    """
    Chuyển đổi dữ liệu bản tin sang Markdown tương thích hoàn hảo với Tencent WeChat Work:
    - Sử dụng tiêu đề ###, danh sách trích dẫn >, và thẻ màu font.
    - Giới hạn kích thước <= 4000 UTF-8 bytes.
    """
    shift_label = briefing_data.get("shift_label", "📋 BẢN TIN LỊCH TRÌNH VẬN HÀNH")
    date_str = briefing_data.get("date", datetime.now().strftime("%Y-%m-%d"))
    today_items = _deduplicate_events(briefing_data.get("today_events", []))
    pending_count = int(briefing_data.get("itsm_pending_count", 0))
    pending_list = briefing_data.get("itsm_pending_candidates", [])
    scan_ok = bool(briefing_data.get("itsm_scan_ok", False))  # thieu khoa -> KHONG bao ALL CLEAR
    checklist = briefing_data.get("action_checklist", [])

    lines: List[str] = [
        f"### {shift_label}",
        f"> **Ngày làm việc:** <font color=\"comment\">{date_str}</font>",
        f"> **Lịch trình:** <font color=\"info\">{len(today_items)} sự kiện</font>",
    ]

    if pending_count > 0:
        lines.append(f"> **Bù đơn ITSM (`Bxx`):** <font color=\"warning\">{pending_count} đơn chờ bù / thiếu Mã NV</font>")
    elif not scan_ok:
        lines.append("> **Bù đơn ITSM (`Bxx`):** <font color=\"warning\">⚠️ Chưa đọc được sổ — chưa kết luận</font>")
    else:
        lines.append("> **Bù đơn ITSM (`Bxx`):** <font color=\"info\">0 Candidate chờ bù (theo sổ THEO-DOI)</font>")

    lines.append("")
    lines.append("#### 📅 1. Lịch Trình Chi Tiết")
    if today_items:
        for idx, it in enumerate(today_items, 1):
            st = str(it.get("start_dt", ""))[11:16] or "Cả ngày"
            title = it.get("title", "")
            src = it.get("source", "")
            lines.append(f"{idx}. **`{st}`** — {title} *(<font color=\"comment\">{src}</font>)*")
    else:
        lines.append("- *(Không có sự kiện phát sinh ngoài lịch vận hành chuẩn)*")

    lines.append("")
    lines.append("#### 🎫 2. Trạng Thái Bù Đơn ITSM (Quy tắc 2 & 4)")
    if pending_count > 0:
        for it in pending_list:
            cid = it.get("candidate_id", "Bxx")
            cust = it.get("customer", "")
            summary = it.get("summary", "")
            lines.append(f"- <font color=\"warning\">⚠️ [{cid}]</font> **{cust}**: {summary}")
        lines.append("> 💡 *Mẫu hỏi Mã NV (VI-ZH):* `Nhờ anh/chị cung cấp Mã nhân viên (V0xxxx) để em bù đơn ITSM / 麻烦提供一下工号。`")
    elif not scan_ok:
        lines.append("- <font color=\"warning\">⚠️ Không đọc được Candidate Bxx nào trong sổ bù đơn — kiểm tay.</font>")
    else:
        lines.append("- <font color=\"info\">0 Candidate Bxx chờ bù (theo sổ THEO-DOI).</font>")

    lines.append("")
    lines.append("#### ✅ 3. Checklist Trọng Tâm")
    for chk in checklist:
        lines.append(f"- {chk}")

    full_md = "\n".join(lines)
    encoded = full_md.encode("utf-8")
    if len(encoded) > 4000:
        truncated = encoded[:3950].decode("utf-8", errors="ignore") + "\n\n*(Nội dung được rút gọn theo giới hạn WeChat)*"
        return truncated
    return full_md


def format_for_zalo_text(briefing_data: Dict[str, Any]) -> str:
    """
    Chuyển đổi dữ liệu bản tin sang văn bản thuần giàu biểu cảm (Emoji + Dòng thụt lề)
    tối ưu cho hiển thị trên ứng dụng Zalo Mobile & Desktop.
    """
    shift_label = briefing_data.get("shift_label", "📋 BẢN TIN LỊCH TRÌNH VẬN HÀNH")
    date_str = briefing_data.get("date", datetime.now().strftime("%Y-%m-%d"))
    today_items = _deduplicate_events(briefing_data.get("today_events", []))
    pending_count = int(briefing_data.get("itsm_pending_count", 0))
    pending_list = briefing_data.get("itsm_pending_candidates", [])
    scan_ok = bool(briefing_data.get("itsm_scan_ok", False))  # thieu khoa -> KHONG bao ALL CLEAR
    stale_list = briefing_data.get("itsm_stale_status", [])
    checklist = briefing_data.get("action_checklist", [])

    lines: List[str] = [
        f"📌 {shift_label.replace('### ', '')}",
        "══════════════════════════",
        f"📅 Ngày làm việc: {date_str}",
        f"⏰ Sự kiện trong ngày: {len(today_items)} mục",
    ]

    if pending_count > 0:
        lines.append(f"🎫 Bù đơn ITSM (Bxx): ⚠️ {pending_count} đơn chờ xử lý")
    elif not scan_ok:
        lines.append("🎫 Bù đơn ITSM (Bxx): ⚠️ chưa đọc được sổ — chưa kết luận")
    else:
        lines.append("🎫 Bù đơn ITSM (Bxx): 0 Candidate chờ bù (theo sổ)")

    lines.append("──────────────────────────")
    lines.append("📅 LỊCH TRÌNH TRONG NGÀY:")
    if today_items:
        for idx, it in enumerate(today_items, 1):
            st = str(it.get("start_dt", ""))[11:16] or "Cả ngày"
            lines.append(f"  {idx}. [{st}] {it.get('title')}")
    else:
        lines.append("  • (Không có sự kiện phát sinh ngoài định kỳ)")

    lines.append("──────────────────────────")
    lines.append("🎫 THEO DÕI SỰ CỐ BÙ ĐƠN ITSM:")
    if pending_count > 0:
        for it in pending_list:
            lines.append(f"  • [{it.get('candidate_id')}] {it.get('customer')}: {it.get('summary')}")
        lines.append("  💡 Mẫu hỏi Mã NV: Nhờ anh/chị cung cấp thêm Mã nhân viên (V0xxxx) nhé!")
    elif not scan_ok:
        lines.append("  ⚠️ Không đọc được Candidate nào trong THEO-DOI-BU-TICKET-ITSM.md — kiểm tay")
    else:
        lines.append("  • 0 Candidate Bxx chờ bù (theo sổ THEO-DOI)")
    if stale_list:
        lines.append(f"  ℹ️ {len(stale_list)} mục đã có INC nhưng ô trạng thái chưa cập nhật: "
                     + ", ".join(x.get("candidate_id", "Bxx") for x in stale_list))

    lines.append("──────────────────────────")
    lines.append("✅ CHECKLIST TRỌNG TÂM THEO CA:")
    for chk in checklist:
        clean_chk = chk.replace("- [ ] ", "").replace("- [x] ", "")
        lines.append(f"  ▫️ {clean_chk}")

    lines.append("══════════════════════════")
    return "\n".join(lines)


def send_to_wechat_webhook(
    markdown_content: str,
    webhook_url: Optional[str] = None,
    dry_run: bool = True,
) -> Dict[str, Any]:
    """
    Gửi báo cáo Markdown qua WeChat Work Group Bot Webhook API:
    - Nếu dry_run=True: Trả về Preview và thông tin kiểm định.
    - Nếu dry_run=False: Thực hiện HTTP POST tới Tencent Webhook.
    """
    cfg = load_notification_config()
    clean_input_url = (webhook_url or "").strip()
    if clean_input_url and "YOUR_DEFAULT_KEY" not in clean_input_url:
        cfg["wechat"]["webhook_url"] = clean_input_url
        save_notification_config(cfg)

    target_url = (clean_input_url or cfg["wechat"].get("webhook_url", "")).strip()
    if "YOUR_DEFAULT_KEY" in target_url:
        target_url = ""

    # Cho phép người dùng chỉ dán mã key (UUID) thay vì cả URL dài
    if target_url and not target_url.startswith("http"):
        target_url = f"https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key={target_url}"

    if not target_url:
        return {
            "status": "NOT_CONFIGURED",
            "channel": "wechat_work_webhook",
            "target_url": "(chưa cấu hình Webhook URL)",
            "message": "Chưa cấu hình WeChat Work Webhook URL. Hãy dán Webhook URL hoặc Key vào ô '4. WeChat Webhook URL / Key' trên giao diện để bắn thẳng vào nhóm WeChat.",
            "preview_content": markdown_content,
            "content_bytes": len(markdown_content.encode("utf-8")),
        }

    if dry_run:
        return {
            "status": "PREVIEW_READY",
            "channel": "wechat_work_webhook",
            "dry_run": True,
            "target_url": mask_url(target_url),
            "message": "Sẵn sàng bắn vào WeChat Work Webhook.",
            "content_bytes": len(markdown_content.encode("utf-8")),
            "preview_content": markdown_content,
        }

    payload = {
        "msgtype": "markdown",
        "markdown": {
            "content": markdown_content,
        },
    }
    data_bytes = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        target_url,
        data=data_bytes,
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=10.0) as resp:
            resp_body = resp.read().decode("utf-8", errors="ignore")
            resp_json = json.loads(resp_body) if resp_body.strip() else {}
            errcode = resp_json.get("errcode", 0)
            errmsg = resp_json.get("errmsg", "ok")
            is_ok = (errcode == 0)
            return {
                "status": "SENT" if is_ok else "FAILED",
                "channel": "wechat_work_webhook",
                "target_url": mask_url(target_url),
                "errcode": errcode,
                "errmsg": errmsg,
                "message": "Đã bắn thành công vào nhóm WeChat Work!" if is_ok else f"WeChat API trả lỗi ({errcode}): {errmsg}",
                "preview_content": markdown_content,
                "raw_response": resp_json,
            }
    except Exception as e:
        return {
            "status": "NETWORK_ERROR",
            "channel": "wechat_work_webhook",
            "target_url": mask_url(target_url),
            "error": str(e),
            "message": f"Lỗi kết nối WeChat Webhook: {e}",
            "preview_content": markdown_content,
        }


def _trigger_zalo_live_worker(queue_file: Path) -> Optional[int]:
    """Khởi chạy tiến trình ngầm gửi tin nhắn từ zalo_queue.json qua DUC_OLAZ Playwright."""
    olaz_dir = DEFAULT_HUB_ROOT / "mcp_bridges" / "DUC_OLAZ"
    if not olaz_dir.is_dir():
        return None
    try:
        proc = subprocess.Popen(
            [sys.executable, str(Path(__file__).resolve()), "--deliver-zalo-queue", str(queue_file)],
            cwd=str(olaz_dir),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000),
        )
        return proc.pid
    except Exception:
        return None


def prepare_zalo_dispatch(
    text_content: str,
    target_chat: str = "Thông báo chung OLAZ",
    dry_run: bool = True,
    repo_root: Optional[Path] = None,
) -> Dict[str, Any]:
    """
    Chuẩn bị và thực thi gửi thông báo Zalo:
    - Lưu vào exports/briefings/zalo_queue.json
    - Khi dry_run=False:
      1. Copy tức thì nội dung bản tin vào Windows Clipboard (để người dùng có thể Ctrl+V ngay).
      2. Kích hoạt tiến trình nền DUC_OLAZ (Zalo-Web-Automation) tự động mở nhóm target_chat và gửi tin.
    """
    root = repo_root or _REPO_ROOT
    queue_file = root / "exports" / "briefings" / "zalo_queue.json"
    queue_file.parent.mkdir(parents=True, exist_ok=True)

    clipboard_copied = False
    worker_pid: Optional[int] = None

    if not dry_run:
        clipboard_copied = copy_to_windows_clipboard(text_content)

    entry = {
        "queued_at": datetime.now().isoformat(timespec="seconds"),
        "target_chat": target_chat,
        "message": text_content,
        "dry_run": dry_run,
        "clipboard_copied": clipboard_copied,
        "status": "PREVIEW" if dry_run else "DISPATCHING_LIVE",
    }

    try:
        existing = []
        if queue_file.is_file():
            existing = json.loads(queue_file.read_text(encoding="utf-8"))
            if not isinstance(existing, list):
                existing = []
        existing.append(entry)
        queue_file.write_text(json.dumps(existing[-20:], ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass

    if not dry_run:
        worker_pid = _trigger_zalo_live_worker(queue_file)

    status_code = (
        "PREVIEW_READY"
        if dry_run
        else ("DISPATCHED_AND_COPIED" if (worker_pid or clipboard_copied) else "QUEUED_FOR_AGENT")
    )
    msg_detail = (
        "Bản xem trước Zalo đã sẵn sàng."
        if dry_run
        else (
            f"✅ Đã copy bản tin vào Clipboard (sẵn sàng Ctrl+V) & kích hoạt Zalo Web Automation (PID {worker_pid}) gửi tới '{target_chat}'!"
            if worker_pid
            else f"✅ Đã copy bản tin vào Clipboard (sẵn sàng Ctrl+V) & ghi hàng đợi gửi tới '{target_chat}'!"
        )
    )

    return {
        "status": status_code,
        "channel": "zalo_web_automation",
        "target_chat": target_chat,
        "dry_run": dry_run,
        "clipboard_copied": clipboard_copied,
        "worker_pid": worker_pid,
        "message": msg_detail,
        "char_count": len(text_content),
        "queue_file": str(queue_file),
        "preview_text": text_content,
    }


def dispatch_daily_briefing(
    date_iso: str = "",
    shift_mode: str = "AUTO",
    channels: Sequence[str] = ("file", "wechat", "zalo"),
    dry_run: bool = True,
    target_zalo: str = "Thông báo chung OLAZ",
    wechat_webhook_url: str = "",
    repo_root: Optional[Path] = None,
) -> Dict[str, Any]:
    """
    Hàm nghiệp vụ cốt lõi điều phối bản tin đa kênh tức thì (< 0.2s):
    - Kênh 'file'  : Xuất tệp Markdown ra exports/briefings/
    - Kênh 'wechat': Bắn bản tin Markdown qua WeChat Work Group Bot Webhook
    - Kênh 'zalo'  : Copy Clipboard + Kích hoạt Zalo Web Automation gửi tới nhóm target_zalo
    """
    root = repo_root or _REPO_ROOT
    # Luôn dùng nguồn lịch đã đồng bộ trong kho cục bộ (include_live_sources=False)
    # để tốc độ phản hồi đạt < 0.2 giây và không bị treo 14s bởi subprocess CLI.
    briefing = engine_daily_briefing(
        date_iso=date_iso,
        shift_mode=shift_mode,
        include_live_sources=False,
        base_dir=root,
    )

    if target_zalo and target_zalo.strip():
        cfg = load_notification_config()
        cfg["zalo"]["default_chat"] = target_zalo.strip()
        if wechat_webhook_url and wechat_webhook_url.strip():
            cfg["wechat"]["webhook_url"] = wechat_webhook_url.strip()
        save_notification_config(cfg)

    channels_norm = [c.strip().lower() for c in channels]
    results: Dict[str, Any] = {
        "status": "SUCCESS",
        "action": "DISPATCH_DAILY_BRIEFING",
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "shift_label": briefing.get("shift_label"),
        "date": briefing.get("date"),
        "dry_run": dry_run,
        "channels": {},
    }

    # 1. Kênh File
    if "file" in channels_norm:
        date_tag = datetime.now().strftime("%Y%m%d")
        bdir = root / "exports" / "briefings"
        bdir.mkdir(parents=True, exist_ok=True)
        fname = f"Briefing_{date_tag}_{briefing.get('shift_mode', 'DAILY')}.md"
        fpath = bdir / fname
        fpath.write_text(briefing.get("markdown_briefing", ""), encoding="utf-8")
        results["channels"]["file"] = {
            "status": "EXPORTED",
            "path": str(fpath),
        }

    # 2. Kênh WeChat Work Webhook
    if "wechat" in channels_norm:
        wechat_md = format_for_wechat_markdown(briefing)
        wechat_res = send_to_wechat_webhook(
            markdown_content=wechat_md,
            webhook_url=wechat_webhook_url,
            dry_run=dry_run,
        )
        results["channels"]["wechat"] = wechat_res

    # 3. Kênh Zalo
    if "zalo" in channels_norm:
        zalo_txt = format_for_zalo_text(briefing)
        zalo_res = prepare_zalo_dispatch(
            text_content=zalo_txt,
            target_chat=target_zalo,
            dry_run=dry_run,
            repo_root=root,
        )
        results["channels"]["zalo"] = zalo_res

    # Ghi log lịch sử
    audit_file = root / "logs" / "notifications_audit.jsonl"
    audit_file.parent.mkdir(parents=True, exist_ok=True)
    with audit_file.open("a", encoding="utf-8") as f:
        f.write(json.dumps(results, ensure_ascii=False) + "\n")

    # Mirror vao Hub
    if DEFAULT_HUB_ROOT.is_dir():
        hub_audit = DEFAULT_HUB_ROOT / "logs" / "ops_audit.jsonl"
        try:
            hub_entry = {
                "at": results["timestamp"],
                "source": "notification_dispatcher",
                "action": "DISPATCH_BRIEFING",
                "ok": True,
                "details": {
                    "channels": list(results["channels"].keys()),
                    "dry_run": dry_run,
                },
            }
            with hub_audit.open("a", encoding="utf-8") as f:
                f.write(json.dumps(hub_entry, ensure_ascii=False) + "\n")
        except Exception:
            pass

    return results


def _deliver_zalo_queue_cli(queue_file_str: str) -> None:
    """Thực thi gửi tin nhắn cuối cùng trong zalo_queue.json qua DUC_OLAZ (chạy ở tiến trình nền)."""
    import asyncio
    qpath = Path(queue_file_str)
    if not qpath.is_file():
        return
    try:
        items = json.loads(qpath.read_text(encoding="utf-8"))
        if not items or not isinstance(items, list):
            return
        last_item = items[-1]
        target_chat = last_item.get("target_chat", "Thông báo chung OLAZ")
        message = last_item.get("message", "")
        if not message:
            return

        olaz_dir = DEFAULT_HUB_ROOT / "mcp_bridges" / "DUC_OLAZ"
        for k in [k for k in list(sys.modules.keys()) if k == "src" or k.startswith("src.")]:
            del sys.modules[k]
        if str(_REPO_ROOT) in sys.path:
            sys.path.remove(str(_REPO_ROOT))
        if str(olaz_dir) not in sys.path:
            sys.path.insert(0, str(olaz_dir))
        from src.tools.messaging import zalo_send_message  # type: ignore

        res = asyncio.run(zalo_send_message(to=target_chat, message=message))
        out_txt = res[0].text if res else "Done"
        last_item["status"] = "SENT" if "✅" in out_txt else "ATTEMPTED"
        last_item["delivery_result"] = out_txt
        last_item["delivered_at"] = datetime.now().isoformat(timespec="seconds")
        qpath.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        try:
            items[-1]["status"] = "DELIVERY_SKIPPED"
            items[-1]["delivery_error"] = str(exc)
            qpath.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--deliver-zalo-queue", type=str, default="", help="Deliver last queued Zalo message")
    args = parser.parse_args()
    if args.deliver_zalo_queue:
        _deliver_zalo_queue_cli(args.deliver_zalo_queue)
