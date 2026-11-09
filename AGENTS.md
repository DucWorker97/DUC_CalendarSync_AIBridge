# Quy tắc làm việc & Kiến trúc Dự án `DUC_CalendarSync_AIBridge` (Workspace Rules)

Tập tin này định nghĩa bối cảnh, kiến trúc, quy tắc phát triển và lịch sử phiên làm việc cho trợ lý AI khi làm việc trực tiếp trong dự án **`DUC_CalendarSync_AIBridge`** (Tác giả: Nguyễn Duy Đức — GitHub: `DucWorker97`).

---

## 1. Tổng Quan Dự Án & Kiến Trúc Cốt Lõi

- **Tên dự án:** `DUC_CalendarSync_AIBridge` (`v1.4.0`)
- **Tài khoản GitHub chính thức:** `DucWorker97` (`https://github.com/DucWorker97/DUC_CalendarSync_AIBridge.git`)
- **Mục tiêu:** MCP Server (`stdio` JSON-RPC 2.0) chuyên trách **Đồng bộ Lịch biểu Tự động Đa Nền tảng 2 Chiều**, **Tự động Nhắc Lịch Sự cố Chờ Bù Đơn ITSM (`Bxx`)**, **Sinh Bản tin Đầu ca / Cuối ca** và **Bắn Thông Báo Tức Thì Đa Kênh (WeChat Work & Zalo)**.
- **3 Nền tảng Lịch tích hợp:**
  1. **Google Calendar (`src/gcalendar_adapter.py`):** Đẩy/cập nhật/xóa/truy vấn trực tiếp qua `gcalendar` OAuth CLI (`mutate recurring` / `mutate create` / `mutate delete` / `readonly list` với `--send-updates none`, có kiểm tra chống trùng lặp Idempotent `_find_existing_gcal_event`) + sinh liên kết 1-Click Google Calendar Web URL kèm `RRULE`.
  2. **Microsoft Outlook Desktop (`src/outlook_adapter.py`):** Đồng bộ/cập nhật/xóa/truy vấn trực tiếp qua MAPI COM (`win32com.client`, thư mục `Lịch` `olFolderCalendar = 9`, `MeetingStatus = 0` không gửi email làm phiền người ngoài, chống trùng lặp theo thẻ `[SMARTOPS-ID:...]`) + liên kết 1-Click Outlook Web Deep-Link.
  3. **Chuẩn Quốc tế iCalendar RFC 5545 (`src/ics_exporter.py`):** Xuất tệp `.ics` đơn lẻ và `DUC_SmartOps_Master_Schedule.ics` kèm múi giờ `Asia/Ho_Chi_Minh`, `RRULE`, `VALARM` và chữ ký `MD5`.
- **Kênh Thông Báo Tức Thì Đa Nền Tảng (`src/notification_dispatcher.py`):**
  1. **WeChat Work Group Bot Webhook (企业微信):** Xuất bản tin Markdown chuyên dụng chuẩn thẻ màu `<font color="info|warning|comment">`, giới hạn an toàn 4096 bytes, hỗ trợ HTTP POST & Dry-Run.
  2. **Zalo Web Automation:** Chuyển đổi Markdown sang Rich-Text phù hợp Zalo, định dạng danh sách checklist, hàng rào phân cách, hỗ trợ hàng đợi tự động `exports/briefings/zalo_queue.json`.

---

## 2. Cấu Trúc Thư Mục Hiện Tại

```text
DUC_CalendarSync_AIBridge/
├── server.py                        # Entrypoint MCP Server (11 tools + --selftest + --register-mcp + --register-tasks)
├── src/
│   ├── __init__.py
│   ├── models.py                    # SmartOpsCalendarEvent + 3 chuỗi lịch vận hành chuẩn
│   ├── calendar_engine.py           # Bộ điều phối trung tâm (10 nghiệp vụ cốt lõi)
│   ├── notification_dispatcher.py   # Bộ điều phối bắn bản tin tức thì đa kênh (WeChat & Zalo)
│   ├── scheduler_runner.py          # Trình chạy tự động hóa ca sáng/chiều & tích hợp bắn tin
│   ├── gcalendar_adapter.py         # Cầu nối Google Calendar OAuth CLI + 1-Click Web URL
│   ├── outlook_adapter.py           # Cầu nối Microsoft Outlook Desktop MAPI COM + Web URL
│   ├── ics_exporter.py              # Bộ xuất chuẩn RFC 5545 (.ics) + MD5
│   └── itsm_reminder_watcher.py     # Quét sổ cái THEO-DOI-BU-TICKET-ITSM.md & tự đặt lịch nhắc Bxx
├── scripts/                         # Bộ kịch bản 1-Click & PowerShell quản trị Task Scheduler
├── tests/
│   └── test_calendar_mcp.py         # Bộ kiểm thử Pytest (10/10 PASS)
├── exports/                         # Thư mục lưu tệp .ics, custom_events.json & briefings/
├── HANDOFF-MACHINE-TRANSFER.md      # Tài liệu bàn giao chuyển máy & tiếp quản
├── pytest.ini                       # Cấu hình Pytest (pythonpath = .)
├── requirements.txt
├── README.md
└── Push_To_GitHub_DucWorker97.bat   # Kịch bản đẩy GitHub tài khoản DucWorker97
```

---

## 3. Danh Mục 11 MCP Tools Đã Triển Khai (`v1.4.0`)

1. `cal_status`: Kiểm tra kết nối Google Calendar OAuth, Microsoft Outlook Desktop MAPI, kho `.ics` và số lượng đơn `Bxx` chờ bù.
2. `cal_sync_all`: Đồng bộ tức thì 3 chuỗi lịch vận hành chuẩn (`DAILY_BACKUP_1745`, `WEEKLY_REPORT_SUN_1500`, `MONTHLY_ITSM_28_1600`) và lịch tùy chỉnh lên cả 3 nền tảng.
3. `cal_create_event`: Đặt lịch sự kiện mới tức thì (hỗ trợ lặp `NONE`, `DAILY_WORKDAYS`, `WEEKLY_SUN`, `MONTHLY_28`).
4. `cal_remind_pending_itsm`: Quét sổ cái `THEO-DOI-BU-TICKET-ITSM.md` (hoặc nhận trực tiếp Candidate `Bxx`) để tự động ghim lịch nhắc bù đơn + mẫu câu hỏi Mã nhân viên (`V0xxxx` — Quy tắc 4).
5. `cal_list_events`: Liệt kê toàn bộ lịch trình đang quản lý kèm liên kết 1-Click.
6. `cal_delete_event`: Xóa sự kiện tùy chỉnh theo `event_id` trên kho `.ics`, Microsoft Outlook Desktop & Google Calendar (bảo vệ 3 chuỗi lịch hệ thống).
7. `cal_query_schedule`: Đọc lịch trình thực tế 2 chiều trong ngày/tuần hợp nhất từ Managed Store, Google Calendar & Outlook Desktop.
8. `cal_check_conflicts`: Kiểm tra xung đột lịch họp (Free-Busy Finder) và tự động gợi ý 3 khung giờ trống gần nhất trong giờ hành chính.
9. `cal_update_event`: Dời lịch (reschedule) hoặc cập nhật tiêu đề, thời lượng, mô tả, địa điểm của sự kiện tùy chỉnh theo `event_id`.
10. `cal_daily_briefing`: Sinh Bản tin Lịch trình Đầu ca (`MORNING` - 08:00) / Cuối ca (`EVENING` - 17:15) kèm trạng thái ITSM & Checklist hành động.
11. `cal_dispatch_notification`: Định dạng và bắn Bản tin Lịch trình tức thì qua các kênh cộng tác (WeChat Work Webhook, Zalo Web Automation) kèm chế độ Dry-Run an toàn.

---

## 4. Nguyên Tắc Bắt Buộc, An Toàn & Bảo Mật Khi Phát Triển

1. **Đọc trước khi code:** Đọc đầy đủ `HANDOFF-MACHINE-TRANSFER.md` → `README.md` → `server.py` → `src/calendar_engine.py` trước khi chỉnh sửa mã nguồn.
2. **Tuân thủ quy tắc hệ sinh thái `DUC_SmartOps_AIHub`:**
   - **Quy tắc 2 & Quy tắc 4 (ITSM Smart Reminder):** Mọi sự cố chờ bù đơn (`Bxx`) trong `THEO-DOI-BU-TICKET-ITSM.md` chưa có mã `INC...` hoặc thiếu Mã nhân viên (`V0xxxx` / 6 số) khi tạo lịch nhắc qua `cal_remind_pending_itsm` bắt buộc đính kèm mẫu câu hỏi Mã NV song ngữ VI–ZH.
   - **Quy tắc 11 (Báo cáo tuần MES):** Lịch nhắc `WEEKLY_REPORT_SUN_1500` (15:00 Chủ Nhật) luôn nhắc kiểm toán 9 lớp, tiêu đề nằm trên bảng metadata, font `Microsoft YaHei`, xuất PDF qua Word COM và tạo thư nháp Outlook chờ duyệt 4D.
   - **Quy tắc 12 (Bảo toàn dữ liệu 100%):** Lịch nhắc `DAILY_BACKUP_1745` (17:45 Thứ 2 – Thứ 7) theo dõi sao lưu Kho Nóng (Snapshot ZIP + MD5) và Kho Lạnh (Video H.265).
3. **Không gửi thư mời ra ngoài (Zero External Spam & Non-blocking):**
   - Trên Google Calendar CLI: Luôn truyền cờ `--send-updates none`.
   - Trên Microsoft Outlook COM: Ưu tiên `GetActiveObject("Outlook.Application")` có fallback sang `Dispatch`, luôn đặt `appt.MeetingStatus = 0` (`olNonMeeting`) và không thêm `Recipients`.
4. **Chống trùng lặp tuyệt đối (Idempotency):**
   - Mọi sự kiện tạo ra đều gắn định danh `[SMARTOPS-ID:<event_id>]`. Trước khi tạo mới phải kiểm tra sự kiện đã tồn tại trên Google Calendar / Outlook hay chưa.
5. **Bảo mật, Git Hygiene & Định danh Git Author:**
   - Tuyệt đối không đưa `.env`, `*token*.json`, `*credentials*.json`, `auth.json`, hoặc tệp lịch sinh động (`exports/*.ics`, `exports/custom_events.json`) vào Git.
   - Luôn duy trì `git config user.name "DucWorker97"` trong kho lưu trữ này.
   - Chạy đủ 3 bước kiểm tra trước khi commit: `python -m compileall server.py src tests`, `python server.py --selftest`, và `python -m pytest tests/ -q`.

---

## 5. Lịch Sử Phiên Làm Việc

| Ngày | Agent | Việc đã làm |
|---|---|---|
| 09/11/2026 | Antigravity (Session 1) | Khởi tạo kiến trúc **DUC_CalendarSync_AIBridge v1.0.0** gồm 5 MCP Tools (`cal_status`, `cal_sync_all`, `cal_create_event`, `cal_remind_pending_itsm`, `cal_list_events`), 3 chuỗi lịch vận hành tự động (`DAILY_BACKUP_1745`, `WEEKLY_REPORT_SUN_1500`, `MONTHLY_ITSM_28_1600`), bộ xuất iCalendar RFC 5545 (`src/ics_exporter.py`), adapter Google Calendar OAuth CLI (`src/gcalendar_adapter.py`), adapter Outlook MAPI COM (`src/outlook_adapter.py`), bộ quét nhắc bù đơn ITSM (`src/itsm_reminder_watcher.py`) và bộ kiểm thử tự động (`tests/test_calendar_mcp.py`). |
| 09/11/2026 | Antigravity (Session 2 — Handoff) | Chuẩn bị bàn giao chuyển máy: Nâng cấp `src/outlook_adapter.py` sang `GetActiveObject("Outlook.Application")`; tối ưu `locate_gcalendar_cli()` trong `src/gcalendar_adapter.py`; bổ sung tương thích kép MCP SDK v1/v2 (`MCPServer` / `FastMCP`) trong `server.py`; thêm `pytest.ini` & fixture `tmp_path` cục bộ tương thích Python 3.12 trên Windows; tạo `AGENTS.md`, `HANDOFF-MACHINE-TRANSFER.md`, `exports/.gitkeep`; đồng bộ Git lên GitHub `DucWorker97/DUC_CalendarSync_AIBridge`. |
| 09/11/2026 | Antigravity (Session 2 — Collaborator) | Cấp quyền cộng tác viên (`Collaborator - write`, Invitation ID `342344645`) cho tài khoản **`DucWorker19`** trên repository GitHub `DucWorker97/DUC_CalendarSync_AIBridge`. |
| 09/11/2026 | Antigravity (Session 3 — Sync & v1.2.0) | Tiếp quản trên máy chính (`D:\DucWorker\DUC_SmartOps_AIHub\mcp_bridges\DUC_CalendarSync_AIBridge`): Hợp nhất cấu hình `pytest.ini`, tương thích kép `MCPServer`/`FastMCP`, fallback `gcalendar` CLI & Outlook COM; đăng ký `DUC_CalendarSync_AIBridge` vào `mcp_config.json`; đồng bộ thật `cal_sync_all(dry_run=False)` lên Google Calendar & Outlook Desktop; nâng cấp lên **v1.2.0** với 10 MCP Tools (`+cal_delete_event`, `+cal_query_schedule`, `+cal_check_conflicts`, `+cal_update_event`, `+cal_daily_briefing`, `+--register-mcp`) và đạt `7/7` Pytest PASS. |
| 09/11/2026 | Antigravity (Session 4 — Portfolio & CI) | Chuẩn hóa Portfolio GitHub `DucWorker97`: Thiết lập GitHub Actions CI Matrix đa nền tảng (`.github/workflows/ci.yml` — Windows & Ubuntu, Python 3.10, 3.11, 3.12); bổ sung giấy phép mã nguồn mở `LICENSE` (MIT); tái cấu trúc và nâng cấp `README.md` chuẩn Showcase với bộ huy hiệu động, sơ đồ kiến trúc Mermaid chi tiết, bảng đặc tả 10 MCP Tools và tài liệu tích hợp đa Client. |
| 09/11/2026 | Antigravity (Session 5 — Hub Interconnection) | Liên thông toàn diện **DUC_CalendarSync_AIBridge v1.2.0** với hệ sinh thái `DUC_SmartOps_AIHub`: Đăng ký MCP Server vào `mcp_config_hub.json` (13 servers); nâng cấp `app_engine/calendar_sync/calendar_service.py` ủy quyền gọi trực tiếp `calendar_engine` v1.2.0; mở rộng 6 REST API endpoints mới trong `app_engine/web_ui/routes_api.py`; tích hợp Tab 6 `📅 Đồng Bộ Lịch & ITSM` trên Control Center Web UI (`templates/index.html`); mở rộng CLI `--briefing` và `--remind-itsm`; xác minh 18/18 Web UI tests PASS và 7/7 Bridge tests PASS. |
| 09/11/2026 | Antigravity (Session 6 — Scheduled Automations & v1.3.0) | Triển khai Hướng 4 — Lịch Quét & Vận Hành Tự Động Hàng Ngày: Xây dựng module `src/scheduler_runner.py` (v1.3.0) xử lý ca sáng (08:00) và ca chiều (17:15) kèm xuất Markdown và ghi nhật ký audit; thiết lập kịch bản tự động hóa cấp OS qua Windows Task Scheduler (`schtasks` với `pythonw.exe` hoàn toàn không hiện popup CMD); bổ sung bộ scripts quản trị `scripts/` (PowerShell & 1-Click Batches); tạo 2 Antigravity Scheduled Automations Sidecars (`smartops-morning-briefing` và `smartops-evening-briefing`); mở rộng Web UI Tab 6 quản trị lịch trình Windows Task Scheduler; nâng tổng số bài kiểm thử lên 9/9 Bridge tests PASS (100%) và 19/19 Hub Web UI tests PASS (100%). |
| 09/11/2026 | Antigravity (Session 7 — Multi-Channel Dispatch & v1.4.0) | Triển khai Hướng 1 — Tích Hợp Thông Báo Tức Thì Đa Kênh (WeChat Work & Zalo Webhook): Xây dựng module `src/notification_dispatcher.py` (v1.4.0) hỗ trợ chuẩn hóa Markdown màu sắc WeChat Work (`<font color="...">`), Rich-Text Zalo kèm hàng rào phân cách, kiểm soát dung lượng thẻ $\le 4000$ bytes; nâng cấp `src/scheduler_runner.py` cờ `--dispatch`, `--channels`, `--confirm-send`; bổ sung MCP Tool thứ 11 `cal_dispatch_notification` trong `server.py`; liên thông DUC_SmartOps_AIHub với 2 API endpoints `/api/calendar/notifications/config` và `/api/calendar/notifications/dispatch`; thêm giao diện thẻ điều khiển bắn tin tức thì đa kênh trên Web UI Tab 6; xác minh 10/10 Bridge tests PASS (100%) và 20/20 Hub Web UI tests PASS (100%). |


