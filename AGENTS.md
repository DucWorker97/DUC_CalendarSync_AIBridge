# AGENTS.md — DUC_CalendarSync_AIBridge

> Quy tắc làm việc & Theo dõi lịch sử AI Agent cho dự án **DUC_CalendarSync_AIBridge** (Multi-Platform Calendar Sync & ITSM Smart Reminder MCP Server)

---

## 1. Quy tắc bắt buộc cho AI Agent tiếp theo

1. **ĐỌC TRƯỚC KHI CODE:** Đọc đầy đủ `HANDOFF-MACHINE-TRANSFER.md` → `README.md` → `server.py` → `src/calendar_engine.py` trước khi chỉnh sửa mã nguồn.
2. **TUÂN THỦ QUY TẮC HỆ SINH THÁI `DUC_SmartOps_AIHub`:**
   - **Quy tắc 2 & Quy tắc 4 (ITSM Smart Reminder):** Mọi sự cố chờ bù đơn (`Bxx`) trong `THEO-DOI-BU-TICKET-ITSM.md` chưa có mã `INC...` hoặc thiếu Mã nhân viên (`V0xxxx` / 6 số) khi tạo lịch nhắc qua `cal_remind_pending_itsm` bắt buộc đính kèm mẫu câu hỏi Mã NV song ngữ VI–ZH.
   - **Quy tắc 11 (Báo cáo tuần MES):** Lịch nhắc `WEEKLY_REPORT_SUN_1500` (15:00 Chủ Nhật) luôn nhắc kiểm toán 9 lớp, tiêu đề nằm trên bảng metadata, font `Microsoft YaHei`, xuất PDF qua Word COM và tạo thư nháp Outlook chờ duyệt 4D.
   - **Quy tắc 12 (Bảo toàn dữ liệu 100%):** Lịch nhắc `DAILY_BACKUP_1745` (17:45 Thứ 2 – Thứ 7) theo dõi sao lưu Kho Nóng (Snapshot ZIP + MD5) và Kho Lạnh (Video H.265).
3. **AN TOÀN KẾT NỐI ĐA NỀN TẢNG (NON-BLOCKING & PRIVACY):**
   - **Outlook Desktop COM (`src/outlook_adapter.py`):** Luôn dùng `win32com.client.GetActiveObject("Outlook.Application")` (non-blocking) thay vì `Dispatch("Outlook.Application")` mù quáng để tránh treo tiến trình nền (`OUTLOOK.EXE -Embedding`) khi Outlook chưa mở. Luôn đặt `MeetingStatus = 0` (`olNonMeeting`) để chỉ tạo lịch nhắc cá nhân, tuyệt đối không tự gửi email mời họp ra ngoài.
   - **Google Calendar CLI (`src/gcalendar_adapter.py`):** Chỉ gọi file thực thi nhị phân (`gcalendar*.exe`) đã cài đặt, chống trùng lặp Idempotent qua thẻ `[SMARTOPS-ID:<event_id>]` và luôn trả kèm 1-Click Web URL + tệp `.ics` chuẩn RFC 5545.
4. **BẢO MẬT & GIT HYGIENE:**
   - Tuyệt đối không đưa `.env`, `*token*.json`, `*credentials*.json`, `auth.json`, hoặc tệp lịch sinh động (`exports/*.ics`, `exports/custom_events.json`) vào Git.
5. **KIỂM ĐỊNH TRƯỚC KHI COMMIT:**
   - Chạy đủ 3 bước kiểm tra: `python -m compileall server.py src tests`, `python server.py --selftest`, và `python -m pytest tests/ -q`.

---

## 2. Lịch sử Phiên làm việc

| Ngày | Agent | Việc đã làm |
|---|---|---|
| 09/11/2026 | Antigravity (Session 1) | Khởi tạo kiến trúc **DUC_CalendarSync_AIBridge v1.0.0** gồm 5 MCP Tools (`cal_status`, `cal_sync_all`, `cal_create_event`, `cal_remind_pending_itsm`, `cal_list_events`), 3 chuỗi lịch vận hành tự động (`DAILY_BACKUP_1745`, `WEEKLY_REPORT_SUN_1500`, `MONTHLY_ITSM_28_1600`), bộ xuất iCalendar RFC 5545 (`src/ics_exporter.py`), adapter Google Calendar OAuth CLI (`src/gcalendar_adapter.py`), adapter Outlook MAPI COM (`src/outlook_adapter.py`), bộ quét nhắc bù đơn ITSM (`src/itsm_reminder_watcher.py`) và bộ kiểm thử tự động (`tests/test_calendar_mcp.py`). |
| 09/11/2026 | Antigravity (Session 2 — Handoff) | Chuẩn bị bàn giao chuyển máy: Nâng cấp `src/outlook_adapter.py` sang `GetActiveObject("Outlook.Application")` non-blocking (triệt tiêu lỗi treo `OUTLOOK.EXE` nền khi chưa mở Outlook); tối ưu `locate_gcalendar_cli()` trong `src/gcalendar_adapter.py` chỉ gọi binary `.exe` thực tế; bổ sung tương thích kép MCP SDK v1/v2 (`MCPServer` / `FastMCP`) trong `server.py`; thêm `pytest.ini` & fixture `tmp_path` cục bộ tương thích Python 3.12 trên Windows; xác minh `compileall`, `--selftest`, `pytest` (5/5 PASS) và JSON-RPC 2.0 `stdio` handshake (5/5 tools); tạo `AGENTS.md`, `HANDOFF-MACHINE-TRANSFER.md`, `exports/.gitkeep`; đồng bộ Git lên GitHub `DucWorker97/DUC_CalendarSync_AIBridge` và đóng gói ZIP bàn giao kèm SHA-256. |
