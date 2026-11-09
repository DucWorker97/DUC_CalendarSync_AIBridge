# HANDOFF — CHUYỂN MÁY TÍNH & TIẾP QUẢN DỰ ÁN | DUC_CalendarSync_AIBridge

> **Ngày lập:** 09/11/2026  
> **Tác giả:** Nguyễn Duy Đức (`DucWorker97`) + Antigravity (AI Assistant)  
> **Mục đích:** Tài liệu bàn giao chuyển máy tính và tiếp quản phát triển dự án **DUC_CalendarSync_AIBridge v1.0.0** — MCP Server đồng bộ lịch biểu đa nền tảng (Google Calendar OAuth CLI + 1-Click Web URL, Microsoft Outlook Desktop MAPI COM, chuẩn quốc tế iCalendar RFC 5545 `.ics`) & tự động nhắc bù đơn Ticket ITSM (`Bxx`).

---

## 1. THÔNG TIN GIT & REPOSITORY

| Mục | Giá trị |
|---|---|
| **Git Root (máy đóng gói)** | `D:\repository portfolio\DUC_CalendarSync_AIBridge\DUC_CalendarSync_AIBridge` |
| **Đường dẫn mục tiêu trên máy mới** | `D:\DucWorker\DUC_SmartOps_AIHub\mcp_bridges\DUC_CalendarSync_AIBridge` (hoặc `D:\repository portfolio\DUC_CalendarSync_AIBridge`) |
| **Branch hiện tại** | `main` |
| **Commit bàn giao (HEAD)** | `chore(handoff): prepare repository transfer to new machine` (đồng bộ 100% với `origin/main`) |
| **Remote `origin`** | `https://github.com/DucWorker97/DUC_CalendarSync_AIBridge.git` |
| **Quyền truy cập Collaborator** | Đã cấp quyền **`write` (`push`)** cho tài khoản **`DucWorker19`** (Invitation ID: `342344645` — https://github.com/DucWorker97/DUC_CalendarSync_AIBridge/invitations) |
| **Trạng thái Working Tree** | Clean (`0` file modified/untracked; các tệp sinh động `exports/*.ics`, `.pytest_tmp/`, `__pycache__/` đã được ignore chuẩn xác) |

---

## 2. NỘI DUNG ĐÃ HOÀN THÀNH

### 2.1. Máy chủ MCP (`server.py`) — 5 Công cụ chuẩn Model Context Protocol (`stdio` JSON-RPC 2.0)
1. **`cal_status`**: Kiểm tra trạng thái kết nối Google Calendar (OAuth CLI / Web URL), Microsoft Outlook Desktop (MAPI COM), kho `.ics` và tổng hợp số lượng đơn `Bxx` chờ bù trong sổ cái ITSM.
2. **`cal_sync_all`**: Đồng bộ tức thì 3 chuỗi lịch vận hành tự động chuẩn (`DAILY_BACKUP_1745`, `WEEKLY_REPORT_SUN_1500`, `MONTHLY_ITSM_28_1600`) cùng lịch tùy chỉnh lên Google Calendar, Outlook Desktop và xuất tệp `.ics` (hỗ trợ `dry_run`).
3. **`cal_create_event`**: Đặt lịch sự kiện mới tức thì theo thời gian ISO 8601, hỗ trợ quy tắc lặp `NONE`, `DAILY_WORKDAYS`, `WEEKLY_SUN`, `MONTHLY_28`.
4. **`cal_remind_pending_itsm`**: Quét sổ cái `THEO-DOI-BU-TICKET-ITSM.md` (hoặc nhận trực tiếp Candidate `Bxx` mới phát sinh) để tự động ghim sự kiện nhắc việc **`[ITSM-Remind] 🎫 Nhắc Bù Đơn Ticket ITSM`** vào khung giờ hành chính tiếp theo (16:00 chiều hoặc 09:00 sáng ngày làm việc kế tiếp), đính kèm sẵn câu mẫu song ngữ VI–ZH hỏi Mã nhân viên (`V0xxxx`) theo Quy tắc 2 & Quy tắc 4.
5. **`cal_list_events`**: Liệt kê toàn bộ sự kiện đang quản lý kèm liên kết 1-Click Google Calendar, Outlook Web và đường dẫn tệp `.ics`.

### 2.2. Kiến trúc Core & Adapters (`src/`)
- **`src/models.py`**: Định nghĩa `SmartOpsCalendarEvent` (tự sinh `UID` chuẩn, thẻ chống trùng lặp `[SMARTOPS-ID:<event_id>]`) và 3 lịch biểu vận hành mặc định.
- **`src/calendar_engine.py`**: Bộ điều phối trung tâm hợp nhất 3 kênh (`ics`, `outlook`, `gcalendar`) và quản lý lịch tùy chỉnh (`exports/custom_events.json`).
- **`src/ics_exporter.py`**: Xuất tệp `.ics` chuẩn quốc tế RFC 5545 (múi giờ `Asia/Ho_Chi_Minh`, `RRULE`, `VALARM` nhắc trước) + tệp tổng hợp `DUC_SmartOps_Master_Schedule.ics` kèm mã băm MD5 và 1-Click URLs.
- **`src/outlook_adapter.py`**: Kết nối non-blocking tới Microsoft Outlook Desktop đang chạy qua `GetActiveObject("Outlook.Application")` (không treo khi Outlook đóng), tự động Upsert sự kiện kèm `RecurrencePattern` và ép `MeetingStatus = 0` (`olNonMeeting`) để không gửi email làm phiền người khác.
- **`src/gcalendar_adapter.py`**: Tự động phát hiện binary `gcalendar*.exe`, kiểm tra trạng thái OAuth, tìm sự kiện trùng lặp qua thẻ `[SMARTOPS-ID:...]` và đồng bộ trực tiếp hoặc fallback mượt mà về 1-Click Web URL + `.ics`.
- **`src/itsm_reminder_watcher.py`**: Bộ parser bảng Markdown quét sổ cái `THEO-DOI-BU-TICKET-ITSM.md`, phân loại thông minh các dòng đã có mã `INC\d{5,}` / `ĐÃ BÙ` với các dòng `PENDING` / `CHỜ MÃ NV` / `⏳`.

---

## 3. THAY ĐỔI CHƯA HOÀN THÀNH & HẠNG MỤC TIẾP THEO (BACKLOG)

1. **Đăng ký vào `mcp_config.json` trên máy chính (`D:\DucWorker\DUC_SmartOps_AIHub`):**
   - Sau khi clone/giải nén về `D:\DucWorker\DUC_SmartOps_AIHub\mcp_bridges\DUC_CalendarSync_AIBridge`, thêm cấu hình `DUC_CalendarSync_AIBridge` vào `mcp_config.json` (Antigravity / Claude / Cursor).
2. **Xác thực OAuth cho `gcalendar` CLI trên máy mới (tùy chọn):**
   - Nếu muốn đẩy sự kiện trực tiếp qua CLI không cần bấm 1-Click URL, đăng nhập tài khoản Google trong plugin `gcalendar` của Antigravity.
3. **Đồng bộ lần đầu với Outlook Desktop & Sổ cái ITSM thật trên máy nhà máy:**
   - Mở ứng dụng Microsoft Outlook Desktop trên máy mới và gọi `cal_sync_all(dry_run=False)` + `cal_remind_pending_itsm(dry_run=False)` để ghim 3 chuỗi lịch vận hành vào lịch Outlook thật và kiểm tra sổ cái `D:\DucWorker\DUC_SmartOps_AIHub\knowledge_core\01_SoCai_DanhBa\THEO-DOI-BU-TICKET-ITSM.md`.

---

## 4. KẾT QUẢ BUILD, TEST & KIỂM TRA READINESS (09/11/2026)

| Hạng mục kiểm tra | Câu lệnh thực thi | Kết quả | Chi tiết |
|---|---|---|---|
| **Python Syntax Compilation** | `python -m compileall server.py src tests` | **PASS (100%)** | Toàn bộ 9 file `.py` biên dịch sạch trên Python 3.12.10 |
| **Self-Test Mode** | `python server.py --selftest` | **PASS** | `mcp_sdk_ready=true`, `status_check=READY`, `sync_dry_run_events=3`, `itsm_reminder_dry_run=DRY_RUN`, đủ 5 tools |
| **Automated Pytest Suite** | `python -m pytest tests/ -v` | **PASS (5/5 tests — 1.11s)** | Kiểm thử đồng bộ 3 kênh, tạo sự kiện tùy chỉnh, quét sổ cái ITSM (`B88` đã bù / `B89` chờ mã NV), trạng thái All-Clear và `--selftest` |
| **MCP JSON-RPC 2.0 `stdio` Handshake** | Subprocess `initialize` + `tools/list` tới `python server.py` | **PASS (5/5 tools)** | Phản hồi chuẩn giao thức MCP `2024-11-05`, liệt kê chính xác 5 công cụ |
| **Secret & Privacy Scan** | Rà soát toàn bộ mã nguồn, `.gitignore`, Git index | **CLEAN (0 secrets)** | Không có API key, token, mật khẩu hay dữ liệu nhạy cảm nào trong Git |

---

## 5. DEPENDENCY & YÊU CẦU CẤU HÌNH MÁY MỚI

1. **Hệ điều hành:** Windows 10/11 64-bit (khuyến nghị để sử dụng tính năng đồng bộ trực tiếp Microsoft Outlook Desktop MAPI COM qua `pywin32`; chế độ `.ics` và Google Calendar hoạt động đa nền tảng).
2. **Python Runtime:** Python 3.10+ (đã kiểm nghiệm trên **Python 3.12.10**). Có thể dùng chung môi trường ảo `.venv` của `DUC_SmartOps_AIHub` hoặc tạo `.venv` riêng.
3. **Thư viện Python (`requirements.txt`):**
   - `mcp>=1.2.0` (tương thích cả `mcp 1.x` `FastMCP` và `mcp 2.x` `MCPServer`)
   - `pywin32>=306; sys_platform == 'win32'`
   - `pytest>=8.0.0`
4. **Phần mềm tích hợp (Tùy chọn):**
   - **Microsoft Outlook Desktop (Classic):** Mở sẵn khi muốn đồng bộ trực tiếp vào thư mục `Calendar` / `Lịch` qua COM.
   - **Google Calendar CLI (`gcalendar`):** Cài qua plugin Google Calendar của Antigravity nếu muốn đẩy lịch trực tiếp qua OAuth CLI.

---

## 6. SECRET & DỮ LIỆU PHẢI CHUYỂN RIÊNG (KHÔNG ĐƯA VÀO GIT / ZIP)

| Dữ liệu / Cấu hình | Vị trí | Lý do không đưa vào Git/ZIP | Cách xử lý trên máy mới |
|---|---|---|---|
| **Sổ cái ITSM thật (`THEO-DOI-BU-TICKET-ITSM.md`)** | `D:\DucWorker\DUC_SmartOps_AIHub\knowledge_core\01_SoCai_DanhBa\THEO-DOI-BU-TICKET-ITSM.md` | Thuộc kho tri thức cốt lõi của `DUC_SmartOps_AIHub` | Đã có sẵn trong repo `DUC_SmartOps_AIHub` trên máy mới (`locate_itsm_ledger()` tự động nhận diện) |
| **Token OAuth Google Calendar** | `%USERPROFILE%\.gemini\antigravity\plugin_data\gcalendar\` | Chứa token đăng nhập tài khoản Google cá nhân | Đăng nhập lại qua plugin Google Calendar trên máy mới |
| **Tệp lịch sinh động (`exports/*.ics`, `exports/custom_events.json`)** | `exports/` | Artifact sinh tự động theo thời điểm chạy | Tự động tái tạo khi chạy `python server.py --selftest` hoặc gọi tool `cal_sync_all` |
| **Môi trường ảo (`.venv/`) & Cache (`.pytest_tmp/`, `__pycache__/`)** | Thư mục gốc dự án | Phụ thuộc đường dẫn máy cục bộ | Khởi tạo lại trên máy mới |

---

## 7. BLOCKER & BƯỚC TIẾP THEO

- **Blocker hiện tại:** **Không có (0 Blocker).** Mã nguồn đã vượt qua 100% bài kiểm thử offline, self-test và giao thức MCP `stdio`.
- **Bước tiếp theo ngay khi sang máy mới:**
  1. Clone repository từ GitHub (hoặc giải nén gói ZIP bàn giao) vào `D:\DucWorker\DUC_SmartOps_AIHub\mcp_bridges\DUC_CalendarSync_AIBridge`.
  2. Chạy `python server.py --selftest` và `python -m pytest tests/ -q` để xác nhận môi trường máy mới.
  3. Khai báo MCP Server `DUC_CalendarSync_AIBridge` vào `mcp_config.json` của IDE.
  4. Mở Microsoft Outlook Desktop và gọi tool `cal_sync_all` để ghim 3 lịch vận hành định kỳ.

---

## 8. CÂU LỆNH TIẾP QUẢN TRÊN MÁY MỚI (QUICK START)

### Cách 1: Clone trực tiếp từ GitHub (Khuyến nghị)
```powershell
# 1. Di chuyển tới thư mục mcp_bridges của DUC_SmartOps_AIHub (hoặc D:\repository portfolio)
cd "D:\DucWorker\DUC_SmartOps_AIHub\mcp_bridges"

# 2. Clone repository từ GitHub
git clone https://github.com/DucWorker97/DUC_CalendarSync_AIBridge.git
cd DUC_CalendarSync_AIBridge

# 3. Kiểm tra trạng thái Git và cài đặt thư viện
git status
git log -n 3 --oneline --decorate
pip install -r requirements.txt

# 4. Chạy tự kiểm định (Self-Test) & bộ kiểm thử Pytest
python server.py --selftest
python -m pytest tests/ -v
```

### Cách 2: Giải nén từ gói ZIP bàn giao (Đã bao gồm đầy đủ `.git`)
```powershell
# 1. Giải nén gói ZIP vào thư mục đích
Expand-Archive -Path "DUC_CalendarSync_AIBridge_handoff_20261109.zip" -DestinationPath "D:\DucWorker\DUC_SmartOps_AIHub\mcp_bridges\"
cd "D:\DucWorker\DUC_SmartOps_AIHub\mcp_bridges\DUC_CalendarSync_AIBridge"

# 2. Kiểm tra tính toàn vẹn Git & đồng bộ với origin
git status
git remote -v
git fetch origin
git status -sb

# 3. Cài đặt thư viện & chạy kiểm thử nghiệm thu
pip install -r requirements.txt
python server.py --selftest
python -m pytest tests/ -v
```
