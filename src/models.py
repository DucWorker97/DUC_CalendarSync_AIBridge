# -*- coding: utf-8 -*-
"""
models.py — Mô hình Dữ liệu Lịch biểu Chuẩn cho DUC_CalendarSync_AIBridge
=========================================================================
Định nghĩa cấu trúc sự kiện lịch thông minh (SmartOpsCalendarEvent) và
3 lịch biểu vận hành tự động mặc định:
1. 17:45 Hàng ngày (Thứ 2 - Thứ 7): Sao lưu Đám mây Tự động (Hot Snapshot + Cold H.265) & Chưng cất Inbox.
2. 15:00 Chủ Nhật hàng tuần: Tổng hợp & Kiểm toán 9 Lớp Báo cáo Tuần MES Song ngữ (VI-ZH).
3. 16:00 Ngày 28 hàng tháng: Rà soát Sổ cái Sự cố & Khóa sổ Bù đơn Ticket ITSM Tháng.
=========================================================================
"""

from dataclasses import dataclass, asdict
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional


@dataclass
class SmartOpsCalendarEvent:
    """Mô hình chuẩn của một sự kiện lịch biểu đa nền tảng (Google, Outlook, iCalendar .ics)."""
    event_id: str
    title: str
    description: str
    location: str
    start_dt: datetime
    duration_minutes: int = 30
    rrule: Optional[str] = None
    reminder_minutes: int = 15
    category: str = "MES-SmartOps"
    recurrence_type: str = "NONE"  # NONE | DAILY_WORKDAYS | WEEKLY_SUN | MONTHLY_28

    @property
    def end_dt(self) -> datetime:
        return self.start_dt + timedelta(minutes=self.duration_minutes)

    @property
    def uid(self) -> str:
        return f"{self.event_id}@ducworker97.smartops.local"

    @property
    def sync_tag(self) -> str:
        return f"[SMARTOPS-ID:{self.event_id}]"

    def to_dict(self) -> Dict[str, Any]:
        data = asdict(self)
        data["start_dt"] = self.start_dt.isoformat(timespec="seconds")
        data["end_dt"] = self.end_dt.isoformat(timespec="seconds")
        data["uid"] = self.uid
        data["sync_tag"] = self.sync_tag
        return data


def _next_weekday_time(now: datetime, target_hour: int, target_minute: int) -> datetime:
    candidate = now.replace(hour=target_hour, minute=target_minute, second=0, microsecond=0)
    if candidate <= now:
        candidate += timedelta(days=1)
    while candidate.weekday() == 6:  # Chủ Nhật -> nhảy sang Thứ 2
        candidate += timedelta(days=1)
    return candidate


def _next_sunday_time(now: datetime, target_hour: int, target_minute: int) -> datetime:
    candidate = now.replace(hour=target_hour, minute=target_minute, second=0, microsecond=0)
    days_ahead = (6 - candidate.weekday()) % 7
    candidate += timedelta(days=days_ahead)
    if candidate <= now:
        candidate += timedelta(days=7)
    return candidate


def _next_day_of_month(now: datetime, day: int, target_hour: int, target_minute: int) -> datetime:
    year, month = now.year, now.month
    try:
        candidate = now.replace(day=day, hour=target_hour, minute=target_minute, second=0, microsecond=0)
    except ValueError:
        candidate = now.replace(day=28, hour=target_hour, minute=target_minute, second=0, microsecond=0)
    if candidate <= now:
        if month == 12:
            year += 1
            month = 1
        else:
            month += 1
        candidate = candidate.replace(year=year, month=month, day=day)
    return candidate


def get_default_smartops_events(reference_dt: Optional[datetime] = None) -> List[SmartOpsCalendarEvent]:
    """Sinh danh sách 3 lịch biểu vận hành tự động chuẩn của hệ sinh thái MES/IT."""
    now = reference_dt or datetime.now()

    ev_daily_backup = SmartOpsCalendarEvent(
        event_id="DAILY_BACKUP_1745",
        title="[MES-Ops] ☁️ Sao lưu Đám mây Tự động 17:45 & Chưng cất Hòm thư",
        description=(
            "Lịch trình tự động cuối ca làm việc hàng ngày (17:45):\n"
            "1. Dịch vụ nền inbox_watcher tự động kích hoạt sao lưu Kho Nóng (Snapshot ZIP + MD5) và Kho Lạnh (Video H.265).\n"
            "2. Kiểm tra trạng thái HEALTHY trên Web UI Dashboard (http://127.0.0.1:8000 - Tab 5).\n"
            "3. Đảm bảo tuân thủ Quy tắc 12 (Bảo toàn 100% dữ liệu)."
        ),
        location="http://127.0.0.1:8000 (#tab-backup)",
        start_dt=_next_weekday_time(now, 17, 45),
        duration_minutes=15,
        rrule="FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR,SA",
        reminder_minutes=10,
        category="MES-SmartOps-Backup",
        recurrence_type="DAILY_WORKDAYS",
    )

    ev_weekly_report = SmartOpsCalendarEvent(
        event_id="WEEKLY_REPORT_SUN_1500",
        title="[MES-Ops] 📊 Lập & Kiểm toán 9 Lớp Báo cáo Tuần MES Song ngữ (VI-ZH)",
        description=(
            "Quy trình định kỳ chiều Chủ Nhật hàng tuần (15:00):\n"
            "1. Quét nhật ký phiên và sự cố MES trong tuần.\n"
            "2. Biên soạn Báo cáo tuần bản Tiếng Việt (.docx) và Tiếng Trung (.docx + .pdf).\n"
            "3. Chạy kiểm toán tự động 9 Lớp (Gate 6 + Layer 9 Quy tắc 13: Chặn rò rỉ tên dự án nội bộ, Quy tắc 11: Tiêu đề nằm trên bảng, Font Microsoft YaHei).\n"
            "4. Tạo thư nháp Outlook chờ duyệt 4D trước khi gửi."
        ),
        location="knowledge_core/05_LichSu_QuyTac/BaoCao_BoPhanIT/BaoCao-Tuan",
        start_dt=_next_sunday_time(now, 15, 0),
        duration_minutes=45,
        rrule="FREQ=WEEKLY;BYDAY=SU",
        reminder_minutes=30,
        category="MES-SmartOps-WeeklyReport",
        recurrence_type="WEEKLY_SUN",
    )

    ev_monthly_itsm = SmartOpsCalendarEvent(
        event_id="MONTHLY_ITSM_28_1600",
        title="[MES-Ops] 🎫 Đối soát Sổ cái Sự cố & Hoàn tất Bù đơn Ticket ITSM Tháng",
        description=(
            "Quy trình khóa sổ cuối tháng (16:00 ngày 28 hàng tháng):\n"
            "1. Rà soát toàn bộ sự cố có thông báo/tag tên xử lý trong tháng (Quy tắc 1 & Quy tắc 2).\n"
            "2. Thẩm tra Customer qua Danh bạ đối chiếu (Quy tắc 5, 6, 7).\n"
            "3. Hoàn tất tạo và đóng (Resolved/Closed) 100% Candidate Bxx trên BMC Smart IT theo chuẩn Tam ngữ VI-ZH-EN <= 200 ký tự kèm dự đoán Root Cause (Quy tắc 8 & 9)."
        ),
        location="knowledge_core/01_SoCai_DanhBa/THEO-DOI-BU-TICKET-ITSM.md",
        start_dt=_next_day_of_month(now, 28, 16, 0),
        duration_minutes=60,
        rrule="FREQ=MONTHLY;BYMONTHDAY=28",
        reminder_minutes=60,
        category="MES-SmartOps-ITSM",
        recurrence_type="MONTHLY_28",
    )

    return [ev_daily_backup, ev_weekly_report, ev_monthly_itsm]
