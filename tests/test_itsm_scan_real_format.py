# -*- coding: utf-8 -*-
"""
test_itsm_scan_real_format.py — 10/11/2026 (STT 194)
Kiểm bộ quét Candidate Bxx theo ĐỊNH DẠNG THẬT của THEO-DOI-BU-TICKET-ITSM.md (khối `### Bxx` / `### Candidate Bxx`),
đối chiếu file tháng BuDon_ITSM_YYYY-MM.md, và bản tin KHÔNG được báo "ALL CLEAR" khi không đọc được sổ.
Chạy: cd mcp_bridges/DUC_CalendarSync_AIBridge && python -m pytest tests/test_itsm_scan_real_format.py -q
"""
import shutil
import sys
import uuid
from pathlib import Path
from typing import Iterator

import pytest

_BRIDGE_ROOT = Path(__file__).resolve().parents[1]
if str(_BRIDGE_ROOT) not in sys.path:
    sys.path.insert(0, str(_BRIDGE_ROOT))

from src.itsm_reminder_watcher import scan_pending_itsm_candidates, schedule_pending_itsm_reminder  # noqa: E402
from src.notification_dispatcher import format_for_zalo_text, format_for_wechat_markdown  # noqa: E402

REAL_LEDGER = Path(r"D:\DucWorker\DUC_SmartOps_AIHub\knowledge_core\01_SoCai_DanhBa\THEO-DOI-BU-TICKET-ITSM.md")

SAMPLE = """# 🎫 THEO DÕI BÙ TICKET ITSM

### B06. Cấp quyền xem màn 0405 (24/09/2026) — chưa tạo ticket
| # | Ngày | Mã | Xưởng | Mô tả | KQ | Loại | Trạng thái bù |
|---|---|---|---|---|---|---|---|
| 1 | 24/09 | X | 551 | quyền | OK | Quyền | ⏸️ **HOÃN BÙ 24/09** — customer không tra ra |

### B57. Khắc phục lỗi in tem C3-VN7-00608
| # | Ngày | Mã | Xưởng | Mô tả | KQ | Loại | Trạng thái bù |
|---|---|---|---|---|---|---|---|
| 1 | 20/10 | X | 631 | in tem | OK | In | 📋 **SẴN SÀNG BÙ (B57)** (Customer: 189176 范新鹏 Xinpeng Fan, P3) |

### B70. 🛑 KHÔNG BÙ THEO QUYẾT ĐỊNH CỦA DAN — SSPR
| # | Ngày | Mã | Xưởng | Mô tả | KQ | Loại | Trạng thái bù |
|---|---|---|---|---|---|---|---|
| 1 | 27/10 | X | 551 | SSPR | OK | Account | 🛑 **KHÔNG BÙ (B70)** |

### B76. [✅ ĐÃ BÙ INC000000194596] Xưởng 551 rỗng Nhân quân suất
| # | Ngày | Mã | Xưởng | Mô tả | KQ | Loại | Trạng thái bù |
|---|---|---|---|---|---|---|---|
| 1 | 29/10 | X | 551 | 人均率 | OK | Data | ✅ **ĐÃ BÙ (INC000000194596)** (Customer: V005933 Trúc Mai, P3) |

### Candidate B77 — Xưởng C7 & C8 nghẽn ESB (02/11/2026)
| # | Ngày | Mã | Xưởng | Mô tả | KQ | Loại | Trạng thái bù |
|---|---|---|---|---|---|---|---|
| 1 | 02/11 | X | C7 | ESB | OK | Tích hợp | 📋 **SẴN SÀNG BÙ (B77)** (Customer: V01790 Nguyễn Văn Phú - Quản lý C8, P3) |

### Candidate B88 — Xưởng C7-Mower trích lịch sử in phiếu 0702 (05/11/2026)
| # | Ngày | Mã | Xưởng | Mô tả | KQ | Loại | Trạng thái bù |
|---|---|---|---|---|---|---|---|
| 1 | 05/11 | X | 502 | 0702 | OK | Report | 📝 **CHỜ XÁC NHẬN (B88)** (Customer: `159299` 张龙 Long Zhang) |

### Candidate B89 — ✅ ĐÃ BÙ (INC000000194417) — VN7-00377 NiceLabel 10/5
| # | Ngày | Mã | Xưởng | Mô tả | KQ | Loại | Trạng thái bù |
|---|---|---|---|---|---|---|---|
| 1 | 06/11 | X | 501 | license | OK | In | 📋 **SẴN SÀNG BÙ (B89)** — Customer `V02099` Lucy Nguyen |

## 📊 TỔNG KẾT
| 1 | không phải candidate | SẴN SÀNG BÙ |
"""

BUDON = """# BÙ ĐƠN ITSM — Tháng 2026-10
| 45 | 28/10/2026 | INC000000192603 | 189176 | 范新鹏 | Lỗi in tem thùng BTP C3-VN7-00608 | In ấn | B57 |
- **Sẵn sàng bù:** 10 đơn (B37, B40, B41, B57, B59)
"""


@pytest.fixture
def tmp_path() -> Iterator[Path]:
    base = Path(__file__).resolve().parents[1] / ".pytest_tmp"
    base.mkdir(parents=True, exist_ok=True)
    work = base / f"case_{uuid.uuid4().hex[:8]}"
    work.mkdir(parents=True, exist_ok=True)
    try:
        yield work
    finally:
        shutil.rmtree(work, ignore_errors=True)


def _ledger(tmp_path: Path, text: str = SAMPLE, budon: str = BUDON) -> Path:
    p = tmp_path / "THEO-DOI-BU-TICKET-ITSM.md"
    p.write_text(text, encoding="utf-8")
    if budon:
        d = tmp_path / "BU-DON-ITSM-THEO-THANG"
        d.mkdir()
        (d / "BuDon_ITSM_2026-10.md").write_text(budon, encoding="utf-8")
    return p


def test_block_format_classification(tmp_path: Path):
    r = scan_pending_itsm_candidates(ledger_path=_ledger(tmp_path))
    assert r["scan_ok"] is True and r["format"] == "candidate_blocks"
    assert r["total_candidates"] == 7
    assert sorted(x["candidate_id"] for x in r["pending_candidates"]) == ["B77", "B88"]
    assert r["done_count"] == 3                      # B57 (qua BuDon), B76, B89
    assert r["no_comp_count"] == 1 and r["on_hold_count"] == 1
    assert sorted(x["candidate_id"] for x in r["stale_status"]) == ["B57", "B89"]
    b77 = next(x for x in r["pending_candidates"] if x["candidate_id"] == "B77")
    assert "V01790" in b77["customer"]


def test_budon_line_with_several_b_codes_is_not_proof(tmp_path: Path):
    r = scan_pending_itsm_candidates(ledger_path=_ledger(tmp_path, budon="- Sẵn sàng bù: B57, B77 — INC000000999999\n"))
    assert "B57" in [x["candidate_id"] for x in r["pending_candidates"]]


def test_unreadable_ledger_never_all_clear(tmp_path: Path):
    p = _ledger(tmp_path, text="# Sổ trống — không có khối Candidate\n", budon="")
    r = scan_pending_itsm_candidates(ledger_path=p)
    assert r["scan_ok"] is False and r["pending_count"] == 0
    rem = schedule_pending_itsm_reminder(ledger_path=p, targets=(), dry_run=True)
    assert rem["status"] == "SCAN_FAILED_NO_CONCLUSION" and rem["scheduled"] is False


def test_missing_ledger_file(tmp_path: Path):
    r = scan_pending_itsm_candidates(ledger_path=tmp_path / "khong_co.md")
    assert r["ledger_found"] is False and r["scan_ok"] is False


def _brief(**kw):
    b = {"shift_label": "🌅 BẢN TIN ĐẦU CA SÁNG (08:00 — MORNING BRIEFING)", "date": "2026-11-10", "today_events": [],
         "itsm_pending_count": 0, "itsm_pending_candidates": [], "action_checklist": []}
    b.update(kw)
    return b


def test_zalo_text_three_states():
    z_missing = format_for_zalo_text(_brief())                       # thieu itsm_scan_ok
    assert "ALL CLEAR" not in z_missing and "chưa kết luận" in z_missing
    z_zero = format_for_zalo_text(_brief(itsm_scan_ok=True))
    assert "ALL CLEAR" not in z_zero and "0 Candidate" in z_zero
    z_pend = format_for_zalo_text(_brief(itsm_scan_ok=True, itsm_pending_count=1, itsm_pending_candidates=[
        {"candidate_id": "B77", "customer": "V01790", "summary": "ESB"}],
        itsm_stale_status=[{"candidate_id": "B89"}]))
    assert "⚠️ 1 đơn chờ xử lý" in z_pend and "[B77]" in z_pend and "B89" in z_pend


def test_wechat_never_all_clear_without_scan():
    w = format_for_wechat_markdown(_brief())
    assert "ALL CLEAR" not in w and "chưa kết luận" in w


@pytest.mark.skipif(not REAL_LEDGER.is_file(), reason="Chỉ chạy trên máy có knowledge_core")
def test_real_ledger_smoke():
    r = scan_pending_itsm_candidates(ledger_path=REAL_LEDGER)
    ids = {x["candidate_id"] for x in r["pending_candidates"]}
    assert r["scan_ok"] is True and r["total_candidates"] >= 80
    assert {"B75", "B77", "B88"} <= ids and "B76" not in ids and "B89" not in ids
