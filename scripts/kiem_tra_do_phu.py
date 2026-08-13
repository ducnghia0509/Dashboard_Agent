# -*- coding: utf-8 -*-
"""KIỂM TRA ĐỘ PHỦ — cái chặn cho hệ thống qa khỏi mục ruỗng khi thêm nguồn mới.

Đẩy một báo cáo mới lên (vd VHKD An Taxi) thì chỉ mục nhãn dòng TỰ quét, agent trả lời được ngay ở
mức tra cứu. Nhưng nếu không ai khai loại câu hỏi / chỉ tiêu / câu golden cho nó thì nguồn đó cứ
nằm đó, agent trả lời nửa vời, và không ai biết cho tới khi lãnh đạo hỏi trúng.

Script này báo đúng bốn thứ đó. Chạy hằng đêm.

Chạy: .venv/bin/python scripts/kiem_tra_do_phu.py [--json]
Mã thoát: 0 = sạch, 1 = có mục cần xử lý.
"""
import argparse
import json
import os
import sys
import time

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.normpath(os.path.join(_HERE, ".."))
sys.path.insert(0, _ROOT)

from servers.common import doc_chi_tieu as dct      # noqa: E402
from servers.common import phan_loai as pl          # noqa: E402
from servers.common import row_index as ri          # noqa: E402
from servers.common import source_catalog as sc     # noqa: E402

NGAY_CHUA_INDEX_TOI_DA = 3


def _kpi_glossary() -> list:
    p = os.path.join(_ROOT, "kpi_glossary.json")
    if not os.path.exists(p):
        return []
    with open(p, encoding="utf-8") as fh:
        return json.load(fh)


def _golden() -> list:
    p = os.path.join(_ROOT, "eval", "qa_golden", "questions.json")
    if not os.path.exists(p):
        return []
    try:
        with open(p, encoding="utf-8") as fh:
            return json.load(fh)
    except json.JSONDecodeError:
        return []


def kiem_tra() -> dict:
    entries = sc.search()
    so_tay = pl.so_tay()

    # 1 · report_type có file nhưng KHÔNG loại câu hỏi nào nhận
    rt_da_khai = {r for lo in so_tay["loai"] for r in lo.get("report_type", [])}
    rt_co_file = {}
    for e in entries:
        rt = e.get("report_type")
        if rt:
            rt_co_file.setdefault(rt, {"so_file": 0, "so_sheet": 0})
            rt_co_file[rt]["so_file"] += 1
            rt_co_file[rt]["so_sheet"] += len(e.get("sheets") or [])
    chua_khai = [{"report_type": k, **v} for k, v in sorted(rt_co_file.items())
                 if k not in rt_da_khai]

    # 2 · chỉ tiêu glossary chưa có nguồn (dark KPI)
    dark = [{"chi_tieu": r.get("chi_tieu"), "man": r.get("man_hien_thi")}
            for r in _kpi_glossary() if r.get("co_tren_dashboard") == "chua"]

    # 3 · loại câu hỏi đã khai nhưng golden chưa có câu nào
    g = _golden()
    loai_trong_golden = {q.get("loai") for q in g if isinstance(q, dict)}
    thieu_golden = [lo["id"] for lo in so_tay["loai"] if lo["id"] not in loai_trong_golden]

    # 4 · file về đã lâu mà chưa vào chỉ mục
    da_index = {}
    if os.path.exists(ri.DB_PATH):
        con = ri._connect()
        try:
            da_index = {r["path"]: r["mtime"] for r in con.execute("SELECT path, mtime FROM quet")}
        finally:
            con.close()
    nguong = time.time() - NGAY_CHUA_INDEX_TOI_DA * 86400
    chua_index = [{"file": e["file"], "report_type": e.get("report_type")}
                  for e in entries
                  if e["path"] not in da_index and (e.get("mtime") or 0) < nguong]

    # 5 · file không suy được kỳ (không lọc theo kỳ nào thấy được)
    ky_khong_ro = sc.ky_khong_ro_list()

    # 6 · bố cục có file nhưng doc_chi_tieu chưa khai
    ct_da_khai = {c["id"] for c in dct.danh_sach_chi_tieu()}

    return {
        "report_type_chua_gan_loai_cau_hoi": chua_khai,
        "chi_tieu_chua_co_nguon": dark,
        "loai_cau_hoi_chua_co_golden": thieu_golden,
        "file_qua_han_chua_index": chua_index,
        "file_khong_ro_ky": ky_khong_ro,
        "tong_quan": {
            "so_file": len(entries),
            "so_report_type": len(rt_co_file),
            "so_loai_cau_hoi": len(so_tay["loai"]),
            "so_chi_tieu_chuan": len(ct_da_khai),
            "chi_muc": ri.san_sang(),
        },
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()
    kq = kiem_tra()
    if args.json:
        print(json.dumps(kq, ensure_ascii=False, indent=2))
        return 0

    tq = kq["tong_quan"]
    print(f"ĐỘ PHỦ AGENT QA — {tq['so_file']} file · {tq['so_report_type']} loại báo cáo · "
          f"{tq['so_loai_cau_hoi']} loại câu hỏi · {tq['so_chi_tieu_chuan']} chỉ tiêu chuẩn")
    print(f"Chỉ mục: {tq['chi_muc']}")

    can_xu_ly = 0
    muc = [
        ("Loại báo cáo CHƯA gắn loại câu hỏi nào", "report_type_chua_gan_loai_cau_hoi",
         lambda x: f"{x['report_type']}: {x['so_file']} file, {x['so_sheet']} sheet"),
        ("File quá hạn CHƯA vào chỉ mục", "file_qua_han_chua_index",
         lambda x: f"{x['file']} ({x['report_type']})"),
        ("File KHÔNG suy được kỳ (vắng mặt ở mọi phép lọc theo kỳ)", "file_khong_ro_ky",
         lambda x: f"{x['file']} ({x.get('report_type')})"),
        ("Loại câu hỏi CHƯA có câu golden", "loai_cau_hoi_chua_co_golden", lambda x: str(x)),
        ("Chỉ tiêu trong glossary CHƯA có nguồn", "chi_tieu_chua_co_nguon",
         lambda x: str(x["chi_tieu"])[:70]),
    ]
    for tieu_de, khoa, fmt in muc:
        ds = kq[khoa]
        if not ds:
            print(f"\n✓ {tieu_de}: không có")
            continue
        can_xu_ly += len(ds)
        print(f"\n⚠ {tieu_de}: {len(ds)}")
        for x in ds[:12]:
            print("   -", fmt(x))
        if len(ds) > 12:
            print(f"   … còn {len(ds) - 12} mục (dùng --json để xem hết)")

    print(f"\n{'=' * 64}\n{can_xu_ly} mục cần xử lý.")
    return 1 if can_xu_ly else 0


if __name__ == "__main__":
    sys.exit(main())
