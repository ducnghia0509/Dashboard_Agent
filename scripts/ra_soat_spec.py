#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""RÀ SOÁT KHÔ 95 SPEC — spec nào đang đọc hụt, đọc rỗng, hoặc kêu cảnh báo.

VÌ SAO CẦN
----------
Tầng trích xuất đã tự tính sẵn hai con số nói lên "engine bỏ sót gì" mà chưa ai gom lại nhìn
một lượt:

  · `dong`                        — số bản ghi engine LẤY ĐƯỢC từ file;
  · `bỏ N dòng không qua bộ lọc`  — số dòng engine ĐỌC THẤY nhưng loại đi.

Tỷ lệ bỏ cao bất thường, hoặc `dong = 0`, là dấu hiệu spec đang đọc hụt file — kiểu lỗi IM
LẶNG nhất trong hệ: file vẫn nạp, cron vẫn "THÀNH CÔNG", dashboard vẫn có số, chỉ là thiếu.

KHÔNG GHI MỘT DÒNG NÀO VÀO DB. `run(spec, path, write=False)`.

Đây là SÀNG LỌC, không phải kết luận: bỏ dòng là hành vi ĐÚNG với phần lớn spec (dòng tiêu đề,
dòng trống, dòng tổng của file). Bảng này chỉ nói "chỗ nào đáng mở ra xem", theo đúng bài học
đã trả giá — quan hệ giữa các con số phải được KHAI, không được suy.

Chạy:
    .venv/bin/python scripts/ra_soat_spec.py                 # tất cả spec, file mới nhất mỗi kỳ
    .venv/bin/python scripts/ra_soat_spec.py --spec atx_doixe
    .venv/bin/python scripts/ra_soat_spec.py --json > ra_soat.json
"""
import argparse
import json
import os
import re
import sys
import traceback

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _HERE)
sys.path.insert(0, os.path.normpath(os.path.join(_HERE, "..")))

import spec_extract as se                                              # noqa: E402

_RE_BO_LOC = re.compile(r"bỏ\s+([\d.,]+)\s+dòng không qua bộ lọc")


def _so(s: str) -> int:
    return int(re.sub(r"[.,\s]", "", s or "0") or 0)


def _bo_loc(canh_bao) -> int:
    """Tổng số dòng bị bộ lọc loại, cộng qua MỌI câu cảnh báo của file.

    Cộng chứ không lấy câu đầu: một file HQKD ngày có 17 sheet × 4 vùng nên đẻ hàng chục câu
    'bỏ N dòng', lấy một câu là báo thiếu cả chục lần.
    """
    return sum(_so(m.group(1)) for c in (canh_bao or []) for m in [_RE_BO_LOC.search(str(c))] if m)


def ra_soat_mot_spec(sp: dict, gioi_han_file: int = None) -> dict:
    ket = {"id": sp.get("id"), "report_type": sp.get("report_type"),
           "folder": ((sp.get("nguon") or {}).get("folder") or ""),
           "so_file": 0, "dong": 0, "bo_loc": 0, "file_rong": [], "loi": [], "canh_bao": {}}
    try:
        files, _ = se.quet_nguon(sp)
        files, _bo = se.loc_file_moi_nhat(sp, files)
    except Exception as e:                                             # noqa: BLE001
        ket["loi"].append(f"quét nguồn: {type(e).__name__}: {e}")
        return ket
    if gioi_han_file:
        files = files[:gioi_han_file]
    for f in files:
        ket["so_file"] += 1
        try:
            r = se.run(sp, f, write=False)
        except Exception as e:                                         # noqa: BLE001
            ket["loi"].append(f"{os.path.basename(f)[:44]}: {type(e).__name__}: {e}")
            continue
        ket["dong"] += int(r.get("dong") or 0)
        ket["bo_loc"] += _bo_loc(r.get("canh_bao"))
        if not r.get("dong"):
            ket["file_rong"].append(os.path.basename(f)[:56])
        for c in (r.get("canh_bao") or []):
            # Gom theo LOẠI (bỏ số) để bảng không thành tường chữ; số lượt vẫn giữ.
            loai = _RE_BO_LOC.sub("bỏ N dòng không qua bộ lọc", str(c))[:96]
            ket["canh_bao"][loai] = ket["canh_bao"].get(loai, 0) + 1
    tong = ket["dong"] + ket["bo_loc"]
    ket["ty_le_bo"] = round(100.0 * ket["bo_loc"] / tong, 1) if tong else None
    return ket


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--spec", help="chỉ rà 1 spec (id)")
    ap.add_argument("--max-file", type=int, default=None, help="giới hạn số file mỗi spec")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()

    ids = ([a.spec] if a.spec
           else sorted(os.path.splitext(os.path.basename(p))[0]
                       for p in __import__("glob").glob(os.path.join(se.SPEC_DIR, "*.json"))
                       if not os.path.basename(p).startswith("_")))
    ket_qua = []
    for i, sid in enumerate(ids, 1):
        if not a.json:
            print(f"[{i}/{len(ids)}] {sid}", file=sys.stderr, flush=True)
        try:
            sp = se.load_spec(sid)
        except Exception as e:                                         # noqa: BLE001
            ket_qua.append({"id": sid, "loi": [f"load_spec: {e}"], "dong": 0, "bo_loc": 0,
                            "so_file": 0, "file_rong": [], "canh_bao": {}, "ty_le_bo": None})
            continue
        try:
            ket_qua.append(ra_soat_mot_spec(sp, a.max_file))
        except Exception:                                              # noqa: BLE001
            ket_qua.append({"id": sid, "loi": [traceback.format_exc()[-300:]], "dong": 0,
                            "bo_loc": 0, "so_file": 0, "file_rong": [], "canh_bao": {},
                            "ty_le_bo": None})

    if a.json:
        print(json.dumps(ket_qua, ensure_ascii=False, indent=1))
        return 0

    xau = [k for k in ket_qua if k["loi"] or k["file_rong"] or (k["ty_le_bo"] or 0) >= 20]
    print()
    print(f"{'spec':30s} {'file':>5s} {'dòng':>8s} {'bỏ':>8s} {'%bỏ':>6s}  ghi chú")
    print("-" * 104)
    for k in sorted(ket_qua, key=lambda x: -(x["ty_le_bo"] or 0)):
        ghi = []
        if k["loi"]:
            ghi.append(f"LỖI: {k['loi'][0][:52]}")
        if k["file_rong"]:
            ghi.append(f"{len(k['file_rong'])} file ra 0 dòng")
        if not k["so_file"]:
            ghi.append("không có file nào")
        print(f"{k['id'][:30]:30s} {k['so_file']:5d} {k['dong']:8d} {k['bo_loc']:8d} "
              f"{(str(k['ty_le_bo']) + '%' if k['ty_le_bo'] is not None else '—'):>6s}  "
              f"{' · '.join(ghi)}")
    print("-" * 104)
    print(f"{len(ket_qua)} spec · {sum(k['dong'] for k in ket_qua)} dòng đọc được · "
          f"{sum(k['bo_loc'] for k in ket_qua)} dòng bị lọc · {len(xau)} spec đáng mở ra xem")
    print("\nĐây là SÀNG LỌC, không phải kết luận: bỏ dòng là hành vi ĐÚNG với phần lớn spec.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
