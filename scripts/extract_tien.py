# -*- coding: utf-8 -*-
"""DẪN XUẤT KHỐI TIỀN (GỘP 1 MỐI) từ Báo cáo tiền tập đoàn — 1 lần đọc, neo theo MÃ:
  · SDT  (03B_SODU_TIEN) — số dư Tiền mặt/Tiền gửi/Tiền vay theo CÔNG TY (sheet 'TC01_SD TIỀN').
  · VAY  (04_VAY)        — dư nợ đầu/cuối + vay thêm/trả nợ theo CÔNG TY × NGÂN HÀNG (mục TIỀN VAY).
  · THUCHI (03_DONGTIEN) — thu/chi theo KHOẢN MỤC (mã 1-9) từng pháp nhân (sheet 'BC THU CHI_T*_<CTY>').

Thay 3 extractor rời (extract_sodu_tien / extract_vay / extract_thuchi) — cùng nguồn, gom để nhất
quán + hết nhân đôi. Neo theo MÃ/section (không cứng vị trí). Import idempotent theo (source_file,
report_type) qua template_filler.import_filled.

Chạy: .venv/bin/python scripts/extract_tien.py <file.xlsx> --period 2026-06
"""
import argparse
import os
import re
import sys

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from servers import template_filler as tf  # noqa: E402
from servers.common import source_catalog as SC  # noqa: E402
from servers.common import be_bridge as bb  # noqa: E402

SD_SHEET = "TC01_SD TIỀN"
# Hậu tố sheet 'BC THU CHI_T<m>_<SUFFIX>' -> mã pháp nhân (T6+ tách AN thành ANTAXI/ANKS, cùng AAG).
_SUFFIX_CO = {"TC": "TC", "VFQN": "VFQN", "XANH": "XVP", "HUNGTHINH": "HT", "AN": "AAG",
              "ANTAXI": "AAG", "ANKS": "AAG"}
# Cột số dư theo LOẠI TIỀN (SD TIỀN) -> cột template 03B.
_SEC_COL = {"tien mat": "Tiền mặt (tỷ)", "tien gui": "Tiền gửi NH (tỷ)",
            "tien vay": "Số dư tiền vay (tỷ) — đối chiếu 04_VAY"}
_NAME_ALIAS = {"an taxi": "AAG", "an ks": "AAG", "an khach san": "AAG", "global ai": "GA",
               "htx xanh tuyen quang": "HTX_XTQ", "htx xanh vinh phuc": "HTX_XVP",
               "xanh tuyen quang": "HTX_XTQ"}
_KYHAN = re.compile(r"(ngan han|trung han|dai han)")   # dòng con kỳ hạn dưới ngân hàng -> bỏ (cộng trùng)


def _norm(s):
    return bb.remove_diacritics("" if s is None else str(s)).strip().lower()


def _resolve_co(name):
    return bb.master.resolve_company_code(name) or _NAME_ALIAS.get(_norm(name))


def _num(v):
    return float(v) if isinstance(v, (int, float)) else None


# ───────────────────────── SD TIỀN -> SDT + VAY ─────────────────────────
def _sd_extract(rows, period):
    """Neo theo MÃ/section: dòng SECTION (cột0+cột1=LOẠI TIỀN), CÔNG TY (cột2), NGÂN HÀNG (cột3).
    SDT = số dư cuối theo công ty × loại tiền. VAY = đầu/cuối + vay thêm/trả theo công ty × bank."""
    hdr_i = next((i for i, r in enumerate(rows[:15])
                  if any("dau ky" in _norm(c) for c in r) and any("den ngay" in _norm(c) for c in r)), None)
    if hdr_i is None:
        return [], [], set(), "Không thấy header (ĐẦU KỲ / ĐẾN NGÀY HIỆN TẠI)"
    hdr = rows[hdr_i]
    c_dau = next(j for j, c in enumerate(hdr) if "dau ky" in _norm(c))
    c_cuoi = next(j for j, c in enumerate(hdr) if "den ngay" in _norm(c))
    deltas = [j for j, c in enumerate(hdr) if _norm(c).startswith("+/-")]

    def _bank_vals(r):
        dau = _num(r[c_dau]) if c_dau < len(r) else None
        cuoi = _num(r[c_cuoi]) if c_cuoi < len(r) else None
        vt = sum(v for j in deltas if j < len(r) and (v := _num(r[j])) and v > 0)
        tn = -sum(v for j in deltas if j < len(r) and (v := _num(r[j])) and v < 0)
        return dau, cuoi, vt, tn

    def _vay_rec(code, bank, dau, cuoi, vt, tn):
        return {"Kỳ": period, "Đơn vị": code, "Ngân hàng": bank,
                "Dư nợ đầu kỳ (tỷ)": round((dau or 0) / 1e9, 9),
                "Vay thêm trong kỳ (tỷ)": round(vt / 1e9, 9),
                "Trả nợ trong kỳ (tỷ)": round(tn / 1e9, 9),
                "Dư nợ cuối kỳ (tỷ)": round((cuoi or 0) / 1e9, 9)}

    sdt, vay, unresolved = [], [], set()
    section = company = None
    vay_seen = False        # 'TIỀN VAY' liệt kê 2 lần cùng số -> chỉ đọc khối ĐẦU (tránh nhân đôi)
    in_vay = False
    orphans = []            # bank ĐỨNG TRƯỚC dòng công ty (layout TC) -> gán khi subtotal cuối khớp

    def _flush(code, subtotal_cuoi):
        nonlocal orphans
        if orphans and subtotal_cuoi is not None:
            tong = sum(b[2] or 0.0 for b in orphans)   # b=(bank,dau,cuoi,vt,tn)
            if abs(tong - subtotal_cuoi) <= max(abs(subtotal_cuoi) * 0.005, 1e4):
                for b in orphans:
                    vay.append(_vay_rec(code, *b))
        orphans = []

    for r in rows[hdr_i + 1:]:
        c0 = str(r[0]).strip() if r[0] not in (None, "") else ""
        c1 = str(r[1]).strip() if len(r) > 1 and r[1] not in (None, "") else ""
        c2 = str(r[2]).strip() if len(r) > 2 and r[2] not in (None, "") else ""
        c3 = str(r[3]).strip() if len(r) > 3 and r[3] not in (None, "") else ""
        if c0 and c1:                                   # dòng SECTION
            section = next((k for k in _SEC_COL if k in _norm(c1)), None)
            in_vay = ("tien vay" == section) and not vay_seen
            if section == "tien vay":
                vay_seen = True
            company, orphans = None, []
            continue
        if not section:
            continue
        if c2:                                          # dòng CÔNG TY (subtotal)
            code = _resolve_co(c2)
            cuoi = _num(r[c_cuoi]) if c_cuoi < len(r) else None
            if code is None:
                unresolved.add(c2)
                company, orphans = None, []
                continue
            if in_vay:
                _flush(code, cuoi)                        # gán các bank orphan (đứng trước) cho cty này
            company = code
            if cuoi is not None:
                sdt.append({"Kỳ": period, "Đơn vị": code, _SEC_COL[section]: round(cuoi / 1e9, 9)})
            continue
        if in_vay and c3 and not _KYHAN.search(_norm(c3)):   # dòng NGÂN HÀNG
            dau, cuoi, vt, tn = _bank_vals(r)
            if not any((dau, cuoi, vt, tn)):
                continue
            if company:                                  # bank đứng SAU công ty của nó (layout thường)
                vay.append(_vay_rec(company, c3, dau, cuoi, vt, tn))
            else:                                        # bank đứng TRƯỚC công ty -> chờ subtotal
                orphans.append((c3, dau, cuoi, vt, tn))
    return sdt, vay, unresolved, None


# ───────────────────────── BC THU CHI -> THUCHI ─────────────────────────
_COL_KY = "Kỳ / Ngày"
_COL_CTY = "Mã Công ty (auto từ CC)"
_COL_LOAI = "Loại (Thu/Chi)"
_COL_KM = "Khoản mục (Thu bán hàng, Thu đầu tư, Chi NCC, Chi tài chính, Chi đầu tư TS…)"
_COL_TH = "Thực hiện (tỷ)"


def _co_of_sheet(sheet):
    up = sheet.upper()
    if "TỔNG" in up or "TONG" in up:
        return None
    return _SUFFIX_CO.get(up.rsplit("_", 1)[-1].strip())


def _val_col(rows):
    """Cột giá trị tổng = 'TM + TG + T.VAY' (gồm vay, theo yêu cầu). Mặc định 3 (cột D)."""
    for r in rows[:12]:
        for j, c in enumerate(r):
            s = str(c).strip().upper() if c is not None else ""
            if "TM" in s and "VAY" in s:
                return j
    return 3


def _thuchi_extract(wb, period):
    recs = []
    sheets = [s for s in wb.sheetnames if re.search(r"THU CHI_T\d", s.upper()) and _co_of_sheet(s)]
    for sh in sheets:
        co = _co_of_sheet(sh)
        # HT (Xe tải Hưng Thịnh): DÒNG TIỀN lấy từ sheet 'LCTT' của BCTC (spec 50 chỉ tiêu, dòng
        # 18-19, khớp kế toán tới đồng — xem agent_cli._derive_lctt_ht), KHÔNG lấy từ Báo cáo tiền
        # tập đoàn để tránh NẠP ĐÔI + khác nguồn. (SDT/VAY của HT vẫn lấy bình thường ở _sd_extract.)
        if co == "HT":
            continue
        rows = [list(r) for r in wb[sh].iter_rows(values_only=True)]
        vc = _val_col(rows)
        sec = None                                       # 'A' (thu, mục I) / 'B' (chi, mục II)
        for r in rows:
            c0 = str(r[0]).strip() if r[0] not in (None, "") else ""
            c1 = str(r[1]).strip() if len(r) > 1 and r[1] not in (None, "") else ""
            if c0 == "I":
                sec = "A"; continue
            if c0 == "II":
                sec = "B"; continue
            if sec and re.fullmatch(r"\d+", c0) and c1:  # khoản mục CẤP 1 (1,2,3…) — bỏ 1.1 / '+…'
                v = r[vc] if vc < len(r) and isinstance(r[vc], (int, float)) else None
                if v is None:
                    continue
                recs.append({_COL_KY: period, _COL_CTY: co,
                             _COL_LOAI: "Thu" if sec == "A" else "Chi",
                             _COL_KM: c1, _COL_TH: round(v / 1e9, 9)})
    return recs, sheets


def extract(path: str, period: str, cong_ty: str = None) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)
    src = SC.source_id_from_path(path)
    out = {"ok": True}

    # 1) SDT + VAY
    if SD_SHEET in wb.sheetnames:
        rows = [list(r) for r in wb[SD_SHEET].iter_rows(values_only=True)]
        sdt, vay, unresolved, err = _sd_extract(rows, period)
        if err:
            out["sd_error"] = err
        if unresolved:
            out["unresolved_companies"] = sorted(unresolved)
        if sdt:
            p = os.path.join(tf.FILLED_DIR, f"TIEN_{period}_03B_SODU_TIEN.xlsx")
            tf.fill("03B_SODU_TIEN", sdt, p)
            out["sdt"] = tf.import_filled(p, cong_ty=None, source_file=src).get("rows_imported")
        if vay:
            p = os.path.join(tf.FILLED_DIR, f"TIEN_{period}_04_VAY.xlsx")
            tf.fill("04_VAY", vay, p)
            out["vay"] = tf.import_filled(p, cong_ty=None, source_file=src).get("rows_imported")
    else:
        out["sd_error"] = f"Không thấy sheet {SD_SHEET}"

    # 2) THUCHI
    tc, sheets = _thuchi_extract(wb, period)
    if tc:
        p = os.path.join(tf.FILLED_DIR, f"TIEN_{period}_03_DONGTIEN.xlsx")
        tf.fill("03_DONGTIEN", tc, p)
        out["thuchi"] = tf.import_filled(p, cong_ty=None, source_file=src).get("rows_imported")
        out["thuchi_sheets"] = sheets
    if not any(k in out for k in ("sdt", "vay", "thuchi")):
        return {"ok": False, "error": "Không bóc được SDT/VAY/THUCHI nào", **out}
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--period", required=True)
    ap.add_argument("--cong-ty", dest="cong_ty", default=None)
    a = ap.parse_args()
    import json
    print(json.dumps(extract(a.file, a.period, a.cong_ty), ensure_ascii=False))
