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
        return [], [], set(), None, "Không thấy header (ĐẦU KỲ / ĐẾN NGÀY HIỆN TẠI)"
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
        # DƯ NỢ đầu/cuối theo BANK (giữ, nguồn SD TIỀN). Vay thêm/trả=0 ở đây — nay lấy PER-BANK từ
        # báo cáo ngân hàng BCTH 2 (cột Tổng vay/Đã thanh toán) trong extract_vay_kyhan (chạy sau).
        # KHÔNG dùng delta SD TIỀN (sai) cũng KHÔNG dùng BC THU CHI mức công ty (làm bẩn chart theo NH).
        return {"Kỳ": period, "Đơn vị": code, "Ngân hàng": bank,
                "Dư nợ đầu kỳ (tỷ)": round((dau or 0) / 1e9, 9),
                "Vay thêm trong kỳ (tỷ)": 0.0,
                "Trả nợ trong kỳ (tỷ)": 0.0,
                "Dư nợ cuối kỳ (tỷ)": round((cuoi or 0) / 1e9, 9)}

    sdt, vay, unresolved, dachi = [], [], set(), None
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
        # 'Đã chi nhưng chưa có chứng từ' (memo dưới TIỀN MẶT) -> cảnh báo TC (guide 21/7). Bắt theo TÊN
        # dòng (c2), chặn TRƯỚC nhánh công ty (nếu không sẽ rơi vào unresolved). Lấy cả đầu & cuối kỳ.
        if section == "tien mat" and c2 and "da chi" in _norm(c2) and "chung tu" in _norm(c2):
            dachi = {"cuoi": _num(r[c_cuoi]) if c_cuoi < len(r) else None,
                     "dau": _num(r[c_dau]) if c_dau < len(r) else None}
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
    return sdt, vay, unresolved, dachi, None


# ───────────────────────── BC THU CHI -> THUCHI ─────────────────────────
_COL_KY = "Kỳ / Ngày"
_COL_CTY = "Mã Công ty (auto từ CC)"
_COL_LOAI = "Loại (Thu/Chi)"
_COL_KM = "Khoản mục (Thu bán hàng, Thu đầu tư, Chi NCC, Chi tài chính, Chi đầu tư TS…)"
_COL_HT = "Hình thức (TM/TG/Đối trừ CN/Vay)"   # -> dim2 (import map sẵn); tách TM/gửi/vay
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
        # HT (Xe tải Hưng Thịnh): DÒNG TIỀN nay lấy từ 'BC THU CHI_T*_HUNGTHINH' NHƯ MỌI PHÁP NHÂN
        # (ĐẢO 2026-07-21 — xem memory ht-dongtien-from-lctt: bỏ GỌI _derive_lctt_ht, lấy chung nguồn
        # để khớp báo cáo tập đoàn + hết nạp đôi). TRƯỚC ĐÂY skip HT ở đây (lấy từ LCTT); nay KHÔNG skip
        # nữa — _derive_lctt_ht đã không được gọi nên KHÔNG đếm đôi. (SDT/VAY HT vẫn ở _sd_extract.)
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
                # TÁCH THEO HÌNH THỨC (spec 2026-07-23): cột D (vc) = E(TM) + F(NH/gửi) + G(TVAY) — verify
                # khớp per khoản mục. Emit 1 dòng/hình thức (Hình thức -> dim2) THAY dòng tổng D: Σ = D nên
                # inflow/outflow (Σ amount) KHÔNG đổi; có thêm 'tiền mặt/gửi thu-chi' lọc theo dim2. Bỏ dòng =0.
                loai = "Thu" if sec == "A" else "Chi"
                for _off, _ht in ((1, "Tiền mặt"), (2, "Tiền gửi"), (3, "Tiền vay")):
                    _j = vc + _off
                    _v = r[_j] if _j < len(r) and isinstance(r[_j], (int, float)) else None
                    if not _v:
                        continue
                    recs.append({_COL_KY: period, _COL_CTY: co, _COL_LOAI: loai,
                                 _COL_KM: c1, _COL_HT: _ht, _COL_TH: round(_v / 1e9, 9)})
    return recs, sheets


def extract(path: str, period: str, cong_ty: str = None) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)
    src = SC.source_id_from_path(path)
    out = {"ok": True}

    # 1) SDT + VAY
    if SD_SHEET in wb.sheetnames:
        rows = [list(r) for r in wb[SD_SHEET].iter_rows(values_only=True)]
        sdt, vay, unresolved, dachi, err = _sd_extract(rows, period)
        # VAY: dòng per-BANK giữ dư nợ đầu/cuối (SD TIỀN); vay thêm/trả=0 ở đây — nay lấy PER-BANK từ
        # báo cáo ngân hàng (BCTH 2) trong extract_vay_kyhan (chạy sau, wired autofill). KHÔNG còn emit
        # dòng '(Cả <cty>)' mức công ty (làm bẩn chart theo ngân hàng + chart đi vay/trả trống).
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
        # 'Đã chi nhưng chưa có chứng từ' (cảnh báo TC) -> DACHI_CCT direct-insert (KHÔNG qua template).
        # amount=cuối, amount2=đầu. Idempotent (DELETE trước). Twin lấy dataset_id/ngay/khoi từ dòng SDT
        # TC vừa nạp ở trên. Port từ extract_sodu_tien (legacy) -> nay tất định trong extract_tien (wired).
        if dachi is not None:
            from servers.common import be_bridge as bb
            import json as _json
            _db = bb.db.get_db()
            _db.execute("DELETE FROM raw_rows WHERE source_file=? AND period_month=? AND report_type=?",
                        (src, period, "DACHI_CCT"))
            _tw = _db.execute("SELECT dataset_id, ngay, khoi FROM raw_rows WHERE source_file=? AND "
                              "period_month=? AND report_type='SDT' AND cong_ty='TC' LIMIT 1",
                              (src, period)).fetchone()
            if _tw:
                _db.execute(
                    "INSERT INTO raw_rows(dataset_id,report_type,row_index,ngay,cong_ty,khoi,cost_center,"
                    "period_month,amount,amount2,dim1,dim2,dim3,payload,source_file) "
                    "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (_tw["dataset_id"], "DACHI_CCT", 6000000, _tw["ngay"], "TC", _tw["khoi"], None, period,
                     round((dachi["cuoi"] or 0) * 1e-9, 9), round((dachi["dau"] or 0) * 1e-9, 9),
                     "Đã chi nhưng chưa có chứng từ", None, None,
                     _json.dumps({"unit": "ty", "nguon": "TC01_SD TIỀN - Đã chi chưa có chứng từ"},
                                 ensure_ascii=False), src))
                _db.commit()
                out["dachi_cct"] = round((dachi["cuoi"] or 0) * 1e-9, 9)
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
