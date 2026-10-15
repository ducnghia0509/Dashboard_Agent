# -*- coding: utf-8 -*-
"""Lấy SỐ DASHBOARD ĐANG HIỂN THỊ của một màn rồi tìm chỉ tiêu theo từ khoá (14/10/2026).

Gọi chính builder của tc-admin-api trong tiến trình con (không cần đăng nhập, DB chỉ đọc). Số trả về là số FE sẽ vẽ,
KHÔNG phải số tự tính lại — dùng để đặt cạnh số tính từ file nguồn."""
import json
import os
import re
import subprocess

from . import man_fe as mf

API_DIR = os.environ.get("TC_API_DIR") or "/home/itadmin/apps/tc-console/tc-admin-api"
API_PY = os.path.join(API_DIR, ".venv", "bin", "python")
_DUMP = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "scripts", "man_hien_thi_dump.py")
_PORT = {"prod": "5434", "test": "5435"}
_NGUONG_HIEN = 16  # tối đa số mục trả về


def _env_api():
    ev = {}
    with open(os.path.join(API_DIR, ".env"), encoding="utf-8") as fh:
        for ln in fh:
            ln = ln.strip()
            if ln and not ln.startswith("#") and "=" in ln:
                k, _, v = ln.partition("=")
                ev[k.strip()] = v.strip().strip("'\"")
    return ev


def lay_man(man, period=None, env="prod", grain="month", tu_ngay=None, den_ngay=None, cong_ty=None, khoi=None, timeout=120):
    if env not in _PORT:
        raise ValueError("env phải là prod hoặc test")
    ev = _env_api()
    url = ev.get("DATABASE_URL") or ""
    for p in _PORT.values():
        url = url.replace(f":{p}/", f":{_PORT[env]}/")
    e = {**os.environ, **ev, "DATABASE_URL": url, "PYTHONPATH": API_DIR,
         "PGOPTIONS": "-c default_transaction_read_only=on -c statement_timeout=60000"}
    arg = {"man": man, "period": period, "grain": grain, "tu_ngay": tu_ngay, "den_ngay": den_ngay,
           "cong_ty": cong_ty, "khoi": khoi}
    r = subprocess.run([API_PY, _DUMP, json.dumps(arg)], capture_output=True, text=True, timeout=timeout, cwd=API_DIR, env=e)
    if r.returncode != 0:
        raise RuntimeError("không dựng được số hiển thị: " + (r.stderr or "")[-300:])
    out = json.loads(r.stdout)
    if isinstance(out, dict) and out.get("loi"):
        raise ValueError(f"{out['loi']}. Kỳ có: {', '.join(out.get('ky_co') or [])}")
    return out


def _kd(s):
    from .mapping_check import _kd as k
    return k(s)


def _duyet(node, duong, kw, ra):
    """Duyệt JSON kết quả; gom các dict (mục KPI/dòng bảng) có chuỗi chứa từ khoá hoặc mã."""
    if len(ra) >= _NGUONG_HIEN * 3:
        return
    if isinstance(node, dict):
        if any(isinstance(v, str) and kw in _kd(v) for v in node.values()):
            ra.append((duong, node))
            return                                           # dừng ở mục khớp, khỏi nhặt con của nó
        for k, v in node.items():
            _duyet(v, f"{duong}.{k}" if duong else k, kw, ra)
    elif isinstance(node, list):
        for i, v in enumerate(node):
            _duyet(v, f"{duong}[{i}]", kw, ra)


def _lay(node, truong):
    for k in truong.split("."):
        if not isinstance(node, dict) or k not in node:
            return None
        node = node[k]
    return node


def _ngu_canh(goc):
    """Trường mô tả ĐI KÈM con số (ngày chốt, nguồn, ghi chú, đơn vị) — để bot nói được số đó tính tới ngày nào,
    lấy từ đâu. Thiếu phần này bot đọc tồn kho (ảnh chụp tại ngày) thành 'luỹ kế'."""
    if not isinstance(goc, dict):
        return {}
    return {k: v for k, v in goc.items()
            if re.match(r"(ngay|nguon|ghiChu|donVi|ky|thieuNguon|tonHoaDonTuDong)", k)
            and len(json.dumps(v, ensure_ascii=False, default=str)) <= 600}


def so_hien_thi(man, tu_khoa, period=None, env="prod", grain="month", tu_ngay=None, den_ngay=None, cong_ty=None, khoi=None):
    kw = _kd(tu_khoa)
    if len(kw) < 2:
        raise ValueError("từ khoá quá ngắn")
    # Nhận cả TÊN màn ("Quản lý tồn kho SRVF") chứ không chỉ mã — bot không được hỏi ngược người dùng mã màn.
    ung_vien = mf.tim_man(man)
    if not ung_vien:
        return {"loi": f"không nhận ra màn '{man}'", "danh_muc_man": mf.danh_muc_man()}
    man_chon = ung_vien[0]
    man = man_chon["man"]
    res = lay_man(man, period, env, grain, tu_ngay, den_ngay, cong_ty, khoi)
    goc = res.get(man) if isinstance(res, dict) and isinstance(res.get(man), dict) else res

    # 1) Theo NHÃN THẺ trên FE (đúng chữ người dùng nhìn thấy) -> trường dữ liệu -> giá trị.
    tu = kw.split()
    theo_nhan = []
    for n in mf.nhan_tren_man(man):
        nk = _kd(n["nhan"])
        diem = sum(1 for w in tu if w in nk.split()) / len(tu)
        if kw in nk or diem >= 0.75:
            theo_nhan.append((-(2 if nk == kw else 1 if kw in nk else diem), n))
    theo_nhan = [{"nhan_tren_man": n["nhan"], "truong": n["truong"], "gia_tri": _lay(goc, n["truong"])}
                 for _, n in sorted(theo_nhan, key=lambda x: x[0])]

    # 2) Theo chuỗi trong JSON (mã chỉ tiêu CT33, tên dòng bảng…) — cách cũ, hợp với màn bảng điểm.
    ra = []
    _duyet(res, "", kw, ra)
    muc = []
    for duong, d in ra[:_NGUONG_HIEN]:
        gon = {k: v for k, v in d.items() if not isinstance(v, (dict, list)) or k in ("theoThang",)}
        if len(json.dumps(gon, ensure_ascii=False, default=str)) > 1500:
            gon = {k: v for k, v in list(gon.items())[:14]}
        muc.append({"vi_tri_trong_ket_qua": duong, "gia_tri": gon})
    out = {"man": man, "ten_man": man_chon["ten_man"], "nhom_man": man_chon["nhom"], "env": env, "ky": period,
           "tu_khoa": tu_khoa, "grain": grain, "tu_ngay": tu_ngay, "den_ngay": den_ngay,
           "khop_theo_nhan_the": theo_nhan, "so_muc_khop": len(ra), "cat_bot": max(0, len(ra) - _NGUONG_HIEN),
           "so_hien_thi": muc, "ngu_canh_cua_so": _ngu_canh(goc),
           "luu_y": "Đây là số builder của dashboard trả ra (đúng số FE vẽ, toàn quyền admin, chưa lọc theo user). "
                    "Đơn vị tiền xem trường donVi (thường tỷ đồng). Trường 'nguon' do chính dashboard khai. "
                    "Chỉ tiêu TỒN KHO/SỐ DƯ/CÔNG NỢ là ẢNH CHỤP tại ngày chốt (xem ngu_canh_cua_so.ngayChot*), "
                    "KHÔNG phải luỹ kế cộng dồn các ngày. "
                    "Muốn đối chiếu với file nguồn thì gọi tiếp check_ton_kho_xe / check_chi_tiet_vs_tong / check_thang_vs_ngay."}
    if len(ung_vien) > 1:
        out["man_khac_co_the"] = [{"man": m["man"], "ten_man": m["ten_man"], "nhom": m["nhom"]} for m in ung_vien[1:3]]
    if not theo_nhan and not ra:
        # Không khớp gì: đưa NGUYÊN danh sách thẻ của màn kèm giá trị để bot tự chọn đúng thẻ, KHÔNG trả rỗng
        # rồi đòi người dùng "tên chỉ tiêu nguyên văn".
        out["nhan_co_tren_man"] = [{"nhan_tren_man": n["nhan"], "truong": n["truong"], "gia_tri": _lay(goc, n["truong"])}
                                   for n in mf.nhan_tren_man(man)]
    return out
