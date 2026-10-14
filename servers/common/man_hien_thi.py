# -*- coding: utf-8 -*-
"""Lấy SỐ DASHBOARD ĐANG HIỂN THỊ của một màn rồi tìm chỉ tiêu theo từ khoá (14/10/2026).

Gọi chính builder của tc-admin-api trong tiến trình con (không cần đăng nhập, DB chỉ đọc). Số trả về là số FE sẽ vẽ,
KHÔNG phải số tự tính lại — dùng để đặt cạnh số tính từ file nguồn."""
import json
import os
import subprocess

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


def so_hien_thi(man, tu_khoa, period=None, env="prod", grain="month", tu_ngay=None, den_ngay=None, cong_ty=None, khoi=None):
    kw = _kd(tu_khoa)
    if len(kw) < 2:
        raise ValueError("từ khoá quá ngắn")
    res = lay_man(man, period, env, grain, tu_ngay, den_ngay, cong_ty, khoi)
    ra = []
    _duyet(res, "", kw, ra)
    muc = []
    for duong, d in ra[:_NGUONG_HIEN]:
        gon = {k: v for k, v in d.items() if not isinstance(v, (dict, list)) or k in ("theoThang",)}
        if len(json.dumps(gon, ensure_ascii=False, default=str)) > 1500:
            gon = {k: v for k, v in list(gon.items())[:14]}
        muc.append({"vi_tri_trong_ket_qua": duong, "gia_tri": gon})
    return {"man": man, "env": env, "ky": period, "tu_khoa": tu_khoa, "grain": grain, "so_muc_khop": len(ra),
            "cat_bot": max(0, len(ra) - _NGUONG_HIEN), "so_hien_thi": muc,
            "luu_y": "Đây là số builder của dashboard trả ra (đúng số FE vẽ, toàn quyền admin, chưa lọc theo user). "
                     "Đơn vị tiền xem trường donVi (thường tỷ đồng). Trường 'nguon' do chính dashboard khai. "
                     "Muốn đối chiếu với file nguồn thì gọi tiếp check_chi_tiet_vs_tong / check_thang_vs_ngay."}
