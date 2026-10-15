# -*- coding: utf-8 -*-
"""Mốc lượt chạy '====' / 'XONG' của 6 cron dùng `pull_nguon` -> panel "Chạy tự động" chấm đúng màu.

Chạy: python scripts/test_moc_panel_cron.py
      (PANEL_SRC=<đường dẫn source_bridge.py> để đổi bản panel; mặc định checkout TEST của tc-admin-api)

Vì sao có test này (15/10/2026): panel (`source_bridge._job_last_run`) coi một lượt là khối sau dòng
'=' x20 cuối cùng, XANH khi có 'XONG', ĐỎ khi không. 6 cron đăng ký panel 09/10 (duan_quantri,
tuoino_ngay, sodu_nganhang, vay_ngay, htxt_daily, kehoach_doanhthu) không in hai mốc này -> đo log
thật 15/10 cả 12 log (test+prod) đều 'fail — chạy dở / lỗi không rõ' dù rc=0. Test khoá:
  · lượt chạy đúng -> ok; có dòng CẢNH BÁO / KHÔNG VỀ -> warn (vẫn XONG);
  · rc≠0 / exception giữa chừng -> fail, KHÔNG có 'XONG', có lý do 'DỪNG:' (trừ exception);
  · cờ tắt -> off (câu log chứa 'BO QUA: dang TAT');
  · lượt mới nằm SAU lượt cũ trong cùng log -> panel đọc đúng lượt mới.
Không chạm DB / máy gửi: mọi subprocess, kéo file, xoá DB đều giả lập.
"""
import ast
import contextlib
import io
import json
import os
import re
import sys
import tempfile
import traceback
import types
import unicodedata

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.normpath(os.path.join(_HERE, "..", "cron")))
sys.path.insert(0, _HERE)

import pull_nguon  # noqa: E402

PANEL_SRC = os.environ.get("PANEL_SRC") or \
    "/home/itadmin/AI_Dashboard_QT/AI_coding/tc-admin-api/app/bridges/source_bridge.py"


def _nap_panel(src):
    """Lấy ĐÚNG `_job_last_run` + phụ trợ từ source_bridge.py (không import app.* của API)."""
    tree = ast.parse(open(src, encoding="utf-8").read())
    fn = {"_bo_dau_hoa", "_gom_trung", "_job_last_run"}
    var = {"_CANH_BAO_MARK", "_GHI_CHU_NGUON", "_NHAN_NGAN"}
    nodes = [n for n in tree.body
             if (isinstance(n, ast.FunctionDef) and n.name in fn)
             or (isinstance(n, ast.Assign) and any(getattr(t, "id", None) in var for t in n.targets))]
    ns = {"re": re, "unicodedata": unicodedata, "os": os}
    exec(compile(ast.Module(body=nodes, type_ignores=[]), src, "exec"), ns)
    return ns["_job_last_run"]


if not os.path.isfile(PANEL_SRC):
    print(f"BỎ QUA: không thấy {PANEL_SRC}")
    sys.exit(0)
job_last_run = _nap_panel(PANEL_SRC)
loi = []


def chk(ten, dk, them=""):
    print(f"  [{'OK ' if dk else 'SAI'}] {ten}" + (f"  -> {them}" if not dk and them else ""))
    if not dk:
        loi.append(ten)


class R:  # kết quả subprocess.run giả
    def __init__(self, rc=0, out="", err=""):
        self.returncode, self.stdout, self.stderr = rc, out, err


_TMP = tempfile.mkdtemp(prefix="moc_panel_")
pull_nguon.LOGS_DIR = _TMP                       # cờ tắt đọc ở đây, không đụng logs thật
pull_nguon.refresh_metadata = lambda *a, **k: True
pull_nguon.keo = lambda *a, **k: {"xin": 0, "ve": 0, "thieu": []}


def chay(mod, argv, sub, truoc=""):
    """Chạy mod.main() như crontab (stdout+stderr cùng đổ vào log) -> (log, panel)."""
    if sub is not None:
        mod.subprocess.run = sub
    buf = io.StringIO()
    sys.argv = ["x"] + argv
    with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
        try:
            mod.main()
        except SystemExit:
            pass
        except Exception:                         # noqa: BLE001 — traceback cũng vào log như cron thật
            traceback.print_exc()
    p = os.path.join(_TMP, "log.log")
    with open(p, "w", encoding="utf-8") as fh:
        fh.write(truoc + buf.getvalue())
    return buf.getvalue(), job_last_run(p)


def no(*a, **k):
    raise RuntimeError("hỏng giữa chừng")


def tat(job, bat):
    co = os.path.join(_TMP, f"cron_{job}.disabled")
    if bat:
        open(co, "w").close()
    elif os.path.exists(co):
        os.remove(co)


def bo_ca(ten, mod, job, argv, ok, sai, canh_bao=None):
    """4 ca chung: ok / rc≠0 / exception / tắt (+ cảnh báo nếu có)."""
    print(ten)
    log, r = chay(mod, argv, ok)
    chk("lượt đúng -> ok", r["status"] == "ok", f"{r['status']} {r['text']}")
    chk("có '====' và 'XONG'", "=" * 20 in log and "XONG" in log)
    cu = log
    log, r = chay(mod, argv, sai)
    chk("rc≠0 -> fail, không XONG, có 'DỪNG:'",
        r["status"] == "fail" and "XONG" not in log and "DỪNG:" in r["text"], f"{r['status']} {r['text']}")
    log, r = chay(mod, argv, no, truoc=cu)
    chk("exception (sau một lượt ok) -> fail ở lượt MỚI", r["status"] == "fail" and "XONG" not in log,
        f"{r['status']} {r['text']}")
    tat(job, True)
    log, r = chay(mod, argv, ok)
    tat(job, False)
    chk("cờ tắt -> off", r["status"] == "off", f"{r['status']} {r['text']}")
    if canh_bao:
        log, r = chay(mod, argv, canh_bao)
        chk("có CẢNH BÁO -> warn (vẫn XONG)", r["status"] == "warn" and "XONG" in log,
            f"{r['status']} {r['text']}")


# ── htxt_daily ──────────────────────────────────────────────────────────────────────────
import cron_htxt_daily as ch  # noqa: E402
ch.xoa_ban_bi_thay = lambda *a: 0
_AN = json.dumps([{"file_dang_an": "HUNGTHINH::x.xlsx", "written": 3,
                   "canh_bao": ["CẢNH BÁO: file x.xlsx nằm trong hidden_files — 3 dòng vừa ghi bị che"]}])
bo_ca("1. cron_htxt_daily", ch, "htxt_daily", ["--env", "test"],
      ok=lambda *a, **k: R(0, '[{"written": 5}]'),
      sai=lambda *a, **k: R(2, "", "Traceback ..."),
      canh_bao=lambda *a, **k: R(0, _AN))

# ── duan_quantri ────────────────────────────────────────────────────────────────────────
import cron_duan_quantri as cd  # noqa: E402
cd.xoa_ban_bi_thay = lambda *a: 0
pull_nguon.read_metadata = lambda: [{"company": "DUAN", "report_type": "duanpqdbthauphu",
                                     "fileName": "DB báo cáo thầu phụ.xlsx"}]
bo_ca("2. cron_duan_quantri", cd, "duan_quantri", ["--env", "test"],
      ok=lambda *a, **k: R(0, '[{"file": "a.xlsx", "ok": true}]'),
      sai=lambda *a, **k: R(1, "", "Traceback ..."))
pull_nguon.read_metadata = lambda: []
log, r = chay(cd, ["--env", "test"], lambda *a, **k: R(0, "[]"))
chk("không có file nào -> warn (cảnh báo + XONG), không fail", r["status"] == "warn", r["text"])

# ── sodu_nganhang / vay_ngay ────────────────────────────────────────────────────────────
import cron_sodu_nganhang as cs  # noqa: E402
import cron_vay_ngay as cv  # noqa: E402
for i, (mod, job, ghi) in enumerate(((cs, "sodu_nganhang", "ĐÃ GHI: 13 (đơn vị,ngân hàng), 0 bỏ qua."),
                                     (cv, "vay_ngay", "ĐÃ GHI: 8 dòng, 0 bỏ qua.")), 3):
    bo_ca(f"{i}. cron_{job}", mod, job, ["--env", "test"],
          ok=lambda *a, _g=ghi, **k: R(0, f'{{"ok": []}}\n\n{_g}\n'),
          sai=lambda *a, **k: R(1, "", "Traceback ..."))
    log, r = chay(mod, ["--env", "test"], lambda *a, _g=ghi, **k: R(0, f"\n{_g}\n"))
    chk("XONG mang số 'ĐÃ GHI'", ghi in r["text"], r["text"])

# ── tuoino_ngay ─────────────────────────────────────────────────────────────────────────
import cron_tuoino_ngay as ct  # noqa: E402
from datetime import date  # noqa: E402
os.makedirs(os.path.join(_TMP, "XVP", "baocaotuoinongay"), exist_ok=True)
_f = os.path.join(_TMP, "XVP", "baocaotuoinongay", "B.6.XVP.D.20261013.Baocaotuoinophaithu.xlsx")
open(_f, "w").close()
pull_nguon.RECEIVED_DIR = _TMP
_kq = {"o": {"written": 10, "tong_ty": 1.0}}
sys.modules["derive_congno_tuoino"] = types.SimpleNamespace(
    ngay_trong_ten=lambda p: date.today() if "Baocaotuoino" in p else None,
    derive=lambda *a: _kq["o"])
ct.core.moi_truong = lambda job: {"test": {"database_url": "postgresql://gia/khong-dung"}}
pull_nguon.read_metadata = lambda: [{"company": "XVP", "report_type": "baocaotuoinongay",
                                     "fileName": os.path.basename(_f)}]
print("5. cron_tuoino_ngay")
log, r = chay(ct, ["--env", "test"], None)
chk("lượt đúng -> ok", r["status"] == "ok", f"{r['status']} {r['text']}")
_kq["o"] = {"ok": False, "error": "không thấy sheet 'Báo cáo tuổi nợ'"}
log, r = chay(ct, ["--env", "test"], None)
chk("1 file lỗi (exit 1) -> fail, không XONG, có 'DỪNG:'",
    r["status"] == "fail" and "XONG" not in log and "DỪNG:" in r["text"], f"{r['status']} {r['text']}")
tat("tuoino_ngay", True)
log, r = chay(ct, ["--env", "test"], None)
tat("tuoino_ngay", False)
chk("cờ tắt -> off", r["status"] == "off", r["text"])

# ── kehoach_doanhthu ────────────────────────────────────────────────────────────────────
import cron_kehoach_doanhthu as ck  # noqa: E402
from datetime import datetime  # noqa: E402
ck.THU_MUC = _TMP
_nam = datetime.now(ck.VN).year
open(os.path.join(_TMP, f"0.KH.GR.Y.{_nam}.Kehoachdoanhthu.xlsx"), "w").close()
_js = json.dumps({"ok": True, "written": 80, "ky": {f"{_nam}-10": {"tong_ty": 1834.3, "so_khoi": 8}}})
bo_ca("6. cron_kehoach_doanhthu", ck, "kehoach_doanhthu", ["--env", "test"],
      ok=lambda *a, **k: R(0, _js),
      sai=lambda *a, **k: R(1, '{"ok": false}', "Traceback ..."))
ck.THU_MUC = os.path.join(_TMP, "khong-co")
log, r = chay(ck, ["--env", "test"], lambda *a, **k: R(0, _js))
chk("không có file năm -> fail có 'DỪNG:'", r["status"] == "fail" and "DỪNG:" in r["text"], r["text"])

print()
print("KẾT QUẢ:", "TẤT CẢ ĐẠT" if not loi else f"{len(loi)} SAI: {loi}")
sys.exit(1 if loi else 0)
