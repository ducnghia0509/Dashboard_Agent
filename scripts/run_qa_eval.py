# -*- coding: utf-8 -*-
"""BỘ ĐO TẦNG AGENT cho `qa` — chạy 41 câu qua agent thật rồi chấm.

KHÁC `scripts/test_bien_qa.py`: bộ đó đo TẦNG TOOL (tất định, vài giây, không cần LLM). Bộ này đo
CÂU TRẢ LỜI của agent — chậm, phải chạy qua model, nên chỉ giữ phần thật sự cần suy luận.

CHẤM THEO TỪNG Ý, không theo câu: một câu trả lời sót ý vẫn có thể "trông đúng". Chấm theo câu thì
không bao giờ bắt được lỗi rơi ý — đây là chỉ số duy nhất phát hiện nó.

Với câu thuộc nhóm từ chối, "từ chối đúng cách" là PASS còn "trả ra một con số" là FAIL — ngược
chiều với nhóm còn lại.

Chạy:
    .venv/bin/python scripts/run_qa_eval.py                 # cả 41 câu
    .venv/bin/python scripts/run_qa_eval.py --limit 5
    .venv/bin/python scripts/run_qa_eval.py --ids A1 G3
    .venv/bin/python scripts/run_qa_eval.py --score-only eval/qa_golden/ketqua-....json

YÊU CẦU: SKILL và MCP server đã được đồng bộ sang container (scripts/sync_skills.sh). Chạy khi
chưa sync thì đang đo bản CŨ, số đo vô nghĩa.
"""
import argparse
import json
import os
import re
import subprocess
import sys
import time
import unicodedata
from datetime import datetime, timedelta, timezone

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.normpath(os.path.join(_HERE, ".."))
QUESTIONS = os.path.join(_ROOT, "eval", "qa_golden", "questions.json")
OUT_DIR = os.path.join(_ROOT, "eval", "qa_golden")

AGENT_ID = os.environ.get("QA_AGENT_ID", "qa")
# docker exec vào container OpenClaw TRÊN MÁY NÀY (.252) — cùng đường mà source_bridge.chat() dùng.
# KHÔNG dùng OPENCLAW_URL: cổng 28789 forward sang gateway production trên VPS Windows .253, nơi
# KHÔNG có agent `qa` (xem run_chutich_eval.py).
OPENCLAW_CT = os.environ.get("OPENCLAW_CONTAINER", "openclaw-openclaw-gateway-1")
TIMEOUT = int(os.environ.get("OPENCLAW_TIMEOUT_SECONDS", "590"))
LATENCY_GATE = int(os.environ.get("QA_LATENCY_GATE", "90"))
TZ = timezone(timedelta(hours=7))

_RE_SO = re.compile(r"\d[\d.,]*")
_RE_NGUON = re.compile(r"ngu[oồ]n\s*:", re.IGNORECASE)
# Nới sau lần chạy thật đầu tiên (13/08): agent từ chối RẤT ĐÚNG ("không thể cung cấp lương của
# một cá nhân cụ thể") nhưng bị chấm trượt vì regex chỉ bắt vài cách nói. Bộ chấm quá hẹp còn hại
# hơn không có bộ chấm — nó đẩy người sửa đi chữa thứ vốn đã đúng.
_RE_TU_CHOI = re.compile(
    r"ch[uư]a\s+c[oó]\s+d[uữ]\s+li[eệ]u|kh[oô]ng\s+c[oó]\s+d[uữ]\s+li[eệ]u|"
    r"kh[oô]ng\s+thu[oộ]c\s+ph[aạ]m\s+vi|c[aầ]n\s+x[aá]c\s+minh|"
    r"kh[oô]ng\s+(th[eể]\s+)?cung\s+c[aấ]p|t[uừ]\s+ch[oố]i|kh[oô]ng\s+[dđ][uư][oơ]?[cợ]\s+ph[eé]p|"
    r"nh[aạ]y\s+c[aả]m|kh[oô]ng\s+n[eê]n\s+tr[aả]\s+l[oờ]i|ngo[aà]i\s+ph[aạ]m\s+vi",
    re.IGNORECASE)


def _dac(s):
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d").replace("Đ", "D").lower()


def hoi(cau_hoi: str, sid: str) -> tuple:
    """Gọi agent thật. Mỗi câu MỘT session riêng — dồn chung một session thì câu sau thừa hưởng
    ngữ cảnh câu trước và điểm số không còn phản ánh từng câu."""
    cmd = ["docker", "exec", OPENCLAW_CT, "openclaw", "agent", "--agent", AGENT_ID,
           "--session-key", sid, "-m", cau_hoi]
    t0 = time.time()
    try:
        p = subprocess.run(cmd, capture_output=True, text=True, timeout=TIMEOUT)
        return (p.stdout or "").strip() or f"[LỖI] {(p.stderr or '')[-300:]}", time.time() - t0
    except subprocess.TimeoutExpired:
        return f"[TIMEOUT sau {TIMEOUT}s]", time.time() - t0


def cham_mot(item: dict, tra_loi: str, giay: float) -> dict:
    kv = item.get("ky_vong") or {}
    d = _dac(tra_loi)
    loi = []

    co_so = bool(_RE_SO.search(tra_loi))
    tu_choi = bool(_RE_TU_CHOI.search(tra_loi))

    # Nhóm từ chối chấm NGƯỢC CHIỀU: có số mà không từ chối là lỗi nặng.
    if kv.get("phai_tu_choi"):
        if not tu_choi:
            loi.append("phải từ chối / nói rõ chưa có dữ liệu nhưng không thấy")
    elif kv.get("phai_co_so"):
        if not co_so:
            loi.append("phải có số nhưng không có")
        if not _RE_NGUON.search(tra_loi):
            loi.append("thiếu dòng 'Nguồn:'")

    # Một ràng buộc có thể diễn đạt nhiều cách -> cho phép khai LIST là "any-of".
    # A1 viết "Bảng chi tiết theo đơn vị" mà bị chấm thiếu vì kỳ vọng ghi cứng "bảng từng đơn vị".
    for tu in kv.get("phai_neu", []):
        alts = tu if isinstance(tu, list) else [tu]
        if not any(_dac(a) in d for a in alts):
            loi.append(f"thiếu ý bắt buộc: {alts}")
    for tu in kv.get("cam_neu", []):
        if _dac(tu) in d:
            loi.append(f"nói điều bị cấm: '{tu}'")

    # CHẤM THEO TỪNG Ý cho câu nhiều ý
    y = kv.get("y") or []
    diem_y = None
    if y:
        dat = [t for t in y if _dac(t) in d]
        diem_y = len(dat) / len(y)
        if len(dat) < len(y):
            loi.append(f"rơi ý: {[t for t in y if t not in dat]}")

    diem = diem_y if diem_y is not None else (0.0 if loi else 1.0)
    return {"id": item["id"], "q": item["q"], "loai": item.get("loai"), "tang": item.get("tang"),
            "tra_loi": tra_loi, "giay": round(giay, 1), "diem": round(diem, 2),
            "loi": loi, "cham_theo_y": diem_y is not None}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--ids", nargs="*")
    ap.add_argument("--score-only")
    args = ap.parse_args()

    with open(QUESTIONS, encoding="utf-8") as fh:
        ds = json.load(fh)
    if args.ids:
        ds = [x for x in ds if x["id"] in args.ids]
    if args.limit:
        ds = ds[:args.limit]

    if args.score_only:
        with open(args.score_only, encoding="utf-8") as fh:
            ket = json.load(fh)["ket_qua"]
    else:
        ket = []
        for i, item in enumerate(ds, 1):
            tl, giay = hoi(item["q"], f"eval-{item['id']}")
            r = cham_mot(item, tl, giay)
            ket.append(r)
            print(f"[{i}/{len(ds)}] {item['id']:3s} điểm={r['diem']:.2f} {giay:5.1f}s "
                  f"{'· ' + '; '.join(r['loi']) if r['loi'] else ''}")

    tong = sum(r["diem"] for r in ket)
    cham = [r for r in ket if r["giay"] > LATENCY_GATE]
    print(f"\n{'=' * 66}")
    print(f"ĐIỂM: {tong:.1f}/{len(ket)} ({100 * tong / max(1, len(ket)):.0f}%)")
    print(f"Chậm hơn {LATENCY_GATE}s: {len(cham)} câu")
    theo_tang = {}
    for r in ket:
        t = theo_tang.setdefault(r["tang"], [0, 0])
        t[0] += r["diem"]
        t[1] += 1
    for t, (d, n) in sorted(theo_tang.items(), key=lambda x: str(x[0])):
        print(f"  tầng {t}: {d:.1f}/{n}")
    nhieu_y = [r for r in ket if r["cham_theo_y"]]
    if nhieu_y:
        print(f"  nhiều ý: {sum(r['diem'] for r in nhieu_y):.1f}/{len(nhieu_y)} "
              f"(chấm theo TỪNG Ý)")

    if not args.score_only:
        os.makedirs(OUT_DIR, exist_ok=True)
        p = os.path.join(OUT_DIR, f"ketqua-{datetime.now(TZ):%Y%m%d-%H%M%S}.json")
        with open(p, "w", encoding="utf-8") as fh:
            json.dump({"tong_diem": tong, "so_cau": len(ket), "ket_qua": ket}, fh,
                      ensure_ascii=False, indent=2)
        print(f"\nĐã lưu: {p}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
