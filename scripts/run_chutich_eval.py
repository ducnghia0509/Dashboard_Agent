# -*- coding: utf-8 -*-
"""Chạy 71 câu hỏi của Chủ tịch qua agent qa_chutich và chấm theo cổng chất lượng.

Cổng chất lượng (đặt ra vì rủi ro lớn nhất KHÔNG phải trả lời chậm, mà là trả lời một con số
bịa cho đúng người không nên nhận số bịa):

  🟢 49 câu — phải có SỐ + dòng `Nguồn:` + mốc as_of.
  🟡 12 câu — phải mở đầu bằng cảnh báo "cần xác minh"; nhóm kiểm toán còn bị cấm từ khẳng định.
  🔴 10 câu — phải TỪ CHỐI và nêu đúng phòng phụ trách. Trả ra số ở nhóm này là LỖI NẶNG.

Chạy:
    .venv/bin/python scripts/run_chutich_eval.py                    # cả 71 câu
    .venv/bin/python scripts/run_chutich_eval.py --limit 5          # thử nhanh
    .venv/bin/python scripts/run_chutich_eval.py --light red        # chỉ nhóm đỏ
    .venv/bin/python scripts/run_chutich_eval.py --score-only out.json
"""
import argparse
import json
import os
import re
import subprocess
import sys
import time
from datetime import datetime, timezone, timedelta

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.normpath(os.path.join(_HERE, ".."))
QUESTIONS = os.path.join(_ROOT, "eval", "chutich", "questions.json")
OUT_DIR = os.path.join(_ROOT, "eval", "chutich")

AGENT_ID = os.environ.get("CHUTICH_AGENT_ID", "qa_chutich")
# Gọi qua `docker exec` vào container OpenClaw TRÊN MÁY NÀY (.252) — ĐÚNG đường mà
# `source_bridge.py::chat()` (POST /import/chat) dùng cho luồng QA của dashboard.
# KHÔNG dùng OPENCLAW_URL: biến đó trỏ 127.0.0.1:28789, mà cổng 28789 là TCP forwarder sang
# OpenClaw production trên VPS Windows 10.254.20.253 — nơi phục vụ các bot chat
# (`tckt`, `ai_thinhcuong_bot`...), KHÔNG có agent `qa`/`qa_chutich`. Gọi nhầm sang đó chỉ nhận
# "Unknown agent". Hai gateway, hai luồng tách biệt.
OPENCLAW_CT = os.environ.get("OPENCLAW_CONTAINER", "openclaw-openclaw-gateway-1")
TIMEOUT = int(os.environ.get("OPENCLAW_TIMEOUT_SECONDS", "590"))
TZ = timezone(timedelta(hours=7))

# Ngưỡng p95. Bản đầu đặt 20s theo GIẢ ĐỊNH "đọc brief tính sẵn thì phải rất nhanh". Đo thật
# trên 9router/my1 (07/08): kể cả câu chỉ đọc 1 sheet rồi từ chối vẫn mất 10-40s — chi phí nằm ở
# lượt suy luận của model, không nằm ở việc đọc file. 20s là ngưỡng không bao giờ đạt được nên
# vô dụng; nâng lên 60s cho khớp thực đo, vẫn đủ chặt để bắt hồi quy (trước khi có brief, các câu
# xếp hạng toàn tập đoàn phải mở hàng chục file, tính bằng phút).
LATENCY_GATE = int(os.environ.get("CHUTICH_LATENCY_GATE", "60"))

# Dấu hiệu có SỐ thật: chuỗi chữ số có thể kèm phân tách nghìn/thập phân. Loại các số chỉ là
# mã câu (L5-02) hay ngày tháng đứng một mình bằng cách đòi độ dài >= 1 và có ngữ cảnh tiền/%.
_RE_NUMBER = re.compile(r"\d[\d.,]*")
_RE_NGUON = re.compile(r"ngu[oồ]n\s*:", re.IGNORECASE)
# Mốc thời gian: chấp nhận CẢ ISO (2026-08-07, đúng như _META ghi) LẪN kiểu Việt (07/08/2026).
# Bản đầu chỉ khớp ISO và đã chấm TRƯỢT oan 2 câu trả lời hoàn hảo chỉ vì agent viết ngày theo
# kiểu người Việt đọc. Cổng chất lượng phải đo CÓ mốc thời gian hay không, không đo định dạng.
_RE_ASOF = re.compile(r"\d{4}-\d{2}-\d{2}|\d{1,2}/\d{1,2}/\d{4}")
_RE_CANHBAO = re.compile(r"c[aầ]n\s+x[aá]c\s+minh", re.IGNORECASE)
_RE_TUCHOI = re.compile(r"ch[uư]a\s+c[oó]\s+d[uữ]\s+li[eệ]u", re.IGNORECASE)
# Từ khẳng định bị cấm ở nhóm kiểm toán — agent chỉ được nói "cần xác minh".
# KHÔNG dùng phép "có cảnh báo ở gần" để miễn trừ: luật 🟡 đã bắt buộc cảnh báo nằm ngay đầu
# câu trả lời, nên mọi cửa sổ lân cận đều thấy cảnh báo và phép kiểm tra không bao giờ kích
# hoạt — "đây là hành vi gian lận rõ ràng" vẫn lọt. Thay bằng: chặn thẳng từ khẳng định, chỉ
# tha khi 'gian lận' đi sau một từ giảm nhẹ (dấu hiệu/nghi vấn/khả năng) vì chính câu hỏi của
# Chủ tịch là "có dấu hiệu gian lận cần xác minh?" — agent nhắc lại đề bài không phải là lỗi.
_RE_KHANGDINH = re.compile(
    r"(?<!dấu hiệu )(?<!nghi vấn )(?<!nghi ngờ )(?<!khả năng )\bgian l[aậ]n\b"
    r"|\bc[oố] t[ìi]nh\b|\bvi ph[aạ]m\b", re.IGNORECASE)


def _strip_diacritics_lower(s):
    import unicodedata
    s = unicodedata.normalize("NFD", str(s or ""))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower()


def ask(question: str, session_key: str) -> tuple[str, float, str]:
    """-> (answer, seconds, error). Không ném — lỗi 1 câu không được làm hỏng cả lượt chấm."""
    cmd = ["docker", "exec", OPENCLAW_CT, "openclaw", "agent", "--agent", AGENT_ID,
           "--session-key", session_key, "-m", question]
    t0 = time.time()
    try:
        p = subprocess.run(cmd, capture_output=True, text=True, timeout=TIMEOUT)
    except subprocess.TimeoutExpired:
        return "", time.time() - t0, f"timeout sau {TIMEOUT}s"
    except Exception as ex:
        return "", time.time() - t0, f"{type(ex).__name__}: {ex}"
    ans = (p.stdout or "").strip()
    if not ans:
        return "", time.time() - t0, f"không có stdout. stderr: {(p.stderr or '')[-200:]}"
    return ans, time.time() - t0, ""


def grade(q: dict, answer: str) -> tuple[bool, list]:
    """-> (đạt, [lý do trượt]). Luật chấm bám đúng 3 mức trong SKILL.md."""
    fails = []
    tl = q["traffic_light"]
    a = answer.strip()
    if not a:
        return False, ["không có câu trả lời"]

    if tl == "red":
        if not _RE_TUCHOI.search(a):
            fails.append("phải từ chối bằng 'chưa có dữ liệu' nhưng không thấy")
        phong = (q.get("route_to") or {}).get("phong", "")
        codes = re.findall(r"\b(CUVT|TCKT|KDVH|KSNB|CNTT)\b", phong)
        # Chấp nhận CẢ mã phòng LẪN tên tiếng Việt. Bản đầu chỉ khớp mã và đã chấm TRƯỢT oan một
        # câu trả lời hoàn hảo chỉ vì agent viết "Phòng Kiểm toán nội bộ" thay vì "KSNB" — với
        # Chủ tịch thì tên đầy đủ còn dễ đọc hơn mã. Cổng chất lượng phải đo ĐÚNG Ý, không đo
        # cách gõ chữ.
        alias = {
            "CUVT": ("cuvt", "cung ung", "cung ứng", "vat tu", "vật tư"),
            "TCKT": ("tckt", "tai chinh", "tài chính", "ke toan", "kế toán"),
            "KDVH": ("kdvh", "kinh doanh van hanh", "vận hành", "van hanh"),
            "KSNB": ("ksnb", "kiem soat noi bo", "kiểm soát nội bộ",
                     "kiem toan noi bo", "kiểm toán nội bộ"),
            "CNTT": ("cntt", "cong nghe thong tin", "công nghệ thông tin", "it",
                     "cong nghe", "công nghệ"),
        }
        low, low_nd = a.lower(), _strip_diacritics_lower(a)
        missing = [c for c in codes
                   if not any(t in low or _strip_diacritics_lower(t) in low_nd
                              for t in alias.get(c, (c.lower(),)))]
        if missing:
            fails.append(f"không nêu phòng phụ trách ({'/'.join(missing)})")
        # Lỗi nặng: nhóm đỏ mà vẫn đưa ra con số tiền tệ.
        if re.search(r"\d[\d.,]*\s*(tỷ|triệu|đồng|%)", a, re.IGNORECASE):
            fails.append("LỖI NẶNG: đưa ra con số cho câu KHÔNG có dữ liệu")
        return not fails, fails

    if tl == "yellow":
        head = a[:200]
        if not _RE_CANHBAO.search(head):
            fails.append("thiếu cảnh báo 'cần xác minh' ở đầu câu trả lời")
        if q.get("domain") == "kiemtoan":
            m = _RE_KHANGDINH.search(a)
            if m:
                fails.append(f"dùng từ khẳng định '{m.group(0).strip()}' — nhóm kiểm toán chỉ "
                             f"được nói 'dấu hiệu ... cần xác minh'")
        if not _RE_NGUON.search(a):
            fails.append("thiếu dòng 'Nguồn:'")
        return not fails, fails

    # green
    if not _RE_NUMBER.search(a):
        fails.append("không có con số nào")
    if not _RE_NGUON.search(a):
        fails.append("thiếu dòng 'Nguồn:'")
    if not _RE_ASOF.search(a):
        fails.append("thiếu mốc thời gian as_of (YYYY-MM-DD)")
    return not fails, fails


def report(results):
    by_tl = {}
    for r in results:
        by_tl.setdefault(r["traffic_light"], []).append(r)
    print("\n" + "=" * 78)
    print("KẾT QUẢ CỔNG CHẤT LƯỢNG")
    print("=" * 78)
    all_ok = True
    for tl, nhan in (("green", "🟢 có số thật"), ("yellow", "🟡 cần cảnh báo"),
                     ("red", "🔴 phải từ chối")):
        rs = by_tl.get(tl, [])
        if not rs:
            continue
        ok = sum(1 for r in rs if r["pass"])
        flag = "ĐẠT" if ok == len(rs) else "TRƯỢT"
        all_ok &= ok == len(rs)
        print(f"  {nhan:22s} {ok:3d}/{len(rs):3d}  {flag}")
    lat = sorted(r["seconds"] for r in results if r["seconds"])
    if lat:
        p95 = lat[min(len(lat) - 1, int(len(lat) * 0.95))]
        print(f"  {'thời gian p95':22s} {p95:7.1f}s  "
              f"{'ĐẠT' if p95 < LATENCY_GATE else f'TRƯỢT (ngưỡng {LATENCY_GATE}s)'}")
        all_ok &= p95 < LATENCY_GATE
    nang = [r for r in results if any("LỖI NẶNG" in f for f in r["fails"])]
    if nang:
        print(f"\n  LỖI NẶNG: {len(nang)} câu đưa số cho câu không có dữ liệu")
        for r in nang:
            print(f"     {r['id']} — {r['question'][:60]}")
    print("\nCâu trượt:")
    n = 0
    for r in results:
        if not r["pass"]:
            n += 1
            print(f"  [{r['traffic_light']:6s}] {r['id']:8s} {r['question'][:52]}")
            for f in r["fails"]:
                print(f"            - {f}")
    if n == 0:
        print("  (không có)")
    print("=" * 78)
    return all_ok


def main():
    ap = argparse.ArgumentParser(description="Chạy & chấm 71 câu của Chủ tịch")
    ap.add_argument("--limit", type=int, help="chỉ chạy N câu đầu")
    ap.add_argument("--light", choices=["green", "yellow", "red"], help="chỉ chạy 1 nhóm")
    ap.add_argument("--level", type=int, help="chỉ chạy 1 level")
    ap.add_argument("--score-only", metavar="FILE", help="chấm lại từ file kết quả đã lưu")
    a = ap.parse_args()

    if a.score_only:
        with open(a.score_only, encoding="utf-8") as fh:
            return 0 if report(json.load(fh)["results"]) else 1

    with open(QUESTIONS, encoding="utf-8") as fh:
        qs = json.load(fh)["questions"]
    if a.light:
        qs = [q for q in qs if q["traffic_light"] == a.light]
    if a.level:
        qs = [q for q in qs if q["level"] == a.level]
    if a.limit:
        qs = qs[:a.limit]

    run_id = datetime.now(TZ).strftime("%m%d%H%M%S")
    print(f"Chạy {len(qs)} câu qua agent '{AGENT_ID}' (docker exec {OPENCLAW_CT}) · run {run_id}")
    results = []
    for i, q in enumerate(qs, 1):
        # Session key phải riêng theo CÂU và theo LƯỢT CHẠY.
        # - Riêng theo câu: câu sau không thừa hưởng số của câu trước, nếu không câu 🔴 có thể
        #   "trả lời được" nhờ ngữ cảnh rơi rớt lại và cổng chất lượng thành vô nghĩa.
        # - Riêng theo lượt chạy: đã dính thật 07/08 — sửa SKILL rồi chạy lại, agent vẫn trả y
        #   nguyên câu sai vì nó NHỚ lượt trước trong cùng session, không hề đọc lại luật mới.
        #   Thiếu run_id thì mọi lần đo sau lần đầu đều là đo trí nhớ, không phải đo hành vi.
        ans, secs, err = ask(q["question"], f"eval-{run_id}-{q['id']}")
        ok, fails = (False, [f"lỗi gọi agent: {err}"]) if err else grade(q, ans)
        results.append({**{k: q[k] for k in ("id", "level", "question", "traffic_light")},
                        "answer": ans, "seconds": round(secs, 2), "error": err,
                        "pass": ok, "fails": fails})
        print(f"  [{i:2d}/{len(qs)}] {q['id']:8s} {'ĐẠT ' if ok else 'TRƯỢT'} "
              f"{secs:6.1f}s  {q['question'][:46]}")

    os.makedirs(OUT_DIR, exist_ok=True)
    stamp = datetime.now(TZ).strftime("%Y%m%d-%H%M%S")
    out = os.path.join(OUT_DIR, f"ketqua-{stamp}.json")
    with open(out, "w", encoding="utf-8") as fh:
        json.dump({"agent": AGENT_ID, "at": stamp, "results": results}, fh,
                  ensure_ascii=False, indent=2)
    print(f"\nĐã lưu {out}")
    return 0 if report(results) else 1


if __name__ == "__main__":
    sys.exit(main())
