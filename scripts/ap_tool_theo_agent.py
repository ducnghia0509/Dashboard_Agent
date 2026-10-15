# -*- coding: utf-8 -*-
"""ÁP GIỚI HẠN TOOL THEO AGENT vào openclaw.json (đợt 1 an toàn bot, 15/10/2026) — NGƯỜI DÙNG tự chạy.

Vì sao: 3 agent `check`, `keo_file`, `mapping` không có khối `tools` riêng → hưởng profile toàn cục `coding`
(exec/write/web…) và MỌI MCP server, kể cả `dashboard_ingest__import_execute` (ghi DB). Nội dung cần áp nằm ở
`deploy/openclaw_tools_theo_agent.json` (khoá bắt đầu bằng '_' là ghi chú, bị bỏ khi áp).

openclaw.json thuộc user `node` của container (mode 600, itadmin trên host KHÔNG đọc được) → mọi thao tác đi
qua `docker exec <container>`; file mới ghi bằng user node nên giữ nguyên chủ/quyền.

Chạy:
  .venv/bin/python scripts/ap_tool_theo_agent.py            # CHẠY THỬ: in diff (đã che bí mật) + dry-run OpenClaw
  .venv/bin/python scripts/ap_tool_theo_agent.py --ap       # sao lưu có dấu thời gian -> ghi -> validate lại

Các bước khi --ap: (1) đọc cấu hình hiện tại; (2) dừng nếu thiếu agent nào trong 3 agent; (3) THAY TRỌN
`agents.entries.<id>.tools` của 3 agent; (4) `openclaw config patch --dry-run` (schema chính chủ) phải ok;
(5) sao lưu `openclaw.json.bak-tool-theo-agent-<YYYYmmdd-HHMMSS>` (cp -p trong container); (6) ghi atomic (tmp +
mv); (7) `openclaw config validate`, hỏng thì TỰ chép bản sao lưu về. KHÔNG restart gateway — nhưng gateway có
theo dõi file cấu hình, thay đổi có thể có hiệu lực ngay ở lượt chat kế tiếp.
Hoàn tác: lệnh `docker exec … cp -p <bản sao lưu> <openclaw.json>` được in ra sau khi ghi.
In diff KHÔNG lộ token: giá trị dưới khoá chứa token/key/secret/pass/auth/cookie và mọi chuỗi giống khoá ngẫu
nhiên (≥24 ký tự chữ-số liền) đều thành '<ẨN>'.
"""
import argparse
import copy
import datetime
import difflib
import json
import os
import re
import subprocess
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
PATCH_FILE = os.path.join(_HERE, "..", "deploy", "openclaw_tools_theo_agent.json")
CONTAINER = "openclaw-openclaw-gateway-1"
CFG = "/home/node/.openclaw/openclaw.json"
_KHOA_NHAY_CAM = re.compile(r"token|key|secret|pass|auth|cookie|credential", re.I)
_CHUOI_BI_MAT = re.compile(r"[A-Za-z0-9_\-+/=]{24,}")
_KHOA_TEN_TOOL = {"allow", "deny", "alsoAllow"}


def _dx(args, input_=None, check=True):
    r = subprocess.run(["docker", "exec", "-i", CONTAINER, *args], input=input_, capture_output=True, text=True,
                       timeout=120)
    if check and r.returncode != 0:
        sys.exit(f"DỪNG: docker exec {' '.join(args[:3])}… lỗi rc={r.returncode}: {(r.stderr or r.stdout)[-400:]}")
    return r


def _che(x, khoa=""):
    if isinstance(x, dict):
        return {k: _che(v, k) for k, v in x.items()}
    if isinstance(x, list):
        if khoa in _KHOA_TEN_TOOL and all(isinstance(v, str) for v in x):
            return x                                          # tên tool (dashboard_check__check_…) — không phải bí mật
        return [_che(v, khoa) for v in x]
    if isinstance(x, str) and (_KHOA_NHAY_CAM.search(khoa) or _CHUOI_BI_MAT.search(x)):
        return "<ẨN>"
    if _KHOA_NHAY_CAM.search(khoa) and not isinstance(x, (bool, type(None))):
        return "<ẨN>"
    return x


def _patch():
    with open(PATCH_FILE, encoding="utf-8") as fh:
        d = json.load(fh)
    p = {k: v for k, v in d.items() if not k.startswith("_")}
    ent = p.get("agents", {}).get("entries", {})
    if set(p) != {"agents"} or set(p["agents"]) != {"entries"} or not ent \
            or any(set(v) != {"tools"} for v in ent.values()):
        sys.exit("DỪNG: file patch chỉ được chứa agents.entries.<id>.tools")
    return p


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ap", action="store_true", help="ghi thật (mặc định chỉ chạy thử)")
    a = ap.parse_args()

    patch = _patch()
    ids = sorted(patch["agents"]["entries"])
    cu = json.loads(_dx(["cat", CFG]).stdout)
    thieu = [i for i in ids if i not in cu.get("agents", {}).get("entries", {})]
    if thieu:
        sys.exit(f"DỪNG: openclaw.json không có agent {thieu} — kiểm lại tên agent trước khi áp")
    moi = copy.deepcopy(cu)
    for i in ids:
        cu_tools = cu["agents"]["entries"][i].get("tools")
        if cu_tools:
            print(f"  lưu ý: agent '{i}' đang có tools={json.dumps(_che(cu_tools), ensure_ascii=False)} -> bị THAY")
        moi["agents"]["entries"][i]["tools"] = copy.deepcopy(patch["agents"]["entries"][i]["tools"])
    noi_dung = json.dumps(moi, ensure_ascii=False, indent=2) + "\n"
    json.loads(noi_dung)                                          # JSON hợp lệ

    diff = list(difflib.unified_diff(json.dumps(_che(cu), ensure_ascii=False, indent=2).splitlines(),
                                     json.dumps(_che(moi), ensure_ascii=False, indent=2).splitlines(),
                                     "openclaw.json (hiện tại)", "openclaw.json (sau khi áp)", lineterm="", n=2))
    print("\n".join(diff) if diff else "(không đổi gì — cấu hình đã đúng)")
    if not diff:
        return

    rp = [x for i in ids for x in ("--replace-path", f"agents.entries.{i}.tools")]
    r = _dx(["openclaw", "config", "patch", "--stdin", "--dry-run", "--json", *rp], json.dumps(patch), check=False)
    try:
        kq = json.loads(r.stdout[r.stdout.index("{"):])
    except ValueError:
        sys.exit(f"DỪNG: dry-run OpenClaw không trả JSON: {(r.stdout or r.stderr)[-400:]}")
    if not kq.get("ok"):
        sys.exit(f"DỪNG: OpenClaw từ chối cấu hình: {json.dumps(kq.get('errors'), ensure_ascii=False)[:800]}")
    print("dry-run OpenClaw: ok (schema hợp lệ)")
    if not a.ap:
        print("CHẠY THỬ — chưa ghi. Thêm --ap để áp.")
        return

    bak = f"{CFG}.bak-tool-theo-agent-{datetime.datetime.now():%Y%m%d-%H%M%S}"
    _dx(["cp", "-p", CFG, bak])
    _dx(["sh", "-c", f'cat > "{CFG}.tmp-tool-theo-agent" && chmod 600 "{CFG}.tmp-tool-theo-agent" '
                     f'&& mv "{CFG}.tmp-tool-theo-agent" "{CFG}"'], noi_dung)
    v = _dx(["openclaw", "config", "validate"], check=False)
    if v.returncode != 0:
        _dx(["cp", "-p", bak, CFG])
        sys.exit(f"DỪNG: validate sau khi ghi hỏng -> ĐÃ chép lại bản sao lưu. {(v.stderr or v.stdout)[-400:]}")
    sau = json.loads(_dx(["cat", CFG]).stdout)
    for i in ids:
        if sau["agents"]["entries"][i].get("tools") != patch["agents"]["entries"][i]["tools"]:
            sys.exit(f"DỪNG: đọc lại thấy tools của '{i}' không như mong đợi — kiểm tay (bản sao lưu: {bak})")
    print(f"ĐÃ ÁP cho {ids}. Sao lưu: {bak}")
    print(f"Hoàn tác: docker exec {CONTAINER} cp -p {bak} {CFG}")
    print("Không restart gateway. Kiểm: nhắn thử từng bot một câu; bot nào báo 'không còn tool' -> hoàn tác.")


if __name__ == "__main__":
    main()
