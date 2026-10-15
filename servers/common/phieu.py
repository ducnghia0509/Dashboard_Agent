# -*- coding: utf-8 -*-
"""KHO PHIẾU LỖI + GIAO VIỆC NỀN cho Agent tổng (`quan_ly`, bot `ai_quan_ly_bot`) — 15/10/2026.

Hợp đồng: `wt/hopdong_dot2_agent_tong_20261015.md` mục B. Schema phiếu: `agents/phieu_loi.schema.json` (v1).

THIẾT KẾ
  1. KHÔNG CHẠM DB. Mọi đọc/ghi phiếu đi qua API nội bộ `/internal/phieu/...` của tc-admin-api (mục A.3) với
     header `X-Phieu-Token` (= env PHIEU_INTERNAL_TOKEN). Máy trạng thái là CODE ở API: tool chỉ ĐỀ NGHỊ chuyển,
     API từ chối bước sai (4xx -> lỗi nổi lên nguyên văn cho Agent tổng). Khoá lạc quan bằng `version`: tool đọc
     version mới nhất ngay trước khi ghi, lệch (409) thì đọc lại và thử thêm đúng một lần.
  2. GIAO VIỆC CHẠY NỀN. Lượt chat của Agent tổng bị cắt ~500-590s; một lượt tim_loi có thể dài hơn. Nên
     `giao_viec` chỉ ghi file việc + thả thread rồi TRẢ VỀ NGAY. Thread: semaphore 3 TOÀN TIẾN TRÌNH (hợp đồng:
     tối đa 3 việc song song toàn hệ thống — chỉ tiến trình MCP này chạy bot con) + khoá theo phiếu (một phiếu một
     việc một lúc; giao trùng bị từ chối ngay, không xếp hàng). Trạng thái việc ghi atomic ra
     `memory/phieu/viec/<viec_id>.json`; `phieu_viec` đọc lại. Tiến trình chết giữa chừng: lần khởi động sau đánh
     dấu các việc treo là `gian_doan` (không tự chạy lại — phiếu vẫn ở trạng thái cũ, Agent tổng giao lại).
  3. GỌI BOT CON đúng như tc-admin-api `bot_service._call_agent_docker` (bản prod 15/10/2026):
     `docker exec <OPENCLAW_CONTAINER> openclaw agent --agent <vai> --session-key agent:<vai>:phieu-<id>-<n> -m <prompt>`,
     timeout `PHIEU_AGENT_TIMEOUT` (mặc định 590s như OPENCLAW_TIMEOUT_SECONDS), stdout rỗng = lỗi. Khoá phiên PHẢI
     mang tiền tố `agent:<id agent>:` — thiếu là OpenClaw rơi về agent mặc định.
  4. KẾT QUẢ BOT CON = MỘT khối ```json của đúng khối vai (tim_loi -> `chan_doan`, nghiem_thu -> `nghiem_thu`),
     validate bằng jsonschema theo `$defs` của schema v1. Hỏng -> hỏi lại MỘT lần trong cùng phiên kèm lỗi; vẫn
     hỏng -> đăng "bot <vai> trả kết quả không hợp lệ", GIỮ trạng thái. Trạng thái đề nghị do CODE suy từ khối (không
     để LLM tự chọn): chan_doan.loai=DUNG -> KHONG_PHAI_LOI; còn lại -> CHO_CHOT_NGHIEP_VU (DA_DUYET chỉ đi qua lệnh
     người `DUYET PHAM VI`). nghiem_thu KHONG_DAT -> MO_LAI; DAT / DAT_MOT_PHAN -> giữ CHO_NGHIEM_THU, mời người
     `DONG PL-…` (đóng phiếu là cổng người).
  5. sua_loi KHÔNG chạy gì (bot sửa lỗi là Claude Code do IT chạy): chỉ khi phiếu DA_DUYET, ghi
     `memory/phieu/viec_sua/<id>.json` (phiếu + chỉ dẫn) và đăng tin chờ IT nhận việc.
  6. Trần: 6 lần giao/phiếu — API giữ (`POST /internal/phieu/{id}/giao` gọi TRƯỚC khi chạy bot con, tăng
     `so_lan_giao`); tool kiểm sớm thêm bằng max(file việc của phiếu, so_lan_giao). 12 lần chuyển do API giữ.
Nhật ký: `memory/phieu/audit.jsonl`.
"""
import csv
import datetime
import io
import json
import os
import re
import secrets
import subprocess
import threading
import urllib.error
import urllib.parse
import urllib.request

import jsonschema

ROOT = os.environ.get("TC_ROOT") or "/home/itadmin/AI_Dashboard_QT"
AGENT_DIR = os.environ.get("TC_AGENT_DIR") or os.path.join(ROOT, "Dashboard_Agent")
_REPO = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
SCHEMA_FILE = os.path.join(_REPO, "agents", "phieu_loi.schema.json")
STATE_DIR = os.environ.get("PHIEU_STATE_DIR") or os.path.join(AGENT_DIR, "memory", "phieu")
VIEC_DIR = os.path.join(STATE_DIR, "viec")
VIEC_SUA_DIR = os.path.join(STATE_DIR, "viec_sua")
AUDIT = os.path.join(STATE_DIR, "audit.jsonl")

VAI_QL = "quan_ly"
VAI_CHAY = ("tim_loi", "nghiem_thu", "check", "mapping")
VAI_GIAO = VAI_CHAY + ("sua_loi",)
KHOI_CUA_VAI = {"tim_loi": "chan_doan", "nghiem_thu": "nghiem_thu"}
SONG_SONG = 3
MAX_GIAO_MOI_PHIEU = 6
MAX_TRA_LOI_LUU = 6000
MAX_PROMPT = 100_000                     # một đối số argv tối đa ~128 KiB (MAX_ARG_STRLEN)
TRANG_THAI_KET = ("DONG", "HUY", "TRUNG")
_RE_ID = re.compile(r"PL-\d{8}-\d{3}")
_RE_VIEC = re.compile(r"V-\d{14}-[0-9a-f]{4}")

_sem = threading.BoundedSemaphore(SONG_SONG)
_khoa = threading.Lock()                 # giữ _dang_lam + đếm lần giao
_dang_lam = {}                           # ticket_id -> viec_id đang chờ/chạy
_threads = {}                            # viec_id -> Thread (test chờ được)


class PhieuLoi(RuntimeError):
    """Lỗi gọi API phiếu — mang mã HTTP để nơi gọi phân biệt 409 (lệch version) / 4xx (API từ chối)."""

    def __init__(self, msg, status=None):
        super().__init__(msg)
        self.status = status


# ───────────────────────── tiện ích ─────────────────────────
def _bay_gio():
    return datetime.datetime.now()


def _iso():
    return _bay_gio().isoformat(timespec="seconds")


def _mk():
    for d in (VIEC_DIR, VIEC_SUA_DIR):
        os.makedirs(d, exist_ok=True)


def _ghi_atomic(path, obj):
    _mk()
    tmp = f"{path}.tmp{os.getpid()}-{threading.get_ident()}"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(obj, fh, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def _doc_json(path):
    try:
        with open(path, encoding="utf-8") as fh:
            x = json.load(fh)
    except (OSError, ValueError):
        return None
    return x if isinstance(x, dict) else None


def _audit(hanh_dong, **kv):
    _mk()
    with _khoa_audit, open(AUDIT, "a", encoding="utf-8") as fh:
        fh.write(json.dumps({"luc": _iso(), "hanh_dong": hanh_dong, **kv}, ensure_ascii=False, default=str) + "\n")


_khoa_audit = threading.Lock()


def _id(x):
    if not _RE_ID.fullmatch(str(x or "")):
        raise ValueError(f"mã phiếu không hợp lệ: {x!r} (dạng PL-YYYYMMDD-NNN)")
    return str(x)


def dinh_dang_so(v):
    """Số đầy đủ kiểu Việt Nam (1.720.000), không làm tròn; None -> '—'."""
    if v is None:
        return "—"
    if isinstance(v, float) and not v.is_integer():
        nguyen, _, le = repr(v).partition(".")
        return f"{int(nguyen):,}".replace(",", ".") + "," + le
    return f"{int(v):,}".replace(",", ".")


def _so(x):
    if not isinstance(x, dict):
        return "—"
    return f"{dinh_dang_so(x.get('gia_tri'))} {x.get('don_vi', '')}".strip()


# ───────────────────────── API nội bộ (mục A.3) ─────────────────────────
def _api_url():
    return (os.environ.get("PHIEU_API_URL") or "http://127.0.0.1:8011").rstrip("/")


def _goi(method, path, body=None, params=None):
    tok = os.environ.get("PHIEU_INTERNAL_TOKEN") or ""
    if not tok:
        raise PhieuLoi("thiếu PHIEU_INTERNAL_TOKEN trong .env của Dashboard_Agent — IT cần điền rồi restart "
                       "dashboard-mcp-phieu")
    url = _api_url() + path
    if params:
        q = {k: v for k, v in params.items() if v not in (None, "")}
        if q:
            url += "?" + urllib.parse.urlencode(q)
    data = json.dumps(body, ensure_ascii=False).encode("utf-8") if body is not None else None
    req = urllib.request.Request(url, data=data, method=method,
                                 headers={"X-Phieu-Token": tok, "Content-Type": "application/json",
                                          "Accept": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=float(os.environ.get("PHIEU_API_TIMEOUT", "30"))) as r:
            raw = r.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        chi_tiet = e.read().decode("utf-8", "replace")[:800]
        try:
            j = json.loads(chi_tiet)
            chi_tiet = j.get("detail", j) if isinstance(j, dict) else j
            if isinstance(chi_tiet, dict) and chi_tiet.get("message"):          # {error_code, message} của API
                chi_tiet = f"{chi_tiet.get('error_code', '')}: {chi_tiet['message']}".lstrip(": ")
        except ValueError:
            pass
        raise PhieuLoi(f"API phiếu từ chối {method} {path}: HTTP {e.code} — {chi_tiet}", e.code) from None
    except (urllib.error.URLError, OSError) as e:
        raise PhieuLoi(f"không gọi được API phiếu {_api_url()}: {e}") from None
    return json.loads(raw) if raw.strip() else {}


def kenh(channel_id=None):
    """channel_id tường minh, hoặc kênh "Xử lý lỗi dashboard" khai ở env PHIEU_CHANNEL_ID (bot quan_ly không thấy
    id kênh trong lượt chat — id chỉ nằm trong session key do API đặt)."""
    c = channel_id if channel_id not in (None, "") else os.environ.get("PHIEU_CHANNEL_ID")
    if c in (None, ""):
        raise ValueError("thiếu channel_id và env PHIEU_CHANNEL_ID chưa đặt (IT điền id kênh in ra từ "
                         "tao_kenh_agent_tong.py)")
    try:
        return int(c)
    except (TypeError, ValueError):
        raise ValueError(f"channel_id không hợp lệ: {c!r}") from None


def chuan(x):
    """Bản ghi API -> dạng phẳng {id, trang_thai, version, channel_id, root_message_id, so_lan_*, data}.
    Nhận cả hai dạng: hàng bảng (`ticket_id`, `data` jsonb không có id/trang_thai) hoặc phiếu theo schema."""
    if not isinstance(x, dict):
        raise PhieuLoi(f"API trả phiếu không phải object: {str(x)[:200]}")
    if isinstance(x.get("phieu"), dict):
        x = x["phieu"]
    data = dict(x["data"]) if isinstance(x.get("data"), dict) else {
        k: v for k, v in x.items() if k in ("nguon_log", "chan_doan", "quyet_dinh", "sua", "phat_hanh", "nghiem_thu")}
    pid = x.get("ticket_id") or x.get("id") or data.get("id")
    tt = x.get("trang_thai") or data.get("trang_thai")
    data.update({"id": pid, "phien_ban": 1, "trang_thai": tt})
    return {"id": pid, "trang_thai": tt, "version": x.get("version"), "channel_id": x.get("channel_id"),
            "root_message_id": x.get("root_message_id"), "so_lan_chuyen": x.get("so_lan_chuyen"),
            "so_lan_giao": x.get("so_lan_giao"), "data": data}


def api_tao(nguon_log, channel_id):
    if not isinstance(nguon_log, dict):
        raise ValueError("nguon_log phải là object theo schema v1")
    loi = loi_khoi("nguon_log", nguon_log)
    if loi:
        raise ValueError(f"nguon_log sai schema: {loi}")
    channel_id = kenh(channel_id)
    p = chuan(_goi("POST", "/internal/phieu", {"nguon_log": nguon_log, "channel_id": channel_id}))
    _audit("tao", id=p["id"], channel_id=channel_id)
    return p


def api_doc(pid):
    return chuan(_goi("GET", f"/internal/phieu/{_id(pid)}"))


def api_liet_ke(trang_thai=None, channel_id=None):
    r = _goi("GET", "/internal/phieu", params={"trang_thai": trang_thai, "channel_id": channel_id})
    ds = r.get("items", r.get("phieu", r.get("data", []))) if isinstance(r, dict) else r
    return [chuan(x) for x in (ds or [])]


def _voi_version(pid, lam):
    """Gọi `lam(version)` với version mới nhất; 409 (lệch version) -> đọc lại, thử thêm đúng một lần."""
    for lan in (1, 2):
        v = api_doc(pid)["version"]
        try:
            return lam(v)
        except PhieuLoi as e:
            if e.status != 409 or lan == 2:
                raise


def api_ghi_khoi(pid, ten, vai, gia_tri):
    loi = loi_khoi(ten, gia_tri)
    if loi:
        raise ValueError(f"khối {ten} sai schema: {loi}")
    r = _voi_version(pid, lambda v: _goi("PUT", f"/internal/phieu/{_id(pid)}/khoi/{ten}",
                                         {"vai": vai, "version": v, "gia_tri": gia_tri}))
    _audit("ghi_khoi", id=pid, khoi=ten, vai=vai)
    return r


def api_chuyen(pid, den, ghi_chu, vai=VAI_QL):
    r = _voi_version(pid, lambda v: _goi("POST", f"/internal/phieu/{_id(pid)}/chuyen",
                                         {"vai": vai, "den": den, "ghi_chu": ghi_chu or "", "version": v}))
    _audit("chuyen", id=pid, den=den, vai=vai)
    return chuan(r) if isinstance(r, dict) and (r.get("ticket_id") or r.get("id") or r.get("phieu")) else r


def api_giao(pid, giao_cho, ghi_chu, viec_id):
    """Ghi nhận MỘT lần giao ở API (`POST /internal/phieu/{id}/giao`, gọi TRƯỚC khi chạy bot con): API giữ trần
    6 lần/phiếu và luật sua_loi chỉ khi DA_DUYET — từ chối thì không chạy gì."""
    r = _voi_version(pid, lambda v: _goi("POST", f"/internal/phieu/{_id(pid)}/giao",
                                         {"vai": VAI_QL, "giao_cho": giao_cho, "version": v,
                                          "ghi_chu": (ghi_chu or "")[:3900], "viec_id": viec_id}))
    _audit("api_giao", id=pid, giao_cho=giao_cho, viec_id=viec_id)
    return r


def api_dang_tin(pid, noi_dung):
    r = _goi("POST", f"/internal/phieu/{_id(pid)}/dang-tin", {"noi_dung": str(noi_dung)[:4000]})
    _audit("dang_tin", id=pid, do_dai=len(str(noi_dung)))
    return r


# ───────────────────────── schema ─────────────────────────
_schema_cache = {}


def _schema():
    if "s" not in _schema_cache:
        with open(SCHEMA_FILE, encoding="utf-8") as fh:
            _schema_cache["s"] = json.load(fh)
    return _schema_cache["s"]


def loi_khoi(ten, gia_tri):
    """Validate một khối (`chan_doan`, `nghiem_thu`, `nguon_log`…) theo schema v1, dùng `$defs` của bản chính.
    Trả chuỗi lỗi (tối đa 5 lỗi, có đường dẫn) hoặc None nếu hợp lệ."""
    s = _schema()
    if ten not in s["properties"]:
        return f"không có khối '{ten}' trong schema v1"
    sub = dict(s["properties"][ten])
    sub["$defs"] = s["$defs"]
    sub["$schema"] = s.get("$schema")
    v = jsonschema.Draft202012Validator(sub)
    loi = sorted(v.iter_errors(gia_tri), key=lambda e: list(e.absolute_path))
    if not loi:
        return None
    return "; ".join(f"{'/'.join(map(str, e.absolute_path)) or '(gốc)'}: {e.message}" for e in loi[:5])


def trich_khoi(tra_loi, ten):
    """Câu trả lời bot con -> (khối, lỗi). Lấy khối ```json CUỐI CÙNG; chấp nhận cả {"<ten>": {...}} bọc ngoài."""
    khoi = re.findall(r"```json\s*\n?(.*?)```", tra_loi or "", re.S | re.I)
    if not khoi:
        return None, "không có khối ```json nào trong câu trả lời"
    try:
        x = json.loads(khoi[-1])
    except ValueError as e:
        return None, f"khối ```json không phải JSON hợp lệ: {e}"
    if isinstance(x, dict) and set(x) == {ten} and isinstance(x[ten], dict):
        x = x[ten]
    if not isinstance(x, dict):
        return None, "khối JSON phải là một object"
    loi = loi_khoi(ten, x)
    return (None, f"sai schema khối {ten}: {loi}") if loi else (x, None)


# ───────────────────────── gọi bot con ─────────────────────────
def chay_agent_docker(vai, session_key, prompt):
    """Bản chép `bot_service._call_agent_docker` (tc-admin-api prod, 15/10/2026). Trả stdout; lỗi -> ném."""
    cont = os.environ.get("OPENCLAW_CONTAINER") or "openclaw-openclaw-gateway-1"
    duoi = session_key.split(":", 2)[-1] if session_key.startswith("agent:") else session_key
    khoa = f"agent:{vai}:{duoi}"
    han = float(os.environ.get("PHIEU_AGENT_TIMEOUT", "590"))
    cmd = ["docker", "exec", cont, "openclaw", "agent", "--agent", vai, "--session-key", khoa, "-m", prompt]
    try:
        p = subprocess.run(cmd, capture_output=True, text=True, timeout=han)
    except subprocess.TimeoutExpired:
        raise TimeoutError(f"agent '{vai}' quá {han:.0f}s") from None
    ans = (p.stdout or "").strip()
    if not ans:
        raise RuntimeError(f"agent '{vai}' trả rỗng (rc={p.returncode}): {(p.stderr or '')[-300:]}")
    return ans


_chay_agent = chay_agent_docker          # test thay bằng hàm giả


_HUONG_DAN_KHOI = {
    "chan_doan": ("Trả lời bằng ĐÚNG MỘT khối ```json là object `chan_doan` theo schema phiếu v1 "
                  "(bắt buộc: so_dung, so_prod, loai, do_tin, nguyen_nhan, bang_chung[≥1]; số là {gia_tri, don_vi}, "
                  "don_vi ∈ dong/usd/xe/cai/pct/khac, KHÔNG làm tròn; loai ∈ DUNG/FILE/MAPPING/SPEC/NAP/BUILDER/FE; "
                  "do_tin ∈ cao/vua/thap; câu hỏi nghiệp vụ cho người ở cau_hoi_nghiep_vu[{ma, cau_hoi, lua_chon, "
                  "mac_dinh_de_xuat}]; so_phai_ra_sau_sua[{ky, key_json, gia_tri}] để nghiệm thu)."),
    "nghiem_thu": ("Trả lời bằng ĐÚNG MỘT khối ```json là object `nghiem_thu` theo schema phiếu v1 "
                   "(bắt buộc ket_qua ∈ DAT/KHONG_DAT/DAT_MOT_PHAN; so_prod_sau[{ky, key_json, gia_tri, khop_so_dung}] "
                   "đo bằng tl_so_dashboard(env=prod) cho TỪNG mục chan_doan.so_phai_ra_sau_sua; do_qua; "
                   "noi_dung_tra_loi_user = câu dán vào cột 'IT fix lỗi')."),
}


def _prompt(vai, phieu, chi_dan):
    khoi = KHOI_CUA_VAI.get(vai)
    pj = json.dumps(phieu["data"], ensure_ascii=False, indent=1)
    if len(pj) > MAX_PROMPT - 4000:
        pj = pj[:MAX_PROMPT - 4000] + "\n…(cắt bớt — gọi lại dữ liệu bằng tool nếu cần)"
    dau = (f"[Việc từ Agent tổng — phiếu {phieu['id']}, trạng thái {phieu['trang_thai']}]\n"
           f"Chỉ dẫn: {chi_dan or '(không có)'}\n\n")
    cuoi = (f"\n\n{_HUONG_DAN_KHOI[khoi]}" if khoi else
            "\n\nTrả lời ngắn bằng tiếng Việt: kết luận, số (lấy từ tool), bằng chứng.")
    return f"{dau}Phiếu (JSON):\n```json\n{pj}\n```{cuoi}"


# ───────────────────────── việc nền ─────────────────────────
def _viec_path(vid):
    if not _RE_VIEC.fullmatch(str(vid or "")):
        raise ValueError(f"viec_id không hợp lệ: {vid!r}")
    return os.path.join(VIEC_DIR, vid + ".json")


def _so_viec_cua_phieu(pid):
    _mk()
    n = 0
    for f in os.listdir(VIEC_DIR):
        if f.endswith(".json"):
            x = _doc_json(os.path.join(VIEC_DIR, f))
            n += bool(x and x.get("ticket_id") == pid)
    return n + (1 if os.path.exists(os.path.join(VIEC_SUA_DIR, pid + ".json")) else 0)


def _luu_viec(v):
    v["cap_nhat"] = _iso()
    _ghi_atomic(_viec_path(v["viec_id"]), v)


def don_viec_treo():
    """Gọi khi tiến trình khởi động: việc `cho`/`dang_chay` của tiến trình khác (đã chết) -> `gian_doan`."""
    _mk()
    n = 0
    for f in os.listdir(VIEC_DIR):
        if not f.endswith(".json"):
            continue
        v = _doc_json(os.path.join(VIEC_DIR, f))
        if v and v.get("trang_thai") in ("cho", "dang_chay") and v.get("pid") != os.getpid():
            v.update(trang_thai="gian_doan", loi="tiến trình MCP khởi động lại khi việc chưa xong — giao lại")
            _luu_viec(v)
            n += 1
    return n


def _de_nghi_den(vai, khoi, phieu):
    """Trạng thái đề nghị suy TẤT ĐỊNH từ khối (không để LLM tự chọn). None = giữ nguyên."""
    if vai == "tim_loi":
        if khoi.get("loai") == "DUNG":
            return "KHONG_PHAI_LOI"
        return "CHO_CHOT_NGHIEP_VU"
    if vai == "nghiem_thu":
        return "MO_LAI" if khoi.get("ket_qua") == "KHONG_DAT" else None
    return None


def _tom_tat(vai, pid, khoi, den, phieu):
    nl = phieu["data"].get("nguon_log") or {}
    if vai == "tim_loi":
        d = [f"[tim_loi] {pid} · {nl.get('man')} · {nl.get('chi_tieu')} · {nl.get('ky')}",
             f"Loại: {khoi.get('loai')} (tầng {khoi.get('tang_lech', '?')}, độ tin {khoi.get('do_tin')}). "
             f"Số đúng {_so(khoi.get('so_dung'))} · PROD đang {_so(khoi.get('so_prod'))}.",
             f"Nguyên nhân: {str(khoi.get('nguyen_nhan'))[:500]}"]
        if khoi.get("de_xuat_sua"):
            d.append(f"Đề xuất sửa: {str(khoi['de_xuat_sua'])[:400]}")
        for q in khoi.get("cau_hoi_nghiep_vu") or []:
            lc = " / ".join(q.get("lua_chon") or [])
            d.append(f"Câu hỏi {q['ma']}: {q['cau_hoi']}" + (f" [{lc}]" if lc else "")
                     + (f" (đề xuất: {q['mac_dinh_de_xuat']})" if q.get("mac_dinh_de_xuat") else ""))
        if den == "CHO_CHOT_NGHIEP_VU":
            ky = nl.get("ky") if re.fullmatch(r"\d{4}-\d{2}(-\d{2})?", str(nl.get("ky"))) else "*"
            d.append(f"Chờ người duyệt: `DUYET PHAM VI {pid} khoi=<khối> man={nl.get('man') or '*'} ky={ky}`"
                     + (f" (trả lời câu hỏi trước: `TRA LOI {pid} <mã>=<lựa chọn>`)"
                        if khoi.get("cau_hoi_nghiep_vu") else ""))
        elif den == "KHONG_PHAI_LOI":
            d.append(f"Dashboard ĐÚNG — người xác nhận đóng: `DONG {pid}`")
        return "\n".join(d)
    if vai == "nghiem_thu":
        d = [f"[nghiem_thu] {pid}: {khoi.get('ket_qua')}"]
        for s in khoi.get("so_prod_sau") or []:
            d.append(f"- {s.get('ky')} {s.get('key_json')}: {dinh_dang_so(s.get('gia_tri'))}"
                     f" {'khớp' if s.get('khop_so_dung') else 'LỆCH'}")
        if khoi.get("ket_qua") == "DAT":
            d.append(f"Đạt — người đóng phiếu: `DONG {pid}`")
        elif den == "MO_LAI":
            d.append("Không đạt — phiếu MỞ LẠI, Agent tổng giao chẩn đoán lại.")
        return "\n".join(d)
    return f"[{vai}] {pid}: xong"


def _lam_viec(v):
    """Thân thread: chờ semaphore -> gọi bot con (≤2 lần) -> PUT khối -> chuyển -> đăng tin. Luôn nhả khoá."""
    pid, vai = v["ticket_id"], v["vai"]
    try:
        with _sem:
            v.update(trang_thai="dang_chay", bat_dau=_iso())
            _luu_viec(v)
            phieu = api_doc(pid)
            ten = KHOI_CUA_VAI.get(vai)
            msg = _prompt(vai, phieu, v.get("chi_dan"))
            khoi, loi, tra_loi = None, None, ""
            for lan in (1, 2):
                v["lan_thu"] = lan
                _luu_viec(v)
                try:
                    tra_loi = _chay_agent(vai, v["session_key"], msg)
                except Exception as e:  # noqa: BLE001 — mọi lỗi gọi bot đều báo, không thử lại
                    loi = f"gọi bot {vai} lỗi: {e}"
                    break
                v["tra_loi_tho"] = tra_loi[-MAX_TRA_LOI_LUU:]
                if not ten:
                    loi = None
                    break
                khoi, loi = trich_khoi(tra_loi, ten)
                if khoi is not None:
                    break
                msg = (f"Câu trả lời trước KHÔNG hợp lệ: {loi}\nGửi lại ĐÚNG MỘT khối ```json `{ten}` hợp lệ "
                       f"(không kèm khối json nào khác). {_HUONG_DAN_KHOI[ten]}")
            if loi:
                bao = (f"bot {vai} trả kết quả không hợp lệ ({loi[:300]}) — giữ trạng thái {phieu['trang_thai']}"
                       if "gọi bot" not in loi else f"{loi[:400]} — giữ trạng thái {phieu['trang_thai']}")
                v.update(trang_thai="loi", loi=bao)
                _dang_tin_an_toan(pid, bao, v)
                return
            if not ten:                                   # check / mapping: không có khối, đăng nguyên ý chính
                v.update(trang_thai="xong", ket_qua={"tra_loi": tra_loi[:1500]})
                _dang_tin_an_toan(pid, f"[{vai}] {pid}:\n{tra_loi[:1500]}", v)
                return
            api_ghi_khoi(pid, ten, vai, khoi)
            den = _de_nghi_den(vai, khoi, phieu)
            if den:
                api_chuyen(pid, den, f"{vai} xong việc {v['viec_id']}", vai=vai)
            v.update(trang_thai="xong", ket_qua={"khoi": ten, "chuyen_den": den,
                                                 "tom_tat": _tom_tat(vai, pid, khoi, den, phieu)})
            _dang_tin_an_toan(pid, v["ket_qua"]["tom_tat"], v)
    except Exception as e:  # noqa: BLE001 — thread nền không được chết câm
        v.update(trang_thai="loi", loi=f"{type(e).__name__}: {e}"[:800])
        _dang_tin_an_toan(pid, f"việc {v['viec_id']} ({vai}) lỗi: {str(e)[:300]}", v)
    finally:
        v["ket_thuc"] = _iso()
        _luu_viec(v)
        _audit("viec_xong", viec_id=v["viec_id"], id=pid, vai=vai, trang_thai=v["trang_thai"])
        with _khoa:
            if _dang_lam.get(pid) == v["viec_id"]:
                del _dang_lam[pid]


def _dang_tin_an_toan(pid, noi_dung, v):
    try:
        api_dang_tin(pid, noi_dung)
        v["da_dang_tin"] = True
    except Exception as e:  # noqa: BLE001
        v["loi_dang_tin"] = str(e)[:300]


def giao_viec(vai, ids, chi_dan=""):
    """Giao `vai` cho từng phiếu trong `ids`. Trả NGAY danh sách {id, viec_id | tu_choi}. Không chờ bot con."""
    if vai not in VAI_GIAO:
        raise ValueError(f"vai không hợp lệ: {vai!r} — chỉ {list(VAI_GIAO)}")
    if isinstance(ids, str):
        ids = [ids]
    ra = []
    for raw in list(dict.fromkeys(ids or [])):
        try:
            pid = _id(raw)
            ra.append(_giao_mot(vai, pid, chi_dan))
        except (ValueError, PhieuLoi) as e:
            ra.append({"id": raw, "tu_choi": str(e)})
    return {"vai": vai, "ket_qua": ra, "dang_chay_toan_he_thong": dang_chay(),
            "luu_y": "việc chạy NỀN — hỏi tiến độ bằng phieu_viec(viec_id) hoặc phieu_bang_tong"}


def _giao_mot(vai, pid, chi_dan):
    with _khoa:
        if pid in _dang_lam:
            return {"id": pid, "tu_choi": f"phiếu đang có việc {_dang_lam[pid]} — một phiếu một việc một lúc"}
        _dang_lam[pid] = "dang_giao"
    try:
        p = api_doc(pid)
        n = max(_so_viec_cua_phieu(pid), int(p.get("so_lan_giao") or 0))
        if n >= MAX_GIAO_MOI_PHIEU:
            return _nha(pid, {"id": pid, "tu_choi": f"phiếu đã giao {n} lần (trần {MAX_GIAO_MOI_PHIEU}) — "
                                                   "báo người xem lại phiếu"})
        tt = p["trang_thai"]
        if tt in TRANG_THAI_KET:
            return _nha(pid, {"id": pid, "tu_choi": f"phiếu ở {tt} — không giao việc"})
        if vai == "sua_loi":
            if tt != "DA_DUYET":
                return _nha(pid, {"id": pid, "tu_choi": f"sua_loi chỉ giao khi DA_DUYET (đang {tt}) — phạm vi phải "
                                                       "được người duyệt bằng `DUYET PHAM VI` trước"})
            api_giao(pid, vai, chi_dan, None)
            return _nha(pid, _giao_sua(api_doc(pid), chi_dan))
        if vai == "tim_loi" and tt not in ("MOI", "MO_LAI", "DANG_CHAN_DOAN"):
            return _nha(pid, {"id": pid, "tu_choi": f"tim_loi chỉ giao khi MOI/MO_LAI/DANG_CHAN_DOAN (đang {tt})"})
        if vai == "nghiem_thu" and tt != "CHO_NGHIEM_THU":
            return _nha(pid, {"id": pid, "tu_choi": f"nghiem_thu chỉ giao khi CHO_NGHIEM_THU (đang {tt})"})
        vid = f"V-{_bay_gio():%Y%m%d%H%M%S}-{secrets.token_hex(2)}"
        api_giao(pid, vai, chi_dan, vid)
        if vai == "tim_loi" and tt in ("MOI", "MO_LAI"):
            api_chuyen(pid, "DANG_CHAN_DOAN", f"giao tim_loi ({vid})", vai=VAI_QL)
        v = {"viec_id": vid, "vai": vai, "ticket_id": pid, "chi_dan": chi_dan or "", "trang_thai": "cho",
             "tao_luc": _iso(), "pid": os.getpid(), "lan_giao": n + 1,
             "session_key": f"agent:{vai}:phieu-{pid}-{n + 1}"}
        _luu_viec(v)
        with _khoa:
            _dang_lam[pid] = vid
        _audit("giao", viec_id=vid, id=pid, vai=vai)
        t = threading.Thread(target=_lam_viec, args=(v,), name=f"phieu-{vid}", daemon=True)
        _threads[vid] = t
        t.start()
        return {"id": pid, "viec_id": vid, "session_key": v["session_key"]}
    except BaseException:
        with _khoa:
            if _dang_lam.get(pid) == "dang_giao":
                del _dang_lam[pid]
        raise


def _nha(pid, kq):
    with _khoa:
        if _dang_lam.get(pid) == "dang_giao":
            del _dang_lam[pid]
    return kq


def _giao_sua(p, chi_dan):
    pid = p["id"]
    if p["trang_thai"] != "DA_DUYET":
        return {"id": pid, "tu_choi": f"sua_loi chỉ giao khi DA_DUYET (đang {p['trang_thai']}) — phạm vi phải "
                                      "được người duyệt bằng `DUYET PHAM VI` trước"}
    duong = os.path.join(VIEC_SUA_DIR, pid + ".json")
    _ghi_atomic(duong, {"ticket_id": pid, "luc": _iso(), "chi_dan": chi_dan or "", "phieu": p["data"],
                        "version": p["version"], "pham_vi_duoc_duyet":
                            (p["data"].get("quyet_dinh") or {}).get("pham_vi_duoc_duyet")})
    _audit("giao_sua", id=pid, file=duong)
    tin = f"{pid}: chờ IT nhận việc sửa (Claude Code bot-sua-loi) — phiếu ghi ở {duong}"
    try:
        api_dang_tin(pid, tin)
    except PhieuLoi as e:
        return {"id": pid, "file_viec_sua": duong, "loi_dang_tin": str(e)}
    return {"id": pid, "file_viec_sua": duong}


def doc_viec(vid):
    v = _doc_json(_viec_path(vid))
    if not v:
        raise ValueError(f"không có việc {vid}")
    v = dict(v)
    v.pop("pid", None)
    if v.get("tra_loi_tho"):
        v["tra_loi_tho"] = v["tra_loi_tho"][-1500:]
    return v


def dang_chay():
    with _khoa:
        return {k: x for k, x in _dang_lam.items() if x != "dang_giao"}


def cho(timeout=None):
    """Chờ mọi thread việc xong (dùng trong test)."""
    for t in list(_threads.values()):
        t.join(timeout)


# ───────────────────────── gộp bước ─────────────────────────
def tao_tu_log(van_ban, channel_id, tim_man=None):
    """Khối log dán từ Excel -> tạo TẤT CẢ phiếu một lần (dòng trùng màn+chỉ tiêu+kỳ -> TRUNG). Trả bảng tóm tắt."""
    from . import phieu_log
    channel_id = kenh(channel_id)
    dong = phieu_log.tach_log(van_ban, tim_man=tim_man)
    lo = f"LO-{_bay_gio():%Y%m%d%H%M%S}"
    id_theo_tt, ra = {}, []
    for d in dong:
        nl = dict(d["nguon_log"])
        nl["file_log"] = f"chat kênh {channel_id} · {lo}"
        nl["lo"] = lo
        if d["trung_voi_thu_tu"] and id_theo_tt.get(d["trung_voi_thu_tu"]):
            nl["trung_voi"] = id_theo_tt[d["trung_voi_thu_tu"]]
        lq = [id_theo_tt[t] for t in d["lien_quan_thu_tu"] if id_theo_tt.get(t)]
        if lq:
            nl["lien_quan"] = lq
        dong_ra = {"dong": d["thu_tu"], "man": nl["man"], "chi_tieu": nl["chi_tieu"], "ky": nl["ky"],
                   "so_bao_cao": _so(nl.get("so_bao_cao")), "so_dashboard": _so(nl.get("so_dashboard")),
                   "canh_bao": d["canh_bao"]}
        try:
            p = api_tao(nl, channel_id)
            id_theo_tt[d["thu_tu"]] = p["id"]
            dong_ra.update(id=p["id"], trang_thai=p["trang_thai"] or "MOI")
            if nl.get("trung_voi"):
                api_chuyen(p["id"], "TRUNG", f"trùng {nl['trung_voi']} (cùng màn + chỉ tiêu + kỳ)", vai=VAI_QL)
                dong_ra.update(trang_thai="TRUNG", trung_voi=nl["trung_voi"])
            if lq:
                dong_ra["lien_quan"] = lq
        except (PhieuLoi, ValueError) as e:
            dong_ra["loi"] = str(e)[:300]
        ra.append(dong_ra)
    md = ["| dòng | phiếu | màn | chỉ tiêu | kỳ | số báo cáo | số dashboard | trạng thái | ghi chú |",
          "|---|---|---|---|---|---|---|---|---|"]
    for r in ra:
        gc = "; ".join(filter(None, [f"trùng {r['trung_voi']}" if r.get("trung_voi") else "",
                                     f"liên quan {', '.join(r['lien_quan'])}" if r.get("lien_quan") else "",
                                     *r["canh_bao"], r.get("loi", "")]))
        md.append(f"| {r['dong']} | {r.get('id', '—')} | {r['man']} | {r['chi_tieu']} | {r['ky']} | "
                  f"{r['so_bao_cao']} | {r['so_dashboard']} | {r.get('trang_thai', 'LỖI')} | {gc} |")
    return {"lo": lo, "so_dong": len(dong), "so_phieu_tao": sum(1 for r in ra if r.get("id")),
            "so_trung": sum(1 for r in ra if r.get("trang_thai") == "TRUNG"),
            "so_loi": sum(1 for r in ra if r.get("loi")), "dong": ra, "bang_markdown": "\n".join(md)}


_CHO_AI = {"MOI": "Agent tổng giao tim_loi", "DANG_CHAN_DOAN": "bot tim_loi", "CHO_CHOT_NGHIEP_VU": "NGƯỜI duyệt/trả lời",
           "HOLD": "NGƯỜI (đang hold)", "DA_DUYET": "Agent tổng giao sua_loi", "DANG_SUA": "IT (bot-sua-loi)",
           "CHO_MERGE": "NGƯỜI merge PR", "CHO_PHAT_HANH": "NGƯỜI restart/deploy", "CHO_NGHIEM_THU": "bot nghiem_thu",
           "MO_LAI": "Agent tổng giao tim_loi", "KHONG_PHAI_LOI": "NGƯỜI xác nhận DONG", "DONG": "—", "HUY": "—",
           "TRUNG": "—"}


def _viec_moi_nhat():
    _mk()
    ra = {}
    for f in sorted(os.listdir(VIEC_DIR)):
        v = _doc_json(os.path.join(VIEC_DIR, f)) if f.endswith(".json") else None
        if v and v.get("ticket_id"):
            ra[v["ticket_id"]] = v
    return ra


def bang_tong(channel_id=None, trang_thai=None):
    ds = api_liet_ke(trang_thai, channel_id)
    viec = _viec_moi_nhat()
    ds.sort(key=lambda p: p["id"] or "")
    md = ["| phiếu | màn | chỉ tiêu | kỳ | trạng thái | việc | chờ ai |", "|---|---|---|---|---|---|---|"]
    dem = {}
    for p in ds:
        nl = p["data"].get("nguon_log") or {}
        v = viec.get(p["id"])
        vt = f"{v['vai']} {v['trang_thai']}" if v else "—"
        cho_ai = _CHO_AI.get(p["trang_thai"], "?")
        if v and v.get("trang_thai") in ("cho", "dang_chay"):
            cho_ai = f"bot {v['vai']}"
        dem[p["trang_thai"]] = dem.get(p["trang_thai"], 0) + 1
        md.append(f"| {p['id']} | {nl.get('man', '?')} | {nl.get('chi_tieu', '?')} | {nl.get('ky', '?')} | "
                  f"{p['trang_thai']} | {vt} | {cho_ai} |")
    tong = " · ".join(f"{k} {n}" for k, n in sorted(dem.items()))
    return {"so_phieu": len(ds), "theo_trang_thai": dem, "dang_chay": dang_chay(),
            "bang_markdown": f"**{len(ds)} phiếu** ({tong or 'trống'})\n" + "\n".join(md)}


def giao_hang_loat(vai, loc_trang_thai, chi_dan="", channel_id=None):
    """Giao `vai` cho MỌI phiếu đang ở `loc_trang_thai` (vd MOI -> tim_loi). Semaphore 3 + trần 6 lần/phiếu giữ
    nguyên: phiếu vượt trần/đang có việc bị từ chối từng cái, phần còn lại xếp hàng chạy dần."""
    if not loc_trang_thai:
        raise ValueError("phải nêu loc_trang_thai (vd MOI)")
    ds = api_liet_ke(loc_trang_thai, channel_id)
    ids = sorted(p["id"] for p in ds if p["trang_thai"] == loc_trang_thai)
    if not ids:
        return {"vai": vai, "ket_qua": [], "luu_y": f"không có phiếu nào ở {loc_trang_thai}"}
    return giao_viec(vai, ids, chi_dan)


def _ngay_vn(s):
    try:
        return datetime.datetime.fromisoformat(str(s).replace("Z", "+00:00")).strftime("%d/%m/%Y")
    except ValueError:
        return str(s or "")[:10]


def _cot_cho_phieu(p):
    d = p["data"]
    tt, cd, nt = p["trang_thai"], d.get("chan_doan") or {}, d.get("nghiem_thu") or {}
    nl = d.get("nguon_log") or {}
    if tt == "DONG" and nt.get("ket_qua") in ("DAT", "DAT_MOT_PHAN"):
        so = "; ".join(f"{s.get('ky')} {dinh_dang_so(s.get('gia_tri'))}" for s in nt.get("so_prod_sau") or [])
        it = nt.get("noi_dung_tra_loi_user") or f"Đã fix, số sau sửa {so}"
        return f"Đã fix {_ngay_vn(nt.get('ngay'))}: {it}".strip(), "Đã xong"
    if tt in ("DONG", "KHONG_PHAI_LOI") and cd.get("loai") == "DUNG":
        return f"Dashboard ĐÚNG ({_so(cd.get('so_dung'))}): {cd.get('nguyen_nhan', '')}", "Không phải lỗi"
    if tt == "TRUNG":
        return f"Trùng {nl.get('trung_voi', '')}", "Trùng"
    if tt == "HOLD":
        q = "; ".join(x.get("cau_hoi", "") for x in cd.get("cau_hoi_nghiep_vu") or [])
        return f"HOLD — chờ chốt nghiệp vụ: {q}".strip(), "Hold"
    if tt == "HUY":
        return "Huỷ", "Huỷ"
    if tt == "CHO_CHOT_NGHIEP_VU":
        q = "; ".join(x.get("cau_hoi", "") for x in cd.get("cau_hoi_nghiep_vu") or [])
        return (f"Đã chẩn đoán ({cd.get('loai')}): {cd.get('nguyen_nhan', '')}" + (f" — cần chốt: {q}" if q else "")
                ), "Chờ duyệt"
    return f"Đang xử lý ({tt})", "Đang xử lý"


def cot_it_fix(channel_id=None, ids=None, lo=None):
    """TSV hai cột 'IT fix lỗi' · 'Trạng thái' theo ĐÚNG thứ tự dòng log gốc (nguon_log.dong_excel) để dán
    ngược vào Excel. Không nêu ids/lo -> lô dán GẦN NHẤT của kênh."""
    if ids:
        ds = [api_doc(i) for i in ([ids] if isinstance(ids, str) else ids)]
    elif channel_id is not None or os.environ.get("PHIEU_CHANNEL_ID"):
        ds = api_liet_ke(None, kenh(channel_id))
        cac_lo = sorted({(p["data"].get("nguon_log") or {}).get("lo") or "" for p in ds} - {""})
        lo = lo or (cac_lo[-1] if cac_lo else None)
        if lo:
            ds = [p for p in ds if (p["data"].get("nguon_log") or {}).get("lo") == lo]
    else:
        raise ValueError("cần channel_id hoặc ids (hoặc đặt env PHIEU_CHANNEL_ID)")
    ds.sort(key=lambda p: ((p["data"].get("nguon_log") or {}).get("dong_excel") or 10 ** 6, p["id"]))
    buf = io.StringIO()
    w = csv.writer(buf, dialect="excel-tab", lineterminator="\n")
    bang = []
    for p in ds:
        it, tt = _cot_cho_phieu(p)
        w.writerow([it, tt])
        nl = p["data"].get("nguon_log") or {}
        bang.append({"dong": nl.get("dong_excel"), "id": p["id"], "it_fix": it, "trang_thai": tt})
    return {"lo": lo, "so_dong": len(ds), "tsv": buf.getvalue(), "bang": bang,
            "huong_dan": "Dán `tsv` vào ô cột 'IT fix lỗi' của dòng log đầu tiên (2 cột: IT fix lỗi, Trạng thái)."}
