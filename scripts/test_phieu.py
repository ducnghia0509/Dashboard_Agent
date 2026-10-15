# -*- coding: utf-8 -*-
"""Test OFFLINE cho MCP `dashboard_phieu` (Agent tổng). Chạy: .venv/bin/python scripts/test_phieu.py
Không đụng API/DB/OpenClaw thật: API nội bộ `/internal/phieu` GIẢ bằng http.server cục bộ (máy trạng thái rút gọn
đúng bảng chuyển của hợp đồng mục A.2), bot con GIẢ bằng thay `phieu._chay_agent`.
Trọng tâm: gọi đúng endpoint + token; giao nền trả về ngay; ≤3 việc song song; một phiếu một việc; JSON đúng ->
PUT khối + chuyển + đăng tin; JSON hỏng 2 lần -> báo lỗi, giữ trạng thái; sua_loi chỉ DA_DUYET; tách log Xe tải
15/10 (13 dòng -> 13 MOI, 0 TRUNG)."""
import csv
import io
import json
import os
import re
import sys
import tempfile
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_HERE, ".."))
tmp = tempfile.mkdtemp()
TOKEN = "tok-test-123"
os.environ["PHIEU_STATE_DIR"] = os.path.join(tmp, "phieu")
os.environ["PHIEU_INTERNAL_TOKEN"] = TOKEN

# ───────────── API giả ─────────────
CHUYEN = {"MOI": {"DANG_CHAN_DOAN", "TRUNG"}, "DANG_CHAN_DOAN": {"CHO_CHOT_NGHIEP_VU", "KHONG_PHAI_LOI", "DA_DUYET"},
          "CHO_CHOT_NGHIEP_VU": {"DA_DUYET", "HOLD"}, "HOLD": {"CHO_CHOT_NGHIEP_VU"}, "DA_DUYET": {"DANG_SUA"},
          "DANG_SUA": {"CHO_MERGE", "CHO_CHOT_NGHIEP_VU"}, "CHO_MERGE": {"CHO_PHAT_HANH"},
          "CHO_PHAT_HANH": {"CHO_NGHIEM_THU"}, "CHO_NGHIEM_THU": {"DONG", "MO_LAI"}, "MO_LAI": {"DANG_CHAN_DOAN"},
          "KHONG_PHAI_LOI": {"DONG"}}
CHI_NGUOI = {"DA_DUYET", "HUY"}
VAI_KHOI = {"chan_doan": {"tim_loi"}, "nghiem_thu": {"nghiem_thu"}, "nguon_log": {"quan_ly", "tiep_nhan"}}
DB, LOG, LOCK = {}, [], threading.Lock()
SO = [0]
EP_409 = []                                   # đặt số lần cần trả 409 cho lần PUT/chuyển kế tiếp


def _row(t):
    return {"ticket_id": t["id"], "channel_id": t["channel_id"], "root_message_id": 900 + int(t["id"][-3:]),
            "trang_thai": t["tt"], "version": t["v"], "so_lan_chuyen": t["nc"], "so_lan_giao": t["ng"], "data": t["data"]}


class H(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, obj):
        b = json.dumps(obj, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        self.wfile.write(b)

    def _body(self):
        n = int(self.headers.get("Content-Length") or 0)
        return json.loads(self.rfile.read(n)) if n else None

    def _xu_ly(self, method):
        body = self._body()
        path, _, q = self.path.partition("?")
        with LOCK:
            LOG.append({"m": method, "p": path, "q": q, "tok": self.headers.get("X-Phieu-Token"), "b": body})
        if self.headers.get("X-Phieu-Token") != TOKEN:
            return self._send(401, {"detail": "sai token"})
        with LOCK:
            if method == "POST" and path == "/internal/phieu":
                SO[0] += 1
                pid = f"PL-20261015-{SO[0]:03d}"
                DB[pid] = {"id": pid, "channel_id": body["channel_id"], "tt": "MOI", "v": 1, "nc": 0, "ng": 0,
                           "data": {"nguon_log": body["nguon_log"]}, "tin": []}
                return self._send(201, _row(DB[pid]))
            if method == "GET" and path == "/internal/phieu":
                ps = dict(x.split("=", 1) for x in q.split("&") if "=" in x)
                ds = [_row(t) for t in DB.values() if (not ps.get("trang_thai") or t["tt"] == ps["trang_thai"])
                      and (not ps.get("channel_id") or str(t["channel_id"]) == ps["channel_id"])]
                return self._send(200, ds)                      # API thật trả MẢNG
            m = re.fullmatch(r"/internal/phieu/(PL-\d{8}-\d{3})(?:/(khoi)/(\w+)|/(chuyen|dang-tin|giao))?", path)
            if not m or m[1] not in DB:
                return self._send(404, {"detail": "không có"})
            t = DB[m[1]]
            if method == "GET" and not m[2] and not m[4]:
                return self._send(200, _row(t))
            if m[4] == "dang-tin":
                t["tin"].append(body["noi_dung"])
                return self._send(200, {"message_id": 5000 + len(t["tin"])})
            if EP_409:
                EP_409.pop()
                t["v"] += 1                      # ai đó vừa ghi chen
                return self._send(409, {"detail": "version lệch"})
            if body.get("version") != t["v"]:
                return self._send(409, {"detail": "version lệch"})
            if m[4] == "giao":                       # bản rút gọn phieu_service.giao (API thật)
                if body["vai"] != "quan_ly":
                    return self._send(403, {"detail": {"error_code": "SAI_VAI", "message": "Chỉ quan_ly giao việc"}})
                if body["giao_cho"] == "sua_loi" and t["tt"] != "DA_DUYET":
                    return self._send(422, {"detail": {"error_code": "CHUA_DUYET", "message": "chưa DA_DUYET"}})
                if t["ng"] >= 6:
                    return self._send(409, {"detail": {"error_code": "VUOT_TRAN_GIAO",
                                                       "message": f"đã giao {t['ng']}/6 lần"}})
                t["ng"], t["v"] = t["ng"] + 1, t["v"] + 1
                return self._send(200, _row(t))
            if m[2] == "khoi":
                if body["vai"] not in VAI_KHOI.get(m[3], set()):
                    return self._send(403, {"detail": f"vai {body['vai']} không được ghi {m[3]}"})
                t["data"][m[3]] = body["gia_tri"]
                t["v"] += 1
                return self._send(200, _row(t))
            if m[4] == "chuyen":
                den = body["den"]
                if den in CHI_NGUOI or den not in CHUYEN.get(t["tt"], set()):
                    return self._send(422, {"detail": {"error_code": "CHUYEN_SAI",
                                                       "message": f"không chuyển {t['tt']}→{den}"}})
                t["tt"], t["v"], t["nc"] = den, t["v"] + 1, t["nc"] + 1
                return self._send(200, _row(t))
        return self._send(405, {"detail": "?"})

    def do_GET(self):
        self._xu_ly("GET")

    def do_POST(self):
        self._xu_ly("POST")

    def do_PUT(self):
        self._xu_ly("PUT")


srv = ThreadingHTTPServer(("127.0.0.1", 0), H)
threading.Thread(target=srv.serve_forever, daemon=True).start()
os.environ["PHIEU_API_URL"] = f"http://127.0.0.1:{srv.server_address[1]}"

from servers.common import phieu as ph  # noqa: E402
from servers.common import phieu_log as plg  # noqa: E402

ok, loi = 0, []


def chk(ten, cond):
    global ok
    if cond:
        ok += 1
    else:
        loi.append(ten)


def raises(fn, *a, **k):
    try:
        fn(*a, **k)
        return None
    except (ValueError, RuntimeError) as e:
        return str(e)


NL = {"nguoi_bao": "Uyen.lt", "man": "htxt1", "chi_tieu": "Số lượng hàng đã nhập kho", "ky": "2026-07"}
CD_DUNG = {"so_dung": {"gia_tri": 20, "don_vi": "xe"}, "so_prod": {"gia_tri": 48, "don_vi": "xe"},
           "loai": "BUILDER", "tang_lech": "builder", "do_tin": "cao",
           "nguyen_nhan": "builder đếm lô theo HĐ ký trong kỳ thay vì ngày hàng về (mapping dòng 58)",
           "de_xuat_sua": "lọc theo Ngày hàng về", "so_phai_ra_sau_sua": [{"ky": "2026-07", "key_json":
                                                                         "nhapKho.soLuong", "gia_tri": 20}],
           "bang_chung": [{"loai": "file_nguon", "mo_ta": "03_LoHang lọc Đã nhận + tháng 7 = 20 dòng"}]}
TRA_DUNG = "Chẩn đoán xong.\n```json\n" + json.dumps({"chan_doan": CD_DUNG}, ensure_ascii=False) + "\n```"


def tao(n=1, **kw):
    return [ph.api_tao({**NL, **kw}, 77)["id"] for _ in range(n)]


# ── 1. endpoint + token ──
LOG.clear()
p1 = ph.api_tao(NL, 77)
chk("tạo: POST /internal/phieu, body có nguon_log + channel_id",
    LOG[-1]["m"] == "POST" and LOG[-1]["p"] == "/internal/phieu" and LOG[-1]["b"] == {"nguon_log": NL, "channel_id": 77})
chk("mọi request mang X-Phieu-Token", all(x["tok"] == TOKEN for x in LOG))
chk("tạo trả dạng phẳng id/trang_thai/version", p1["id"] == "PL-20261015-001" and p1["trang_thai"] == "MOI"
    and p1["version"] == 1 and p1["data"]["id"] == p1["id"])
d = ph.api_doc(p1["id"])
chk("đọc: GET /internal/phieu/{id}", LOG[-1]["m"] == "GET" and LOG[-1]["p"] == f"/internal/phieu/{p1['id']}"
    and d["data"]["nguon_log"] == NL)
LOG.clear()
ph.api_chuyen(p1["id"], "DANG_CHAN_DOAN", "thử")
chk("chuyển: đọc version rồi POST /chuyen {vai, den, ghi_chu, version}",
    [x["m"] for x in LOG] == ["GET", "POST"] and LOG[-1]["p"].endswith("/chuyen")
    and LOG[-1]["b"] == {"vai": "quan_ly", "den": "DANG_CHAN_DOAN", "ghi_chu": "thử", "version": 1})
chk("API từ chối bước sai -> lỗi nổi lên", "422" in (raises(ph.api_chuyen, p1["id"], "DONG", "") or ""))
chk("DA_DUYET qua tool bị API chặn (cổng người)", raises(ph.api_chuyen, p1["id"], "DA_DUYET", "") is not None)
ph.api_liet_ke("DANG_CHAN_DOAN", 77)
chk("liệt kê: GET /internal/phieu?trang_thai=&channel_id=", LOG[-1]["p"] == "/internal/phieu"
    and "trang_thai=DANG_CHAN_DOAN" in LOG[-1]["q"] and "channel_id=77" in LOG[-1]["q"])
ph.api_dang_tin(p1["id"], "xin chào")
chk("đăng tin: POST /dang-tin {noi_dung}", LOG[-1]["p"].endswith("/dang-tin") and LOG[-1]["b"] == {"noi_dung": "xin chào"})
os.environ["PHIEU_INTERNAL_TOKEN"] = "sai"
chk("sai token -> 401 nổi lên", "401" in (raises(ph.api_doc, p1["id"]) or ""))
os.environ["PHIEU_INTERNAL_TOKEN"] = ""
chk("thiếu token -> báo IT điền .env, không gọi", "PHIEU_INTERNAL_TOKEN" in (raises(ph.api_doc, p1["id"]) or ""))
os.environ["PHIEU_INTERNAL_TOKEN"] = TOKEN
chk("thiếu channel_id + chưa đặt PHIEU_CHANNEL_ID -> báo rõ", "PHIEU_CHANNEL_ID" in (raises(ph.api_tao, NL, None) or ""))
os.environ["PHIEU_CHANNEL_ID"] = "55"
_p55 = ph.api_tao(NL, None)
chk("channel_id bỏ trống -> kênh PHIEU_CHANNEL_ID", LOG[-1]["b"]["channel_id"] == 55)
del os.environ["PHIEU_CHANNEL_ID"]
chk("mã phiếu lạ bị chặn trước khi ghép URL", raises(ph.api_doc, "../x") is not None)
chk("nguon_log sai schema bị chặn trước khi gọi", raises(ph.api_tao, {"man": "x"}, 77) is not None)
EP_409.append(1)
ph.api_ghi_khoi(p1["id"], "chan_doan", "tim_loi", CD_DUNG)
chk("409 lệch version -> đọc lại, thử thêm 1 lần, thành công", DB[p1["id"]]["data"].get("chan_doan") == CD_DUNG)
chk("ghi khối sai schema bị chặn tại tool", raises(ph.api_ghi_khoi, p1["id"], "chan_doan", "tim_loi", {"loai": "X"}))

# ── 2. trích + validate ──
x, e = ph.trich_khoi(TRA_DUNG, "chan_doan")
chk("trích khối bọc {'chan_doan': …}", x == CD_DUNG and e is None)
x, e = ph.trich_khoi("```json\n{\"a\":1}\n```\nsửa lại:\n```json\n" + json.dumps(CD_DUNG) + "\n```", "chan_doan")
chk("lấy khối json CUỐI", x == CD_DUNG)
chk("thiếu khối json -> lỗi", ph.trich_khoi("không có gì", "chan_doan")[1] is not None)
chk("thiếu bang_chung -> lỗi schema", "bang_chung" in ph.trich_khoi(
    "```json\n" + json.dumps({k: v for k, v in CD_DUNG.items() if k != "bang_chung"}) + "\n```", "chan_doan")[1])
chk("don_vi sai enum ($defs/so) -> lỗi", ph.trich_khoi("```json\n" + json.dumps(
    {**CD_DUNG, "so_dung": {"gia_tri": 1, "don_vi": "tấn"}}) + "\n```", "chan_doan")[1] is not None)
chk("nghiem_thu hợp lệ", ph.loi_khoi("nghiem_thu", {"ket_qua": "DAT"}) is None)

# ── 3. giao 5 phiếu tim_loi: ≤3 song song, trả về ngay, cùng phiếu không chạy 2 việc ──
dang, dinh, goi = [0], [0], []
cho_tha = threading.Event()


def agent_cham(vai, key, prompt):
    with LOCK:
        dang[0] += 1
        dinh[0] = max(dinh[0], dang[0])
        goi.append((vai, key, prompt))
    cho_tha.wait(10)
    time.sleep(0.05)
    with LOCK:
        dang[0] -= 1
    return TRA_DUNG


ph._chay_agent = agent_cham
ids5 = tao(5)
LOG.clear()
t0 = time.time()
r = ph.giao_viec("tim_loi", ids5, "đo theo mapping dòng 58")
chk("lượt chat trả về NGAY (<2s) khi bot con còn đang chạy", time.time() - t0 < 2 and not cho_tha.is_set())
chk("5 việc được nhận", sum(1 for k in r["ket_qua"] if k.get("viec_id")) == 5)
time.sleep(0.5)
chk("đúng 3 việc chạy cùng lúc (semaphore 3)", dang[0] == 3 and dinh[0] == 3)
r2 = ph.giao_viec("tim_loi", [ids5[0]])
chk("cùng phiếu đang có việc -> từ chối, không chạy việc thứ 2", "tu_choi" in r2["ket_qua"][0]
    and "một phiếu một việc" in r2["ket_qua"][0]["tu_choi"])
vid0 = r["ket_qua"][0]["viec_id"]
chk("phieu_viec đọc được trạng thái việc", ph.doc_viec(vid0)["trang_thai"] in ("cho", "dang_chay"))
chk("phiếu đã chuyển MOI→DANG_CHAN_DOAN trước khi chạy", all(DB[i]["tt"] == "DANG_CHAN_DOAN" for i in ids5))
_g = [x for x in LOG if x["p"].endswith("/giao")]
chk("mỗi phiếu ghi nhận 1 lần giao ở API (POST /giao {vai: quan_ly, giao_cho, version, viec_id}) TRƯỚC khi chạy",
    len(_g) == 5 and all(x["b"]["vai"] == "quan_ly" and x["b"]["giao_cho"] == "tim_loi" and x["b"]["viec_id"]
                         for x in _g) and all(DB[i]["ng"] == 1 for i in ids5))
chk("lỗi API {error_code, message} hiện gọn", "CHUYEN_SAI: không chuyển" in (raises(ph.api_chuyen, ids5[0], "DONG", "") or ""))
cho_tha.set()
ph.cho(10)
chk("không bao giờ quá 3 việc song song", dinh[0] == 3)
chk("session key agent:tim_loi:phieu-<id>-<n>", all(re.fullmatch(r"agent:tim_loi:phieu-PL-\d{8}-\d{3}-1", k)
                                                      for _, k, _ in goi))
chk("prompt chứa chỉ dẫn + JSON phiếu", all("đo theo mapping dòng 58" in p and '"nguon_log"' in p for *_, p in goi))
# ── 4. JSON đúng -> PUT khối + chuyển + đăng tin ──
t = DB[ids5[0]]
chk("JSON đúng -> PUT chan_doan", t["data"].get("chan_doan") == CD_DUNG)
chk("-> chuyển CHO_CHOT_NGHIEP_VU (code suy, không phải LLM)", t["tt"] == "CHO_CHOT_NGHIEP_VU")
chk("-> đăng tóm tắt vào luồng có số đúng/prod + lệnh DUYET",
    t["tin"] and "Số đúng 20 xe" in t["tin"][-1] and "48 xe" in t["tin"][-1]
    and f"DUYET PHAM VI {ids5[0]} khoi=" in t["tin"][-1])
v = ph.doc_viec(vid0)
chk("việc xong ghi kết quả", v["trang_thai"] == "xong" and v["ket_qua"]["chuyen_den"] == "CHO_CHOT_NGHIEP_VU")
chk("file việc nằm ở memory/phieu/viec/<viec_id>.json",
    os.path.exists(os.path.join(tmp, "phieu", "viec", vid0 + ".json")))
chk("khoá phiếu được nhả sau khi xong", not ph.dang_chay())

# loai DUNG -> KHONG_PHAI_LOI
ph._chay_agent = lambda vai, key, p: "```json\n" + json.dumps({**CD_DUNG, "loai": "DUNG", "tang_lech": "khong"}) + "\n```"
[pd] = tao()
ph.giao_viec("tim_loi", [pd])
ph.cho(10)
chk("loai=DUNG -> KHONG_PHAI_LOI + mời DONG", DB[pd]["tt"] == "KHONG_PHAI_LOI" and f"DONG {pd}" in DB[pd]["tin"][-1])

# ── 5. JSON hỏng 2 lần -> báo lỗi, giữ trạng thái ──
so_goi = []


def agent_hong(vai, key, p):
    so_goi.append(p)
    return "Đây là kết quả: ```json\n{\"loai\": \"BUILDER\",}\n```"


ph._chay_agent = agent_hong
[ph1] = tao()
r = ph.giao_viec("tim_loi", [ph1])
ph.cho(10)
chk("JSON hỏng: gọi bot đúng 2 lần (thử lại 1)", len(so_goi) == 2)
chk("lần 2 mang lời nhắc lỗi", "KHÔNG hợp lệ" in so_goi[1])
chk("hỏng 2 lần -> đăng 'bot tim_loi trả kết quả không hợp lệ'",
    any("bot tim_loi trả kết quả không hợp lệ" in m for m in DB[ph1]["tin"]))
chk("hỏng 2 lần -> giữ trạng thái DANG_CHAN_DOAN, không ghi khối", DB[ph1]["tt"] == "DANG_CHAN_DOAN"
    and "chan_doan" not in DB[ph1]["data"])
chk("việc ghi trạng thái loi", ph.doc_viec(r["ket_qua"][0]["viec_id"])["trang_thai"] == "loi")
lan = []
ph._chay_agent = lambda vai, key, p: (lan.append(1), "sai" if len(lan) == 1 else TRA_DUNG)[1]
[ph2] = tao()
ph.giao_viec("tim_loi", [ph2])
ph.cho(10)
chk("hỏng 1 lần rồi đúng -> thành công", DB[ph2]["tt"] == "CHO_CHOT_NGHIEP_VU" and len(lan) == 2)


def agent_loi(vai, key, p):
    raise TimeoutError("agent 'tim_loi' quá 590s")


ph._chay_agent = agent_loi
[ph3] = tao()
ph.giao_viec("tim_loi", [ph3])
ph.cho(10)
chk("bot con timeout -> báo lỗi, giữ trạng thái", DB[ph3]["tt"] == "DANG_CHAN_DOAN"
    and any("590s" in m for m in DB[ph3]["tin"]))

# ── 6. sua_loi chỉ khi DA_DUYET ──
r = ph.giao_viec("sua_loi", [ids5[1]], "sửa builder htxt1")
chk("sua_loi khi chưa DA_DUYET -> từ chối", "DA_DUYET" in r["ket_qua"][0].get("tu_choi", ""))
chk("… và không ghi viec_sua", not os.path.exists(os.path.join(tmp, "phieu", "viec_sua", ids5[1] + ".json")))
DB[ids5[1]]["tt"] = "DA_DUYET"                         # người đã gõ DUYET PHAM VI (API làm)
DB[ids5[1]]["data"]["quyet_dinh"] = {"pham_vi_duoc_duyet": {"khoi": ["HTXT"], "man": ["htxt1"], "ky": ["2026-07"]},
                                     "nguoi_duyet": "66"}
goi_truoc = len(goi)
r = ph.giao_viec("sua_loi", [ids5[1]], "sửa builder htxt1")
f = os.path.join(tmp, "phieu", "viec_sua", ids5[1] + ".json")
chk("sua_loi khi DA_DUYET -> ghi viec_sua/<id>.json", os.path.exists(f) and r["ket_qua"][0].get("file_viec_sua") == f)
vs = json.load(open(f, encoding="utf-8"))
chk("viec_sua có phiếu + chỉ dẫn + phạm vi duyệt", vs["chi_dan"] == "sửa builder htxt1"
    and vs["phieu"]["nguon_log"] == NL and vs["pham_vi_duoc_duyet"]["man"] == ["htxt1"])
chk("sua_loi đăng 'chờ IT nhận việc sửa'", "chờ IT nhận việc sửa (Claude Code bot-sua-loi)" in DB[ids5[1]]["tin"][-1])
chk("sua_loi KHÔNG chạy bot nào, không đổi trạng thái", len(goi) == goi_truoc and DB[ids5[1]]["tt"] == "DA_DUYET")

# ── 7. nghiem_thu ──
DB[ids5[2]]["tt"] = "CHO_NGHIEM_THU"
ph._chay_agent = lambda vai, key, p: "```json\n" + json.dumps({"ket_qua": "KHONG_DAT", "so_prod_sau": [
    {"ky": "2026-07", "key_json": "nhapKho.soLuong", "gia_tri": 48, "khop_so_dung": False}]}) + "\n```"
r = ph.giao_viec("nghiem_thu", [ids5[2]])
ph.cho(10)
chk("nghiem_thu KHONG_DAT -> PUT nghiem_thu + MO_LAI", DB[ids5[2]]["tt"] == "MO_LAI"
    and DB[ids5[2]]["data"]["nghiem_thu"]["ket_qua"] == "KHONG_DAT")
chk("nghiem_thu session key agent:nghiem_thu:phieu-<id>-2", ph.doc_viec(r["ket_qua"][0]["viec_id"])["session_key"]
    == f"agent:nghiem_thu:phieu-{ids5[2]}-2")
DB[ids5[3]]["tt"] = "CHO_NGHIEM_THU"
ph._chay_agent = lambda vai, key, p: "```json\n" + json.dumps({"ket_qua": "DAT", "ngay": "2026-10-15T12:30:00",
                                                              "noi_dung_tra_loi_user": "Đã sửa, T7 = 20 xe",
                                                              "so_prod_sau": [{"ky": "2026-07", "gia_tri": 20,
                                                                               "khop_so_dung": True}]}) + "\n```"
ph.giao_viec("nghiem_thu", [ids5[3]])
ph.cho(10)
chk("nghiem_thu DAT -> giữ CHO_NGHIEM_THU, mời người DONG", DB[ids5[3]]["tt"] == "CHO_NGHIEM_THU"
    and f"DONG {ids5[3]}" in DB[ids5[3]]["tin"][-1])
chk("nghiem_thu khi phiếu chưa CHO_NGHIEM_THU -> từ chối",
    "CHO_NGHIEM_THU" in ph.giao_viec("nghiem_thu", [ids5[4]])["ket_qua"][0].get("tu_choi", ""))
chk("vai lạ bị từ chối", raises(ph.giao_viec, "exec", [ids5[4]]) is not None)

# ── 8. trần 6 lần giao/phiếu ──
ph._chay_agent = lambda vai, key, p: "xong"
[p6] = tao()
for _ in range(6):
    ph.giao_viec("check", [p6], "đếm")
    ph.cho(10)
r = ph.giao_viec("check", [p6], "đếm")
chk("lần giao thứ 7 bị từ chối (trần 6)", "trần 6" in r["ket_qua"][0].get("tu_choi", ""))
os.remove(ph._viec_path(sorted(f[:-5] for f in os.listdir(ph.VIEC_DIR)
                               if json.load(open(os.path.join(ph.VIEC_DIR, f)))["ticket_id"] == p6)[0]))
r = ph.giao_viec("check", [p6], "đếm")
chk("mất file việc cục bộ -> vẫn chặn lần 7 nhờ so_lan_giao của API", "đã giao 6 lần" in r["ket_qua"][0].get("tu_choi", ""))
_doc_that = ph.api_doc
ph.api_doc = lambda pid: {**_doc_that(pid), "so_lan_giao": 0}          # giả tool đọc thiếu bộ đếm
r = ph.giao_viec("check", [p6], "đếm")
ph.api_doc = _doc_that
chk("… và nếu tool kiểm sớm hụt, API /giao chặn (VUOT_TRAN_GIAO), không chạy bot",
    "VUOT_TRAN_GIAO" in r["ket_qua"][0].get("tu_choi", "") and not ph.dang_chay())
chk("check/mapping: đăng nguyên câu trả lời, không đổi trạng thái", DB[p6]["tt"] == "MOI"
    and DB[p6]["tin"][-1].startswith("[check]"))

# ── 9. việc treo khi khởi động lại ──
treo = {"viec_id": "V-20261015000000-abcd", "vai": "tim_loi", "ticket_id": p6, "trang_thai": "dang_chay",
        "pid": -1}
ph._ghi_atomic(ph._viec_path(treo["viec_id"]), treo)
chk("don_viec_treo đánh dấu gian_doan", ph.don_viec_treo() == 1
    and ph.doc_viec(treo["viec_id"])["trang_thai"] == "gian_doan")
chk("viec_id lạ bị chặn", raises(ph.doc_viec, "../../etc/passwd") is not None)

# ── 10. tách log Xe tải 15/10 (13 dòng thật) ──
LOG_TSV = open(os.path.join(_HERE, "du_lieu_test", "log_mau_xetai_so_gia.tsv"), encoding="utf-8").read()
MAN = {"Kinh doanh xe tải - Bảng điểm KPI": "htxt0", "Cung ứng - vận hành xe tải": "htxt1"}
dong = plg.tach_log(LOG_TSV, tim_man=MAN.get, sheet_cua_man=lambda m: "S")
chk("13 dòng log -> 13 dòng", len(dong) == 13)
chk("0 TRUNG (T7 'Giống lỗi 12/13' là kỳ khác)", all(d["trung_voi_thu_tu"] is None for d in dong))
chk("'Giống lỗi N' -> lien_quan về dòng T8 cùng chỉ tiêu", dong[11]["lien_quan_thu_tu"] == [10]
    and dong[12]["lien_quan_thu_tu"] == [11])
n = [d["nguon_log"] for d in dong]
chk("ô trống kế thừa Tên User + Dashboard mẹ", n[1]["nguoi_bao"] == "quyen.dt.005745" and n[12]["nguoi_bao"] == "Uyen.lt"
    and n[5]["man"] == "htxt0" and n[9]["man"] == "htxt1")
chk("' 12,345,678,901 ' -> 12345678901 dong", n[0]["so_bao_cao"] == {"gia_tri": 12345678901, "don_vi": "dong"})
chk("' -   ' -> null", n[1]["so_dashboard"]["gia_tri"] is None)
chk("' (7)' -> -7 (chênh lệch)", n[1]["chenh_lech"]["gia_tri"] == -7)
chk("'1.5 triệu USD' -> 1500000 usd", n[12]["so_bao_cao"]["gia_tri"] == 1500000 and n[12]["so_bao_cao"]["don_vi"] == "usd")
chk("'1.25 triệu USD' -> 1250000 usd", n[12]["so_dashboard"] == {**n[12]["so_dashboard"], "gia_tri": 1250000,
                                                                 "don_vi": "usd"})
chk("'2,500,000' Giá trị nhập kho -> KHÔNG đoán USD: khac + ghi_chu", n[10]["so_bao_cao"]["gia_tri"] == 2500000
    and n[10]["so_bao_cao"]["don_vi"] == "khac" and n[10]["so_bao_cao"].get("ghi_chu"))
chk("kỳ 'T8/2026' -> 2026-08, 'Ngày 06/08/2026' -> 2026-08-06", n[1]["ky"] == "2026-08" and n[2]["ky"] == "2026-08-06")
chk("'dòng 10' -> dong_mapping S!10", n[1]["dong_mapping"] == ["S!10"])
chk("'dòng 22-29' -> 8 dòng", n[3]["dong_mapping"] == [f"S!{i}" for i in range(22, 30)])
chk("dong_excel = thứ tự dòng dán", [x["dong_excel"] for x in n] == list(range(1, 14)))
chk("mọi nguon_log đúng schema", all(ph.loi_khoi("nguon_log", x) is None for x in n))
_khong_td = plg.tach_log(LOG_TSV.split("\n", 1)[1], tim_man=MAN.get, sheet_cua_man=lambda m: "S")          # bỏ dòng tiêu đề
chk("không dòng tiêu đề vẫn đọc theo thứ tự cột", len(_khong_td) == 13
    and [d["nguon_log"] for d in _khong_td] == n)
chk("số kiểu VN '20.111.222.333đ'", plg.doc_so("20.111.222.333đ") == {"gia_tri": 20111222333, "don_vi": "dong"})
chk("'75%' -> pct", plg.doc_so("75%")["don_vi"] == "pct")

LOG.clear()
ph._mk()
kq = ph.tao_tu_log(LOG_TSV, 88, tim_man=MAN.get)
chk("tao_tu_log: 13 phiếu MOI, 0 TRUNG, 0 lỗi", kq["so_phieu_tao"] == 13 and kq["so_trung"] == 0 and kq["so_loi"] == 0)
chk("tao_tu_log gọi 13 POST /internal/phieu", sum(1 for x in LOG if x["m"] == "POST" and x["p"] == "/internal/phieu") == 13)
id12 = kq["dong"][11]["id"]
chk("lien_quan mang MÃ PHIẾU dòng T8", DB[id12]["data"]["nguon_log"]["lien_quan"] == [kq["dong"][9]["id"]])
chk("bảng markdown đủ 13 dòng", kq["bang_markdown"].count("\n| ") == 13)
_buf = io.StringIO()
csv.writer(_buf, dialect="excel-tab", lineterminator="\n").writerow(
    list(csv.reader(io.StringIO(LOG_TSV), dialect="excel-tab"))[2])                # chép lại dòng 2 (SL xe bán T8)
dup = LOG_TSV.rstrip("\n") + "\n" + _buf.getvalue()
kq2 = ph.tao_tu_log(dup, 89, tim_man=MAN.get)
chk("dòng chép lại cùng màn+chỉ tiêu+kỳ -> TRUNG (API chuyển MOI→TRUNG)", kq2["so_trung"] == 1
    and DB[kq2["dong"][13]["id"]]["tt"] == "TRUNG"
    and DB[kq2["dong"][13]["id"]]["data"]["nguon_log"]["trung_voi"] == kq2["dong"][1]["id"])

# ── 11. bảng tổng + cột IT fix ──
bt = ph.bang_tong(channel_id=88)
chk("bảng tổng kênh 88: 13 phiếu, markdown", bt["so_phieu"] == 13 and "| PL-" in bt["bang_markdown"]
    and "chờ ai" in bt["bang_markdown"])
did = kq["dong"][0]["id"]
DB[did]["tt"] = "KHONG_PHAI_LOI"
DB[did]["data"]["chan_doan"] = {**CD_DUNG, "loai": "DUNG", "so_dung": {"gia_tri": 20111222333, "don_vi": "dong"},
                                "nguyen_nhan": "số log là doanh thu trước VAT"}
DB[kq["dong"][9]["id"]].update(tt="DONG")
DB[kq["dong"][9]["id"]]["data"]["nghiem_thu"] = {"ket_qua": "DAT", "ngay": "2026-10-15T12:30:00",
                                                 "noi_dung_tra_loi_user": "T8 = 68 xe"}
c = ph.cot_it_fix(channel_id=88)
rows = list(csv.reader(io.StringIO(c["tsv"]), dialect="excel-tab"))
chk("cột IT fix: 13 dòng, 2 cột, đúng thứ tự log", len(rows) == 13 and all(len(r) == 2 for r in rows)
    and [b["dong"] for b in c["bang"]] == list(range(1, 14)))
chk("dòng 1 ĐÚNG -> 'Dashboard ĐÚNG (20.111.222.333 dong)'", "20.111.222.333" in rows[0][0]
    and rows[0][1] == "Không phải lỗi")
chk("dòng 10 DONG/DAT -> 'Đã fix 15/10/2026'", rows[9][0].startswith("Đã fix 15/10/2026") and rows[9][1] == "Đã xong")
chk("kênh 88 chỉ lấy lô gần nhất", c["lo"] == kq["lo"])

# ── 12. giao hàng loạt ──
dinh[0], dang[0] = 0, 0
cho_tha.clear()
ph._chay_agent = agent_cham
r = ph.giao_hang_loat("tim_loi", "MOI", "đo", channel_id=88)
time.sleep(0.5)
chk("giao hàng loạt MOI -> tim_loi: 11 phiếu MOI kênh 88 được nhận",
    sum(1 for k in r["ket_qua"] if k.get("viec_id")) == 11)
chk("giao hàng loạt vẫn ≤3 song song", dang[0] == 3)
cho_tha.set()
ph.cho(20)
chk("giao hàng loạt xong vẫn chưa từng quá 3", dinh[0] == 3)

# ── 13. MCP server đăng ký đủ tool ──
from servers import phieu_server  # noqa: E402
import asyncio  # noqa: E402
ten = sorted(t.name for t in asyncio.run(phieu_server.mcp.list_tools()))
CAN = ["phieu_bang_tong", "phieu_chuyen", "phieu_cot_it_fix", "phieu_dang_tin", "phieu_doc", "phieu_giao_hang_loat",
       "phieu_giao_viec", "phieu_liet_ke", "phieu_tao", "phieu_tao_tu_log", "phieu_viec"]
chk(f"server có đủ 11 tool (có {ten})", ten == CAN)

srv.shutdown()
print(f"{ok} PASS, {len(loi)} FAIL")
for t in loi:
    print("  FAIL:", t)
sys.exit(1 if loi else 0)
