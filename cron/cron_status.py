# -*- coding: utf-8 -*-
"""ARTIFACT TRẠNG THÁI cho 2 job kéo tự động — cron_thuchi_daily.py / cron_hqkdngay_daily.py.

VÌ SAO CÓ FILE NÀY: agent giám sát (`update_file_bao_cao_ngay` trên openclaw) trước đây phải bóc
trạng thái từ log văn xuôi bằng awk/grep. Hai lỗi thật đã xảy ra vì thế: `awk` trong container
openclaw KHÔNG hỗ trợ interval `/={20,}/` nên lệnh cắt "khối log cuối" im lặng trả về CẢ FILE (agent
xếp loại theo lượt của mấy ngày trước), và agent trong session cron đọc log sơ sài hơn session tương
tác nên bỏ sót đơn vị. Log văn xuôi là để NGƯỜI đọc; artifact JSON này là để MÁY đọc.

Agent trong container openclaw KHÔNG nối được Postgres (5434/5435 chỉ mở trên loopback của host) nên
không thể tự truy vấn số ngày/số liệu. Cron chạy trên host thì có sẵn kết nối đó — vì vậy chỗ đúng để
kết luận trạng thái là Ở ĐÂY, không phải ở agent.

MẪU SỐ LÀ DANH SÁCH KỲ VỌNG, KHÔNG PHẢI SỐ FILE TÌM THẤY. Cron vẫn log "nạp thành công 11/11" trong
khi phải có 12 đơn vị — đơn vị vắng mặt hoàn toàn khỏi metadata (ca DUAN: `month=null` do tên file
`...D.20268...` không pad số 0) không để lại dấu vết nào trong log. Artifact luôn ghi ĐỦ bản ghi cho
mọi mục kỳ vọng, đơn vị không thấy gì thì state='chua_co_file'.

Ghi 2 nơi:
  logs/status_<job><suffix>.json    — trạng thái lượt GẦN NHẤT (agent đọc file này)
  logs/status_<job><suffix>.jsonl   — lịch sử, mỗi lượt 1 dòng (để biết "đỏ mấy ngày liên tiếp")

Ghi nguyên tử (tmp + os.replace) để agent không bao giờ đọc được file ghi dở.
"""
import json
import os
from datetime import datetime, timedelta, timezone

VN = timezone(timedelta(hours=7))

# Trạng thái của MỘT mục kỳ vọng. Giữ nguyên chuỗi slug — AGENTS.md của agent map trực tiếp sang
# 4 mục trong tin gửi lãnh đạo, đổi tên slug là hỏng tin.
STATE_DU = "du"                       # file mới về + đã có số liệu ngày cần  -> "Đã update file báo cáo"
STATE_CHAM = "cham"                   # có số liệu nhưng thiếu ngày cần        -> "Đã lên số liệu"
STATE_KHONG_XAC_NHAN = "khong_xac_nhan"  # có ngày cần nhưng file dựng sẵn cả tháng (max_ngay > hôm nay)
STATE_LOI_NAP = "loi_nap"             # kéo được nhưng không ra dòng nào       -> "Có file, không có dữ liệu"
STATE_CHUA_CO_FILE = "chua_co_file"    # không có ở nguồn / xin mà không về    -> "Chưa có file báo cáo"
STATES = (STATE_DU, STATE_CHAM, STATE_KHONG_XAC_NHAN, STATE_LOI_NAP, STATE_CHUA_CO_FILE)

# Trạng thái của CẢ LƯỢT chạy.
RUN_OK = "ok"
RUN_TAT = "bi_tat"              # cờ .disabled do UI đặt
RUN_DUNG_SOM = "dung_som"       # receiver chết / metadata lỗi / không có file nào khớp


def ngay_can(now) -> str:
    """Ngày số liệu mà lượt chạy hôm nay đòi hỏi = HÔM QUA, mọi ngày trong tuần như nhau.

    KHÔNG trừ Chủ nhật (user chốt 2026-08-17). Bản 17/08 từng lùi ngày cần khi ngày nhập rơi vào
    Chủ nhật, với lý lẽ "Chủ nhật không ai nhập nên đừng báo chậm oan" — user bác: vẫn muốn thông
    báo từng ngày như bình thường, Chủ nhật chưa nhập thì cứ hiện đúng là chưa nhập. Bản tin phản
    ánh thực tế, không tự bào chữa thay cho nguồn.

    Giữ hàm này (thay vì tính tại chỗ) để 2 job và trường `ngay_can` trong artifact dùng chung một
    định nghĩa — đổi quy ước sau này chỉ sửa ở đây.
    """
    return (now - timedelta(days=1)).strftime("%Y-%m-%d")


def state_from_verify(verify: dict, autofill_ok: bool, today: str, van_tay_cu: str = None,
                      doi_luc_cu: str = None) -> tuple:
    """Suy trạng thái từ kết quả verify — MỘT chỗ duy nhất, để 2 cron không lệch cách xếp.

    BẰNG CHỨNG "KẾ TOÁN ĐÃ NHẬP" LÀ VÂN TAY ĐỔI, không phải "có dòng cho ngày hôm qua". Kế toán
    nhập số vào CHÍNH file tháng đó (giống bên dòng tiền), mà nhiều file dựng sẵn cột cho cả tháng:
    XDV có 98 dòng mỗi ngày 12→15/08 với tổng giống hệt 1.655; HTX Xanh VP có số tới tận 20/08 —
    ngày tương lai. Với những file này "có dòng ngày hôm qua" luôn đúng kể cả khi không ai đụng vào,
    nên trước đây chúng nằm mãi ở `khong_xac_nhan`, tức bản tin không nói được gì suốt nhiều ngày.

    So VÂN TAY (số dòng + tổng |amount| của cả kỳ) với lượt trước thì bắt được đúng thứ cần biết:
    file có ĐỔI hay không. Đổi = ai đó đã nhập/sửa. Không đổi = y nguyên hôm qua, dù có bao nhiêu
    dòng đi nữa.

    Lượt ĐẦU TIÊN chưa có vân tay cũ (`van_tay_cu=None`) thì không kết luận được — lùi về cách cũ
    (`max_ngay > today` -> `khong_xac_nhan`) thay vì đoán bừa. Từ lượt thứ hai trở đi mới đủ dữ kiện.

    So với NGÀY ĐỔI GẦN NHẤT chứ không phải "đổi ngay trong lượt này": job chạy 2 lượt cùng ngày
    (hoặc ai đó bấm "Chạy ngay" trên UI) thì lượt sau thấy vân tay y hệt lượt trước — nếu lấy đó
    làm "chưa cập nhật" thì cả bảng chuyển sang chậm oan, dù sáng nay kế toán đã nhập thật. Vì vậy
    ghi lại `doi_luc` = ngày vân tay đổi lần cuối, và "đủ" nghĩa là đã đổi TRONG HÔM NAY.

    Trả (state, doi_luc) để bên gọi ghi `doi_luc` vào artifact cho lượt sau dùng tiếp.
    """
    if not autofill_ok:
        return STATE_LOI_NAP, doi_luc_cu
    code = (verify or {}).get("code") or ""
    max_ngay = (verify or {}).get("max_ngay") or ""
    van_tay = (verify or {}).get("van_tay")
    doi_luc = today if (van_tay and van_tay_cu and van_tay != van_tay_cu) else doi_luc_cu
    if code == "THIEU_NGAY_HOM_QUA":
        return STATE_CHAM, doi_luc
    if code == "OK_CO_NGAY_HOM_QUA":
        if van_tay and van_tay_cu:
            return (STATE_DU if doi_luc == today else STATE_CHAM), doi_luc
        return (STATE_KHONG_XAC_NHAN if (max_ngay and max_ngay > today) else STATE_DU), doi_luc
    return STATE_LOI_NAP, doi_luc   # KHONG_CO_DONG_NAO, hoặc verify không đọc được kết quả


class StatusWriter:
    """Thu trạng thái trong lúc cron chạy rồi ghi 1 lần ở cuối.

    finish() PHẢI được gọi ở MỌI đường ra của main(), kể cả nhánh lỗi — artifact thiếu đồng nghĩa
    với "cron không chạy" theo cách agent hiểu, nên im lặng thoát sẽ thành báo động giả.
    """

    def __init__(self, job: str, env: str, json_path: str, jsonl_path: str,
                 schedule_vn: str, expected: list, names: dict = None):
        self.json_path = json_path
        self.jsonl_path = jsonl_path
        # Vân tay của LƯỢT TRƯỚC, đọc từ chính artifact cũ trước khi ghi đè. Đây là toàn bộ "trí nhớ"
        # của cơ chế phát hiện file có đổi hay không — mất file này thì lượt kế tiếp không kết luận
        # được (rơi về nhánh không-xác-nhận), không phải kết luận sai.
        self.van_tay_cu, self.doi_luc_cu = {}, {}
        try:
            with open(json_path, encoding="utf-8") as fh:
                for r in (json.load(fh).get("records") or []):
                    if r.get("van_tay"):
                        self.van_tay_cu[r["key"]] = r["van_tay"]
                    if r.get("doi_luc"):
                        self.doi_luc_cu[r["key"]] = r["doi_luc"]
        except (OSError, json.JSONDecodeError, AttributeError):
            pass
        self.run = {
            "job": job,
            "env": env,
            "schedule_vn": schedule_vn,
            "run_at": datetime.now(VN).strftime("%Y-%m-%d %H:%M:%S"),
            "run_date": datetime.now(VN).strftime("%Y-%m-%d"),
            "today": datetime.now(VN).strftime("%Y-%m-%d"),
            "ngay_can": ngay_can(datetime.now(VN)),
            "expected_count": len(expected),
        }
        # Tạo sẵn bản ghi cho MỌI mục kỳ vọng ở state xấu nhất; các bước sau nâng dần lên.
        # `ten` là tên tiếng Việt để agent dùng NGUYÊN VĂN trong tin gửi lãnh đạo — dịch mã thư mục
        # (DUAN -> "Dự án") ở đây, một chỗ, thay vì để agent tự map và có ngày map sai.
        self.records = {
            key: {"key": key, "ten": (names or {}).get(key, key), "state": STATE_CHUA_CO_FILE,
                  "ly_do": "không thấy trong danh sách nguồn của lượt này"}
            for key in expected
        }

    def set_run(self, **kv):
        self.run.update(kv)

    def record(self, key: str, **kv):
        """Cập nhật bản ghi. Key ngoài danh sách kỳ vọng vẫn được ghi (để không mất dữ liệu) kèm
        cờ `ngoai_ke_hoach` — nhìn thấy cờ này là danh sách kỳ vọng đã lệch với thực tế."""
        rec = self.records.setdefault(
            key, {"key": key, "ten": key, "state": STATE_CHUA_CO_FILE, "ngoai_ke_hoach": True})
        rec.update(kv)

    def finish(self, cron_status: str = RUN_OK, note: str = None) -> dict:
        recs = sorted(self.records.values(), key=lambda r: r["key"])
        # `summary` CHỈ đếm bản ghi của kỳ chính (mặc định mọi bản ghi là kỳ chính). Job dòng tiền
        # kéo cả tháng trước — kỳ đó đã chốt nên luôn "đủ"; đếm nó vào summary thì mẫu số phồng lên
        # và tin gửi lãnh đạo báo "đủ 2/2" trong khi tháng này chưa có số liệu.
        summary = {s: sum(1 for r in recs if r.get("state") == s and r.get("ky_chinh", True))
                   for s in STATES}
        # Mẫu số ĐỂ ĐỐI CHIẾU với summary phải cùng phạm vi với summary (chỉ kỳ chính), không thì
        # bên tiêu thụ báo "lệch phép kiểm" oan: job dòng tiền có expected_count=2 (2 kỳ) nhưng
        # summary chỉ đếm 1 kỳ chính.
        expected_chinh = sum(1 for r in recs if r.get("ky_chinh", True))
        payload = {
            **self.run,
            "cron_status": cron_status,
            **({"note": note} if note else {}),
            "expected_chinh_count": expected_chinh,
            "summary": summary,
            "records": recs,
        }
        self._write(payload)
        return payload

    def _write(self, payload: dict):
        os.makedirs(os.path.dirname(self.json_path), exist_ok=True)
        tmp = self.json_path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False, indent=2)
        os.replace(tmp, self.json_path)          # nguyên tử: agent không đọc được bản ghi dở
        with open(self.jsonl_path, "a", encoding="utf-8") as fh:
            fh.write(json.dumps(payload, ensure_ascii=False) + "\n")
