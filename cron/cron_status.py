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


def state_from_verify(verify: dict, autofill_ok: bool, today: str) -> str:
    """Suy trạng thái từ kết quả verify — MỘT chỗ duy nhất, để 2 cron không lệch cách xếp.

    Thứ tự xét là xấu-trước-tốt-sau. Đặc biệt phải xét `max_ngay > today` TRƯỚC nhánh 'đủ':
    HTX XTQ/XVP dựng sẵn cột cho cả tháng nên verify trả OK_CO_NGAY_HOM_QUA kể cả khi kế toán
    chưa nhập gì — lấy OK đó làm 'đủ' là báo xanh oan.

    KHÔNG xét việc file có về trong lượt này hay không: file tháng bị ép kéo lại mỗi ngày nên gần
    như luôn "về", còn bằng chứng kế toán đã cập nhật chỉ nằm ở ngày mới nhất trong DB. Cờ `arrived`
    vẫn được ghi vào bản ghi để tra cứu, chỉ là không dùng để xếp loại.
    """
    if not autofill_ok:
        return STATE_LOI_NAP
    code = (verify or {}).get("code") or ""
    max_ngay = (verify or {}).get("max_ngay") or ""
    if code == "THIEU_NGAY_HOM_QUA":
        return STATE_CHAM
    if code == "OK_CO_NGAY_HOM_QUA":
        return STATE_KHONG_XAC_NHAN if (max_ngay and max_ngay > today) else STATE_DU
    return STATE_LOI_NAP        # KHONG_CO_DONG_NAO, hoặc verify không đọc được kết quả


class StatusWriter:
    """Thu trạng thái trong lúc cron chạy rồi ghi 1 lần ở cuối.

    finish() PHẢI được gọi ở MỌI đường ra của main(), kể cả nhánh lỗi — artifact thiếu đồng nghĩa
    với "cron không chạy" theo cách agent hiểu, nên im lặng thoát sẽ thành báo động giả.
    """

    def __init__(self, job: str, env: str, json_path: str, jsonl_path: str,
                 schedule_vn: str, expected: list, names: dict = None):
        self.json_path = json_path
        self.jsonl_path = jsonl_path
        self.run = {
            "job": job,
            "env": env,
            "schedule_vn": schedule_vn,
            "run_at": datetime.now(VN).strftime("%Y-%m-%d %H:%M:%S"),
            "run_date": datetime.now(VN).strftime("%Y-%m-%d"),
            "today": datetime.now(VN).strftime("%Y-%m-%d"),
            "ngay_can": (datetime.now(VN) - timedelta(days=1)).strftime("%Y-%m-%d"),
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
