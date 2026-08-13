# -*- coding: utf-8 -*-
"""PHÂN LOẠI CÂU HỎI — tách ý + khớp loại + đọc sẵn tham số (kỳ, đơn vị).

Khớp TẤT ĐỊNH bằng từ khoá trong Python, KHÔNG nhờ model tự phân loại: model nhận về công thức
rồi thi hành, thay vì vừa đoán loại vừa đoán cách làm.

Toàn bộ nội dung nghiệp vụ nằm ở `loai_cau_hoi.json` — thêm nguồn báo cáo mới thì sửa JSON, không
đụng file này.

TÁCH Ý: lỗi rơi ý là kiểu hỏng ÂM THẦM (trả lời ý 1 rất đẹp rồi quên ý 3, người đọc không biết là
đã mất một ý) và nặng nhất đúng lúc lịch sử bị nén. Vì vậy ý nào không phân loại được VẪN phải xuất
hiện trong danh sách với loai=None — thà báo "không hiểu ý này" còn hơn im lặng bỏ qua.
"""
import json
import os
import re
import unicodedata

_HERE = os.path.dirname(os.path.abspath(__file__))
_AGENT_ROOT = os.path.normpath(os.path.join(_HERE, "..", ".."))
SO_TAY = os.path.join(_AGENT_ROOT, "loai_cau_hoi.json")

_cache = None


def _dac(s) -> str:
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d").replace("Đ", "D").strip().lower()


def so_tay() -> dict:
    global _cache
    if _cache is None:
        with open(SO_TAY, encoding="utf-8") as fh:
            _cache = json.load(fh)
        for lo in _cache["loai"]:
            lo["_nhan_dien_dac"] = [_dac(t) for t in lo.get("nhan_dien", [])]
    return _cache


# --- Tách ý ---------------------------------------------------------------------------------
# Dấu phân cách MẠNH luôn tách. Liên từ chỉ tách khi CẢ HAI vế đều nhận ra được loại — nếu không,
# "doanh thu và chi phí tháng 7" sẽ bị bẻ thành 2 ý cụt.
_TACH_MANH = re.compile(r"[;\n]+|(?:^|\s)[-•*]\s+")
_TACH_LIEN_TU = re.compile(r"\s+(?:va|con|ngoai ra|dong thoi|kem theo|cung voi)\s+")


def tach_y(cau_hoi: str) -> list:
    tho = [p.strip(" ,.?") for p in _TACH_MANH.split(cau_hoi or "") if p and p.strip(" ,.?")]
    if not tho:
        return []
    ra = []
    for phan in tho:
        dac = _dac(phan)
        vi_tri = [(m.start(), m.end()) for m in _TACH_LIEN_TU.finditer(dac)]
        if not vi_tri:
            ra.append(phan)
            continue
        # thử tách ở từng liên từ, chỉ nhận nếu cả hai vế đều khớp loại
        da_tach = False
        for s, e in vi_tri:
            trai, phai = phan[:s].strip(" ,.?"), phan[e:].strip(" ,.?")
            if not (trai and phai):
                continue
            lt, lp = _khop_loai(trai), _khop_loai(phai)
            # CÙNG loại thì KHÔNG tách: "doanh thu và chi phí tháng 7" là MỘT ý hỏi hai chỉ tiêu.
            # Tách ra sẽ đẻ một ý cụt không có kỳ ("doanh thu") — tệ hơn hẳn để nguyên.
            if lt and lp and lt[0][0]["id"] != lp[0][0]["id"]:
                ra.extend(tach_y(trai))
                ra.extend(tach_y(phai))
                da_tach = True
                break
        if not da_tach:
            ra.append(phan)
    return ra


def _khop_loai(text: str) -> list:
    """Trả [(loai, số từ khoá khớp)] sắp giảm dần."""
    dac = _dac(text)
    diem = []
    for lo in so_tay()["loai"]:
        n = sum(1 for t in lo["_nhan_dien_dac"] if t and t in dac)
        if n:
            diem.append((lo, n))
    diem.sort(key=lambda x: -x[1])
    return diem


# --- Đọc tham số ----------------------------------------------------------------------------
_RE_THANG = re.compile(r"\b(?:thang|t)\s*[.]?\s*(\d{1,2})\b")
_RE_THANG_NAM = re.compile(r"\b(\d{1,2})\s*[/-]\s*(20\d{2})\b")
_RE_KY_ISO = re.compile(r"\b(20\d{2})\s*[-/]\s*(\d{1,2})\b")
_RE_NAM = re.compile(r"\b(?:nam\s*)?(20\d{2})\b")
_RE_QUY = re.compile(r"\bquy\s*([1-4iv]+)\b")
_RE_N_THANG_DAU = re.compile(r"\b(\d{1,2})\s*thang\s*dau\s*nam\b")

_QUY_THANG = {"1": (1, 3), "i": (1, 3), "2": (4, 6), "ii": (4, 6),
              "3": (7, 9), "iii": (7, 9), "4": (10, 12), "iv": (10, 12)}


def doc_ky(text: str) -> dict:
    """Suy kỳ từ câu hỏi. Trả {ky, year, month, dai_ky:[...], mo_ta}. Không đoán bừa: không nêu kỳ
    thì trả rỗng để tầng trên áp quy tắc mặc định (kỳ ĐÃ CHỐT gần nhất) và nói rõ đã chọn kỳ nào."""
    dac = _dac(text)
    nam = thang = None
    dai = None

    m = _RE_KY_ISO.search(dac)
    if m:
        nam, thang = int(m.group(1)), int(m.group(2))
    if thang is None:
        m = _RE_THANG_NAM.search(dac)
        if m:
            thang, nam = int(m.group(1)), int(m.group(2))
    if thang is None:
        m = _RE_THANG.search(dac)
        if m:
            thang = int(m.group(1))
    if nam is None:
        m = _RE_NAM.search(dac)
        if m:
            nam = int(m.group(1))

    m = _RE_N_THANG_DAU.search(dac)
    if m:
        dai = list(range(1, min(12, int(m.group(1))) + 1))
    if dai is None:
        m = _RE_QUY.search(dac)
        if m:
            r = _QUY_THANG.get(m.group(1))
            if r:
                dai = list(range(r[0], r[1] + 1))

    if thang is not None and not 1 <= thang <= 12:
        thang = None
    ra = {"year": nam, "month": thang, "dai_thang": dai,
          "ky": f"{nam:04d}-{thang:02d}" if (nam and thang) else None}
    ra["thieu_nam"] = thang is not None and nam is None
    ra["thieu_ky"] = thang is None and dai is None
    return ra


def _danh_muc_don_vi() -> list:
    """Danh mục đơn vị lấy từ master_data + các token nguồn hay gặp. KHÔNG hardcode trong SKILL."""
    ra = []
    try:
        from . import be_bridge as bb
        md = bb.master_data()
        for c in md.get("companies", []):
            ra.append({"loai": "cong_ty", "ma": c.get("ma"), "ten": c.get("ten")})
        for k in md.get("khoi", []):
            ra.append({"loai": "khoi", "ma": k.get("ma"), "ten": k.get("ten")})
        for cc in md.get("costCenters", []):
            ra.append({"loai": "cost_center", "ma": cc.get("ma"), "ten": cc.get("ten"),
                       "cong_ty": cc.get("congTy"), "khoi": cc.get("khoi")})
    except Exception:                                            # noqa: BLE001
        pass
    # 5 nhóm nội bộ của TC: mỗi nhóm một file riêng, hỏi "TC" phải gộp đủ 5.
    for ma, ten in [("SRVF", "Showroom Vinfast (TC)"), ("DUAN", "Khối dự án (TC)"),
                    ("TRAMSAC", "Trạm sạc (TC)"), ("HO", "Hỗ trợ tập đoàn (TC)"),
                    ("XDV", "Xưởng dịch vụ Vinfast (TC)")]:
        ra.append({"loai": "nhom_noi_bo_tc", "ma": ma, "ten": ten})
    return ra


# Tiền tố loại hình pháp lý — bỏ đi để "Hưng Thịnh" khớp được
# "Công ty TNHH Xuất nhập khẩu và Khai thác Hưng Thịnh".
_RE_TIEN_TO_PL = re.compile(
    r"^(cong ty|cty|co phan|cp|tnhh|mtv|hop tac xa|htx|tap doan|chi nhanh|"
    r"xuat nhap khau|khai thac|dich vu|cong nghe|van tai|va|)\s*", re.IGNORECASE)


def _ten_ngan(ten: str) -> str:
    """Cắt dần tiền tố loại hình pháp lý để lấy phần tên riêng."""
    s = _dac(ten)
    truoc = None
    while s and s != truoc:
        truoc = s
        s = _RE_TIEN_TO_PL.sub("", s, count=1).strip()
    return s


def doc_don_vi(text: str) -> list:
    dac = _dac(text)
    ra, da_thay = [], set()
    for m in _danh_muc_don_vi():
        for khoa in (m.get("ma"), m.get("ten"), _ten_ngan(m.get("ten") or "")):
            k = _dac(khoa)
            if len(k) < 2:
                continue
            # mã ngắn (TC, GA, HT) phải khớp NGUYÊN TỪ, không thì "GA" trúng "giá vốn"
            hit = re.search(rf"(?<![a-z0-9]){re.escape(k)}(?![a-z0-9])", dac) if len(k) <= 4 else (k in dac)
            if hit and m.get("ma") not in da_thay:
                da_thay.add(m.get("ma"))
                ra.append(m)
                break
    return ra


def phan_loai(cau_hoi: str, gioi_han_loai: int = 2) -> dict:
    """Tách ý -> khớp loại -> trả CÔNG THỨC ĐẦY ĐỦ cho từng ý (kèm bat_buoc + bay)."""
    st = so_tay()
    cac_y = tach_y(cau_hoi)
    ds = []
    for i, y in enumerate(cac_y, 1):
        diem = _khop_loai(y)
        ky = doc_ky(y)
        if ky["thieu_ky"] and len(cac_y) > 1:            # ý sau thừa hưởng kỳ của ý trước
            for truoc in reversed(ds):
                if not truoc["tham_so"]["ky_doc_duoc"]["thieu_ky"]:
                    ky = dict(truoc["tham_so"]["ky_doc_duoc"])
                    ky["ke_thua_tu_y_truoc"] = True
                    break
        muc = {
            "y_id": f"y{i}",
            "noi_dung": y,
            "loai": [
                {k: v for k, v in lo.items() if not k.startswith("_")}
                for lo, _n in diem[:gioi_han_loai]
            ],
            "tham_so": {"ky_doc_duoc": ky, "don_vi_doc_duoc": doc_don_vi(y)},
        }
        if not diem:
            muc["canh_bao"] = ("Không khớp loại câu hỏi nào. PHẢI nói rõ với người dùng là chưa hiểu "
                               "ý này thay vì bỏ qua, hoặc hỏi lại cho rõ.")
        ds.append(muc)

    return {
        "so_y": len(ds),
        "y": ds,
        "da_yeu_cau": [m["y_id"] for m in ds],
        "bat_buoc_dau_ra": (
            "Mỗi ý PHẢI có một mục riêng trong câu trả lời, đúng thứ tự người dùng hỏi. "
            "Ý nào không trả lời được thì nói rõ ý đó và vì sao — cấm lặng lẽ chỉ trả lời phần làm được."
        ) if len(ds) > 1 else None,
        "bay_chung": st.get("_bay_chung", []),
    }
