# -*- coding: utf-8 -*-
"""Kiểm <NGU_CANH> của backend mà không phải khởi động FastAPI (BE là repo anh em)."""
import importlib.util
import os

_HERE = os.path.dirname(os.path.abspath(__file__))
_BE = os.path.normpath(os.path.join(_HERE, "..", "..", "DashBoard_AI", "backend", "app",
                                    "source_bridge.py"))


def _nap():
    # BE là repo anh em, venv của Dashboard_Agent không có fastapi/pydantic. Chỉ cần đọc 2 hàm
    # ngữ cảnh nên nhét stub tối thiểu thay vì cài thêm phụ thuộc chỉ để chạy test.
    import sys
    import types
    if "fastapi" not in sys.modules:
        fa = types.ModuleType("fastapi")
        fa.APIRouter = lambda **k: types.SimpleNamespace(
            post=lambda *a, **k: (lambda f: f), get=lambda *a, **k: (lambda f: f))
        fa.HTTPException = type("HTTPException", (Exception,), {})
        sys.modules["fastapi"] = fa
    if "pydantic" not in sys.modules:
        pd = types.ModuleType("pydantic")
        pd.BaseModel = type("BaseModel", (), {})
        sys.modules["pydantic"] = pd
    spec = importlib.util.spec_from_file_location("sb_probe", _BE)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def kiem_ngu_canh() -> bool:
    m = _nap()
    sid = "conv-test-20260813"
    assert m._ngu_canh(sid) == "", "phiên mới phải KHÔNG có ngữ cảnh"
    m._ghi_ngu_canh(sid, "Doanh thu tháng 7?", "Doanh thu T7 = 628,7 tỷ. Nguồn: 11 file BCTC. " + "x" * 5000)
    nc = m._ngu_canh(sid)
    if "Doanh thu tháng 7?" not in nc:
        return False
    if len(nc) > 1200:                      # payload thô bị nhét vào là hỏng mục đích
        return False
    for i in range(5):                      # chỉ giữ 3 lượt gần nhất
        m._ghi_ngu_canh(sid, f"hoi {i}", f"dap {i}")
    return m._ngu_canh(sid).count("- Hỏi:") == 3
