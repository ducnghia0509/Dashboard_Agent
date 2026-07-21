# -*- coding: utf-8 -*-
"""SINH LẠI 3 sheet MD_ (MD_CONGTY / MD_KHOIKD / MD_COSTCENTER) trong Template_chuan.xlsx TỪ YAML.

Để danh mục BÊN TRONG file vàng (VLOOKUP + dropdown) khớp knowledge/*.yaml — knowledge là nguồn.

QUAN TRỌNG: ghi ở tầng ZIP/XML, CHỈ thay <sheetData> của 3 part XML thuộc MD_; mọi part khác
(13 sheet dữ liệu, data-validation x14, công thức, style, sharedStrings) giữ NGUYÊN BYTE.
KHÔNG dùng openpyxl.save (nó xoá data-validation mở rộng x14 của các sheet nhập liệu).

Header (dòng 1) giữ nguyên verbatim; dữ liệu (dòng 2+) ghi dạng inlineStr, tái dùng style của
cột từ dòng dữ liệu cũ. Mặc định ghi ra BẢN COPY để review; --inplace mới ghi đè file vàng.
  python scripts/gen_template_md.py --out /tmp/Template_new.xlsx
  python scripts/gen_template_md.py --inplace          # ghi đè (nên backup trước)
"""
import os
import re
import shutil
import sys
import zipfile
from xml.sax.saxutils import escape

import yaml

KNOWLEDGE = os.environ.get("KNOWLEDGE_DIR", "/home/sysadmin/knowledge")
GOLDEN = os.environ.get("GOLDEN_TEMPLATE", "/home/sysadmin/template_trust/Template_chuan.xlsx")


def _load(name):
    with open(os.path.join(KNOWLEDGE, name), encoding="utf-8") as fh:
        return yaml.safe_load(fh)


def _rows_for_md():
    ccs_raw = _load("cost_centers.yaml") or []
    # "aggregate-only" (vd GR) = mã KHÔNG có cost center nào gán -> loại khỏi MD_CONGTY (giữ khớp
    # gen_org.py). Suy từ dữ liệu, không cần cờ role trong companies.yaml (đồng bộ 2 script).
    used = {str(c["ma_congty"]).strip() for c in ccs_raw if c.get("ma_congty")}
    companies = [(str(c["ma_congty"]).strip(), str(c["ten_congty"]).strip())
                 for c in (_load("companies.yaml") or [])
                 if str(c["ma_congty"]).strip() in used]
    khoi = [(str(k["ma_khoi"]).strip(), str(k["ten_khoi"]).strip())
            for k in (_load("khoi.yaml") or [])]
    ccs = [(str(c["ma_cost_center"]).strip(), str(c.get("ten_cost_center") or "").strip(),
            str(c.get("ma_congty") or "").strip(), str(c.get("ma_khoi") or "").strip())
           for c in ccs_raw if c.get("ma_cost_center")]
    return {"MD_CONGTY": companies, "MD_KHOIKD": khoi, "MD_COSTCENTER": ccs}


def _col_letter(idx):  # 1 -> A
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def _sheet_paths(zf):
    """name sheet -> part path (xl/worksheets/sheetN.xml) qua workbook.xml + rels."""
    wb = zf.read("xl/workbook.xml").decode("utf-8")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid_target = dict(re.findall(r'Id="(rId\d+)"[^>]*?Target="([^"]+)"', rels))
    out = {}
    for name, rid in re.findall(r'<sheet[^>]*?name="([^"]+)"[^>]*?r:id="(rId\d+)"', wb):
        tgt = rid_target.get(rid)
        if not tgt:
            continue
        out[name] = ("xl/" + tgt) if not tgt.startswith("/") else tgt.lstrip("/")
    return out


def _styles_from_row2(xml, ncol):
    """Lấy style index (thuộc tính s) của từng cột từ dòng dữ liệu cũ (r=2) để tái dùng."""
    m = re.search(r'<row[^>]*\br="2"[^>]*>(.*?)</row>', xml, re.S)
    styles = {}
    if m:
        for cm in re.finditer(r'<c\b([^>]*)/?>', m.group(1)):
            attrs = cm.group(1)
            ref = re.search(r'r="([A-Z]+)\d+"', attrs)
            sty = re.search(r's="(\d+)"', attrs)
            if ref:
                styles[ref.group(1)] = sty.group(1) if sty else None
    return styles


def _build_rows(rows, ncol, styles):
    out = []
    for i, tup in enumerate(rows, start=2):
        cells = []
        for j in range(ncol):
            col = _col_letter(j + 1)
            val = tup[j] if j < len(tup) else ""
            s = styles.get(col)
            sattr = f' s="{s}"' if s else ""
            cells.append(f'<c r="{col}{i}"{sattr} t="inlineStr">'
                         f'<is><t xml:space="preserve">{escape("" if val is None else str(val))}</t></is></c>')
        out.append(f'<row r="{i}">' + "".join(cells) + "</row>")
    return "".join(out)


def _rewrite_sheet_xml(xml, rows, ncol):
    styles = _styles_from_row2(xml, ncol)
    header = re.search(r'(<row[^>]*\br="1"[^>]*>.*?</row>)', xml, re.S)
    header_xml = header.group(1) if header else ""
    new_data = "<sheetData>" + header_xml + _build_rows(rows, ncol, styles) + "</sheetData>"
    xml2 = re.sub(r'<sheetData\b[^>]*>.*?</sheetData>|<sheetData\s*/>', new_data, xml, count=1, flags=re.S)
    # cập nhật dimension ref (nếu có) theo số dòng mới
    last = 1 + len(rows)
    xml2 = re.sub(r'(<dimension ref=")[^"]*(")', rf'\g<1>A1:{_col_letter(ncol)}{last}\g<2>', xml2)
    return xml2


def build(out_path=None, inplace=False):
    if not inplace and not out_path:
        raise SystemExit("Cần --out <path> (bản copy) hoặc --inplace (ghi đè file vàng).")
    dest = GOLDEN if inplace else out_path
    md = _rows_for_md()
    ncols = {"MD_CONGTY": 2, "MD_KHOIKD": 2, "MD_COSTCENTER": 4}
    src = GOLDEN
    with zipfile.ZipFile(src) as zin:
        paths = _sheet_paths(zin)
        targets = {}  # part path -> new xml
        for sheet, rows in md.items():
            p = paths.get(sheet)
            if not p or p not in zin.namelist():
                raise SystemExit(f"Không tìm thấy part XML cho sheet {sheet}")
            xml = zin.read(p).decode("utf-8")
            targets[p] = _rewrite_sheet_xml(xml, rows, ncols[sheet])
        names = zin.namelist()
        infos = {i.filename: i for i in zin.infolist()}
        # ghi zip mới: copy nguyên byte mọi part, chỉ thay 3 part MD_
        tmp = (dest + ".tmp")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                data = targets[name].encode("utf-8") if name in targets else zin.read(name)
                zi = zipfile.ZipInfo(name, date_time=infos[name].date_time)
                zi.compress_type = zipfile.ZIP_DEFLATED
                zi.external_attr = infos[name].external_attr
                zout.writestr(zi, data)
    if inplace and src == dest:
        shutil.move(tmp, dest)
    else:
        shutil.move(tmp, dest)
    return {"ok": True, "out": dest, "inplace": inplace,
            "written": {k: len(v) for k, v in md.items()}}


if __name__ == "__main__":
    import json
    out = sys.argv[sys.argv.index("--out") + 1] if "--out" in sys.argv else None
    print(json.dumps(build(out, inplace=("--inplace" in sys.argv)), ensure_ascii=False))
