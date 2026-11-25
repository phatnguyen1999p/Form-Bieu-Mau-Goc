"""
[NEW 20/11/2026] Bộ sinh hồ sơ Quan trắc môi trường lao động (MTLĐ) theo KẾ HOẠCH TỪNG PHÒNG.

Dùng chung cho mọi hồ sơ MTLĐ (khởi nguồn từ hồ sơ QUATEST 3 – KTN Cát Lái 2026):
  - BM3030.01 Ghi chép: trang thông tin + bảng theo dõi khối lượng theo phòng + mỗi phòng 1 trang gồm
      bảng "Máy đo sử dụng" (tick ☐), bảng đo trực tiếp (ô xám = không đo), ồn dải tần (ghi tên file đo),
      bảng lấy mẫu hơi khí độc / vi sinh (cột Bơm tick ☐), bảng Ecgonomy (OWAS & vị trí lao động).
  - BM3030.02 Biên bản: mỗi chỉ tiêu 1 dòng, ghi sẵn các phòng thực hiện, SL / kí hiệu / niêm phong để chấm.

Đầu vào: file Excel kế hoạch có các cột Phòng | Chỉ tiêu | SL | Điểm đo (Nhóm, Ghi chú tùy chọn).
  Cột "Điểm đo" ghi dạng "P304: 02; P303: 01" -> tách thành từng khu vực với số điểm.
  Ghi không rõ số (ví dụ "Phòng XD") -> cả cụm được tính 1 khu vực với số điểm = SL.
Mẫu: templates/MTLD_KeHoach_mau.xlsx
"""
import os
import re
import copy
import shutil
from collections import OrderedDict

import docx
import openpyxl
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

from .config import BASE_DIR, Z_DRIVE_DIR
from .normalizer import (
    DOC_NUMBER_BLANK_PLACEHOLDER, is_blank_doc_number, doc_number_for_filename,
    normalize_address, normalize_customer_name, build_output_folder_name, sanitize_for_filename,
    LOCATION_DOTS,
)

TPL01 = os.path.join(BASE_DIR, "CO2 - MTLD", "BM3030.01- Ghi chép - Q326D091792.docx")
TPL02 = os.path.join(BASE_DIR, "CO2 - MTLD", "BM3030.02- BBLM - Q326D091792.docx")
FONT = "Times New Roman"
USABLE = 18.0           # bề rộng vùng in (cm) của mẫu A4
LA = WD_ALIGN_PARAGRAPH.LEFT
CE = WD_ALIGN_PARAGRAPH.CENTER

# ==============================================================================
# 1. DANH MỤC CHỈ TIÊU
# ==============================================================================
# Chỉ tiêu đo đọc trực tiếp -> (tiêu đề cột, bề rộng cột cm hoặc None = mặc định)
DIRECT = OrderedDict([
    ("Nhiệt độ", ("t\n(°C)", None)),
    ("Độ ẩm", ("RH\n(%)", None)),
    ("Tốc độ gió", ("V gió\n(m/s)", None)),
    ("Bức xạ nhiệt", ("Bức xạ\nnhiệt\n(W/m²)", None)),
    ("Điện từ trường", ("E (kV/m)\n/ H (A/m)", 2.2)),
    ("Ánh sáng", ("Ánh sáng\n(Lux)", None)),
    ("Ồn chung", ("Ồn\n(tên file đo)", 2.4)),
    ("Bụi hô hấp", ("Bụi HH\n(mg/m³)", None)),
    ("Bụi toàn phần", ("Bụi TP\n(mg/m³)", None)),
    ("CO₂", ("CO₂\n(ppm)", None)),
])

# Tên gọi khác -> tên chuẩn (so khớp không dấu cách, chữ thường)
ALIASES = [
    (r"^nhi[eệ]t\s*đ[oộ]", "Nhiệt độ"),
    (r"^đ[oộ]\s*[aẩ]m", "Độ ẩm"),
    (r"(t[oố]c\s*đ[oộ]|v[aậ]n\s*t[oố]c)\s*gi[oó]", "Tốc độ gió"),
    (r"b[uứ]c\s*x[aạ]\s*nhi[eệ]t", "Bức xạ nhiệt"),
    (r"đi[eệ]n\s*t[uừ]\s*tr[uư][oờ]ng|electromagnetic|\bemf\b", "Điện từ trường"),
    (r"^(ánh\s*sáng|độ\s*sáng|độ\s*rọi|cường\s*độ\s*(chiếu\s*)?sáng)", "Ánh sáng"),
    (r"d[aả]i\s*t[aầ]n", "Ồn dải tần"),
    (r"^(ồn|tiếng\s*ồn|độ\s*ồn)", "Ồn chung"),
    (r"b[uụ]i\s*h[oô]\s*h[aấ]p", "Bụi hô hấp"),
    (r"b[uụ]i\s*to[aà]n\s*ph[aầ]n", "Bụi toàn phần"),
    (r"^co\s*[2₂]$", "CO₂"),
]


def canon_indicator(name):
    s = str(name or "").strip()
    low = s.lower()
    for pat, canon in ALIASES:
        if re.search(pat, low):
            return canon
    return s


def kind_of(ct):
    """direct | octave | ecg | assess | micro | sample"""
    low = ct.lower()
    if ct in DIRECT:
        return "direct"
    if ct == "Ồn dải tần":
        return "octave"
    if re.search(r"tiếp xúc nghề nghiệp|yếu tố tiếp xúc", low):
        return "assess"
    if re.search(r"ecgonom|ergonom|owas|vị trí lao động|tư thế", low):
        return "ecg"
    if re.search(r"nấm|vi sinh|vi khuẩn|vsv", low):
        return "micro"
    return "sample"


GROUP_OF_KIND = {"ecg": "Ecgonomy", "assess": "Ecgonomy", "micro": "Vi sinh", "sample": "Hơi khí độc", "octave": "Vật lý"}
GROUP_OF_DIRECT = {"Nhiệt độ": "Vi khí hậu", "Độ ẩm": "Vi khí hậu", "Tốc độ gió": "Vi khí hậu", "Bức xạ nhiệt": "Vi khí hậu",
                   "Ánh sáng": "Vật lý", "Ồn chung": "Vật lý", "Điện từ trường": "Vật lý", "Bụi hô hấp": "Bụi", "Bụi toàn phần": "Bụi",
                   "CO₂": "Hơi khí độc"}
GROUP_ORDER = ["Vi khí hậu", "Vật lý", "Bụi", "Hơi khí độc", "Vi sinh", "Ecgonomy"]

# Máy đo để tick ☐ tại hiện trường (mỗi phòng): nhãn -> (chỉ tiêu áp dụng, các máy)
DEVICES = OrderedDict([
    # [24/11/2026] Nhiệt độ, độ ẩm đo bằng skSATO SK-170GT (thay Testo 625)
    ("Nhiệt độ, độ ẩm", (("Nhiệt độ", "Độ ẩm"), ["skSATO SK-170GT", "Testo 440", "Climomaster 6533-2G"])),
    ("Tốc độ gió", (("Tốc độ gió",), ["Climomaster 6533-2G", "Climomaster 6501"])),
    ("Bức xạ nhiệt", (("Bức xạ nhiệt",), ["skSATO SK-170GT"])),
    ("Điện từ trường", (("Điện từ trường",), ["EMF450"])),
    ("Ánh sáng", (("Ánh sáng",), ["Testo 440 – máy 1", "Testo 440 – máy 2"])),
    ("Ồn", (("Ồn chung", "Ồn dải tần"), ["Rion NL-53"])),
    ("Bụi", (("Bụi hô hấp", "Bụi toàn phần"), ["Temptop PMD 371"])),
    ("CO₂", (("CO₂",), ["testo 440"])),
])
PUMPS_LEGEND = "Buck = Buck Genie; SKC = SKC PCXR8; Kitagawa = Kitagawa ASP 1200"
PUMP_CELL = "☐ Buck\n☐ SKC\n☐ Kitagawa"
MICRO_CELL = "☐ Bio Culture"
DOTS_KHAC = "Khác: ……………"

# Thiết bị (mục Thiết bị sử dụng) + quy trình N2.TC theo nhóm chỉ tiêu có trong hồ sơ
EQUIP_RULES = [
    # [24/11/2026] skSATO SK-170GT đo cả nhiệt độ, độ ẩm và bức xạ nhiệt -> 1 dòng; EMF450 cho điện từ trường
    (("Nhiệt độ", "Độ ẩm", "Bức xạ nhiệt"), ["Thiết bị đo nhiệt độ, độ ẩm, bức xạ nhiệt skSATO SK-170GT, Japan"]),
    (("Tốc độ gió",), ["Máy đo tốc độ gió Kanomax Climomaster 6533-2G, Japan"]),
    (("Điện từ trường",), ["Máy đo điện từ trường EMF450"]),
    (("Ánh sáng",), ["Máy đo độ sáng Testo 440 Kit Lux, Testo – Germany"]),
    (("Ồn chung", "Ồn dải tần"), ["Thiết bị đo mức âm thanh Rion NL-53, Japan"]),
    (("Bụi hô hấp", "Bụi toàn phần"), ["Máy đo nồng độ bụi Temptop PMD 371, USA"]),
    (("CO₂",), ["Thiết bị đo nồng độ CO2 trong không khí testo 440, Germany"]),
    (("@sample",), ["Bơm lấy mẫu Buck Genie, USA; bơm lấy mẫu SKC PCXR8, USA",
                    "Bơm lấy mẫu Kitagawa ASP 1200, Japan", "Impinger + Cassette"]),
    (("@micro",), ["Thiết bị lấy mẫu vi sinh trong không khí Buck – Bio Culture B30120, USA"]),
    (("@ecg",), ["Thước dây, thước đo góc (đánh giá Ecgonomy)"]),
]
PROC_RULES = [
    (("Nhiệt độ", "Độ ẩm", "Tốc độ gió", "Bức xạ nhiệt", "Ánh sáng"), "N2.TC.01"),
    (("Ồn chung", "Ồn dải tần"), "N2.TC.02"),
    (("Điện từ trường",), "N2.TC.05"),
    (("Bụi hô hấp", "Bụi toàn phần"), "N2.TC.06"),
    (("CO₂", "@sample"), "N2.TC.07"),
    (("@micro",), "N2.TC.09"),
    (("@ecg",), "N2.TC.11, N2.TC.12"),
]

BBLM_LABEL = {
    "Nhiệt độ": "Đo nhiệt độ môi trường lao động",
    "Độ ẩm": "Đo độ ẩm môi trường lao động",
    "Tốc độ gió": "Đo tốc độ gió môi trường lao động",
    "Bức xạ nhiệt": "Đo bức xạ nhiệt môi trường lao động",
    "Ánh sáng": "Đo độ sáng môi trường lao động",
    "Ồn chung": "Đo độ ồn chung môi trường lao động",
    "Ồn dải tần": "Đo độ ồn theo dải tần môi trường lao động",
    "Bụi hô hấp": "Đo nồng độ bụi hô hấp môi trường lao động",
    "Bụi toàn phần": "Đo nồng độ bụi toàn phần môi trường lao động",
    "CO₂": "Đo nồng độ CO2 trong không khí môi trường lao động",
}

# ==============================================================================
# 2. ĐỌC KẾ HOẠCH
# ==============================================================================
HEADER_KEYS = {
    "phong": r"^(phòng|khu vực|bộ phận|phân xưởng)",
    "ct": r"^chỉ tiêu",
    "sl": r"^(sl|số lượng)",
    "diem": r"điểm đo|vị trí",
    "nhom": r"^nhóm",
    "gc": r"^ghi chú",
}


def _find_header(ws):
    for r in range(1, 15):
        vals = [str(ws.cell(r, c).value or "").strip().lower() for c in range(1, ws.max_column + 1)]
        cols = {}
        for key, pat in HEADER_KEYS.items():
            for ci, v in enumerate(vals):
                if v and re.search(pat, v) and key not in cols and ci not in cols.values():
                    cols[key] = ci
        if {"phong", "ct", "sl"} <= set(cols):
            return r, cols
    return None, None


def load_plan(path, sheet=None):
    """Trả về OrderedDict phòng -> [dict(ct, sl, diem, nhom)] (bỏ dòng SL = 0, dòng cộng / tiêu đề nhóm)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    sheets = [wb[sheet]] if sheet else ([wb["Ke hoach theo phong"]] if "Ke hoach theo phong" in wb.sheetnames else wb.worksheets)
    for ws in sheets:
        hr, cols = _find_header(ws)
        if not hr:
            continue
        rooms = OrderedDict()
        last_room = None
        for row in ws.iter_rows(min_row=hr + 1, values_only=True):
            g = lambda k: row[cols[k]] if k in cols and cols[k] < len(row) else None
            phong, ct, sl = g("phong"), g("ct"), g("sl")
            if not ct or str(ct).strip().lower().startswith(("cộng", "tổng")):
                continue
            try:
                sl = int(float(sl or 0))
            except (TypeError, ValueError):
                continue
            phong = str(phong).strip() if phong not in (None, "") else last_room
            if not phong or sl <= 0 or str(phong).lower().startswith(("cộng", "tổng")):
                continue
            last_room = phong
            rooms.setdefault(phong, []).append(dict(
                ct=canon_indicator(ct), sl=sl,
                diem=str(g("diem") or "").strip(), nhom=str(g("nhom") or "").strip(),
                gc=str(g("gc") or "").strip()))
        if rooms:
            return rooms
    raise ValueError(f"Không tìm thấy bảng kế hoạch (cần các cột Phòng | Chỉ tiêu | SL | Điểm đo) trong {path}")


def parse_locs(diem, sl, room):
    """'P304: 02; P303: 01' -> [('P304', 2), ('P303', 1)]; không khớp SL -> cả cụm là 1 khu vực."""
    if not diem:
        return [(room, sl)]
    out = []
    for s in [x.strip() for x in diem.split(";") if x.strip()]:
        m = re.match(r"^(.*?):\s*(\d+)$", s)
        out.append((m.group(1).strip(), int(m.group(2))) if m else (s, None))
    known = sum(n for _, n in out if n)
    unk = [i for i, (_, n) in enumerate(out) if n is None]
    if not unk and known == sl:
        return out
    if len(unk) == 1 and sl - known > 0:
        out[unk[0]] = (out[unk[0]][0], sl - known)
        return out
    return [(diem, sl)]


def _norm_loc(s):
    s = re.sub(r"\s*–\s*thay F\d+", "", s)
    s = re.sub(r"\s*\(thay F\d+\)", "", s)
    return s.strip()


def build_room_data(rooms, fixes=None):
    """fixes: {(phòng, chỉ tiêu): [(khu vực, số điểm), ...]} để sửa dòng kế hoạch ghi điểm đo chưa rõ."""
    fixes = fixes or {}
    out = OrderedDict()
    for p, items in rooms.items():
        direct, sampled, ecg, octave, assess = OrderedDict(), [], [], [], []
        for it in items:
            ct = it["ct"]
            locs = fixes.get((p, ct)) or parse_locs(it["diem"], it["sl"], p)
            locs = [(_norm_loc(l), n) for l, n in locs]
            k = kind_of(ct)
            it["kind"] = k
            it["nhom"] = it["nhom"] or (GROUP_OF_DIRECT.get(ct) if k == "direct" else GROUP_OF_KIND[k])
            for l, n in locs:
                if k == "direct":
                    direct.setdefault(l, {})[ct] = direct.get(l, {}).get(ct, 0) + n
                elif k == "ecg":
                    ecg += [l] * n
                elif k == "assess":
                    assess += [l] * n
                elif k == "octave":
                    octave += [l] * n
                else:
                    sampled += [(l, ct, k, it.get("gc", ""))] * n
        out[p] = dict(items=items, direct=direct, sampled=sampled, ecg=ecg, octave=octave, assess=assess)
    return out


def _present_keys(data):
    keys = set()
    for rd in data.values():
        for it in rd["items"]:
            keys.add(it["ct"])
            if it["kind"] in ("sample", "micro", "ecg", "assess"):
                keys.add("@" + it["kind"])
    return keys


def equipment_lines(data):
    keys = _present_keys(data)
    lines = [e for trig, eqs in EQUIP_RULES if keys & set(trig) for e in eqs]
    return [f"- {e}{'.' if i == len(lines) - 1 else ';'}" for i, e in enumerate(lines)]


def procedure_text(data):
    keys = _present_keys(data)
    codes = [c for trig, c in PROC_RULES if keys & set(trig)]
    return "QUATEST 3 3030:2025; " + ", ".join(codes)


def is_room_code(p):
    """Mã phòng ngắn (QT, HTD, IT-XT (mới), Kho) -> True; tên dài (Xưởng may) -> False."""
    return bool(re.match(r"^[A-ZĐ0-9][A-ZĐ0-9\-]{0,7}(\s*\(mới\))?$", p)) or len(p) <= 4


def is_blank_sample(it):
    """Mẫu blank (trắng hiện trường / vận chuyển): có chữ 'blank' ở tên chỉ tiêu hoặc ghi chú kế hoạch.
    Không tính vào bảng theo dõi khối lượng, nhưng vẫn giữ dòng trong bảng lấy mẫu."""
    return "blank" in f"{it.get('ct', '')} {it.get('gc', '')}".lower()


def is_blank_loc(s):
    """Khu vực để chấm điền tay ('……', 'Khu vực: ……')."""
    return bool(re.fullmatch(r"(khu vực:?)?[\s.…]*", str(s or "").strip().lower()))


def loc_cell(s):
    return "" if is_blank_loc(s) else s


def room_title(p):
    return f"Phòng {p}" if is_room_code(p) else p


# ==============================================================================
# 3. TIỆN ÍCH DOCX
# ==============================================================================
def set_run_font(run, size=13, bold=False, italic=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), FONT)


def para_fmt(p, before=2, after=2, line=None, align=None):
    pf = p.paragraph_format
    pf.space_before, pf.space_after = Pt(before), Pt(after)
    pf.left_indent = pf.right_indent = pf.first_line_indent = Cm(0)
    if line:
        pf.line_spacing = line
    if align is not None:
        p.alignment = align


def _strip_numpr(p):
    ppr = p._p.pPr
    if ppr is not None:
        for e in ppr.findall(qn("w:numPr")):
            ppr.remove(e)


def set_cell(cell, lines, size=13, bold=False, italic=False, align=None, before=2, after=2, line=None):
    if isinstance(lines, str):
        lines = [lines]
    ps = cell.paragraphs
    for extra in ps[1:]:
        extra._p.getparent().remove(extra._p)
    p = ps[0]
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    for i, txt in enumerate(lines):
        if i > 0:
            p = cell.add_paragraph()
        _strip_numpr(p)
        para_fmt(p, before, after, line, align)
        parts = txt.split("\n")
        for j, part in enumerate(parts):
            r = p.add_run(part)
            set_run_font(r, size, bold, italic)
            if j < len(parts) - 1:
                r.add_break()
    return cell


def shade(cell, fill="A6A6A6"):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
    tcpr.append(shd)


def valign(cell, v="center"):
    tcpr = cell._tc.get_or_add_tcPr()
    for e in tcpr.findall(qn("w:vAlign")):
        tcpr.remove(e)
    va = OxmlElement("w:vAlign"); va.set(qn("w:val"), v); tcpr.append(va)


def row_props(row, header=False, height_cm=None):
    trpr = row._tr.get_or_add_trPr()
    trpr.append(OxmlElement("w:cantSplit"))
    if header:
        trpr.append(OxmlElement("w:tblHeader"))
    if height_cm:
        h = OxmlElement("w:trHeight")
        h.set(qn("w:val"), str(int(height_cm * 567))); h.set(qn("w:hRule"), "atLeast")
        trpr.append(h)


def new_table(doc, widths_cm):
    t = doc.add_table(rows=0, cols=len(widths_cm))
    t.style = doc.styles["Table Grid"]
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); t._tbl.tblPr.append(lay)
    for gc, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), widths_cm):
        gc.set(qn("w:w"), str(int(w * 567)))
    t._widths = widths_cm
    return t


def add_row(t, values, size=11, bold=False, header=False, height=None, aligns=None, shade_idx=(), italic=False):
    r = t.add_row()
    row_props(r, header, height)
    for i, (c, w) in enumerate(zip(r.cells, t._widths)):
        c.width = Cm(w)
        v = values[i] if i < len(values) else ""
        al = aligns[i] if aligns and i < len(aligns) and aligns[i] is not None else (CE if header else LA)
        set_cell(c, v, size=size, bold=bold, italic=italic, align=al, before=1, after=1)
        valign(c, "center")
        if header:
            shade(c, "F2F2F2")
        if i in shade_idx:
            shade(c)
    return r


def spacer(doc):
    p = doc.add_paragraph(); para_fmt(p, 0, 0)
    p.paragraph_format.line_spacing = Pt(6)
    set_run_font(p.add_run(""), 6)
    return p


def heading(doc, text, size=13, before=6, after=3, page_break=False):
    p = doc.add_paragraph(); para_fmt(p, before, after, align=LA)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.page_break_before = page_break
    set_run_font(p.add_run(text), size, True)
    return p


def note(doc, text, size=10, italic=True, keep=False):
    p = doc.add_paragraph(); para_fmt(p, 1, 1, align=LA)
    p.paragraph_format.keep_with_next = keep
    set_run_font(p.add_run(text), size, False, italic)
    return p


def remove_after(el):
    body = el.getparent()
    nxt = el.getnext()
    while nxt is not None:
        n2 = nxt.getnext()
        if nxt.tag != qn("w:sectPr"):
            body.remove(nxt)
        nxt = n2


def ucells(row):
    out, prev = [], None
    for c in row.cells:
        if c._tc is prev:
            continue
        prev = c._tc
        out.append(c)
    return out


# ==============================================================================
# 4. BM3030.01 – GHI CHÉP
# ==============================================================================
def build_bm01(data, info):
    d = docx.Document(TPL01)
    t0, t1 = d.tables[0], d.tables[1]
    c0 = ucells(t0.rows[0])
    set_cell(c0[0], "Số:"); set_cell(c0[1], info["doc"]); set_cell(c0[-2], "Ngày:"); set_cell(c0[-1], info["date_blank"])
    for ri, v in {1: info["customer"], 2: info["request"], 3: info["place"], 4: info["time"], 7: info["procedure"], 8: " "}.items():
        set_cell(ucells(t1.rows[ri])[-1], v, align=LA)
    set_cell(ucells(t1.rows[5])[-1], "", align=LA)
    set_cell(ucells(t1.rows[6])[0], info["equipment"], align=LA)
    remove_after(t1._tbl)
    body = d.element.body
    spacer(d)

    # ---- bảng theo dõi khối lượng
    sum_cols = OrderedDict([("Nhiệt độ", ("Nhiệt độ",)), ("Độ ẩm", ("Độ ẩm",)), ("Gió", ("Tốc độ gió",)),
                            ("BX nhiệt", ("Bức xạ nhiệt",)), ("Điện từ", ("Điện từ trường",)), ("Ánh sáng", ("Ánh sáng",)), ("Ồn", ("Ồn chung", "Ồn dải tần")),
                            ("Bụi", ("Bụi hô hấp", "Bụi toàn phần")), ("CO₂", ("CO₂",)),
                            ("HKĐ", ("@sample",)), ("Vi sinh", ("@micro",)), ("Ecg", ("@ecg",)), ("TXNN", ("@assess",))])
    keys = _present_keys(data)
    sum_cols = OrderedDict((k, v) for k, v in sum_cols.items() if keys & set(v))
    heading(d, f"BẢNG THEO DÕI KHỐI LƯỢNG THEO PHÒNG{(' (' + info['plan_title'] + ')') if info.get('plan_title') else ''}", 12, before=2)
    note(d, "Số trong ô = số điểm theo kế hoạch (không tính mẫu blank). Đo xong phòng nào đánh dấu ✓ cột cuối. "
            "HKĐ = hơi khí độc lấy mẫu; Ecg = Ecgonomy (OWAS & vị trí lao động)"
            + ("; TXNN = đánh giá yếu tố tiếp xúc nghề nghiệp." if "@assess" in keys else "."), 10, keep=True)
    n = len(sum_cols)
    wc = min(1.3, (USABLE - 2.2 - 1.2 - 1.2) / max(n, 1))
    w_xong = 1.2
    w_phong = max(2.2, USABLE - n * wc - 1.2 - w_xong)
    w = [w_phong] + [wc] * n + [1.2, w_xong]
    tb = new_table(d, w)
    add_row(tb, ["Phòng"] + list(sum_cols) + ["Tổng", "Xong"], 10, True, header=True)
    tot = [0] * (n + 1)
    for p, rd in data.items():
        cnt = []
        for trig in sum_cols.values():
            cnt.append(sum(it["sl"] for it in rd["items"]
                           if (it["ct"] in trig or ("@" + it["kind"]) in trig) and not is_blank_sample(it)))
        cnt.append(sum(cnt))
        tot = [a + b for a, b in zip(tot, cnt)]
        add_row(tb, [p] + [str(v) if v else "" for v in cnt] + ["☐"], 10, aligns=[LA] + [CE] * (n + 2))
    add_row(tb, ["Tổng"] + [str(v) for v in tot] + [""], 10, True, aligns=[LA] + [CE] * (n + 2))
    if "@ecg" in keys:
        spacer(d)
        note(d, "Mã OWAS (ghi số vào bảng Ecgonomy):  Lưng: 1 thẳng · 2 cúi · 3 thẳng + vặn · 4 cúi + vặn   |   "
                "Tay: 1 hai tay dưới vai · 2 một tay trên vai · 3 hai tay trên vai   |   "
                "Chân: 1 ngồi · 2 đứng 2 chân thẳng · 3 đứng 1 chân thẳng · 4 hai chân khuỵu · 5 một chân khuỵu · 6 một chân quỳ · 7 đi lại   |   "
                "Vật nặng: 1 <10 kg · 2 10–20 kg · 3 >20 kg", 10, italic=False)
        note(d, "Vị trí lao động (QĐ 3733/2002/QĐ-BYT):  Ngồi: bề mặt LV 59–83 cm, chỗ để chân (1) rộng ≥60 · (2) sâu sàn ≥65 · (3) sâu gối ≥45 cm   |   "
                "Đứng: bề mặt LV 72–99 cm, (1) sâu bàn chân ≥15 · (2) cao bàn chân ≥15 · (3) khoảng trống sau lưng ≥90 cm   |   "
                "Mắt–vật 25–50 cm; góc nhìn: ngả sau ≤15°, cúi trước ≤45°", 10, italic=False)

    # ---- mỗi phòng 1 trang
    for p, rd in data.items():
        heading(d, room_title(p).upper(), 14, before=4, after=1, page_break=True)
        note(d, f"Ngày đo: {info['date_blank']}      Người đo: …………………………      Đại diện phòng: …………………………", 11, italic=False, keep=True)
        cts_room = {it["ct"] for it in rd["items"]}
        dev_rows = [(lab, opts) for lab, (trig, opts) in DEVICES.items() if cts_room & set(trig)]
        has_pump = any(s[2] == "sample" for s in rd["sampled"])
        if dev_rows or has_pump:
            tdv = new_table(d, [3.2, USABLE - 3.2])
            add_row(tdv, ["Máy đo sử dụng", "Đánh dấu ☒ máy dùng thực tế"], 10, True, header=True, aligns=[LA, LA])
            for lab, opts in dev_rows:
                add_row(tdv, [lab, "   ".join("☐ " + o for o in opts) + "   ☐ " + DOTS_KHAC], 11, aligns=[LA, LA])
            if has_pump:
                add_row(tdv, ["Bơm lấy mẫu", f"Tick theo từng mẫu ở cột “Bơm” bên dưới ({PUMPS_LEGEND})"], 11, italic=True, aligns=[LA, LA])

        if rd["direct"]:
            used = [k for k in DIRECT if any(k in v for v in rd["direct"].values())]
            heading(d, "Đo trực tiếp: " + ", ".join(k if k == "CO₂" else k.lower() for k in used), 12, before=4, after=2)
            wi = 1.45 if len(used) <= 5 else (1.3 if len(used) <= 7 else 1.15)
            wcols = [DIRECT[k][1] or wi for k in used]
            base = [0.9, 3.0, 0, 1.3]
            base[2] = USABLE - sum(base) - sum(wcols)
            tbl = new_table(d, base + wcols)
            add_row(tbl, ["TT", "Khu vực", "Vị trí cụ thể\n(bàn / người LĐ)", "Giờ đo"] + [DIRECT[k][0] for k in used], 9, True, header=True)
            stt = 0
            for loc, cts in rd["direct"].items():
                for k in range(max(cts.values())):
                    stt += 1
                    sh = [4 + j for j, ct in enumerate(used) if k >= cts.get(ct, 0)]
                    add_row(tbl, [str(stt), loc_cell(loc), "", ""] + [""] * len(used), 11, height=0.85, shade_idx=sh,
                            aligns=[CE, LA] + [CE] * (2 + len(used)))
            note(d, "Ô xám = không đo chỉ tiêu đó tại điểm này.", 9)

        if rd["octave"]:
            heading(d, "Đo ồn theo dải tần số", 12, before=4, after=2)
            w = [0.9, 3.0, 0, 1.3, 4.0]
            w[2] = USABLE - sum(w)
            tbl = new_table(d, w)
            add_row(tbl, ["TT", "Khu vực", "Vị trí cụ thể", "Giờ đo", "Tên file đo\n(ồn chung + dải tần)"], 9, True, header=True)
            for i, loc in enumerate(rd["octave"], 1):
                add_row(tbl, [str(i), loc_cell(loc), "", "", ""], 11, height=0.85, aligns=[CE, LA, None, None, None])

        if rd["sampled"]:
            heading(d, "Lấy mẫu hơi khí độc / vi sinh", 12, before=4, after=2)
            w = [0.9, 2.6, 2.3, 3.0, 2.1, 1.5, 1.3, 1.3, 1.3, 0]
            w[-1] = USABLE - sum(w)
            tbl = new_table(d, w)
            add_row(tbl, ["TT", "Khu vực", "Vị trí cụ thể", "Chỉ tiêu", "Bơm", "Kí hiệu\nmẫu", "Q\n(L/ph)",
                          "Giờ\nbắt đầu", "Giờ\nkết thúc", "Ghi chú"], 9, True, header=True)
            for i, (loc, ct, k, gc) in enumerate(rd["sampled"], 1):
                r_ = add_row(tbl, [str(i), loc_cell(loc), "", ct] + [""] * 5 + [gc], 11, height=0.85, aligns=[CE, LA, None, LA, LA] + [CE] * 5)
                set_cell(r_.cells[4], MICRO_CELL if k == "micro" else PUMP_CELL, 9, align=LA, before=1, after=1)

        if rd["ecg"]:
            heading(d, "Ecgonomy – OWAS & vị trí lao động", 12, before=4, after=2)
            w = [0.9, 1.7, 0, 1.15, 1.15, 1.15, 1.15, 1.15, 1.15, 1.15, 1.15, 0.95, 0.95, 0.95]
            w[2] = USABLE - sum(w)
            tbl = new_table(d, w)
            add_row(tbl, ["TT", "Khu vực", "Người LĐ / công việc", "Ngồi/\nĐứng", "Lưng", "Tay", "Chân", "Vật\nnặng",
                          "Cao\nbề mặt\n(cm)", "Mắt–\nvật\n(cm)", "Góc\nnhìn\n(°)", "(1)\ncm", "(2)\ncm", "(3)\ncm"], 8.5, True, header=True)
            for i, loc in enumerate(rd["ecg"], 1):
                preset = ""
                m = re.match(r"^(\S+)\s+–\s+(.+)$", loc)
                if m and not re.search(r"\d", m.group(2)):
                    loc, preset = m.group(1), m.group(2)
                add_row(tbl, [str(i), loc_cell(loc), preset, "N / Đ"] + [""] * 10, 10, height=0.95, aligns=[CE, LA, LA] + [CE] * 11)
            note(d, "(1)(2)(3): Ngồi = rộng / sâu tại sàn / sâu tại gối chỗ để chân; Đứng = sâu / cao chỗ bàn chân / "
                    "khoảng trống sau lưng. Chụp hình tư thế: ☐", 9)

        if rd["assess"]:
            heading(d, "Đánh giá yếu tố tiếp xúc nghề nghiệp", 12, before=4, after=2)
            w = [0.9, 2.6, 3.4, 0, 1.9, 2.6]
            w[3] = USABLE - sum(w)
            tbl = new_table(d, w)
            add_row(tbl, ["TT", "Khu vực", "Người LĐ / công việc", "Yếu tố tiếp xúc\n(hóa chất, bụi, ồn, nhiệt...)",
                          "Thời gian\ntiếp xúc\n(giờ/ca)", "Phương tiện\nbảo vệ cá nhân"], 9, True, header=True)
            for i, loc in enumerate(rd["assess"], 1):
                add_row(tbl, [str(i), loc_cell(loc), "", "", "", ""], 11, height=1.1, aligns=[CE, LA, LA, LA, CE, LA])
        note(d, "Ghi chú: " + "." * 150, 11, italic=False)

    spacer(d)
    p = d.add_paragraph(); para_fmt(p, 6, 2, align=WD_ALIGN_PARAGRAPH.RIGHT)
    p.paragraph_format.keep_together = True
    set_run_font(p.add_run("Người thực hiện          "), 13, True)
    sp = body.find(qn("w:sectPr")); body.remove(sp); body.append(sp)
    return d


# ==============================================================================
# 5. BM3030.02 – BIÊN BẢN
# ==============================================================================
def build_bm02(data, info):
    d = docx.Document(TPL02)
    t0, t1, t2, t3 = d.tables[:4]
    c0 = ucells(t0.rows[0])
    set_cell(c0[0], "Số:"); set_cell(c0[1], info["doc"]); set_cell(c0[-2], "Ngày:"); set_cell(c0[-1], info["date_blank"])
    r1 = ucells(t1.rows[1])
    set_cell(r1[-3], info["doc"]); set_cell(r1[-1], info["reg_date"])
    for ri, v in {2: info["customer"], 3: info["address"], 4: info["request"], 5: info["place"], 6: info["time"], 8: info["procedure"]}.items():
        set_cell(ucells(t1.rows[ri])[-1], v, align=LA)
    set_cell(ucells(t1.rows[7])[-1], info["equipment"], align=LA)

    body = d.element.body
    paras, el = [], t1._tbl.getnext()
    while el is not None and el.tag == qn("w:p"):
        paras.append(el); el = el.getnext()
    for e in paras[1:]:
        body.remove(e)
    P = Paragraph(paras[0], d._body)
    for r in list(P.runs):
        r._element.getparent().remove(r._element)
    para_fmt(P, 2, 2, 1.15, LA)
    set_run_font(P.add_run(f"Nội dung thực hiện theo báo giá số {info['quote']}. Vị trí thực hiện được chỉ định bởi đại diện "
                           f"{info['representative']}, số lượng vị trí cụ thể như sau:"), 13)

    agg = OrderedDict()
    for p, rd in data.items():
        for it in rd["items"]:
            a = agg.setdefault(it["ct"], dict(nhom=it["nhom"], kind=it["kind"], rooms=[]))
            if p not in a["rooms"]:
                a["rooms"].append(p)
    rows = sorted(agg.items(), key=lambda kv: GROUP_ORDER.index(kv[1]["nhom"]) if kv[1]["nhom"] in GROUP_ORDER else 99)

    tmpl_row = t2.rows[2]._tr
    for r_ in list(t2.rows)[2:]:
        t2._tbl.remove(r_._tr)
    all_codes = all(is_room_code(p) for p in data)
    for i, (ct, a) in enumerate(rows, 1):
        t2._tbl.append(copy.deepcopy(tmpl_row))
        row = t2.rows[-1]
        trpr = row._tr.get_or_add_trPr()
        if trpr.find(qn("w:cantSplit")) is None:
            trpr.append(OxmlElement("w:cantSplit"))
        is_sample = a["kind"] in ("sample", "micro")
        if ct in BBLM_LABEL:
            label = BBLM_LABEL[ct]
        elif a["kind"] == "ecg":
            label = "Đánh giá Ecgonomy (tư thế lao động OWAS & vị trí lao động)"
        elif a["kind"] == "assess":
            label = "Đánh giá yếu tố tiếp xúc nghề nghiệp"
        elif a["kind"] == "micro":
            label = f"Lấy mẫu vi sinh ({ct[0].lower() + ct[1:]}) trong không khí môi trường lao động"
        else:
            label = "Lấy mẫu " + (ct[0].lower() + ct[1:] if ct.lower().startswith("hơi") else "hơi " + ct) + " trong không khí môi trường lao động"
        if all(is_blank_loc(p) for p in a["rooms"]):
            where = "…" * 22
        elif len(data) > 1 and len(a["rooms"]) == len(data):
            where = "Các phòng" if all_codes else "Các khu vực"
        elif all_codes:
            where = "Phòng " + ", ".join(a["rooms"])
        else:
            where = ", ".join(a["rooms"])
        vals = [str(i), where, label, "" if is_sample else "x", "x" if is_sample else "", "……",
                "………" if is_sample else "-", "………" if is_sample else "-"]
        for j, (c, v) in enumerate(zip(ucells(row), vals)):
            set_cell(c, v, 12, align=LA if j in (1, 2) else CE, line=1.15 if j in (1, 2) else 1.5)
            valign(c, "center")

    W2 = [1.15, 4.15, 5.3, 0.9, 1.3, 1.6, 1.8, 1.8]
    for gc, wv in zip(t2._tbl.tblGrid.findall(qn("w:gridCol")), W2):
        gc.set(qn("w:w"), str(int(wv * 567)))
    for row in t2.rows:
        gi = 0
        for tc in row._tr.findall(qn("w:tc")):
            span = tc.tcPr.find(qn("w:gridSpan")) if tc.tcPr is not None else None
            n_ = int(span.get(qn("w:val"))) if span is not None else 1
            tcw = tc.get_or_add_tcPr().get_or_add_tcW()
            tcw.set(qn("w:w"), str(int(sum(W2[gi:gi + n_]) * 567))); tcw.set(qn("w:type"), "dxa")
            gi += n_
    for hr in t2.rows[:2]:
        trpr = hr._tr.get_or_add_trPr()
        for tag in ("w:tblHeader", "w:cantSplit"):
            if trpr.find(qn(tag)) is None:
                trpr.append(OxmlElement(tag))
        for c in hr.cells:
            for p_ in c.paragraphs:
                p_.paragraph_format.keep_with_next = True
    el = t2._tbl.getnext()
    while el is not None and el.tag == qn("w:p"):
        nx = el.getnext()
        if "Vị trí thực hiện được chỉ định" in "".join(x.text or "" for x in el.iter(qn("w:t"))):
            el.getparent().remove(el)
        el = nx

    notes = [info["condition_note"], f"Ký hiệu phiếu ghi chép: {info['doc']}/1;"]
    for i, row in enumerate(t3.rows):
        set_cell(ucells(row)[0], notes[i] if i < len(notes) else "", align=LA)
    return d


# ==============================================================================
# 6. HÀM CHÍNH
# ==============================================================================
def find_existing_z_folder(doc_number, z_base=None):
    """Quy ước: Z đã có thư mục cùng mã Q326D... thì dùng lại thư mục đó."""
    base = z_base or Z_DRIVE_DIR
    if not doc_number or doc_number == DOC_NUMBER_BLANK_PLACEHOLDER or not os.path.isdir(base):
        return None
    for name in os.listdir(base):
        if name.startswith(doc_number) and os.path.isdir(os.path.join(base, name)):
            return name
    return None


def generate_mtld(plan_path, customer, address, doc_number=None, short_customer=None, request=None,
                  quote=None, reg_date=None, representative=None, multi_day=False, plan_title=None,
                  condition_note=None, fixes=None, sheet=None, out_root=None, export_pdf=True,
                  copy_to_z=True, scan_files=None, z_base=None, year="2026"):
    from .pdf_exporter import convert_docx_to_pdf, repeat_pdf_copies

    blank = (not doc_number) or is_blank_doc_number(doc_number)
    doc = DOC_NUMBER_BLANK_PLACEHOLDER if blank else doc_number.strip()
    cust = normalize_customer_name(customer)
    addr = normalize_address(address)
    date_blank = f"…./…./{year}"
    data = build_room_data(load_plan(plan_path, sheet), fixes)
    info = dict(
        doc=doc, date_blank=date_blank, reg_date=reg_date or date_blank,
        customer=cust, address=addr, request=request or "Quan trắc môi trường lao động",
        place=LOCATION_DOTS + " - " + addr,
        time=(f"Bắt đầu từ ngày {date_blank} kết thúc ngày {date_blank}" if multi_day
              else f"Bắt đầu từ.................... kết thúc lúc..................... ngày {date_blank}"),
        equipment=equipment_lines(data), procedure=procedure_text(data),
        quote=quote or "QONV2……………", representative=representative or cust, plan_title=plan_title,
        condition_note=condition_note or "Tại thời điểm thực hiện, các khu vực hoạt động bình thường, có người lao động làm việc.",
    )

    folder = find_existing_z_folder(doc, z_base) if not blank else None
    if not folder:
        folder, short_kh, _ = build_output_folder_name(doc, cust, short_customer_name=short_customer, short_content="MTLĐ")
    short_kh = sanitize_for_filename(short_customer or build_output_folder_name(doc, cust)[1])
    out_dir = os.path.join(out_root or os.path.join(BASE_DIR, "output"), folder)
    os.makedirs(out_dir, exist_ok=True)
    dn = doc_number_for_filename(doc)
    f1 = os.path.join(out_dir, f"{dn} - BM3030.01 - Ghi chep - {short_kh}.docx")
    f2 = os.path.join(out_dir, f"{dn} - BM3030.02 - BBLM - {short_kh}.docx")
    build_bm01(data, info).save(f1)
    build_bm02(data, info).save(f2)
    res = dict(output_dir=out_dir, folder=folder, bm01_file=f1, bm02_file=f2,
               rooms=len(data), total=sum(it["sl"] for rd in data.values() for it in rd["items"] if not is_blank_sample(it)),
               equipment=info["equipment"], procedure=info["procedure"])
    if export_pdf:
        p1, n1 = convert_docx_to_pdf(f1)
        p2, n2 = convert_docx_to_pdf(f2)
        repeat_pdf_copies(p2)          # biên bản lập 2 bản -> PDF chứa sẵn 2 bản
        res.update(bm01_pdf=p1, bm01_pages=n1, bm02_pdf=p2, bm02_pages=n2)
    if copy_to_z:
        from .form_cloner import copy_order_to_z_drive
        res["z_dir"] = copy_order_to_z_drive(out_dir, folder, z_base, scan_files)
    return res
