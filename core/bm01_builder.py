"""
Module chuyên trách xây dựng và chuẩn hóa biểu mẫu BM3030.01
(Phiếu ghi chép hiện trường / Phiếu ghi chép giám định tại hiện trường)
Bao gồm các module con chuyên biệt cho từng lĩnh vực:
- BSC: Tủ an toàn sinh học (Trang 1 bảng biểu phòng sạch + nối Ghi Chep BSC.pdf)
- Khí nén: Co giãn 1 hoặc nhiều vị trí (Vị trí 01, 02, 03...), có/không có vi sinh
- Phòng sạch: Bảng đo hạt bụi chuẩn + các bảng đo thứ cấp (Gió, Sáng, Khói, Vi sinh)
- Hybrid: Kết hợp Khí nén + Phòng sạch hoặc các lĩnh vực phức hợp
- MTLĐ & Kho lạnh
Tuân thủ 100% 11 Quy tắc soạn thảo chuẩn mực (GEMINI.md).
"""
import os
import re
import copy
import docx
import fitz
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

from .config import BASE_DIR, CATEGORIES, infer_method_type
from .normalizer import (
    normalize_customer_name,
    normalize_date_str,
    normalize_address,
    normalize_location,
    clean_indicator_name,
    format_equipment_multiline,
    clean_text_no_status_or_quantitative,
    get_default_execution_date_dotted
)
from .docx_processor import (
    set_cell_text_preserve_formatting,
    update_header_hl,
    clean_spaces_before_colon,
    finalize_document_formatting,
    is_bm01_measurement_table
)


def add_6pt_spacer_after(element, doc):
    """Thêm đoạn trắng cách bảng cỡ 6pt theo Quy tắc 3."""
    p_xml = parse_xml(f'''
        <w:p {nsdecls("w")}>
            <w:pPr>
                <w:spacing w:before="0" w:after="0" w:line="120" w:lineRule="exact"/>
                <w:rPr>
                    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
                    <w:sz w:val="12"/>
                    <w:szCs w:val="12"/>
                </w:rPr>
            </w:pPr>
        </w:p>
    ''')
    element.addnext(p_xml)


def ensure_single_signature_at_end(doc):
    """
    [NEW 28/10/2026] Quy tắc: 'Người thực hiện' chỉ ký 1 LẦN DUY NHẤT ở cuối ghi chép BM3030.01.
    - Xóa mọi đoạn 'Người thực hiện' nằm rải rác dưới từng bảng đo (body-level).
    - Đặt lại đúng 1 đoạn chữ ký ở cuối tài liệu (sau bảng đo cuối cùng, trước sectPr).
      Giữ định dạng đoạn chữ ký gốc của mẫu nếu có, nếu không thì tạo mới theo mẫu Phòng sạch
      (căn phải, đậm, 12pt, thụt phải 1133 dxa).
    - Đưa sectPr của body về cuối cùng (một số bảng thứ cấp được append sau sectPr).
    """
    body = doc.element.body
    sig_paras = []
    for el in list(body.iterchildren()):
        if el.tag == qn('w:p'):
            txt = ''.join(t.text or '' for t in el.iter(qn('w:t'))).strip()
            if txt == 'Người thực hiện':
                sig_paras.append(el)

    if sig_paras:
        sig_el = sig_paras[-1]
    else:
        sig_el = parse_xml(
            f'<w:p {nsdecls("w")}><w:pPr><w:tabs><w:tab w:val="left" w:pos="2835"/></w:tabs>'
            f'<w:spacing w:before="80" w:after="0"/><w:ind w:right="1133"/><w:contextualSpacing/><w:jc w:val="right"/>'
            f'<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
            f'<w:t>Người thực hiện</w:t></w:r></w:p>'
        )
    for el in sig_paras:
        body.remove(el)

    sectPr = body.find(qn('w:sectPr'))
    if sectPr is not None:
        body.remove(sectPr)
    body.append(sig_el)
    if sectPr is not None:
        body.append(sectPr)
    return doc


# ==============================================================================
# 1. HELPER SOẠN THẢO TABLE 0 & TABLE 1 TRÊN TRANG 1 CHO BM3030.01
# ==============================================================================

def fill_bm01_header_and_table1(doc, customer_info):
    """
    Điền chuẩn Table 0 và Table 1 cho Trang đầu của BM3030.01 theo đúng Quy tắc.
    """
    doc_number = customer_info.get("doc_number", "")
    exec_date = customer_info.get("execution_date") or get_default_execution_date_dotted()
    customer_name = normalize_customer_name(customer_info.get("customer_name", ""))
    address = normalize_address(customer_info.get("address", ""))
    location = normalize_location(customer_info.get("location", ""), address)
    request_content = clean_text_no_status_or_quantitative(customer_info.get("request_content", ""))
    time_text = f"Bắt đầu từ........... kết thúc lúc........... ngày {exec_date}"

    procedure_text = customer_info.get("procedure", "") or customer_info.get("standard_procedure", "QUATEST 3 3030:2025")
    raw_equipment = customer_info.get("equipment", "")
    equipment_formatted = format_equipment_multiline(raw_equipment)
    eq_lines = [line.strip() for line in equipment_formatted.split("\n") if line.strip()]

    # Table 0: Số và Ngày hiện trường
    t0 = doc.tables[0]
    set_cell_text_preserve_formatting(t0.rows[0].cells[1], doc_number, bold=False, font_size=Pt(13))
    set_cell_text_preserve_formatting(t0.rows[0].cells[4], exec_date, bold=False, font_size=Pt(13))

    # Căn giữa dấu ':' Table 0 nếu có
    for r in t0.rows:
        if len(r.cells) >= 3 and r.cells[1].text.strip() == ":":
            pPr = r.cells[1].paragraphs[0]._p.get_or_add_pPr()
            jc = pPr.find(qn('w:jc'))
            if jc is not None:
                jc.set(qn('w:val'), 'center')
            else:
                pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))

    # Table 1: Thông tin chung
    t1 = doc.tables[1]
    for cell in t1.rows[0].cells:
        cell.text = ""

    # Row 1: Đơn vị yêu cầu
    set_cell_text_preserve_formatting(t1.rows[1].cells[2], customer_name, bold=False, font_size=Pt(13))
    # Row 2: Nội dung yêu cầu
    set_cell_text_preserve_formatting(t1.rows[2].cells[2], request_content, bold=False, font_size=Pt(13))
    # Row 3: Địa điểm thực hiện
    set_cell_text_preserve_formatting(t1.rows[3].cells[2], location, bold=False, font_size=Pt(13))
    # Row 4: Thời gian thực hiện
    set_cell_text_preserve_formatting(t1.rows[4].cells[2], time_text, bold=False, font_size=Pt(13))
    t1.rows[4].cells[2].paragraphs[0].paragraph_format.left_indent = Inches(0)
    t1.rows[4].cells[2].paragraphs[0].paragraph_format.right_indent = Inches(0)

    # Row 6: Thiết bị sử dụng
    c_eq = t1.rows[6].cells[0] if len(t1.rows[6].cells) == 1 else (t1.rows[6].cells[2] if len(t1.rows[6].cells) >= 3 else t1.rows[6].cells[-1])
    for p in list(c_eq.paragraphs):
        p._p.getparent().remove(p._p)
    for eq_line in eq_lines:
        p = c_eq.add_paragraph()
        p.paragraph_format.left_indent = Inches(0)
        p.paragraph_format.right_indent = Inches(0)
        p.paragraph_format.first_line_indent = Inches(0)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(eq_line)
        r.font.name = "Times New Roman"
        r._r.get_or_add_rPr().set(qn('w:rFonts'), "Times New Roman")
        r.font.size = Pt(13)
        r.bold = False

    # Row 7: Quy trình thực hiện
    set_cell_text_preserve_formatting(t1.rows[7].cells[2], procedure_text, bold=False, font_size=Pt(13))

    # Row 8: Nội dung/Quá trình thực hiện (để trống per standard)
    if len(t1.rows) > 8:
        set_cell_text_preserve_formatting(t1.rows[8].cells[2], "", bold=False, font_size=Pt(13))

    add_6pt_spacer_after(t1._tbl, doc)


# ==============================================================================
# 2. MODULE BM01: TỦ AN TOÀN SINH HỌC (BSC)
# ==============================================================================

def render_bsc_pdf_pages_to_images(pdf_path, dpi=300):
    """Render các trang của Ghi Chep BSC.pdf sang ảnh PNG 300dpi phục vụ nhúng Word."""
    doc_pdf = fitz.open(pdf_path)
    scratch_imgs_dir = os.path.join(BASE_DIR, "scratch", f"bsc_imgs_{dpi}")
    os.makedirs(scratch_imgs_dir, exist_ok=True)
    images = []
    for i, page in enumerate(doc_pdf):
        img_out = os.path.join(scratch_imgs_dir, f"page_{i+1}.png")
        if not os.path.exists(img_out):
            pix = page.get_pixmap(dpi=dpi)
            pix.save(img_out)
        images.append(img_out)
    return images


def _calc_bm01_p1_dot_lines(customer_info):
    """Tính số dòng chấm lửng (147 dấu '.') lấp đầy Trang 1 của BM3030.01 theo số dòng thiết bị."""
    eq_formatted = format_equipment_multiline(customer_info.get("equipment", ""))
    eq_count = len([l for l in eq_formatted.split("\n") if l.strip()])
    if eq_count >= 8:
        return 10
    if eq_count == 7:
        return 11
    if eq_count in (5, 6):
        return 12
    if eq_count in (3, 4):
        return 14
    return 15


def build_bm01_bsc(customer_info, num_positions=1, output_path=None):
    """
    Xây dựng BM3030.01 cho Tủ an toàn sinh học (BSC):
    - Trang 1: Bảng biểu theo phòng sạch + các hàng chấm lấp đầy trang 1 để KTV điền tay.
    - Trang 2+: Nối 3 trang bảng đo BSC từ Ghi Chep BSC.pdf cho mỗi tủ (num_positions tủ -> 3*N trang).
    """
    template_ps = os.path.join(BASE_DIR, "Phong sach, Vi sinh", "BM3030.01- Ghi chép - 116326.docx")
    doc = docx.Document(template_ps)

    # Cập nhật Header HL:01/08/2025
    update_header_hl(doc, "HL:01/08/2025")

    # Xóa các bảng đo phía sau Table 1 (Table 2, 3, 4, 5...)
    while len(doc.tables) > 2:
        doc.tables[-1]._tbl.getparent().remove(doc.tables[-1]._tbl)

    # Xóa các paragraph thừa trong body
    for p in list(doc.paragraphs):
        p._p.getparent().remove(p._p)

    # Điền Table 0 và Table 1
    fill_bm01_header_and_table1(doc, customer_info)

    # Thêm hàng chấm lửng tràn hết lề layout lấp đầy khoảng trắng Trang 1 (Quy tắc 6)
    dot_line = "." * 147
    num_dots_p1 = _calc_bm01_p1_dot_lines(customer_info)
    for _ in range(num_dots_p1):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(dot_line)
        r.font.name = "Times New Roman"
        r._r.get_or_add_rPr().set(qn('w:rFonts'), "Times New Roman")
        r.font.size = Pt(13)

    # Nối Ghi Chep BSC.pdf vào ghi chép: Section mới cho 3 trang bảng đo × số lượng tủ
    pdf_bsc_source = os.path.join(BASE_DIR, "BSC", "Ghi Chep BSC.pdf")
    pdf_images = render_bsc_pdf_pages_to_images(pdf_bsc_source, dpi=300)

    sec2 = doc.add_section(WD_SECTION_START.NEW_PAGE)
    sec2.top_margin = Inches(0.4)
    sec2.bottom_margin = Inches(0.4)
    sec2.left_margin = Inches(0.4)
    sec2.right_margin = Inches(0.4)
    sec2.header.is_linked_to_previous = False
    for t in list(sec2.header.tables):
        t._tbl.getparent().remove(t._tbl)
    for p in list(sec2.header.paragraphs):
        p.text = ""

    pos_count = max(1, int(num_positions or customer_info.get("num_positions") or 1))
    page_idx = 0
    for _cab in range(pos_count):
        for img_path in pdf_images:
            if page_idx > 0:
                doc.add_page_break()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run()
            r.add_picture(img_path, width=Inches(7.4))
            page_idx += 1

    finalize_document_formatting(doc)

    if output_path:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        doc.save(output_path)

    return doc


def prepare_bsc_temp_p1(customer_info, output_path):
    """
    Soạn riêng file docx Trang 1 của BM3030.01 phục vụ xuất PDF vector chất lượng cao
    sau đó ghép với 3 trang vector gốc của Ghi Chep BSC.pdf (nhân theo số lượng tủ).
    """
    template_ps = os.path.join(BASE_DIR, "Phong sach, Vi sinh", "BM3030.01- Ghi chép - 116326.docx")
    doc_p1 = docx.Document(template_ps)
    while len(doc_p1.tables) > 2:
        doc_p1.tables[-1]._tbl.getparent().remove(doc_p1.tables[-1]._tbl)
    for p in list(doc_p1.paragraphs):
        p._p.getparent().remove(p._p)

    update_header_hl(doc_p1, "HL:01/08/2025")
    fill_bm01_header_and_table1(doc_p1, customer_info)

    dot_line = "." * 147
    num_dots_p1 = _calc_bm01_p1_dot_lines(customer_info)
    for _ in range(num_dots_p1):
        p = doc_p1.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(dot_line)
        r.font.name = "Times New Roman"
        r._r.get_or_add_rPr().set(qn('w:rFonts'), "Times New Roman")
        r.font.size = Pt(13)

    finalize_document_formatting(doc_p1)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc_p1.save(output_path)
    return output_path


# ==============================================================================
# 3. MODULE BM01: KHÍ NÉN & KHÍ CÓ ÁP (NHIỀU VỊ TRÍ)
# ==============================================================================

def prepare_single_khi_nen_table_element(src_tbl, has_micro=False):
    """
    Chuẩn hóa element bảng Khí nén từ template Q326D122266:
    - Nếu không có vi sinh: xóa hàng vi sinh (hàng index 3).
    - Thay thế HDTN... và MTTN... bằng chuỗi chấm lửng '………………………………'.
    - Format Row 0 Cell 0 gồm đúng 3 paragraphs chấm lửng font 11 italic.
    - Áp dụng fixed layout và cantSplit.
    """
    tbl_elem = copy.deepcopy(src_tbl._tbl)

    # Xóa hàng Vi sinh nếu không yêu cầu vi sinh
    tr_list = tbl_elem.findall(qn('w:tr'))
    if not has_micro and len(tr_list) == 5:
        tbl_elem.remove(tr_list[3])

    # Thay thế HDTN và MTTN
    for tr in tbl_elem.findall(qn('w:tr')):
        trPr = tr.get_or_add_trPr()
        if trPr.find(qn('w:cantSplit')) is None:
            trPr.append(OxmlElement('w:cantSplit'))
        for tc in tr.findall(qn('w:tc')):
            for p in tc.findall(qn('w:p')):
                for r in p.findall(qn('w:r')):
                    for t in r.findall(qn('w:t')):
                        if t.text:
                            t.text = re.sub(r'HDTN\w+', '………………………………', t.text)
                            t.text = re.sub(r'MTTN\w+', '………………………………', t.text)

    # Format Row 0 Cell 0 gồm đúng 3 paragraphs per Quy tắc 6
    tr0 = tbl_elem.findall(qn('w:tr'))[0]
    tc0 = tr0.findall(qn('w:tc'))[0]
    for p in list(tc0.findall(qn('w:p'))):
        tc0.remove(p)

    for p_txt in ["   Tên khu vực thực hiện: ………", "   …………………………………", "   …………………………………"]:
        p_new = parse_xml(
            f'<w:p {nsdecls("w")}><w:pPr><w:spacing w:before="40" w:after="40"/><w:rPr>'
            f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="22"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="22"/></w:rPr>'
            f'<w:t>{p_txt}</w:t></w:r></w:p>'
        )
        tc0.append(p_new)

    return tbl_elem


def build_bm01_khi_nen(customer_info, num_positions=1, indicators=None, output_path=None):
    """
    Xây dựng BM3030.01 cho Khí nén:
    - Co giãn đúng số lượng bảng tương ứng với số vị trí yêu cầu (Vị trí 01, Vị trí 02...).
    - Khóa fixed layout, font Times New Roman 13pt bold lề trái cho tiêu đề bảng đo.
    - [FIX 28/10/2026] Chữ ký 'Người thực hiện' chỉ 1 lần ở cuối ghi chép (không đặt dưới từng bảng đo).
    """
    template_kn = os.path.join(BASE_DIR, "Khi nen", "BM3030.01- Ghi chép - Q326D122266.docx")
    doc = docx.Document(template_kn)

    # Cập nhật Header HL:01/08/2025
    update_header_hl(doc, "HL:01/08/2025")

    # Cập nhật thông tin khách hàng vào Table 0
    from .docx_processor import update_customer_info
    update_customer_info(doc, customer_info, indicators=indicators)

    # Kiểm tra xem có vi sinh khí nén hay không
    combined_str = f"{customer_info.get('request_content', '')} {' '.join(str(i) for i in (indicators or []))}".lower()
    has_micro = any(k in combined_str for k in ["vi sinh", "micro", "vsv", "khuẩn", "nấm"])

    # Xác định số vị trí đo
    pos_count = max(1, int(num_positions))

    # Co giãn số lượng bảng đo trong tài liệu:
    # Table 0: Thông tin khách hàng
    # Table 1, 2, 3...: Các bảng đo
    while len(doc.tables) - 1 > pos_count and len(doc.tables) > 2:
        doc.tables[-1]._tbl.getparent().remove(doc.tables[-1]._tbl)

    while len(doc.tables) - 1 < pos_count:
        last_tbl = doc.tables[-1]
        p_title = doc.add_paragraph()
        p_title.text = f"Khí nén (Vị trí {len(doc.tables):02d}):"
        p_title.runs[0].font.name = "Times New Roman"
        p_title.runs[0].font.size = Pt(13)
        p_title.runs[0].bold = True
        p_title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        new_tbl = copy.deepcopy(last_tbl._tbl)
        doc.element.body.append(new_tbl)
        add_6pt_spacer_after(new_tbl, doc)

    # [NEW 11/11/2026] Chỉ giữ khối chỉ tiêu được yêu cầu (Tổng dầu / Particle / Vi sinh / Điểm sương).
    # Không nhận diện được chỉ tiêu nào -> giữ nguyên như cũ (chỉ bỏ Vi sinh nếu không yêu cầu).
    ind_str = " ".join(str(i) for i in (indicators or [])).lower()
    wanted = {
        "Tổng dầu": any(k in ind_str for k in ["dầu", "oil"]),
        "Particle": any(k in ind_str for k in ["bụi", "hạt", "particle"]),
        "Điểm sương": any(k in ind_str for k in ["điểm sương", "dew"]),
    }
    prune_by_indicator = any(wanted.values())

    # Chuẩn hóa từng bảng đo
    for tbl_idx in range(1, len(doc.tables)):
        tbl = doc.tables[tbl_idx]
        tr_list = tbl._tbl.findall(qn('w:tr'))
        if not has_micro and len(tr_list) == 5:
            tbl._tbl.remove(tr_list[3]) # Xóa hàng Vi sinh
        if prune_by_indicator:
            restart_tc_props = None
            for tr in tbl._tbl.findall(qn('w:tr'))[1:]:
                tcs = tr.findall(qn('w:tc'))
                if len(tcs) < 2:
                    continue
                block = "".join(tcs[1].itertext()).strip()
                key = next((k for k in wanted if block.startswith(k)), None)
                if key is None or wanted[key]:
                    continue
                vm = tcs[0].find('.//' + qn('w:vMerge'))
                if vm is not None and vm.get(qn('w:val')) == 'restart':
                    # Hàng bị xóa đang giữ ô gộp 'CHỈ TIÊU THỰC HIỆN' -> chuyển ô đó sang hàng kế tiếp
                    nxt = tr.getnext()
                    if nxt is not None and nxt.tag == qn('w:tr'):
                        nxt_tcs = nxt.findall(qn('w:tc'))
                        if nxt_tcs:
                            nxt.replace(nxt_tcs[0], copy.deepcopy(tcs[0]))
                tbl._tbl.remove(tr)

        for r in tbl.rows:
            trPr = r._tr.get_or_add_trPr()
            if trPr.find(qn('w:cantSplit')) is None:
                trPr.append(OxmlElement('w:cantSplit'))
            for c in r.cells:
                if 'HDTN' in c.text or 'MTTN' in c.text:
                    t = re.sub(r'HDTN\w+', '………………………………', c.text)
                    t = re.sub(r'MTTN\w+', '………………………………', t)
                    set_cell_text_preserve_formatting(c, t, bold=False, font_size=Pt(13))

        # Format Row 0 Cell 0 gồm đúng 3 paragraphs font 11 italic
        if len(tbl.rows) > 0 and len(tbl.rows[0].cells) > 0:
            c0 = tbl.rows[0].cells[0]
            for p in list(c0.paragraphs):
                p._p.getparent().remove(p._p)
            for p_txt in ["   Tên khu vực thực hiện: ………", "   …………………………………", "   …………………………………"]:
                p_new = c0.add_paragraph()
                p_new.paragraph_format.space_before = Pt(2)
                p_new.paragraph_format.space_after = Pt(2)
                r = p_new.add_run(p_txt)
                r.italic = True
                r.font.name = "Times New Roman"
                r._r.get_or_add_rPr().set(qn('w:rFonts'), "Times New Roman")
                r.font.size = Pt(11)

    # Format tiêu đề các bảng đo 'Khí nén:' hoặc 'Khí nén (Vị trí XX):'
    heading_count = 0
    for p in doc.paragraphs:
        if "khí nén" in p.text.lower():
            heading_count += 1
            if pos_count > 1:
                p.text = f"Khí nén (Vị trí {heading_count:02d}):"
            else:
                p.text = "Khí nén:"
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.left_indent = Inches(0)
            p.paragraph_format.right_indent = Inches(0)
            p.paragraph_format.first_line_indent = Inches(0)
            for r in p.runs:
                r.bold = True
                r.font.name = "Times New Roman"
                r._r.get_or_add_rPr().set(qn('w:rFonts'), "Times New Roman")
                r.font.size = Pt(13)

    finalize_document_formatting(doc)
    ensure_single_signature_at_end(doc)  # [FIX 28/10/2026] ký 1 lần ở cuối ghi chép

    if output_path:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        doc.save(output_path)

    return doc


# ==============================================================================
# 4. MODULE BM01: PHÒNG SẠCH & BẢNG ĐO THỨ CẤP DYNAMIC
# ==============================================================================

SECONDARY_MEASUREMENT_CONFIGS = {
    'air_velocity': {
        'title': 'Tốc độ gió',
        'equipment': 'Kanomax Climomaster 6501, Japan',
        'loc_label': 'Mặt cấp gió',
        'col3_header': 'Giá trị đo (m/s)',
        'col3_unit': '…………… m/s'
    },
    'light': {
        'title': 'Cường độ sáng',
        'equipment': 'Testo 540 / Testo 440, Germany',
        'loc_label': 'Mặt làm việc',
        'col3_header': 'Giá trị đo (lux)',
        'col3_unit': '…………… lux'
    },
    'airflow_pattern': {
        'title': 'Hướng dòng khí',
        'equipment': 'Máy tạo khói khí dung (Smoke pencil / Aerosol generator)',
        'loc_label': 'Vùng làm việc',
        'col3_header': 'Hình thái dòng khí',
        'col3_unit': '[  ] Đạt     [  ] Không đạt'
    },
    'vskk': {
        'title': 'Lấy mẫu vi sinh trong không khí',
        'equipment': 'Buck – Bio Culture B30120, USA',
        'loc_label': 'Không khí phòng sạch',
        'col3_header': 'Thể tích lấy mẫu (L)',
        'col3_unit': '…………… L'
    },
    'hepa_leak': {
        'title': 'Quét rò rỉ màng lọc HEPA',
        'equipment': 'ATI aerosol generator 6D và aerosol photometer 2i, USA',
        'loc_label': 'Màng lọc HEPA',
        'col3_header': 'Độ xuyên thấu (%)',
        'col3_unit': '…………… %'
    },
    'swab': {
        'title': 'Lấy mẫu vi sinh bề mặt (Swab)',
        'equipment': 'Dụng cụ lấy mẫu vi sinh bề mặt chuyên dụng',
        'loc_label': 'Bề mặt tiếp xúc',
        'col3_header': 'Diện tích lấy mẫu',
        'col3_unit': '10cm * 10cm'
    },
    # [14/11/2026] Vi sinh không khí phương pháp đặt đĩa (lắng, tối đa 4 giờ - N2.TC.09)
    'vskk_settle': {
        'title': 'Lấy mẫu vi sinh trong không khí (đặt đĩa)',
        'equipment': 'Khay inox, đĩa môi trường',
        'loc_label': 'Không khí',
        'col3_header': 'Thời gian đặt đĩa',
        'col3_unit': 'Từ ………… đến …………'
    },
    # [14/11/2026] Vi sinh bề mặt phương pháp đĩa tiếp xúc (Rodac Ø55 mm, ~25 cm2 - N2.TC.09)
    'contact_plate': {
        'title': 'Lấy mẫu vi sinh bề mặt (đĩa tiếp xúc)',
        'equipment': 'Đĩa tiếp xúc Rodac',
        'loc_label': 'Bề mặt tiếp xúc',
        'col3_header': 'Diện tích lấy mẫu',
        'col3_unit': '~ 25 cm²'
    },
    # [14/11/2026] Chênh áp qua màng lọc khí cấp HEPA
    'filter_dp': {
        'title': 'Chênh áp màng lọc HEPA',
        'equipment': 'TSI 8375, USA',
        'loc_label': 'Màng lọc HEPA',
        'col3_header': 'Chênh áp (Pa)',
        'col3_unit': '…………… Pa'
    }
}


def build_secondary_measurement_table(ref_table_2, measure_type, num_rows=8, custom_col3_unit=None):
    """
    Tạo bảng đo thứ cấp 4 cột theo đúng chuẩn format header của Bảng đo bụi (Table 2):
    - Khóa fixed layout w:type="fixed".
    - Row 0: Tên khu vực (span 2, 3227 dxa) & Tình trạng HVAC (span 2, 6368 dxa).
    - Row 1: Left trống & Right thông tin phép đo (Tên, Thiết bị, Thời gian, Vị trí...).
    - Row 2: Tiêu đề STT (800 dxa), Vị trí đo (5227 dxa), Kết quả (3568 dxa).
    - Row 3+: Hàng dữ liệu kèm dòng chấm không rớt dòng.
    """
    cfg = SECONDARY_MEASUREMENT_CONFIGS[measure_type]

    tbl_xml = f'''
    <w:tbl {nsdecls("w")}>
        <w:tblPr>
            <w:tblW w:w="9595" w:type="dxa"/>
            <w:tblLayout w:type="fixed"/>
            <w:jc w:val="center"/>
            <w:tblBorders>
                <w:top w:val="single" w:sz="12" w:space="0" w:color="auto"/>
                <w:left w:val="single" w:sz="12" w:space="0" w:color="auto"/>
                <w:bottom w:val="single" w:sz="12" w:space="0" w:color="auto"/>
                <w:right w:val="single" w:sz="12" w:space="0" w:color="auto"/>
                <w:insideH w:val="single" w:sz="12" w:space="0" w:color="auto"/>
                <w:insideV w:val="single" w:sz="12" w:space="0" w:color="auto"/>
            </w:tblBorders>
        </w:tblPr>
        <w:tblGrid>
            <w:gridCol w:w="800"/>
            <w:gridCol w:w="2427"/>
            <w:gridCol w:w="2800"/>
            <w:gridCol w:w="3568"/>
        </w:tblGrid>
    </w:tbl>
    '''
    new_tbl = parse_xml(tbl_xml)

    # Hàng 0: Khu vực & HVAC
    r0 = parse_xml(f'<w:tr {nsdecls("w")}><w:trPr><w:cantSplit/></w:trPr></w:tr>')
    c0 = copy.deepcopy(ref_table_2.rows[0].cells[0]._tc)
    tcPr0 = c0.get_or_add_tcPr()
    gs0 = tcPr0.find(qn('w:gridSpan'))
    if gs0 is None: tcPr0.append(parse_xml(f'<w:gridSpan {nsdecls("w")} w:val="2"/>'))
    else: gs0.set(qn('w:val'), '2')
    tcW0 = tcPr0.find(qn('w:tcW'))
    if tcW0 is not None: tcW0.set(qn('w:w'), '3227')
    r0.append(c0)

    c1 = copy.deepcopy(ref_table_2.rows[0].cells[1]._tc)
    tcPr1 = c1.get_or_add_tcPr()
    gs1 = tcPr1.find(qn('w:gridSpan'))
    if gs1 is None: tcPr1.append(parse_xml(f'<w:gridSpan {nsdecls("w")} w:val="2"/>'))
    else: gs1.set(qn('w:val'), '2')
    tcW1 = tcPr1.find(qn('w:tcW'))
    if tcW1 is not None: tcW1.set(qn('w:w'), '6368')
    r0.append(c1)
    new_tbl.append(r0)

    # Hàng 1: Trống & Thông tin phép đo
    r1 = parse_xml(f'<w:tr {nsdecls("w")}><w:trPr><w:cantSplit/></w:trPr></w:tr>')
    c0_1 = copy.deepcopy(ref_table_2.rows[1].cells[0]._tc)
    tcPr0_1 = c0_1.get_or_add_tcPr()
    gs0_1 = tcPr0_1.find(qn('w:gridSpan'))
    if gs0_1 is None: tcPr0_1.append(parse_xml(f'<w:gridSpan {nsdecls("w")} w:val="2"/>'))
    else: gs0_1.set(qn('w:val'), '2')
    tcW0_1 = tcPr0_1.find(qn('w:tcW'))
    if tcW0_1 is not None: tcW0_1.set(qn('w:w'), '3227')
    r1.append(c0_1)

    c1_1 = copy.deepcopy(ref_table_2.rows[1].cells[1]._tc)
    tcPr1_1 = c1_1.get_or_add_tcPr()
    gs1_1 = tcPr1_1.find(qn('w:gridSpan'))
    if gs1_1 is None: tcPr1_1.append(parse_xml(f'<w:gridSpan {nsdecls("w")} w:val="2"/>'))
    else: gs1_1.set(qn('w:val'), '2')
    tcW1_1 = tcPr1_1.find(qn('w:tcW'))
    if tcW1_1 is not None: tcW1_1.set(qn('w:w'), '6368')

    paras = c1_1.findall(qn('w:p'))
    if len(paras) > 0:
        p0 = paras[0]
        for r in list(p0.findall(qn('w:r'))): p0.remove(r)
        p0.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="24"/></w:rPr><w:t>{cfg["title"]}</w:t></w:r>'))

    if len(paras) > 1:
        p1 = paras[1]
        for r in list(p1.findall(qn('w:r'))): p1.remove(r)
        p1.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">+ Thiết bị sử dụng:  {cfg["equipment"]}</w:t></w:r>'))

    if len(paras) > 2:
        p2 = paras[2]
        for r in list(p2.findall(qn('w:r'))): p2.remove(r)
        p2.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr><w:t>+ Thời gian thực hiện: ………………………………</w:t></w:r>'))

    if len(paras) > 3:
        p3 = paras[3]
        for r in list(p3.findall(qn('w:r'))): p3.remove(r)
        p3.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">+ Vị trí kiểm tra:  </w:t></w:r>'))
        p3.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:sz w:val="22"/></w:rPr><w:sym w:font="Wingdings" w:char="F0A8"/></w:r>'))
        p3.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve"> {cfg["loc_label"]}   </w:t></w:r>'))
        p3.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:sz w:val="22"/></w:rPr><w:sym w:font="Wingdings" w:char="F0A8"/></w:r>'))
        p3.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve"> khác: ……………………</w:t></w:r>'))

    if len(paras) > 4:
        p4 = paras[4]
        for r in list(p4.findall(qn('w:r'))): p4.remove(r)
        p4.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr><w:t>+ Kết quả đo: ghi nhận theo bảng bên dưới</w:t></w:r>'))

    while len(paras) > 5:
        c1_1.remove(paras[-1])
        paras = c1_1.findall(qn('w:p'))

    r1.append(c1_1)
    new_tbl.append(r1)

    # Hàng 2: Header 3 cột (STT, Vị trí đo, Kết quả)
    r2_xml = f'''
    <w:tr {nsdecls("w")}>
        <w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>
        <w:tc>
            <w:tcPr><w:tcW w:w="800" w:type="dxa"/><w:gridSpan w:val="1"/><w:vAlign w:val="center"/></w:tcPr>
            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="40" w:after="40"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/></w:rPr><w:t>STT</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
            <w:tcPr><w:tcW w:w="5227" w:type="dxa"/><w:gridSpan w:val="2"/><w:vAlign w:val="center"/></w:tcPr>
            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="40" w:after="40"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/></w:rPr><w:t>Vị trí đo</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
            <w:tcPr><w:tcW w:w="3568" w:type="dxa"/><w:gridSpan w:val="1"/><w:vAlign w:val="center"/></w:tcPr>
            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="40" w:after="40"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/></w:rPr><w:t>{cfg["col3_header"]}</w:t></w:r></w:p>
        </w:tc>
    </w:tr>
    '''
    new_tbl.append(parse_xml(r2_xml))

    # Hàng 3+: Dữ liệu
    dots_loc = "………………………………………………"
    col3_unit_val = custom_col3_unit if custom_col3_unit else cfg["col3_unit"]
    for idx in range(1, num_rows + 1):
        row_xml = f'''
        <w:tr {nsdecls("w")}>
            <w:trPr><w:cantSplit/><w:trHeight w:val="454" w:hRule="atLeast"/></w:trPr>
            <w:tc>
                <w:tcPr><w:tcW w:w="800" w:type="dxa"/><w:gridSpan w:val="1"/><w:vAlign w:val="center"/></w:tcPr>
                <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="40" w:after="40"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/></w:rPr><w:t>{idx:02d}</w:t></w:r></w:p>
            </w:tc>
            <w:tc>
                <w:tcPr><w:tcW w:w="5227" w:type="dxa"/><w:gridSpan w:val="2"/><w:vAlign w:val="center"/></w:tcPr>
                <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="40" w:after="40"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/></w:rPr><w:t>{dots_loc}</w:t></w:r></w:p>
            </w:tc>
            <w:tc>
                <w:tcPr><w:tcW w:w="3568" w:type="dxa"/><w:gridSpan w:val="1"/><w:vAlign w:val="center"/></w:tcPr>
                <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="40" w:after="40"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/></w:rPr><w:t>{col3_unit_val}</w:t></w:r></w:p>
            </w:tc>
        </w:tr>
        '''
        new_tbl.append(parse_xml(row_xml))

    # [14/11/2026] Giữ trọn bảng trên 1 trang (keepNext cho mọi hàng trừ hàng cuối),
    # không để bảng bị cắt đôi qua 2 trang.
    all_rows = new_tbl.findall(qn('w:tr'))
    for tr in all_rows[:-1]:
        for p in tr.iter(qn('w:p')):
            pPr = p.find(qn('w:pPr'))
            if pPr is None:
                pPr = OxmlElement('w:pPr')
                p.insert(0, pPr)
            if pPr.find(qn('w:keepNext')) is None:
                kn = OxmlElement('w:keepNext')
                pstyle = pPr.find(qn('w:pStyle'))
                if pstyle is not None:
                    pstyle.addnext(kn)
                else:
                    pPr.insert(0, kn)

    return new_tbl


def prepare_original_vskk_table_element(src_vskk_tbl_el, ref_dust_tbl, num_rows=8):
    """
    Chuẩn hóa bảng ghi chép Vi sinh gốc (Table 5 của mẫu BM3030.01 Phòng sạch, Vi sinh):
    - Giữ nguyên 100% checkbox Wingdings (APBuck, SKC-Biostage, TS lấy mẫu / TS blank, MM lấy mẫu / MM blank, Khác)
      và số thứ tự tự động numId='36' của mẫu QUATEST 3 gốc.
    - Khóa fixed layout (9595 dxa = 3228 + 6367 dxa) để cột HVAC không bị rớt dòng.
    - Chuẩn hóa Hàng 0 Cột 0 (Tên khu vực thực hiện: 3 dòng in nghiêng 11pt), xóa dòng trống thừa ở Hàng 0 Cột 1,
      xóa chữ 'CHỈ TIÊU THỰC HIỆN' ở Hàng 1 Cột 0, rút gọn chấm lửng '+ Lưu lượng lấy mẫu:' nằm gọn 1 dòng.
    - Co giãn số hàng dữ liệu theo num_rows (số vị trí vi sinh + 3 dòng dự phòng).
    """
    tbl_el = copy.deepcopy(src_vskk_tbl_el)

    tblPr = tbl_el.find(qn('w:tblPr'))
    if tblPr is not None:
        tblW = tblPr.find(qn('w:tblW'))
        if tblW is None:
            tblPr.insert(0, parse_xml(f'<w:tblW {nsdecls("w")} w:w="9595" w:type="dxa"/>'))
        else:
            tblW.set(qn('w:w'), '9595')
            tblW.set(qn('w:type'), 'dxa')
        if tblPr.find(qn('w:tblLayout')) is None:
            tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))

    trs = tbl_el.findall(qn('w:tr'))
    for tr in trs:
        tcs = tr.findall(qn('w:tc'))
        for ci, tc in enumerate(tcs):
            tcPr = tc.get_or_add_tcPr()
            tcW = tcPr.find(qn('w:tcW'))
            target_w = '3228' if ci == 0 else '6367'
            if tcW is None:
                tcPr.insert(0, parse_xml(f'<w:tcW {nsdecls("w")} w:w="{target_w}" w:type="dxa"/>'))
            else:
                tcW.set(qn('w:w'), target_w)
                tcW.set(qn('w:type'), 'dxa')

    # Hàng 0 Cột 0: đồng bộ 100% với ô Tên khu vực thực hiện của Bảng đo bụi (3 dòng in nghiêng 11pt)
    if len(trs) > 0:
        tcs0 = trs[0].findall(qn('w:tc'))
        if len(tcs0) >= 1 and len(ref_dust_tbl.rows) > 0:
            ref_tc0 = copy.deepcopy(ref_dust_tbl.rows[0].cells[0]._tc)
            ref_tcW = ref_tc0.get_or_add_tcPr().find(qn('w:tcW'))
            if ref_tcW is not None:
                ref_tcW.set(qn('w:w'), '3228')
                ref_tcW.set(qn('w:type'), 'dxa')
            trs[0].replace(tcs0[0], ref_tc0)
        # Hàng 0 Cột 1: xóa các đoạn rỗng thừa cuối ô HVAC (P5, P6)
        tcs0 = trs[0].findall(qn('w:tc'))
        if len(tcs0) >= 2:
            ps_hvac = tcs0[1].findall(qn('w:p'))
            while len(ps_hvac) > 5 and "".join(ps_hvac[-1].itertext()).strip() == "":
                tcs0[1].remove(ps_hvac[-1])
                ps_hvac = tcs0[1].findall(qn('w:p'))

    # Hàng 1 Cột 0: để trống (xóa 'CHỈ TIÊU THỰC HIỆN'); Hàng 1 Cột 1: rút gọn chấm lửng không rớt dòng
    if len(trs) > 1:
        tcs1 = trs[1].findall(qn('w:tc'))
        if len(tcs1) >= 1:
            for p_el in tcs1[0].findall(qn('w:p')):
                for r_el in list(p_el.findall(qn('w:r'))):
                    p_el.remove(r_el)
        if len(tcs1) >= 2:
            for p_el in tcs1[1].findall(qn('w:p')):
                p_txt = "".join(p_el.itertext())
                if "Lưu lượng lấy" in p_txt:
                    runs = p_el.findall(qn('w:r'))
                    if runs:
                        r0_pr = copy.deepcopy(runs[0].find(qn('w:rPr')))
                        for r_el in runs:
                            p_el.remove(r_el)
                        new_r = parse_xml(
                            f'<w:r {nsdecls("w")}><w:t xml:space="preserve">+ Lưu lượng lấy mẫu: {"\u2026" * 13}</w:t></w:r>'
                        )
                        if r0_pr is not None:
                            new_r.insert(0, r0_pr)
                        p_el.append(new_r)
                elif "Ký hiệu" in p_txt and "đĩa" in p_txt:
                    for t_el in p_el.iter(qn('w:t')):
                        if t_el.text:
                            t_el.text = t_el.text.replace("đĩa :", "đĩa:")

    # Co giãn số hàng dữ liệu (3 hàng header + num_rows hàng dữ liệu)
    target_total = 3 + max(3, int(num_rows))
    trs = tbl_el.findall(qn('w:tr'))
    while len(trs) > target_total:
        tbl_el.remove(trs[-1])
        trs = tbl_el.findall(qn('w:tr'))
    while len(trs) < target_total and len(trs) >= 4:
        tbl_el.append(copy.deepcopy(trs[-1]))
        trs = tbl_el.findall(qn('w:tr'))

    for tr in tbl_el.findall(qn('w:tr')):
        trPr = tr.get_or_add_trPr()
        if trPr.find(qn('w:cantSplit')) is None:
            trPr.append(OxmlElement('w:cantSplit'))

    return tbl_el


def build_bm01_phong_sach(customer_info, indicators=None, num_samples=8, num_vskk_samples=None, output_path=None):
    """
    Xây dựng BM3030.01 cho Phòng sạch:
    - Bảo toàn 100% format chuẩn của Bảng đo hạt bụi (Table 2):
      + Không có tiêu đề phía trên Table 2.
      + Hàng 0 Cột 1 giữ nguyên checkbox Wingdings & hiện trạng HVAC.
      + Hàng 1 Cột 1 giữ nguyên checkbox Particle (Lasair, AeroTrak, 1m, 1.2m, file zeroing, file đo...).
      + Co giãn số hàng STT theo số lượng điểm đo + 2-3 dòng dự phòng.
    - Bảo toàn 100% Bảng ghi chép Vi sinh gốc (Table 5 của mẫu gốc) khi có lấy mẫu vi sinh không khí.
    - Tự động sinh thêm các bảng đo thứ cấp nếu có chỉ tiêu tương ứng (Tốc độ gió, Cường độ sáng, Hướng dòng khí, Swab).
    """
    template_ps = os.path.join(BASE_DIR, "Phong sach, Vi sinh", "BM3030.01- Ghi chép - 116326.docx")
    doc = docx.Document(template_ps)

    # Cập nhật Header HL
    update_header_hl(doc, "HL:01/08/2025")

    # Lưu bản sao Bảng Vi sinh gốc (Table 5) trước khi dọn các bảng thừa sau Table 2
    orig_vskk_tbl_el = copy.deepcopy(doc.tables[5]._tbl) if len(doc.tables) >= 6 else None

    # Xóa các bảng thừa cũ sau Table 2 (Table 3, 4, 5...)
    while len(doc.tables) > 3:
        tbl_del = doc.tables[-1]
        prev = tbl_del._tbl.getprevious()
        if prev is not None and prev.tag.endswith('p'):
            prev.getparent().remove(prev)
        tbl_del._tbl.getparent().remove(tbl_del._tbl)

    # Điền Table 0 và Table 1
    fill_bm01_header_and_table1(doc, customer_info)

    # Chuẩn hóa Table 2 (Bảng đo bụi)
    tbl_dust = doc.tables[2]
    tblPr_dust = tbl_dust._tbl.tblPr
    if tblPr_dust.find(qn('w:tblLayout')) is None:
        tblPr_dust.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))

    # Row 0 Cell 0: 3 paragraphs chấm lửng font 11 italic chuẩn
    if len(tbl_dust.rows) > 0 and len(tbl_dust.rows[0].cells) > 0:
        c0 = tbl_dust.rows[0].cells[0]
        if len(c0.paragraphs) > 0:
            p0 = c0.paragraphs[0]
            p0.text = "   Tên khu vực thực hiện: ………"
            for r in p0.runs:
                r.italic = True
                r.font.name = 'Times New Roman'
                r.font.size = Pt(11)
        if len(c0.paragraphs) > 1:
            p1 = c0.paragraphs[1]
            p1.text = "   …………………………………"
            for r in p1.runs:
                r.italic = True
                r.font.name = 'Times New Roman'
                r.font.size = Pt(11)
        if len(c0.paragraphs) > 2:
            p2 = c0.paragraphs[2]
            p2.text = "   …………………………………"
            for r in p2.runs:
                r.italic = True
                r.font.name = 'Times New Roman'
                r.font.size = Pt(11)

    # Row 1 Cell 1: Rút ngắn chấm lửng File zeroing và File đo để không rớt dòng
    if len(tbl_dust.rows) > 1 and len(tbl_dust.rows[1].cells) > 1:
        c1_dust = tbl_dust.rows[1].cells[1]
        for p in c1_dust.paragraphs:
            if 'Thời gian lấy' in p.text:
                for r in p.runs:
                    if '…' in r.text or '...' in r.text:
                        r.text = ' ………………………………'
            if 'File zeroing' in p.text:
                runs = p.runs
                if len(runs) >= 3:
                    runs[2].text = ' File zeroing ………………………………'
                    for r in runs[3:]: r.text = ''
            if 'File đo' in p.text:
                runs = p.runs
                if len(runs) >= 3:
                    runs[2].text = ' File đo………………………………………'
                    for r in runs[3:]: r.text = ''

    # Co giãn số hàng STT: 3 hàng header + target_samples hàng
    target_samples = max(4, int(num_samples))
    target_total_rows = 3 + target_samples
    while len(tbl_dust.rows) > target_total_rows:
        tbl_dust._tbl.remove(tbl_dust.rows[-1]._tr)
    while len(tbl_dust.rows) < target_total_rows and len(tbl_dust.rows) >= 4:
        tbl_dust._tbl.append(copy.deepcopy(tbl_dust.rows[-1]._tr))

    for r in tbl_dust.rows:
        trPr = r._tr.get_or_add_trPr()
        if trPr.find(qn('w:cantSplit')) is None:
            trPr.append(OxmlElement('w:cantSplit'))

    # Kiểm tra các bảng đo bổ sung theo yêu cầu của hồ sơ
    combined_info = f"{customer_info.get('request_content', '')} {' '.join(str(i) for i in (indicators or []))}".lower()
    has_vskk = any(k in combined_info for k in ["vi sinh không khí", "vi sinh trong không khí", "vskk", "air micro"])
    # [14/11/2026] Đặt đĩa (lắng) không dùng bảng Vi sinh gốc (máy Buck/SKC) mà dùng bảng thời gian đặt đĩa
    has_settle = any(k in combined_info for k in ["đặt đĩa", "đĩa lắng", "settle plate"])
    has_contact = any(k in combined_info for k in ["đĩa tiếp xúc", "contact plate", "rodac"])
    if has_vskk and has_settle:
        orig_vskk_tbl_el = None

    # Khi Bảng đo bụi lớn (>= 8 dòng) hoặc hồ sơ có nhiều bảng (Bụi + Vi sinh),
    # lấp đầy nửa dưới Trang 1 bằng các dòng chấm lửng (147 dấu '.') và sang Trang 2 cho Bảng đo bụi
    # để bảng không bị cắt đôi qua 2 trang làm lặp lại 3 hàng header lớn.
    if target_samples >= 8 or (has_vskk and target_samples >= 6):
        dot_line_p1 = '.' * 147
        num_dot_lines_p1 = _calc_bm01_p1_dot_lines(customer_info)
        for _ in range(num_dot_lines_p1):
            p_dot_xml = (
                f'<w:p {nsdecls("w")}>'
                f'<w:pPr><w:spacing w:before="40" w:after="40" w:line="276" w:lineRule="auto"/>'
                f'<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:pPr>'
                f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>'
                f'<w:t>{dot_line_p1}</w:t></w:r></w:p>'
            )
            tbl_dust._tbl.addprevious(parse_xml(p_dot_xml))
        p_pb_dust = parse_xml(
            f'<w:p {nsdecls("w")}><w:pPr><w:pageBreakBefore/>'
            f'<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
            f'<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>'
        )
        tbl_dust._tbl.addprevious(p_pb_dust)

    add_6pt_spacer_after(tbl_dust._tbl, doc)

    # [14/11/2026] Số dòng từng bảng thứ cấp theo số vị trí thực tế, vd "air_velocity:3;filter_dp:3"
    row_overrides = {}
    for item in str(customer_info.get("bm01_rows") or "").replace(",", ";").split(";"):
        if ":" in item:
            k, v = item.split(":", 1)
            try:
                row_overrides[k.strip()] = max(1, int(v.strip()))
            except ValueError:
                pass

    secondary_needed = []
    # [14/11/2026] chỉ quét rò rỉ mới cần bảng HEPA leak; "chênh áp màng lọc HEPA" có bảng riêng
    if any(k in combined_info for k in ["rò rỉ", "leak", "xuyên thấu"]):
        secondary_needed.append('hepa_leak')
    if any(k in combined_info for k in ["tốc độ gió", "vận tốc gió", "trao đổi khí"]):
        secondary_needed.append('air_velocity')
    if "chênh áp" in combined_info and any(k in combined_info for k in ["màng lọc", "hepa"]):
        secondary_needed.append('filter_dp')
    if any(k in combined_info for k in ["cường độ sáng", "độ rọi", "độ sáng"]):
        secondary_needed.append('light')
    if any(k in combined_info for k in ["hướng dòng khí", "hình thái dòng khí", "khói"]):
        secondary_needed.append('airflow_pattern')
    if has_vskk and orig_vskk_tbl_el is None:
        secondary_needed.append('vskk_settle' if has_settle else 'vskk')
    if any(k in combined_info for k in ["vi sinh bề mặt", "swab", "vsbm"]):
        secondary_needed.append('contact_plate' if has_contact else 'swab')

    for stype in secondary_needed:
        sec_rows = target_samples
        if stype == 'filter_dp':
            sec_rows = 4  # [14/11/2026] thường chỉ 1-2 màng lọc
        elif stype in ('vskk_settle', 'contact_plate') and num_vskk_samples:
            sec_rows = int(num_vskk_samples)
        if stype in row_overrides:
            sec_rows = row_overrides[stype]
        if stype == 'hepa_leak' and target_samples >= 8:
            # [11/11/2026] Bảng đo bụi đã chiếm trọn Trang 2 -> đưa Bảng quét rò rỉ HEPA sang trọn Trang 3
            # (không bị cắt đôi qua 2 trang), chừa 15 dòng cho nhiều màng lọc HEPA.
            doc.element.body.append(parse_xml(
                f'<w:p {nsdecls("w")}><w:pPr><w:pageBreakBefore/>'
                f'<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
                f'<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>'
            ))
            sec_rows = 15
        sec_tbl = build_secondary_measurement_table(tbl_dust, stype, num_rows=sec_rows)
        doc.element.body.append(sec_tbl)
        add_6pt_spacer_after(sec_tbl, doc)

    # Chèn Bảng Vi sinh gốc của QUATEST 3 (nếu có yêu cầu lấy mẫu vi sinh trong không khí)
    if has_vskk and orig_vskk_tbl_el is not None:
        vskk_rows = int(num_vskk_samples) if num_vskk_samples else target_samples
        if target_samples >= 6:
            p_pb_vskk = parse_xml(
                f'<w:p {nsdecls("w")}><w:pPr><w:pageBreakBefore/>'
                f'<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
                f'<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>'
            )
            doc.element.body.append(p_pb_vskk)
        vskk_tbl_ready = prepare_original_vskk_table_element(orig_vskk_tbl_el, tbl_dust, num_rows=vskk_rows)
        doc.element.body.append(vskk_tbl_ready)
        add_6pt_spacer_after(vskk_tbl_ready, doc)

    finalize_document_formatting(doc)
    ensure_single_signature_at_end(doc)  # [FIX 28/10/2026] ký 1 lần ở cuối ghi chép

    if output_path:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        doc.save(output_path)

    return doc


# ==============================================================================
# 4b. MODULE BM01: LẤY MẪU VI SINH BỀ MẶT (SWAB / VSBM)
# ==============================================================================

def build_bm01_vi_sinh_be_mat(customer_info, production_lines=None, num_samples=21, output_path=None):
    """
    Xây dựng BM3030.01 cho hồ sơ Lấy mẫu vi sinh bề mặt (Swab / VSBM):
    - Chuẩn mực biểu mẫu QUATEST 3 theo mẫu 'i.   Lấy mẫu vi sinh/ vệ sinh công nghiệp' (Table 13 trong BM3030.01 master).
    - Trang 1: Header Table 0 + Table 1 Thông tin chung + các hàng chấm lửng lấp đầy Trang 1.
    - Phân tách theo từng dây chuyền sản xuất / vị trí đo (ví dụ: Dây chuyền sản xuất nắp DMCA01, Dây chuyền sản xuất phôi DMPA01).
    - Mỗi dây chuyền 1 bảng vi sinh chuẩn gồm khối HVAC, Vi sinh (APBuck/SKC) và Vi sinh bề mặt (+ Vị trí thực hiện, + Diện tích quét, + Dung dịch/thể tích, Kí hiệu mẫu).
    - Ký 1 lần duy nhất ở cuối tài liệu (sau bảng đo cuối cùng).
    """
    template_ps = os.path.join(BASE_DIR, "Phong sach, Vi sinh", "BM3030.01- Ghi chép - 116326.docx")
    doc = docx.Document(template_ps)
    update_header_hl(doc, "HL:01/08/2025")

    # Tải mẫu Table 13 (Lấy mẫu vi sinh/ vệ sinh công nghiệp) từ master SOP BM3030.01
    master_path = os.path.join(BASE_DIR, "QuyTrinh", "1. QUATEST3 3030.2025 - QUY TRÌNH MÔI TRƯỜNG + BIỂU MẪU", "BM3030.01-Ghi chep hien truong-02-08-final.docx")
    master_doc = docx.Document(master_path)
    t13_src = master_doc.tables[13]

    # Xóa các bảng thừa cũ sau Table 1
    while len(doc.tables) > 2:
        tbl_del = doc.tables[-1]
        prev = tbl_del._tbl.getprevious()
        if prev is not None and prev.tag.endswith('p'):
            prev.getparent().remove(prev)
        tbl_del._tbl.getparent().remove(tbl_del._tbl)

    # Điền Table 0 và Table 1
    fill_bm01_header_and_table1(doc, customer_info)

    # Lấp đầy Trang 1 bằng các hàng chấm lửng tràn lề layout ('.' * 147)
    dot_line_p1 = '.' * 147
    num_dot_lines_p1 = _calc_bm01_p1_dot_lines(customer_info)
    for _ in range(num_dot_lines_p1):
        p_dot_xml = (
            f'<w:p {nsdecls("w")}>'
            f'<w:pPr><w:spacing w:before="40" w:after="40" w:line="276" w:lineRule="auto"/>'
            f'<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>'
            f'<w:t>{dot_line_p1}</w:t></w:r></w:p>'
        )
        doc.element.body.append(parse_xml(p_dot_xml))

    # Xác định danh sách dây chuyền / bảng đo
    lines_cfg = []
    if production_lines and isinstance(production_lines, list) and len(production_lines) > 0:
        lines_cfg = production_lines
    else:
        lines_info = customer_info.get("lines") or customer_info.get("production_lines")
        if lines_info:
            lines_cfg = lines_info
        elif "dmca01" in str(customer_info).lower() or "dmpa01" in str(customer_info).lower():
            lines_cfg = [
                {"name": "Dây chuyền sản xuất nắp DMCA01", "samples": 9},
                {"name": "Dây chuyền sản xuất phôi DMPA01", "samples": 12}
            ]
        else:
            lines_cfg = [
                {"name": "Dây chuyền sản xuất", "samples": int(num_samples) if num_samples else 10}
            ]

    custom_area = customer_info.get("sampling_area") or customer_info.get("swab_area") or "10cm * 10cm"

    def _make_t13_instance(line_name, area=custom_area):
        t_elem = copy.deepcopy(t13_src._tbl)
        doc_dummy = docx.Document()
        doc_dummy.element.body.append(t_elem)
        tbl = doc_dummy.tables[0]
        
        # Row 0: Tên khu vực
        c0 = tbl.rows[0].cells[0]
        c0.paragraphs[0].text = f"   Tên khu vực thực hiện: {line_name}"
        if c0.paragraphs[0].runs:
            c0.paragraphs[0].runs[0].font.name = "Times New Roman"
            c0.paragraphs[0].runs[0].font.size = Pt(11)
            c0.paragraphs[0].runs[0].font.italic = True
        
        # Row 3: Vị trí thực hiện
        c3 = tbl.rows[3].cells[2]
        c3.text = "…………………………………………………………………………"
        if c3.paragraphs and c3.paragraphs[0].runs:
            c3.paragraphs[0].runs[0].font.name = "Times New Roman"
            c3.paragraphs[0].runs[0].font.size = Pt(11)
            
        # Row 4: Diện tích quét
        c4 = tbl.rows[4].cells[2]
        c4.text = area
        if c4.paragraphs and c4.paragraphs[0].runs:
            c4.paragraphs[0].runs[0].font.name = "Times New Roman"
            c4.paragraphs[0].runs[0].font.size = Pt(11)
            c4.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
        # Row 5: Dung dịch sử dụng
        c5 = tbl.rows[5].cells[2]
        c5.text = "…………………………………………………………………………"
        if c5.paragraphs and c5.paragraphs[0].runs:
            c5.paragraphs[0].runs[0].font.name = "Times New Roman"
            c5.paragraphs[0].runs[0].font.size = Pt(11)
            
        # Row 6: Kí hiệu mẫu
        c6 = tbl.rows[6].cells[2]
        c6.text = "…………………………………………………………………………"
        if c6.paragraphs and c6.paragraphs[0].runs:
            c6.paragraphs[0].runs[0].font.name = "Times New Roman"
            c6.paragraphs[0].runs[0].font.size = Pt(11)
            
        return t_elem

    # Render từng bảng theo form chuẩn trên từng trang độc lập
    for idx, lcfg in enumerate(lines_cfg, start=1):
        lname = lcfg.get("name", f"Khu vực lấy mẫu {idx}") if isinstance(lcfg, dict) else str(lcfg)

        # Ngắt trang trước mỗi bảng
        p_pb_xml = (
            f'<w:p {nsdecls("w")}><w:pPr><w:pageBreakBefore/>'
            f'<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
            f'<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>'
        )
        doc.element.body.append(parse_xml(p_pb_xml))

        # Tiêu đề bảng lề trái tuyệt đối, in đậm font 13 theo chuẩn biểu mẫu
        p_title_xml = (
            f'<w:p {nsdecls("w")}><w:pPr><w:jc w:val="left"/>'
            f'<w:spacing w:before="60" w:after="60"/><w:ind w:left="0"/></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>'
            f'<w:t>i.   Lấy mẫu vi sinh/ vệ sinh công nghiệp - {lname}:</w:t></w:r></w:p>'
        )
        doc.element.body.append(parse_xml(p_title_xml))

        # Bảng vi sinh chuẩn Table 13
        vsbm_tbl = _make_t13_instance(lname, area=custom_area)
        doc.element.body.append(vsbm_tbl)
        add_6pt_spacer_after(vsbm_tbl, doc)

        # Lấp đầy khoảng trống còn lại bằng các hàng chấm lửng
        num_dots = 7 if idx < len(lines_cfg) else 5
        for _ in range(num_dots):
            p_d_xml = (
                f'<w:p {nsdecls("w")}>'
                f'<w:pPr><w:spacing w:before="40" w:after="40" w:line="276" w:lineRule="auto"/>'
                f'<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:pPr>'
                f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>'
                f'<w:t>{dot_line_p1}</w:t></w:r></w:p>'
            )
            doc.element.body.append(parse_xml(p_d_xml))

    finalize_document_formatting(doc)
    ensure_single_signature_at_end(doc)

    if output_path:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        doc.save(output_path)

    return doc


# ==============================================================================
# 5. MODULE BM01: HYBRID (KẾT HỢP KHÍ NÉN + PHÒNG SẠCH)
# ==============================================================================

def build_bm01_hybrid_khi_nen_phong_sach(customer_info, num_kn_positions=2, num_ps_samples=8, indicators=None, output_path=None):
    """
    Xây dựng BM3030.01 cho đơn hàng kết hợp Khí nén + Phòng sạch (như Nextern Q326D133266):
    - Trang 1: Table 1 thông tin + 11 hàng chấm lửng lấp đầy trang 1.
    - Page break sang Trang 2:
      + Bảng Khí nén (Vị trí 01, 02...).
      + Bảng đo hạt bụi phòng sạch (Particle Table).
    """
    template_ps = os.path.join(BASE_DIR, "Phong sach, Vi sinh", "BM3030.01- Ghi chép - 116326.docx")
    template_kn = os.path.join(BASE_DIR, "Khi nen", "BM3030.01- Ghi chép - Q326D122266.docx")

    doc = docx.Document(template_ps)
    doc_kn = docx.Document(template_kn)

    # Cập nhật Header HL
    update_header_hl(doc, "HL:01/08/2025")

    # Xóa các bảng đo phía sau Table 2 (Table 3, 4, 5...)
    # [FIX 27/10/2026] Xóa luôn đoạn tiêu đề ngay trước mỗi bảng bị xóa
    # (bản cũ để sót tiêu đề 'Đo chênh áp phòng sạch' dù hồ sơ không yêu cầu chênh áp)
    while len(doc.tables) > 3:
        tbl_del = doc.tables[-1]
        prev = tbl_del._tbl.getprevious()
        if prev is not None and prev.tag.endswith('p'):
            prev.getparent().remove(prev)
        tbl_del._tbl.getparent().remove(tbl_del._tbl)

    # Điền Table 0 và Table 1
    fill_bm01_header_and_table1(doc, customer_info)

    # Chuẩn hóa Bảng đo hạt bụi phòng sạch (Table 2)
    tbl_dust = doc.tables[2]
    tblPr_dust = tbl_dust._tbl.tblPr
    if tblPr_dust.find(qn('w:tblLayout')) is None:
        tblPr_dust.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))

    # Format Row 0 Cell 0
    if len(tbl_dust.rows) > 0 and len(tbl_dust.rows[0].cells) > 0:
        c0 = tbl_dust.rows[0].cells[0]
        if len(c0.paragraphs) > 0: c0.paragraphs[0].text = "   Tên khu vực thực hiện: ………"
        if len(c0.paragraphs) > 1: c0.paragraphs[1].text = "   …………………………………"
        if len(c0.paragraphs) > 2: c0.paragraphs[2].text = "   …………………………………"
        for p in c0.paragraphs:
            for r in p.runs:
                r.italic = True
                r.font.name = "Times New Roman"
                r.font.size = Pt(11)

    # Rút ngắn chấm lửng Table 2
    if len(tbl_dust.rows) > 1 and len(tbl_dust.rows[1].cells) > 1:
        c1_dust = tbl_dust.rows[1].cells[1]
        for p in c1_dust.paragraphs:
            if 'Thời gian lấy' in p.text:
                for r in p.runs:
                    if '…' in r.text or '...' in r.text: r.text = ' ………………………………'
            if 'File zeroing' in p.text:
                runs = p.runs
                if len(runs) >= 3:
                    runs[2].text = ' File zeroing ………………………………'
                    for r in runs[3:]: r.text = ''
            if 'File đo' in p.text:
                runs = p.runs
                if len(runs) >= 3:
                    runs[2].text = ' File đo………………………………………'
                    for r in runs[3:]: r.text = ''

    # Co giãn số hàng STT bụi
    target_dust_rows = 3 + num_ps_samples
    while len(tbl_dust.rows) > target_dust_rows:
        tbl_dust._tbl.remove(tbl_dust.rows[-1]._tr)

    for r in tbl_dust.rows:
        trPr = r._tr.get_or_add_trPr()
        if trPr.find(qn('w:cantSplit')) is None:
            trPr.append(OxmlElement('w:cantSplit'))

    # Chuẩn bị điểm chèn trước Bảng bụi
    insert_point = tbl_dust._tbl

    # [FIX 27/10/2026] Bố cục trang của BM01 Hybrid:
    #   Trang 1: Table 0 + Table 1 + hàng chấm lửng lấp đầy phần còn lại của Trang 1
    #   Trang 2..N+1: mỗi vị trí Khí nén một trang (tiêu đề 'Khí nén (Vị trí 0x):' có pageBreakBefore)
    #   Trang cuối: Bảng đo bụi phòng sạch (đoạn ngắt trang mảnh ngay trước bảng)
    # (bản cũ chèn 11 hàng chấm SAU đoạn ngắt trang -> Trang 1 trống nửa dưới, hàng chấm rơi sang Trang 2,
    #  và 2 bảng khí nén chảy liền nhau làm 'Người thực hiện' bị rớt lẻ loi sang đầu trang sau)
    eq_count = len([l for l in format_equipment_multiline(customer_info.get("equipment", "")).split("\n") if l.strip()])
    num_dot_lines_p1 = max(6, 16 - max(0, eq_count - 4))

    # Thêm hàng chấm lửng lấp đầy khoảng trắng Trang 1 của BM01
    dot_line_p1 = '.' * 147
    for _ in range(num_dot_lines_p1):
        p_xml = (
            f'<w:p {nsdecls("w")}>'
            f'<w:pPr><w:spacing w:before="40" w:after="40" w:line="276" w:lineRule="auto"/>'
            f'<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>'
            f'<w:t>{dot_line_p1}</w:t></w:r></w:p>'
        )
        insert_point.addprevious(parse_xml(p_xml))

    # Chèn các bảng Khí nén cho từng vị trí
    src_kn_tbl = doc_kn.tables[1]
    for pos in range(1, num_kn_positions + 1):
        # Tiêu đề Khí nén (mỗi vị trí bắt đầu ở đầu trang mới)
        p_kn = doc.add_paragraph()
        p_kn.paragraph_format.page_break_before = True
        p_kn.text = f"Khí nén (Vị trí {pos:02d}):"
        p_kn.runs[0].font.name = "Times New Roman"
        p_kn.runs[0].font.size = Pt(13)
        p_kn.runs[0].font.bold = True
        p_kn.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_kn.paragraph_format.space_before = Pt(0)
        p_kn.paragraph_format.space_after = Pt(2)
        insert_point.addprevious(p_kn._p)

        # Bảng Khí nén
        kn_elem = prepare_single_khi_nen_table_element(src_kn_tbl, has_micro=False)
        insert_point.addprevious(kn_elem)

        # [FIX 28/10/2026] KHÔNG chèn 'Người thực hiện' dưới từng bảng khí nén nữa —
        # người thực hiện chỉ ký 1 lần ở cuối ghi chép (xem ensure_single_signature_at_end).

    # Ngắt trang mảnh trước Bảng bụi phòng sạch (Bảng bụi bắt đầu trọn vẹn ở trang mới)
    p_pb_dust = parse_xml(
        f'<w:p {nsdecls("w")}><w:pPr><w:pageBreakBefore/>'
        f'<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
        f'<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>'
    )
    insert_point.addprevious(p_pb_dust)

    finalize_document_formatting(doc)
    ensure_single_signature_at_end(doc)  # [FIX 28/10/2026] ký 1 lần ở cuối ghi chép

    if output_path:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        doc.save(output_path)

    return doc


# ==============================================================================
# 6. UNIFIED DISPATCHER CHO BM3030.01
# ==============================================================================

def build_bm3030_01(
    customer_info,
    category_key="phong_sach",
    indicators=None,
    num_samples=None,
    num_positions=1,
    output_path=None
):
    """
    Dispatcher tự động lựa chọn module BM3030.01 chuẩn xác theo từng loại hồ sơ:
    - BSC -> build_bm01_bsc
    - Khí nén -> build_bm01_khi_nen
    - Khí nén + Phòng sạch -> build_bm01_hybrid_khi_nen_phong_sach
    - Vi sinh bề mặt thuần túy -> build_bm01_vi_sinh_be_mat
    - Phòng sạch / Vi sinh -> build_bm01_phong_sach
    - MTLĐ / Kho lạnh -> build tương ứng
    """
    combined_text = f"{category_key} {customer_info.get('request_content', '')} {customer_info.get('short_content', '')} {' '.join(str(i) for i in (indicators or []))}".lower()

    # 1. Tủ an toàn sinh học (BSC)
    if category_key == "bsc" or "bsc" in combined_text or "an toàn sinh học" in combined_text:
        return build_bm01_bsc(customer_info, num_positions=num_positions, output_path=output_path)

    # 2. Khí nén + Phòng sạch kết hợp
    has_kn = any(k in combined_text for k in ["khí nén", "khi nen", "tổng dầu", "điểm sương"])
    has_ps = any(k in combined_text for k in ["phòng sạch", "phong sach", "nồng độ bụi phòng sạch"])
    if (category_key == "khi_nen_phong_sach") or (has_kn and has_ps):
        return build_bm01_hybrid_khi_nen_phong_sach(
            customer_info,
            num_kn_positions=num_positions if num_positions > 1 else 2,
            num_ps_samples=num_samples if num_samples else 8,
            indicators=indicators,
            output_path=output_path
        )

    # 3. Khí nén đơn thuần
    if category_key == "khi_nen" or (has_kn and not has_ps):
        return build_bm01_khi_nen(
            customer_info,
            num_positions=num_positions,
            indicators=indicators,
            output_path=output_path
        )

    # 4. Vi sinh bề mặt thuần túy (không có đo bụi phòng sạch)
    has_swab = any(k in combined_text for k in ["vi sinh bề mặt", "swab", "vsbm"])
    has_ps_dust = any(k in combined_text for k in ["nồng độ bụi phòng sạch", "hạt bụi", "bụi phòng sạch", "aerotrak", "lasair"])
    if category_key == "vi_sinh_be_mat" or (has_swab and not has_ps_dust and not has_kn):
        return build_bm01_vi_sinh_be_mat(
            customer_info,
            production_lines=customer_info.get("production_lines"),
            num_samples=num_samples,
            output_path=output_path
        )

    # 5. Phòng sạch & Vi sinh (Mặc định)
    return build_bm01_phong_sach(
        customer_info,
        indicators=indicators,
        num_samples=num_samples if num_samples else 8,
        num_vskk_samples=customer_info.get("num_vskk_bm01"),
        output_path=output_path
    )
