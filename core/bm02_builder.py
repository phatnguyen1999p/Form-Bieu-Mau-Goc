"""
Module chuyên trách xây dựng và chuẩn hóa biểu mẫu BM3030.02
(Biên bản kiểm tra / giám định tại hiện trường / Biên bản lấy mẫu)
Tuân thủ 100% 11 Quy tắc soạn thảo chuẩn mực (GEMINI.md).
Đảm bảo tài liệu luôn đạt đúng 2 trang hoàn hảo, không bị nhảy sang trang 3.
"""
import os
import re
import copy
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
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
    update_header_hl,
    update_customer_info,
    adjust_sample_table,
    finalize_document_formatting,
    set_cell_text_preserve_formatting
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


def build_bm3030_02(
    template_path,
    customer_info,
    indicators=None,
    num_samples=None,
    category_key="phong_sach",
    output_path=None
):
    """
    Xây dựng tài liệu BM3030.02 hoàn chỉnh và tối ưu hóa layout 2 trang chuẩn mực.
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Không tìm thấy mẫu BM3030.02: {template_path}")

    doc = docx.Document(template_path)

    # 1. Cập nhật Header HL:01/08/2025 theo Quy tắc 3
    update_header_hl(doc, "HL:01/08/2025")

    is_bsc = (category_key == "bsc" or "bsc" in str(customer_info.get("short_content", "")).lower())

    # 2. Cập nhật thông tin khách hàng và thông tin chung vào Table 0 & Table 1
    norm_indicators = indicators or customer_info.get("indicators", [])
    update_customer_info(doc, customer_info, indicators=norm_indicators)

    # 3. Format 'Thời gian thực hiện' để vừa vặn 1 dòng duy nhất (Quy tắc 2)
    t1 = doc.tables[1]
    exec_date = customer_info.get("execution_date") or get_default_execution_date_dotted()
    time_text = f"Bắt đầu từ........... kết thúc lúc........... ngày {exec_date}"
    for r in t1.rows:
        if any("thời gian thực hiện" in c.text.lower() for c in r.cells):
            c_target = r.cells[2]
            c_target.text = time_text
            if c_target.paragraphs:
                p0 = c_target.paragraphs[0]
                if p0.runs:
                    p0.runs[0].font.name = "Times New Roman"
                    p0.runs[0].font.size = Pt(13)
                p0.paragraph_format.left_indent = Inches(0)
                p0.paragraph_format.right_indent = Inches(0)
            break

    # 4. Điều chỉnh Table 2
    t2 = doc.tables[2]
    num_data_items = len(norm_indicators) if norm_indicators else (num_samples if num_samples else 1)
    cabinet_list = []
    if is_bsc:
        # Bảng BSC: hỗ trợ 1 tủ hoặc nhiều tủ (ví dụ Tủ cấp 1 + Tủ cấp 2)
        cabinets = customer_info.get("cabinets")
        num_pos = int(customer_info.get("num_positions") or 1)
        if customer_info.get("blank_cabinets") or customer_info.get("blank_cabinet_name"):
            cabinet_list = ["." * 31 for _ in range(num_pos)]
        elif cabinets and isinstance(cabinets, (list, tuple)) and len(cabinets) > 0:
            cabinet_list = [str(c).strip() for c in cabinets if str(c).strip()]
        elif isinstance(cabinets, str) and cabinets.strip():
            cabinet_list = [c.strip() for c in cabinets.replace(";", "\n").split("\n") if c.strip()]
        elif num_pos > 1:
            cabinet_list = [f"Tủ an toàn sinh học số {i}" for i in range(1, num_pos + 1)]
        else:
            cabinet_list = ["Tủ an toàn sinh học cấp 2"]

        proto_tr = copy.deepcopy(t2.rows[1]._tr)
        for r in list(t2.rows)[1:]:
            t2._tbl.remove(r._tr)

        t2_header_tr = t2.rows[0]._tr.get_or_add_trPr()
        if t2_header_tr.find(qn('w:tblHeader')) is None:
            t2_header_tr.append(OxmlElement('w:tblHeader'))

        # [NEW 09/11/2026] Giữ hàng tiêu đề đi liền với hàng tủ đầu tiên: nếu hàng tủ không vừa cuối Trang 1
        # thì cả bảng sang Trang 2, không để hàng tiêu đề trơ trọi ở cuối Trang 1.
        for c_h in t2.rows[0].cells:
            for p_h in c_h.paragraphs:
                pPr_h = p_h._p.get_or_add_pPr()
                if pPr_h.find(qn('w:keepNext')) is None:
                    pPr_h.insert(0, OxmlElement('w:keepNext'))

        # [NEW 09/11/2026] Chừa chấm rộng rãi + giãn dòng 1.5 để KTV viết tay (mỗi dòng chấm vừa 1 dòng của cột):
        #   Tên tủ: tên tủ + 3 dòng chấm (để trắng tên -> 4 dòng chấm)
        #   Hãng sản xuất / Serial: mỗi nhãn + 2 dòng chấm; Mã quản lý, Ký hiệu ghi chép: 2 dòng chấm.
        ELL = '…'
        bsc_line_spacing = 1.5
        # [NEW 09/11/2026] Thêm mục Model: ô gồm 3 mục Hãng sản xuất, Serial, Model (cả tiêu đề cột và hàng tủ)
        BSC_HSX_DOT_LINES = 1
        hsx_serial = "\n".join(sum(([lbl] + [ELL * 12] * BSC_HSX_DOT_LINES
                                    for lbl in ("Hãng sản xuất:", "Serial:", "Model:")), []))
        # Tiêu đề cột giữ 2 dòng ("Hãng sản xuất" / "Serial / Model") để hàng tiêu đề không cao thêm,
        # nhờ vậy Tủ 1 vẫn vừa cuối Trang 1.
        hdr_cell = t2.rows[0].cells[2]
        if "Model" not in hdr_cell.text:
            last_p = hdr_cell.paragraphs[-1]._p
            t_nodes = last_p.findall('.//' + qn('w:t'))
            if t_nodes:
                t_nodes[0].text = "Serial / Model"
                for t_n in t_nodes[1:]:
                    t_n.text = ""
        for idx, cab_name in enumerate(cabinet_list, start=1):
            t2._tbl.append(copy.deepcopy(proto_tr))
            r_cur = t2.rows[-1]
            trPr_cur = r_cur._tr.get_or_add_trPr()
            if trPr_cur.find(qn('w:cantSplit')) is None:
                trPr_cur.append(OxmlElement('w:cantSplit'))
            trH = trPr_cur.find(qn('w:trHeight'))
            if trH is None:
                trPr_cur.append(parse_xml(f'<w:trHeight {nsdecls("w")} w:val="794"/>'))
            else:
                trH.set(qn('w:val'), '794')

            if not cab_name or re.fullmatch(r'[\s.…]+', cab_name):
                cab_text = "\n".join([ELL * 9] * 4)
            else:
                cab_text = "\n".join([cab_name] + [ELL * 9] * 3)

            set_cell_text_preserve_formatting(r_cur.cells[0], str(idx), bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_text_preserve_formatting(r_cur.cells[1], cab_text, bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.LEFT)
            set_cell_text_preserve_formatting(r_cur.cells[2], hsx_serial, bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.LEFT)
            set_cell_text_preserve_formatting(r_cur.cells[3], "\n".join([ELL * 5] * 2), bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_text_preserve_formatting(r_cur.cells[4], "\n".join([ELL * 4] * 2), bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.CENTER)
            for c in r_cur.cells:
                for p in c.paragraphs:
                    p.paragraph_format.line_spacing = bsc_line_spacing
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
    else:
        num_target_rows = len(norm_indicators) if norm_indicators and len(norm_indicators) > 0 else (num_samples if num_samples else 1)
        adjust_sample_table(
            doc,
            num_samples=num_target_rows,
            indicators=norm_indicators,
            default_location=customer_info.get("location"),
            category_key=category_key
        )

        # Điều chỉnh khoảng cách dòng trong Table 2 để đảm bảo vừa vặn tuyệt đối 2 trang
        # Nếu có từ 4 chỉ tiêu trở lên, dùng line_spacing = 1.15 để tránh nhảy trang 3
        num_t2_data_rows = len(t2.rows) - 2
        spacing_val = 1.15 if num_t2_data_rows >= 4 else 1.5
        sb_val = Pt(1) if num_t2_data_rows >= 4 else Pt(2)
        sa_val = Pt(1) if num_t2_data_rows >= 4 else Pt(2)
        for r in t2.rows[2:]:
            for c in r.cells:
                for p in c.paragraphs:
                    p.paragraph_format.line_spacing = spacing_val
                    p.paragraph_format.space_before = sb_val
                    p.paragraph_format.space_after = sa_val

    for r in t2.rows:
        trPr = r._tr.get_or_add_trPr()
        if trPr.find(qn('w:cantSplit')) is None:
            trPr.append(OxmlElement('w:cantSplit'))

    # 5. TỐI ƯU HÓA LAYOUT 2 TRANG HOÀN HẢO THEO QUY TẮC 7
    # 5.1. Dọn dẹp khoảng trống thừa trước Table 1
    p_between = t1._tbl.getprevious()
    while p_between is not None and p_between.tag.endswith('p') and p_between != doc.tables[0]._tbl:
        prev_node = p_between.getprevious()
        p_between.getparent().remove(p_between)
        p_between = prev_node
    add_6pt_spacer_after(doc.tables[0]._tbl, doc)

    # 5.2. Dọn dẹp khoảng trống thừa giữa Table 1 và Table 2
    p_between2 = t2._tbl.getprevious()
    while p_between2 is not None and p_between2.tag.endswith('p') and p_between2 != t1._tbl:
        prev_node = p_between2.getprevious()
        p_between2.getparent().remove(p_between2)
        p_between2 = prev_node
    add_6pt_spacer_after(t1._tbl, doc)

    # 5.3 & 5.4. Thêm nội dung Mục 9 và xử lý ngắt trang
    quote_num = str(customer_info.get("quotation_number") or customer_info.get("quote_number") or "QONV2……………").strip()
    cust_name = normalize_customer_name(clean_text_no_status_or_quantitative(customer_info.get("customer_name", ""))).strip().rstrip(".,;")
    if not cust_name:
        cust_name = "khách hàng"

    # [NEW 09/11/2026] Không có số báo giá (QONV) nhưng có số PO -> dùng số PO.
    #   Có số báo giá thật       : 'theo báo giá số QONV26015439.'
    #   Không có QONV, có số PO   : 'theo đơn đặt hàng (PO) số E900192188.'
    #   Không có cả hai           : 'theo báo giá số QONV2……………' (chừa chấm điền tay)
    po_num = str(customer_info.get("po_number") or "").strip()
    quote_is_placeholder = (not quote_num) or ('…' in quote_num) or ('...' in quote_num)
    if quote_is_placeholder and po_num:
        quote_prefix = f"Nội dung thực hiện theo đơn đặt hàng (PO) số {po_num}"
        if not po_num.endswith('.'):
            quote_prefix += "."
    else:
        quote_prefix = f"Nội dung thực hiện theo báo giá số {quote_num}"
        if not quote_num.endswith(('.', '…')):
            quote_prefix += "."
    sec9_text = f"{quote_prefix} Vị trí thực hiện được chỉ định bởi đại diện {cust_name}, số lượng vị trí cụ thể như sau:"

    p_sec9_xml = (
        f'<w:p {nsdecls("w")}><w:pPr><w:spacing w:before="40" w:after="40" w:line="276" w:lineRule="auto"/>'
        f'<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>'
        f'<w:t>{sec9_text}</w:t></w:r></w:p>'
    )

    # BẮT BUỘC CHO TẤT CẢ HỒ SƠ: Table 2 bắt đầu ngay dưới Mục 9 ở Trang 1
    # TUYỆT ĐỐI KHÔNG ngắt trang trước Table 2 và KHÔNG chèn hàng chấm lửng giả ở Trang 1
    t2._tbl.addprevious(parse_xml(p_sec9_xml))

    # 5.5. Sau Table 2: Chuẩn hóa và co giãn số dòng chấm lửng (147 dấu '.') trên Trang 2 trước khối chữ ký
    # Chỉ khớp các paragraph thuần dấu chấm (re.fullmatch(r'[\s.…]+', txt)) hoặc paragraph rỗng,
    # TUYỆT ĐỐI không ghi đè hoặc xóa dòng 'Chỉ tiêu thực hiện lấy mẫu vi sinh...'
    has_micro_line = False
    pure_dot_els_p2 = []
    empty_els_p2 = []
    anchor_el_p2 = None
    curr = t2._tbl.getnext()
    while curr is not None:
        if curr.tag == qn('w:p'):
            txt_p = "".join(curr.itertext()).strip()
            if txt_p.startswith("Chỉ tiêu thực hiện lấy mẫu vi sinh"):
                has_micro_line = True
            elif not txt_p:
                if anchor_el_p2 is None:
                    empty_els_p2.append(curr)
            elif re.fullmatch(r'[\s.\u2026]+', txt_p) is not None:
                pure_dot_els_p2.append(curr)
            elif anchor_el_p2 is None and (
                txt_p.startswith("Ký hiệu phiếu ghi chép")
                or txt_p.startswith("Biên bản này được lập")
            ):
                anchor_el_p2 = curr
        curr = curr.getnext()

    for el_del in pure_dot_els_p2 + empty_els_p2:
        el_del.getparent().remove(el_del)

    # Co giãn số dòng chấm lửng trên Trang 2:
    # Nếu chỉ có 1 mẫu/1 tủ, Table 2 nằm gọn trên Trang 1 -> cần pageBreakBefore ở phần tử đầu Trang 2
    # và cần nhiều dòng chấm lửng (khoảng 8-14 dòng) để lấp đầy Trang 2.
    # Nếu có 2 mẫu trở lên, Table 2 tự tách sang Trang 2 -> chỉ cần 1-2 dòng chấm lửng (hoặc 0 nếu >=4 mẫu).
    num_t2_rows = len(t2.rows)
    if is_bsc:
        num_cabs = max(1, len(cabinet_list))
        # [09/11/2026] Hàng tủ cao hơn (giãn dòng 1.5) + tiêu đề keepNext: tùy độ dài Trang 1, Tủ 1 có thể nằm
        # cuối Trang 1 hoặc cả bảng sang Trang 2. KHÔNG ép ngắt trang sau bảng nữa (nếu bảng đã sang Trang 2,
        # ngắt trang sẽ đẩy chữ ký sang Trang 3). Số dòng chấm chọn để vừa 2 trang trong cả hai trường hợp.
        need_pb_after_t2 = False
        if num_cabs == 1:
            target_p2_dots = 18
        elif num_cabs == 2:
            target_p2_dots = 10
        else:
            target_p2_dots = max(0, 14 - 4 * (num_cabs - 1))
    else:
        num_data_rows = max(1, num_t2_rows - 2)
        need_pb_after_t2 = (num_data_rows == 1)
        if num_data_rows == 1:
            target_p2_dots = 8
        elif num_data_rows == 2:
            target_p2_dots = 2
        elif num_data_rows == 3:
            target_p2_dots = 1
        else:
            target_p2_dots = 0

        if has_micro_line and target_p2_dots > 0:
            target_p2_dots = max(0, target_p2_dots - 1)

    dot_line_p2 = '.' * 147
    if anchor_el_p2 is not None and target_p2_dots > 0:
        for idx_dot in range(target_p2_dots):
            # Nếu Table 2 nằm hoàn toàn ở Trang 1 (1 mẫu/1 tủ), đặt ngắt trang tại dòng chấm đầu tiên của Trang 2
            pb_tag = '<w:pageBreakBefore/>' if (need_pb_after_t2 and idx_dot == 0 and not has_micro_line) else ''
            p_p2_xml = (
                f'<w:p {nsdecls("w")}><w:pPr>{pb_tag}<w:spacing w:before="40" w:after="40" w:line="276" w:lineRule="auto"/>'
                f'<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:pPr>'
                f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>'
                f'<w:t>{dot_line_p2}</w:t></w:r></w:p>'
            )
            anchor_el_p2.addprevious(parse_xml(p_p2_xml))

    # Nếu có has_micro_line và need_pb_after_t2, gán pageBreakBefore cho dòng micro_line
    if has_micro_line and need_pb_after_t2:
        curr = t2._tbl.getnext()
        while curr is not None:
            if curr.tag == qn('w:p') and "Chỉ tiêu thực hiện lấy mẫu vi sinh" in "".join(curr.itertext()):
                pPr = curr.get_or_add_pPr()
                if pPr.find(qn('w:pageBreakBefore')) is None:
                    pPr.append(OxmlElement('w:pageBreakBefore'))
                break
            curr = curr.getnext()
    elif need_pb_after_t2 and not has_micro_line and target_p2_dots == 0 and anchor_el_p2 is not None:
        pPr = anchor_el_p2.get_or_add_pPr()
        if pPr.find(qn('w:pageBreakBefore')) is None:
            pPr.append(OxmlElement('w:pageBreakBefore'))

    # Xóa các paragraph rỗng thừa phía sau Table 2 (nhưng không xóa page break paragraph)
    for p in list(doc.paragraphs[4:]):
        if not p.text.strip() and not p.paragraph_format.page_break_before:
            p._p.getparent().remove(p._p)

    # 6. Chuẩn hóa format tổng thể
    finalize_document_formatting(doc)

    # [OPT 05/11/2026] Đối với BSC 1 tủ có nhiều thiết bị (>= 6 dòng):
    # Tinh chỉnh nhẹ spacing của khối thiết bị và Table 2 để bảo đảm Table 2 nằm trọn ở cuối Trang 1,
    # Trang 2 chứa 14 dòng chấm lửng + chữ ký, đạt đúng chuẩn 2 trang hoàn hảo theo Quy tắc 7.
    if is_bsc:
        eq_formatted = format_equipment_multiline(customer_info.get("equipment", ""))
        eq_count = len([l for l in eq_formatted.split("\n") if l.strip()])
        if eq_count >= 6:
            for r in t1.rows:
                for c in r.cells:
                    if any(kw in c.text for kw in ["TSI 8375", "Kanomax", "ATI aerosol", "Testo 540", "Testo 440", "Extech"]):
                        for p in c.paragraphs:
                            p.paragraph_format.space_before = Pt(1)
                            p.paragraph_format.space_after = Pt(1)
                            p.paragraph_format.line_spacing = 1.05
            # [09/11/2026] Không còn ép giãn dòng 1.25 cho hàng tủ: hàng tủ giữ giãn dòng 1.5 để viết tay.

        # [NEW 09/11/2026] Theo chỉnh tay của người dùng (hồ sơ Sanofi Q326D139329): bỏ khoảng cách 2pt trên/dưới
        # của hàng "Quy trình thực hiện", hàng "Nội dung/ Quá trình thực hiện" (Table 1) và câu dẫn Mục 9
        # -> tiết kiệm ~12pt để hàng Tủ 1 (giãn dòng 1.5) vừa cuối Trang 1. Làm SAU finalize_document_formatting
        # vì bước đó đặt lại spacing 2pt cho mọi ô bảng.
        for r in t1.rows:
            label = r.cells[0].text.strip() if r.cells else ""
            if label.startswith("Quy trình thực hiện") or label.startswith("Nội dung/"):
                seen = set()
                for c in r.cells:
                    if id(c._tc) in seen:
                        continue
                    seen.add(id(c._tc))
                    for p in c.paragraphs:
                        p.paragraph_format.space_before = Pt(0)
                        p.paragraph_format.space_after = Pt(0)
        p_prev = t2._tbl.getprevious()
        if p_prev is not None and p_prev.tag == qn('w:p') and "".join(p_prev.itertext()).startswith("Nội dung thực hiện theo"):
            sp = p_prev.find(qn('w:pPr') + '/' + qn('w:spacing'))
            if sp is not None:
                sp.set(qn('w:before'), '0')
                sp.set(qn('w:after'), '0')

    if output_path:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        doc.save(output_path)

    return doc
