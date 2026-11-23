"""
Module điều phối việc nhân bản hồ sơ biểu mẫu QUATEST 3 (Master Engine)
Tự động sinh trọn bộ:
  - BM3030.02: Biên bản giám định hiện trường / Biên bản lấy mẫu
  - BM3030.01: Phiếu ghi chép hiện trường chuyên biệt theo lĩnh vực
  - Xuất PDF tự động và ghép vector nét cao
  - Tự động cập nhật tiêu đề cuộc trò chuyện vào database (Quy tắc 10)
Tuân thủ 100% 11 Quy tắc soạn thảo chuẩn mực (GEMINI.md).
"""
import os
import re
import datetime
import random
import shutil
import docx

from .config import (
    CATEGORIES,
    OUTPUT_DIR,
    BASE_DIR,
    Z_DRIVE_DIR,
    resolve_equipment_and_procedure,
    infer_method_type
)
from .normalizer import (
    normalize_customer_name,
    normalize_date_str,
    normalize_address,
    normalize_location,
    extract_short_customer_name,
    extract_short_content,
    build_output_folder_name,
    auto_generate_doc_number,
    auto_generate_date,
    clean_indicator_name,
    format_equipment_multiline,
    clean_text_no_status_or_quantitative,
    get_default_execution_date_dotted,
    update_conversation_db_title,
    is_blank_doc_number,
    doc_number_for_filename,
    sanitize_for_filename,
    DOC_NUMBER_BLANK_PLACEHOLDER
)
from .bm02_builder import build_bm3030_02
from .bm01_builder import build_bm3030_01
from .pdf_exporter import export_order_pdfs
from .order_profiles import ORDER_PROFILES, get_profile, detect_profile


def sanitize_filename(name):
    """Làm sạch tên file/thư mục, loại bỏ ký tự không hợp lệ trên Windows."""
    clean = re.sub(r'[\\/*?:"<>|]', '_', str(name))
    clean = re.sub(r'\s+', ' ', clean).strip()
    return clean[:80]


def widen_blank_doc_number_cell(docx_path, extra_dxa=700):
    """
    [NEW 27/10/2026] Khi để trắng mã hồ sơ: nới rộng ô mã số ở header Table 0 ('Số: | Q326D…………… | | Ngày: | ...')
    bằng cách lấy bớt bề rộng của cột trống kế bên, để placeholder 'Q326D……………' nằm gọn 1 dòng
    và đủ chỗ cho KTV điền tay 6 chữ số. Tổng bề rộng bảng giữ nguyên.
    """
    from docx.oxml.ns import qn
    d = docx.Document(docx_path)
    if not d.tables:
        return False
    t0 = d.tables[0]
    if not t0.rows or not t0.rows[0].cells or not t0.rows[0].cells[0].text.strip().startswith("Số"):
        return False
    tbl = t0._tbl
    grid = tbl.find(qn('w:tblGrid'))
    cols = grid.findall(qn('w:gridCol')) if grid is not None else []
    if len(cols) < 3:
        return False
    w1, w2 = int(cols[1].get(qn('w:w'))), int(cols[2].get(qn('w:w')))
    if w2 - extra_dxa < 1000:
        return False
    total = sum(int(c.get(qn('w:w'))) for c in cols)
    cols[1].set(qn('w:w'), str(w1 + extra_dxa))
    cols[2].set(qn('w:w'), str(w2 - extra_dxa))
    tcs = t0.rows[0]._tr.findall(qn('w:tc'))
    if len(tcs) >= 3:
        for tc, delta in ((tcs[1], extra_dxa), (tcs[2], -extra_dxa)):
            tcW = tc.find(qn('w:tcPr') + '/' + qn('w:tcW'))
            if tcW is None:
                continue
            w_type = tcW.get(qn('w:type'))
            w_val = int(tcW.get(qn('w:w')))
            if w_type == 'pct':
                tcW.set(qn('w:w'), str(int(round(w_val + delta * 5000 / total))))
            elif w_type == 'dxa':
                tcW.set(qn('w:w'), str(w_val + delta))
    d.save(docx_path)
    return True


Z_SUBDIR_THUC_HIEN = "Thuc hien"   # [NEW 09/11/2026] file biên bản / ghi chép do hệ thống sinh ra
Z_SUBDIR_SCAN = "Scan"             # [NEW 09/11/2026] file gốc ban đầu (phiếu yêu cầu, PO... bản scan)


def copy_order_to_z_drive(target_dir, subfolder_name=None, z_base_dir=None, scan_files=None):
    # [NEW 09/11/2026] Cấu trúc trên ổ Z:
    #   Z:\...\2026\[Thư mục hồ sơ]\Thuc hien\  <- toàn bộ file trong output/[Thư mục hồ sơ]
    #   Z:\...\2026\[Thư mục hồ sơ]\Scan\       <- file gốc ban đầu (scan_files, ví dụ Input/xxxx.pdf)
    # Trả về đường dẫn thư mục 'Thuc hien' (pdf_exporter đồng bộ PDF vào đây).
    """
    [NEW 04/11/2026] Sao chép một bản trọn vẹn của thư mục hồ sơ (gồm docx, pdf...) sang ổ đĩa Z:
    Đường dẫn chuẩn: Z:\\9.NV2 - MOI TRUONG\\1.Ho So Moi Truong\\2026\\[subfolder_name] (Quy tắc 13)
    Nếu ổ Z không truy cập được hoặc chưa kết nối, ghi log cảnh báo an toàn.
    """
    base_dest = z_base_dir or Z_DRIVE_DIR
    if not os.path.exists(base_dest):
        print(f"[WARN] O dia Z hoac duong dan dich khong ton tai / chua ket noi: {base_dest}")
        return None

    if not os.path.exists(target_dir):
        print(f"[WARN] Thu muc nguon khong ton tai: {target_dir}")
        return None

    folder_name = subfolder_name or os.path.basename(target_dir)
    order_dir = os.path.join(base_dest, folder_name)
    dest_dir = os.path.join(order_dir, Z_SUBDIR_THUC_HIEN)

    if scan_files:
        if isinstance(scan_files, str):
            scan_files = [s.strip() for s in scan_files.split(";") if s.strip()]
        scan_dir = os.path.join(order_dir, Z_SUBDIR_SCAN)
        try:
            os.makedirs(scan_dir, exist_ok=True)
            for sf in scan_files:
                if os.path.isfile(sf):
                    shutil.copy2(sf, os.path.join(scan_dir, os.path.basename(sf)))
                    print(f"[INFO] Da sao chep file goc sang Scan: {os.path.basename(sf)}")
                else:
                    print(f"[WARN] Khong tim thay file goc: {sf}")
        except Exception as e:
            print(f"[WARN] Loi khi sao chep file goc sang Scan ({scan_dir}): {e}")

    try:
        os.makedirs(dest_dir, exist_ok=True)
        copied_files = []
        for item in os.listdir(target_dir):
            if item.startswith("~$"):
                continue  # Bỏ qua tệp khóa tạm của Word
            s_path = os.path.join(target_dir, item)
            d_path = os.path.join(dest_dir, item)
            if os.path.isfile(s_path):
                shutil.copy2(s_path, d_path)
                copied_files.append(d_path)
            elif os.path.isdir(s_path):
                shutil.copytree(s_path, d_path, dirs_exist_ok=True)
                copied_files.append(d_path)

        print(f"[INFO] Da sao chep thanh cong 1 ban sang o Z ({len(copied_files)} tep): {dest_dir}")
        return dest_dir
    except Exception as e:
        print(f"[WARN] Loi khi sao chep sang o Z ({dest_dir}): {e}")
        return None


def clone_form_set(
    category_key=None,
    customer_name="Công ty Cổ phần ABC",
    address="123 Đường XYZ, Phường 1, TP. HCM",
    location=None,
    num_samples=4,
    num_samples_bm01=None,
    num_positions=1,
    doc_number=None,
    date=None,
    registration_date=None,
    execution_date=None,
    request_content=None,
    equipment=None,
    procedure=None,
    indicators=None,
    custom_template_bm02=None,
    custom_template_bm01=None,
    output_base_dir=None,
    generate_bm01=True,
    generate_bm02=True,
    short_customer_name=None,
    short_content=None,
    profile=None,
    export_pdf=False,
    conversation_id=None,
    blank_doc_number=False,
    blank_cabinets=False,
    cabinets=None,
    num_vskk_bm01=None,
    bm01_rows=None,
    micro_targets=None,
    quotation_number=None,
    quote_number=None,
    po_number=None,
    scan_files=None,
    production_lines=None,
    copy_to_z=True,
    z_drive_dir=None
):
    """
    Nhân bản trọn bộ hồ sơ (BM3030.01 và BM3030.02) cho khách hàng mới.
    Tự động nhận diện profile mẫu, tối ưu bố cục 2 trang và xuất PDF vector.
    """
    # 1. Nhận diện hoặc áp dụng Profile mẫu nếu có
    order_data_for_detection = {
        "request_content": request_content,
        "indicators": indicators,
        "category": category_key,
        "customer_name": customer_name
    }
    selected_profile_id = profile if profile and profile in ORDER_PROFILES else (
        category_key if category_key in ORDER_PROFILES else detect_profile(order_data_for_detection)
    )
    profile_cfg = get_profile(selected_profile_id)

    # Nếu category_key chưa truyền, lấy từ profile
    if not category_key:
        category_key = profile_cfg.get("category_key", "phong_sach") if profile_cfg else "phong_sach"

    category_cfg = CATEGORIES.get(category_key, CATEGORIES["phong_sach"])

    # 2. Chuẩn hóa thông tin hồ sơ
    # [NEW 27/10/2026] blank_doc_number=True (hoặc doc_number='BLANK'/'để trắng'/'...'):
    #   -> Mã hồ sơ = 'Q326D……………' và Ngày đăng ký = '…./…./YYYY' để điền tay; KHÔNG sinh mã ngẫu nhiên.
    if blank_doc_number or is_blank_doc_number(doc_number):
        blank_doc_number = True
        doc_number = DOC_NUMBER_BLANK_PLACEHOLDER
        if not registration_date or is_blank_doc_number(registration_date):
            registration_date = get_default_execution_date_dotted()
            date = None
    final_doc_number = doc_number.strip() if doc_number and str(doc_number).strip() and str(doc_number).strip() != "None" else auto_generate_doc_number()
    final_reg_date = normalize_date_str(registration_date) if registration_date and str(registration_date).strip() and str(registration_date).strip() != "None" else (
        normalize_date_str(date) if date and str(date).strip() and str(date).strip() != "None" else auto_generate_date()
    )
    final_exec_date = normalize_date_str(execution_date) if execution_date and str(execution_date).strip() and str(execution_date).strip() != "None" else get_default_execution_date_dotted()

    final_customer_name = normalize_customer_name(customer_name)
    final_address = normalize_address(address)
    final_location = normalize_location(location, final_address)
    final_quote_number = str(quotation_number or quote_number or "").strip() or "QONV2……………"

    # Nội dung yêu cầu
    if request_content and str(request_content).strip() and str(request_content).strip() != "None":
        final_request = clean_text_no_status_or_quantitative(str(request_content).strip())
    elif profile_cfg and profile_cfg.get("default_request"):
        final_request = profile_cfg["default_request"]
    else:
        final_request = category_cfg.get("default_request", "")

    # Chuẩn hóa danh sách indicators (bổ sung chủ thể lấy mẫu/đo, is_measure và is_sample chuẩn xác)
    normalized_indicators = []
    if indicators:
        for ind in indicators:
            if isinstance(ind, dict):
                ind_copy = dict(ind)
                cleaned_name = clean_indicator_name(ind_copy.get("name", ""), category_key=category_key)
                ind_copy["name"] = cleaned_name
                if "is_measure" not in ind_copy or "is_sample" not in ind_copy:
                    m, s = infer_method_type(cleaned_name)
                    ind_copy.setdefault("is_measure", m)
                    ind_copy.setdefault("is_sample", s)
                normalized_indicators.append(ind_copy)
            else:
                cleaned_name = clean_indicator_name(str(ind), category_key=category_key)
                m, s = infer_method_type(cleaned_name)
                normalized_indicators.append({
                    "name": cleaned_name,
                    "is_measure": m,
                    "is_sample": s
                })
    elif profile_cfg and profile_cfg.get("indicators"):
        normalized_indicators = profile_cfg["indicators"]
    else:
        normalized_indicators = category_cfg.get("indicators", [])

    # Nếu truyền danh sách indicators rõ ràng mà num_samples để mặc định là 4
    if indicators and len(normalized_indicators) > 0 and num_samples == 4 and len(normalized_indicators) != 4:
        num_samples = len(normalized_indicators)

    parsed_cabinets = None
    if cabinets:
        if isinstance(cabinets, (list, tuple)):
            parsed_cabinets = [str(c).strip() for c in cabinets if str(c).strip()]
        else:
            parsed_cabinets = [c.strip() for c in str(cabinets).replace(";", "\n").split("\n") if c.strip()]
        if parsed_cabinets and num_positions == 1 and len(parsed_cabinets) > 1:
            num_positions = len(parsed_cabinets)

    # Tự động giải quyết thiết bị và quy trình áp dụng chuẩn SOP
    req_for_sop = f"{final_request} {' '.join(parsed_cabinets)}" if parsed_cabinets else final_request
    final_equipment, final_procedure = resolve_equipment_and_procedure(
        category_key=category_key,
        indicators=normalized_indicators,
        request_content=req_for_sop,
        custom_equipment=equipment,
        custom_procedure=procedure
    )

    customer_info = {
        "customer_name": final_customer_name,
        "address": final_address,
        "location": final_location,
        "doc_number": final_doc_number,
        "date": final_exec_date,
        "registration_date": final_reg_date,
        "execution_date": final_exec_date,
        "request_content": final_request,
        "equipment": final_equipment,
        "procedure": final_procedure,
        "standard_procedure": final_procedure,
        "indicators": normalized_indicators,
        "short_customer_name": short_customer_name,
        "short_content": short_content,
        "num_positions": num_positions,
        "cabinets": parsed_cabinets,
        "blank_cabinets": blank_cabinets,
        "num_vskk_bm01": num_vskk_bm01,
        "bm01_rows": bm01_rows,  # [14/11/2026] số dòng bảng đo thứ cấp BM3030.01, vd "air_velocity:3;filter_dp:3"
        "micro_targets": micro_targets,
        "quotation_number": final_quote_number,
        "quote_number": final_quote_number,
        # [NEW 09/11/2026] Số PO của khách hàng: dùng cho Mục 9 BM3030.02 khi không có số báo giá QONV
        "po_number": str(po_number or "").strip(),
        "production_lines": production_lines
    }

    # 3. Xác định thư mục output theo nguyên tắc: "số hồ sơ - tên khách hàng - nội dung vắn tắc" (Quy tắc 9)
    out_dir = output_base_dir if output_base_dir else OUTPUT_DIR
    subfolder_name, clean_short_kh, clean_short_cnt = build_output_folder_name(
        doc_number=final_doc_number,
        customer_name=customer_name,
        short_customer_name=short_customer_name,
        short_content=short_content,
        request_content=final_request,
        indicators=normalized_indicators,
        category_key=category_key
    )
    target_dir = os.path.join(out_dir, subfolder_name)
    os.makedirs(target_dir, exist_ok=True)

    customer_info["short_customer_name"] = clean_short_kh
    customer_info["short_content"] = clean_short_cnt

    clean_kh = sanitize_for_filename(clean_short_kh) if clean_short_kh else sanitize_filename(customer_name)

    results = {
        "success": True,
        "profile": selected_profile_id,
        "category": category_cfg["name"],
        "customer_name": final_customer_name,
        "short_customer_name": clean_short_kh,
        "short_content": clean_short_cnt,
        "subfolder_name": subfolder_name,
        "doc_number": final_doc_number,
        "num_positions": num_positions,
        "cabinets": parsed_cabinets,
        "date": final_exec_date,
        "registration_date": final_reg_date,
        "execution_date": final_exec_date,
        "output_dir": target_dir,
        "z_drive_copy": None,
        "bm02_file": None,
        "bm01_file": None,
        "bm02_pdf": None,
        "bm01_pdf": None,
        "errors": []
    }

    # 4. Nhân bản BM3030.02 (Biên bản giám định hiện trường / BBLM)
    if generate_bm02:
        template_bm02 = custom_template_bm02 if custom_template_bm02 else category_cfg.get("template_bm02_source")
        if not template_bm02 or not os.path.exists(template_bm02):
            # Fallback sang mẫu phòng sạch tiêu chuẩn
            template_bm02 = os.path.join(BASE_DIR, "Phong sach, Vi sinh", "BM3030.02- BBLM - Q326D108033.docx")

        if template_bm02 and os.path.exists(template_bm02):
            try:
                out_bm02_name = f"{doc_number_for_filename(final_doc_number)} - BM3030.02 - BBLM - {clean_kh}.docx"
                out_bm02_path = os.path.join(target_dir, out_bm02_name)

                build_bm3030_02(
                    template_path=template_bm02,
                    customer_info=customer_info,
                    indicators=normalized_indicators,
                    num_samples=num_samples,
                    category_key=category_key,
                    output_path=out_bm02_path
                )
                results["bm02_file"] = out_bm02_path
            except Exception as e:
                results["errors"].append(f"Lỗi nhân bản BM3030.02: {str(e)}")
        else:
            results["errors"].append(f"Không tìm thấy mẫu BM3030.02: {template_bm02}")

    # 5. Nhân bản BM3030.01 (Phiếu ghi chép hiện trường)
    if generate_bm01:
        try:
            out_bm01_name = f"{doc_number_for_filename(final_doc_number)} - BM3030.01 - Ghi chep - {clean_kh}.docx"
            out_bm01_path = os.path.join(target_dir, out_bm01_name)

            build_bm3030_01(
                customer_info=customer_info,
                category_key=category_key,
                indicators=normalized_indicators,
                num_samples=num_samples_bm01 if num_samples_bm01 is not None else num_samples,
                num_positions=num_positions,
                output_path=out_bm01_path
            )
            results["bm01_file"] = out_bm01_path
        except Exception as e:
            results["errors"].append(f"Lỗi nhân bản BM3030.01: {str(e)}")

    # 5b. [NEW 27/10/2026] Mã hồ sơ để trắng -> nới ô mã số trên header cho placeholder không bị rớt dòng
    if blank_doc_number:
        for _f in (results.get("bm02_file"), results.get("bm01_file")):
            if _f and os.path.exists(_f) and _f.lower().endswith(".docx"):
                try:
                    widen_blank_doc_number_cell(_f)
                except Exception as e:
                    results["errors"].append(f"Lỗi nới ô mã hồ sơ: {str(e)}")

    # 6. Tùy chọn xuất PDF tự động (Word COM + Vector PDF ghép)
    if export_pdf:
        try:
            export_order_pdfs(results, customer_info=customer_info)
        except Exception as e:
            results["errors"].append(f"Lỗi xuất PDF: {str(e)}")

    # 7. Cập nhật tiêu đề cuộc hội thoại vào SQLite DB theo Quy tắc 10
    try:
        update_conversation_db_title(
            doc_number=final_doc_number,
            customer_name=customer_name,
            cid=conversation_id
        )
    except Exception as e:
        print(f"[WARN] Không thể cập nhật tiêu đề DB: {e}")

    # 8. [NEW 04/11/2026] Tự động sao chép 1 bản sang ổ đĩa Z theo Quy tắc 13
    if copy_to_z:
        try:
            z_dest = copy_order_to_z_drive(
                target_dir=target_dir,
                subfolder_name=subfolder_name,
                z_base_dir=z_drive_dir,
                scan_files=scan_files
            )
            results["z_drive_copy"] = z_dest
        except Exception as e:
            print(f"[WARN] Lỗi khi sao chép sang ổ Z: {e}")
            results["z_drive_copy"] = None

    if results["errors"] and not results["bm02_file"] and not results["bm01_file"]:
        results["success"] = False

    return results


def process_order(order_data, export_pdf=False, conversation_id=None, copy_to_z=True):
    """
    Hàm giao diện cấp cao (Unified API) nhận vào dictionary thông tin đơn hàng
    và tự động sinh toàn bộ hồ sơ. Tương thích linh hoạt với mọi biến thể tên trường.
    """
    cat = (
        order_data.get("category_key")
        or order_data.get("category")
        or order_data.get("template_type")
    )
    doc_num = (
        order_data.get("doc_number")
        or order_data.get("request_code")
        or order_data.get("so_yeu_cau")
        or order_data.get("ho_so_so")
    )
    addr = order_data.get("address") or order_data.get("customer_address", "")
    loc = order_data.get("location") or order_data.get("execution_location")
    reg_date = (
        order_data.get("registration_date")
        or order_data.get("reg_date")
        or order_data.get("ngay_dang_ky")
        or order_data.get("date")
    )
    exec_date = (
        order_data.get("execution_date")
        or order_data.get("exec_date")
        or order_data.get("ngay_thuc_hien")
    )
    quote_num = (
        order_data.get("quotation_number")
        or order_data.get("quote_number")
        or order_data.get("bao_gia")
        or order_data.get("so_bao_gia")
    )

    return clone_form_set(
        category_key=cat,
        customer_name=order_data.get("customer_name", ""),
        address=addr,
        location=loc,
        num_samples=order_data.get("num_samples", 4),
        num_samples_bm01=order_data.get("num_samples_bm01"),
        num_positions=order_data.get("num_positions", 1),
        doc_number=doc_num,
        date=order_data.get("date"),
        registration_date=reg_date,
        execution_date=exec_date,
        request_content=order_data.get("request_content"),
        equipment=order_data.get("equipment"),
        procedure=order_data.get("procedure"),
        indicators=order_data.get("indicators"),
        output_base_dir=order_data.get("output_dir"),
        short_customer_name=order_data.get("short_customer_name"),
        short_content=order_data.get("short_content"),
        profile=order_data.get("profile"),
        export_pdf=export_pdf or order_data.get("export_pdf", False),
        conversation_id=conversation_id,
        blank_doc_number=bool(order_data.get("blank_doc_number", False)),
        blank_cabinets=bool(order_data.get("blank_cabinets", False) or order_data.get("blank_cabinet_name", False)),
        cabinets=order_data.get("cabinets"),
        num_vskk_bm01=order_data.get("num_vskk_bm01"),
        bm01_rows=order_data.get("bm01_rows"),
        micro_targets=order_data.get("micro_targets"),
        quotation_number=quote_num,
        po_number=order_data.get("po_number") or order_data.get("po") or order_data.get("so_po"),
        scan_files=order_data.get("scan_files") or order_data.get("scan") or order_data.get("input_files"),
        copy_to_z=order_data.get("copy_to_z", copy_to_z),
        z_drive_dir=order_data.get("z_drive_dir")
    )
