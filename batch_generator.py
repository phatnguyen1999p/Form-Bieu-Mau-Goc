"""
Module xử lý nhân bản hàng loạt từ file Excel danh sách khách hàng
"""
import os
import sys
import openpyxl

try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

from core.form_cloner import clone_form_set
from core.config import CATEGORIES, BASE_DIR
from core.order_profiles import detect_profile

def map_category_string(cat_str):
    """Ánh xạ chuỗi tên lĩnh vực người dùng nhập sang category_key chuẩn."""
    if not cat_str:
        return "phong_sach"
    s = str(cat_str).lower().strip()
    if any(k in s for k in ["bsc", "tủ an toàn", "tủ sinh học"]):
        return "bsc"
    elif any(k in s for k in ["khí nén", "nén", "nitơ", "nito", "compressed"]):
        return "khi_nen"
    elif any(k in s for k in ["co2", "mtld", "lao động", "môi trường"]):
        return "co2_mtld"
    elif any(k in s for k in ["kho lạnh", "kho", "lạnh", "cold"]):
        return "kho_lanh"
    elif any(k in s for k in ["sạch", "phòng sạch", "clean", "vi sinh"]):
        return "phong_sach"
    return "phong_sach"

def process_batch_excel(excel_path, output_dir=None, export_pdf=False, progress_callback=None):
    """
    Đọc file Excel và sinh hàng loạt biểu mẫu cho tất cả khách hàng.
    Tự động nhận diện profile và tối ưu layout chuẩn 11 quy tắc.
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Không tìm thấy file Excel: {excel_path}")
        
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        return {"total": 0, "success": 0, "failed": 0, "results": []}
        
    header_row = [str(c).strip().lower() if c else "" for c in rows[0]]
    
    # Tìm chỉ số các cột dựa vào tên cột
    def find_col_idx(keywords):
        for idx, h in enumerate(header_row):
            if any(k in h for k in keywords):
                return idx
        return -1
        
    col_cat = find_col_idx(["lĩnh vực", "biểu mẫu", "loại"])
    col_customer = find_col_idx(["khách hàng", "đơn vị", "tên"])
    col_short_customer = find_col_idx(["tên rút gọn", "tên ngắn", "viết tắt", "rút gọn"])
    col_short_content = find_col_idx(["nội dung vắn tắt", "vắn tắt", "tóm tắt"])
    col_address = find_col_idx(["địa chỉ"])
    col_location = find_col_idx(["địa điểm", "vị trí"])
    col_num = find_col_idx(["số lượng", "chỉ tiêu", "mẫu"])
    col_doc_num = find_col_idx(["mã số", "số phiếu", "hồ sơ"])
    col_date = find_col_idx(["ngày"])
    col_content = find_col_idx(["nội dung"])
    col_note = find_col_idx(["ghi chú", "chi tiết"])

    # Fallback theo vị trí cột mặc định nếu không khớp tiêu đề
    if col_customer == -1: col_customer = 2
    if col_address == -1: col_address = 3
    if col_location == -1: col_location = 4
    if col_num == -1: col_num = 5

    data_rows = rows[1:]
    total = len(data_rows)
    success_count = 0
    failed_count = 0
    batch_results = []

    for idx, r in enumerate(data_rows, 1):
        # Bỏ qua dòng trống
        if not r or not any(r):
            continue
            
        c_name = str(r[col_customer]).strip() if col_customer < len(r) and r[col_customer] is not None else ""
        if not c_name or c_name == "None":
            continue
            
        short_c_name = str(r[col_short_customer]).strip() if col_short_customer != -1 and col_short_customer < len(r) and r[col_short_customer] is not None and str(r[col_short_customer]).strip() != "None" else None
        short_cnt_val = str(r[col_short_content]).strip() if col_short_content != -1 and col_short_content < len(r) and r[col_short_content] is not None and str(r[col_short_content]).strip() != "None" else None

        c_addr = str(r[col_address]).strip() if col_address < len(r) and r[col_address] is not None else ""
        c_loc = str(r[col_location]).strip() if col_location < len(r) and r[col_location] is not None else c_addr
        
        # Số lượng chỉ tiêu
        raw_num = r[col_num] if col_num < len(r) and r[col_num] is not None else 4
        try:
            num_samples = int(float(raw_num))
        except (ValueError, TypeError):
            num_samples = 4
            
        cat_str = str(r[col_cat]) if col_cat != -1 and col_cat < len(r) and r[col_cat] is not None else ""
        cat_key = map_category_string(cat_str)
        
        doc_num = str(r[col_doc_num]).strip() if col_doc_num != -1 and col_doc_num < len(r) and r[col_doc_num] is not None else ""
        date_val = str(r[col_date]).strip() if col_date != -1 and col_date < len(r) and r[col_date] is not None else ""
        req_content = str(r[col_content]).strip() if col_content != -1 and col_content < len(r) and r[col_content] is not None else ""
        
        # Chỉ tiêu chi tiết nếu có
        indicators_list = None
        if col_note != -1 and col_note < len(r) and r[col_note]:
            items = [item.strip() for item in str(r[col_note]).replace(";", "\n").split("\n") if item.strip()]
            if items:
                indicators_list = [{"name": item, "is_measure": True, "is_sample": False} for item in items]
                num_samples = max(num_samples, len(indicators_list))

        # Nhận diện profile tự động
        note_str = str(r[col_note]) if col_note != -1 and col_note < len(r) and r[col_note] else ""
        context_str = f"{cat_str} {req_content} {note_str}"
        detected_profile = detect_profile(context_str)

        try:
            res = clone_form_set(
                category_key=cat_key,
                customer_name=c_name,
                address=c_addr,
                location=c_loc,
                num_samples=num_samples,
                doc_number=doc_num if doc_num and doc_num != "None" else None,
                date=date_val if date_val and date_val != "None" else None,
                request_content=req_content if req_content and req_content != "None" else None,
                indicators=indicators_list,
                output_base_dir=output_dir,
                short_customer_name=short_c_name,
                short_content=short_cnt_val,
                profile=detected_profile,
                export_pdf=export_pdf
            )
            if res["success"]:
                success_count += 1
            else:
                failed_count += 1
            batch_results.append(res)
        except Exception as e:
            failed_count += 1
            batch_results.append({
                "success": False,
                "customer_name": c_name,
                "errors": [str(e)]
            })

        if progress_callback:
            progress_callback(idx, total, c_name)

    return {
        "total": len(batch_results),
        "success": success_count,
        "failed": failed_count,
        "results": batch_results
    }

if __name__ == "__main__":
    sample_file = os.path.join(BASE_DIR, "danh_sach_khach_hang_mau.xlsx")
    print(f"Bắt đầu xử lý file mẫu: {sample_file}")
    
    def on_prog(idx, total, name):
        print(f"[{idx}/{total}] Đã xử lý: {name}")
        
    summary = process_batch_excel(sample_file, progress_callback=on_prog)
    print("\n" + "="*50)
    print(f"TỔNG KẾT: Thành công {summary['success']}/{summary['total']} hồ sơ (Thất bại: {summary['failed']})")
