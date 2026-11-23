"""
Giao diện dòng lệnh (CLI) để nhân bản biểu mẫu QUATEST 3
Hỗ trợ cả chế độ nhập tương tác (interactive) và truyền tham số command-line
"""
import os
import sys
import argparse
import subprocess

try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

from core.config import CATEGORIES, BASE_DIR, OUTPUT_DIR
from core.form_cloner import clone_form_set, auto_generate_doc_number, auto_generate_date
from batch_generator import process_batch_excel

def print_banner():
    print("=" * 65)
    print("   HỆ THỐNG NHÂN BẢN VÀ TÁI CHẾ BIỂU MẪU HIỆN TRƯỜNG QUATEST 3   ")
    print("   (Biên bản giám định hiện trường BM3030.02 & Phiếu ghi chép BM3030.01)")
    print("=" * 65)

def interactive_mode():
    print_banner()
    print("\n[CHỌN LĨNH VỰC NGHIỆP VỤ]")
    cat_keys = list(CATEGORIES.keys())
    for idx, key in enumerate(cat_keys, 1):
        print(f"  {idx}. {CATEGORIES[key]['name']}")
    
    while True:
        choice = input("\nNhập số thứ tự lĩnh vực (1-5) [mặc định 1]: ").strip()
        if not choice:
            cat_key = cat_keys[0]
            break
        if choice.isdigit() and 1 <= int(choice) <= len(cat_keys):
            cat_key = cat_keys[int(choice) - 1]
            break
        print("Lựa chọn không hợp lệ, vui lòng nhập lại.")
        
    cfg = CATEGORIES[cat_key]
    print(f"\n=> Đã chọn: {cfg['name']}")
    
    print("\n[NHẬP THÔNG TIN KHÁCH HÀNG]")
    while True:
        customer_name = input("1. Tên khách hàng / Đơn vị yêu cầu (*): ").strip()
        if customer_name:
            break
        print("Tên khách hàng không được để trống!")
        
    address = input("2. Địa chỉ đơn vị yêu cầu: ").strip()
    if not address:
        address = "………………………………………………………………………………………………"
        
    location = input("3. Địa điểm thực hiện giám định / lấy mẫu: ").strip()
    if not location:
        location = address
        
    while True:
        raw_num = input("4. Số lượng chỉ tiêu / vị trí lấy mẫu [mặc định 4]: ").strip()
        if not raw_num:
            num_samples = 4
            break
        if raw_num.isdigit() and int(raw_num) > 0:
            num_samples = int(raw_num)
            break
        print("Vui lòng nhập số nguyên dương!")
        
    default_doc_num = auto_generate_doc_number()
    doc_number = input(f"5. Mã số phiếu / hồ sơ [mặc định {default_doc_num}]: ").strip()
    if not doc_number:
        doc_number = default_doc_num
        
    default_date = auto_generate_date()
    date_val = input(f"6. Ngày thực hiện [mặc định {default_date}]: ").strip()
    if not date_val:
        date_val = default_date

    print("\n[CHỌN CHỈ TIÊU]")
    print("Danh mục chỉ tiêu mẫu có sẵn:")
    def_inds = cfg.get("indicators", [])
    for idx, ind in enumerate(def_inds, 1):
        print(f"  [{idx}] {ind['name']}")
    use_custom = input("Bạn có muốn dùng danh mục mẫu trên không? (Y/n): ").strip().lower()
    selected_indicators = None
    if use_custom == 'n':
        print("Nhập tên các chỉ tiêu cụ thể (mỗi dòng 1 chỉ tiêu, gõ dòng trống để kết thúc):")
        custom_inds = []
        while True:
            line = input(f"  Chỉ tiêu #{len(custom_inds)+1}: ").strip()
            if not line:
                break
            custom_inds.append({"name": line, "is_measure": True, "is_sample": False})
        if custom_inds:
            selected_indicators = custom_inds
            num_samples = max(num_samples, len(selected_indicators))

    print("\nĐang tiến hành nhân bản tài liệu...")
    res = clone_form_set(
        category_key=cat_key,
        customer_name=customer_name,
        address=address,
        location=location,
        num_samples=num_samples,
        doc_number=doc_number,
        date=date_val,
        indicators=selected_indicators
    )
    
    if res["success"]:
        print("\n" + "=" * 65)
        print("   NHÂN BẢN HỒ SƠ THÀNH CÔNG!   ")
        print("=" * 65)
        print(f"Thư mục lưu: {res['output_dir']}")
        if res.get("z_drive_copy"):
            print(f"- Đã sao chép sang ổ Z: {res['z_drive_copy']}")
        if res["bm02_file"]:
            print(f"- Biên bản BM3030.02: {os.path.basename(res['bm02_file'])}")
        if res["bm01_file"]:
            print(f"- Phiếu ghi chép BM3030.01: {os.path.basename(res['bm01_file'])}")
            
        open_folder = input("\nBạn có muốn mở thư mục chứa file vừa tạo không? (Y/n): ").strip().lower()
        if open_folder != 'n':
            os.startfile(res["output_dir"])
    else:
        print("\nCó lỗi xảy ra trong quá trình nhân bản:")
        for err in res["errors"]:
            print(f"- {err}")

def main():
    parser = argparse.ArgumentParser(description="Nhân bản biểu mẫu giám định & ghi chép hiện trường QUATEST 3")
    parser.add_argument("--batch", type=str, help="Đường dẫn file Excel để chạy hàng loạt")
    parser.add_argument("--category", type=str, choices=list(CATEGORIES.keys()), help="Khóa lĩnh vực (phong_sach, khi_nen, co2_mtld, kho_lanh, bsc)")
    parser.add_argument("--customer", type=str, help="Tên khách hàng / Đơn vị yêu cầu")
    parser.add_argument("--address", type=str, default="", help="Địa chỉ khách hàng")
    parser.add_argument("--location", type=str, default="", help="Địa điểm thực hiện")
    parser.add_argument("--num-samples", type=int, default=4, help="Số lượng chỉ tiêu / lấy mẫu")
    parser.add_argument("--doc-number", type=str, default="", help="Mã số hồ sơ / Số phiếu")
    parser.add_argument("--date", type=str, default="", help="Ngày thực hiện (DD/MM/YYYY)")
    parser.add_argument("--short-name", type=str, default=None, help="Tên khách hàng rút gọn (tùy chọn)")
    parser.add_argument("--short-content", type=str, default=None, help="Nội dung vắn tắt (tùy chọn, VD: Phòng sạch, Nước uống, vskk - swab)")
    parser.add_argument("--profile", type=str, default=None, help="Preset profile (bsc, khi_nen, phong_sach, phong_sach_vskk_swab, khi_nen_phong_sach, co2_mtld, kho_lanh)")
    parser.add_argument("--positions", type=int, default=1, help="Số lượng vị trí / điểm đo (áp dụng cho khí nén, BSC...)")
    parser.add_argument("--samples-bm01", type=int, default=None, help="Số lượng dòng/vị trí đo cho BM3030.01")
    parser.add_argument("--pdf", action="store_true", help="Tự động xuất tệp PDF và ghép vector chất lượng cao")
    parser.add_argument("--output", type=str, default=None, help="Thư mục output tùy chọn")

    args = parser.parse_args()

    # Chế độ chạy hàng loạt từ Excel
    if args.batch:
        print_banner()
        print(f"Chạy chế độ hàng loạt từ file: {args.batch}")
        summary = process_batch_excel(args.batch, output_dir=args.output, export_pdf=args.pdf)
        print(f"\nKết quả: Thành công {summary['success']}/{summary['total']} hồ sơ (Thất bại: {summary['failed']})")
        return

    # Chế độ có tham số
    if args.customer:
        cat_key = args.category if args.category else "phong_sach"
        res = clone_form_set(
            category_key=cat_key,
            customer_name=args.customer,
            address=args.address if args.address else args.customer,
            location=args.location if args.location else (args.address if args.address else args.customer),
            num_samples=args.num_samples,
            num_samples_bm01=args.samples_bm01,
            num_positions=args.positions,
            doc_number=args.doc_number if args.doc_number else None,
            date=args.date if args.date else None,
            output_base_dir=args.output,
            short_customer_name=args.short_name,
            short_content=args.short_content,
            profile=args.profile,
            export_pdf=args.pdf
        )
        if res["success"]:
            print(f"Nhân bản thành công vào: {res['output_dir']}")
            if res.get("z_drive_copy"):
                print(f"Đã sao chép sang ổ Z: {res['z_drive_copy']}")
        else:
            print(f"Lỗi: {res['errors']}")
        return

    # Nếu không có tham số nào, vào chế độ nhập tương tác
    interactive_mode()

if __name__ == "__main__":
    main()
