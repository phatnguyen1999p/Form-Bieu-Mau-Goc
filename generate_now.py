"""
Công cụ thực thi nhanh việc nhân bản biểu mẫu hiện trường QUATEST 3 (CLI Tool)
Dành cho Chatbot, Agent hoặc người dùng gọi trực tiếp từ dòng lệnh.
Hỗ trợ Profile mẫu từ các hồ sơ thực tế, tự động nhận diện và xuất cả Word lẫn PDF vector.
"""
import os
import sys
import json
import argparse

try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core import (
    clone_form_set,
    process_order,
    ORDER_PROFILES,
    get_profile,
    detect_profile
)


def main():
    parser = argparse.ArgumentParser(description="Sinh biểu mẫu hiện trường QUATEST 3 chuẩn mực")
    parser.add_argument("--customer", help="Tên khách hàng / Đơn vị yêu cầu")
    parser.add_argument("--address", help="Địa chỉ khách hàng")
    parser.add_argument("--location", default=None, help="Địa điểm thực hiện (mặc định lấy theo địa chỉ chuẩn hóa)")
    parser.add_argument("--num_samples", type=int, default=4, help="Số lượng chỉ tiêu / mẫu")
    parser.add_argument("--num_positions", type=int, default=1, help="Số lượng vị trí đo (ví dụ: số vị trí khí nén)")
    parser.add_argument("--category", default=None, help="Mã lĩnh vực (phong_sach, khi_nen, bsc, co2_mtld, kho_lanh...)")
    parser.add_argument("--profile", default=None, help="Mã profile mẫu (bsc, khi_nen, phong_sach, khi_nen_phong_sach...)")
    parser.add_argument("--doc_number", default=None, help="Mã số hồ sơ (ví dụ: Q326D132918)")
    parser.add_argument("--blank_doc", action="store_true", help="Để trắng mã hồ sơ (Q326D……………) và ngày đăng ký (…./…./YYYY) để điền tay")
    parser.add_argument("--num_samples_bm01", type=int, default=None, help="Số dòng đo bụi phòng sạch trong BM3030.01 (mặc định 8)")
    parser.add_argument("--num_vskk_bm01", type=int, default=None, help="Số dòng lấy mẫu vi sinh không khí trong BM3030.01")
    parser.add_argument("--bm01_rows", default=None, help="Số dòng từng bảng đo thứ cấp BM3030.01, ví dụ: 'air_velocity:3;filter_dp:3;vskk_settle:7;contact_plate:7'")
    parser.add_argument("--cabinets", default=None, help="Danh sách tên tủ BSC phân cách bởi dấu chấm phẩy (;)")
    parser.add_argument("--blank_cabinets", action="store_true", help="Chừa hàng chấm lửng (...............................) cho phần tên tủ BSC")
    parser.add_argument("--micro_targets", default=None, help="Chỉ tiêu vi sinh thực hiện cụ thể (ví dụ: Tổng số vi sinh vật hiếu khí)")
    parser.add_argument("--production_lines", default=None, help="Danh sách dây chuyền sản xuất / phân xưởng (ví dụ: 'Nắp DMCA01:9; Phôi DMPA01:12')")
    parser.add_argument("--quote", "--quotation_number", "--quote_number", dest="quotation_number", default=None, help="Mã số báo giá đầy đủ (ví dụ: QONV26015439, mặc định: QONV2……………)")
    parser.add_argument("--po", "--po_number", dest="po_number", default=None, help="Số đơn đặt hàng (PO) của khách hàng, dùng cho Mục 9 khi không có số báo giá QONV (ví dụ: E900192188)")
    parser.add_argument("--scan", dest="scan_files", default=None, help="File gốc ban đầu (phiếu yêu cầu/PO scan) chép vào thư mục 'Scan' trên ổ Z; nhiều file cách nhau bởi ';' (ví dụ: Input/20261109155059847.pdf)")
    parser.add_argument("--date", default=None, help="Ngày đăng ký hồ sơ (DD/MM/YYYY)")
    parser.add_argument("--reg_date", default=None, help="Ngày đăng ký nhận hồ sơ (DD/MM/YYYY)")
    parser.add_argument("--exec_date", default=None, help="Ngày thực hiện hiện trường (mặc định để chấm lửng …./…./2026)")
    parser.add_argument("--request", default=None, help="Nội dung yêu cầu")
    parser.add_argument("--equipment", default=None, help="Danh sách thiết bị sử dụng")
    parser.add_argument("--procedure", default=None, help="Quy trình áp dụng")
    parser.add_argument("--indicators", default=None, help="Danh sách chỉ tiêu cụ thể phân cách bởi dấu chấm phẩy (;)")
    parser.add_argument("--short_customer", default=None, help="Tên rút gọn của khách hàng")
    parser.add_argument("--short_content", default=None, help="Nội dung vắn tắt cho tên thư mục")
    parser.add_argument("--pdf", action="store_true", help="Tự động xuất cả bộ hồ sơ sang PDF")
    parser.add_argument("--no_z_copy", action="store_true", help="Không sao chép sang ổ đĩa Z")
    parser.add_argument("--z_drive_dir", default=None, help="Đường dẫn tùy chỉnh ổ đĩa Z (mặc định: Z:\\9.NV2 - MOI TRUONG\\1.Ho So Moi Truong\\2026)")
    parser.add_argument("--json_file", default=None, help="Đường dẫn file JSON chứa thông tin đơn hàng đầy đủ")
    parser.add_argument("--json", action="store_true", help="Xuất kết quả dưới dạng JSON")

    args = parser.parse_args()

    # Trường hợp truyền qua file JSON
    if args.json_file:
        if not os.path.exists(args.json_file):
            print(f"❌ Không tìm thấy file JSON: {args.json_file}")
            sys.exit(1)
        with open(args.json_file, "r", encoding="utf-8") as f:
            order_data = json.load(f)
        if args.pdf:
            order_data["export_pdf"] = True
        if args.blank_doc:
            order_data["blank_doc_number"] = True
        if args.no_z_copy:
            order_data["copy_to_z"] = False
        if args.z_drive_dir:
            order_data["z_drive_dir"] = args.z_drive_dir
        res = process_order(order_data)
    else:
        if not args.customer or not args.address:
            parser.print_help()
            sys.exit(1)

        # Parse indicators string if provided
        parsed_indicators = None
        if args.indicators:
            items = [i.strip() for i in args.indicators.replace(";", "\n").split("\n") if i.strip()]
            if items:
                parsed_indicators = items

        # Parse production_lines if provided
        parsed_production_lines = None
        if args.production_lines:
            try:
                parsed_production_lines = json.loads(args.production_lines)
            except Exception:
                # Format: "Nắp DMCA01:9; Phôi DMPA01:12"
                pl_items = []
                for p_item in args.production_lines.split(";"):
                    p_item = p_item.strip()
                    if not p_item:
                        continue
                    if ":" in p_item:
                        p_name, p_s = p_item.rsplit(":", 1)
                        try:
                            pl_items.append({"name": p_name.strip(), "samples": int(p_s.strip())})
                        except ValueError:
                            pl_items.append({"name": p_item, "samples": 10})
                    else:
                        pl_items.append({"name": p_item, "samples": 10})
                if pl_items:
                    parsed_production_lines = pl_items

        res = clone_form_set(
            category_key=args.category,
            profile=args.profile,
            customer_name=args.customer,
            address=args.address,
            location=args.location,
            num_samples=args.num_samples,
            num_positions=args.num_positions,
            doc_number=args.doc_number,
            date=args.date,
            registration_date=args.reg_date or args.date,
            execution_date=args.exec_date,
            request_content=args.request,
            equipment=args.equipment,
            procedure=args.procedure,
            indicators=parsed_indicators,
            short_customer_name=args.short_customer,
            short_content=args.short_content,
            export_pdf=args.pdf,
            blank_doc_number=args.blank_doc,
            blank_cabinets=args.blank_cabinets,
            num_samples_bm01=args.num_samples_bm01,
            num_vskk_bm01=args.num_vskk_bm01,
            bm01_rows=args.bm01_rows,
            cabinets=args.cabinets,
            micro_targets=args.micro_targets,
            quotation_number=args.quotation_number,
            po_number=args.po_number,
            scan_files=args.scan_files,
            production_lines=parsed_production_lines,
            copy_to_z=not args.no_z_copy,
            z_drive_dir=args.z_drive_dir
        )

    if args.json:
        print(json.dumps(res, ensure_ascii=False, indent=2))
    else:
        print("=" * 70)
        if res.get("success"):
            print(f"✅ ĐÃ TẠO THÀNH CÔNG BỘ HỒ SƠ QUATEST 3 CHO: {res['customer_name']}")
            print(f"📁 Thư mục lưu trữ: {res['output_dir']}")
            if res.get("z_drive_copy"):
                print(f"💾 Đã sao chép 1 bản sang ổ Z: {res['z_drive_copy']}")
            print(f"📄 BM3030.02 (Biên bản hiện trường): {os.path.basename(res['bm02_file'])}")
            if res.get("bm01_file"):
                print(f"📝 BM3030.01 (Phiếu ghi chép hiện trường): {os.path.basename(res['bm01_file'])}")
            if res.get("bm02_pdf"):
                print(f"📑 BM3030.02 PDF: {os.path.basename(res['bm02_pdf'])} ({res.get('bm02_page_count', 0)} trang/bản × {res.get('bm02_pdf_copies', 1)} bản = {res.get('bm02_pdf_total_pages', res.get('bm02_page_count', 0))} trang)")
            if res.get("bm01_pdf"):
                print(f"📑 BM3030.01 PDF: {os.path.basename(res['bm01_pdf'])} ({res.get('bm01_page_count', 0)} trang)")
            print(f"📌 Mã hồ sơ: {res['doc_number']} | Profile: {res.get('profile')} | Thư mục: {res['subfolder_name']}")
        else:
            print("❌ Có lỗi xảy ra trong quá trình sinh hồ sơ:")
            for err in res.get("errors", []):
                print(f"  - {err}")
        print("=" * 70)


if __name__ == "__main__":
    main()
