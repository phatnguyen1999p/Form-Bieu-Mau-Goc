"""
Sinh hồ sơ Quan trắc môi trường lao động (MTLĐ) theo KẾ HOẠCH TỪNG PHÒNG (CLI).

  - BM3030.01 Ghi chép: mỗi phòng 1 trang (chọn máy ☐, đo trực tiếp có ô xám, ồn ghi tên file đo,
    lấy mẫu hơi khí độc có cột Bơm, Ecgonomy OWAS) + bảng theo dõi khối lượng.
  - BM3030.02 Biên bản: mỗi chỉ tiêu 1 dòng, ghi sẵn các phòng thực hiện.

Đầu vào: file Excel kế hoạch (mẫu: templates/MTLD_KeHoach_mau.xlsx) gồm các cột
  Phòng | Chỉ tiêu | SL | Điểm đo  (vd Điểm đo: "P304: 02; P303: 01").

Ví dụ:
  python generate_mtld.py --plan "Input/KeHoach.xlsx" --customer "CÔNG TY TNHH ABC" ^
      --address "Lô A1, KCN Tân Bình, TP.HCM" --doc Q326D145000 --reg_date 18/11/2026 --pdf
  (bỏ --doc để mã hồ sơ để trắng Q326D……………)
"""
import os
import sys
import json
import argparse

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core.mtld_builder import generate_mtld


def parse_fixes(s):
    """'QT|Ồn chung=B201:1;P304:1 || HTD|Ồn chung=A106:1;Phòng còn lại:2'"""
    fixes = {}
    for part in [x.strip() for x in (s or "").split("||") if x.strip()]:
        key, val = part.split("=", 1)
        room, ct = [x.strip() for x in key.split("|", 1)]
        locs = []
        for seg in [x.strip() for x in val.split(";") if x.strip()]:
            name, n = seg.rsplit(":", 1)
            locs.append((name.strip(), int(n)))
        fixes[(room, ct)] = locs
    return fixes


def main():
    ap = argparse.ArgumentParser(description="Sinh Ghi chép + Biên bản MTLĐ theo kế hoạch từng phòng")
    ap.add_argument("--plan", required=True, help="File Excel kế hoạch (cột Phòng | Chỉ tiêu | SL | Điểm đo)")
    ap.add_argument("--sheet", default=None, help="Tên sheet kế hoạch (mặc định tự tìm)")
    ap.add_argument("--customer", required=True, help="Đơn vị yêu cầu")
    ap.add_argument("--address", required=True, help="Địa chỉ (tự chuẩn hóa)")
    ap.add_argument("--short_customer", default=None, help="Tên khách hàng rút gọn cho tên thư mục/file")
    ap.add_argument("--doc", "--doc_number", dest="doc", default=None, help="Mã hồ sơ Q326D... (bỏ trống = để trắng)")
    ap.add_argument("--reg_date", default=None, help="Ngày đăng ký (DD/MM/YYYY), mặc định để chấm")
    ap.add_argument("--quote", default=None, help="Số báo giá QONV... (mặc định QONV2……………)")
    ap.add_argument("--request", default=None, help="Nội dung yêu cầu (mặc định: Quan trắc môi trường lao động)")
    ap.add_argument("--representative", default=None, help="Câu mục 9: 'được chỉ định bởi đại diện ...' (mặc định = đơn vị yêu cầu)")
    ap.add_argument("--condition", default=None, help="Câu ghi chú tình trạng hoạt động trong biên bản")
    ap.add_argument("--plan_title", default=None, help="Ghi chú cạnh tiêu đề bảng theo dõi (vd 'theo Kế hoạch QTMTLĐ 2026')")
    ap.add_argument("--multi_day", action="store_true", help="Thời gian thực hiện dạng 'từ ngày … kết thúc ngày …'")
    ap.add_argument("--fix", default=None, help="Sửa điểm đo chưa rõ: 'Phòng|Chỉ tiêu=Khu vực:SL;Khu vực:SL || ...'")
    ap.add_argument("--scan", default=None, help="File scan gốc chép vào thư mục Scan trên Z (nhiều file cách ';')")
    ap.add_argument("--pdf", action="store_true", help="Xuất PDF (biên bản PDF chứa sẵn 2 bản)")
    ap.add_argument("--no_z_copy", action="store_true", help="Không chép sang ổ Z")
    ap.add_argument("--out_root", default=None, help="Thư mục gốc lưu (mặc định output/)")
    ap.add_argument("--json", action="store_true", help="In kết quả dạng JSON")
    a = ap.parse_args()

    res = generate_mtld(
        a.plan, a.customer, a.address, doc_number=a.doc, short_customer=a.short_customer, request=a.request,
        quote=a.quote, reg_date=a.reg_date, representative=a.representative, multi_day=a.multi_day,
        plan_title=a.plan_title, condition_note=a.condition, fixes=parse_fixes(a.fix), sheet=a.sheet,
        out_root=a.out_root, export_pdf=a.pdf, copy_to_z=not a.no_z_copy, scan_files=a.scan,
    )
    if a.json:
        print(json.dumps(res, ensure_ascii=False, indent=2))
    else:
        print(f"Thư mục: {res['output_dir']}")
        print(f"Phòng: {res['rooms']} | Tổng chỉ tiêu: {res['total']}")
        for k in ("bm01_file", "bm02_file", "bm01_pdf", "bm02_pdf", "z_dir"):
            if res.get(k):
                print(f"{k}: {res[k]}")


if __name__ == "__main__":
    main()
