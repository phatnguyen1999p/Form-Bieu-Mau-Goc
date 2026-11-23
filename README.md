# HỆ THỐNG NHÂN BẢN VÀ TÁI CHẾ BIỂU MẪU HIỆN TRƯỜNG QUATEST 3

Hệ thống mã nguồn tự động hóa việc nhân bản, làm sạch và tái sử dụng các biểu mẫu giám định & lấy mẫu hiện trường của **Trung tâm Kỹ thuật Tiêu chuẩn Đo lường Chất lượng 3 (QUATEST 3)**.

Mỗi khi bạn chuẩn bị đi lấy mẫu hoặc giám định cho một khách hàng mới, chỉ cần cung cấp:
- **Tên khách hàng / Đơn vị yêu cầu**
- **Địa chỉ khách hàng**
- **Địa điểm thực hiện giám định / lấy mẫu**
- **Số lượng chỉ tiêu / mẫu đo đạc**

Hệ thống sẽ tự động nhân bản trọn bộ hồ sơ gồm cả **Biên bản giám định hiện trường (BM3030.02)** và **Phiếu ghi chép hiện trường (BM3030.01)** với bảng chỉ tiêu được co giãn linh hoạt đúng số dòng yêu cầu, đồng thời làm sạch toàn bộ dữ liệu đo cũ để biểu mẫu hoàn toàn sạch sẽ, sẵn sàng cho kỹ thuật viên sử dụng.

---

## 📁 PHÂN TÍCH CẤU TRÚC THƯ MỤC GỐC

Thư mục gốc chứa các biểu mẫu của 5 lĩnh vực nghiệp vụ giám định kỹ thuật:

| Lĩnh vực | Thư mục gốc | Biểu mẫu sử dụng | Đặc điểm & Các chỉ tiêu tiêu biểu |
| :--- | :--- | :--- | :--- |
| **Phòng sạch, Vi sinh** | `Phong sach, Vi sinh/` | BM3030.01 + BM3030.02 | Nồng độ hạt bụi (ISO 14644), vi sinh không khí, nhiệt độ - độ ẩm, đo chênh áp phòng, số lần trao đổi gió, độ ồn. |
| **Khí nén, Khí Nitơ** | `Khi nen/` | BM3030.01 + BM3030.02 | Phân tích tổng dầu (oil aerosol & vapor), nhiệt độ điểm sương (dew point), nồng độ hạt bụi khí có áp, khí Nitơ. |
| **CO2 - Môi trường LĐ** | `CO2 - MTLD/` | BM3030.01 + BM3030.02 | Quan trắc môi trường làm việc văn phòng/nhà máy: Khí CO2 và tổng chất hữu cơ bay hơi (TVOC). |
| **Kho lạnh** | `Kho lanh/` | BM3030.01 + BM3030.02 | Đánh giá độ đồng đều nhiệt độ và độ ẩm kho bảo quản thuốc/vắc-xin (nhiệt kế tự ghi Haxo-8 mapping). |
| **Tủ an toàn sinh học** | `BSC/` | BM3030.01 + BM3030.02 | Hiệu năng an toàn tủ BSC theo EN 12469:2000 (Vận tốc inflow, downflow, HEPA leak test, độ sáng, độ ồn, UVC, khói). |

### Hai biểu mẫu cốt lõi trong mỗi bộ hồ sơ:
1. **BM3030.02 (Biên bản giám định hiện trường / BBLM)**: Biên bản làm việc pháp lý giữa QUATEST 3 và khách hàng, thể hiện danh mục các chỉ tiêu / mẫu kiểm tra (Bảng Table 2 gồm 8 cột chuẩn hoặc 5 cột đối với BSC).
2. **BM3030.01 (Phiếu ghi chép giám định tại hiện trường)**: Bản ghi chép kỹ thuật để kỹ thuật viên ghi số liệu đo đạc trực tiếp, lưu mã file đo (file zeroing, file đo particle, giá trị đo thô).

---

## 🚀 HƯỚNG DẪN SỬ DỤNG

Bạn có 4 cách cực kỳ linh hoạt để sử dụng hệ thống:

## 🚀 CÁCH SỬ DỤNG HỆ THỐNG MỚI (MODULAR ARCHITECTURE)

Hệ thống đã được tái cấu trúc hoàn chỉnh thành bộ module lõi tại `core/`, tuân thủ nghiêm ngặt **11 Quy tắc sinh biểu mẫu QUATEST 3 (GEMINI.md)**. Bạn có các cách dùng sau mà **không cần phải viết bất kỳ dòng code tạm (scratch script) nào**:

### Cách 1: Một lệnh duy nhất qua `generate_now.py` (Khuyên dùng nhất)
Chỉ cần gọi lệnh dòng lệnh kèm theo preset profile:
```powershell
# 1. Hồ sơ Tủ an toàn sinh học (BSC) -> Tự động ghép 3 trang vector PDF
python generate_now.py --profile bsc --doc Q326D132918 --customer "PHÒNG THÍ NGHIỆM CÔNG NGHỆ SINH HỌC" --address "Khu phố 6, phường Linh Trung, TP. Thủ Đức, TP.HCM" --pdf

# 2. Hồ sơ Khí nén 3 vị trí đo
python generate_now.py --profile khi_nen --num_positions 3 --doc Q326D122318 --customer "CÔNG TY TNHH THỰC PHẨM NAM DƯƠNG" --address "KCN Hiệp Phước, Nhà Bè, TP.HCM" --pdf

# 3. Hồ sơ Phòng sạch 5 chỉ tiêu (Bụi, Nhiệt ẩm, Chênh áp, Ồn, Sáng)
python generate_now.py --profile phong_sach --doc Q326D124674 --customer "CÔNG TY TNHH THỰC PHẨM CÁT TÂN" --address "KCN Tân Phú Trung, Củ Chi, TP.HCM" --pdf

# 4. Hồ sơ Hybrid Khí nén + Phòng sạch
python generate_now.py --profile khi_nen_phong_sach --num_positions 2 --num_samples_bm01 8 --doc Q326D133266 --customer "CÔNG TY TNHH NEXTERN VIỆT NAM" --address "Nhà xưởng C1B-D5, Lô C-1B-CN, Đường DE5, Khu Công Nghiệp Mỹ Phước 3, Phường Thới Hòa, Thành phố Hồ Chí Minh" --pdf

# 4b. Mã hồ sơ để trắng (chưa có số phiếu): thêm --blank_doc -> mã Q326D…………… và ngày đăng ký …./…./2026 để điền tay
python generate_now.py --profile khi_nen_phong_sach --blank_doc --num_positions 2 --num_samples_bm01 8 --customer "CÔNG TY TNHH NEXTERN VIỆT NAM" --address "..." --pdf

# 5. Hồ sơ Vi sinh không khí & Swab bề mặt
python generate_now.py --profile phong_sach_vskk_swab --doc Q326D132223 --customer "CÔNG TY CỔ PHẦN HULO CARE" --address "KDC Him Lam, Phường 11, Quận 6, TP.HCM" --pdf
```

### Cách 1b: Hồ sơ Môi trường lao động theo kế hoạch từng phòng (`generate_mtld.py`)
Điền kế hoạch vào `templates/MTLD_KeHoach_mau.xlsx` (cột Phòng | Chỉ tiêu | SL | Điểm đo), rồi:
```powershell
python generate_mtld.py --plan "Input/KeHoach.xlsx" --customer "CÔNG TY TNHH ABC" --address "Lô A1, KCN Tân Bình, TP.HCM" --doc Q326D145000 --reg_date 18/11/2026 --pdf
# Bỏ --doc để để trắng mã hồ sơ; thêm --multi_day nếu đo nhiều ngày; --scan "Input/... - Scan PYC.pdf" để chép scan lên Z
```
Ra Ghi chép (mỗi phòng 1 trang: chọn máy ☐, ô xám = không đo, ồn ghi tên file đo, cột Bơm Buck/SKC/Kitagawa, Ecgonomy) và Biên bản (mỗi chỉ tiêu 1 dòng). Chi tiết: Quy tắc 14.

### Cách 2: Gọi trực tiếp qua Python API (`process_order`)
Bất kỳ script Python nào cũng có thể tạo trọn bộ hồ sơ chỉ bằng 3 dòng:
```python
from core.form_cloner import process_order

res = process_order({
    "doc_number": "Q326D124674",
    "customer_name": "CÔNG TY TNHH THỰC PHẨM CÁT TÂN",
    "customer_address": "Lô A2-5, Đường D4, KCN Tân Phú Trung, Củ Chi, Tp.HCM",
    "profile": "phong_sach",
    "indicators": [
        "Đo nồng độ bụi phòng sạch",
        "Đo nhiệt độ, độ ẩm phòng sạch",
        "Đo chênh áp phòng sạch",
        "Đo độ ồn phòng sạch",
        "Đo độ sáng phòng sạch"
    ]
}, export_pdf=True)

print("Hồ sơ đã tạo tại:", res["output_dir"])
```

### Cách 3: Giao diện đồ họa Desktop (Desktop GUI)
- Chạy `python app_gui.py` hoặc nhấp đúp vào file `run_app.bat`.

### Cách 4: Nhân bản hàng loạt từ file Excel (Batch Generator)
- Chạy `python app_cli.py --batch danh_sach_khach_hang_mau.xlsx --pdf`

---

## 🏗️ CẤU TRÚC KIẾN TRÚC MODULE LÕI (`core/`)

```
core/
├── config.py             # SOP templates, quy trình, chỉ tiêu & thiết bị tương ứng
├── normalizer.py         # Quy tắc 1, 2, 4, 5, 8, 9, 10, 11 (Chuẩn hóa 2 cấp địa chỉ, ngày chấm lửng, thiết bị, chủ thể, DB title)
├── docx_processor.py     # Nền tảng xử lý XML, font, styles, căn lề ô, cantSplit, spacer 6pt
├── bm02_builder.py       # Module chuyên trách BM3030.02, thuật toán căn chỉnh đảm bảo CHÍNH XÁC 2 TRANG
├── bm01_builder.py       # Module chuyên trách BM3030.01: BSC, Khí nén 1-N vị trí, Phòng sạch & Hybrid
├── order_profiles.py     # Thư viện Preset profiles của các hồ sơ mẫu & heuristic nhận diện tự động
├── pdf_exporter.py       # Xuất PDF qua Microsoft Word COM và ghép vector PDF độ nét cao
├── form_cloner.py        # Dispatcher trung tâm tiếp nhận yêu cầu và điều phối toàn hệ thống
└── __init__.py           # Public API tinh gọn
```
- `--doc-number`: Mã số hồ sơ (tùy chọn)
- `--date`: Ngày thực hiện DD/MM/YYYY (tùy chọn)

---

### Cách 4: Nhân bản hàng loạt từ file Excel (Batch Generator)
Hệ thống đã chuẩn bị sẵn file Excel mẫu: **`danh_sach_khach_hang_mau.xlsx`**.
1. Mở file `danh_sach_khach_hang_mau.xlsx` và nhập danh sách khách hàng của bạn.
2. Chạy lệnh:
```powershell
python app_cli.py --batch danh_sach_khach_hang_mau.xlsx
```
Toàn bộ hồ sơ của các khách hàng trong file Excel sẽ được tạo đồng loạt!

---

## 📂 QUY TẮC ĐẶT TÊN VÀ LƯU TRỮ TÀI LIỆU OUTPUT

Mỗi lần nhân bản, tài liệu được lưu trữ tự động vào thư mục `output/` theo cấu trúc chuẩn:
**`"số hồ sơ - tên khách hàng - nội dung vắn tắc"`**

```
output/
└── [Mã_Số] - [Tên_Khách_Hàng] - [Nội_Dung_Vắn_Tắt]/
    ├── [Mã_Số] - BM3030.02 - BBLM - [Tên_KH].docx
    └── [Mã_Số] - BM3030.01 - Ghi chep - [Tên_KH].docx
```

**Ví dụ thực tế:**
- `Q326D124674 - Cát Tân - Phòng sạch`
- `Q326D120711 - Greencans Industries - Nước uống`
- `Q326D117382 - Thụy Khuê - Nuoc sinh hoat`
- `Q326D124796 - K pack - vskk - swab`

*(Ghi chú: VSKK = Vi sinh không khí, VSBM = Vi sinh bề mặt / swab)*

---

## 🛠️ CẤU TRÚC MÃ NGUỒN DỰ ÁN

```
c:\Users\phat.nt\Desktop\Source\Form Bieu Mau Goc\
├── app_gui.py                     # Giao diện đồ họa Desktop hiện đại (CustomTkinter)
├── app_cli.py                     # Giao diện dòng lệnh (hỗ trợ interactive & arguments)
├── batch_generator.py             # Xử lý sinh hàng loạt từ file Excel
├── create_excel_template.py       # Script tạo file Excel mẫu
├── run_app.bat                    # Phím tắt 1-click khởi động ứng dụng trên Windows
├── danh_sach_khach_hang_mau.xlsx  # File Excel mẫu chuẩn
├── core/
│   ├── __init__.py
│   ├── config.py                  # Cấu hình danh mục chỉ tiêu, thiết bị, quy trình chuẩn
│   ├── docx_processor.py          # Module xử lý file Word, giữ 100% format, co giãn Table 2
│   └── form_cloner.py             # Logic điều phối nhân bản cả cặp BM3030.01 và BM3030.02
├── output/                        # Thư mục chứa các hồ sơ đã nhân bản
└── README.md                      # Hướng dẫn chi tiết sử dụng
```

---

## ✨ CÁC ĐIỂM NỔI BẬT VỀ MẶT KỸ THUẬT

1. **Bảo toàn 100% định dạng Word gốc**:
   - Sử dụng phương pháp thay thế run-level để giữ nguyên font chữ **Times New Roman**, cỡ chữ (10pt, 11pt, 12pt), thuộc tính **in đậm**, in nghiêng và căn lề.
   - Giữ nguyên các ô gộp (merged cells) của QUATEST 3 trong Table 0 và Table 1.
2. **Co giãn linh hoạt Bảng chỉ tiêu (Table 2)**:
   - Tự động sao chép cấu trúc XML của dòng dữ liệu nguyên mẫu (prototype row).
   - Tự động đánh số thứ tự (STT: 1, 2, ..., N), điền dấu tích `x` cho phương pháp Đo hoặc Lấy mẫu, điền dấu chấm `……………………………………………………` cho vị trí đo.
3. **Làm sạch số liệu cũ**:
   - Tự động xóa sạch các giá trị đo cũ của khách hàng trước đó trong các bảng đo BM3030.01, giúp kỹ thuật viên có mẫu trắng hoàn hảo để đi hiện trường.

---

## 🤖 CÀI ĐẶT TRÊN CLAUDE MỚI (PLUGIN `bieumau`)

Trên máy mới chỉ cần có Claude Code (app desktop hoặc CLI). Trong ô chat của Claude gõ lần lượt:

```
/plugin marketplace add phatnguyen1999p/Form-Bieu-Mau-Goc
/plugin install bieumau@bieumau-marketplace
```

Khởi động lại phiên rồi nói **"cài đặt biểu mẫu cho máy này"** — skill `cai-dat` sẽ tải mã nguồn về `Desktop\Source\Form Bieu Mau Goc`, cài Python + thư viện (`requirements.txt`), kiểm tra Word, ổ Z và chạy thử.

Các skill trong plugin:

| Skill | Dùng khi |
| :--- | :--- |
| `cai-dat` | Máy mới, kiểm tra môi trường |
| `tao-ho-so` | Có scan phiếu yêu cầu trong `Input/`, cần tạo bộ BM3030.01 + BM3030.02 |
| `erp-nhan-ho-so` | Nhận hồ sơ, phân công, tạo YCTN trên ERP |

Cập nhật plugin khi repo có bản mới: `/plugin marketplace update bieumau-marketplace`.

---

## 📝 NHẬT KÝ THAY ĐỔI
- **27/10/2026 (Claude Cowork):** Hỗ trợ mã hồ sơ để trắng (`--blank_doc`), sửa lỗi chuẩn hóa `Đường`/`Đ.`, sửa thiết bị & quy trình cho profile `khi_nen_phong_sach`, sửa bố cục trang BM01 Hybrid. Chi tiết xem mục *NHẬT KÝ THAY ĐỔI CODE* cuối `GEMINI.md`.
