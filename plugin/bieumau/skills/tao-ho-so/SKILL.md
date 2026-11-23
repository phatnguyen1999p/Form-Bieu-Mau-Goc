---
name: tao-ho-so
description: Tạo bộ biểu mẫu hiện trường QUATEST 3 (BM3030.02 Biên bản + BM3030.01 Ghi chép, kèm PDF, chép lên ổ Z) cho một khách hàng từ file scan phiếu yêu cầu hoặc thông tin người dùng đưa — BSC, khí nén, phòng sạch, vi sinh không khí / swab, kho lạnh, CO2, môi trường lao động (MTLĐ). Dùng khi người dùng nói "làm biểu mẫu", "tạo hồ sơ Q326D…", "có scan mới trong Input", "làm biên bản / ghi chép cho công ty …".
---

# Tạo hồ sơ biểu mẫu hiện trường

Dự án ở `%USERPROFILE%\Desktop\Source\Form Bieu Mau Goc` (không thấy → chạy skill `cai-dat` trước). Mọi lệnh chạy ở gốc dự án.

**Trước khi làm, đọc quy tắc:** `GEMINI.md` (14 quy tắc + nhật ký thay đổi) và `QUY_TAC_SOAN_THAO_BIEU_MAU.md`. Hai file này là chuẩn; skill này chỉ tóm tắt quy trình và các quy định người dùng chốt thêm.

## 1. Lấy thông tin

- Scan mới thường nằm trong `Input\` với tên tự động kiểu `20261114144734901.pdf`. Đọc PDF (PyMuPDF; scan ảnh thì xem ảnh trang) để lấy: **mã hồ sơ Q326D…**, ngày đăng ký, tên + địa chỉ khách hàng, địa điểm thực hiện, nội dung yêu cầu / chỉ tiêu / số mẫu / số vị trí, số báo giá QONV… hoặc số PO.
- Thiếu thông tin hoặc chọn lựa chưa rõ (chỉ tiêu nào đo, bao nhiêu điểm) → hỏi người dùng, không tự đoán.
- Chưa có mã hồ sơ → dùng `--blank_doc` (Quy tắc 12).
- **Đổi tên scan ngay** khi có mã: `<Mã hồ sơ> - <Tên KH rút gọn> - Scan PYC.pdf` (vd `Q326D125256 - Dược Sài Gòn - Scan PYC.pdf`). Một file scan gộp nhiều hồ sơ → tách thành từng file. Không đổi tên các file cũ đã xử lý.
- Tên khách hàng đổi tên pháp nhân: Công ty CP Sanofi Việt Nam → **Công ty Cổ phần Opella Healthcare Việt Nam** (từ 09/11/2026, cùng MST 0312233458), hồ sơ mới ghi tên mới dù giấy tờ in tên cũ; tên rút gọn `Opella Healthcare Việt Nam`.

## 2. Kiểm tra ổ Z trước khi sinh

Tìm mã Q326D trong `Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026`:
- **Đã có thư mục cùng mã** (dù tên khác mẫu) → đó là cùng một hồ sơ: sinh với `--no_z_copy`, đặt tên thư mục `output\` trùng tên thư mục Z đó, rồi chép file vào `<thư mục Z>\Thuc hien\` (kiểm tra trùng tên, **không ghi đè**). Thư mục cũ cũng là nguồn tốt để lấy địa chỉ / địa điểm.
- Chưa có → để mặc định, script tự chép sang Z (Quy tắc 13), scan vào `Scan\`, file sinh ra vào `Thuc hien\`.
- Không có ổ Z → `--no_z_copy`, báo người dùng.

## 3. Chọn profile và chạy

`generate_now.py` (xem `python generate_now.py -h`). Profile có sẵn trong `core/order_profiles.py`: `bsc`, `khi_nen`, `phong_sach`, `phong_sach_vskk_swab`, `khi_nen_phong_sach`, `kho_lanh`, `vi_sinh_be_mat`.

```
python generate_now.py --profile khi_nen --num_positions 3 --doc_number Q326D122318 --customer "CÔNG TY TNHH ..." --address "..." --scan "Input/Q326D122318 - ... - Scan PYC.pdf" --pdf
```
Tham số hay dùng: `--num_positions`, `--num_samples_bm01`, `--num_vskk_bm01`, `--bm01_rows`, `--cabinets` / `--blank_cabinets` (BSC), `--indicators "a;b;c"`, `--quote QONV…` hoặc `--po …`, `--reg_date`, `--exec_date`, `--short_customer`, `--short_content`, `--json`.

**Môi trường lao động theo kế hoạch từng phòng** → `generate_mtld.py`: điền kế hoạch theo mẫu `templates/MTLD_KeHoach_mau.xlsx` (Phòng | Chỉ tiêu | SL | Điểm đo), lưu vào `Input\`, rồi:
```
python generate_mtld.py --plan "Input/KeHoach.xlsx" --customer "..." --address "..." --doc Q326D145000 --reg_date 18/11/2026 --scan "Input/... - Scan PYC.pdf" --pdf
```
Thêm `--multi_day` nếu đo nhiều ngày. Chi tiết: Quy tắc 14.

Không profile nào khớp → dùng Python API `core.form_cloner.process_order({...})` hoặc viết script dựng tay trong `scratch\`, nhưng vẫn theo đúng quy tắc trong `GEMINI.md`. Không sửa code `core/` khi người dùng chưa đồng ý.

## 4. Quy định đã chốt (luôn áp dụng)

- **Địa điểm thực hiện**: chừa đúng **90 dấu chấm** liền một dòng trước ` - [địa chỉ]` (hằng `LOCATION_DOTS` trong `core/normalizer.py`); script dựng tay cũng dùng 90.
- **Thiết bị MTLĐ**: bơm lấy mẫu khí ghi "Bơm lấy mẫu khí Kitagawa ASP 1200, Japan" (không dùng Buck Allin-1); máy đo bụi ghi "Máy đo bụi Temptop PMD 371, USA" (không dùng Metone Aerocet 831 / bơm SKC khi bụi đo trực tiếp).
- Tên thư mục output: `<Mã> - <Tên KH rút gọn> - <Nội dung vắn tắt>` (Quy tắc 9), trừ khi Z đã có thư mục cùng mã (mục 2).

## 5. Kiểm tra kết quả

- Mở PDF, xem từng trang: BM3030.02 đúng 2 trang, bảng chỉ tiêu đủ dòng, không trang trắng thừa, mã hồ sơ / ngày / tên KH / địa chỉ đúng chuẩn hóa, dòng địa điểm có chấm.
- Báo người dùng: thư mục output, đường dẫn trên Z, các điểm còn để trống cần điền tay.
- Sau đó nếu người dùng muốn nhận hồ sơ trên ERP → skill `erp-nhan-ho-so`.
