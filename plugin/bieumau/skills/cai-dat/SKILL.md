---
name: cai-dat
description: Cài đặt dự án "Form Bieu Mau Goc" (nhân bản biểu mẫu hiện trường QUATEST 3) trên máy mới — tải mã nguồn từ GitHub, cài Python và thư viện, kiểm tra Word, ổ Z, chạy thử. Dùng khi người dùng nói "cài đặt", "setup máy mới", "cài biểu mẫu cho máy này", "kiểm tra môi trường", hoặc khi chưa tìm thấy thư mục dự án.
---

# Cài đặt "Form Bieu Mau Goc" trên máy mới

Mục tiêu: máy Windows mới tinh (chỉ có Claude) chạy được `generate_now.py` và xuất được PDF. Trong lúc cài **không** sửa code, không ghi gì lên ổ Z, không chạy sinh hồ sơ thật (chỉ chạy thử với `--no_z_copy` vào thư mục tạm). Báo người dùng từng bước: có / thiếu / đã cài.

## 1. Mã nguồn

- Vị trí chuẩn: `%USERPROFILE%\Desktop\Source\Form Bieu Mau Goc`. Nếu đã có thư mục và có `generate_now.py` → dùng luôn, chạy `git pull` nếu là git repo.
- Chưa có → cần Git (`git --version`; thiếu thì hỏi người dùng rồi cài: `winget install --id Git.Git -e`). Sau đó:
  ```
  git clone https://github.com/phatnguyen1999p/Form-Bieu-Mau-Goc.git "%USERPROFILE%\Desktop\Source\Form Bieu Mau Goc"
  ```
- Các thư mục `Input\`, `output\`, `scratch\` không có trên GitHub (dữ liệu khách hàng) → tạo `Input\` và `output\` rỗng.
- Từ đây mọi lệnh chạy ở gốc dự án. Đọc `CLAUDE.md`, `README.md`; quy tắc soạn biểu mẫu ở `GEMINI.md` và `QUY_TAC_SOAN_THAO_BIEU_MAU.md`.

## 2. Python

- Cần Python ≥ 3.10 trên PATH (`python --version`). Thiếu → hỏi người dùng rồi cài: `winget install --id Python.Python.3.12 -e`, mở lại terminal. Nếu `python` mở Microsoft Store thì thử `py`.
- Cài thư viện: `python -m pip install -r requirements.txt`
- Kiểm tra: `python -c "import docx, fitz, win32com.client, openpyxl, pandas, xlsxwriter, PIL, customtkinter; print('OK')"`
- pywin32 báo lỗi DLL → `python -m pywin32_postinstall -install` (hoặc `python Scripts\pywin32_postinstall.py -install` trong thư mục Python).

## 3. Microsoft Word (bắt buộc để xuất PDF)

- `python -c "import win32com.client as w; a=w.DispatchEx('Word.Application'); print(a.Version); a.Quit()"`
- Không có Word → vẫn sinh được .docx, chỉ không dùng được `--pdf`. Báo người dùng.

## 4. Ổ mạng Z (chỉ kiểm tra, không ghi)

- `Test-Path "Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026"`
- Không có → nhắc người dùng map ổ Z, hoặc khi sinh hồ sơ thêm `--no_z_copy`. Có thể đặt đường dẫn khác bằng biến môi trường `QUATEST3_Z_DRIVE` (hỏi người dùng trước khi `setx`).

## 5. Chạy thử (không đụng Z)

```
python generate_now.py --profile khi_nen --num_positions 1 --blank_doc --customer "CÔNG TY TNHH CHẠY THỬ" --address "Lô A1, KCN Tân Bình, TP.HCM" --no_z_copy --pdf
```
- Kiểm tra trong `output\` có thư mục mới với BM3030.01 + BM3030.02 (.docx, và .pdf nếu có Word). Mở PDF xem nhanh rồi **xóa thư mục chạy thử** (hỏi người dùng trước khi xóa).

## 6. Tổng kết

Liệt kê: đường dẫn dự án, phiên bản Python, thư viện OK/thiếu, Word OK/thiếu, ổ Z OK/thiếu, kết quả chạy thử. Gợi ý bước tiếp: dùng skill `tao-ho-so` để làm hồ sơ đầu tiên (bỏ file scan phiếu yêu cầu vào `Input\`).
