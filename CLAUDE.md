# Form Bieu Mau Goc — hướng dẫn cho Claude

Dự án nhân bản biểu mẫu hiện trường QUATEST 3 (BM3030.01 Ghi chép + BM3030.02 Biên bản).

- Quy tắc soạn biểu mẫu (bắt buộc đọc trước khi sinh/sửa hồ sơ): @GEMINI.md và `QUY_TAC_SOAN_THAO_BIEU_MAU.md`.
- Cách chạy: `README.md` (`generate_now.py`, `generate_mtld.py`, `core.form_cloner.process_order`).
- Quy trình làm việc đóng gói trong plugin `bieumau` (thư mục `plugin/bieumau`): skill `cai-dat`, `tao-ho-so`, `erp-nhan-ho-so`.
- `Input/`, `output/`, `scratch/` chứa dữ liệu khách hàng, không đưa lên GitHub.
- Không ghi đè file trên ổ Z; không sửa code `core/` khi người dùng chưa đồng ý.
