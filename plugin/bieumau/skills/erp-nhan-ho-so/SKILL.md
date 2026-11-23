---
name: erp-nhan-ho-so
description: Nhận hồ sơ và phân công trên ERP QUATEST 3 (Odoo) — Xác nhận đánh giá, Phân công N2 - Môi trường, tạo Yêu cầu thử nghiệm (YCTN) khi có mẫu gửi phòng thí nghiệm. Dùng khi người dùng nói "nhận hồ sơ trên ERP", "phân công hồ sơ", "tạo YCTN".
---

# ERP: nhận hồ sơ, phân công, tạo YCTN

ERP: https://erp.quatest3.com.vn → Khối Nghiệp vụ > Hồ sơ yêu cầu. Dùng trình duyệt tích hợp; người dùng tự đăng nhập (không nhập mật khẩu thay). Đặt khung nhìn rộng ~1400px để Odoo hiện dạng danh sách (hẹp quá sẽ ra kanban, thiếu cột).

**Tìm hồ sơ:** trường "Số." là mã N2NV…, mã Q326D… nằm ở cột "Hồ sơ tổng". Tìm theo Số với Q326D… không ra → lọc theo **Khách hàng** (tên công ty).

## 1. Nhận và phân công (mọi hồ sơ: lấy mẫu lẫn đo)

1. Mở hồ sơ → **Xác nhận đánh giá** (trạng thái → Chờ phân công).
2. **Phân công công việc** → Assign Groups: **N2 - Môi trường**; Phân công cho: người được người dùng chỉ định (mặc định của nhóm: **N Tiến Phát - N2** và **N Xuân Nam - N2**; máy mới hỏi lại người dùng) → Xác nhận (→ Đang thực hiện).

## 2. Tạo YCTN (chỉ khi có mẫu gửi phòng thí nghiệm)

- Tab **Yêu cầu nội bộ** → **Tạo YCTN** (tạo phiếu Nháp N2NV….1) → **Chi tiết**.
- **Ngày hẹn trả KQ** điền trước: mặc định = ngày lấy mẫu + 7 ngày; mẫu tỉnh xa về HCM trễ thì tính từ ngày mẫu về HCM; người dùng có thể chỉ định khác. Ô ngày mở lịch riêng (gõ phím không nhận) → chọn ngày trên lịch → bấm **Lưu** (biểu tượng đám mây).
- Thêm chỉ tiêu: **Dữ liệu thử nghiệm (tìm theo nền mẫu)** → chọn **Mẫu TT3** (nền mẫu) trước, rồi "Thêm một dòng". Tên mẫu theo khách hàng = Tên trên PKQ; số mẫu theo hợp đồng. Kiểu tác vụ **Mới** = thêm nhóm mẫu mới, nhóm cũ giữ nguyên. Trong hộp chọn chỉ tiêu dùng ô **Tìm kiếm** trên cùng, không bấm vào ô lọc cột.
- Một chỉ tiêu có nhiều phương pháp (vd TSS, COD) → **hỏi người dùng chọn**, không tự chọn.
- **Tách YCTN theo nền mẫu**: không khí MTLĐ một phiếu, nước thải một phiếu.
- Nước thải: tick **Combo chỉ tiêu** → Mẫu TT3 "MT | Nước thải" → Gói chỉ số (vd QCVN 40:2011/BTNMT Full) → xóa các chỉ tiêu đo hiện trường (nhiệt độ, pH, clo dư) → Xác nhận. BBLM phải có ô ghi kết quả đo hiện trường các chỉ tiêu này.
- MTLĐ: hỏi người dùng chỉ tiêu nào đo tại hiện trường, chỉ tiêu nào lấy mẫu gửi PTN, trước khi tạo YCTN. Chỉ tiêu đo hiện trường không đưa vào YCTN.
- Lưu ý lỗi ERP: form "Dữ liệu thử nghiệm" nhớ chỉ tiêu + tên mẫu lần trước → xóa dòng cũ, sửa tên trước khi Xác nhận. Nút xóa / kiểu tác vụ "Xóa" chỉ xóa chỉ tiêu, để lại nhóm mẫu rỗng. YCTN nháp bị lẫn mẫu sai → **Hủy** (hỏi người dùng) rồi tạo YCTN mới (.2, .3…).
- **Không bấm "Xác nhận & gửi cho PTN"** khi người dùng chưa duyệt. Mọi nút xác nhận / gửi / hủy đều hỏi người dùng trước.
