# HỆ THỐNG QUY TẮC SOẠN THẢO BIỂU MẪU HIỆN TRƯỜNG QUATEST 3
## (Áp dụng cho BM3030.01 - Ghi chép & BM3030.02 - Biên bản lấy mẫu / Giám định)

Tài liệu này tổng hợp toàn bộ các quy tắc chuẩn mực bắt buộc phải tuân thủ khi sinh hoặc chỉnh sửa các biểu mẫu hiện trường của QUATEST 3 (Phòng sạch, Vi sinh, Khí nén, CO2 môi trường lao động, Tủ BSC, Kho lạnh...).

---

### PHẦN 1: QUY TẮC TRÍCH XUẤT MÃ HỒ SƠ VÀ NGÀY THÁNG
1. **Mã số hồ sơ**:
   - Vị trí trích xuất: Ở **Trang 1 của Phiếu yêu cầu**, phần *"Phần dành cho Trung tâm Kỹ thuật 3"* (phía trên chữ ký của người nhận yêu cầu / NVGĐ Dương Ngọc Bảo Hân).
   - Mã số luôn có đầy đủ tiền tố `Q326D...` (ví dụ `Q326D122015`), tuyệt đối không được cắt bớt chỉ lấy đuôi số `122015`.
2. **Ngày đăng ký (nhận hồ sơ)**:
   - Lấy đúng theo ngày in phía trên chữ ký của NVGĐ ở trang 1 Phiếu yêu cầu (ví dụ: `Ngày 29 tháng 09 năm 2026` -> `29/09/2026`).
   - Tuyệt đối **KHÔNG** lấy ngày phát hành của Báo giá (trang 5).
   - Điền ngày đăng ký vào dòng `Phiếu yêu cầu & Nhận TĐ/GĐ: Q326D... Ngày: 29/09/2026`.
3. **Ngày hiện trường (Header Table 0 & Thời gian thực hiện)**:
   - **Header Table 0 của BM3030.02**: Ô `Ngày:` bắt buộc để định dạng chấm lửng `…./…./YYYY` (ví dụ `…./…./2026`), **KHÔNG ĐƯỢC ĐIỀN NGÀY ĐĂNG KÝ** vì ngày thực hiện hiện trường chưa diễn ra tại thời điểm nhận hồ sơ.
   - **Thời gian thực hiện (cả BM3030.01 và BM3030.02)**:
     - Cú pháp chuẩn: `Bắt đầu từ.................... kết thúc lúc..................... ngày …./…./YYYY` (ví dụ `ngày …./…./2026`).
     - Viết hoa chữ `Bắt đầu`.
     - Xóa bỏ tiền tố ngày tháng cũ (như `Ngày ... tháng ... năm ..., `).
     - Ngày để dạng chấm lửng `ngày …./…./2026`, không điền ngày đăng ký vào đây trừ khi có lịch hẹn hiện trường cụ thể.

---

### PHẦN 2: QUY TẮC ĐỊNH DẠNG FONT CHỮ VÀ KIỂU CHỮ (FONT & STYLING)
1. **Font chữ & Cỡ chữ**:
   - Sử dụng thống nhất font chữ **Times New Roman**, cỡ chữ **13pt** (`Pt(13)`) cho toàn bộ các nội dung văn bản, thông tin khách hàng, số liệu, tiêu đề và bảng biểu.
2. **Quy cách in đậm / in thường**:
   - **Không in đậm** (`bold=False`) các thông tin điền mới (Tên đơn vị yêu cầu, Địa chỉ, Địa điểm thực hiện, Thời gian thực hiện, Thiết bị sử dụng, Quy trình thực hiện, Nội dung yêu cầu, Tên chỉ tiêu trong bảng, Số liệu đo...). Giữ đúng phong cách biểu mẫu hành chính gốc của QUATEST 3.
   - Chỉ in đậm (`bold=True`) tiêu đề chính của các bảng đo đạc hiện trường đặt phía trên bảng.
3. **Ngày hiệu lực biểu mẫu (`HL:...`)**:
   - Cụm từ ngày hiệu lực ở góc trên bên phải bảng Header Table 0 (ví dụ: `HL: 29/09/2026`, `HL:01/8/2025`) là **ngày hiệu lực cố định của form biểu mẫu**, tuyệt đối **KHÔNG** được thay đổi theo ngày làm việc tại hiện trường.
4. **Nội dung yêu cầu (Mục 2 trong Table 1 BM01 và Mục 4 trong Table 1 BM02)**:
   - Bắt buộc nằm trên cùng một hàng tại cột 3 (Column 2) ngay sau dấu hai chấm `:` (`4. Nội dung yêu cầu : [Nội dung yêu cầu]`).
   - Tuyệt đối không để trống ô cột 3 sau dấu `:` rồi đẩy nội dung xuống một hàng gộp riêng bên dưới. Tự động xóa bỏ hàng rác phụ của template cũ bên dưới dòng `Nội dung yêu cầu`.

---

### PHẦN 3: QUY TẮC TRÌNH BÀY THIẾT BỊ SỬ DỤNG (EQUIPMENT)
3.1. **Khí nén (N2.TC.KN.01)**:
     - Có chỉ tiêu đo hạt bụi trong khí nén:
       - Nếu yêu cầu đo hạt bụi kích thước đến 0,1 µm (`0,1um`, `0.1um`, `0,1-5,0um`...) -> `Máy đo nồng độ hạt bụi Lasair III 110 và phụ kiện HPD II, USA;`
       - Nếu đo nồng độ hạt thông thường (≥ 0,3 µm) -> `Máy đo nồng độ hạt AeroTrak 9350 và phụ kiện HPD III, USA;`
     - Có chỉ tiêu đo điểm sương -> `Thiết bị đo nhiệt độ điểm sương;`
     - Có chỉ tiêu lấy mẫu vi sinh khí nén -> `Thiết bị lấy mẫu vi sinh khí nén SAS Pinocchio Super II.`
     - Có chỉ tiêu lấy mẫu tổng dầu -> 2 dòng (cập nhật 11/11/2026): `- Bộ dụng cụ lấy mẫu tổng dầu trong khí nén;` và `- Giấy lọc thủy tinh Pall life Sciences, ống than hoạt tính ORBO®.` (dòng cuối dùng `;` nếu còn thiết bị khác phía sau)
     - Quy trình áp dụng: `QUATEST 3 3030:2025; N2.TC.KN.01` (bổ sung `N2.TC.09` nếu có vi sinh).

   3.2. **Phòng sạch (N2.TC.PS.01 & N2.TC.09)**:
     - Đo nồng độ hạt bụi thông thường (≥ 0,3 µm, ISO 7-9, Grade C, D) -> `Thiết bị đo nồng độ hạt TSI AeroTrak 9350, USA;`
     - Đo hạt bụi siêu mịn (kích thước đến 0,1 µm, ISO 1-6, Grade A, B) -> `Máy đo nồng độ hạt bụi PMS Lasair III 110, USA;`
     - Đo nhiệt độ, độ ẩm -> `Thiết bị đo nhiệt độ, độ ẩm trong không khí Testo 625, Germany;`
     - Đo chênh lệch áp suất -> `Thiết bị đo chênh lệch áp suất TSI 8375, USA;`
     - Đo độ ồn -> `Thiết bị đo độ ồn Rion NL 53, Japan;`
     - Đo độ sáng -> `Thiết bị đo độ sáng Testo 440, Germany;`
     - Đo lưu lượng khí -> `Thiết bị đo lưu lượng TSI 8375, USA;`
     - Lấy mẫu vi sinh trong không khí -> `Thiết bị lấy mẫu vi sinh trong không khí Buck – Bio Culture B30120, USA.`
     - Quét rò rỉ màng lọc HEPA phòng sạch -> `Máy tạo sol khí ATI aerosol generator, model 6D và quang kế aerosol photometer, model 2i, USA;` (BM3030.01 có thêm bảng `Quét rò rỉ màng lọc HEPA` cột `Độ xuyên thấu (%)`; khi bảng đo bụi ≥ 8 dòng thì bảng HEPA sang trọn Trang 3 với 15 dòng).
     - Quy trình áp dụng: `QUATEST 3 3030:2025; N2.TC.PS.01` (ghép `N2.TC.09` nếu có vi sinh, `N2.TC.01`, `N2.TC.02` nếu có vi khí hậu/ồn).

   3.3. **Tủ an toàn sinh học (BSC), Tủ sạch (LAF), Tủ hút (N2.TC.BSC.01/02, N2.TC.LAF.01, N2.TC.T.01)**:
     - Tốc độ gió vào (inflow) -> `Thiết bị đo lưu lượng TSI 8375/8380, USA;` hoặc `Máy đo tốc độ gió Kanomax Climomaster 6501, Japan;`
     - Tốc độ gió cấp (downflow) -> `Máy đo tốc độ gió Kanomax Climomaster 6501, Japan;`
     - Kiểm tra rò rỉ lọc HEPA -> `Máy tạo sol khí ATI aerosol generator, model 6D và quang kế aerosol photometer, model 2i, USA;`
     - Kiểm tra hình thái dòng khí -> `Máy tạo khói khí dung (Smoke pencil / Aerosol generator);`
     - Cường độ sáng -> `Máy đo độ rọi Testo 540 / Testo 440, Germany;`
     - Đo độ ồn -> `Máy đo độ ồn Rion NL-53, Japan;`
     - Đo độ rung -> `Máy đo độ rung RION VM-85, Japan;`
     - Đo cường độ UVC -> `Máy đo cường độ bức xạ UVC Extech SDL470, USA;`
     - Đo nồng độ hạt bụi trong tủ -> `Máy đếm hạt TSI AeroTrak 9350, USA.`
     - Quy trình áp dụng: `QUATEST 3 3030:2025; EN 12469:2000; N2.TC.BSC.02` (tủ BSC) hoặc `N2.TC.LAF.01` (tủ LAF) hoặc `N2.TC.T.01` (tủ hút).

   3.4. **Mapping nhiệt độ / Kho lạnh (N2.TC.MP.01)**:
     - Thiết bị ghi nhiệt ẩm tự động -> `Thiết bị tự ghi nhiệt độ độ ẩm Haxo-8 (mapping);`
     - Thiết bị kiểm tra đối chứng -> `Thiết bị đo nhiệt độ độ ẩm Testo 625, Germany.`
     - Quy trình áp dụng: `QUATEST 3 3030:2025; N2.TC.MP.01`.

   3.5. **Vi khí hậu & Môi trường lao động (N2.TC.01 đến N2.TC.08)**:
     - **[CẬP NHẬT 24/11/2026]** Nhiệt độ, độ ẩm, bức xạ nhiệt -> `Thiết bị đo nhiệt độ, độ ẩm, bức xạ nhiệt skSATO SK-170GT, Japan;` (thay Testo 625); tốc độ gió -> `Máy đo tốc độ gió Kanomax Climomaster 6533-2G, Japan;`
     - Đo độ rọi -> `Máy đo độ sáng Testo 440 Kit Lux, Testo – Germany;`
     - Đo độ ồn -> `Thiết bị đo mức âm thanh Rion NL-53, Japan;`
     - Đo điện từ trường -> `Máy đo điện từ trường EMF450;` (thay Holaday HI-3604)
     - Đo bụi MTLĐ trực tiếp -> `Máy đo nồng độ bụi Temptop PMD 371, USA;` (thay Metone Aerocet 831 / bơm SKC)
     - Bơm lấy mẫu hơi khí -> `Bơm lấy mẫu khí Kitagawa ASP 1200, Japan;`
     - Đo khí CO2 -> `Thiết bị đo nồng độ CO2 trong không khí testo 440, Germany;`. Không dùng Drager X-am 8000 cho MTLĐ (chỉ giữ cho TVOC của profile `co2_mtld`).
     - Test nhanh hơi khí độc -> `Bơm hút khí Gastec / Kitagawa và ống phát hiện nhanh chuyên dụng.`

   *Tuyệt đối không đưa các thiết bị của chỉ tiêu không được yêu cầu.*
2. **Xuống dòng riêng biệt**:
   - Mỗi thiết bị đo đạc / lấy mẫu phải nằm trên một dòng riêng biệt để người đọc dễ kiểm soát danh mục thiết bị mang theo.
3. **Ký tự đầu dòng**:
   - Mỗi đầu dòng thiết bị bắt đầu bằng đúng 1 dấu gạch ngang (`- Tên thiết bị...`).
   - Xóa bỏ triệt để bullet / numbering tự động (`w:numPr`) của Word để tránh lỗi phát sinh trùng lặp 2 dấu gạch (`- - `).
4. **Căn lề và thụt đầu dòng**:
   - Thụt lề dòng đầu tiên (`first_line_indent`) và thụt lề trái (`left_indent`), lề phải (`right_indent`) của tất cả các dòng thiết bị phải set về đúng **0cm** (không thụt lề).

---

### PHẦN 4: QUY TẮC LÀM SẠCH THÔNG TIN TRẠNG THÁI & ĐỊNH LƯỢNG
1. **Loại bỏ thông tin trạng thái phòng sạch**:
   - Tuyệt đối **không nêu trạng thái phòng sạch** (ví dụ: *"ở trạng thái nghỉ"*, *"trạng thái nghỉ"*, *"trạng thái hoạt động"*, *"at-rest"*, v.v.) trong tên khu vực, địa điểm thực hiện hay trong bảng chỉ tiêu khi soạn thảo form mẫu.
2. **Loại bỏ thông số định lượng trong tên chỉ tiêu**:
   - Về phần chỉ tiêu: **Chỉ ghi tên của chỉ tiêu**, tuyệt đối **không ghi mức đánh giá, yêu cầu tiêu chuẩn hay giới hạn định lượng**.
   - *Ví dụ đúng*:
     - `Nồng độ hạt bụi`
     - `Nhiệt độ điểm sương khí nén`
     - `Vi sinh khí nén`
     - `Đo chênh áp`
   - *Ví dụ sai cần loại bỏ*:
     - `Nồng độ hạt bụi (kích thước 0,1-5,0um)` -> Sai
     - `Đo nồng độ hạt bụi (ISO Class 5 - Kích thước hạt ≥ 0,5µm)` -> Sai
     - `Đo nhiệt độ, độ ẩm trong không khí (Yêu cầu: 21-24°C, 55-65%)` -> Sai
3. **Giữ nguyên mã tiêu chuẩn kỹ thuật**:
   - Giữ nguyên các mã tiêu chuẩn phương pháp như `ISO 8573`, `ISO 14644`, không được xóa nhầm mã tiêu chuẩn thành cụm từ lơ lửng như `theo`.

---

### PHẦN 5: QUY TẮC BẢNG GHI CHÉP SỐ LIỆU ĐO (BM3030.01)
1. **Ô mã hồ sơ (Table 0, Cell 0,0)**:
   - Điền mã số hồ sơ `Q326D...` vào ô góc trên bên trái của Table 0.
2. **Số lượng bảng đo hiện trường (Khí nén)**:
   - Co giãn số lượng bảng ghi chép hiện trường (Table 1, Table 2, Table 3...) tương ứng với đúng số lượng vị trí yêu cầu trong hợp đồng/báo giá (ví dụ: 03 vị trí -> có đủ 3 bảng đo riêng biệt).
3. **Tiêu đề đặt phía trên mỗi bảng đo**:
   - Ghi tên chỉ tiêu tương ứng ngay phía trên bảng đo (Format căn lề trái tuyệt đối: `alignment = LEFT`, `<w:jc w:val="left"/>`, thụt lề = 0cm).
4. **Không can thiệp mã hóa thủ công (AI không tự cấp mã)**:
   - Phần điền vị trí đo/lấy mẫu: Để dòng chấm `…………………………………………` để kỹ thuật viên tự tay điền tên vị trí cụ thể tại hiện trường.
   - Phần thông số đĩa phía trên bảng vi sinh để dòng chấm trống:
     ```text
     + Ký hiệu đĩa:
     - TS lấy mẫu: ……………………
     - TS blank: ……………………
     ```

---

### PHẦN 6: QUY TẮC BIÊN BẢN LẤY MẪU / GIÁM ĐỊNH (BM3030.02)
1. **Thời gian thực hiện**:
   - Điều chỉnh thụt lề trái (`left_indent`) về **0cm**, thụt lề phải (`right_indent`) về **0cm** (`first_line_indent = 0cm`).
   - Xóa bỏ triệt để các thuộc tính `w:ind` âm và tab ẩn `\t`.
   - Cột nhãn: `Thời gian thực hiện` (không chèn `:` nếu dấu `:` ở ô riêng biệt).
   - Nội dung thời gian: `Bắt đầu từ.................... kết thúc lúc..................... ngày …./…./YYYY`.
2. **Bảng chỉ tiêu (Table 2) - Quy tắc KHÔNG ĐỂ TRỐNG (để dòng chấm điền tay):**
   - **Cột Vị trí lấy mẫu (đo)**: **TUYỆT ĐỐI KHÔNG ĐỂ TRỐNG**, để dòng chấm `………………………………………………………………………………………………` để người dùng/KTV điền tay vào.
   - **Cột Số lượng**: **TUYỆT ĐỐI KHÔNG ĐỂ TRỐNG**, để dòng chấm `…………` để người dùng điền tay vào.
   - **Cột Kí hiệu mẫu (nếu có)**: **TUYỆT ĐỐI KHÔNG ĐỂ TRỐNG**, để dòng chấm `…………` để người dùng điền tay vào (không tự cấp mã `1, 2, 3`).
   - **Cột Niêm phong mẫu (nếu có)**: **TUYỆT ĐỐI KHÔNG ĐỂ TRỐNG**, để dòng chấm `…………` để người dùng điền tay vào.
   - **Cột Phương pháp thực hiện**: Đánh dấu `x` vào ô tương ứng (`Đo` hoặc `Lấy mẫu`).
   - **Giãn dòng viết tay (Line Spacing 1.5)**: Thiết lập khoảng cách dòng **1.5 lines** (`line_spacing = 1.5`) cho tất cả paragraph trong các ô dữ liệu của Table 2 để tạo độ cao thông thoáng, giúp kỹ thuật viên viết tay thuận tiện, không bị chen chúc.
   - **Chống ngắt dòng qua 2 trang (`cantSplit`)**: Bắt buộc thiết lập thuộc tính `<w:cantSplit/>` cho toàn bộ các hàng của Table 2 để đảm bảo không có hàng nào bị ngắt đôi chia thành 2 trang giấy khác nhau.
3. **Chỉ tiêu thực hiện lấy mẫu vi sinh**:
   - Tự động trích xuất đúng chỉ tiêu vi sinh thực tế theo nội dung hồ sơ/báo giá (ví dụ: `Chỉ tiêu thực hiện lấy mẫu vi sinh bề mặt: Tổng số vi sinh vật hiếu khí, tổng số nấm men - nấm mốc` hoặc `Chỉ tiêu thực hiện lấy mẫu vi sinh trong không khí: ...`). Ẩn/xóa dòng này nếu hồ sơ không yêu cầu vi sinh.
4. **Quy tắc Mục 9 (Nội dung/ Quá trình thực hiện) trong Table 1 của BM3030.02 (BẮT BUỘC)**:
   - **Vị trí**: Nằm ngay phía dưới hàng đề mục `"9. Nội dung/ Quá trình thực hiện :"` tại cuối Table 1 trên Trang 1.
   - **Nội dung ghi sẵn chuẩn mực**:
     `Nội dung thực hiện theo báo giá số {quotation_number}. Vị trí thực hiện được chỉ định bởi đại diện {customer_name}, số lượng vị trí cụ thể như sau:`
   - **Quy cách chi tiết**:
     + **Số báo giá (`quotation_number`)**: Bắt buộc là **mã full đầy đủ** của số báo giá (ví dụ: `QONV26015439`, `QONV26016078`...). Tuyệt đối không để rút gọn cụt lủn hay in cứng `QONV2......`. Nếu hồ sơ chưa có số báo giá cụ thể, để sẵn placeholder chừa chấm điền tay: `QONV2……………`.
     + **Tên khách hàng (`customer_name`)**: Tự động trích xuất tên đơn vị yêu cầu / khách hàng đã qua hàm chuẩn hóa `normalize_customer_name` (ví dụ: `Công ty TNHH Hon Chuan Vietnam`, `Công ty Cổ phần P/S`, `Công ty TNHH Thực phẩm Quốc tế Nam Dương`...). Nếu chưa xác định tên, ghi `khách hàng`.
     + **Định dạng văn bản**: Font **Times New Roman, cỡ 13pt**, không in đậm (`bold=False`), căn lề trái, `space_before = 2pt`, `space_after = 2pt`, giãn dòng `line_spacing = 1.15`.
     + **Nối tiếp trực tiếp sang Table 2:** Phía dưới câu ghi sẵn, bắt đầu ngay Table 2 ở cuối Trang 1, TUYỆT ĐỐI KHÔNG chèn hàng chấm lửng giả ở Trang 1.
5. **Bố cục Table 2 liền mạch cho TẤT CẢ các hồ sơ (Phòng sạch, Khí nén, Vi sinh, BSC, Vi khí hậu, MTLĐ, Kho lạnh...):**
   - **Trang 1**: Table 2 bắt đầu ngay phía dưới câu ghi sẵn của Mục 9 (tuyệt đối không chèn ngắt trang trước Table 2, không chèn các hàng chấm lửng giả ở Trang 1). Hàng tiêu đề Table 2 và (các) hàng dữ liệu đầu tiên nằm gọn ở cuối Trang 1.
   - **Trang 2**: Các hàng dữ liệu còn lại tự động trôi sang đầu Trang 2 nhờ `<w:cantSplit/>`.
   - **Dưới Table 2 trên Trang 2**:
     * Dòng chỉ tiêu vi sinh (nếu có): `Chỉ tiêu thực hiện lấy mẫu vi sinh...`
     * Co giãn số hàng chấm lửng tràn lề (`.` * 147) vừa vặn trước khối chữ ký:
       - 1 mẫu / 1 tủ (Table 2 nằm trọn ở Trang 1): Có ngắt trang trước phần tử đầu Trang 2 và chèn 8-14 hàng chấm lửng lấp đầy Trang 2.
       - 2 mẫu (chuẩn như hồ sơ Hon Chuan Vietnam): Dưới Table 2 trên Trang 2 chèn đúng **2 hàng chấm lửng**.
       - 3 mẫu: Dưới Table 2 trên Trang 2 chèn **1 hàng chấm lửng**.
       - 4 mẫu trở lên: Không chèn thêm chấm lửng (0 hàng chấm).
       - Nếu có dòng chỉ tiêu vi sinh, tự động giảm bớt 1-2 hàng chấm lửng tương ứng.
     * `Ký hiệu phiếu ghi chép: {doc_number}/1;`
     * `Biên bản này được lập thành 2 bản: 1 bản giao đại diện đơn vị yêu cầu, 1 bản lưu hồ sơ tại TTKT 3`
     * Khối chữ ký: `Đại diện đơn vị yêu cầu` - `Người thực hiện`.
   - **Kết quả**: Tất cả hồ sơ luôn đạt chuẩn **CHÍNH XÁC 2 TRANG**, bố cục cân đối và chuyên nghiệp.

---

### PHẦN 7: QUY TẮC CHUẨN HÓA TÊN KHÁCH HÀNG, ĐỊA CHỈ & VIẾT HOA - VIẾT THƯỜNG
0. **Chuẩn hóa Tên khách hàng / Đơn vị yêu cầu khi bị IN HOA TOÀN BỘ**:
   - Nếu tên khách hàng bị in hoa toàn bộ, bắt buộc chuyển lại sang chữ hoa - thường chuẩn tên cơ quan, tổ chức tiếng Việt: viết hoa chữ cái đầu của từ/cụm từ chỉ loại hình và lĩnh vực hoạt động (`Công ty`, `Cổ phần`, `Chi nhánh`, `Kỹ nghệ`, `Súc sản`, `Thực phẩm`, `Công nghệ`, `Sinh học`...), viết hoa tên riêng (`Việt Nam`, `Cát Tân`...) và giữ nguyên in hoa các từ viết tắt (`TNHH`, `MTV`, `SX`, `TM`, `DV`, `P/S`...).
   - Ví dụ:
     + `CÔNG TY CỔ PHẦN VIỆT NAM KỸ NGHỆ SÚC SẢN` -> `Công ty Cổ phần Việt Nam Kỹ nghệ Súc sản`
     + `CÔNG TY CỔ PHẦN P/S` -> `Công ty Cổ phần P/S`
1. **Mở rộng các từ viết tắt địa chỉ thành nguyên chữ**:
   - `KCN` -> `Khu công nghiệp` (ví dụ: `KCN Tân Đông Hiệp B` -> `Khu công nghiệp Tân Đông Hiệp B`, `KCN Vĩnh Lộc` -> `Khu công nghiệp Vĩnh Lộc`)
   - `KCX` -> `Khu chế xuất` (ví dụ: `KCX Tân Thuận` -> `Khu chế xuất Tân Thuận`, `KCX Linh Trung` -> `Khu chế xuất Linh Trung`)
   - `KDC` -> `Khu dân cư` (ví dụ: `KDC Thới An` -> `Khu dân cư Thới An`)
   - `CCN` -> `Cụm công nghiệp` (ví dụ: `CCN Tân Mỹ` -> `Cụm công nghiệp Tân Mỹ`)
   - `KĐT` -> `Khu đô thị` (ví dụ: `KĐT Sala` -> `Khu đô thị Sala`, `KĐT Phú Mỹ Hưng` -> `Khu đô thị Phú Mỹ Hưng`)
   - `KKT` -> `Khu kinh tế` (ví dụ: `KKT Dung Quất` -> `Khu kinh tế Dung Quất`)
   - `KCNC` -> `Khu công nghệ cao`
   - `Tp.HCM` / `TP.HCM` / `TP. HCM` / `Tp. HCM` -> `Thành phố Hồ Chí Minh`
   - `QL.` / `QL` -> `Quốc lộ` (ví dụ: `QL. 1A` -> `Quốc lộ 1A`, `QL 13` -> `Quốc lộ 13`)
   - `TL.` / `TL` -> `Tỉnh lộ` (ví dụ: `TL. 8` -> `Tỉnh lộ 8`)
   - `Đ.` / `Đg.` -> `đường`
   - `P.` -> `phường` (nếu tên chữ: `P. Thới An` -> `phường Thới An`) / `Phường` (nếu tên số: `P. 1` -> `Phường 1`)
   - `Q.` -> `quận` (nếu tên chữ: `Q. Gò Vấp` -> `quận Gò Vấp`) / `Quận` (nếu tên số: `Q. 1` -> `Quận 1`)
   - `H.` -> `huyện` (ví dụ: `H. Hóc Môn` -> `huyện Hóc Môn`)
   - `TX.` -> `thị xã` (ví dụ: `TX. Bến Cát` -> `thị xã Bến Cát`)
   - `TT.` -> `thị trấn` (ví dụ: `TT. Hóc Môn` -> `thị trấn Hóc Môn`)
   - `X.` -> `xã` (ví dụ: `X. Tân Thới Hiệp` -> `xã Tân Thới Hiệp`)
   - `KP.` -> `Khu phố` (ví dụ: `KP. 1` -> `Khu phố 1`)
   - `TDP.` -> `Tổ dân phố` (ví dụ: `TDP. 5` -> `Tổ dân phố 5`)
2. **Ví dụ chuẩn hóa**:
   - Gốc: `"A87, KDC Thới An, Đường Lê Thị Riêng, Khu Phố 1, Phường Thới An, TP.HCM, Việt Nam"`
   - Sau chuẩn hóa: `"A87, Khu dân cư Thới An, đường Lê Thị Riêng, Khu phố 1, phường Thới An, Thành phố Hồ Chí Minh"`
3. **Loại bỏ chữ "Việt Nam" ở cuối**:
   - Tuyệt đối không ghi `Việt Nam`, `Vietnam`, `VN` ở phần cuối chuỗi địa chỉ.
4. **Thành phố trực thuộc Trung ương**:
   - Bắt buộc viết hoa chữ `Thành phố`:
     + `Thành phố Hồ Chí Minh` (thay thế cho `TP.HCM`, `Tp.HCM`, `TP. HCM`, `thành phố Hồ Chí Minh`...)
     + `Thành phố Hà Nội` (thay thế cho `TP. Hà Nội`, `Tp. Hà Nội`, `Hà Nội`...)
     + `Thành phố Đà Nẵng` (thay thế cho `TP. Đà Nẵng`, `Đà Nẵng`...)
     + `Thành phố Hải Phòng` (thay thế cho `TP. Hải Phòng`, `Hải Phòng`...)
     + `Thành phố Cần Thơ` (thay thế cho `TP. Cần Thơ`, `Cần Thơ`...)
     + `Thành phố Huế` (nếu có)
5. **Tỉnh và Thành phố thuộc tỉnh**:
   - Viết thường chữ `tỉnh` và `thành phố`:
     + Tỉnh: `tỉnh Bình Dương`, `tỉnh Đồng Nai`, `tỉnh Long An`, `tỉnh Bà Rịa - Vũng Tàu`...
     + Thành phố thuộc tỉnh: `thành phố Thủ Dầu Một`, `thành phố Dĩ An`, `thành phố Thuận An`, `thành phố Biên Hòa`, `thành phố Thủ Đức`...
6. **Quận / Huyện / Thị xã**:
   - Tên riêng là chữ -> viết thường: `quận Bình Thạnh`, `quận Tân Bình`, `quận Gò Vấp`, `huyện Hóc Môn`, `huyện Củ Chi`, `thị xã Bến Cát`...
   - Tên riêng là số -> viết hoa từ chỉ đơn vị: `Quận 1`, `Quận 12`...
7. **Phường / Xã / Thị trấn**:
   - Tên riêng là chữ -> viết thường: `phường Thới An`, `phường An Khánh`, `phường Bến Nghé`, `xã Tân Thới Hiệp`, `thị trấn Hóc Môn`...
   - Tên riêng là số -> viết hoa từ chỉ đơn vị: `Phường 1`, `Phường 14`...
8. **Khu phố / Thôn / Ấp / Tổ dân phố**:
   - `Khu phố`: viết hoa chữ `Khu`, viết thường chữ `phố` (`Khu phố 1`, `Khu phố 2`, `Khu phố Đông Chiêu`).
   - `Ấp`: viết hoa chữ `Ấp` (`Ấp 1`, `Ấp Bình Mỹ`...).
   - `Thôn`: viết hoa chữ `Thôn` (`Thôn 1`, `Thôn Phú Mỹ`...).
   - `Tổ dân phố`: viết hoa chữ `Tổ`, viết thường `dân phố` (`Tổ dân phố 1`, `Tổ dân phố 5`...).
9. **Đường / Phố / Đại lộ / Hẻm / Ngõ / Ngách**:
   - Viết thường danh từ chỉ loại đường giao thông: `đường Lê Thị Riêng`, `đường số 3`, `phố Tràng Tiền`, `đại lộ Mai Chí Thọ`, `hẻm 45`, `ngõ 12`...
10. **Khu công nghiệp / Khu chế xuất / Khu dân cư / Cụm công nghiệp**:
   - Mở rộng viết đầy đủ và viết hoa chữ cái đầu: `Khu công nghiệp Tân Bình`, `Khu chế xuất Tân Thuận`, `Khu dân cư Thới An`, `Cụm công nghiệp Tân Mỹ`... (viết hoa chữ `Khu`, `Cụm`, viết thường các từ sau `công nghiệp`, `chế xuất`, `dân cư`).
11. **Tra cứu Google / Web Search để xác minh địa danh & cập nhật đơn vị hành chính mới**:
   - **Được phép và khuyến khích tra cứu Google:** Khi thông tin địa chỉ hoặc địa điểm thực hiện trong yêu cầu / hợp đồng / báo giá viết tắt, ghép liền cụm danh từ hoặc thiếu từ chỉ cấp hành chính (ví dụ: `"KCN Vĩnh Lộc Hóc Môn, Tp.HCM"`), người soạn thảo/AI được phép và chủ động tra cứu Google / Web Search để xác định chính xác các cấp đơn vị hành chính và phân tách rõ ràng bằng dấu phẩy.
   - **Ví dụ chuẩn hóa:** `"KCN Vĩnh Lộc Hóc Môn, Tp.HCM"` $\rightarrow$ `"Khu công nghiệp Vĩnh Lộc, xã Hóc Môn, Thành phố Hồ Chí Minh"`.
   - **Lưu ý đơn vị hành chính mới cho tất cả các địa chỉ:** Luôn cập nhật theo đơn vị hành chính mới nhất (các phường/xã mới sau khi sắp xếp, sáp nhập tại TP.HCM, Hà Nội; các thành phố trực thuộc tỉnh mới thành lập như `thành phố Bến Cát`, `thành phố Tân Uyên`, `thành phố Thủ Đức`...). Tuyệt đối không giữ tên đơn vị hành chính cũ đã bị bãi bỏ hoặc sáp nhập.

---

### PHẦN 8: QUY TẮC ĐẶT TÊN THƯ MỤC LƯU FILE XUẤT (OUTPUT FOLDER)
1. **Cấu trúc chuẩn mực bắt buộc**:
   - Cú pháp: `"số hồ sơ - tên khách hàng - nội dung vắn tắc"`
   - Phân tách 3 thành phần chính bằng dấu khoảng trắng, gạch nối, khoảng trắng: ` - `.
   - Tuyệt đối không dùng dấu gạch dưới `_` làm dấu phân cách chính (không dùng `Q326D..._CONG TY...`, bắt buộc dùng dạng `Q326D... - [Tên KH] - [Nội dung]`).
2. **Ví dụ thực tế chuẩn mực**:
   - `Q326D124674 - Cát Tân - Phòng sạch`
   - `Q326D120711 - Greencans Industries - Nước uống`
   - `Q326D117382 - Thụy Khuê - Nuoc sinh hoat`
   - `Q326D124796 - K pack - vskk - swab`
3. **Quy ước từ viết tắt cho "Nội dung vắn tắt"**:
   - `VSKK`: Vi sinh không khí (hoặc viết thường `vskk`).
   - `VSBM`: Vi sinh bề mặt (hoặc viết thường `vsbm`, `swab`).
   - `vskk - swab`: Khi hồ sơ có cả lấy mẫu vi sinh trong không khí và vi sinh bề mặt (swab).
   - `Phòng sạch`: Đo kiểm chất lượng phòng sạch (nồng độ hạt bụi, chênh áp, nhiệt ẩm, trao đổi khí...).
   - `Khí nén`: Kiểm tra khí nén (tổng dầu, điểm sương, hạt bụi khí nén...).
   - `Khí Nitơ`: Kiểm tra khí nitơ có áp.
   - `Nước uống`: Kiểm nghiệm / lấy mẫu nước uống, nước ăn uống (QCVN 01-1).
   - `Nuoc sinh hoat` (hoặc `Nước sinh hoạt`): Nước sinh hoạt (QCVN 02).
   - `Nước thải`: Nước thải công nghiệp / y tế (QCVN 28, QCVN 40).
   - `MTLĐ`: Quan trắc môi trường lao động (CO2, TVOC, vi khí hậu, tiếng ồn, độ rung, bụi MTLĐ...).
   - `BSC`: Giám định tủ an toàn sinh học.
   - `LAF`: Đo kiểm tủ sạch dòng khí tầng.
   - `Tủ hút`: Kiểm tra tủ hút hóa chất.
   - `Kho lạnh`: Mapping nhiệt độ, độ ẩm kho lạnh, tủ bảo quản dược phẩm/vắc-xin.
4. **Quy tắc trích xuất "Tên khách hàng" vắn tắt**:
   - Lược bỏ tiền tố loại hình doanh nghiệp: `CÔNG TY TNHH`, `CÔNG TY CỔ PHẦN`, `CÔNG TY CP`, `CÔNG TY TNHH MTV`, `TỔNG CÔNG TY`, `TẬP ĐOÀN`, `DOANH NGHIỆP TƯ NHÂN`, `DNTN`, `CHI NHÁNH`, `CTY`...
   - Lược bỏ các cụm từ ngành nghề chung không mang tính định danh riêng ở đầu (như `CÔNG NGHỆ SẠCH`, `SẢN XUẤT THƯƠNG MẠI DỊCH VỤ`, `SX TM DV`...) nếu phía sau còn tên thương hiệu riêng rõ ràng (ví dụ: `CÔNG TY TNHH CÔNG NGHỆ SẠCH CÁT TÂN` -> `Cát Tân`).
   - Nếu có tên viết tắt trong ngoặc đơn (ví dụ: `(Vietcombank)`), ưu tiên sử dụng tên viết tắt đó.
   - Viết hoa dạng Title Case (`Cát Tân`, `Greencans Industries`, `Thụy Khuê`, `K pack`).

---

## 10. QUY TẮC ĐẶT TÊN CUỘC HỘI THOẠI (CONVERSATION TITLE)
- **Cấu trúc chuẩn mực:** `[Mã số hồ sơ] - [Tên khách hàng]` (ví dụ: `Q326D126588 - Hulo Care`, `Q326D120423 - Nhật Khang`).
- **Cập nhật tự động:** Ngay khi xác định hoặc nhận được mã hồ sơ và tên khách hàng trong phiên làm việc, tự động cập nhật tiêu đề cuộc hội thoại vào cơ sở dữ liệu `conversation_summaries.db` (`UPDATE conversation_summaries SET title = ? WHERE conversation_id = ?`) để danh sách lịch sử chat của Antigravity hiển thị chuẩn xác tên hồ sơ và khách hàng, giúp người dùng dễ dàng theo dõi và tra cứu.

---

## 11. QUY TẮC GHI ĐỊA ĐIỂM THỰC HIỆN (CHỪA HÀNG CHẤM ĐIỀN TAY)
- **Không điền cứng tên phòng/khu vực/hệ thống/ống dẫn cụ thể:** Dòng `Địa điểm thực hiện:` trong Table 0 của BM3030.01 và Table 1 của BM3030.02 tuyệt đối không điền sẵn tên cụ thể của vị trí hay thiết bị (như *"Ống dẫn khí nén hệ lọc"*, *"Phòng mổ"*, *"Khu vực cân"*...).
- **Cấu trúc chuẩn mực bắt buộc:** `Địa điểm thực hiện: [90 dấu chấm] - [Địa chỉ đã chuẩn hóa]`
- **[CẬP NHẬT 22/11/2026] Quy cách chấm:** Chừa **90 dấu chấm liền nhau trên một dòng** (không xuống dòng giữa dải chấm) ở phía trước ` - ` để kỹ thuật viên tự tay điền khu vực, phòng hoặc vị trí cụ thể khi đến hiện trường. Địa chỉ được phép xuống dòng thứ 2. Trong code dùng hằng số `LOCATION_DOTS` (`core/normalizer.py`), không gõ cứng số chấm ở nơi khác.
- **Ví dụ chuẩn:**
  + Thay vì ghi: `"Địa điểm thực hiện: Ống dẫn khí nén hệ lọc - Lô C20a-3, đường số 14, Khu công nghiệp Hiệp Phước, xã Hiệp Phước, huyện Nhà Bè, Thành phố Hồ Chí Minh"`
  + Bắt buộc ghi: `"Địa điểm thực hiện: .......................................................................................... - Lô C20a-3, đường số 14, Khu công nghiệp Hiệp Phước, xã Hiệp Phước, huyện Nhà Bè, Thành phố Hồ Chí Minh"`

---

## 12. MÃ HỒ SƠ ĐỂ TRẮNG (CHƯA CÓ SỐ PHIẾU YÊU CẦU)
- Khi người dùng yêu cầu **"mã hồ sơ để trắng"**: KHÔNG tự sinh mã ngẫu nhiên, KHÔNG dùng lại mã cũ của hồ sơ trước.
- Mã hồ sơ ghi trong biểu mẫu: `Q326D……………` (giữ tiền tố `Q326D`, chừa chấm để KTV điền tay) tại: ô `Số:` header Table 0 (BM01 & BM02), dòng `Phiếu yêu cầu & Nhận TĐ/GĐ:`, dòng `Ký hiệu phiếu ghi chép: Q326D……………/1,`.
- Ngày đăng ký (dòng `Phiếu yêu cầu & Nhận TĐ/GĐ ... Ngày:`) cũng để chấm lửng `…./…./YYYY`.
- Ô mã số trên header Table 0 được nới rộng (lấy bớt từ cột trống kế bên, tổng bề rộng bảng giữ nguyên) để placeholder nằm gọn 1 dòng.
- Tên thư mục / file dùng dạng ASCII an toàn: `Q326D...... - [Tên KH] - [Nội dung]` (ví dụ `Q326D...... - Nextern Việt Nam - Khí nén - Phòng sạch`).

---

## 13. TỰ ĐỘNG SAO CHÉP 1 BẢN SANG Ổ ĐĨA Z (LƯU TRỮ HỒ SƠ QUATEST 3)
- **Đường dẫn chuẩn bắt buộc:** `Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026`
- **Nguyên tắc hoạt động:**
  + Từ nay về sau, mọi hồ sơ khi được sinh ra (cả file `.docx` và file `.pdf` nếu có) đều được lưu tại thư mục cục bộ `output/[Thư mục hồ sơ]` VÀ đồng thời tự động sao chép nguyên vẹn 1 bản (cả thư mục) sang đường dẫn: `Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026\[Thư mục hồ sơ]`.
  + Tên thư mục con trên ổ Z phải trùng khớp 100% với tên thư mục output chuẩn theo Quy tắc 9 (`[số hồ sơ] - [tên khách hàng] - [nội dung vắn tắt]`).
  + Không xóa hoặc di chuyển file cục bộ, chỉ sao chép thêm 1 bản (mirror copy).
  + **[CẬP NHẬT 09/11/2026] Cấu trúc thư mục con trên ổ Z (BẮT BUỘC):**
    * `Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026\[Thư mục hồ sơ]\Thuc hien\` — các file biên bản / ghi chép do hệ thống sinh ra (.docx, .pdf).
    * `Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026\[Thư mục hồ sơ]\Scan\` — file gốc ban đầu (phiếu yêu cầu, PO... bản scan trong `Input/`).
    * Truyền file gốc bằng `--scan Input/xxxx.pdf` (CLI `generate_now.py`, nhiều file cách nhau `;`), `scan_files=` (`clone_form_set`) hoặc khóa `scan_files` / `scan` / `input_files` (`process_order`). Không truyền thì không tạo thư mục `Scan`.
    * **[CẬP NHẬT 14/11/2026] Tên file scan gốc:** file scan ban đầu trong `Input/` (tên máy scan kiểu `20261114144734901.pdf`) đổi tên thành `<Số hồ sơ> - <Tên khách hàng rút gọn> - Scan PYC.pdf` (ví dụ `Q326D125256 - Dược Sài Gòn - Scan PYC.pdf`) trước khi truyền vào `--scan`. Scan gộp nhiều hồ sơ thì tách từng file theo tên này.
    * Hồ sơ có nhiều bộ (ví dụ Sanofi `Q326D139329` tách `BSC - LAF` và `Nước uống`): mỗi thư mục hồ sơ đều có `Thuc hien` + `Scan` (chép cùng file scan gốc vào cả hai).
  + Nếu ổ đĩa Z chưa kết nối hoặc mạng bị ngắt quãng, hệ thống ghi nhận cảnh báo an toàn và bảo toàn trọn vẹn file cục bộ, không làm gián đoạn luồng sinh tài liệu.

---

## 14. HỒ SƠ MÔI TRƯỜNG LAO ĐỘNG (MTLĐ) THEO KẾ HOẠCH TỪNG PHÒNG — DÙNG CHUNG (20/11/2026)
- **Công cụ:** `generate_mtld.py` (module `core/mtld_builder.py`). Áp dụng cho mọi hồ sơ quan trắc MTLĐ có nhiều phòng/khu vực (khởi nguồn từ hồ sơ QUATEST 3 – KTN Cát Lái 2026). Hồ sơ chỉ đo CO2/TVOC văn phòng vẫn dùng profile `co2_mtld` cũ.
- **Đầu vào:** file Excel kế hoạch theo mẫu `templates/MTLD_KeHoach_mau.xlsx`: cột `Phòng | Chỉ tiêu | SL | Điểm đo` (Điểm đo dạng `P304: 02; P303: 01`). Điểm đo ghi không rõ số -> cả cụm thành 1 khu vực; sửa bằng `--fix "Phòng|Chỉ tiêu=Khu vực:SL;Khu vực:SL || ..."`.
- **BM3030.01 Ghi chép:** trang 1 thông tin + bảng theo dõi khối lượng theo phòng (cột Xong ☐) + bảng tra mã OWAS / QĐ 3733 (nếu có Ecgonomy). **Mỗi phòng 1 trang** gồm:
  + Bảng `Máy đo sử dụng` tick ☐ (chỉ hiện chỉ tiêu phòng đó có): Nhiệt ẩm `skSATO SK-170GT / Testo 440 / Climomaster 6533-2G`; Gió `Climomaster 6533-2G / 6501`; Bức xạ nhiệt `skSATO SK-170GT`; Điện từ trường `EMF450`; Ánh sáng `Testo 440 – máy 1 / Testo 440 – máy 2`; Ồn `Rion NL-53`; Bụi `Temptop PMD 371`; CO₂ `testo 440`; luôn có `☐ Khác: ……` (cập nhật 24/11/2026).
  + Bảng đo trực tiếp: mỗi khu vực ghi đủ tên ở từng dòng (KHÔNG dùng dấu 〃), **ô xám đậm (A6A6A6) = không đo chỉ tiêu đó tại điểm đó**. Cột Ồn ghi **tên file đo** (không ghi dBA); ồn dải tần chỉ 1 cột tên file đo.
  + Bảng lấy mẫu hơi khí độc / vi sinh: ghi sẵn tên chỉ tiêu; cột `Bơm` tick `☐ Buck ☐ SKC ☐ Kitagawa` (Buck Genie, SKC PCXR8, Kitagawa ASP 1200 – Japan); vi sinh `☐ Bio Culture`.
  + Bảng Ecgonomy (OWAS: Lưng/Tay/Chân/Vật nặng + vị trí lao động).
  + Chữ ký `Người thực hiện` 1 lần ở cuối.
- **BM3030.02 Biên bản:** mỗi chỉ tiêu 1 dòng, cột vị trí ghi sẵn các phòng thực hiện (`Phòng QT, DH...` / `Các phòng`); Số lượng, Kí hiệu, Niêm phong để chấm. Bảng ghi chú cuối giữ đường chấm sẵn của mẫu — KHÔNG gõ thêm dải dấu chấm (bị chồng 2 lớp).
- **Thiết bị sử dụng & Quy trình** tự lọc theo chỉ tiêu có trong kế hoạch (vi khí hậu/ánh sáng N2.TC.01, ồn N2.TC.02, bụi N2.TC.06, CO2 + hơi khí độc N2.TC.07, vi sinh N2.TC.09, Ecgonomy N2.TC.11, N2.TC.12).
- `--multi_day` khi đo nhiều ngày: `Bắt đầu từ ngày …./…./2026 kết thúc ngày …./…./2026`.
- Chỉ tiêu `Đánh giá yếu tố tiếp xúc nghề nghiệp` -> bảng riêng (Người LĐ, Yếu tố tiếp xúc, Thời gian tiếp xúc, PTBVCN). Cột `Ghi chú` của kế hoạch (vd `Mẫu gửi`) được ghi sẵn vào cột Ghi chú bảng lấy mẫu. Hồ sơ chưa chia phòng: Phòng = `Khu vực: ……`, Điểm đo = `……` -> ô khu vực để trống viết tay, biên bản ghi dải chấm.
