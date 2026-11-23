# QUATEST 3 FORM GENERATION RULES
# (Tự động tải cho tất cả các phiên làm việc với Biểu mẫu QUATEST 3)

Mọi thao tác chỉnh sửa code sinh biểu mẫu hoặc soạn thảo hồ sơ Word (BM3030.01 và BM3030.02) BẮT BUỘC phải tuân thủ các quy tắc sau:

## 1. Trích xuất Mã số hồ sơ & Ngày đăng ký
- **Vị trí lấy thông tin chuẩn:** Tại **Trang 1 của Phiếu yêu cầu** (ở phần *"Phần dành cho Trung tâm Kỹ thuật 3"*, nằm ngay phía trên chữ ký của người nhận yêu cầu / NVGĐ, ví dụ Dương Ngọc Bảo Hân).
- **Mã số hồ sơ:** Bắt buộc phải có đầy đủ tiền tố `Q326D...` (ví dụ: `Q326D122015`), tuyệt đối không được chỉ lấy cụm số `122015`.
- **Ngày đăng ký (nhận hồ sơ):** Lấy đúng theo dòng in phía trên chữ ký của NVGĐ ở trang 1 (ví dụ `Ngày 29 tháng 09 năm 2026` -> `29/09/2026`). Tuyệt đối KHÔNG lấy ngày của Báo giá (trang 5).
- Điền ngày đăng ký này vào dòng: `Phiếu yêu cầu & Nhận TĐ/GĐ: Q326D... Ngày: DD/MM/YYYY`.

## 2. Ngày hiện trường & Thời gian thực hiện (Chưa thực hiện hiện trường)
- **Header Table 0 (BM3030.02):** Ô `Ngày:` trên header Table 0 **KHÔNG ĐƯỢC ĐIỀN NGÀY ĐĂNG KÝ**. Bắt buộc để dạng chấm lửng `…./…./YYYY` (ví dụ: `…./…./2026`) vì đây là ngày thực hiện hiện trường, chưa diễn ra lúc nhận hồ sơ.
- **Thời gian thực hiện trong BM3030.01 & BM3030.02:**
  - Định dạng chuẩn: `Bắt đầu từ.................... kết thúc lúc..................... ngày …./…./YYYY` (ví dụ: `ngày …./…./2026`).
  - Viết hoa chữ `Bắt đầu`.
  - Xóa bỏ tiền tố ngày tháng cũ (như `Ngày ... tháng ... năm ..., `).
  - Ngày để dạng chấm lửng `ngày …./…./2026`, KHÔNG điền ngày đăng ký vào đây trừ khi có lịch hẹn hiện trường cụ thể.
  - **Cấu trúc hàng và tự động đánh số thứ tự trong Table 1 (cả BM3030.01 và BM3030.02):**
    + Cột nhãn: BẮT BUỘC giữ/gán `w:numPr` để tự động đánh số thứ tự liên tục theo các đề mục trên (1, 2, 3, 4, 5... -> 6. Thời gian thực hiện). Tuyệt đối không xóa thuộc tính `numPr` của ô này.
    + Cột giữa: Dấu `:` nằm riêng biệt ở cột giữa, căn giữa (`w:jc w:val="center"`). Nếu form mẫu cũ bị gộp cột (gridSpan >= 2) thì phải tách thành 2 ô riêng để dấu `:` nằm cùng cột giữa với các dòng khác.
    + Cột nội dung: `Bắt đầu từ.................... kết thúc lúc..................... ngày …./…./YYYY` (indent = 0cm, Times New Roman 13pt).

## 3. Font chữ, Cỡ chữ & Định dạng bảng
- Font: **Times New Roman**, cỡ **13pt** (`Pt(13)`).
- **Không in đậm** (`bold=False`) các thông tin điền mới (giữ đúng chuẩn biểu mẫu gốc).
- Chỉ in đậm tiêu đề chính của các bảng đo đặt phía trên bảng.
- **Ngày hiệu lực của form (`HL:...`) trên header**: BẮT BUỘC là `HL:01/08/2025` cho tất cả hồ sơ và biểu mẫu (BM3030.01 và BM3030.02).
- **Không có dấu cách trước dấu `:`**: Tuyệt đối không để dấu cách trước dấu hai chấm trong văn bản cũng như các ô bảng (`Số:`, `Ngày:`, `Đơn vị yêu cầu:`, `Địa chỉ:`, `Thời gian thực hiện:`...).
- **Căn lề dọc ô bảng (Cell Vertical Alignment) = Top**: BẮT BUỘC format thuộc tính `<w:vAlign w:val="top"/>` cho TẤT CẢ các cell của TẤT CẢ các bảng (Table 0, Table 1, Table 2...).
- **Spacing các phần tử trong bảng = 2 pt cho Top và Down**: BẮT BUỘC format paragraph spacing `space_before = 2 pt` (Top) và `space_after = 2 pt` (Down) cho tất cả các phần tử trong tất cả các bảng.
- **Cách dòng dưới mỗi bảng có cỡ chữ 6pt (`Pt(6)`):** Dưới mỗi bảng (Table 0, Table 1, Table 2, Table 3...) BẮT BUỘC chèn/duy trì một dòng trắng cách bảng (empty spacer paragraph) với font chữ **Times New Roman, cỡ 6pt** (`Pt(6)` hoặc `<w:sz w:val="12"/>`), `space_before = 0pt`, `space_after = 0pt`, `line = 120` (exact 6pt). Cỡ chữ 6pt đảm bảo phân tách mỹ quan giữa các bảng và các khối nội dung/chữ ký mà không chiếm nhiều khoảng trống dọc gây nhảy trang.
- **Hàng 1 và cột 1 trong bảng KHÔNG CẦN GHI SỐ HỒ SƠ**: Ô Hàng 1 Cột 1 của các bảng (như Table 1) để trống hoàn toàn, không điền mã số hồ sơ vào ô này.

## 4. Thiết bị sử dụng
- Tự động trích xuất đúng các thiết bị tương ứng với các chỉ tiêu có trong báo giá/yêu cầu:
  - **Khí nén (N2.TC.KN.01):**
    - Đo hạt đến 0,1 µm (`0,1um`, `0.1um`, `0,1-5,0um`...) -> `Máy đo nồng độ hạt bụi Lasair III 110 và phụ kiện HPD II, USA;`
    - Đo hạt thông thường (≥ 0,3 µm) -> `Máy đo nồng độ hạt AeroTrak 9350 và phụ kiện HPD III, USA;`
    - Điểm sương -> `Thiết bị đo nhiệt độ điểm sương;`
    - Vi sinh khí nén -> `Thiết bị lấy mẫu vi sinh khí nén SAS Pinocchio Super II.`
    - Tổng dầu -> 2 dòng (cập nhật 11/11/2026): `- Bộ dụng cụ lấy mẫu tổng dầu trong khí nén;` và `- Giấy lọc thủy tinh Pall life Sciences, ống than hoạt tính ORBO®.` (dòng cuối dùng `;` nếu còn thiết bị khác phía sau)
  - **Phòng sạch (N2.TC.PS.01 & N2.TC.09):**
    - Đo hạt ≥ 0,3 µm (ISO 7-9) -> `Thiết bị đo nồng độ hạt TSI AeroTrak 9350, USA;`
    - Đo hạt đến 0,1 µm (ISO 1-6) -> `Máy đo nồng độ hạt bụi PMS Lasair III 110, USA;`
    - Nhiệt độ, độ ẩm -> `Thiết bị đo nhiệt độ, độ ẩm trong không khí Testo 625, Germany;`
    - Chênh áp -> `Thiết bị đo chênh lệch áp suất TSI 8375, USA;`
    - Độ ồn -> `Thiết bị đo độ ồn Rion NL 53, Japan;`
    - Độ sáng -> `Thiết bị đo độ sáng Testo 440, Germany;`
    - Vi sinh không khí -> `Thiết bị lấy mẫu vi sinh trong không khí Buck – Bio Culture B30120, USA.`
  - **Tủ an toàn sinh học (BSC), LAF, Tủ hút (N2.TC.BSC.02, N2.TC.LAF.01, N2.TC.T.01):**
    - Inflow -> `Thiết bị đo lưu lượng TSI 8375/8380, USA;`
    - Downflow -> `Máy đo tốc độ gió Kanomax Climomaster 6501, Japan;`
    - Rò rỉ HEPA -> `Máy tạo sol khí ATI aerosol generator, model 6D và quang kế aerosol photometer, model 2i, USA;`
    - Khói -> `Máy tạo khói khí dung (Smoke pencil / Aerosol generator);`
    - Độ rọi -> `Máy đo độ rọi Testo 540 / Testo 440, Germany;`
    - Độ ồn -> `Máy đo độ ồn Rion NL-53, Japan;`
    - Độ rung -> `Máy đo độ rung RION VM-85, Japan;`
    - Bức xạ UVC -> `Máy đo cường độ bức xạ UVC Extech SDL470, USA;`
    - Đếm hạt trong tủ -> `Máy đếm hạt TSI AeroTrak 9350, USA.`
  - **Mapping nhiệt độ / Kho lạnh (N2.TC.MP.01):**
    - `Thiết bị tự ghi nhiệt độ độ ẩm Haxo-8 (mapping);`
    - `Thiết bị đo nhiệt độ độ ẩm Testo 625, Germany.`
  - **MTLĐ & Vi khí hậu (N2.TC.01 - N2.TC.08):**
    - Vi khí hậu -> `Máy đo nhiệt độ, độ ẩm Testo 625, Germany;`, `Máy đo tốc độ gió Kanomax Climomaster 6533-2G, Japan;`
    - Độ rọi -> `Máy đo độ sáng Testo 440 Kit Lux, Testo – Germany;`
    - Độ ồn -> `Thiết bị đo mức âm thanh Rion NL-53, Japan;`
    - Bụi MTLĐ -> `Bơm lấy mẫu bụi cá nhân SKC / Buck Genie và cassette màng lọc;`
    - CO2 / TVOC -> `Thiết bị đo nồng độ CO2 trong không khí testo 440, Germany;`, `Thiết bị đo nồng độ TVOC Drager Xam 8000, Germany.`
  - KHÔNG đưa thiết bị của chỉ tiêu không được yêu cầu.
- Xuống dòng riêng biệt cho mỗi thiết bị.
- Mỗi đầu dòng bắt đầu bằng đúng 1 dấu gạch ngang (`- Tên thiết bị...`).
- Xóa `w:numPr` để tránh lỗi phát sinh 2 dấu gạch (`- - `).
- Thụt lề `left_indent = 0cm`, `right_indent = 0cm`, `first_line_indent = 0cm`.

## 5. Quy tắc ghi Nội dung yêu cầu & Tên chỉ tiêu kèm Chủ thể lấy mẫu / đo
- **Nội dung yêu cầu (Mục 2 trong Table 1 của BM3030.01 và Mục 4 trong Table 1 của BM3030.02):**
  + Chỉ ghi đúng và ngắn gọn nội dung công việc thực tế được yêu cầu (ví dụ: `Đo đạc và đánh giá hệ thống LAF tại khu vực cân nguyên liệu.`, `Kiểm tra chất lượng hệ thống khí nén.`, `Đo kiểm chất lượng không khí phòng sạch.`).
  + **TUYỆT ĐỐI KHÔNG ghi kèm viện dẫn tiêu chuẩn** như `theo ISO 14644-1:2015`, `theo ISO 8573`, `theo QCVN...` vào dòng Nội dung yêu cầu, vì các tiêu chuẩn và phương pháp thực hiện này đã được trình bày riêng biệt và đầy đủ tại mục *"Quy trình thực hiện / Tiêu chuẩn áp dụng"*.
  + **Nằm trên cùng một hàng tại cột 3 (Column 2) ngay sau dấu `:`**:
    * Trong cả BM3030.01 và BM3030.02, nội dung yêu cầu bắt buộc phải nằm ở cột 3 ngay sau dấu hai chấm `:` trên cùng một dòng đề mục (`4. Nội dung yêu cầu : [Nội dung yêu cầu]`).
    * Tuyệt đối không để trống ô cột 3 sau dấu `:` rồi đẩy nội dung xuống một hàng gộp riêng bên dưới (tránh lỗi văn bản bị thụt lề trái và dấu `:` bị trơ trọi một mình trên hàng trên).
    * Tự động xóa bỏ hàng rác phụ của template cũ bên dưới dòng `Nội dung yêu cầu` nếu có.
- **Bắt buộc ghi đầy đủ `[Hành động: Đo / Lấy mẫu...] + [Tên chỉ tiêu] + [Chủ thể lấy mẫu / đo]`:**
  + Khi ghi tên chỉ tiêu tại cột `Loại mẫu/ Mô tả/ Chỉ tiêu` (Table 2 của BM3030.02) cũng như danh mục chỉ tiêu, **BẮT BUỘC phải có chủ thể / môi trường lấy mẫu hoặc đo đi kèm**, tuyệt đối không ghi cụt lủn chỉ mỗi tên thông số (không ghi trống không là `Nồng độ hạt bụi`, `Nhiệt độ điểm sương`, `Hàm lượng tổng dầu`, `Đo chênh áp`).
  + **Ví dụ chuẩn mực theo từng chủ thể:**
    * **Phòng sạch:** `Đo nồng độ bụi phòng sạch` (hoặc `Đo nồng độ hạt bụi phòng sạch`), `Đo nhiệt độ, độ ẩm phòng sạch`, `Đo chênh áp phòng sạch`, `Đo độ ồn phòng sạch`, `Đo độ sáng phòng sạch`, `Đo lưu lượng và tính toán số lần trao đổi khí phòng sạch`, `Lấy mẫu vi sinh trong không khí phòng sạch`, `Lấy mẫu vi sinh bề mặt`.
    * **Khí nén / Khí có áp:** `Đo nồng độ bụi trong khí nén` (hoặc `Đo nồng độ hạt bụi trong khí nén`), `Lấy mẫu tổng dầu trong khí nén`, `Đo nhiệt độ điểm sương trong khí nén`, `Lấy mẫu vi sinh trong khí nén`, `Lấy mẫu khí Nitơ`, `Đo nồng độ bụi trong khí có áp`.
    * **Nước:** `Lấy mẫu nước uống`, `Lấy mẫu nước sinh hoạt`, `Lấy mẫu nước thải`, `Lấy mẫu nước tinh khiết`, `Lấy mẫu nước RO`, `Lấy mẫu nước cấp`, `Lấy mẫu nước ngầm`, `Lấy mẫu nước mặt`.
    * **Môi trường lao động (MTLĐ):** `Đo nồng độ CO2 trong không khí môi trường lao động`, `Đo nồng độ TVOC trong không khí môi trường lao động`, `Đo vi khí hậu môi trường lao động`, `Đo độ ồn môi trường lao động`, `Đo độ sáng môi trường lao động`, `Lấy mẫu bụi môi trường lao động`.
    * **Tủ BSC / LAF / Tủ hút / Kho lạnh:** `Đo nồng độ bụi trong tủ an toàn sinh học`, `Đo nồng độ bụi trong tủ LAF`, `Đánh giá độ đồng đều nhiệt độ kho bảo quản thuốc` (hoặc `kho lạnh`).
- **Loại bỏ trạng thái phòng sạch & thông số định lượng:**
  + KHÔNG nêu trạng thái phòng sạch ("trạng thái nghỉ", "ở trạng thái nghỉ", "trạng thái hoạt động", "at-rest"...) trong tên khu vực hay chỉ tiêu.
  + KHÔNG ghi mức đánh giá hay yêu cầu định lượng vào tên chỉ tiêu (như `kích thước 0,1-5,0um`, `ISO Class 5`, `21-24°C`, `+8-12 Pa`, `≤ 55 dBA`).
  + Giữ nguyên các mã tiêu chuẩn phương pháp (ví dụ: `ISO 8573`, `ISO 14644`) tại mục Quy trình thực hiện, chỉ loại bỏ các phân cấp phòng sạch `ISO Class 1-9`.

## 6. Bảng ghi chép hiện trường (BM3030.01)
- Hàng 1 và cột 1 trong bảng (Table 1): Không ghi số hồ sơ (để trống).
- Tiêu đề đặt phía trên bảng đo phải **Format lề trái tuyệt đối** (`align=LEFT`, `<w:jc w:val="left"/>`, indents = 0cm).
- **Khí nén:** Co giãn số lượng bảng ghi chép hiện trường (Table 1, Table 2, Table 3...) tương ứng với đúng số lượng vị trí yêu cầu trong hợp đồng/báo giá (ví dụ: 03 vị trí -> có đủ 3 bảng đo riêng biệt).
- **Phòng sạch / Môi trường:** Co giãn số dòng theo số lượng dự kiến + bổ sung **2 - 3 dòng trắng dự phòng**.
- **Quy tắc định lượng dải chấm lửng (`...` / `…`) và chống ngắt dòng làm nhảy trang trong BM3030.01:**
  + **Tất cả các dải chấm lửng BẮT BUỘC phải nằm gọn trên 1 dòng duy nhất**, tuyệt đối không để dải chấm quá dài làm Word ngắt dòng sang dòng thứ 2 (rơi vài dấu chấm `...` xuống dưới), làm tăng chiều cao ô và đẩy bảng nhảy trang sang trang tiếp theo.
  + **Hàng 0 Cột 0 (Tên khu vực thực hiện):**
    * Paragraph 1: `   Tên khu vực thực hiện: ………` (khoảng 3-4 ký tự `…` / 9-12 chấm).
    * Paragraph 2 & 3: `   …………………………………` (khoảng 11-12 ký tự `…` / 30-36 chấm).
    * Giữ nguyên cấu trúc 3 paragraph gốc với font Times New Roman 11pt, in nghiêng (`italic=True`). Tuyệt đối không để dải chấm quá dài (như 43 chấm) làm ô bị phình từ 3 dòng lên 5 dòng. TUYỆT ĐỐI KHÔNG điền địa chỉ công ty vào đây.
  + **Khối thông tin phép đo (Hàng 1 Cột 1 - Particle, Tốc độ gió, Cường độ sáng, Hướng dòng khí...):**
    * Dòng thời gian: `+ Thời gian lấy mẫu: ………………………………` / `+ Thời gian thực hiện: ………………………………` (khoảng 12-14 ký tự `…`, vừa vặn 1 dòng).
    * Dòng vị trí khác: `+ Vị trí kiểm tra: ... khác: ……………………` (khoảng 6-8 ký tự `…`, vừa vặn 1 dòng).
    * Dòng `File zeroing` và `File đo` trong Table 2: `[checkbox] File zeroing ………………………………` và `[checkbox] File đo………………………………………` (giữ nguyên tab stop, checkbox Wingdings, rút gọn chấm lửng để `File đo` không bị ngắt cụm rớt dòng thành 2 - 3 dòng).
  + **Cột Vị trí đo tại các hàng dữ liệu (Bảng 3, 4, 5...):**
    * Dải chấm điền tay: `………………………………………………` (khoảng 18-20 ký tự `…` / 40-44 chấm), vừa vặn ô bề rộng 261 pt (5227 dxa), không bị ngắt đôi sang dòng thứ 2.
  + **Khóa cố định bề rộng bảng (`<w:tblLayout w:type="fixed"/>`):**
    * BẮT BUỘC áp dụng thuộc tính `<w:tblLayout w:type="fixed"/>` cho Table 2 và toàn bộ các bảng đo mở rộng trong BM3030.01.
    * Ngăn Microsoft Word tự động co hẹp cột bên phải khi render, đảm bảo khối hiện trạng HVAC (`ngang`, `tường`, `xuất (nghỉ)`) không bao giờ bị rớt dòng.
- **Giữ nguyên 100% format chuẩn của Bảng đo hạt bụi (Table 2) trong BM3030.01:**
  + Tuyệt đối **KHÔNG chèn tiêu đề** "Đo nồng độ hạt bụi" phía trên Table 2 (biểu mẫu gốc không có tiêu đề này).
  + **Hàng 0 Cột 1:** Giữ nguyên 100% hiện trạng HVAC, các ô checkbox Wingdings và tab stop của biểu mẫu gốc.
  + **Hàng 1 Cột 0:** Để trống hoàn toàn như mẫu gốc (tuyệt đối KHÔNG điền "CHỈ TIÊU THỰC HIỆN").
  + **Hàng 1 Cột 1:** Giữ nguyên 100% khối `Particle` với tất cả các ô checkbox Wingdings (Lasair III 110, DPC 9350, chiều cao 1m, 1,2 m, file zeroing, file đo...). Tuyệt đối KHÔNG ghi đè bằng plain text làm mất checkbox và tab stop.
  + **Hàng 2:** Giữ nguyên tiêu đề 2 cột `Ký hiệu vị trí đo` và `File ghi`.
  + **Từ Hàng 3 trở đi:** Giữ nguyên số thứ tự tự động (`numPr`) ở Cột 0 và ô trống ở Cột 1 (chiều cao 454 dxa, căn giữa `vAlign=center`), tuyệt đối KHÔNG ghi đè `VT-01` hay `File đo: ………………………`. Chỉ co giãn số hàng theo số lượng vị trí yêu cầu (`target = 3 + num_dust + extra_rows`).
  + **Miễn trừ ghi đè format:** Các bảng ghi chép hiện trường BM3030.01 (Table 2...) được giữ nguyên định dạng gốc, không bị ghi đè thuộc tính `vAlign=top` và `spacing=2pt` của các bảng thông tin.
- Vị trí đo để dòng chấm `…………………………………………` để kỹ thuật viên tự ghi tay.
- Thông số đĩa trên đầu bảng vi sinh để dòng chấm trống (`- TS lấy mẫu: ……………………`, `- TS blank: ……………………`).
- **Chữ ký `Người thực hiện` chỉ ký 1 LẦN ở cuối ghi chép:** KHÔNG đặt `Người thực hiện` dưới từng bảng đo (từng vị trí khí nén, bảng bụi, bảng thứ cấp...). Toàn bộ BM3030.01 chỉ có đúng 1 đoạn chữ ký, căn phải, nằm sau bảng đo cuối cùng của tài liệu (hàm `ensure_single_signature_at_end()` trong `core/bm01_builder.py`).
- **Tủ an toàn sinh học (BSC):**
  + **Trang đầu của BM3030.01:** Thiết lập Table 1 theo đúng cấu trúc bảng biểu phòng sạch (Header table QUATEST 3, Table 0, Table 1 có đầy đủ các mục: Đơn vị yêu cầu, Nội dung yêu cầu, Địa điểm thực hiện, Thời gian thực hiện, Thiết bị sử dụng, Quy trình thực hiện, Nội dung/Quá trình thực hiện).
  + **Hàng chấm điền tay:** Phía dưới Table 1 trên trang đầu, chèn các hàng chấm lửng tràn hết toàn bộ bề rộng margin layout từ lề trái sang lề phải (độ dài 147 dấu chấm `.` trên mỗi dòng để vừa vặn chiều rộng bảng mà không bị rớt dòng) và lấp đầy toàn bộ phần trang trống còn lại của trang 1 (11 hàng chấm với `space_before=2pt`, `space_after=2pt`, `line_spacing=1.15`) để kỹ thuật viên có tối đa khoảng trống ghi chép hiện trường, tuyệt đối không để đoạn chấm ngắn ngủn lơ lửng hay để trống nửa dưới trang giấy.
  + **Tích hợp `Ghi Chep BSC.pdf`:** Nối tệp `Ghi Chep BSC.pdf` (3 trang bảng đo BSC) vào ngay sau trang đầu của biểu mẫu ghi chép BM3030.01, giúp tài liệu hoàn chỉnh 4 trang (trang 1 thông tin + 3 trang bảng đo chi tiết), chỉ cần soạn trang đầu của ghi chép.

## 7. File Biên bản lấy mẫu / Giám định (BM3030.02)
- Hàng "Thời gian thực hiện": Đồng bộ chuẩn 3 cột với Table 1: Cột 1 có số thứ tự tự động (`w:numPr`), Cột 2 là dấu `:`, Cột 3 là nội dung thời gian (`left_indent = 0cm`, `right_indent = 0cm`, xóa các thuộc tính `w:ind` âm và tab ẩn `\t`). Nội dung chuẩn: `Bắt đầu từ.................... kết thúc lúc..................... ngày …./…./YYYY`.
- **Bảng chỉ tiêu (Table 2) - Quy tắc KHÔNG ĐỂ TRỐNG (để dòng chấm điền tay) & Ghi rõ Chủ thể lấy mẫu/đo:**
  - Cột `Tên gọi/ Khu vực/ Vị trí lấy mẫu (đo)`: **TUYỆT ĐỐI KHÔNG ĐỂ TRỐNG**, để dòng chấm `………………………………………………………………………………………………` để người dùng/KTV điền tay vào.
  - Cột `Loại mẫu/ Mô tả/ Chỉ tiêu`: **BẮT BUỘC ghi đầy đủ hành động + chỉ tiêu + chủ thể lấy mẫu/đo** (ví dụ: `Đo nồng độ bụi phòng sạch`, `Đo nồng độ bụi trong khí nén`, `Lấy mẫu tổng dầu trong khí nén`, `Đo nhiệt độ điểm sương trong khí nén`, `Lấy mẫu nước uống`, `Lấy mẫu nước sinh hoạt`...). Tuyệt đối không ghi cụt lủn thiếu chủ thể (như chỉ ghi `Nồng độ hạt bụi` hay `Tổng dầu`).
  - Cột `Số lượng`: **TUYỆT ĐỐI KHÔNG ĐỂ TRỐNG**, để dòng chấm `…………` để người dùng điền tay vào.
  - Cột `Kí hiệu mẫu (nếu có)`: **TUYỆT ĐỐI KHÔNG ĐỂ TRỐNG**, để dòng chấm `…………` để người dùng điền tay vào (không tự cấp mã `1, 2, 3`).
  - Cột `Niêm phong mẫu (nếu có)`: **TUYỆT ĐỐI KHÔNG ĐỂ TRỐNG**, để dòng chấm `…………` để người dùng điền tay vào.
  - Cột `Phương pháp thực hiện`: Đánh dấu `x` vào ô tương ứng (`Đo` hoặc `Lấy mẫu`).
  - **Giãn dòng viết tay:** Thiết lập khoảng cách dòng **1.5 lines** (`line_spacing = 1.5`) cho các paragraph trong các ô dữ liệu để tạo khoảng trống cao ráo cho KTV viết tay vào các dòng chấm.
  - **Chống ngắt dòng qua 2 trang:** Bắt buộc áp dụng thuộc tính `<w:cantSplit/>` cho tất cả các hàng (`trPr`) của bảng để các hàng không bao giờ bị cắt đôi sang 2 trang.
- **Xử lý bố cục liền mạch & Ngắt trang Table 2 trên BM3030.02 (BẮT BUỘC ÁP DỤNG CHO TẤT CẢ CÁC HỒ SƠ VỀ SAU):**
  + **Bỏ cách dòng thừa:** Xóa bỏ toàn bộ các paragraph rỗng thừa giữa Table 0 và Table 1 (chỉ duy trì 1 dòng spacer 6pt), và giữa Table 1 và Table 2.
  + **BẮT BUỘC Table 2 bắt đầu ngay trên Trang 1:**
    * Table 2 bắt đầu ngay phía dưới câu văn của Mục 9 (`Nội dung thực hiện theo báo giá số... Vị trí thực hiện được chỉ định bởi đại diện...`).
    * **TUYỆT ĐỐI KHÔNG chèn ngắt trang trước Table 2** và **TUYỆT ĐỐI KHÔNG chèn hàng chấm lửng giả ở cuối Trang 1**.
    * Hàng tiêu đề Table 2 và hàng dữ liệu đầu tiên (hoặc các hàng vừa vặn) nằm gọn ở cuối Trang 1.
  + **Phân bổ Table 2 sang Trang 2 & Co giãn dòng chấm lửng:**
    * Khi có từ 2 hàng dữ liệu trở lên, các hàng tiếp theo tự động trôi sang đầu Trang 2 nhờ thuộc tính chống cắt dòng `<w:cantSplit/>`.
    * Phía dưới Table 2 trên Trang 2:
      - Dòng chỉ tiêu vi sinh (nếu hồ sơ có yêu cầu): `Chỉ tiêu thực hiện lấy mẫu vi sinh bề mặt: ...` hoặc `Chỉ tiêu thực hiện lấy mẫu vi sinh trong không khí: ...`.
      - Co giãn số hàng chấm lửng tràn lề (`.` * 147) vừa vặn để trang giấy cân đối trước khối chữ ký:
        + Nếu có 1 mẫu / 1 tủ (Table 2 nằm trọn ở Trang 1): Trang 2 bắt đầu bằng ngắt trang (`<w:pageBreakBefore/>`) và có khoảng 8-14 hàng chấm lửng lấp đầy Trang 2.
        + Nếu có 2 mẫu (chuẩn như hồ sơ Hon Chuan Vietnam): Dưới Table 2 trên Trang 2 chèn đúng **2 hàng chấm lửng**.
        + Nếu có 3 mẫu: Dưới Table 2 trên Trang 2 chèn **1 hàng chấm lửng**.
        + Nếu có 4 mẫu trở lên: **Không chèn chấm lửng** (0 hàng chấm) để bảo đảm khối chữ ký nằm trọn vẹn ở Trang 2.
        + Nếu có dòng chỉ tiêu vi sinh, tự động giảm bớt 1-2 hàng chấm lửng tương ứng.
      - Dòng `Ký hiệu phiếu ghi chép: {doc_number}/1;`
      - Dòng `Biên bản này được lập thành 2 bản: 1 bản giao đại diện đơn vị yêu cầu, 1 bản lưu hồ sơ tại TTKT 3`
      - Khối chữ ký: `Đại diện đơn vị yêu cầu` - `Người thực hiện`.
    * **Kết quả:** Mọi hồ sơ (Phòng sạch, Khí nén, Vi sinh, BSC, Vi khí hậu, MTLĐ, Kho lạnh...) luôn đạt chuẩn **CHÍNH XÁC 2 TRANG HOÀN HẢO**, bố cục liền mạch, không có khoảng trắng thừa và không bao giờ bị nhảy sang Trang 3.
- **Chỉ tiêu thực hiện lấy mẫu vi sinh:** Tự động trích xuất đúng chỉ tiêu vi sinh thực tế theo nội dung hồ sơ/báo giá (ví dụ: `Chỉ tiêu thực hiện lấy mẫu vi sinh bề mặt: Tổng số vi sinh vật hiếu khí, tổng số nấm men - nấm mốc` hoặc `Chỉ tiêu thực hiện lấy mẫu vi sinh trong không khí: ...`). Ẩn/xóa dòng này nếu hồ sơ không yêu cầu vi sinh.
- **Quy tắc Mục 9 (Nội dung/ Quá trình thực hiện) trong BM3030.02 (BẮT BUỘC áp dụng cho mọi biên bản sau này):**
  + **Vị trí hiển thị:** Tại cuối Table 1 của BM3030.02, ngay phía dưới dòng đề mục `"9. Nội dung/ Quá trình thực hiện :"`.
  + **[NEW 09/11/2026] Không có số báo giá QONV nhưng có số PO của khách hàng -> dùng số PO:** câu Mục 9 đổi thành `Nội dung thực hiện theo đơn đặt hàng (PO) số {po_number}. Vị trí thực hiện được chỉ định bởi đại diện {customer_name}, số lượng vị trí cụ thể như sau:` (ví dụ `E900192188` - Sanofi). Truyền bằng `--po E900192188` (CLI `generate_now.py`), `po_number=` (`clone_form_set`) hoặc khóa `po_number` / `po` / `so_po` (`process_order`). Có số báo giá thật thì vẫn ưu tiên số báo giá; không có cả hai thì để `QONV2……………`.
  + **Câu văn ghi sẵn bắt buộc:**
    `Nội dung thực hiện theo báo giá số {quotation_number}. Vị trí thực hiện được chỉ định bởi đại diện {customer_name}, số lượng vị trí cụ thể như sau:`
  + **Quy cách chi tiết:**
    * **Số báo giá (`quotation_number`):** BẮT BUỘC là **mã full đầy đủ** của số báo giá (ví dụ: `QONV26015439`, `QONV26016078`...). Tuyệt đối không để rút gọn cụt lủn hay in cứng `QONV2......`. Nếu hồ sơ chưa có số báo giá cụ thể, để sẵn placeholder chừa chấm điền tay: `QONV2……………`.
    * **Đại diện khách hàng (`customer_name`):** Tự động điền tên đầy đủ của đơn vị yêu cầu / khách hàng sau khi đã qua chuẩn hóa (`normalize_customer_name`), ví dụ: `Công ty TNHH Hon Chuan Vietnam`, `Công ty Cổ phần P/S`, `Công ty TNHH Thực phẩm Quốc tế Nam Dương`... Nếu chưa có thông tin khách hàng, để chữ `khách hàng`.
    * **Định dạng văn bản:** Font **Times New Roman, cỡ 13pt**, không in đậm (`bold=False`), căn lề trái, khoảng cách đoạn `space_before = 2pt`, `space_after = 2pt`, giãn dòng `line_spacing = 1.15`.
    * **Nối liền ngay sang Table 2:** Phía dưới câu ghi sẵn trên, bắt đầu ngay Table 2 ở cuối Trang 1, TUYỆT ĐỐI KHÔNG chèn hàng chấm lửng giả ở giữa Mục 9 và Table 2.
- **Bố cục Table 2 cho hồ sơ Tủ an toàn sinh học (BSC) (1 tủ hoặc nhiều tủ / chừa trống tên tủ):**
  + **Trang 1:** Table 2 bắt đầu ngay phía dưới câu dẫn Mục 9 (`Nội dung thực hiện theo báo giá số... Vị trí thực hiện được chỉ định bởi đại diện...`). Tuyệt đối KHÔNG chèn ngắt trang trước Table 2 và KHÔNG chèn các hàng chấm lửng ở Trang 1.
  + **Hàng tiêu đề Table 2:** BẮT BUỘC có thuộc tính `<w:tblHeader/>` để Microsoft Word tự động lặp lại hàng tiêu đề của bảng ở đầu Trang 2 khi có nhiều tủ.
  + **Trường hợp 1 tủ BSC:**
    * **Trang 1:** Table 2 (gồm tiêu đề + hàng Tủ 1) nằm gọn gàng, đẹp mắt ở cuối Trang 1 ngay sau Mục 9.
    * **Trang 2:** Chèn ngắt trang ngay sau Table 2; Trang 2 có 14 hàng chấm lửng tràn lề layout (`.` * 147) dành cho KTV ghi chép hiện trường, và kết thúc bằng khối chữ ký rộng rãi, cao ráo.
  + **Trường hợp 2 tủ BSC:**
    * **Tủ 1:** Nằm ở cuối Trang 1 ngay sau hàng tiêu đề Table 2.
    * **Tủ 2:** Nằm ở đầu Trang 2 ngay sau hàng tiêu đề được tự động lặp lại.
    * Dưới hàng Tủ 2 trên Trang 2, chèn đúng **10 hàng chấm lửng** tràn lề layout (`.` * 147) lấp đầy khoảng trắng trước khối chữ ký.
  + Cả các hàng tủ đều có thuộc tính chống cắt ngang dòng `<w:cantSplit/>` và thiết lập chiều cao hàng tối thiểu `<w:trHeight w:val="794"/>` (~1.4 cm) tạo khoảng trống rộng rãi cho kỹ thuật viên viết tay.
  + **Quy cách chừa chấm điền tay cho các cột (CŨ, đã được thay bằng mục `[CẬP NHẬT 09/11/2026]` bên dưới):**
    * **Cột `Tên tủ`:** Chừa dòng chấm `\u2026` * 44 (khoảng 4 dòng chấm lửng) để viết tay thông tin tên tủ/vị trí.
    * **Cột `Hãng sản xuất / Serial`:** Chừa sẵn tiêu đề và chấm lửng 4 dòng:
      `Hãng sản xuất: ……………………\n……………………………………………………\nSerial: ………………………………\n………………………………`
    * **Cột `Mã quản lý`:** `…………`
    * **Cột `Ký hiệu ghi chép`:** `………`
  + **[CẬP NHẬT 09/11/2026] Quy cách chừa chấm điền tay cho các cột** (mỗi dòng chấm vừa đúng 1 dòng của cột, không rớt dòng; giãn dòng **1.5**, `space_before/after = 2pt` cho mọi ô của hàng tủ):
    * **Cột `Tên tủ`:** Ghi tên tủ (ví dụ `Tủ an toàn sinh học cấp 2`, `Tủ cấy vi sinh (LAF)`) + **3 dòng chấm** (`…` × 9) bên dưới. Nếu để trắng tên tủ (`--blank_cabinets`) -> **4 dòng chấm**.
    * **Cột `Hãng sản xuất / Serial / Model`:** tiêu đề cột 2 dòng `Hãng sản xuất` / `Serial / Model` (không thêm dòng thứ 3 để hàng tiêu đề không cao thêm). Hàng tủ gồm 3 mục: `Hãng sản xuất:`, `Serial:`, `Model:`, mỗi mục + **1 dòng chấm** (`…` × 12) — hằng `BSC_HSX_DOT_LINES = 1` trong `core/bm02_builder.py`. Đã thử 2 dòng chấm mỗi mục: Tủ 1 rớt sang Trang 2 và hồ sơ 2-3 tủ nhảy sang Trang 3, nên giữ 1 dòng.
    * **Cột `Mã quản lý`:** 2 dòng `…` × 5. **Cột `Ký hiệu ghi chép`:** 2 dòng `…` × 4.
  + **[CẬP NHẬT 09/11/2026] Hàng tiêu đề Table 2 có `keepNext`:** luôn đi liền với hàng tủ đầu tiên. Nếu Trang 1 dài (quy trình 2 dòng, nhiều thiết bị...) làm Tủ 1 không vừa cuối Trang 1 thì cả bảng tự sang Trang 2 — không để hàng tiêu đề trơ trọi ở cuối Trang 1. Vì vị trí bảng phụ thuộc nội dung Trang 1, **KHÔNG ép ngắt trang sau Table 2** (kể cả 1 tủ). Số hàng chấm lửng Trang 2: 1 tủ = 18, 2 tủ = 10, ≥3 tủ = `max(0, 14 - 4*(số tủ - 1))`; đã kiểm tra bằng Word đều đúng 2 trang dù bảng bắt đầu ở Trang 1 hay Trang 2. Các mô tả "1 tủ: ngắt trang sau Table 2, 14 hàng chấm" và "2 tủ: Tủ 1 ở cuối Trang 1" ở trên chỉ còn đúng khi Trang 1 đủ chỗ.
  + **[CẬP NHẬT 09/11/2026 - theo chỉnh tay của người dùng] Bỏ khoảng cách 2pt** (`space_before = space_after = 0`) cho hàng `Quy trình thực hiện`, hàng `Nội dung/ Quá trình thực hiện` (Table 1) và câu dẫn Mục 9 của biên bản BSC (áp dụng SAU `finalize_document_formatting`). Tiết kiệm ~12pt để hàng Tủ 1 vừa cuối Trang 1 kể cả khi quy trình dài 2 dòng (ví dụ hồ sơ BSC + LAF).
  + **Kết quả:** Hồ sơ BSC (1 tủ hay 2 tủ) luôn đạt chuẩn **chính xác 2 trang**, cân đối thẩm mỹ tuyệt đối giữa cả 2 trang giấy.

- **[NEW 11/11/2026] PDF Biên bản in sẵn 2 bản:** File PDF `BM3030.02 - BBLM` luôn chứa **2 bản liên tiếp** (2 trang × 2 = 4 trang) để in một lần, không cần chỉnh số bản khi in. File `.docx` vẫn giữ 1 bản để chỉnh sửa. Nếu 1 bản có số trang lẻ thì chèn trang trắng giữa 2 bản (in 2 mặt bản sau vẫn sang tờ mới). Hằng `BM02_PDF_COPIES = 2` trong `core/pdf_exporter.py`; `bm02_page_count` vẫn là số trang của 1 bản.

## 8. Quy tắc chuẩn hóa Tên khách hàng, Địa chỉ & Viết hoa - viết thường
- **Chuẩn hóa Tên khách hàng / Đơn vị yêu cầu khi bị IN HOA TOÀN BỘ (`normalize_customer_name`):**
  + Nếu tên khách hàng trên phiếu yêu cầu / báo giá bị **IN HOA TOÀN BỘ**, BẮT BUỘC chuyển lại về dạng chữ hoa - chữ thường chuẩn theo quy tắc viết hoa tên cơ quan, tổ chức tiếng Việt (Nghị định 30/2020/NĐ-CP):
    * Viết hoa chữ cái đầu của các từ/cụm từ ghép chỉ loại hình doanh nghiệp và chức năng, ngành nghề hoạt động (`Công ty`, `Tổng công ty`, `Chi nhánh`, `Cổ phần`, `Trách nhiệm hữu hạn`, `Một thành viên`, `Kỹ nghệ`, `Súc sản`, `Thực phẩm`, `Dược phẩm`, `Công nghệ`, `Sinh học`, `Sản xuất`, `Thương mại`, `Dịch vụ`...).
    * Giữ viết hoa các chữ cái đầu của tên riêng / địa danh (`Việt Nam`, `Cát Tân`, `Hulo Care`...) và giữ nguyên **IN HOA** các từ viết tắt / tên thương hiệu viết tắt (`TNHH`, `MTV`, `CP`, `SX`, `TM`, `DV`, `XNK`, `P/S`...).
  + **Ví dụ bắt buộc:**
    * `CÔNG TY CỔ PHẦN VIỆT NAM KỸ NGHỆ SÚC SẢN` -> `Công ty Cổ phần Việt Nam Kỹ nghệ Súc sản`
    * `CÔNG TY CỔ PHẦN P/S` -> `Công ty Cổ phần P/S`
    * `CÔNG TY TNHH THỰC PHẨM CÁT TÂN` -> `Công ty TNHH Thực phẩm Cát Tân`
- **Mở rộng viết tắt địa chỉ thành nguyên chữ:**
  + `KCN` -> `Khu công nghiệp` (ví dụ: `KCN Tân Đông Hiệp B` -> `Khu công nghiệp Tân Đông Hiệp B`, `KCN Vĩnh Lộc` -> `Khu công nghiệp Vĩnh Lộc`).
  + `KCX` -> `Khu chế xuất` (ví dụ: `KCX Tân Thuận` -> `Khu chế xuất Tân Thuận`, `KCX Linh Trung` -> `Khu chế xuất Linh Trung`).
  + `KDC` -> `Khu dân cư` (ví dụ: `KDC Thới An` -> `Khu dân cư Thới An`).
  + `CCN` -> `Cụm công nghiệp` (ví dụ: `CCN Tân Mỹ` -> `Cụm công nghiệp Tân Mỹ`).
  + `KĐT` -> `Khu đô thị` (ví dụ: `KĐT Sala` -> `Khu đô thị Sala`, `KĐT Phú Mỹ Hưng` -> `Khu đô thị Phú Mỹ Hưng`).
  + `KKT` -> `Khu kinh tế` (ví dụ: `KKT Dung Quất` -> `Khu kinh tế Dung Quất`).
  + `KCNC` -> `Khu công nghệ cao`.
  + `Tp.HCM` / `TP.HCM` / `TP. HCM` / `Tp. HCM` -> `Thành phố Hồ Chí Minh`.
  + `QL.` / `QL` -> `Quốc lộ` (ví dụ: `QL. 1A` -> `Quốc lộ 1A`, `QL 13` -> `Quốc lộ 13`).
  + `TL.` / `TL` -> `Tỉnh lộ` (ví dụ: `TL. 8` -> `Tỉnh lộ 8`).
  + `Đ.` / `Đg.` -> `đường`.
  + `P.` -> `phường` (nếu tên chữ: `P. Thới An` -> `phường Thới An`) / `Phường` (nếu tên số: `P. 1` -> `Phường 1`).
  + `Q.` -> `quận` (nếu tên chữ: `Q. Gò Vấp` -> `quận Gò Vấp`) / `Quận` (nếu tên số: `Q. 1` -> `Quận 1`).
  + `H.` -> `huyện` (ví dụ: `H. Hóc Môn` -> `huyện Hóc Môn`).
  + `TX.` -> `thị xã` (ví dụ: `TX. Bến Cát` -> `thị xã Bến Cát`).
  + `TT.` -> `thị trấn` (ví dụ: `TT. Hóc Môn` -> `thị trấn Hóc Môn`).
  + `X.` -> `xã` (ví dụ: `X. Tân Thới Hiệp` -> `xã Tân Thới Hiệp`).
  + `KP.` -> `Khu phố` (ví dụ: `KP. 1` -> `Khu phố 1`).
  + `TDP.` -> `Tổ dân phố` (ví dụ: `TDP. 5` -> `Tổ dân phố 5`).
- **Ví dụ chuẩn:** `"A87, KDC Thới An, Đường Lê Thị Riêng, Khu Phố 1, Phường Thới An, TP.HCM, Việt Nam"` chuẩn hóa thành:
  `"A87, Khu dân cư Thới An, đường Lê Thị Riêng, Khu phố 1, phường Thới An, Thành phố Hồ Chí Minh"`.
- **Bỏ "Việt Nam" ở cuối:** Bỏ hẳn `, Việt Nam`, `, Vietnam`, `, VN` hoặc `Việt Nam` ở cuối chuỗi địa chỉ, không ghi vào hồ sơ.
- **Thành phố trực thuộc Trung ương:** Viết hoa chữ `Thành phố`:
  + `Thành phố Hồ Chí Minh` (chuẩn hóa từ `TP.HCM`, `Tp.HCM`, `TP. HCM`, `TP Hồ Chí Minh`, `thành phố Hồ Chí Minh`...)
  + `Thành phố Hà Nội` (chuẩn hóa từ `TP. Hà Nội`, `Tp. Hà Nội`, `Hà Nội`...)
  + `Thành phố Đà Nẵng` (chuẩn hóa từ `TP. Đà Nẵng`, `Đà Nẵng`...)
  + `Thành phố Hải Phòng` (chuẩn hóa từ `TP. Hải Phòng`, `Hải Phòng`...)
  + `Thành phố Cần Thơ` (chuẩn hóa từ `TP. Cần Thơ`, `Cần Thơ`...)
  + `Thành phố Huế` (nếu có)
- **Tỉnh và Thành phố thuộc tỉnh:** Viết thường chữ `tỉnh` và `thành phố`:
  + Ví dụ tỉnh: `tỉnh Bình Dương`, `tỉnh Đồng Nai`, `tỉnh Long An`, `tỉnh Bà Rịa - Vũng Tàu`...
  + Ví dụ thành phố thuộc tỉnh: `thành phố Thủ Dầu Một`, `thành phố Dĩ An`, `thành phố Thuận An`, `thành phố Biên Hòa`, `thành phố Thủ Đức`...
- **Quận / Huyện / Thị xã:**
  + Tên riêng là chữ -> viết thường: `quận Bình Thạnh`, `quận Tân Bình`, `quận Gò Vấp`, `huyện Hóc Môn`, `huyện Củ Chi`, `thị xã Bến Cát`...
  + Tên riêng là số -> viết hoa từ chỉ đơn vị hành chính: `Quận 1`, `Quận 12`...
- **Phường / Xã / Thị trấn:**
  + Tên riêng là chữ -> viết thường: `phường Thới An`, `phường An Khánh`, `phường Bến Nghé`, `xã Tân Thới Hiệp`, `thị trấn Hóc Môn`...
  + Tên riêng là số -> viết hoa từ chỉ đơn vị hành chính: `Phường 1`, `Phường 14`...
- **Khu phố / Thôn / Ấp / Tổ dân phố:**
  + `Khu phố`: viết hoa chữ `Khu`, viết thường chữ `phố` (ví dụ: `Khu phố 1`, `Khu phố Đông Chiêu`).
  + `Ấp`: viết hoa chữ `Ấp` (`Ấp 1`, `Ấp 2`, `Ấp Bình Mỹ`...).
  + `Thôn`: viết hoa chữ `Thôn` (`Thôn 1`, `Thôn Phú Mỹ`...).
  + `Tổ dân phố`: viết hoa chữ `Tổ`, viết thường `dân phố` (`Tổ dân phố 1`, `Tổ dân phố 5`...).
- **Đường / Phố / Đại lộ / Ngõ / Hẻm / Ngách:**
  + Viết thường danh từ chỉ loại đường giao thông: `đường Lê Thị Riêng`, `đường số 3`, `phố Tràng Tiền`, `đại lộ Mai Chí Thọ`, `hẻm 45`, `ngõ 12`...
- **Khu công nghiệp / Khu chế xuất / Khu dân cư / Cụm công nghiệp:**
  + Mở rộng viết đầy đủ và viết hoa chữ cái đầu: `Khu công nghiệp Tân Bình`, `Khu chế xuất Tân Thuận`, `Khu dân cư Thới An`, `Cụm công nghiệp Tân Mỹ`... (viết hoa chữ `Khu`, `Cụm`, viết thường các từ sau `công nghiệp`, `chế xuất`, `dân cư`).
- **Tra cứu Google / Web Search để xác minh địa danh & cập nhật đơn vị hành chính mới:**
  + **Được phép và khuyến khích tra cứu Google:** Khi thông tin địa chỉ hoặc địa điểm thực hiện trong yêu cầu / báo giá viết tắt, ghép liền cụm danh từ hoặc thiếu từ chỉ cấp hành chính (ví dụ: `"KCN Vĩnh Lộc Hóc Môn, Tp.HCM"`), trợ lý được phép và chủ động tra cứu Google / Web Search để xác định chính xác các cấp đơn vị hành chính và phân tách rõ ràng bằng dấu phẩy.
  + **Ví dụ chuẩn hóa:** `"KCN Vĩnh Lộc Hóc Môn, Tp.HCM"` $\rightarrow$ `"Khu công nghiệp Vĩnh Lộc, xã Hóc Môn, Thành phố Hồ Chí Minh"`.
  + **Lưu ý đơn vị hành chính mới cho tất cả các địa chỉ:** Luôn chủ động cập nhật theo đơn vị hành chính mới nhất (các phường/xã mới sau khi sắp xếp, sáp nhập tại TP.HCM, Hà Nội; các thành phố trực thuộc tỉnh mới thành lập như `thành phố Bến Cát`, `thành phố Tân Uyên`, `thành phố Thủ Đức`...). Tuyệt đối không giữ tên đơn vị hành chính cũ đã bị bãi bỏ hoặc sáp nhập.
- **Quy tắc sử dụng địa chỉ 2 cấp hành chính (BẮT BUỘC áp dụng từ nay về sau):**
  + BẮT BUỘC chỉ sử dụng địa chỉ hành chính 2 cấp gồm: `[Cấp xã/phường/thị trấn], [Cấp tỉnh/thành phố trực thuộc Trung ương]` (ví dụ: `xã Cần Giuộc, tỉnh Long An`, `phường Thới An, Thành phố Hồ Chí Minh`, `xã Hóc Môn, Thành phố Hồ Chí Minh`...).
  + TUYỆT ĐỐI KHÔNG tự động suy diễn chèn thêm cấp quận/huyện/thị xã nếu văn bản khách hàng không có, và KHÔNG viết 3 cấp hành chính kiểu `xã Long Hậu, huyện Cần Giuộc, tỉnh Long An`.
  + Ví dụ: Địa chỉ khách hàng ghi `KCN Long Hậu, Cần Giuộc, Long An` -> chuẩn hóa thành `Khu công nghiệp Long Hậu, xã Cần Giuộc, tỉnh Long An`.

## 9. Quy tắc đặt tên Thư mục lưu file xuất (Folder Output)
- **Cấu trúc chuẩn mực bắt buộc:** `"số hồ sơ - tên khách hàng - nội dung vắn tắc"`
  + Ba thành phần được phân tách bằng dấu cách, gạch nối, dấu cách: ` - `.
  + Tuyệt đối không dùng dấu gạch dưới `_` làm dấu phân cách các thành phần chính (không dùng `Q326D..._CONG TY...`, bắt buộc dùng `Q326D... - [Tên KH] - [Nội dung]`).
- **Ví dụ chuẩn mực:**
  + `Q326D124674 - Cát Tân - Phòng sạch`
  + `Q326D120711 - Greencans Industries - Nước uống`
  + `Q326D117382 - Thụy Khuê - Nuoc sinh hoat`
  + `Q326D124796 - K pack - vskk - swab`
- **Quy ước từ viết tắt cho "Nội dung vắn tắt":**
  + `VSKK`: Vi sinh không khí (hoặc viết thường `vskk`).
  + `VSBM`: Vi sinh bề mặt (hoặc viết thường `vsbm`, `swab`).
  + `vskk - swab`: Khi hồ sơ có cả lấy mẫu vi sinh không khí và vi sinh bề mặt (swab).
  + `Phòng sạch`: Đo kiểm chất lượng phòng sạch (nồng độ hạt bụi, chênh áp, nhiệt ẩm, trao đổi khí...).
  + `Khí nén`: Kiểm tra khí nén (tổng dầu, nhiệt độ điểm sương, nồng độ hạt bụi trong khí nén...).
  + `Khí Nitơ`: Kiểm tra khí nitơ có áp.
  + `Nước uống`: Kiểm nghiệm / lấy mẫu nước uống, nước ăn uống (QCVN 01-1).
  + `Nuoc sinh hoat` (hoặc `Nước sinh hoạt`): Nước sinh hoạt (QCVN 02).
  + `Nước thải`: Nước thải công nghiệp / y tế (QCVN 28, QCVN 40).
  + `MTLĐ`: Quan trắc môi trường lao động (CO2, TVOC, vi khí hậu, tiếng ồn, độ rung, bụi MTLĐ...).
  + `BSC`: Giám định tủ an toàn sinh học cấp I, II (inflow, downflow, HEPA leak test, ánh sáng, ồn, UVC...).
  + `LAF`: Đo kiểm tủ sạch dòng khí tầng.
  + `Tủ hút`: Kiểm tra tủ hút hóa chất.
  + `Kho lạnh`: Mapping nhiệt độ, độ ẩm kho lạnh, tủ bảo quản dược phẩm/vắc-xin.
- **Quy tắc trích xuất "Tên khách hàng" vắn tắt:**
  + Lược bỏ các tiền tố loại hình pháp lý: `CÔNG TY TNHH`, `CÔNG TY CỔ PHẦN`, `CÔNG TY CP`, `CÔNG TY TNHH MTV`, `TỔNG CÔNG TY`, `TẬP ĐOÀN`, `DOANH NGHIỆP TƯ NHÂN`, `DNTN`, `CHI NHÁNH`, `CTY`...
  + Lược bỏ các cụm từ ngành nghề chung không mang tính định danh riêng ở đầu (như `CÔNG NGHỆ SẠCH`, `SẢN XUẤT THƯƠNG MẠI DỊCH VỤ`, `SX TM DV`...) nếu phía sau còn tên thương hiệu riêng rõ ràng (ví dụ: `CÔNG TY TNHH CÔNG NGHỆ SẠCH CÁT TÂN` -> `Cát Tân`).
  + Nếu có tên viết tắt trong ngoặc đơn (ví dụ: `(Vietcombank)`), ưu tiên sử dụng tên viết tắt đó.
  + Viết hoa dạng Title Case (`Cát Tân`, `Greencans Industries`, `Thụy Khuê`, `K pack`).

## 10. Quy tắc đặt tên Cuộc hội thoại (Conversation Title)
- **Cấu trúc chuẩn mực:** `[Mã số hồ sơ] - [Tên khách hàng]` (ví dụ: `Q326D126588 - Hulo Care`, `Q326D120423 - Nhật Khang`).
- **Cập nhật tự động:** Ngay khi xác định hoặc nhận được mã hồ sơ và tên khách hàng trong phiên làm việc, tự động cập nhật tiêu đề cuộc hội thoại vào cơ sở dữ liệu `conversation_summaries.db` (`UPDATE conversation_summaries SET title = ? WHERE conversation_id = ?`) để danh sách lịch sử chat của Antigravity hiển thị chuẩn xác tên hồ sơ và khách hàng, giúp người dùng dễ dàng theo dõi và tra cứu.

## 11. Quy tắc ghi Địa điểm thực hiện (Chừa hàng chấm để điền tay)
- **Không điền cứng tên phòng/khu vực/hệ thống/ống dẫn cụ thể:** Dòng `Địa điểm thực hiện:` trong Table 0 của BM3030.01 và Table 1 của BM3030.02 tuyệt đối không điền sẵn tên cụ thể của vị trí hay thiết bị (như *"Ống dẫn khí nén hệ lọc"*, *"Phòng mổ"*, *"Khu vực cân"*...).
- **Cấu trúc chuẩn mực bắt buộc:**
  ```text
  ......................................................................
  ................................... - [Địa chỉ đã chuẩn hóa]
  ```
- **Mục đích & Quy cách chấm:** Chừa tối đa không gian chấm lửng (khoảng hơn 1 dòng chấm: 1 dòng đầy 70 chấm + 35 chấm trước dấu gạch ngang) để kỹ thuật viên có đủ khoảng trống viết tay chi tiết vị trí cụ thể tại hiện trường (tên phân xưởng, phòng kiểm nghiệm QC, lầu, tòa nhà...) mà không làm chật chội dòng địa chỉ.
- **Ví dụ chuẩn:**
  + Thay vì ghi: `"Địa điểm thực hiện: Ống dẫn khí nén hệ lọc - Lô C20a-3, đường số 14, Khu công nghiệp Hiệp Phước, xã Hiệp Phước, huyện Nhà Bè, Thành phố Hồ Chí Minh"`
  + Bắt buộc ghi:
    `Địa điểm thực hiện : ......................................................................`
    `                     ................................... - Lô C20a-3, đường số 14, Khu công nghiệp Hiệp Phước, xã Hiệp Phước, huyện Nhà Bè, Thành phố Hồ Chí Minh`

## 12. Mã hồ sơ để trắng (chưa có số Phiếu yêu cầu)
- Khi người dùng yêu cầu **"mã hồ sơ để trắng"**: KHÔNG tự sinh mã ngẫu nhiên, KHÔNG dùng lại mã cũ của hồ sơ trước.
- Mã hồ sơ ghi trong biểu mẫu: `Q326D……………` (giữ tiền tố `Q326D`, chừa chấm để KTV điền tay) tại: ô `Số:` header Table 0 (BM01 & BM02), dòng `Phiếu yêu cầu & Nhận TĐ/GĐ:`, dòng `Ký hiệu phiếu ghi chép: Q326D……………/1,`.
- Ngày đăng ký (dòng `Phiếu yêu cầu & Nhận TĐ/GĐ ... Ngày:`) cũng để chấm lửng `…./…./YYYY`.
- Ô mã số trên header Table 0 được nới rộng (lấy bớt từ cột trống kế bên, tổng bề rộng bảng giữ nguyên) để placeholder nằm gọn 1 dòng.
- Tên thư mục / file dùng dạng ASCII an toàn: `Q326D...... - [Tên KH] - [Nội dung]` (ví dụ `Q326D...... - Nextern Việt Nam - Khí nén - Phòng sạch`).
- Cách gọi: `python generate_now.py ... --blank_doc` hoặc `process_order({..., "blank_doc_number": True})` hoặc truyền `doc_number="BLANK"` / `"để trắng"` / `"..."`.

## 13. Tự động sao chép 1 bản sang Ổ đĩa Z (Lưu trữ hồ sơ QUATEST 3)
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

---

## 14. HỒ SƠ MÔI TRƯỜNG LAO ĐỘNG (MTLĐ) THEO KẾ HOẠCH TỪNG PHÒNG — DÙNG CHUNG (20/11/2026)
- **Công cụ:** `generate_mtld.py` (module `core/mtld_builder.py`). Áp dụng cho mọi hồ sơ quan trắc MTLĐ có nhiều phòng/khu vực (khởi nguồn từ hồ sơ QUATEST 3 – KTN Cát Lái 2026). Hồ sơ chỉ đo CO2/TVOC văn phòng vẫn dùng profile `co2_mtld` cũ.
- **Đầu vào:** file Excel kế hoạch theo mẫu `templates/MTLD_KeHoach_mau.xlsx`: cột `Phòng | Chỉ tiêu | SL | Điểm đo` (Điểm đo dạng `P304: 02; P303: 01`). Điểm đo ghi không rõ số -> cả cụm thành 1 khu vực; sửa bằng `--fix "Phòng|Chỉ tiêu=Khu vực:SL;Khu vực:SL || ..."`.
- **BM3030.01 Ghi chép:** trang 1 thông tin + bảng theo dõi khối lượng theo phòng (cột Xong ☐) + bảng tra mã OWAS / QĐ 3733 (nếu có Ecgonomy). **Mỗi phòng 1 trang** gồm:
  + Bảng `Máy đo sử dụng` tick ☐ (chỉ hiện chỉ tiêu phòng đó có): Nhiệt ẩm `Testo 625 / Testo 440 / Climomaster 6533-2G`; Gió `Climomaster 6533-2G / 6501`; Ánh sáng `Testo 440 – máy 1 / Testo 440 – máy 2`; Ồn `Rion NL-53`; Bụi `Metone Aerocet 831`; CO₂ `testo 440`; luôn có `☐ Khác: ……`.
  + Bảng đo trực tiếp: mỗi khu vực ghi đủ tên ở từng dòng (KHÔNG dùng dấu 〃), **ô xám đậm (A6A6A6) = không đo chỉ tiêu đó tại điểm đó**. Cột Ồn ghi **tên file đo** (không ghi dBA); ồn dải tần chỉ 1 cột tên file đo.
  + Bảng lấy mẫu hơi khí độc / vi sinh: ghi sẵn tên chỉ tiêu; cột `Bơm` tick `☐ Buck ☐ SKC ☐ Kitagawa` (Buck Genie, SKC PCXR8, Kitagawa ASP 1200 – Japan); vi sinh `☐ Bio Culture`.
  + Bảng Ecgonomy (OWAS: Lưng/Tay/Chân/Vật nặng + vị trí lao động).
  + Chữ ký `Người thực hiện` 1 lần ở cuối.
- **BM3030.02 Biên bản:** mỗi chỉ tiêu 1 dòng, cột vị trí ghi sẵn các phòng thực hiện (`Phòng QT, DH...` / `Các phòng`); Số lượng, Kí hiệu, Niêm phong để chấm. Bảng ghi chú cuối giữ đường chấm sẵn của mẫu — KHÔNG gõ thêm dải dấu chấm (bị chồng 2 lớp).
- **Thiết bị sử dụng & Quy trình** tự lọc theo chỉ tiêu có trong kế hoạch (vi khí hậu/ánh sáng N2.TC.01, ồn N2.TC.02, bụi N2.TC.06, CO2 + hơi khí độc N2.TC.07, vi sinh N2.TC.09, Ecgonomy N2.TC.11, N2.TC.12).
- `--multi_day` khi đo nhiều ngày: `Bắt đầu từ ngày …./…./2026 kết thúc ngày …./…./2026`.
- Chỉ tiêu `Đánh giá yếu tố tiếp xúc nghề nghiệp` -> bảng riêng (Người LĐ, Yếu tố tiếp xúc, Thời gian tiếp xúc, PTBVCN). Cột `Ghi chú` của kế hoạch (vd `Mẫu gửi`) được ghi sẵn vào cột Ghi chú bảng lấy mẫu. Hồ sơ chưa chia phòng: Phòng = `Khu vực: ……`, Điểm đo = `……` -> ô khu vực để trống viết tay, biên bản ghi dải chấm.

---

## NHẬT KÝ THAY ĐỔI CODE (để đồng bộ giữa Claude Cowork ↔ Antigravity)
- 20/11/2026: thêm `core/mtld_builder.py` + `generate_mtld.py` + `templates/MTLD_KeHoach_mau.xlsx` (hồ sơ MTLĐ theo kế hoạch từng phòng, quy tắc 14).

### 09/11/2026 — Claude Code: BSC chừa chỗ viết tay rộng hơn, bỏ N2.25ĐG khỏi form ghi chép BSC
1. **`core/bm02_builder.py`** (nhánh BSC của Table 2, đánh dấu `[NEW 09/11/2026]`):
   - Hàng tủ: tên tủ + 3 dòng chấm (để trắng -> 4 dòng), `Hãng sản xuất:`/`Serial:` mỗi mục 2 dòng chấm, `Mã quản lý`/`Ký hiệu ghi chép` 2 dòng chấm; giãn dòng 1.5, spacing 2pt. Bỏ việc tự thay tên chứa "tủ an toàn sinh học" bằng dòng chấm (tên tủ được giữ, có dòng chấm bên dưới).
   - Hàng tiêu đề Table 2 thêm `keepNext` (không để tiêu đề trơ trọi cuối Trang 1).
   - Bỏ ngắt trang ép sau Table 2 cho 1 tủ; 1 tủ = 18 hàng chấm Trang 2. Bỏ đoạn `[OPT 05/11/2026]` ép giãn dòng 1.25 cho hàng tủ (vẫn giữ phần nén khối thiết bị).
   - Đã kiểm tra bằng Word (6 trường hợp: 1 tủ 8 thiết bị, 1 tủ 3 thiết bị, 1 tủ Trang 1 dài, 2 tủ, 3 tủ, 2 tủ để trắng tên): tất cả đúng 2 trang.
1a. **Tủ 1 lên Trang 1** (theo bản người dùng tự chỉnh trong Word): `core/bm02_builder.py` đặt spacing 0pt trên/dưới cho hàng `Quy trình thực hiện`, `Nội dung/ Quá trình thực hiện` và câu dẫn Mục 9 (chỉ BSC). Kiểm tra lại 6 trường hợp bằng Word: đều đúng 2 trang; trường hợp 2 tủ BSC + LAF và 1 tủ có Trang 1 dài nay đều có Tủ 1 ở cuối Trang 1. Biên bản Sanofi dùng nguyên bản người dùng đã lưu.
1d. **Ổ Z chia thư mục con `Thuc hien` / `Scan`**: `core/form_cloner.py` (`copy_order_to_z_drive(..., scan_files=)` chép file sinh ra vào `Thuc hien`, file gốc vào `Scan`, trả về đường dẫn `Thuc hien` để `pdf_exporter` đồng bộ PDF đúng chỗ; tham số `scan_files` trong `clone_form_set`/`process_order`), `generate_now.py` (cờ `--scan`). Đã chuyển 2 thư mục Sanofi `Q326D139329` trên Z sang cấu trúc mới.
1c. **Thêm mục Model** vào ô `Hãng sản xuất / Serial` của biên bản BSC (tiêu đề `Hãng sản xuất` / `Serial / Model`; hàng tủ: Hãng sản xuất, Serial, Model mỗi mục 1 dòng chấm). Kiểm tra lại 6 trường hợp bằng Word: đều 2 trang, Tủ 1 ở Trang 1.
1b. **Mục 9 dùng số PO khi không có QONV**: `core/bm02_builder.py` (câu Mục 9), `core/form_cloner.py` (tham số `po_number`, khóa `po_number`/`po`/`so_po` trong `process_order`), `generate_now.py` (cờ `--po`). Hồ sơ Sanofi `Q326D139329` dùng PO `E900192188`.
2. **`BSC/Ghi Chep BSC.pdf`**: xóa chữ `N2.25ĐG` ở ô `Số hồ sơ` (trang 1), giữ gạch chân để điền tay. Bản gốc lưu tại `BSC/_backup/Ghi Chep BSC - goc (co N2.25DG).pdf`. Đã xóa cache ảnh `scratch/bsc_imgs_300` để file Word nhúng ảnh mới.
3. Hồ sơ `Q326D139329` - Công ty Cổ phần Sanofi Việt Nam: tách 2 bộ `BSC - LAF` (2 tủ) và `Nước uống` (theo số mẫu năm 2025 hồ sơ 966N2.25DG: 7 nước uống + 3 nước cấp có đo pH, clo dư tại hiện trường; hệ thống chưa có profile nước nên bộ nước được ghép bằng script riêng từ mẫu phòng sạch BM02 + bảng "Lấy mẫu nước" của BM3030.01 gốc).

### 05/11/2026 — Antigravity: Tối ưu Bố cục BSC 1 tủ (>= 6 thiết bị) & Chống phát sinh trang trắng thừa trước Ngắt trang
- **Mục đích:** Khắc phục trường hợp hồ sơ BSC có danh mục thiết bị dài (>= 6 thiết bị như inflow, downflow, HEPA leak, khói, độ rọi, độ ồn, UVC) làm tràn bảng hoặc đẩy ngắt trang sang Trang 3.
- **Các file đã cập nhật:**
  1. `core/docx_processor.py` (`ensure_6pt_spacing_after_all_tables`): Thêm điều kiện nhận diện nếu phần tử tiếp theo (hoặc sau các paragraph rỗng) mang thuộc tính `<w:pageBreakBefore/>` thì xóa sạch paragraph rỗng thừa và KHÔNG chèn thêm 6pt spacer paragraph trước ngắt trang, tránh lỗi Word render spacer thành 1 trang trắng riêng biệt.
  2. `core/bm02_builder.py`: 
     - Dọn dẹp triệt để các paragraph rỗng thừa giữa Table 2 và đoạn ngắt trang/chữ ký trên Trang 2.
     - Tinh chỉnh khoảng cách dòng (spacing) cho khối thiết bị trong Table 1 (`space_before=1pt`, `space_after=1pt`, `line_spacing=1.05`) và hàng Table 2 (`line_spacing=1.25`) khi có >= 6 thiết bị, bảo đảm Table 2 (1 tủ) luôn nằm trọn vẹn ở cuối Trang 1 và tài liệu đạt đúng chuẩn **CHÍNH XÁC 2 TRANG HOÀN HẢO**.
  3. Xử lý hoàn tất 02 hồ sơ từ thư mục `Input/`:
     - `Q326D136247` - Công ty Cổ phần Phân bón Bình Điền (01 tủ ESCO SC2-4E1, Báo giá QONV26015972, ngày nhận 26/10/2026).
     - `Q326D136839` - Chi nhánh Công ty TNHH Nước giải khát Swire Coca-Cola Việt Nam tại tỉnh Tây Ninh (01 tủ Bio Air SafeMate ECO+1.5, Báo giá QONV26015971, ngày nhận 27/10/2026).

### 04/11/2026 — Antigravity: Tự động sao chép 1 bản hồ sơ sang Ổ đĩa Z (Quy tắc 13)
- **Mục đích:** Tự động đồng bộ và sao chép 1 bản lưu trữ sang thư mục mạng hồ sơ môi trường năm 2026 của QUATEST 3 theo yêu cầu người dùng (`Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026`).
- **Các file đã cập nhật:**
  1. `core/config.py`: Khai báo hằng số `Z_DRIVE_DIR = os.environ.get("QUATEST3_Z_DRIVE", r"Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026")`.
  2. `core/form_cloner.py`: Bổ sung hàm `copy_order_to_z_drive(target_dir, subfolder_name, z_base_dir)` và tích hợp vào Bước 8 của `clone_form_set()`, `process_order()`. Bỏ qua tệp khóa `~$*` của Word khi sao chép.
  3. `core/pdf_exporter.py`: Bổ sung cơ chế tự động đồng bộ file PDF sang ổ Z khi xuất PDF.
  4. `core/__init__.py`: Export `Z_DRIVE_DIR` và `copy_order_to_z_drive`.
  5. `generate_now.py`, `app_cli.py`, `app_gui.py`: Bổ sung tham số tùy chọn `--no_z_copy`, `--z_drive_dir` và hiển thị đường dẫn sao lưu ổ Z khi tạo thành công.

### 27/10/2026 — Claude (Cowork): Hồ sơ NEXTERN Việt Nam (Khí nén 2 vị trí + Phòng sạch, mã để trắng)
Các file đã sửa (đã ghi đè trực tiếp vào thư mục dự án, các đoạn sửa có đánh dấu `[FIX 27/10/2026]` / `[NEW 27/10/2026]`):

1. **`core/normalizer.py`**
   - [FIX] Regex chuẩn hóa `Đường` (mục 14): bản cũ `\b(?:Đg\.?|Đ\.?|Đường)` khớp nhầm chữ `Đ` đầu từ → `"Đường DE5"` thành `"đường ường DE5"`, `"Đại An"` thành `"đường ại An"`, `"Đồng Nai"` thành `"đường ồng Nai"`. Sửa thành `\b(?:Đường|Đg\.|Đ\.)` (viết tắt bắt buộc có dấu chấm).
   - [NEW] Hằng số `DOC_NUMBER_BLANK_PLACEHOLDER = "Q326D……………"`, `DOC_NUMBER_BLANK_FILENAME = "Q326D......"`, hàm `is_blank_doc_number()`, `doc_number_for_filename()`.
   - `build_output_folder_name()` dùng `doc_number_for_filename()` cho thành phần mã số.
2. **`core/config.py` → `resolve_equipment_and_procedure()`**
   - [FIX] Thêm nhánh `category_key == "khi_nen_phong_sach"`. Bản cũ rơi vào fallback phòng sạch → in ra thiết bị Testo 625, TSI 8375, Rion, Buck (không được yêu cầu) và THIẾU thiết bị khí nén; quy trình sai `N2.TC.09, N2.TC.PS.01`.
   - Nhánh mới tách chỉ tiêu theo chủ thể (khí nén / phòng sạch) nên `Đo nhiệt độ điểm sương trong khí nén` không còn bị hiểu nhầm là chỉ tiêu nhiệt độ - độ ẩm phòng sạch. Dấu câu chuẩn: `;` cuối mỗi dòng, `.` ở dòng cuối. Quy trình: `QUATEST 3 3030:2025; N2.TC.KN.01; N2.TC.PS.01`.
3. **`core/bm01_builder.py` → `build_bm01_hybrid_khi_nen_phong_sach()`**
   - [FIX] Khi xóa Table 3, 4, 5 của mẫu phòng sạch thì xóa luôn đoạn tiêu đề ngay trước bảng (bản cũ sót tiêu đề `Đo chênh áp phòng sạch`).
   - [FIX] Bố cục trang: hàng chấm lửng nay nằm ở **Trang 1** (bản cũ chèn sau đoạn ngắt trang nên Trang 1 trống nửa dưới, hàng chấm rơi sang Trang 2). Số hàng chấm = `max(6, 16 - max(0, số_thiết_bị - 4))`.
   - [FIX] Mỗi vị trí khí nén bắt đầu trang mới (tiêu đề `Khí nén (Vị trí 0x):` có `pageBreakBefore`), bảng bụi phòng sạch bắt đầu trang mới (đoạn ngắt trang mảnh). Kết quả 2 vị trí: 4 trang = Trang 1 thông tin | Trang 2 KN vị trí 01 | Trang 3 KN vị trí 02 | Trang 4 bụi phòng sạch. Hết lỗi `Người thực hiện` rớt lẻ sang đầu trang sau.
4. **`core/form_cloner.py`**
   - [NEW] Tham số `blank_doc_number` cho `clone_form_set()`; `process_order()` đọc khóa `blank_doc_number`.
   - [NEW] Hàm `widen_blank_doc_number_cell(docx_path, extra_dxa=700)` chạy sau khi sinh BM01/BM02 khi để trắng mã.
   - Tên file BM01/BM02 dùng `doc_number_for_filename()`.
5. **`core/order_profiles.py`**: sửa dấu câu `default_equipment` của profile `khi_nen_phong_sach` (dòng tổng dầu `.` → `;`).
6. **`generate_now.py`**: thêm cờ `--blank_doc` và `--num_samples_bm01` (số dòng đo bụi phòng sạch trong BM01, mặc định theo số chỉ tiêu nếu không truyền → nên truyền 8).

Lệnh đã dùng để sinh hồ sơ NEXTERN:
```powershell
python generate_now.py --profile khi_nen_phong_sach --blank_doc --num_positions 2 --num_samples_bm01 8 `
  --customer "CÔNG TY TNHH NEXTERN VIỆT NAM" `
  --address "Nhà xưởng C1B-D5, Lô C-1B-CN, Đường DE5, Khu Công Nghiệp Mỹ Phước 3, Phường Thới Hòa, Thành phố Hồ Chí Minh, Việt Nam" `
  --request "Lấy mẫu, phân tích chất lượng khí nén và giám định chất lượng phòng sạch." `
  --indicators "Lấy mẫu tổng dầu trong khí nén;Đo nhiệt độ điểm sương trong khí nén;Đo nồng độ bụi trong khí nén;Đo nồng độ bụi phòng sạch" `
  --short_customer "Nextern Việt Nam" --short_content "Khí nén - Phòng sạch" --pdf
```
Lưu ý: Thư mục cũ `output/Q326D133266 - Nextern...` được sinh với địa chỉ SAI (lấy từ ví dụ README: "KCN Đại An Mở Rộng, Hải Dương") và thiết bị SAI — không dùng lại. Địa chỉ đúng theo Phiếu yêu cầu (Input/133266.pdf): `Nhà xưởng C1B-D5, Lô C-1B-CN, đường DE5, Khu công nghiệp Mỹ Phước 3, phường Thới Hòa, Thành phố Hồ Chí Minh`.

### 28/10/2026 — Claude (Cowork): `Người thực hiện` chỉ ký 1 lần ở cuối BM3030.01
File đã sửa: **`core/bm01_builder.py`** (đánh dấu `[FIX 28/10/2026]` / `[NEW 28/10/2026]`):
- [FIX] `build_bm01_hybrid_khi_nen_phong_sach()`: bỏ đoạn chèn `Người thực hiện` dưới mỗi bảng `Khí nén (Vị trí 0x)` (bản cũ in chữ ký ở Trang 2, Trang 3 và cuối tài liệu).
- [NEW] Hàm `ensure_single_signature_at_end(doc)`: xóa mọi đoạn `Người thực hiện` ở cấp body, giữ đúng 1 đoạn (định dạng của mẫu gốc nếu có, nếu không tạo mới: căn phải, đậm, 12pt, thụt phải 1133 dxa) đặt ở cuối tài liệu ngay trước `sectPr`; đồng thời đưa `sectPr` về cuối body (bảng thứ cấp phòng sạch trước đây bị append sau `sectPr`).
- Gọi hàm này sau `finalize_document_formatting(doc)` trong `build_bm01_khi_nen()`, `build_bm01_phong_sach()`, `build_bm01_hybrid_khi_nen_phong_sach()` (không áp dụng cho BSC). Lưu ý: mẫu Khí nén đơn thuần (`Khi nen/BM3030.01- Ghi chép - Q326D122266.docx`) vốn không có chữ ký → nay được bổ sung 1 chữ ký ở cuối.
- Đã kiểm tra: hybrid 2 vị trí, khí nén 3 vị trí, phòng sạch + tốc độ gió → mỗi file đúng 1 `Người thực hiện`, nằm ở trang cuối.

### 30/10/2026 — Antigravity: Hồ sơ P/S (`Q326D131671`) & Việt Nam Kỹ nghệ Súc sản (`Q326D134804`)
Các file đã sửa:
1. **`core/normalizer.py` (cùng `core/form_cloner.py`, `core/docx_processor.py`, `core/bm01_builder.py`)**:
   - [NEW] Bổ sung `normalize_customer_name()` tự động chuẩn hóa tên khách hàng bị IN HOA TOÀN BỘ sang chữ hoa - thường chuẩn tên tổ chức tiếng Việt (ví dụ: `CÔNG TY CỔ PHẦN VIỆT NAM KỸ NGHỆ SÚC SẢN` → `Công ty Cổ phần Việt Nam Kỹ nghệ Súc sản`, `CÔNG TY CỔ PHẦN P/S` → `Công ty Cổ phần P/S`).
   - Giữ nguyên chữ viết tắt có dấu gạch chéo như `P/S` trong `to_title_case_vn()`, bổ sung `sanitize_for_filename()` chuyển `/` và `\` thành `-` (`P-S`) khi tạo tên thư mục/file mà không dùng dấu `_`.
2. **`core/config.py`**: nhánh `bsc` tự động nhận diện `cấp 1` / `cấp 2` để xuất đúng quy trình `N2.TC.BSC.01` và/hoặc `N2.TC.BSC.02`, loại bỏ từ khóa quá rộng `"dòng khí"` ở `has_smoke`, chuẩn hóa dấu `;` và `.` cuối danh sách thiết bị.
3. **`core/bm02_builder.py` & `core/docx_processor.py`**:
   - [FIX] Sửa lỗi mục `5.5` nhận nhầm dòng `"Chỉ tiêu thực hiện lấy mẫu vi sinh trong không khí: ..."` là dòng chấm lửng và ghi đè thành `'.' * 75`; nay chỉ khớp đoạn thuần dấu chấm (`re.fullmatch(r'[\s.…]+', txt)`) và tự động chèn đủ số hàng chấm tràn lề `'.' * 147` trên Trang 2.
   - [NEW] Hỗ trợ nhiều tủ BSC (`cabinets` / `num_positions`) trên Table 2 của BM3030.02, tinh chỉnh độ dài chấm lửng các cột `Hãng sản xuất / Serial`, `Mã quản lý`, `Ký hiệu ghi chép` và cột `Số lượng` (`829 dxa`) nằm gọn 1 dòng.
4. **`core/bm01_builder.py`, `core/form_cloner.py`, `core/pdf_exporter.py`, `generate_now.py`**:
   - [NEW] Bảo toàn bảng ghi chép `Vi sinh` gốc (`Table 5` của mẫu Phòng sạch với đầy đủ checkbox Wingdings `APBuck`, `SKC-Biostage`, `TS lấy mẫu / TS blank`, `MM lấy mẫu / MM blank` và `numId='36'`) thông qua `prepare_original_vskk_table_element()`, hỗ trợ `--num_vskk_bm01` độc lập với `--num_samples_bm01`, lấp đầy chấm lửng Trang 1 khi bảng đo sang Trang 2.
   - [NEW] Hỗ trợ nối `Ghi Chep BSC.pdf` nhân theo số lượng tủ (`num_positions`, ví dụ 2 tủ → 1 + 3×2 = 7 trang) cho cả `.docx` và `.pdf` vector, bổ sung cờ `--cabinets`, `--num_vskk_bm01`, `--micro_targets`.

### 01/11/2026 — Antigravity: Chuẩn hóa Mục 9 "Nội dung/ Quá trình thực hiện" cho BM3030.02
Các file đã sửa:
1. **`core/bm02_builder.py`**:
   - [NEW] Tự động chèn câu ghi sẵn chuẩn mực phía dưới Mục 9 ("9. Nội dung/ Quá trình thực hiện :"):
     `"Nội dung thực hiện theo báo giá số {quotation_number}. Vị trí thực hiện được chỉ định bởi đại diện {customer_name}, số lượng vị trí cụ thể như sau:"`
     với `quotation_number` (mặc định `QONV2......` hoặc số báo giá cụ thể được truyền vào) và `customer_name` đã qua hàm chuẩn hóa `normalize_customer_name` (ví dụ: `Công ty Cổ phần P/S`).
   - [NEW] Giảm tương ứng 2 hàng chấm lửng ở Trang 1 (`num_dot_lines_p1`) để bù trừ cho 2 dòng hiển thị của câu trên, bảo đảm Trang 1 vừa vặn 100%.
   - [FIX] Tinh chỉnh `target_p2_dots` (với 3 chỉ tiêu: `target_p2_dots = 1`) đảm bảo 100% các loại hồ sơ (Phòng sạch, Khí nén, BSC, Hybrid...) luôn đạt chuẩn 2 trang không bao giờ bị nhảy sang trang 3.
2. **`core/form_cloner.py` & `generate_now.py`**:
   - [NEW] Bổ sung tham số `quotation_number` / `quote_number` và cờ CLI `--quote` / `--quotation_number` để người dùng hoặc AI có thể truyền số báo giá thực tế vào hồ sơ khi cần.
3. **`GEMINI.md` & `QUY_TAC_SOAN_THAO_BIEU_MAU.md`**: Cập nhật quy tắc bắt buộc áp dụng từ nay về sau.

### 02/11/2026 — Antigravity: Chuẩn hóa Bố cục BM3030.02 cho Tủ an toàn sinh học (BSC 2 tủ / chừa chấm tên tủ)
Các file đã sửa:
1. **`core/bm02_builder.py`**:
   - [NEW] Thêm thuộc tính `<w:tblHeader/>` cho hàng tiêu đề Table 2 của BSC để tự động lặp lại ở đầu trang sau.
   - [NEW] Hỗ trợ bố cục 2 tủ BSC: Không chèn ngắt trang trước Table 2, để Tủ 1 nằm ở cuối Trang 1 và Tủ 2 nằm ở đầu Trang 2; thiết lập `<w:trHeight w:val="794"/>` và `<w:cantSplit/>` cho cả hai hàng tủ.
   - [NEW] Chừa chấm lửng cột `Tên tủ` (`\u2026` * 44, tương đương 4 dòng chấm), cột `Hãng sản xuất / Serial` (4 dòng chấm kèm nhãn `Hãng sản xuất:` và `Serial:`).
   - [NEW] Điều chỉnh số hàng chấm lửng Trang 2 (`target_p2_dots = 10`) khi có 2 tủ để lấp đầy khoảng trống hoàn hảo trước khối chữ ký, bảo đảm tài liệu đúng 2 trang.
2. **`core/form_cloner.py` & `generate_now.py`**:
   - [NEW] Thêm tham số `blank_cabinets` / cờ CLI `--blank_cabinets` để hỗ trợ chừa trống tên tủ cho kỹ thuật viên tự điền tay tại hiện trường.
3. **`GEMINI.md` & `QUY_TAC_SOAN_THAO_BIEU_MAU.md`**: Cập nhật quy tắc bố cục 2 tủ BSC và nhật ký thay đổi.

### 02/11/2026 (Chiều) — Antigravity: Tối ưu Bố cục BSC 1 tủ trên Trang 1 & Tăng khoảng trống viết tay Địa điểm thực hiện
Các file đã sửa:
1. **`core/normalizer.py` & `core/docx_processor.py`**:
   - [NEW] Tăng không gian viết tay cho dòng `Địa điểm thực hiện` với tiền tố hơn 1 dòng chấm lửng (`.` * 70 + `\n` + `.` * 35 + ` - [Địa chỉ]`), giúp KTV có đủ diện tích viết tay tên phân xưởng, phòng kiểm nghiệm mà không bị chật chội dòng địa chỉ.
2. **`core/bm02_builder.py`**:
   - [NEW] Cho phép Table 2 của hồ sơ BSC (cả 1 tủ và 2 tủ) bắt đầu ngay dưới Mục 9 trên Trang 1, loại bỏ hoàn toàn khoảng trắng thừa ở cuối Trang 1.
   - [NEW] Với 1 tủ BSC: Table 2 nằm trọn vẹn ở cuối Trang 1; chèn ngắt trang ngay sau Table 2 để Trang 2 có 14 hàng chấm lửng ghi chép và khối chữ ký cao ráo, thoáng đãng.
3. **`core/bm01_builder.py`**:
   - [NEW] Cập nhật `_calc_bm01_p1_dot_lines` giảm 1 hàng chấm lửng bù trừ cho dòng chấm bổ sung của Địa điểm thực hiện, bảo đảm Trang 1 BM3030.01 luôn vừa vặn 100%.

### 11/11/2026 — Claude Code: Quét rò rỉ màng lọc HEPA phòng sạch & Thành phố Đồng Nai (hồ sơ Q326D067460 - Lâm Việt)
Các file đã sửa:
1. **`core/config.py`**:
   - [NEW] Profile/lĩnh vực `phong_sach` nhận diện chỉ tiêu `hepa` / `rò rỉ` -> thêm thiết bị `Máy tạo sol khí ATI aerosol generator, model 6D và quang kế aerosol photometer, model 2i, USA;`.
   - [FIX] `infer_method_type`: `quét rò rỉ` / `rò rỉ` là phương pháp **Đo** (trước đây rơi vào từ khóa `màng lọc` -> bị đánh dấu nhầm là Lấy mẫu).
2. **`core/bm01_builder.py`**:
   - [NEW] Bảng thứ cấp `hepa_leak` (`Quét rò rỉ màng lọc HEPA`, cột `Độ xuyên thấu (%)`). Khi bảng đo bụi ≥ 8 dòng: ngắt trang, bảng HEPA nằm trọn Trang 3 với 15 dòng.
3. **`core/normalizer.py` & `core/docx_processor.py`**:
   - [NEW] Thêm `Đồng Nai` vào danh sách thành phố trực thuộc Trung ương -> giữ viết hoa `Thành phố Đồng Nai` (theo cách ghi trên hồ sơ thực tế Q326D084627).

### 11/11/2026 (Chiều) — Claude Code: Khí nén chỉ đo tổng dầu (Q326D139353 Phenikaa, Q326D139348 PSI Việt Nam, Q326D139363 Quí Long)
1. **`core/config.py`**: [FIX] Profile `khi_nen` dùng mẫu BBLM phòng sạch (`Phong sach, Vi sinh/BM3030.02- BBLM - Q326D108033.docx`). Mẫu khí nén cũ có đoạn viền chấm làm biên bản 1 chỉ tiêu nhảy sang 3 trang; mẫu mới cho đúng 2 trang (đã thử 1 chỉ tiêu × 1 vị trí và 3 chỉ tiêu × 2 vị trí).
2. **`core/bm01_builder.py` → `build_bm01_khi_nen()`**: [NEW] Chỉ giữ khối chỉ tiêu được yêu cầu trong bảng khí nén (`Tổng dầu` / `Particle` / `Điểm sương`, nhận theo từ khóa của chỉ tiêu). Nếu không nhận ra chỉ tiêu nào thì giữ nguyên như cũ.
3. Scan gộp `Input/20261111154557244.pdf` (30 trang) đã tách thành 3 file 10 trang trong `Input/` (`<mã> - <KH> - scan.pdf`). Trang 18 gốc (báo giá PSI 2/3) bị ngược nên đã xoay 180°.
4. Lưu ý: Phenikaa có Phiếu kiểm soát in mã Q326D139352 nhưng Phiếu yêu cầu ghi tay Q326D139353; người dùng chọn **Q326D139353**. Mục 9 dùng số của báo giá kèm theo (QONV26016982 / QONV26016980), không dùng QONV26016728 ở dòng 11 của Phiếu yêu cầu.
5. **Dụng cụ lấy mẫu tổng dầu** (theo yêu cầu người dùng, áp dụng từ nay): thay `Dụng cụ lấy mẫu tổng dầu (Regulator, Charcoal tube, Cassette).` bằng 2 dòng `- Bộ dụng cụ lấy mẫu tổng dầu trong khí nén;` và `- Giấy lọc thủy tinh Pall life Sciences, ống than hoạt tính ORBO®.` trong `core/config.py` (SOP N2.TC.KN.01, `standard_equipment` khi_nen, nhánh `khi_nen` và `khi_nen_phong_sach` của `resolve_equipment_and_procedure`) và `core/order_profiles.py` (profile `khi_nen`, `khi_nen_phong_sach`). Đã sinh lại 3 hồ sơ Q326D139353 / Q326D139348 / Q326D139363.
6. **PDF Biên bản 2 bản**: `core/pdf_exporter.py` (`repeat_pdf_copies`, `BM02_PDF_COPIES = 2`) lặp PDF BM3030.02 thành 2 bản; `generate_now.py` in thêm số bản. Đã sinh lại 3 hồ sơ trên (PDF biên bản 4 trang, ở cả output và ổ Z).

### 14/11/2026 — Claude Code: Phòng sạch đặt đĩa / đĩa tiếp xúc / chênh áp màng lọc HEPA (hồ sơ Q326D125256 - Dược Sài Gòn)
1. **`core/config.py`** (nhánh `phong_sach` của `resolve_equipment_and_procedure`):
   - [FIX] `hepa` đơn thuần không còn kéo theo máy ATI 6D/2i; chỉ `rò rỉ` / `leak` / `xuyên thấu` mới là quét rò rỉ. `Đo chênh áp màng lọc HEPA` -> `Thiết bị đo chênh lệch áp suất TSI 8375, USA;`.
   - [FIX] `vận tốc gió` / `tốc độ gió` -> `Máy đo tốc độ gió Kanomax Climomaster 6501, Japan;` (trước đây ra nhầm `Thiết bị đo lưu lượng TSI 8375`); `lưu lượng` / `trao đổi khí` vẫn là TSI 8375.
   - [NEW] Vi sinh không khí phương pháp đặt đĩa (`đặt đĩa`, `đĩa lắng`) -> `Khay inox và đĩa môi trường lấy mẫu vi sinh không khí (phương pháp đặt đĩa);` thay cho máy Buck. Vi sinh bề mặt `đĩa tiếp xúc` / `rodac` -> `Đĩa tiếp xúc Rodac lấy mẫu vi sinh bề mặt.`
   - [FIX] Nhận thêm cụm `vi sinh trong không khí` cho vi sinh không khí (trước đây bị bỏ sót khi hồ sơ có cả vi sinh bề mặt).
2. **`core/bm01_builder.py`**:
   - [NEW] Bảng thứ cấp `vskk_settle` (cột `Thời gian đặt đĩa`, dùng thay bảng Vi sinh gốc máy Buck khi đặt đĩa), `contact_plate` (`Diện tích lấy mẫu ~ 25 cm²`), `filter_dp` (`Chênh áp màng lọc HEPA`, cột `Chênh áp (Pa)`, mặc định 4 dòng).
   - [NEW] Bảng thứ cấp có `keepNext` cho mọi hàng trừ hàng cuối -> bảng không bị cắt qua 2 trang.
   - [NEW] Số dòng từng bảng thứ cấp: cờ `--bm01_rows "air_velocity:3;filter_dp:3"` (`generate_now.py`), tham số `bm01_rows` (`clone_form_set`) / khóa `bm01_rows` (`process_order`).
3. **`core/docx_processor.py`**: hồ sơ có cả vi sinh không khí và bề mặt -> `Chỉ tiêu thực hiện lấy mẫu vi sinh trong không khí và vi sinh bề mặt: ...`.
4. Tên file scan gốc mới: `<Số hồ sơ> - <Tên KH> - Scan PYC.pdf` (Quy tắc 13).
