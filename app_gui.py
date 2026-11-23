"""
Giao diện đồ họa (GUI Desktop) Hệ thống Nhân bản và Tái chế Biểu mẫu Hiện trường QUATEST 3
Hỗ trợ tạo hồ sơ đơn lẻ và nhân bản hàng loạt từ file Excel.
Sử dụng CustomTkinter hiện đại, trực quan, dễ sử dụng.
"""
import os
import sys
import datetime
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

try:
    import customtkinter as ctk
    ctk.set_appearance_mode("System")  # Modes: "System", "Dark", "Light"
    ctk.set_default_color_theme("blue")  # Themes: "blue", "green", "dark-blue"
    USE_CTK = True
except ImportError:
    USE_CTK = False

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from core.config import CATEGORIES, BASE_DIR, OUTPUT_DIR
from core.form_cloner import clone_form_set, auto_generate_doc_number, auto_generate_date
from batch_generator import process_batch_excel


class FormGeneratorApp:
    def __init__(self):
        if USE_CTK:
            self.root = ctk.CTk()
        else:
            self.root = tk.Tk()
            
        self.root.title("QUATEST 3 - Hệ Thống Nhân Bản & Tái Chế Biểu Mẫu Hiện Trường")
        self.root.geometry("980x760")
        self.root.minsize(880, 680)

        # Trạng thái dữ liệu
        self.cat_map = {
            "Phòng sạch, Vi sinh": "phong_sach",
            "Khí nén, Khí Nitơ": "khi_nen",
            "CO2 - Môi trường lao động": "co2_mtld",
            "Kho lạnh (Mapping)": "kho_lanh",
            "Tủ an toàn sinh học (BSC)": "bsc"
        }
        self.selected_indicators_vars = []

        self._build_ui()

    def _build_ui(self):
        if USE_CTK:
            # Tiêu đề Header
            header_frame = ctk.CTkFrame(self.root, corner_radius=0, fg_color=("#1f538d", "#14375e"))
            header_frame.pack(fill="x", padx=0, pady=0)

            lbl_title = ctk.CTkLabel(
                header_frame,
                text="HỆ THỐNG NHÂN BẢN & TÁI CHẾ BIỂU MẪU HIỆN TRƯỜNG QUATEST 3",
                font=ctk.CTkFont(size=18, weight="bold"),
                text_color="white"
            )
            lbl_title.pack(pady=(12, 2))

            lbl_sub = ctk.CTkLabel(
                header_frame,
                text="Tự động tái chế và điền thông tin cho Biên bản giám định (BM3030.02) & Phiếu ghi chép (BM3030.01)",
                font=ctk.CTkFont(size=12),
                text_color="#e0e0e0"
            )
            lbl_sub.pack(pady=(0, 10))

            # Tab View
            self.tabview = ctk.CTkTabview(self.root, width=940)
            self.tabview.pack(fill="both", expand=True, padx=15, pady=(5, 10))

            self.tab_single = self.tabview.add("  Tạo Hồ Sơ Đơn Lẻ  ")
            self.tab_batch = self.tabview.add("  Nhân Bản Hàng Loạt (Excel)  ")
            self.tab_about = self.tabview.add("  Hướng Dẫn & Cấu Trúc  ")

            self._build_single_tab()
            self._build_batch_tab()
            self._build_about_tab()

            # Footer / Status
            footer_frame = ctk.CTkFrame(self.root, height=35, corner_radius=0)
            footer_frame.pack(fill="x", side="bottom")

            self.lbl_status = ctk.CTkLabel(footer_frame, text="Sẵn sàng thực hiện.", font=ctk.CTkFont(size=12))
            self.lbl_status.pack(side="left", padx=15, pady=4)

            btn_open_out = ctk.CTkButton(
                footer_frame,
                text="📂 Mở Thư Mục Output",
                width=160,
                height=26,
                command=self.open_output_dir
            )
            btn_open_out.pack(side="right", padx=15, pady=4)

    def _build_single_tab(self):
        container = ctk.CTkScrollableFrame(self.tab_single, corner_radius=6)
        container.pack(fill="both", expand=True, padx=5, pady=5)

        # 1. Chọn Lĩnh Vực
        lbl_sec1 = ctk.CTkLabel(container, text="1. LĨNH VỰC NGHIỆP VỤ & BIỂU MẪU", font=ctk.CTkFont(size=14, weight="bold"))
        lbl_sec1.grid(row=0, column=0, columnspan=2, sticky="w", pady=(5, 5))

        self.cb_category = ctk.CTkComboBox(
            container,
            values=list(self.cat_map.keys()),
            width=360,
            command=self._on_category_changed
        )
        self.cb_category.set("Phòng sạch, Vi sinh")
        self.cb_category.grid(row=1, column=0, sticky="w", padx=(10, 0), pady=(0, 10))

        # Checkbox chọn sinh biểu mẫu
        cb_frame = ctk.CTkFrame(container, fg_color="transparent")
        cb_frame.grid(row=1, column=1, sticky="w", padx=10, pady=(0, 10))
        self.var_bm02 = ctk.BooleanVar(value=True)
        self.var_bm01 = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(cb_frame, text="Biên bản BM3030.02 (BBLM)", variable=self.var_bm02).pack(side="left", padx=(0, 15))
        ctk.CTkCheckBox(cb_frame, text="Phiếu ghi chép BM3030.01", variable=self.var_bm01).pack(side="left")

        # 2. Thông tin khách hàng
        lbl_sec2 = ctk.CTkLabel(container, text="2. THÔNG TIN KHÁCH HÀNG & HIỆN TRƯỜNG", font=ctk.CTkFont(size=14, weight="bold"))
        lbl_sec2.grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 5))

        # Tên khách hàng
        ctk.CTkLabel(container, text="Tên khách hàng / Đơn vị yêu cầu (*):").grid(row=3, column=0, sticky="w", padx=10, pady=2)
        self.entry_customer = ctk.CTkEntry(container, width=550, placeholder_text="Ví dụ: Công ty Cổ phần Dược phẩm BioPharm Việt Nam")
        self.entry_customer.grid(row=3, column=1, sticky="w", padx=10, pady=2)

        # Địa chỉ
        ctk.CTkLabel(container, text="Địa chỉ đơn vị yêu cầu:").grid(row=4, column=0, sticky="w", padx=10, pady=2)
        self.entry_address = ctk.CTkEntry(container, width=550, placeholder_text="Ví dụ: Lô C2, Đường số 3, KCN Tân Đông Hiệp B, Dĩ An, Bình Dương")
        self.entry_address.grid(row=4, column=1, sticky="w", padx=10, pady=2)

        # Địa điểm thực hiện
        ctk.CTkLabel(container, text="Địa điểm thực hiện giám định:").grid(row=5, column=0, sticky="w", padx=10, pady=2)
        self.entry_location = ctk.CTkEntry(container, width=550, placeholder_text="Ví dụ: Nhà máy BioPharm - Xưởng 2, Phòng sạch cấp B và C")
        self.entry_location.grid(row=5, column=1, sticky="w", padx=10, pady=2)

        # Mã số hồ sơ & Ngày thực hiện
        sub_info_frame = ctk.CTkFrame(container, fg_color="transparent")
        sub_info_frame.grid(row=6, column=0, columnspan=2, sticky="w", padx=0, pady=4)

        ctk.CTkLabel(sub_info_frame, text="Mã số phiếu / hồ sơ:").pack(side="left", padx=(10, 5))
        self.entry_doc_num = ctk.CTkEntry(sub_info_frame, width=140)
        self.entry_doc_num.insert(0, auto_generate_doc_number())
        self.entry_doc_num.pack(side="left", padx=5)

        btn_regen_doc = ctk.CTkButton(sub_info_frame, text="🔄 Sinh mã", width=70, command=self._regen_doc_num)
        btn_regen_doc.pack(side="left", padx=5)

        ctk.CTkLabel(sub_info_frame, text="Ngày thực hiện:").pack(side="left", padx=(25, 5))
        self.entry_date = ctk.CTkEntry(sub_info_frame, width=110)
        self.entry_date.insert(0, auto_generate_date())
        self.entry_date.pack(side="left", padx=5)

        ctk.CTkLabel(sub_info_frame, text="Số lượng chỉ tiêu:").pack(side="left", padx=(25, 5))
        self.entry_num_samples = ctk.CTkEntry(sub_info_frame, width=60)
        self.entry_num_samples.insert(0, "4")
        self.entry_num_samples.pack(side="left", padx=5)

        # 3. Danh mục chỉ tiêu
        lbl_sec3 = ctk.CTkLabel(container, text="3. DANH MỤC CHỈ TIÊU / LẤY MẪU (Table 2)", font=ctk.CTkFont(size=14, weight="bold"))
        lbl_sec3.grid(row=7, column=0, columnspan=2, sticky="w", pady=(15, 5))

        self.indicators_container = ctk.CTkFrame(container)
        self.indicators_container.grid(row=8, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
        self._populate_indicators("phong_sach")

        # Nút tạo hồ sơ
        btn_generate = ctk.CTkButton(
            container,
            text="🚀 TẠO HỒ SƠ BIỂU MẪU MỚI",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=42,
            fg_color="#1f7a1f",
            hover_color="#145214",
            command=self._on_single_generate
        )
        btn_generate.grid(row=9, column=0, columnspan=2, sticky="ew", padx=10, pady=(20, 10))

    def _build_batch_tab(self):
        container = ctk.CTkFrame(self.tab_batch)
        container.pack(fill="both", expand=True, padx=10, pady=10)

        lbl_desc = ctk.CTkLabel(
            container,
            text="Nhân bản hàng loạt hàng chục hoặc hàng trăm hồ sơ khách hàng từ file Excel (.xlsx).\nHệ thống tự động đọc từng dòng và tạo cặp biểu mẫu hoàn chỉnh cho từng khách hàng.",
            font=ctk.CTkFont(size=13),
            justify="left"
        )
        lbl_desc.pack(anchor="w", padx=15, pady=(15, 10))

        # Chọn file Excel
        file_frame = ctk.CTkFrame(container, fg_color="transparent")
        file_frame.pack(fill="x", padx=15, pady=10)

        self.entry_excel_path = ctk.CTkEntry(file_frame, width=600, placeholder_text="Đường dẫn file Excel danh sách khách hàng (.xlsx)")
        default_sample = os.path.join(BASE_DIR, "danh_sach_khach_hang_mau.xlsx")
        if os.path.exists(default_sample):
            self.entry_excel_path.insert(0, default_sample)
        self.entry_excel_path.pack(side="left", padx=(0, 10))

        btn_browse = ctk.CTkButton(file_frame, text="🔍 Chọn file...", width=100, command=self._browse_excel)
        btn_browse.pack(side="left", padx=5)

        btn_open_sample = ctk.CTkButton(file_frame, text="📄 Mở file mẫu", width=100, fg_color="#555", hover_color="#333", command=self._open_sample_excel)
        btn_open_sample.pack(side="left", padx=5)

        # Nút thực hiện
        self.btn_run_batch = ctk.CTkButton(
            container,
            text="🚀 BẮT ĐẦU NHÂN BẢN HÀNG LOẠT TỪ EXCEL",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=38,
            fg_color="#1f7a1f",
            hover_color="#145214",
            command=self._on_run_batch
        )
        self.btn_run_batch.pack(fill="x", padx=15, pady=(15, 10))

        self.progressbar = ctk.CTkProgressBar(container)
        self.progressbar.pack(fill="x", padx=15, pady=5)
        self.progressbar.set(0)

        # Khung log
        ctk.CTkLabel(container, text="Nhật ký xử lý:", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", padx=15, pady=(10, 2))
        self.txt_log = ctk.CTkTextbox(container, height=220)
        self.txt_log.pack(fill="both", expand=True, padx=15, pady=(0, 15))

    def _build_about_tab(self):
        container = ctk.CTkScrollableFrame(self.tab_about)
        container.pack(fill="both", expand=True, padx=10, pady=10)

        about_text = """
HỆ THỐNG NHÂN BẢN VÀ TÁI CHẾ BIỂU MẪU HIỆN TRƯỜNG QUATEST 3

1. MỤC ĐÍCH:
- Tái sử dụng và nhân bản nhanh các biểu mẫu chuẩn mực của QUATEST 3 cho các đợt đi hiện trường:
  * BM3030.02: Biên bản giám định hiện trường / Biên bản lấy mẫu (BBLM).
  * BM3030.01: Phiếu ghi chép kết quả đo đạc tại hiện trường.
- Tự động thay thế thông tin khách hàng, số hồ sơ, ngày tháng và tự động điều chỉnh số dòng của bảng chỉ tiêu (Table 2) đúng theo yêu cầu.
- Làm sạch các số liệu đo cũ để kỹ thuật viên ghi chép tay tại hiện trường.

2. CÁC LĨNH VỰC NGHIỆP VỤ HỖ TRỢ:
  1. Phòng sạch, Vi sinh: Nồng độ hạt bụi, vi sinh không khí, chênh áp, nhiệt ẩm, trao đổi gió.
  2. Khí nén, Khí Nitơ: Tổng dầu, nhiệt độ điểm sương, nồng độ hạt bụi trong khí nén, khí Nitơ.
  3. CO2 - Môi trường lao động: Khí CO2 và tổng chất hữu cơ bay hơi (TVOC).
  4. Kho lạnh: Đánh giá độ đồng đều nhiệt độ và độ ẩm kho bảo quản thuốc/vắc-xin.
  5. Tủ an toàn sinh học (BSC): Vận tốc inflow, downflow, HEPA leak test, cường độ sáng, độ ồn, UVC.

3. ĐẶC ĐIỂM KỸ THUẬT:
- Giữ nguyên 100% định dạng file Word gốc (Font Times New Roman, cỡ chữ, căn lề, viền bảng, header/footer).
- Bảng chỉ tiêu co giãn linh hoạt từ 1 đến hàng chục dòng, tự động đánh số thứ tự (STT: 1, 2, 3...).
- File kết quả được tự động sắp xếp vào thư mục Output theo cấu trúc chuẩn:
  Output/[Số_Hồ_Sơ] - [Tên_Khách_Hàng] - [Nội_Dung_Vắn_Tắt]/
    ├── [Số_Hồ_Sơ] - BM3030.02 - BBLM - [Tên_KH].docx
    └── [Số_Hồ_Sơ] - BM3030.01 - Ghi chep - [Tên_KH].docx

  Ví dụ:
    * Q326D124674 - Cát Tân - Phòng sạch
    * Q326D120711 - Greencans Industries - Nước uống
    * Q326D117382 - Thụy Khuê - Nuoc sinh hoat
    * Q326D124796 - K pack - vskk - swab
    (VSKK: Vi sinh không khí, VSBM: Vi sinh bề mặt / swab)
"""
        lbl_info = ctk.CTkLabel(container, text=about_text, font=ctk.CTkFont(size=12), justify="left")
        lbl_info.pack(anchor="w", padx=10, pady=10)

    def _on_category_changed(self, choice):
        cat_key = self.cat_map.get(choice, "phong_sach")
        self._populate_indicators(cat_key)

    def _populate_indicators(self, cat_key):
        for widget in self.indicators_container.winfo_children():
            widget.destroy()
            
        self.selected_indicators_vars = []
        cfg = CATEGORIES.get(cat_key, CATEGORIES["phong_sach"])
        indicators = cfg.get("indicators", [])

        lbl = ctk.CTkLabel(self.indicators_container, text=f"Tích chọn các chỉ tiêu muốn áp dụng ({cfg['name']}):", font=ctk.CTkFont(size=12, slant="italic"))
        lbl.pack(anchor="w", padx=10, pady=(5, 5))

        for ind in indicators:
            var = ctk.BooleanVar(value=True)
            cb = ctk.CTkCheckBox(
                self.indicators_container,
                text=ind["name"],
                variable=var,
                command=self._update_num_samples_from_checkboxes
            )
            cb.pack(anchor="w", padx=15, pady=3)
            self.selected_indicators_vars.append((ind, var))

        self._update_num_samples_from_checkboxes()

    def _update_num_samples_from_checkboxes(self):
        count = sum(1 for _, v in self.selected_indicators_vars if v.get())
        if count > 0:
            self.entry_num_samples.delete(0, "end")
            self.entry_num_samples.insert(0, str(count))

    def _regen_doc_num(self):
        self.entry_doc_num.delete(0, "end")
        self.entry_doc_num.insert(0, auto_generate_doc_number())

    def _on_single_generate(self):
        customer = self.entry_customer.get().strip()
        if not customer:
            messagebox.showwarning("Thiếu thông tin", "Vui lòng nhập Tên khách hàng / Đơn vị yêu cầu!")
            return

        address = self.entry_address.get().strip()
        location = self.entry_location.get().strip()
        doc_num = self.entry_doc_num.get().strip()
        date_val = self.entry_date.get().strip()

        try:
            num_samples = int(self.entry_num_samples.get().strip())
        except ValueError:
            num_samples = 4

        cat_choice = self.cb_category.get()
        cat_key = self.cat_map.get(cat_choice, "phong_sach")

        # Lấy danh sách chỉ tiêu được tích chọn
        selected_inds = [ind for ind, v in self.selected_indicators_vars if v.get()]
        if not selected_inds:
            selected_inds = None

        self.lbl_status.configure(text=f"Đang tạo hồ sơ cho: {customer}...")
        self.root.update_idletasks()

        try:
            res = clone_form_set(
                category_key=cat_key,
                customer_name=customer,
                address=address if address else customer,
                location=location if location else (address if address else customer),
                num_samples=num_samples,
                doc_number=doc_num,
                date=date_val,
                indicators=selected_inds,
                generate_bm01=self.var_bm01.get(),
                generate_bm02=self.var_bm02.get()
            )

            if res["success"]:
                self.lbl_status.configure(text=f"✅ Đã tạo thành công hồ sơ: {customer}")
                succ_msg = f"Đã nhân bản trọn bộ hồ sơ thành công!\n\nThư mục lưu:\n{res['output_dir']}"
                if res.get("z_drive_copy"):
                    succ_msg += f"\n\nĐã sao chép sang ổ Z:\n{res['z_drive_copy']}"
                succ_msg += "\n\nBạn có muốn mở thư mục chứa file ngay bây giờ không?"
                ans = messagebox.askyesno(
                    "Thành công",
                    succ_msg
                )
                if ans:
                    os.startfile(res["output_dir"])
            else:
                err_msg = "\n".join(res["errors"])
                messagebox.showerror("Lỗi nhân bản", f"Có lỗi xảy ra:\n{err_msg}")
                self.lbl_status.configure(text="❌ Có lỗi xảy ra trong quá trình nhân bản.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi không xác định: {str(e)}")
            self.lbl_status.configure(text="❌ Lỗi thực thi.")

    def _browse_excel(self):
        f = filedialog.askopenfilename(
            title="Chọn file Excel danh sách khách hàng",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if f:
            self.entry_excel_path.delete(0, "end")
            self.entry_excel_path.insert(0, f)

    def _open_sample_excel(self):
        sample_path = os.path.join(BASE_DIR, "danh_sach_khach_hang_mau.xlsx")
        if os.path.exists(sample_path):
            os.startfile(sample_path)
        else:
            messagebox.showinfo("Thông báo", f"Không tìm thấy file mẫu tại {sample_path}")

    def _on_run_batch(self):
        excel_path = self.entry_excel_path.get().strip()
        if not os.path.exists(excel_path):
            messagebox.showerror("Lỗi", f"Không tìm thấy file Excel:\n{excel_path}")
            return

        self.btn_run_batch.configure(state="disabled")
        self.txt_log.delete("1.0", "end")
        self.txt_log.insert("end", f"Bắt đầu xử lý file Excel: {excel_path}\n")
        self.progressbar.set(0)

        def worker():
            def on_progress(idx, total, name):
                prog = idx / max(1, total)
                self.progressbar.set(prog)
                self.txt_log.insert("end", f"[{idx}/{total}] ✅ Đã tạo hồ sơ: {name}\n")
                self.txt_log.see("end")

            try:
                summary = process_batch_excel(excel_path, progress_callback=on_progress)
                self.txt_log.insert("end", "\n" + "=" * 50 + "\n")
                self.txt_log.insert("end", f"HOÀN THÀNH: Thành công {summary['success']}/{summary['total']} hồ sơ (Thất bại: {summary['failed']})\n")
                self.lbl_status.configure(text=f"✅ Hoàn thành nhân bản {summary['success']}/{summary['total']} hồ sơ từ Excel.")
                messagebox.showinfo("Hoàn thành", f"Đã nhân bản thành công {summary['success']}/{summary['total']} hồ sơ!")
            except Exception as e:
                self.txt_log.insert("end", f"\n❌ LỖI: {str(e)}\n")
                messagebox.showerror("Lỗi", f"Quá trình xử lý bị lỗi: {str(e)}")
            finally:
                self.btn_run_batch.configure(state="normal")

        threading.Thread(target=worker, daemon=True).start()

    def open_output_dir(self):
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        os.startfile(OUTPUT_DIR)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = FormGeneratorApp()
    app.run()
