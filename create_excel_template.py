"""
Tạo file Excel mẫu để người dùng nhập thông tin hàng loạt khách hàng
"""
import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def generate_sample_excel(file_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DanhSachKhachHang"
    
    headers = [
        "STT",
        "Lĩnh vực / Biểu mẫu",
        "Tên khách hàng / Đơn vị yêu cầu",
        "Địa chỉ đơn vị yêu cầu",
        "Địa điểm thực hiện giám định",
        "Số lượng chỉ tiêu / mẫu",
        "Mã số phiếu / hồ sơ (tùy chọn)",
        "Ngày thực hiện (tùy chọn)",
        "Nội dung yêu cầu (tùy chọn)",
        "Ghi chú / Chỉ tiêu chi tiết (tùy chọn)"
    ]
    
    # Header styling
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='BFBFBF')
    )
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        
    sample_rows = [
        [
            1,
            "Phòng sạch, Vi sinh",
            "Công ty Cổ phần Dược phẩm BioPharm Việt Nam",
            "Lô C2, Đường số 3, KCN Tân Đông Hiệp B, Dĩ An, Bình Dương",
            "Nhà máy BioPharm - Phòng vô trùng cấp B & C",
            5,
            "Q326D108033",
            "15/10/2026",
            "Quan trắc chất lượng phòng sạch và lấy mẫu vi sinh trong không khí",
            "Đo bụi; Đo nhiệt ẩm; Đo chênh áp; Vi sinh không khí; Trao đổi gió"
        ],
        [
            2,
            "Khí nén, Khí Nitơ",
            "Chi nhánh Công ty TNHH Bia Anheuser-Busch InBev",
            "Khu công nghiệp VSIP II-A, Tân Uyên, Bình Dương",
            "Phân xưởng chiết rót bia lon - Đường ống khí nén máy đóng nắp",
            4,
            "",
            "18/10/2026",
            "Lấy mẫu phân tích tổng dầu, đo nhiệt độ điểm sương và nồng độ hạt bụi",
            "Tổng dầu; Điểm sương; Nồng độ hạt bụi; Khí Nito"
        ],
        [
            3,
            "CO2 - Môi trường lao động",
            "Ngân hàng TMCP Ngoại Thương Việt Nam (Vietcombank)",
            "Số 5 Công trường Mê Linh, Bến Nghé, Quận 1, TP. HCM",
            "Tòa nhà Vietcombank Tower - Tầng 10 đến Tầng 12",
            2,
            "",
            "20/10/2026",
            "Quan trắc môi trường làm việc văn phòng (CO2 và TVOC)",
            "Đo nồng độ TVOC; Đo nồng độ CO2"
        ],
        [
            4,
            "Kho lạnh",
            "Công ty TNHH Vận tải Dược phẩm MedLog",
            "Kho Logistics Tân Bình, 120 Trường Chinh, Q. Tân Bình, TP. HCM",
            "Kho lạnh bảo quản dược phẩm số 1 (Dải nhiệt 2-8 độ C)",
            3,
            "",
            "",
            "Đánh giá độ đồng đều nhiệt độ và độ ẩm kho bảo quản thuốc",
            "Đồng đều nhiệt độ; Đồng đều độ ẩm"
        ],
        [
            5,
            "Tủ an toàn sinh học (BSC)",
            "Trung tâm Kiểm nghiệm Y tế Quốc gia",
            "70 Bà Huyện Thanh Quan, Quận 3, TP. Hồ Chí Minh",
            "Phòng Xét nghiệm Sinh học phân tử - Tủ BSC số 1 và số 2",
            2,
            "",
            "",
            "Giám định an toàn và hiệu năng tủ an toàn sinh học theo EN 12469:2000",
            "Inflow; Downflow; HEPA Leak; Cường độ sáng; Độ ồn; UVC"
        ]
    ]
    
    data_font = Font(name="Times New Roman", size=10)
    for row_idx, row_data in enumerate(sample_rows, 2):
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = data_font
            cell.border = thin_border
            if col_idx in [1, 6]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                
    ws.row_dimensions[1].height = 35
    for r in range(2, len(sample_rows) + 2):
        ws.row_dimensions[r].height = 28
        
    col_widths = {1: 8, 2: 24, 3: 38, 4: 45, 5: 45, 6: 18, 7: 22, 8: 18, 9: 35, 10: 35}
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width
        
    wb.save(file_path)
    print(f"Sample Excel generated successfully at: {file_path}")

if __name__ == "__main__":
    base_dir = r"c:\Users\phat.nt\Desktop\Source\Form Bieu Mau Goc"
    target = os.path.join(base_dir, "danh_sach_khach_hang_mau.xlsx")
    generate_sample_excel(target)
