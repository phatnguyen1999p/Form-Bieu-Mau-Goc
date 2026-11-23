"""
Module chuẩn hóa thông tin dữ liệu cho hệ thống hồ sơ biểu mẫu QUATEST 3
Tuân thủ nghiêm ngặt 11 Quy tắc soạn thảo chuẩn mực (GEMINI.md).
"""
import os
import re
import datetime
import random
import sqlite3

# ==============================================================================
# 1. CHUẨN HÓA MÃ HỒ SƠ & NGÀY THÁNG (Quy tắc 1 & 2)
# ==============================================================================

def normalize_date_str(d_str, default_year="2026"):
    """
    Chuẩn hóa chuỗi ngày tháng sang định dạng DD/MM/YYYY (VD: '02/10' -> '02/10/2026').
    Nếu là ngày hiện trường chưa thực hiện, trả về chuỗi chấm lửng '…./…./YYYY'.
    """
    if not d_str:
        return ""
    s = str(d_str).strip()
    if "…" in s or "..." in s:
        # Giữ nguyên chuỗi chấm lửng
        return s
    if re.match(r"^\d{1,2}/\d{1,2}$", s):
        s = f"{s}/{default_year}"
    parts = s.split("/")
    if len(parts) == 3:
        try:
            day = f"{int(parts[0]):02d}"
            month = f"{int(parts[1]):02d}"
            year = parts[2]
            return f"{day}/{month}/{year}"
        except ValueError:
            pass
    return s


def auto_generate_doc_number(year=None):
    """Tự động sinh mã số hồ sơ theo quy cách QUATEST 3 (VD: Q326D + 6 chữ số)."""
    now = datetime.datetime.now()
    yr = str(year if year else now.year)[-2:]
    year_prefix = f"Q3{yr}D"
    random_digits = f"{random.randint(100000, 999999)}"
    return f"{year_prefix}{random_digits}"


# [NEW 27/10/2026] Hỗ trợ "Mã hồ sơ để trắng" (chưa có số phiếu yêu cầu)
DOC_NUMBER_BLANK_PLACEHOLDER = "Q326D……………"   # Ghi trong biểu mẫu: giữ tiền tố Q326D, chừa chấm để KTV điền tay
DOC_NUMBER_BLANK_FILENAME = "Q326D......"      # Dùng cho tên thư mục / tên file (ASCII, hợp lệ trên Windows)
BLANK_DOC_TOKENS = {"blank", "trong", "trắng", "de trang", "để trắng", "...", "…", "-", "q326d", "q326d...", "q326d……………"}


def is_blank_doc_number(value):
    """True nếu người dùng yêu cầu để trắng mã hồ sơ (giá trị là token BLANK/để trắng/...)."""
    if value is None:
        return False
    v = str(value).strip().lower()
    return v in BLANK_DOC_TOKENS or v == DOC_NUMBER_BLANK_PLACEHOLDER.lower()


def doc_number_for_filename(doc_number):
    """Chuyển mã hồ sơ placeholder sang dạng an toàn cho tên file/thư mục."""
    if doc_number == DOC_NUMBER_BLANK_PLACEHOLDER:
        return DOC_NUMBER_BLANK_FILENAME
    return doc_number


def auto_generate_date():
    """Tự động lấy ngày hiện tại định dạng DD/MM/YYYY."""
    now = datetime.datetime.now()
    return now.strftime("%d/%m/%Y")


def get_default_execution_date_dotted(year="2026"):
    """Trả về chuỗi ngày hiện trường chấm lửng theo Quy tắc 2: '…./…./YYYY'."""
    return f"…./…./{year}"


# ==============================================================================
# 2. CHUẨN HÓA ĐỊA CHỈ & ĐỊA ĐIỂM (Quy tắc 8 & 11)
# ==============================================================================

def normalize_address(address_str):
    """
    Chuẩn hóa địa chỉ theo quy tắc viết hoa - viết thường chuẩn hành chính Việt Nam
    và mở rộng các từ viết tắt thành nguyên chữ (Quy tắc 8):
    - Mở rộng KCN, KCX, KDC, CCN, KĐT, KKT, KCNC, QL, TL, Đ., P., Q., H., TX., TT., X., KP., TDP.
    - Loại bỏ 'Việt Nam', 'Vietnam', 'VN' ở cuối địa chỉ.
    - Thành phố trực thuộc Trung ương: viết hoa chữ 'Thành phố' (Thành phố Hồ Chí Minh, Thành phố Hà Nội...).
    - Tỉnh và Thành phố thuộc tỉnh: viết thường chữ 'tỉnh' và 'thành phố' (tỉnh Bình Dương, thành phố Bến Cát...).
    - Quận / Huyện / Thị xã: viết thường nếu tên là chữ; viết hoa nếu tên là số (Quận 1, Quận 12).
    - Phường / Xã / Thị trấn: viết thường nếu tên là chữ; viết hoa nếu tên là số (Phường 1, Phường 14).
    - Khu phố: viết hoa chữ 'Khu', viết thường chữ 'phố'.
    - Quy tắc 2 cấp hành chính: Bắt buộc chỉ dùng [Cấp xã/phường/thị trấn], [Cấp tỉnh/TP trực thuộc TW].
    """
    if not address_str:
        return ""
    s = str(address_str).strip()

    # Lưu lại tiền tố dòng chấm điền tay (nếu có) trước dấu gạch ngang (VD: "......... - [Địa chỉ]")
    prefix = ""
    m_prefix = re.match(r"^([.…]+\s*-\s*)(.*)$", s)
    if m_prefix:
        prefix = m_prefix.group(1)
        s = m_prefix.group(2)

    # 1. Bỏ 'Việt Nam', 'Vietnam', 'VN' ở cuối chuỗi
    s = re.sub(r'[,.\s]+(?:Việt\s*Nam|Vietnam|VN)\s*$', '', s, flags=re.IGNORECASE)

    # 2. Mở rộng các từ viết tắt khu công nghiệp, khu chế xuất, khu dân cư... thành nguyên chữ
    s = re.sub(r'\bKCN\b', 'Khu công nghiệp', s, flags=re.IGNORECASE)
    s = re.sub(r'\bKCX\b', 'Khu chế xuất', s, flags=re.IGNORECASE)
    s = re.sub(r'\bKDC\b', 'Khu dân cư', s, flags=re.IGNORECASE)
    s = re.sub(r'\bCCN\b', 'Cụm công nghiệp', s, flags=re.IGNORECASE)
    s = re.sub(r'\bKĐT\b', 'Khu đô thị', s, flags=re.IGNORECASE)
    s = re.sub(r'\bKKT\b', 'Khu kinh tế', s, flags=re.IGNORECASE)
    s = re.sub(r'\bKCNC\b', 'Khu công nghệ cao', s, flags=re.IGNORECASE)

    # Chuẩn hóa viết hoa chữ cái đầu cho các cụm khu
    s = re.sub(r'\b[kK]hu\s+[cC]ông\s+[nN]ghiệp\b', 'Khu công nghiệp', s)
    s = re.sub(r'\b[kK]hu\s+[cC]hế\s+[xX]uất\b', 'Khu chế xuất', s)
    s = re.sub(r'\b[kK]hu\s+[dD]ân\s+[cC]ư\b', 'Khu dân cư', s)
    s = re.sub(r'\b[kK]hu\s+[đĐ]ô\s+[tT]hị\b', 'Khu đô thị', s)
    s = re.sub(r'\b[kK]hu\s+[kK]inh\s+[tT]ế\b', 'Khu kinh tế', s)
    s = re.sub(r'\b[cC]ụm\s+[cC]ông\s+[nN]ghiệp\b', 'Cụm công nghiệp', s)
    s = re.sub(r'\b[kK]hu\s+[cC]ông\s+[nN]ghệ\s+[cC]ao\b', 'Khu công nghệ cao', s)

    # Tách và chuẩn hóa các cụm khu công nghiệp kèm địa danh viết liền (ví dụ: KCN Vĩnh Lộc Hóc Môn -> Khu công nghiệp Vĩnh Lộc, xã Hóc Môn)
    s = re.sub(r'\b(?:Khu\s+công\s+nghiệp|KCN)\s+Vĩnh\s+Lộc\s+Hóc\s+Môn\b', 'Khu công nghiệp Vĩnh Lộc, xã Hóc Môn', s, flags=re.IGNORECASE)

    # Bổ sung tiền tố 'thành phố' cho các thành phố thuộc tỉnh nếu bị thiếu tiền tố trước tên tỉnh
    provincial_cities_pattern = r'(?:Dĩ\s*An|Thuận\s*An|Bến\s*Cát|Tân\s*Uyên|Thủ\s*Dầu\s*Một)'
    s = re.sub(rf'([,;]\s*)({provincial_cities_pattern})(?=\s*,\s*(?:[tT]ỉnh\s+)?Bình\s*Dương)', r'\1thành phố \2', s)
    s = re.sub(r'([,;]\s*)(Thủ\s*Đức)(?=\s*,\s*(?:[tT]hành\s*phố\s*|TP\.?\s*)?Hồ\s*Chí\s*Minh)', r'\1thành phố \2', s, flags=re.IGNORECASE)

    # 3. Quốc lộ, Tỉnh lộ (QL., TL.)
    s = re.sub(r'\b(?:QL\.?|Quốc\s*[lL]ộ)\s*', 'Quốc lộ ', s)
    s = re.sub(r'\b(?:TL\.?|Tỉnh\s*[lL]ộ)\s*', 'Tỉnh lộ ', s)

    # 4. Thành phố trực thuộc Trung ương -> 'Thành phố <Tên>' (Mở rộng Tp.HCM, TP.HCM...)
    s = re.sub(r'\b(?:TP\.?\s*HCM|TP\.?\s*Hồ\s*Chí\s*Minh|Thành\s*phố\s*Hồ\s*Chí\s*Minh|Tp\.?\s*Hồ\s*Chí\s*Minh|Tp\.?\s*HCM)\b', 'Thành phố Hồ Chí Minh', s, flags=re.IGNORECASE)
    s = re.sub(r'\b(?:TP\.?\s*Hà\s*Nội|Thành\s*phố\s*Hà\s*Nội|Tp\.?\s*Hà\s*Nội)\b', 'Thành phố Hà Nội', s, flags=re.IGNORECASE)
    s = re.sub(r'\b(?:TP\.?\s*Đà\s*Nẵng|Thành\s*phố\s*Đà\s*Nẵng|Tp\.?\s*Đà\s*Nẵng)\b', 'Thành phố Đà Nẵng', s, flags=re.IGNORECASE)
    s = re.sub(r'\b(?:TP\.?\s*Hải\s*Phòng|Thành\s*phố\s*Hải\s*Phòng|Tp\.?\s*Hải\s*Phòng)\b', 'Thành phố Hải Phòng', s, flags=re.IGNORECASE)
    s = re.sub(r'\b(?:TP\.?\s*Cần\s*Thơ|Thành\s*phố\s*Cần\s*Thơ|Tp\.?\s*Cần\s*Thơ)\b', 'Thành phố Cần Thơ', s, flags=re.IGNORECASE)
    s = re.sub(r'\b(?:TP\.?\s*Huế|Thành\s*phố\s*Huế|Tp\.?\s*Huế)\b', 'Thành phố Huế', s, flags=re.IGNORECASE)

    # Đứng sau dấu phẩy riêng lẻ ở cuối địa chỉ
    s = re.sub(r'([,;]\s*)Hồ\s*Chí\s*Minh(?=[,;.]|$)', r'\1Thành phố Hồ Chí Minh', s, flags=re.IGNORECASE)
    s = re.sub(r'([,;]\s*)Hà\s*Nội(?=[,;.]|$)', r'\1Thành phố Hà Nội', s, flags=re.IGNORECASE)
    s = re.sub(r'([,;]\s*)Đà\s*Nẵng(?=[,;.]|$)', r'\1Thành phố Đà Nẵng', s, flags=re.IGNORECASE)
    s = re.sub(r'([,;]\s*)Hải\s*Phòng(?=[,;.]|$)', r'\1Thành phố Hải Phòng', s, flags=re.IGNORECASE)
    s = re.sub(r'([,;]\s*)Cần\s*Thơ(?=[,;.]|$)', r'\1Thành phố Cần Thơ', s, flags=re.IGNORECASE)

    # 5. Các thành phố trực thuộc tỉnh: viết thường 'thành phố ...'
    # [11/11/2026] Bổ sung 'đồng nai' theo cách ghi 'Thành phố Đồng Nai' trên các hồ sơ thực tế (Q326D084627, Q326D067460)
    central_cities_lower = ['hồ chí minh', 'hà nội', 'đà nẵng', 'hải phòng', 'cần thơ', 'huế', 'đồng nai']
    def replace_provincial_city(m):
        c_name = m.group(1).strip()
        if c_name.lower() in central_cities_lower:
            return f'Thành phố {c_name}'
        return f'thành phố {c_name}'

    s = re.sub(r'\b(?:Thành\s*[pP]hố|TP\.?|Tp\.?)\s*([A-ZÀ-ỸĐ][\w\s–-]+?)(?=[,;\.]|$)', replace_provincial_city, s)

    # 6. Tỉnh: viết thường 'tỉnh ...'
    s = re.sub(r'\bTỉnh\s+', 'tỉnh ', s)

    # 7. Quận: (mở rộng Q.)
    # Nếu tên là số (Quận 1, Quận 12): viết hoa 'Quận'
    s = re.sub(r'\b(?:quận|Quận|Q\.?)\s*(\d+)\b', r'Quận \1', s, flags=re.IGNORECASE)
    # Nếu tên là chữ (quận Bình Thạnh, quận Tân Bình): viết thường 'quận'
    s = re.sub(r'\b(?:Quận|Q\.)\s*([A-ZÀ-ỸĐ][^\d,;.]+)', r'quận \1', s)

    # 8. Huyện, Thị xã: viết thường (mở rộng H., TX.)
    s = re.sub(r'\b(?:Huyện|H\.)\s*', 'huyện ', s)
    s = re.sub(r'\b(?:Thị\s*[xX]ã|TX\.?)\s*', 'thị xã ', s)

    # 9. Phường: (mở rộng P.)
    # Nếu tên là số (Phường 1, Phường 12): viết hoa 'Phường'
    s = re.sub(r'\b(?:phường|Phường|P\.?)\s*(\d+)\b', r'Phường \1', s, flags=re.IGNORECASE)
    # Nếu tên là chữ (phường Thới An, phường Bến Nghé): viết thường 'phường'
    s = re.sub(r'\b(?:Phường|P\.)\s*([A-ZÀ-ỸĐ][^\d,;.]+)', r'phường \1', s)

    # 10. Xã, Thị trấn: viết thường (mở rộng X., TT.)
    s = re.sub(r'\b(?:Xã|X\.)\s*', 'xã ', s)
    s = re.sub(r'\b(?:Thị\s*[tT]rấn|TT\.?)\s*', 'thị trấn ', s)

    # 11. Khu phố: viết hoa 'Khu', viết thường 'phố' (mở rộng KP.)
    s = re.sub(r'\b(?:Khu\s*[pP]hố|KP\.?)\s*', 'Khu phố ', s)

    # 12. Tổ dân phố, Tổ: viết hoa 'Tổ', viết thường 'dân phố' (mở rộng TDP.)
    s = re.sub(r'\b(?:Tổ\s*[dD]ân\s*[pP]hố|TDP\.?)\s*', 'Tổ dân phố ', s)
    s = re.sub(r'\b(?:Tổ)\s+(\d+)\b', r'Tổ \1', s)

    # 13. Thôn, Ấp: viết hoa 'Thôn', 'Ấp'
    s = re.sub(r'\b(?:thôn)\s+', 'Thôn ', s)
    s = re.sub(r'\b(?:ấp)\s+', 'Ấp ', s)

    # 14. Đường, Phố, Đại lộ, Hẻm, Ngõ: viết thường
    # [FIX 27/10/2026] Ưu tiên khớp 'Đường' trước và bắt buộc có dấu chấm với dạng viết tắt 'Đ.' / 'Đg.'
    # (bản cũ 'Đ\.?' khớp nhầm chữ Đ đầu từ: 'Đường DE5' -> 'đường ường DE5', 'Đồng Nai' -> 'đường ồng Nai')
    s = re.sub(r'\b(?:Đường|Đg\.|Đ\.)\s*', 'đường ', s)
    s = re.sub(r'\b(?:Phố)\s*', 'phố ', s)
    s = re.sub(r'\b(?:Đại\s*[lL]ộ)\s*', 'đại lộ ', s)
    s = re.sub(r'\b(?:Hẻm)\s*', 'hẻm ', s)
    s = re.sub(r'\b(?:Ngõ)\s*', 'ngõ ', s)

    # Dọn dẹp khoảng trắng và dấu phẩy thừa
    s = re.sub(r'\s*,\s*', ', ', s)
    s = re.sub(r',+', ',', s)
    s = re.sub(r'\s+', ' ', s).strip()
    s = s.rstrip(' ;,')

    if prefix:
        s = f"{prefix}{s}"

    return s


# [22/11/2026] Dải chấm điền tay trước địa chỉ ở dòng "Địa điểm thực hiện": 90 dấu chấm, 1 dòng.
LOCATION_DOTS = "." * 90


def normalize_location(location_str, address_str=None):
    """
    Chuẩn hóa Địa điểm thực hiện theo Quy tắc 11:
    - Không điền cứng tên phòng/khu vực/hệ thống/ống dẫn cụ thể.
    - Cấu trúc chuẩn mực: chừa 90 dấu chấm (LOCATION_DOTS) để KTV điền tay khu vực,
      rồi ' - ' và địa chỉ đã chuẩn hóa:
      '..........(90 chấm).......... - [Địa chỉ đã chuẩn hóa]'
    """
    dots_prefix = LOCATION_DOTS
    
    # Ưu tiên lấy location_str nếu có địa chỉ cụ thể
    base_addr = ""
    if location_str and str(location_str).strip():
        s_loc = str(location_str).strip()
        m = re.match(r"^[\s.…\n]+\s*-\s*(.*)$", s_loc)
        if m:
            s_loc = m.group(1).strip()
        if not re.fullmatch(r"[\s.…-]+", s_loc):
            base_addr = s_loc

    if not base_addr and address_str and str(address_str).strip():
        base_addr = address_str.strip()

    clean_addr = normalize_address(base_addr)
    return f"{dots_prefix} - {clean_addr}"


# ==============================================================================
# 3. CHUẨN HÓA TÊN KHÁCH HÀNG & NỘI DUNG VẮN TẮT (Quy tắc 9 & 10)
# ==============================================================================

COMPANY_TYPES = [
    r"chi\s+nhánh\s+công\s+ty\s+trách\s+nhiệm\s+hữu\s+hạn\s+một\s+thành\s+viên",
    r"chi\s+nhánh\s+công\s+ty\s+trách\s+nhiệm\s+hữu\s+hạn\s+mtv",
    r"chi\s+nhánh\s+công\s+ty\s+trách\s+nhiệm\s+hữu\s+hạn",
    r"chi\s+nhánh\s+công\s+ty\s+cổ\s+phần",
    r"chi\s+nhánh\s+công\s+ty\s+tnhh\s+mtv",
    r"chi\s+nhánh\s+công\s+ty\s+tnhh",
    r"chi\s+nhánh\s+công\s+ty\s+cp",
    r"chi\s+nhánh\s+công\s+ty",
    r"chi\s+nhánh\s+cty\s+tnhh",
    r"chi\s+nhánh\s+cty\s+cp",
    r"chi\s+nhánh\s+cty",
    r"chi\s+nhánh",
    r"công\s+ty\s+trách\s+nhiệm\s+hữu\s+hạn\s+một\s+thành\s+viên",
    r"công\s+ty\s+trách\s+nhiệm\s+hữu\s+hạn\s+mtv",
    r"công\s+ty\s+trách\s+nhiệm\s+hữu\s+hạn",
    r"công\s+ty\s+tnhh\s+một\s+thành\s+viên",
    r"công\s+ty\s+tnhh\s+mtv",
    r"công\s+ty\s+tnhh",
    r"công\s+ty\s+cổ\s+phần",
    r"công\s+ty\s+cp",
    r"công\s+ty\s+liên\s+doanh",
    r"công\s+ty\s+hợp\s+danh",
    r"doanh\s+nghiệp\s+tư\s+nhân",
    r"dntn",
    r"tổng\s+công\s+ty",
    r"tập\s+đoàn",
    r"cty\s+tnhh\s+mtv",
    r"cty\s+tnhh",
    r"cty\s+cp",
    r"cty",
    r"công\s+ty",
    r"hợp\s+tác\s+xã",
    r"htx",
    r"nhà\s+máy",
    r"xí\s+nghiệp",
]

GENERIC_INDUSTRIES = [
    r"sản\s+xuất\s+thương\s+mại\s+dịch\s+vụ",
    r"sản\s+xuất\s+thương\s+mại",
    r"thương\s+mại\s+dịch\s+vụ",
    r"thương\s+mại\s+và\s+dịch\s+vụ",
    r"thương\s+mại\s+xuất\s+nhập\s+khẩu",
    r"dịch\s+vụ\s+thương\s+mại",
    r"sx\s*[-–—/&]\s*tm\s*[-–—/&]\s*dv",
    r"sx\s*tm\s*dv",
    r"tm\s*[-–—/&]\s*dv",
    r"tm\s*dv",
    r"sx\s*[-–—/&]\s*tm",
    r"sx\s*tm",
    r"xnk",
    r"xuất\s+nhập\s+khẩu",
    r"đầu\s+tư\s+và\s+phát\s+triển",
    r"đầu\s+tư\s+phát\s+triển",
    r"đầu\s+tư\s+xây\s+dựng",
    r"đầu\s+tư",
    r"công\s+nghệ\s+sạch",
    r"công\s+nghệ",
    r"dược\s+phẩm",
    r"thiết\s+bị\s+y\s+tế",
]


UPPERCASE_ACRONYMS = {
    "TNHH", "MTV", "CP", "CTCP", "DNTN", "HTX",
    "SX", "TM", "DV", "XNK", "TMDV", "SXTM", "SXTMDV",
    "KCN", "KCX", "KDC", "KĐT", "CCN", "KCNC",
    "PTN", "TTKT", "VN", "TP", "HCM", "HN",
    "USA", "UK", "EU", "ASEAN", "APAC",
    "ABC", "ABT", "ACB", "BIDV", "VCB", "CTG", "MB", "VPB", "SHB",
    "FPT", "VNPT", "EVN", "PVN", "VRG", "VISSAN",
    "SGS", "DKSH", "DHL", "UPS", "IBM", "HP", "LG", "JVC", "BYD", "BMW",
    "GSK", "MSD", "3M", "GMP", "GLP", "GSP", "GDP", "ISO", "WHO",
    "CO2", "N2", "O2", "RO", "UV", "UVC", "BSC", "LAF", "HVAC",
}

# Các cụm từ ghép chỉ loại hình tổ chức & lĩnh vực/chức năng hoạt động (theo Nghị định 30/2020/NĐ-CP):
# Viết hoa chữ cái đầu của từ/cụm từ, các âm tiết sau trong cụm từ ghép viết thường.
ORG_AND_INDUSTRY_COMPOUNDS = [
    # Loại hình doanh nghiệp / tổ chức
    ("tổng công ty", "Tổng công ty"),
    ("công ty", "Công ty"),
    ("chi nhánh", "Chi nhánh"),
    ("tập đoàn", "Tập đoàn"),
    ("doanh nghiệp", "Doanh nghiệp"),
    ("tư nhân", "Tư nhân"),
    ("cổ phần", "Cổ phần"),
    ("trách nhiệm hữu hạn", "Trách nhiệm hữu hạn"),
    ("trách nhiệm", "Trách nhiệm"),
    ("hữu hạn", "Hữu hạn"),
    ("một thành viên", "Một thành viên"),
    ("hai thành viên", "Hai thành viên"),
    ("liên doanh", "Liên doanh"),
    ("hợp danh", "Hợp danh"),
    ("hợp tác xã", "Hợp tác xã"),
    ("nhà máy", "Nhà máy"),
    ("xí nghiệp", "Xí nghiệp"),
    ("văn phòng đại diện", "Văn phòng đại diện"),
    ("văn phòng", "Văn phòng"),
    ("trung tâm", "Trung tâm"),
    ("bệnh viện", "Bệnh viện"),
    ("đa khoa", "Đa khoa"),
    ("chuyên khoa", "Chuyên khoa"),
    ("phòng khám", "Phòng khám"),
    ("phòng thí nghiệm", "Phòng Thí nghiệm"),
    ("phòng thử nghiệm", "Phòng Thử nghiệm"),
    ("viện nghiên cứu", "Viện Nghiên cứu"),
    ("trường đại học", "Trường Đại học"),
    ("trường cao đẳng", "Trường Cao đẳng"),
    # Lĩnh vực / ngành nghề / chức năng
    ("xuất nhập khẩu", "Xuất nhập khẩu"),
    ("nước giải khát", "Nước giải khát"),
    ("bất động sản", "Bất động sản"),
    ("trang thiết bị", "Trang thiết bị"),
    ("vật tư", "Vật tư"),
    ("thiết bị", "Thiết bị"),
    ("y tế", "Y tế"),
    ("nha khoa", "Nha khoa"),
    ("thú y", "Thú y"),
    ("dược phẩm", "Dược phẩm"),
    ("mỹ phẩm", "Mỹ phẩm"),
    ("hóa mỹ phẩm", "Hóa mỹ phẩm"),
    ("hóa phẩm", "Hóa phẩm"),
    ("hóa chất", "Hóa chất"),
    ("thực phẩm", "Thực phẩm"),
    ("dinh dưỡng", "Dinh dưỡng"),
    ("giải khát", "Giải khát"),
    ("đồ uống", "Đồ uống"),
    ("kỹ nghệ", "Kỹ nghệ"),
    ("súc sản", "Súc sản"),
    ("thủy sản", "Thủy sản"),
    ("hải sản", "Hải sản"),
    ("nông sản", "Nông sản"),
    ("lâm sản", "Lâm sản"),
    ("khoáng sản", "Khoáng sản"),
    ("nông nghiệp", "Nông nghiệp"),
    ("công nghiệp", "Công nghiệp"),
    ("lâm nghiệp", "Lâm nghiệp"),
    ("thương nghiệp", "Thương nghiệp"),
    ("công nghệ", "Công nghệ"),
    ("sinh học", "Sinh học"),
    ("kỹ thuật", "Kỹ thuật"),
    ("tiêu chuẩn", "Tiêu chuẩn"),
    ("đo lường", "Đo lường"),
    ("chất lượng", "Chất lượng"),
    ("kiểm nghiệm", "Kiểm nghiệm"),
    ("giám định", "Giám định"),
    ("thử nghiệm", "Thử nghiệm"),
    ("thí nghiệm", "Thí nghiệm"),
    ("nghiên cứu", "Nghiên cứu"),
    ("ứng dụng", "Ứng dụng"),
    ("sản xuất", "Sản xuất"),
    ("chế biến", "Chế biến"),
    ("chế tạo", "Chế tạo"),
    ("gia công", "Gia công"),
    ("thương mại", "Thương mại"),
    ("dịch vụ", "Dịch vụ"),
    ("đầu tư", "Đầu tư"),
    ("phát triển", "Phát triển"),
    ("xây dựng", "Xây dựng"),
    ("kiến trúc", "Kiến trúc"),
    ("công trình", "Công trình"),
    ("hạ tầng", "Hạ tầng"),
    ("tư vấn", "Tư vấn"),
    ("thiết kế", "Thiết kế"),
    ("cơ khí", "Cơ khí"),
    ("điện tử", "Điện tử"),
    ("điện lạnh", "Điện lạnh"),
    ("viễn thông", "Viễn thông"),
    ("thông tin", "Thông tin"),
    ("truyền thông", "Truyền thông"),
    ("giải pháp", "Giải pháp"),
    ("hệ thống", "Hệ thống"),
    ("phần mềm", "Phần mềm"),
    ("năng lượng", "Năng lượng"),
    ("môi trường", "Môi trường"),
    ("bao bì", "Bao bì"),
    ("dệt may", "Dệt may"),
    ("may mặc", "May mặc"),
    ("da giày", "Da giày"),
    ("giày da", "Giày da"),
    ("nội thất", "Nội thất"),
    ("ngoại thất", "Ngoại thất"),
    ("vật liệu", "Vật liệu"),
    ("kim khí", "Kim khí"),
    ("vận tải", "Vận tải"),
    ("giao thông", "Giao thông"),
    ("tiếp vận", "Tiếp vận"),
    ("kho vận", "Kho vận"),
    ("hàng hải", "Hàng hải"),
    ("hàng không", "Hàng không"),
    ("nhập khẩu", "Nhập khẩu"),
    ("xuất khẩu", "Xuất khẩu"),
    ("phân phối", "Phân phối"),
    ("bán lẻ", "Bán lẻ"),
    ("địa ốc", "Địa ốc"),
    ("tài chính", "Tài chính"),
    ("ngân hàng", "Ngân hàng"),
    ("bảo hiểm", "Bảo hiểm"),
    ("chứng khoán", "Chứng khoán"),
    ("kiểm toán", "Kiểm toán"),
    ("kế toán", "Kế toán"),
    ("giáo dục", "Giáo dục"),
    ("đào tạo", "Đào tạo"),
    ("du lịch", "Du lịch"),
    ("lữ hành", "Lữ hành"),
    ("khách sạn", "Khách sạn"),
    ("nhà hàng", "Nhà hàng"),
    ("ẩm thực", "Ẩm thực"),
    ("quốc tế", "Quốc tế"),
    ("toàn cầu", "Toàn cầu"),
    ("thành phố hồ chí minh", "Thành phố Hồ Chí Minh"),
    ("thành phố hà nội", "Thành phố Hà Nội"),
    ("thành phố đà nẵng", "Thành phố Đà Nẵng"),
    ("thành phố hải phòng", "Thành phố Hải Phòng"),
    ("thành phố cần thơ", "Thành phố Cần Thơ"),
    ("tại tỉnh", "tại tỉnh"),
    ("tại thành phố", "tại thành phố"),
]

LOWERCASE_CONNECTORS = {"và", "tại", "ở", "của", "về", "cho", "với"}


def _is_acronym_token(clean_word):
    """Kiểm tra một token chữ cái có phải từ viết tắt (giữ IN HOA) hay không."""
    if not clean_word:
        return False
    up = clean_word.upper()
    if up in UPPERCASE_ACRONYMS:
        return True
    if any(c.isdigit() for c in up) and any(c.isalpha() for c in up):
        return True
    vowels = set("AEIOUYĂÂĐÊÔƠƯÁÀẢÃẠẮẰẲẴẶẤẦẨẪẬÉÈẺẼẸẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌỐỒỔỖỘỚỜỞỠỢÚÙỦŨỤỨỪỬỮỰÝỲỶỸỴ")
    if len(up) <= 5 and up.isalpha() and not any(ch in vowels for ch in up):
        return True
    return False


def _format_single_token(token):
    """Định dạng 1 token đơn lẻ khi chuyển từ ALL-CAPS sang chữ hoa-thường."""
    if not token:
        return token
    m = re.match(r'^([(\[{"\'“‘]*)(.*?)([)\]}"\'”’,.;:]*)$', token)
    prefix, core, suffix = m.groups() if m else ("", token, "")
    if not core:
        return token

    if "/" in core:
        segs = []
        for seg in core.split("/"):
            if not seg:
                segs.append(seg)
            elif len(seg) <= 3 or _is_acronym_token(seg):
                segs.append(seg.upper())
            else:
                segs.append(_format_single_token(seg))
        return f"{prefix}{'/'.join(segs)}{suffix}"

    if "-" in core and len(core) > 1:
        segs = [_format_single_token(seg) if seg else seg for seg in core.split("-")]
        return f"{prefix}{'-'.join(segs)}{suffix}"

    low = core.lower()
    if low in LOWERCASE_CONNECTORS:
        return f"{prefix}{low}{suffix}"

    if _is_acronym_token(core):
        return f"{prefix}{core.upper()}{suffix}"

    return f"{prefix}{core[0].upper() + core[1:].lower()}{suffix}"


def normalize_customer_name(name):
    """
    [NEW 30/10/2026] Chuẩn hóa tên khách hàng / Đơn vị yêu cầu:
    - Nếu tên khách hàng bị IN HOA TOÀN BỘ (ALL CAPS): chuyển sang chuẩn viết hoa tên cơ quan, tổ chức
      tiếng Việt (ví dụ:
        'CÔNG TY CỔ PHẦN VIỆT NAM KỸ NGHỆ SÚC SẢN' -> 'Công ty Cổ phần Việt Nam Kỹ nghệ Súc sản'
        'CÔNG TY CỔ PHẦN P/S' -> 'Công ty Cổ phần P/S'
        'CÔNG TY TNHH THỰC PHẨM CÁT TÂN' -> 'Công ty TNHH Thực phẩm Cát Tân').
    - Nếu đã có chữ thường (mixed case), bảo toàn tên riêng và chỉ chuẩn hóa tiền tố loại hình nếu tiền tố bị IN HOA.
    """
    if not name:
        return ""
    s = re.sub(r"\s+", " ", str(name)).strip()
    if not s:
        return ""

    letters = [c for c in s if c.isalpha()]
    if not letters:
        return s

    is_all_caps = all(c.isupper() for c in letters)
    if not is_all_caps:
        prefix_fixes = [
            (r"^CHI NHÁNH CÔNG TY TRÁCH NHIỆM HỮU HẠN MỘT THÀNH VIÊN\b", "Chi nhánh Công ty Trách nhiệm hữu hạn Một thành viên"),
            (r"^CHI NHÁNH CÔNG TY TNHH MỘT THÀNH VIÊN\b", "Chi nhánh Công ty TNHH Một thành viên"),
            (r"^CHI NHÁNH CÔNG TY TNHH MTV\b", "Chi nhánh Công ty TNHH MTV"),
            (r"^CHI NHÁNH CÔNG TY TNHH\b", "Chi nhánh Công ty TNHH"),
            (r"^CHI NHÁNH CÔNG TY CỔ PHẦN\b", "Chi nhánh Công ty Cổ phần"),
            (r"^CHI NHÁNH CÔNG TY CP\b", "Chi nhánh Công ty CP"),
            (r"^CÔNG TY TRÁCH NHIỆM HỮU HẠN MỘT THÀNH VIÊN\b", "Công ty Trách nhiệm hữu hạn Một thành viên"),
            (r"^CÔNG TY TNHH MỘT THÀNH VIÊN\b", "Công ty TNHH Một thành viên"),
            (r"^CÔNG TY TNHH MTV\b", "Công ty TNHH MTV"),
            (r"^CÔNG TY TNHH\b", "Công ty TNHH"),
            (r"^CÔNG TY CỔ PHẦN\b", "Công ty Cổ phần"),
            (r"^CÔNG TY CP\b", "Công ty CP"),
            (r"^TỔNG CÔNG TY\b", "Tổng công ty"),
            (r"^DOANH NGHIỆP TƯ NHÂN\b", "Doanh nghiệp Tư nhân"),
        ]
        for pat, rep in prefix_fixes:
            if re.search(pat, s):
                s = re.sub(pat, rep, s, count=1)
                break
        return s

    tokens = s.split(" ")
    n = len(tokens)
    out = []
    i = 0

    compound_map = {}
    max_len = 1
    for phrase_low, phrase_fmt in ORG_AND_INDUSTRY_COMPOUNDS:
        w_tuple = tuple(phrase_low.split())
        compound_map[w_tuple] = phrase_fmt.split()
        if len(w_tuple) > max_len:
            max_len = len(w_tuple)

    def strip_punct(tok):
        m = re.match(r'^([(\[{"\'“‘]*)(.*?)([)\]}"\'”’,.;:]*)$', tok)
        return m.groups() if m else ("", tok, "")

    while i < n:
        matched_len = 0
        matched_words = None
        for L in range(min(max_len, n - i), 1, -1):
            window = [strip_punct(tokens[i + k]) for k in range(L)]
            valid_span = True
            for k in range(L):
                pre_k, core_k, suf_k = window[k]
                if not core_k:
                    valid_span = False
                    break
                if k > 0 and pre_k:
                    valid_span = False
                    break
                if k < L - 1 and suf_k:
                    valid_span = False
                    break
            if not valid_span:
                continue
            key = tuple(w[1].lower() for w in window)
            if key in compound_map:
                matched_len = L
                fmt_parts = compound_map[key]
                matched_words = []
                for k in range(L):
                    pre_k, _, suf_k = window[k]
                    w_fmt = fmt_parts[k]
                    if i == 0 and k == 0 and w_fmt and w_fmt[0].islower():
                        w_fmt = w_fmt[0].upper() + w_fmt[1:]
                    matched_words.append(f"{pre_k}{w_fmt}{suf_k}")
                break

        if matched_len > 0:
            out.extend(matched_words)
            i += matched_len
        else:
            tok_fmt = _format_single_token(tokens[i])
            if i == 0 and tok_fmt and tok_fmt[0].islower():
                tok_fmt = tok_fmt[0].upper() + tok_fmt[1:]
            out.append(tok_fmt)
            i += 1

    return " ".join(out)


def to_title_case_vn(text):
    """Giữ nguyên chữ hoa thường nếu đã có định dạng, hoặc chuẩn hóa theo quy tắc tên tổ chức nếu là ALL CAPS."""
    if not text:
        return ""
    has_upper = any(c.isupper() for c in text)
    has_lower = any(c.islower() for c in text)
    if has_upper and has_lower:
        return text.strip()
    return normalize_customer_name(text)


def sanitize_for_filename(text):
    """Làm sạch chuỗi dùng cho tên file / thư mục Windows (VD: 'P/S' -> 'P-S', không dùng dấu '_')."""
    if not text:
        return ""
    s = re.sub(r'[\\/]+', '-', str(text))
    s = re.sub(r'[*?:"<>|]', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def extract_short_customer_name(full_name):
    """
    Rút gọn tên khách hàng để tạo tên thư mục và tên file vắn tắt (Quy tắc 9 & 10):
    - Loại bỏ tiền tố loại hình công ty: CÔNG TY TNHH, CÔNG TY CỔ PHẦN, CTY...
    - Loại bỏ cụm ngành nghề chung ở đầu: CÔNG NGHỆ SẠCH, SX TM DV... nếu phía sau còn tên riêng.
    - Ưu tiên tên viết tắt trong ngoặc đơn (nếu có, VD: Vietcombank, ABT).
    - Chuẩn hóa viết hoa dạng Title Case: 'Cát Tân', 'Greencans Industries', 'Thụy Khuê', 'K pack'.
    """
    if not full_name:
        return ""
    s = normalize_customer_name(str(full_name).strip())

    # 1. Nếu có tên viết tắt trong ngoặc đơn ở cuối (ví dụ Vietcombank, ABT)
    match_paren = re.search(r"\(([^)]+)\)$", s)
    if match_paren:
        paren_content = match_paren.group(1).strip()
        if len(paren_content) >= 2 and not re.search(r"chi\s+nhánh|địa\s+chỉ", paren_content, re.IGNORECASE):
            return paren_content

    # 2. Xử lý các phòng thử nghiệm / viện nội bộ QUATEST 3
    if "phòng thử nghiệm công nghệ sinh học" in s.lower():
        return "PTN Công nghệ Sinh học"

    # 3. Xóa các tiền tố loại hình doanh nghiệp
    for pat in COMPANY_TYPES:
        regex = re.compile(r"^" + pat + r"[\s:,-]+", re.IGNORECASE)
        if regex.search(s):
            s = regex.sub("", s).strip()
            break

    # 4. Xóa các cụm ngành nghề generic nếu phía sau còn tên thương hiệu riêng
    for pat in GENERIC_INDUSTRIES:
        regex = re.compile(r"^" + pat + r"[\s:,-]+", re.IGNORECASE)
        if regex.search(s):
            sub_res = regex.sub("", s).strip()
            if len(sub_res) >= 2:
                s = sub_res
                break

    # 5. Xóa thông tin địa danh hành chính thừa ở cuối nếu có
    s = re.sub(r"[\s,–-]+(thành\s+phố\s+hồ\s+chí\s+minh|tp\.?\s*hcm|hà\s+nội|việt\s+nam|vietnam)$", "", s, flags=re.IGNORECASE).strip()

    # Rút gọn khoảng trắng
    s = re.sub(r"\s+", " ", s).strip()
    return to_title_case_vn(s)


def extract_short_content(request_content=None, indicators=None, category_key=None, custom_short_content=None):
    """
    Trích xuất nội dung vắn tắt cho tên thư mục lưu file xuất theo Quy tắc 9:
    Quy ước chuẩn:
    - VSKK: Vi sinh không khí
    - VSBM: Vi sinh bề mặt / swab
    - 'vskk - swab': khi có cả vi sinh không khí và vi sinh bề mặt
    - 'Phòng sạch': kiểm tra phòng sạch, đếm hạt bụi, chênh áp...
    - 'Khí nén': kiểm tra khí nén, điểm sương, tổng dầu, hạt khí nén...
    - 'Khí Nitơ': kiểm tra khí nitơ
    - 'Khí nén - Phòng sạch': kết hợp cả khí nén và phòng sạch
    - 'Nước uống': nước uống, nước ăn uống (QCVN 01-1)
    - 'Nuoc sinh hoat': nước sinh hoạt
    - 'Nước thải': nước thải
    - 'MTLĐ': môi trường lao động, vi khí hậu, CO2, TVOC, bụi cá nhân...
    - 'BSC': tủ an toàn sinh học
    - 'LAF': tủ sạch
    - 'Tủ hút': tủ hút hóa chất
    - 'Kho lạnh': mapping nhiệt độ, kho lạnh, tủ bảo quản
    """
    if custom_short_content and str(custom_short_content).strip():
        return str(custom_short_content).strip()

    texts = []
    if request_content:
        texts.append(str(request_content))
    if indicators:
        for ind in indicators:
            if isinstance(ind, dict):
                texts.append(str(ind.get("name", "")))
            else:
                texts.append(str(ind))
    if category_key:
        texts.append(str(category_key))

    combined = " ".join(texts).lower()

    # Kiểm tra trường hợp hybrid kết hợp Khí nén + Phòng sạch
    has_kn = any(k in combined for k in ["khí nén", "khi nen", "tổng dầu", "điểm sương"])
    has_ps = any(k in combined for k in ["phòng sạch", "phong sach", "cleanroom", "nồng độ bụi phòng sạch"])
    if has_kn and has_ps:
        return "Khí nén - Phòng sạch"

    # 1. Vi sinh không khí (vskk) & Vi sinh bề mặt (swab / vsbm)
    has_vskk = any(k in combined for k in ["vi sinh không khí", "vskk", "vi sinh trong không khí", "air microbiology"])
    has_swab = any(k in combined for k in ["vi sinh bề mặt", "vsbm", "swab", "đĩa tiếp xúc", "tăm bông", "rodac"])

    if has_vskk and has_swab:
        return "vskk - swab"
    elif has_vskk and not any(k in combined for k in ["hạt bụi", "phòng sạch", "clean room", "cleanroom", "iso 14644"]):
        return "vskk"
    elif has_swab and not any(k in combined for k in ["hạt bụi", "phòng sạch", "clean room", "cleanroom"]):
        return "swab"

    # 2. Tủ an toàn sinh học / BSC / LAF / Tủ hút
    if any(k in combined for k in ["bsc", "an toàn sinh học", "en 12469", "inflow", "downflow"]):
        return "BSC"
    if any(k in combined for k in ["laf", "tủ sạch"]):
        return "LAF"
    if any(k in combined for k in ["tủ hút", "fume hood"]):
        return "Tủ hút"

    # 3. Môi trường nước
    if any(k in combined for k in ["nước uống", "nước ăn uống", "nuoc uong", "drinking water", "qcvn 01-1"]):
        return "Nước uống"
    if any(k in combined for k in ["nước sinh hoạt", "nuoc sinh hoat", "domestic water", "qcvn 02"]):
        return "Nuoc sinh hoat"
    if any(k in combined for k in ["nước thải", "nuoc thai", "wastewater"]):
        return "Nước thải"
    if any(k in combined for k in ["nước tinh khiết", "nước ro", "nuoc ro"]):
        return "Nước tinh khiết"

    # 4. Khí nén & Nitơ
    if any(k in combined for k in ["khí nén", "khi nen", "compressed air", "iso 8573", "điểm sương", "tổng dầu"]):
        return "Khí nén"
    if any(k in combined for k in ["khí nitơ", "khí n2", "nito", "nitrogen"]):
        return "Khí Nitơ"

    # 5. Phòng sạch
    if any(k in combined for k in ["phòng sạch", "phong sach", "cleanroom", "clean room", "iso 14644", "nồng độ hạt bụi", "hạt bụi"]):
        return "Phòng sạch"

    # 6. Kho lạnh / Mapping
    if any(k in combined for k in ["kho lạnh", "kho lanh", "mapping", "tủ bảo quản", "độ đồng đều"]):
        return "Kho lạnh"

    # 7. Môi trường lao động / MTLĐ
    if any(k in combined for k in ["mtld", "môi trường lao động", "vi khí hậu", "co2", "tvoc", "tiếng ồn"]):
        return "MTLĐ"

    # Fallback theo category_key
    cat_fallbacks = {
        "phong_sach": "Phòng sạch",
        "khi_nen": "Khí nén",
        "co2_mtld": "MTLĐ",
        "kho_lanh": "Kho lạnh",
        "bsc": "BSC",
        "tu_laf": "LAF"
    }
    return cat_fallbacks.get(category_key, "Hiện trường")


def build_output_folder_name(
    doc_number,
    customer_name,
    short_customer_name=None,
    short_content=None,
    request_content=None,
    indicators=None,
    category_key=None
):
    """
    Sinh tên thư mục lưu file xuất theo đúng nguyên tắc chuẩn (Quy tắc 9):
    'số hồ sơ - tên khách hàng - nội dung vắn tắc'
    Ví dụ:
      - Q326D124674 - Cát Tân - Phòng sạch
      - Q326D132918 - PTN Công nghệ Sinh học - BSC
      - Q326D133266 - Nextern Việt Nam - Khí nén - Phòng sạch
      - Q326D122318 - Nam Dương - Khí nén
    """
    c_short = short_customer_name.strip() if short_customer_name and short_customer_name.strip() else extract_short_customer_name(customer_name)
    c_short_safe = sanitize_for_filename(c_short)
    s_cont = short_content.strip() if short_content and short_content.strip() else extract_short_content(request_content, indicators, category_key)
    s_cont_safe = sanitize_for_filename(s_cont)

    parts = [doc_number_for_filename(str(doc_number).strip())]
    if c_short_safe:
        parts.append(c_short_safe)
    if s_cont_safe:
        parts.append(s_cont_safe)

    raw_folder = " - ".join(parts)
    clean_folder = sanitize_for_filename(raw_folder)
    return clean_folder, c_short_safe, s_cont_safe


# ==============================================================================
# 4. CHUẨN HÓA TÊN CHỈ TIÊU & LOẠI BỎ ĐỊNH LƯỢNG (Quy tắc 5)
# ==============================================================================

def clean_text_no_status_or_quantitative(text):
    """
    Loại bỏ hoàn toàn các thông tin mang tính trạng thái hoặc định lượng khỏi văn bản:
    - Trạng thái phòng sạch: trạng thái nghỉ, at-rest, hoạt động, operational
    - Phân cấp phòng sạch / Class: ISO Class 5, ISO 5, cấp sạch ...
    - Thông tin diện tích: Diện tích: 36 m2, 36 m²
    - Thông số định lượng yêu cầu / ngưỡng đánh giá: 21-24°C, +8 - 12 Pa, ≤ 55 dBA...
    - Viện dẫn tiêu chuẩn trong Nội dung yêu cầu (theo ISO..., theo QCVN...)
    """
    if not text:
        return ""
    s = str(text)

    # 1. Trạng thái phòng sạch
    s = re.sub(r"[-–—*•\t ]*Tại thời điểm thực hiện[^\n]*ở trạng thái nghỉ[^\n]*", "", s, flags=re.IGNORECASE)
    s = re.sub(r"[-–—*•\t ]*Trạng thái phòng:[^\n]*", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\(?\s*ở trạng thái nghỉ\s*\)?", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\(?\s*trạng thái nghỉ\s*\)?", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\(?\s*trạng thái hoạt động\s*\)?", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\(?\s*at-rest\s*\)?", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\(?\s*operational\s*\)?", "", s, flags=re.IGNORECASE)

    # 2. Cấp sạch ISO / Class (chỉ cấp 1-9 của phòng sạch, không xóa mã tiêu chuẩn như ISO 8573, ISO 14644)
    s = re.sub(r"\bISO\s*Class\s*[1-9]\b", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\bISO\s*[1-9]\b", "", s, flags=re.IGNORECASE)

    # 3. Diện tích
    s = re.sub(r"\(?\s*-\s*Diện tích\s*:\s*\d+(\.\d+)?\s*(m2|m²)\s*\)?", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\(?\s*Diện tích\s*:\s*\d+(\.\d+)?\s*(m2|m²)\s*\)?", "", s, flags=re.IGNORECASE)

    # 4. Định lượng trong ngoặc
    s = re.sub(r"\s*\([^)]*(?:yêu cầu|áp lực|clean zone|kích thước|≤|≥|\d+\s*°c|\d+\s*pa|\d+\s*dba|\d+\s*l/phút|vị trí)[^)]*\)", "", s, flags=re.IGNORECASE)

    # 5. Viện dẫn tiêu chuẩn trong Nội dung yêu cầu (theo ISO..., theo QCVN..., theo EN..., theo TCVN...)
    s = re.sub(r"\s+theo\s+(?:tiêu\s+chuẩn\s+)?(?:ISO|QCVN|TCVN|EN|USP|GMP|WHO)\s+[\w\d.:/-]+(?:\s*,\s*(?:ISO|QCVN|TCVN|EN)\s+[\w\d.:/-]+)*", "", s, flags=re.IGNORECASE)

    # Dọn dẹp ký tự thừa
    s = re.sub(r"\s*-\s*$", "", s)
    s = re.sub(r"\(\s*\)", "", s)
    s = re.sub(r"\s+,", ",", s)
    s = re.sub(r"\s+;", ";", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def clean_indicator_name(name, category_key=None):
    """
    Chuẩn hóa tên chỉ tiêu theo Quy tắc 5:
    1. Bắt buộc có đầy đủ: [Hành động: Đo / Lấy mẫu...] + [Tên chỉ tiêu] + [Chủ thể lấy mẫu / đo].
       Ví dụ:
       - 'Đo nồng độ bụi phòng sạch'
       - 'Đo nồng độ bụi trong khí nén'
       - 'Lấy mẫu tổng dầu trong khí nén'
       - 'Đo nhiệt độ điểm sương trong khí nén'
       - 'Lấy mẫu nước uống'
       - 'Lấy mẫu vi sinh trong không khí phòng sạch'
    2. Loại bỏ các mức đánh giá hoặc ngưỡng định lượng trong ngoặc.
    """
    if not name:
        return ""
    s = str(name).strip()
    if s.startswith("…") or s.startswith("."):
        return s

    def _strip_eval_paren(match):
        inner = match.group(1).strip().lower()
        if inner in ("inflow", "downflow"):
            return match.group(0)
        if any(k in inner for k in ["dây chuyền", "day chuyen", "phân xưởng", "phan xuong", "khu vực", "khu vuc", "line", "tủ", "tu "]):
            return match.group(0)
        return ""

    s = re.sub(r"\s*\(([^)]*)\)", _strip_eval_paren, s).strip()
    s = re.sub(r"\s+", " ", s).strip()
    s_low = s.lower()

    # --- 1. Nhóm Khí nén ---
    if any(k in s_low for k in ["tổng dầu", "hàm lượng dầu"]):
        return "Lấy mẫu tổng dầu trong khí nén"
    if "điểm sương" in s_low:
        if "khí nitơ" in s_low or "khí nito" in s_low:
            return "Đo nhiệt độ điểm sương trong khí Nitơ"
        return "Đo nhiệt độ điểm sương trong khí nén"
    if any(k in s_low for k in ["vi sinh khí nén", "vi sinh trong khí nén"]):
        return "Lấy mẫu vi sinh trong khí nén"
    if s_low in ("khí nito", "khí nitơ", "lấy mẫu khí nito", "lấy mẫu khí nitơ"):
        return "Lấy mẫu khí Nitơ"

    # --- 2. Nhóm Nước ---
    if s_low in ("nước uống", "nước ăn uống", "lấy mẫu nước uống", "lấy mẫu nước ăn uống"):
        return "Lấy mẫu nước uống"
    if s_low in ("nước sinh hoạt", "lấy mẫu nước sinh hoạt"):
        return "Lấy mẫu nước sinh hoạt"
    if s_low in ("nước thải", "lấy mẫu nước thải"):
        return "Lấy mẫu nước thải"
    if s_low in ("nước tinh khiết", "lấy mẫu nước tinh khiết"):
        return "Lấy mẫu nước tinh khiết"
    if s_low in ("nước ro", "lấy mẫu nước ro"):
        return "Lấy mẫu nước RO"
    if s_low in ("nước cấp", "lấy mẫu nước cấp"):
        return "Lấy mẫu nước cấp"
    if s_low in ("nước ngầm", "lấy mẫu nước ngầm"):
        return "Lấy mẫu nước ngầm"
    if s_low in ("nước mặt", "lấy mẫu nước mặt"):
        return "Lấy mẫu nước mặt"

    # --- 3. Nhóm Nồng độ bụi / hạt bụi ---
    if any(k in s_low for k in ["nồng độ hạt bụi", "nồng độ bụi", "đếm hạt"]):
        if "khí nén" in s_low or category_key == "khi_nen":
            return "Đo nồng độ bụi trong khí nén"
        if "khí có áp" in s_low:
            return "Đo nồng độ bụi trong khí có áp"
        if "tủ an toàn" in s_low or "bsc" in s_low or category_key == "bsc":
            return "Đo nồng độ bụi trong tủ an toàn sinh học"
        if "laf" in s_low or "tủ sạch" in s_low or category_key == "tu_laf":
            return "Đo nồng độ bụi trong tủ LAF"
        if "tủ" in s_low:
            return s if s_low.startswith("đo ") else f"Đo {s[0].lower() + s[1:]}"
        return "Đo nồng độ bụi phòng sạch"

    # --- 4. Nhóm Phòng sạch & Vi sinh ---
    if s_low in ("chênh áp", "đo chênh áp", "đo chênh áp phòng", "đo chêch áp phòng", "chênh lệch áp suất", "đo chênh lệch áp suất"):
        return "Đo chênh áp phòng sạch"
    if s_low in ("nhiệt độ, độ ẩm", "đo nhiệt độ, độ ẩm", "đo nhiệt độ và độ ẩm", "nhiệt độ và độ ẩm", "đo nhiệt độ, độ ẩm trong không khí", "nhiệt độ, độ ẩm trong không khí"):
        if category_key in ("co2_mtld", "vi_khi_hau_mtld"):
            return "Đo nhiệt độ, độ ẩm môi trường lao động"
        return "Đo nhiệt độ, độ ẩm phòng sạch"
    if s_low in ("độ ồn", "đo độ ồn", "tiếng ồn", "đo tiếng ồn", "đo độ ồn chung"):
        if category_key in ("co2_mtld", "vi_khi_hau_mtld"):
            return "Đo độ ồn môi trường lao động"
        if category_key == "bsc":
            return "Đo độ ồn tủ an toàn sinh học"
        if category_key == "tu_laf":
            return "Đo độ ồn tủ LAF"
        if category_key == "tu_hut":
            return "Đo độ ồn tủ hút"
        return "Đo độ ồn phòng sạch"
    if s_low in ("độ sáng", "đo độ sáng", "độ rọi", "đo độ rọi", "đo độ rọi / chiếu sáng", "chiếu sáng"):
        if category_key in ("co2_mtld", "vi_khi_hau_mtld"):
            return "Đo độ sáng môi trường lao động"
        return "Đo độ sáng phòng sạch"
    if "trao đổi khí" in s_low and "phòng sạch" not in s_low:
        prefix = s if s_low.startswith("đo ") else f"Đo {s[0].lower() + s[1:]}"
        return f"{prefix} phòng sạch"
    if s_low in ("vi sinh không khí", "vi sinh trong không khí", "lấy mẫu vi sinh không khí", "lấy mẫu vi sinh trong không khí"):
        if category_key == "phong_sach" or category_key is None:
            return "Lấy mẫu vi sinh trong không khí phòng sạch"
        return "Lấy mẫu vi sinh trong không khí"
    if s_low in ("vi sinh bề mặt", "lấy mẫu vi sinh bề mặt", "swab"):
        return "Lấy mẫu vi sinh bề mặt"

    # --- 5. Nhóm MTLĐ (CO2, TVOC, Vi khí hậu) ---
    if s_low in ("đo nồng độ co2 trong không khí", "nồng độ co2 trong không khí", "co2", "đo nồng độ co2"):
        return "Đo nồng độ CO2 trong không khí môi trường lao động"
    if s_low in ("đo nồng độ tvoc trong không khí", "nồng độ tvoc trong không khí", "tvoc", "đo nồng độ tvoc"):
        return "Đo nồng độ TVOC trong không khí môi trường lao động"
    if s_low in ("đo vi khí hậu", "vi khí hậu"):
        return "Đo vi khí hậu môi trường lao động"

    return s


# ==============================================================================
# 5. CHUẨN HÓA DANH SÁCH THIẾT BỊ (Quy tắc 4)
# ==============================================================================

def format_equipment_multiline(text):
    """
    Định dạng danh sách thiết bị sử dụng theo Quy tắc 4:
    - Mỗi thiết bị nằm trên một dòng riêng biệt.
    - Mỗi đầu dòng có đúng một dấu gạch ngang '- '.
    - Dòng kết thúc bằng dấu chấm phẩy ';', dòng cuối cùng kết thúc bằng dấu chấm '.'.
    """
    if not text:
        return ""
    text = str(text).replace(r"\n", "\n").replace("\r\n", "\n").replace("\r", "\n")
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    raw_items = []
    for l in lines:
        for sub in l.split(";"):
            if sub.strip():
                raw_items.append(sub.strip())

    cleaned_items = []
    for item in raw_items:
        clean = re.sub(r"^[-–—*•\t ]+", "", item).strip()
        clean = clean.rstrip(";,.")
        if clean:
            cleaned_items.append(f"- {clean};")

    if cleaned_items:
        cleaned_items[-1] = cleaned_items[-1][:-1] + "."
        return "\n".join(cleaned_items)
    return text


# ==============================================================================
# 6. CẬP NHẬT TIÊU ĐỀ CUỘC HỘI THOẠI VÀO SQLITE DB (Quy tắc 10)
# ==============================================================================

def update_conversation_db_title(doc_number, customer_name, cid=None, db_path=None):
    """
    Tự động cập nhật tiêu đề cuộc trò chuyện vào database của Antigravity (Quy tắc 10):
    Cấu trúc chuẩn: '[Mã số hồ sơ] - [Tên khách hàng]'
    """
    if not doc_number or not customer_name:
        return False

    short_name = extract_short_customer_name(customer_name)
    title = f"{doc_number} - {short_name}"

    if db_path is None:
        db_path = r"C:\Users\phat.nt\.gemini\antigravity\conversation_summaries.db"

    if not os.path.exists(db_path):
        return False

    try:
        conn = sqlite3.connect(db_path)
        cur = conn.cursor()
        if cid:
            cur.execute("UPDATE conversation_summaries SET title = ?, preview = ? WHERE conversation_id = ?", (title, title, cid))
        else:
            # Nếu không truyền cid, cập nhật record gần đây nhất
            cur.execute("UPDATE conversation_summaries SET title = ?, preview = ? WHERE rowid = (SELECT max(rowid) FROM conversation_summaries)", (title, title))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"[WARN] Không thể cập nhật tiêu đề DB: {e}")
        return False
