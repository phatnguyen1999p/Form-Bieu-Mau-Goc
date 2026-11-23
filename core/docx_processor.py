"""
Module xử lý Word docx cho hệ thống nhân bản biểu mẫu QUATEST 3
Bảo toàn 100% định dạng: Font Times New Roman cỡ 13, không in đậm, căn lề và viền bảng theo chuẩn biểu mẫu gốc.
Xuống dòng cho mỗi thiết bị sử dụng.
Loại bỏ hoàn toàn thông tin trạng thái phòng sạch và các thông số định lượng/mức đánh giá khỏi form mẫu.
"""
import copy
import re
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls


def normalize_date_str(d_str, default_year="2026"):
    """Chuẩn hóa chuỗi ngày tháng sang định dạng DD/MM/YYYY (VD: '02/10' -> '02/10/2026')."""
    if not d_str:
        return ""
    s = str(d_str).strip()
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


def clean_indicator_name(name, category_key=None):
    """
    Chuẩn hóa tên chỉ tiêu:
    1. Loại bỏ các mức đánh giá, yêu cầu hoặc ngưỡng định lượng trong ngoặc.
    2. Bắt buộc có đầy đủ: [Hành động: Đo / Lấy mẫu...] + [Tên chỉ tiêu] + [Chủ thể lấy mẫu / đo].
       Ví dụ:
       - 'Đo nồng độ bụi phòng sạch'
       - 'Đo nồng độ bụi trong khí nén'
       - 'Lấy mẫu tổng dầu trong khí nén'
       - 'Đo nhiệt độ điểm sương trong khí nén'
       - 'Lấy mẫu nước uống'
       - 'Lấy mẫu nước sinh hoạt'
    """
    if not name:
        return ""
    s = str(name).strip()
    if s.startswith("…") or s.startswith("."):
        return s

    # Loại bỏ phần mở ngoặc đơn ở cuối tên chỉ tiêu nếu chứa thông số / mức đánh giá
    # (Giữ lại các từ khóa kỹ thuật ngắn như (inflow), (downflow) nếu có)
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

    # --- 1. Chuẩn hóa nhóm Khí nén ---
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

    # --- 2. Chuẩn hóa nhóm Nước ---
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

    # --- 3. Chuẩn hóa nhóm Nồng độ bụi / hạt bụi ---
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
        # Mặc định phòng sạch
        return "Đo nồng độ bụi phòng sạch"

    # --- 4. Chuẩn hóa nhóm Phòng sạch & Vi sinh ---
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

    # --- 5. Chuẩn hóa nhóm MTLĐ (CO2, TVOC, Vi khí hậu) ---
    if s_low in ("đo nồng độ co2 trong không khí", "nồng độ co2 trong không khí", "co2", "đo nồng độ co2"):
        return "Đo nồng độ CO2 trong không khí môi trường lao động"
    if s_low in ("đo nồng độ tvoc trong không khí", "nồng độ tvoc trong không khí", "tvoc", "đo nồng độ tvoc"):
        return "Đo nồng độ TVOC trong không khí môi trường lao động"
    if s_low in ("đo vi khí hậu", "vi khí hậu"):
        return "Đo vi khí hậu môi trường lao động"

    return s



def format_equipment_multiline(text):
    """
    Định dạng danh sách thiết bị sử dụng: Mỗi thiết bị nằm trên một dòng riêng biệt.
    Mỗi đầu dòng có đúng một dấu gạch ngang '- '.
    Dòng kết thúc bằng dấu chấm phẩy ';', dòng cuối cùng kết thúc bằng dấu chấm '.'.
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
        # Xóa sạch mọi dấu gạch ngang, dấu chấm, khoảng trắng thừa ở đầu dòng
        clean = re.sub(r"^[-–—*•\t ]+", "", item).strip()
        clean = clean.rstrip(";,.")
        if clean:
            # Mỗi đầu dòng có đúng một dấu '-'
            cleaned_items.append(f"- {clean};")

    if cleaned_items:
        cleaned_items[-1] = cleaned_items[-1][:-1] + "."
        return "\n".join(cleaned_items)
    return text


def clean_text_no_status_or_quantitative(text):
    """
    Loại bỏ hoàn toàn các thông tin mang tính trạng thái hoặc định lượng khỏi văn bản:
    - Trạng thái phòng sạch: trạng thái nghỉ, at-rest, hoạt động, operational
    - Phân cấp phòng sạch / Class: ISO Class 5, ISO 5, cấp sạch ...
    - Thông tin diện tích: Diện tích: 36 m2, 36 m²
    - Thông số định lượng yêu cầu / ngưỡng đánh giá:
        + Yêu cầu: 21-24°C, 55-65%
        + Áp lực dương, yêu cầu: +8 - 12 Pa
        + Yêu cầu: ≤ 55 dBA
        + Clean zone ... 100 L/phút, 1,0 m3
        + Kích thước hạt ≥ 0,5µm
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
    s = re.sub(r"\s*\([^)]*(?:yêu cầu|áp lực|clean zone|kích thước|≤|≥|\d+\s*°c|\d+\s*pa|\d+\s*dba|\d+\s*l/phút)[^)]*\)", "", s, flags=re.IGNORECASE)

    # 5. Viện dẫn tiêu chuẩn trong Nội dung yêu cầu (theo ISO..., theo QCVN..., theo EN..., theo TCVN...)
    s = re.sub(r"\s+theo\s+(?:tiêu\s+chuẩn\s+)?(?:ISO|QCVN|TCVN|EN|USP|GMP|WHO)\s+[\w\d.:/-]+(?:\s*,\s*(?:ISO|QCVN|TCVN|EN)\s+[\w\d.:/-]+)*", "", s, flags=re.IGNORECASE)

    # Dọn dẹp ký tự thừa
    s = re.sub(r"\s*-\s*$", "", s)
    s = re.sub(r"\(\s*\)", "", s)
    s = re.sub(r"\s+,", ",", s)
    s = re.sub(r"\s+;", ";", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def normalize_address(address_str):
    """
    Chuẩn hóa địa chỉ theo quy tắc viết hoa - viết thường chuẩn hành chính Việt Nam
    và mở rộng các từ viết tắt thành nguyên chữ:
    - Mở rộng viết tắt:
      + KCN -> Khu công nghiệp
      + KCX -> Khu chế xuất
      + KDC -> Khu dân cư
      + CCN -> Cụm công nghiệp
      + KĐT -> Khu đô thị
      + KKT -> Khu kinh tế
      + KCNC -> Khu công nghệ cao
      + Tp.HCM / TP.HCM / TP. HCM / Tp. HCM -> Thành phố Hồ Chí Minh
      + QL. / QL -> Quốc lộ
      + TL. / TL -> Tỉnh lộ
      + Đ. / Đg. -> đường
      + P. -> phường (nếu tên chữ) / Phường (nếu tên số)
      + Q. -> quận (nếu tên chữ) / Quận (nếu tên số)
      + H. -> huyện
      + TX. -> thị xã
      + TT. -> thị trấn
      + X. -> xã
      + KP. -> Khu phố
      + TDP. -> Tổ dân phố
    - Loại bỏ 'Việt Nam', 'Vietnam', 'VN' ở cuối địa chỉ.
    - Thành phố trực thuộc Trung ương: viết hoa chữ 'Thành phố' (Thành phố Hồ Chí Minh, Thành phố Hà Nội, Thành phố Đà Nẵng, Thành phố Hải Phòng, Thành phố Cần Thơ, Thành phố Huế).
    - Tỉnh và Thành phố thuộc tỉnh: viết thường chữ 'tỉnh' và 'thành phố' (tỉnh Bình Dương, thành phố Thủ Dầu Một, thành phố Dĩ An...).
    - Quận / Huyện / Thị xã: viết thường nếu tên là chữ (quận Bình Thạnh, huyện Hóc Môn, thị xã Bến Cát); viết hoa nếu tên là số (Quận 1, Quận 12).
    - Phường / Xã / Thị trấn: viết thường nếu tên là chữ (phường Thới An, xã Tân Thới Hiệp, thị trấn Hóc Môn); viết hoa nếu tên là số (Phường 1, Phường 14).
    - Khu phố: viết hoa chữ 'Khu', viết thường chữ 'phố' (Khu phố 1, Khu phố Đông Chiêu).
    - Thôn / Ấp / Tổ dân phố / Tổ: viết hoa 'Thôn', 'Ấp', 'Tổ dân phố', 'Tổ'.
    - Đường / Phố / Đại lộ / Ngõ / Hẻm / Ngách: viết thường (đường Lê Thị Riêng, đường số 3, phố Tràng Tiền, đại lộ Bình Dương, hẻm 45, ngõ 12).
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
    central_cities_lower = ['hồ chí minh', 'hà nội', 'đà nẵng', 'hải phòng', 'cần thơ', 'huế', 'đồng nai']  # [11/11/2026] + Đồng Nai
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

    # 11. Khu phố, Thôn, Ấp, Tổ dân phố (mở rộng KP., TDP.):
    s = re.sub(r'\b(?:[kK]hu\s+[pP]hố|KP\.?)\s*', 'Khu phố ', s)
    s = re.sub(r'\bấp\b', 'Ấp', s)
    s = re.sub(r'\bthôn\b', 'Thôn', s)
    s = re.sub(r'\b(?:[tT]ổ\s+[dD]ân\s+[pP]hố|TDP\.?)\s*', 'Tổ dân phố ', s)
    s = re.sub(r'\btổ\s*(\d+)\b', r'Tổ \1', s)

    # 12. Đường, Phố, Đại lộ, Hẻm, Ngõ, Ngách (mở rộng Đ., Đg.):
    s = re.sub(r'\b(?:Đường|Đ\.|Đg\.)\s*[sS]ố\s*', 'đường số ', s)
    s = re.sub(r'\b(?:Đường|Đ\.|Đg\.)\s*', 'đường ', s)
    s = re.sub(r'\bPhố\s+', 'phố ', s)
    s = re.sub(r'\bĐại\s*[lL]ộ\s+', 'đại lộ ', s)
    s = re.sub(r'\bHẻm\s+', 'hẻm ', s)
    s = re.sub(r'\bNgõ\s+', 'ngõ ', s)
    s = re.sub(r'\bNgách\s+', 'ngách ', s)

    # 13. Dọn dẹp khoảng trắng và dấu phẩy
    s = re.sub(r'\s*,\s*', ', ', s)
    s = re.sub(r'\s+', ' ', s).strip(' ,.')
    return f"{prefix}{s}" if prefix else s




def set_cell_text_preserve_formatting(cell, text, bold=False, italic=None, font_name="Times New Roman", font_size=Pt(13), alignment=None):
    """
    Điền văn bản vào ô của bảng với font Times New Roman cỡ 13, không in đậm (chuẩn biểu mẫu gốc).
    Hỗ trợ xuống dòng cho mỗi thiết bị nếu văn bản chứa ký tự newline '\n'.
    Luôn set thụt lề (left_indent và first_line_indent) của tất cả các dòng về 0cm.
    """
    # Xóa các paragraph phụ nếu có
    for extra_p in list(cell.paragraphs)[1:]:
        p_elem = extra_p._p
        p_elem.getparent().remove(p_elem)
        
    p = cell.paragraphs[0]
    for r in list(p.runs):
        r_elem = r._r
        r_elem.getparent().remove(r_elem)
    
    if alignment is not None:
        p.alignment = alignment
        
    lines = str(text).split("\n") if text is not None else [""]
    final_font_name = font_name if font_name else "Times New Roman"
    target_size = font_size if font_size is not None else Pt(13)
    target_size_pt = target_size if isinstance(target_size, Pt) else Pt(target_size)

    for idx, line in enumerate(lines):
        if idx > 0:
            # Tạo paragraph mới cho mỗi dòng để xuống dòng chuẩn và đẹp trong ô bảng
            p = cell.add_paragraph()
            if alignment is not None:
                p.alignment = alignment
        
        # Xóa triệt để numbering/bullet (numPr), tab stop và ind cũ để tránh thụt lề âm hoặc lệch dòng
        if p._p.pPr is not None:
            numPr = p._p.pPr.find(qn('w:numPr'))
            if numPr is not None:
                p._p.pPr.remove(numPr)
            tabs = p._p.pPr.find(qn('w:tabs'))
            if tabs is not None:
                p._p.pPr.remove(tabs)
            ind = p._p.pPr.find(qn('w:ind'))
            if ind is not None:
                p._p.pPr.remove(ind)

        # Luôn set thụt lề về 0cm cho tất cả các dòng (left, right, first_line = 0)
        p.paragraph_format.left_indent = Inches(0)
        p.paragraph_format.right_indent = Inches(0)
        p.paragraph_format.first_line_indent = Inches(0)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        pPr = p._p.get_or_add_pPr()
        spacing = pPr.find(qn('w:spacing'))
        if spacing is None:
            pPr.append(parse_xml(f'<w:spacing {nsdecls("w")} w:before="40" w:after="40"/>'))
        else:
            spacing.set(qn('w:before'), "40")
            spacing.set(qn('w:after'), "40")
            for attr in [qn('w:beforeAutospacing'), qn('w:afterAutospacing')]:
                if attr in spacing.attrib:
                    del spacing.attrib[attr]
            
        run = p.add_run(line)
        run.font.name = final_font_name
        run._r.get_or_add_rPr().set(qn('w:rFonts'), final_font_name)
        run.font.size = target_size_pt
        run.bold = False if bold is False or bold is None else bool(bold)
        if italic is not None:
            run.italic = italic

    # Đảm bảo vertical alignment của ô là top
    tcPr = cell._tc.get_or_add_tcPr()
    vAlign = tcPr.find(qn('w:vAlign'))
    if vAlign is not None:
        vAlign.set(qn('w:val'), 'top')
    else:
        tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="top"/>'))


def extract_micro_indicators(indicators):
    """
    Trích xuất danh sách các chỉ tiêu vi sinh không khí từ hồ sơ/danh mục chỉ tiêu.
    Ví dụ:
      - 'Vi sinh không khí (Tổng số vsv hiếu khí)' -> 'Tổng số vi sinh vật hiếu khí'
      - 'Vi sinh không khí (Hiếu khí, Nấm men mốc)' -> 'Tổng số vi sinh vật hiếu khí, Tổng số nấm men – nấm mốc'
      - 'Lấy mẫu vi sinh trong không khí (Tổng số vsv hiếu khí... 100 L/phút, 1,0 m3)' -> 'Tổng số vi sinh vật hiếu khí'
    """
    if not indicators:
        return None
    
    micro_texts = []
    has_micro = False
    
    for ind in indicators:
        if isinstance(ind, dict):
            name = ind.get('name', '')
            if ind.get('micro_targets'):
                return ind['micro_targets']
        else:
            name = str(ind)
            
        if any(k in name.lower() for k in ['vi sinh', 'micro', 'khuẩn', 'nấm', 'vsv']):
            has_micro = True
            micro_texts.append(name)
            
    if not has_micro:
        return None
        
    full_text = ' '.join(micro_texts)
    
    # Kiểm tra các nhóm chỉ tiêu vi sinh thông dụng
    targets = []
    if re.search(r'hiếu\s*khí|vsv\s*hiếu\s*khí|vi\s*sinh\s*vật\s*hiếu\s*khí|\btpc\b', full_text, re.I):
        targets.append('Tổng số vi sinh vật hiếu khí')
    if re.search(r'nấm\s*men|nấm\s*mốc|men\s*mốc', full_text, re.I):
        targets.append('Tổng số nấm men – nấm mốc')
    if re.search(r'coliform', full_text, re.I):
        targets.append('Coliforms')
    if re.search(r'staphylococcus|tụ\s*cầu', full_text, re.I):
        targets.append('Staphylococcus aureus')
    if re.search(r'pseudomonas|mủ\s*xanh', full_text, re.I):
        targets.append('Pseudomonas aeruginosa')
    if re.search(r'e\.?\s*coli|escherichia', full_text, re.I):
        targets.append('Escherichia coli')
        
    if targets:
        return ', '.join(targets)
        
    # Nếu không nhận diện được từ khóa cụ thể nhưng có chỉ tiêu vi sinh
    # Kiểm tra xem có nội dung trong dấu ngoặc đơn không: (...)
    m = re.search(r'\((.*?)\)', full_text)
    if m:
        content = m.group(1).split('...')[0].split('-')[0].strip()
        # Loại bỏ các thông số kỹ thuật như 100 L/phút, 1 m3 nếu có
        content = re.sub(r'\d+\s*(L/phút|lít|m3|m³).*', '', content, flags=re.I).strip()
        if content:
            return content
            
    return 'Tổng số vi sinh vật hiếu khí'


def standardize_thoi_gian_row_in_table(tbl, time_date):
    """
    Tìm và chuẩn hóa hàng 'Thời gian thực hiện' trong Table 1:
    - Tìm hàng mẫu (ref_row) có cột giữa là ':' (ví dụ 'Địa điểm thực hiện', 'Đơn vị yêu cầu'...)
    - Đồng bộ hoàn toàn format bảng:
        + Cột 0: 'Thời gian thực hiện' với numPr kế thừa từ bảng để tự động đánh số thứ tự
        + Cột 1: ':' (cột giữa, căn giữa)
        + Cột 2 (hoặc cột nội dung): 'Bắt đầu từ.................... kết thúc lúc..................... ngày {time_date}'
    """
    ref_row = None
    target_row_idx = None
    
    for r_idx, row in enumerate(tbl.rows):
        row_txt = " ".join([c.text.strip() for c in row.cells])
        if "thời gian thực hiện" in row_txt.lower():
            target_row_idx = r_idx
        elif ref_row is None and len(row.cells) >= 2:
            # Tìm hàng mẫu có dấu ':' ở cell 1
            if ("địa điểm thực hiện" in row_txt.lower() or 
                "đơn vị yêu cầu" in row_txt.lower() or 
                "địa chỉ" in row_txt.lower() or
                "thiết bị sử dụng" in row_txt.lower() or
                "quy trình thực hiện" in row_txt.lower()):
                if len(row.cells) > 1 and row.cells[1].text.strip() == ":":
                    ref_row = row

    if target_row_idx is None:
        return False
        
    target_row = tbl.rows[target_row_idx]
    tr = target_row._tr
    tcs = tr.findall(qn('w:tc'))
    if not tcs:
        return False

    # Lấy thông số từ ref_row nếu có
    ref_tc0_w = "2850"
    ref_tc1_w = "377"
    ref_numPr = None
    ref_pPr_tc0 = None
    
    if ref_row is not None:
        ref_tcs = ref_row._tr.findall(qn('w:tc'))
        if len(ref_tcs) >= 2:
            w0 = ref_tcs[0].find(qn('w:tcPr'))
            if w0 is not None and w0.find(qn('w:tcW')) is not None:
                ref_tc0_w = w0.find(qn('w:tcW')).attrib.get(qn('w:w'), "2850")
            w1 = ref_tcs[1].find(qn('w:tcPr'))
            if w1 is not None and w1.find(qn('w:tcW')) is not None:
                ref_tc1_w = w1.find(qn('w:tcW')).attrib.get(qn('w:w'), "377")
                
            p0 = ref_tcs[0].find(qn('w:p'))
            if p0 is not None and p0.find(qn('w:pPr')) is not None:
                ref_pPr_tc0 = copy.deepcopy(p0.find(qn('w:pPr')))
                ref_numPr = ref_pPr_tc0.find(qn('w:numPr'))

    tc0 = tcs[0]
    tcPr0 = tc0.get_or_add_tcPr()
    gridSpan0 = tcPr0.find(qn('w:gridSpan'))
    gs0_val = int(gridSpan0.attrib.get(qn('w:val'), 1)) if gridSpan0 is not None else 1

    if gs0_val >= 2:
        # Trường hợp BM3030.02: Cột 0 bị merge 2 cột
        # 1. Thu hồi gridSpan về 1
        if gridSpan0 is not None:
            tcPr0.remove(gridSpan0)
        w0 = tcPr0.find(qn('w:tcW'))
        if w0 is not None:
            w0.set(qn('w:w'), ref_tc0_w)
        else:
            tcPr0.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{ref_tc0_w}" w:type="dxa"/>'))

        # Format p0 của tc0
        p0 = tc0.find(qn('w:p'))
        if p0 is None:
            p0 = parse_xml(f'<w:p {nsdecls("w")}/>')
            tc0.append(p0)
        for child in list(p0):
            p0.remove(child)
            
        # Thêm pPr chuẩn
        if ref_pPr_tc0 is not None:
            p0.append(copy.deepcopy(ref_pPr_tc0))
        else:
            pPr_xml = (
                f'<w:pPr {nsdecls("w")}>'
                f'  <w:numPr><w:ilvl w:val="0"/><w:numId w:val="29"/></w:numPr>'
                f'  <w:spacing w:before="40" w:after="40" w:line="240" w:lineRule="auto"/>'
                f'  <w:ind w:left="284" w:hanging="284"/>'
                f'  <w:jc w:val="both"/>'
                f'  <w:rPr><w:szCs w:val="26"/><w:lang w:val="pt-BR"/></w:rPr>'
                f'</w:pPr>'
            )
            p0.append(parse_xml(pPr_xml))
            
        run_xml = (
            f'<w:r {nsdecls("w")}>'
            f'  <w:rPr>'
            f'    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
            f'    <w:sz w:val="26"/><w:szCs w:val="26"/>'
            f'  </w:rPr>'
            f'  <w:t>Thời gian thực hiện</w:t>'
            f'</w:r>'
        )
        p0.append(parse_xml(run_xml))

        # 2. Tạo tc1 cho dấu ':'
        tc_colon_xml = (
            f'<w:tc {nsdecls("w")}>'
            f'  <w:tcPr>'
            f'    <w:tcW w:w="{ref_tc1_w}" w:type="dxa"/>'
            f'    <w:vAlign w:val="top"/>'
            f'  </w:tcPr>'
            f'  <w:p>'
            f'    <w:pPr>'
            f'      <w:spacing w:before="40" w:after="40" w:line="240" w:lineRule="auto"/>'
            f'      <w:jc w:val="center"/>'
            f'      <w:rPr><w:szCs w:val="26"/><w:lang w:val="pt-BR"/></w:rPr>'
            f'    </w:pPr>'
            f'    <w:r>'
            f'      <w:rPr>'
            f'        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
            f'        <w:sz w:val="26"/><w:szCs w:val="26"/>'
            f'      </w:rPr>'
            f'      <w:t>:</w:t>'
            f'    </w:r>'
            f'  </w:p>'
            f'</w:tc>'
        )
        tc_colon = parse_xml(tc_colon_xml)
        tr.insert(tr.index(tc0) + 1, tc_colon)

        # 3. Ô nội dung là tcs[1] ban đầu (nay là ô thứ 3)
        tc_content = tcs[1]

    else:
        # Trường hợp BM3030.01: Đã có đủ cột riêng biệt
        p0 = tc0.find(qn('w:p'))
        if p0 is not None:
            pPr0 = p0.find(qn('w:pPr'))
            if pPr0 is None:
                pPr0 = parse_xml(f'<w:pPr {nsdecls("w")}/>')
                p0.insert(0, pPr0)
            if pPr0.find(qn('w:numPr')) is None:
                if ref_numPr is not None:
                    pPr0.insert(0, copy.deepcopy(ref_numPr))
                else:
                    pPr0.insert(0, parse_xml(f'<w:numPr {nsdecls("w")}><w:ilvl w:val="0"/><w:numId w:val="29"/></w:numPr>'))
            # Đảm bảo nội dung text là 'Thời gian thực hiện'
            runs = p0.findall(qn('w:r'))
            if runs:
                t0 = runs[0].find(qn('w:t'))
                if t0 is not None:
                    t0.text = "Thời gian thực hiện"
                for extra_r in runs[1:]:
                    p0.remove(extra_r)
            else:
                p0.append(parse_xml(
                    f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/></w:rPr><w:t>Thời gian thực hiện</w:t></w:r>'
                ))

        # Cột 1 là dấu ':'
        if len(tcs) > 1:
            tc_col = tcs[1]
            p_col = tc_col.find(qn('w:p'))
            if p_col is not None:
                pPr_col = p_col.find(qn('w:pPr'))
                if pPr_col is None:
                    pPr_col = parse_xml(f'<w:pPr {nsdecls("w")}/>')
                    p_col.insert(0, pPr_col)
                jc = pPr_col.find(qn('w:jc'))
                if jc is None:
                    pPr_col.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
                else:
                    jc.set(qn('w:val'), "center")
                runs = p_col.findall(qn('w:r'))
                if runs:
                    t = runs[0].find(qn('w:t'))
                    if t is not None:
                        t.text = ":"
                    for extra_r in runs[1:]:
                        p_col.remove(extra_r)
                else:
                    p_col.append(parse_xml(
                        f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/></w:rPr><w:t>:</w:t></w:r>'
                    ))

        # Cột 2 là nội dung
        tc_content = tcs[2] if len(tcs) > 2 else tcs[-1]

    # Cập nhật ô nội dung
    p_cnt = tc_content.find(qn('w:p'))
    if p_cnt is None:
        p_cnt = parse_xml(f'<w:p {nsdecls("w")}/>')
        tc_content.append(p_cnt)
    for r in p_cnt.findall(qn('w:r')):
        p_cnt.remove(r)
        
    pPr_cnt = p_cnt.find(qn('w:pPr'))
    if pPr_cnt is not None:
        nPr = pPr_cnt.find(qn('w:numPr'))
        if nPr is not None:
            pPr_cnt.remove(nPr)
        ind = pPr_cnt.find(qn('w:ind'))
        if ind is not None:
            pPr_cnt.remove(ind)
        tabs = pPr_cnt.find(qn('w:tabs'))
        if tabs is not None:
            pPr_cnt.remove(tabs)
    else:
        pPr_cnt = parse_xml(f'<w:pPr {nsdecls("w")}/>')
        p_cnt.insert(0, pPr_cnt)
    pPr_cnt.append(parse_xml(f'<w:ind {nsdecls("w")} w:left="0" w:right="0" w:firstLine="0"/>'))

    updated_time_str = f"Bắt đầu từ.................... kết thúc lúc..................... ngày {time_date}"
    r_cnt = parse_xml(
        f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>{updated_time_str}</w:t></w:r>'
    )
    p_cnt.append(r_cnt)
    return True


def clean_spaces_before_colon(doc):
    """
    Loại bỏ mọi dấu cách trước dấu ':' trong toàn bộ văn bản và các bảng:
    - 'Số : ' -> 'Số: '
    - 'Thời gian thực hiện :' -> 'Thời gian thực hiện:'
    - Ô nhãn có khoảng trắng cuối dòng trước cột ':' (ví dụ 'Đơn vị yêu cầu ' -> 'Đơn vị yêu cầu')
    - Các đoạn văn bản có ' :' -> ':'
    """
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if " :" in p.text:
                        for r in p.runs:
                            if " :" in r.text:
                                r.text = r.text.replace(" :", ":")
                    # Nếu là ô đề mục/nhãn có khoảng trắng thừa ở cuối (VD 'Đơn vị yêu cầu ')
                    if p.text.endswith(" ") and p.text.strip() != "":
                        for r in reversed(p.runs):
                            if r.text.endswith(" "):
                                r.text = r.text.rstrip(" ")
                                if not r.text.endswith(" "):
                                    break
    for p in doc.paragraphs:
        if " :" in p.text:
            for r in p.runs:
                if " :" in r.text:
                    r.text = r.text.replace(" :", ":")


def is_bm01_measurement_table(tbl):
    """
    Xác định bảng ghi chép số liệu đo hiện trường của BM3030.01
    (các bảng này đã có định dạng chuẩn mẫu gốc với checkbox, canh giữa số thứ tự, tab stop).
    """
    if len(tbl.rows) == 0:
        return False
    all_txt = " ".join(c.text.strip() for r in tbl.rows for c in r.cells)
    keywords = [
        "Ký hiệu vị trí đo",
        "Particle",
        "Tên khu vực thực hiện",
        "File ghi",
        "Kết quả đo (Pa)",
        "Khu vực thực hiện/ ký hiệu VT",
        "Tên vị trí lấy mẫu",
        "Ký hiệu đĩa"
    ]
    return any(kw in all_txt for kw in keywords)


def set_all_tables_valign_top(doc):
    """
    Format vertical alignment (căn lề trên) = Top cho tất cả các cell của các bảng thông tin,
    nhưng giữ nguyên format chuẩn của các bảng ghi chép hiện trường BM3030.01.
    """
    for tbl in doc.tables:
        if is_bm01_measurement_table(tbl):
            continue
        for row in tbl.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP


def set_all_tables_spacing_top_down(doc, pt_val=2):
    """
    Format paragraph spacing = 2 pt cho top (before) và down (after)
    cho các bảng thông tin (Table 0, Table 1 và bảng chỉ tiêu BM3030.02),
    giữ nguyên spacing gốc của các bảng ghi chép hiện trường BM3030.01.
    """
    dxa_val = str(int(pt_val * 20))
    for tbl in doc.tables:
        if is_bm01_measurement_table(tbl):
            continue
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    pPr = p._p.get_or_add_pPr()
                    spacing = pPr.find(qn('w:spacing'))
                    if spacing is None:
                        spacing = parse_xml(f'<w:spacing {nsdecls("w")} w:before="{dxa_val}" w:after="{dxa_val}"/>')
                        pPr.append(spacing)
                    else:
                        spacing.set(qn('w:before'), dxa_val)
                        spacing.set(qn('w:after'), dxa_val)
                        for attr in [qn('w:beforeAutospacing'), qn('w:afterAutospacing')]:
                            if attr in spacing.attrib:
                                del spacing.attrib[attr]
                    p.paragraph_format.space_before = Pt(pt_val)
                    p.paragraph_format.space_after = Pt(pt_val)


def update_header_hl(doc, target_hl="HL:01/08/2025"):
    """
    Sửa trong header cái HL là 'HL:01/08/2025' cho tất cả hồ sơ:
    Quét qua tất cả sections, header paragraphs và header tables.
    """
    for sec in doc.sections:
        # Header tables
        for tbl in sec.header.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if "hl" in p.text.lower():
                            if re.search(r'HL\s*:\s*[\d/]+', p.text, re.IGNORECASE) or p.text.strip().startswith("HL:"):
                                runs = p.runs
                                if runs:
                                    runs[0].text = target_hl
                                    for extra_r in runs[1:]:
                                        p._p.remove(extra_r._r)
                                else:
                                    p.text = target_hl

        # Header paragraphs
        for p in sec.header.paragraphs:
            if "hl" in p.text.lower():
                if re.search(r'HL\s*:\s*[\d/]+', p.text, re.IGNORECASE) or p.text.strip().startswith("HL:"):
                    runs = p.runs
                    if runs:
                        runs[0].text = target_hl
                        for extra_r in runs[1:]:
                            p._p.remove(extra_r._r)
                    else:
                        p.text = target_hl


def ensure_6pt_spacing_after_all_tables(doc):
    """
    Dưới mỗi bảng cách dòng và dòng đó cho font size là 6:
    Chèn hoặc duy trì duy nhất một paragraph trống ngay dưới mỗi bảng với font size 6pt (w:sz = 12),
    spacing before/after = 0pt, line spacing = exact 6pt (120 dxa).
    Tự động dọn dẹp các empty paragraph thừa/trùng lặp ngay sau bảng để tránh phình trang.
    """
    for tbl in doc.tables:
        nxt = tbl._tbl.getnext()
        # Duyệt và dọn dẹp các paragraph trống liên tiếp ngay sau bảng
        empty_paras = []
        curr = nxt
        while (
            curr is not None 
            and curr.tag.endswith('p') 
            and "".join(curr.itertext()).strip() == ""
            and curr.find('.//' + qn('w:pageBreakBefore')) is None
        ):
            empty_paras.append(curr)
            curr = curr.getnext()

        # [FIX 05/11/2026] Nếu phần tử ngay sau (hoặc sau các paragraph rỗng) có pageBreakBefore:
        # Xóa sạch toàn bộ paragraph rỗng thừa và TUYỆT ĐỐI KHÔNG chèn 6pt spacer trước ngắt trang,
        # để tránh tạo trang trắng rỗng thừa giữa bảng và trang mới (Quy tắc 3 & Quy tắc 7).
        has_pb_after = (curr is not None and curr.find('.//' + qn('w:pageBreakBefore')) is not None)
        if has_pb_after:
            for p_extra in empty_paras:
                p_extra.getparent().remove(p_extra)
            continue

        if empty_paras:
            # Giữ lại paragraph đầu tiên và format chuẩn 6pt
            p_keep = empty_paras[0]
            # Xóa các empty paragraph thừa phía sau
            for p_extra in empty_paras[1:]:
                p_extra.getparent().remove(p_extra)
                
            pPr = p_keep.get_or_add_pPr()
            sp = pPr.find(qn('w:spacing'))
            if sp is not None:
                pPr.remove(sp)
            pPr.append(parse_xml(f'<w:spacing {nsdecls("w")} w:before="0" w:after="0" w:line="120" w:lineRule="exact"/>'))
            
            rPr = pPr.find(qn('w:rPr'))
            if rPr is None:
                rPr = parse_xml(f'<w:rPr {nsdecls("w")}><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="12"/><w:szCs w:val="12"/></w:rPr>')
                pPr.append(rPr)
            else:
                sz = rPr.find(qn('w:sz'))
                if sz is None: rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="12"/>'))
                else: sz.set(qn('w:val'), '12')
                szCs = rPr.find(qn('w:szCs'))
                if szCs is None: rPr.append(parse_xml(f'<w:szCs {nsdecls("w")} w:val="12"/>'))
                else: szCs.set(qn('w:val'), '12')
                
            # Đảm bảo các run bên trong cũng 6pt
            for r in p_keep.findall(qn('w:r')):
                r_rPr = r.find(qn('w:rPr'))
                if r_rPr is None:
                    r_rPr = parse_xml(f'<w:rPr {nsdecls("w")}><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="12"/><w:szCs w:val="12"/></w:rPr>')
                    r.append(r_rPr)
                else:
                    sz = r_rPr.find(qn('w:sz'))
                    if sz is None: r_rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="12"/>'))
                    else: sz.set(qn('w:val'), '12')
                    szCs = r_rPr.find(qn('w:szCs'))
                    if szCs is None: r_rPr.append(parse_xml(f'<w:szCs {nsdecls("w")} w:val="12"/>'))
                    else: szCs.set(qn('w:val'), '12')
        else:
            spacer_xml = f'''
            <w:p {nsdecls("w")}>
                <w:pPr>
                    <w:spacing w:before="0" w:after="0" w:line="120" w:lineRule="exact"/>
                    <w:rPr>
                        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
                        <w:sz w:val="12"/>
                        <w:szCs w:val="12"/>
                    </w:rPr>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
                        <w:sz w:val="12"/>
                        <w:szCs w:val="12"/>
                    </w:rPr>
                    <w:t></w:t>
                </w:r>
            </w:p>
            '''
            p_elem = parse_xml(spacer_xml)
            tbl._tbl.addnext(p_elem)


def ensure_all_rows_cant_split(doc):
    """Áp dụng thuộc tính w:cantSplit cho tất cả các hàng của tất cả các bảng trong tài liệu."""
    for tbl in doc.tables:
        for row in tbl.rows:
            trPr = row._tr.get_or_add_trPr()
            if trPr.find(qn('w:cantSplit')) is None:
                trPr.append(OxmlElement('w:cantSplit'))


def finalize_document_formatting(doc):
    """
    Bước chuẩn hóa cuối cùng cho tài liệu:
    1. Bỏ dấu cách trước dấu ':' (trong bảng và văn bản).
    2. Format cell vertical alignment = Top cho tất cả các bảng.
    3. Chọn spacing các phần tử trong bảng là 2 pt cho top và down (before=2pt, after=2pt).
    4. Hàng 1 và cột 1 trong bảng không cần ghi số hồ sơ (để trống).
    5. Sửa trong header cái HL là 'HL:01/08/2025' cho tất cả hồ sơ.
    6. Dưới mỗi bảng cách dòng và dòng đó cho font size là 6pt.
    7. Áp dụng w:cantSplit cho tất cả các hàng của bảng.
    """
    clean_spaces_before_colon(doc)
    set_all_tables_valign_top(doc)
    
    # Hàng 1 và cột 1 trong bảng không ghi số hồ sơ (để trống)
    for tbl in doc.tables:
        if is_bm01_measurement_table(tbl):
            continue
        if len(tbl.rows) > 0 and len(tbl.rows[0].cells) > 0:
            row0 = tbl.rows[0]
            if not any(c.text.strip().startswith("Số") for c in row0.cells):
                c0 = row0.cells[0]
                if "Q326" in c0.text or "QONV" in c0.text or c0.text.strip().isdigit():
                    set_cell_text_preserve_formatting(c0, "", bold=False, font_size=Pt(13))

    set_all_tables_spacing_top_down(doc, pt_val=2)
    ensure_all_rows_cant_split(doc)
    update_header_hl(doc, "HL:01/08/2025")
    ensure_6pt_spacing_after_all_tables(doc)


def update_customer_info(doc, customer_info, indicators=None):
    """
    Cập nhật thông tin khách hàng và thông tin chung vào tài liệu docx.
    Tất cả sử dụng font Times New Roman cỡ 13, không in đậm.
    Xuống dòng cho mỗi thiết bị sử dụng.
    Loại bỏ thông tin trạng thái phòng và định lượng.
    Điều chỉnh indent left về 0, right về 0cm cho 'Thời gian thực hiện'.
    Cập nhật chỉ tiêu vi sinh không khí theo nội dung hồ sơ.
    """
    from .normalizer import normalize_customer_name, LOCATION_DOTS
    customer_name = normalize_customer_name(clean_text_no_status_or_quantitative(customer_info.get("customer_name", "")))
    address = normalize_address(clean_text_no_status_or_quantitative(customer_info.get("address", "")))
    raw_loc = clean_text_no_status_or_quantitative(customer_info.get("location", ""))
    dots_prefix = LOCATION_DOTS
    if raw_loc:
        m = re.match(r"^[\s.…\n]+\s*-\s*(.*)$", raw_loc)
        if m:
            addr_part = m.group(1).strip()
        else:
            addr_part = raw_loc.strip()
        norm_addr = normalize_address(addr_part)
        location = f"{dots_prefix} - {norm_addr}"
    elif address:
        location = f"{dots_prefix} - {address}"
    else:
        location = ""
    doc_number = customer_info.get("doc_number", "").strip()
    date_val = normalize_date_str(customer_info.get("date", "").strip())
    reg_date = normalize_date_str(customer_info.get("registration_date") or customer_info.get("reg_date") or date_val)
    current_year = reg_date.split("/")[-1] if "/" in reg_date else "2026"
    default_dotted_date = f"…./…./{current_year}"

    explicit_exec = customer_info.get("execution_date") or customer_info.get("exec_date") or ""
    # Nếu execution_date là dấu chấm hoặc rỗng hoặc trùng reg_date mà không có lịch hiện trường rõ ràng
    if explicit_exec and not any(k in explicit_exec for k in ["…", "..."]):
        exec_date = normalize_date_str(explicit_exec)
    else:
        exec_date = default_dotted_date if not explicit_exec else explicit_exec

    header_date = exec_date if exec_date else default_dotted_date
    req_date = reg_date if reg_date else date_val
    time_date = exec_date if exec_date else default_dotted_date

    request_content = clean_text_no_status_or_quantitative(customer_info.get("request_content", ""))
    equipment_custom = customer_info.get("equipment", "")
    all_indicators = indicators if indicators is not None else customer_info.get("indicators")

    # 1. Cập nhật trong các bảng thông tin chung (Table 0 và Table 1)
    for tbl_idx, tbl in enumerate(doc.tables):
        if len(tbl.rows) == 0:
            continue
        all_tbl_txt = " ".join(c.text.strip() for r in tbl.rows for c in r.cells)
        # Chỉ xử lý Table 0 (header) hoặc bảng chứa "Đơn vị yêu cầu" / "Phiếu yêu cầu"
        is_info_tbl = ("Đơn vị yêu cầu" in all_tbl_txt or "Phiếu yêu cầu" in all_tbl_txt or (tbl_idx == 0 and any("Số" in c.text for c in tbl.rows[0].cells)))
        if not is_info_tbl:
            continue

        # Chuẩn hóa cấu trúc và số thứ tự tự động cho hàng Thời gian thực hiện
        standardize_thoi_gian_row_in_table(tbl, time_date)

        ri = 0
        num_rows = len(tbl.rows)
        while ri < num_rows:
            row = tbl.rows[ri]
            row_full_text = " ".join(c.text.strip() for c in row.cells)

            # --- Cập nhật Bảng Header Table 0 (Số: ..., Ngày: ...) ---
            if len(row.cells) >= 3 and any(c.text.strip().startswith("Số") for c in row.cells):
                for ci in range(len(row.cells)):
                    c_txt = row.cells[ci].text.strip()
                    # Giữ cố định ngày hiệu lực HL: của form biểu mẫu, tuyệt đối không thay đổi
                    if "HL:" in c_txt or "Hiệu lực" in c_txt:
                        continue
                    if ("Q326" in c_txt or "QONV" in c_txt or "……" in c_txt or c_txt == "") and doc_number:
                        if ci > 0 and "Số" in row.cells[ci-1].text:
                            set_cell_text_preserve_formatting(row.cells[ci], doc_number, bold=False, font_size=Pt(13))
                    if c_txt.startswith("Ngày:") or c_txt == "Ngày:":
                        if ci + 1 < len(row.cells) and header_date:
                            set_cell_text_preserve_formatting(row.cells[ci+1], header_date, bold=False, font_size=Pt(13))
                        elif header_date:
                            set_cell_text_preserve_formatting(row.cells[ci], f"Ngày: {header_date}", bold=False, font_size=Pt(13))

            elif ri == 0 and len(row.cells) >= 1 and not any(c.text.strip().startswith("Số") for c in row.cells) and "Đơn vị yêu cầu" not in row_full_text:
                # Hàng 1 và cột 1 trong bảng không cần ghi số hồ sơ (để trống)
                set_cell_text_preserve_formatting(row.cells[0], "", bold=False, font_size=Pt(13))

            # --- Cập nhật Phiếu yêu cầu & Nhận TĐ/GĐ (Sử dụng ngày đăng ký) ---
            elif "Phiếu yêu cầu" in row_full_text:
                for ci in range(len(row.cells)):
                    c_txt = row.cells[ci].text.strip()
                    # Giữ cố định ngày hiệu lực HL: của form biểu mẫu
                    if "HL:" in c_txt or "Hiệu lực" in c_txt:
                        continue
                    if ("Q326" in c_txt or "QONV" in c_txt or "……" in c_txt or c_txt.isdigit()) and doc_number:
                        set_cell_text_preserve_formatting(row.cells[ci], doc_number, bold=False, font_size=Pt(13))
                    elif c_txt.startswith("Ngày:") or c_txt == "Ngày:":
                        if ci + 1 < len(row.cells) and req_date:
                            set_cell_text_preserve_formatting(row.cells[ci+1], req_date, bold=False, font_size=Pt(13))
                    elif ("202" in c_txt or "……/……" in c_txt) and "Phiếu yêu cầu" not in c_txt and req_date:
                        set_cell_text_preserve_formatting(row.cells[ci], req_date, bold=False, font_size=Pt(13))

            # --- Cập nhật Đơn vị yêu cầu ---
            elif "Đơn vị yêu cầu" in row_full_text:
                for ci in range(len(row.cells)):
                    c_txt = row.cells[ci].text.strip()
                    if "Đơn vị yêu cầu" not in c_txt and c_txt != ":" and customer_name:
                        set_cell_text_preserve_formatting(row.cells[ci], customer_name, bold=False, font_size=Pt(13))
                        break

            # --- Cập nhật Địa chỉ ---
            elif "Địa chỉ" in row_full_text and "Địa điểm" not in row_full_text:
                for ci in range(len(row.cells)):
                    c_txt = row.cells[ci].text.strip()
                    if "Địa chỉ" not in c_txt and c_txt != ":" and address:
                        set_cell_text_preserve_formatting(row.cells[ci], address, bold=False, font_size=Pt(13))
                        break

            # --- Cập nhật Địa điểm thực hiện ---
            elif "Địa điểm thực hiện" in row_full_text:
                for ci in range(len(row.cells)):
                    c_txt = row.cells[ci].text.strip()
                    if "Địa điểm thực hiện" not in c_txt and c_txt != ":" and location:
                        set_cell_text_preserve_formatting(row.cells[ci], location, bold=False, font_size=Pt(13))
                        break

            # --- Cập nhật Thời gian thực hiện (đã được chuẩn hóa tự động bởi standardize_thoi_gian_row_in_table) ---
            elif "Thời gian thực hiện" in row_full_text:
                pass

            # --- Cập nhật Nội dung yêu cầu ---
            elif "Nội dung yêu cầu" in row_full_text and "Nội dung/ Quá trình" not in row_full_text and "Nội dung/Quá trình" not in row_full_text:
                if request_content:
                    # Gán request_content vào ô nội dung (cột 3 / cell 2) trên cùng hàng với "Nội dung yêu cầu :"
                    target_cell = None
                    if len(row.cells) >= 3:
                        target_cell = row.cells[2]
                    else:
                        for ci in range(len(row.cells)):
                            c_txt = row.cells[ci].text.strip()
                            if "Nội dung yêu cầu" not in c_txt and c_txt != ":":
                                target_cell = row.cells[ci]
                                break
                    if target_cell is not None:
                        set_cell_text_preserve_formatting(target_cell, request_content, bold=False, font_size=Pt(13))

                    # Nếu hàng tiếp theo là hàng phụ của template cũ (chứa nội dung yêu cầu cũ, không phải là Địa điểm thực hiện)
                    if ri + 1 < len(tbl.rows):
                        next_row = tbl.rows[ri + 1]
                        next_c0 = next_row.cells[0].text.strip()
                        if "Địa điểm" not in next_c0:
                            tr = next_row._tr
                            tr.getparent().remove(tr)
                            num_rows -= 1

            # --- Cập nhật Thiết bị sử dụng (Xuống dòng cho mỗi thiết bị) ---
            elif "Thiết bị sử dụng" in row_full_text and "Quy trình" not in row_full_text:
                eq_to_use = equipment_custom if equipment_custom else ""
                has_same_row_eq = False
                for ci in range(len(row.cells)):
                    c_txt = row.cells[ci].text.strip()
                    if "Thiết bị sử dụng" not in c_txt and c_txt != ":" and len(c_txt) > 0:
                        if not eq_to_use:
                            eq_to_use = c_txt
                        formatted_eq = format_equipment_multiline(eq_to_use)
                        set_cell_text_preserve_formatting(row.cells[ci], formatted_eq, bold=False, font_size=Pt(13))
                        has_same_row_eq = True
                        break
                if not has_same_row_eq and ri + 1 < num_rows:
                    next_row = tbl.rows[ri + 1]
                    if not eq_to_use:
                        eq_to_use = next_row.cells[0].text.strip()
                    formatted_eq = format_equipment_multiline(eq_to_use)
                    set_cell_text_preserve_formatting(next_row.cells[0], formatted_eq, bold=False, font_size=Pt(13))
                    ri += 1

            # --- Cập nhật Quy trình thực hiện ---
            elif "Quy trình thực hiện" in row_full_text or "Quy trình áp dụng" in row_full_text:
                proc_to_use = customer_info.get("procedure") or customer_info.get("standard_procedure")
                if proc_to_use:
                    has_same_row_proc = False
                    for ci in range(len(row.cells)):
                        c_txt = row.cells[ci].text.strip()
                        if "Quy trình" not in c_txt and c_txt != ":" and len(c_txt) > 0:
                            set_cell_text_preserve_formatting(row.cells[ci], proc_to_use, bold=False, font_size=Pt(13))
                            has_same_row_proc = True
                            break
                    if not has_same_row_proc and ri + 1 < num_rows:
                        next_row = tbl.rows[ri + 1]
                        set_cell_text_preserve_formatting(next_row.cells[0], proc_to_use, bold=False, font_size=Pt(13))
                        ri += 1

            ri += 1

    # 2. Cập nhật trong các đoạn văn bản (Paragraphs)
    for p in doc.paragraphs:
        txt = p.text.strip()
        if not txt:
            continue
            
        if txt.startswith("Số:") or txt.startswith("Số :"):
            if doc_number or header_date:
                doc_num_str = doc_number if doc_number else "Q326D……………"
                d_str = header_date if header_date else "..../..../2026"
                p.text = f"Số: {doc_num_str}\t\t\t\t\t\tNgày: {d_str}"
                for r in p.runs:
                    r.font.name = "Times New Roman"
                    r._r.get_or_add_rPr().set(qn('w:rFonts'), "Times New Roman")
                    r.font.size = Pt(13)
                    r.bold = False

        elif "Đơn vị yêu cầu" in txt and ":" in txt and customer_name:
            prefix = txt.split(":")[0].strip() + ":"
            p.text = f"{prefix}  {customer_name}"
            for r in p.runs:
                r.font.name = "Times New Roman"
                r._r.get_or_add_rPr().set(qn('w:rFonts'), "Times New Roman")
                r.font.size = Pt(13)
                r.bold = False

        elif "Địa điểm thực hiện" in txt and ":" in txt and location:
            prefix = txt.split(":")[0].strip() + ":"
            p.text = f"{prefix}  {location}"
            for r in p.runs:
                r.font.name = "Times New Roman"
                r._r.get_or_add_rPr().set(qn('w:rFonts'), "Times New Roman")
                r.font.size = Pt(13)
                r.bold = False

        elif "Nội dung yêu cầu" in txt and ":" in txt and request_content:
            prefix = txt.split(":")[0].strip() + ":"
            p.text = f"{prefix}  {request_content}"
            for r in p.runs:
                r.font.name = "Times New Roman"
                r._r.get_or_add_rPr().set(qn('w:rFonts'), "Times New Roman")
                r.font.size = Pt(13)
                r.bold = False

        elif "Thời gian thực hiện" in txt and time_date:
            p.text = f"Thời gian thực hiện: Bắt đầu từ.................... kết thúc lúc..................... ngày {time_date}"
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.left_indent = Inches(0)
            p.paragraph_format.right_indent = Inches(0)
            p.paragraph_format.first_line_indent = Inches(0)
            for r in p.runs:
                r.font.name = "Times New Roman"
                r._r.get_or_add_rPr().set(qn('w:rFonts'), "Times New Roman")
                r.font.size = Pt(13)
                r.bold = False
                    
        elif "Ký hiệu phiếu ghi chép:" in txt and doc_number:
            p.text = f"Ký hiệu phiếu ghi chép: {doc_number}/1,"
            for r in p.runs:
                r.font.name = "Times New Roman"
                r._r.get_or_add_rPr().set(qn('w:rFonts'), "Times New Roman")
                r.font.size = Pt(13)
                r.bold = False

        elif "Chỉ tiêu thực hiện lấy mẫu vi sinh trong không khí" in txt or txt.startswith("Chỉ tiêu thực hiện lấy mẫu vi sinh"):
            micro_targets_str = customer_info.get("micro_targets")
            if not micro_targets_str and all_indicators:
                micro_targets_str = extract_micro_indicators(all_indicators)
            if not micro_targets_str:
                micro_targets_str = extract_micro_indicators([request_content])

            combined_str = f"{request_content} {' '.join(str(i) for i in (all_indicators or []))}".lower()
            is_surface_micro = any(k in combined_str for k in ["bề mặt", "swab", "vsbm"]) and not any(k in combined_str for k in ["không khí", "vskk", "air micro"])
            micro_prefix = "Chỉ tiêu thực hiện lấy mẫu vi sinh bề mặt:" if is_surface_micro else "Chỉ tiêu thực hiện lấy mẫu vi sinh trong không khí:"
            # [14/11/2026] Hồ sơ có cả vi sinh không khí và vi sinh bề mặt
            if (not is_surface_micro and any(k in combined_str for k in ["bề mặt", "swab", "vsbm"])
                    and any(k in combined_str for k in ["vi sinh không khí", "vi sinh trong không khí", "vskk"])):
                micro_prefix = "Chỉ tiêu thực hiện lấy mẫu vi sinh trong không khí và vi sinh bề mặt:"

            if micro_targets_str:
                clean_target = micro_targets_str.strip().rstrip(".")
                p.text = f"{micro_prefix} {clean_target}."
            else:
                has_any_micro = False
                if all_indicators:
                    for ind in all_indicators:
                        iname = ind.get("name", "") if isinstance(ind, dict) else str(ind)
                        if any(k in iname.lower() for k in ["vi sinh", "micro", "khuẩn"]):
                            has_any_micro = True
                            break
                if not has_any_micro and all_indicators is not None:
                    p._p.getparent().remove(p._p)
                    continue
                else:
                    p.text = f"{micro_prefix} Tổng số vi sinh vật hiếu khí."

            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pPr = p._p.get_or_add_pPr()
            jc = pPr.find(qn('w:jc'))
            if jc is None:
                jc = OxmlElement('w:jc')
                pPr.append(jc)
            jc.set(qn('w:val'), 'left')
            ind = pPr.find(qn('w:ind'))
            if ind is not None:
                pPr.remove(ind)
            tabs = pPr.find(qn('w:tabs'))
            if tabs is not None:
                pPr.remove(tabs)

            p.paragraph_format.left_indent = Inches(0)
            p.paragraph_format.right_indent = Inches(0)
            p.paragraph_format.first_line_indent = Inches(0)
            for r in p.runs:
                r.font.name = "Times New Roman"
                r._r.get_or_add_rPr().set(qn('w:rFonts'), "Times New Roman")
                r.font.size = Pt(13)
                r.bold = False

        # Xóa các dòng có thông tin trạng thái phòng sạch khỏi paragraphs
        if any(kw in p.text.lower() for kw in ["trạng thái nghỉ", "ở trạng thái nghỉ", "trạng thái hoạt động", "at-rest"]):
            p.text = ""

    # Bước hoàn thiện định dạng: xóa dấu cách trước ':' và set vAlign top cho mọi ô trong bảng
    finalize_document_formatting(doc)


def find_sample_table(doc):
    """
    Tìm bảng danh mục chỉ tiêu / lấy mẫu.
    """
    for ti, tbl in enumerate(doc.tables):
        if len(tbl.rows) < 1:
            continue
        first_row_txt = " ".join(c.text.strip() for c in tbl.rows[0].cells)
        if "STT" in first_row_txt:
            if "Tên tủ" in first_row_txt or len(tbl.columns) == 5:
                return tbl, 1, "bsc"
            if len(tbl.rows) > 1:
                second_row_txt = " ".join(c.text.strip() for c in tbl.rows[1].cells)
                if any(kw in (first_row_txt + second_row_txt) for kw in ["Vị trí", "Chỉ tiêu", "Loại mẫu", "Phương pháp", "Số lượng"]):
                    if any(w in second_row_txt for w in ["Đo", "Lấy mẫu", "Chỉ tiêu"]):
                        return tbl, 2, "standard"
                    return tbl, 1, "standard"
            return tbl, 1, "standard"
    return None, 0, None


def adjust_sample_table(doc, num_samples, indicators=None, default_indicators=None, default_location=None, category_key=None):
    """
    Điều chỉnh bảng chỉ tiêu / lấy mẫu (Table 2) đúng N dòng theo số lượng yêu cầu.
    Sử dụng font Times New Roman cỡ 13, không in đậm.
    Chỉ ghi tên chỉ tiêu kèm chủ thể lấy mẫu/đo (loại bỏ mức đánh giá, yêu cầu).
    Loại bỏ thông tin trạng thái phòng và định lượng khỏi vị trí lấy mẫu.
    """
    tbl, header_rows, tbl_type = find_sample_table(doc)
    if tbl is None:
        return False
        
    num_samples = max(1, int(num_samples))
    
    indicator_list = []
    if indicators and len(indicators) > 0:
        indicator_list = indicators
    elif default_indicators and len(default_indicators) > 0:
        indicator_list = default_indicators

    # Lấy dòng mẫu (prototype row)
    if len(tbl.rows) > header_rows:
        proto_tr = copy.deepcopy(tbl.rows[header_rows]._tr)
    else:
        proto_tr = copy.deepcopy(tbl.rows[-1]._tr)

    # Xóa các dòng dữ liệu cũ
    for r in list(tbl.rows)[header_rows:]:
        tbl._tbl.remove(r._tr)

    dots = "………………………………………………………………………………………………"

    # Thêm N dòng dữ liệu mới
    for i in range(1, num_samples + 1):
        new_tr = copy.deepcopy(proto_tr)
        tbl._tbl.append(new_tr)
        new_row = tbl.rows[-1]
        num_cells = len(new_row.cells)

        if tbl_type == "bsc":
            # Cột 0: STT
            if num_cells > 0:
                set_cell_text_preserve_formatting(new_row.cells[0], str(i), bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.CENTER)
            # Cột 1: Tên tủ
            if num_cells > 1:
                cabinet_name = f"Tủ an toàn sinh học số {i}" if num_samples > 1 else "Tủ an toàn sinh học"
                set_cell_text_preserve_formatting(new_row.cells[1], cabinet_name, bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.LEFT)
            # Cột 2: Hãng sản xuất / Serial
            if num_cells > 2:
                set_cell_text_preserve_formatting(new_row.cells[2], "Hãng sản xuất: ……………………………………………………………… Serial: ………………………………………………………………", bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.LEFT)
            # Cột 3: Mã quản lý
            if num_cells > 3:
                set_cell_text_preserve_formatting(new_row.cells[3], "……………………………………………………", bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.CENTER)
            # Cột 4: Ký hiệu ghi chép
            if num_cells > 4:
                set_cell_text_preserve_formatting(new_row.cells[4], "…………………………………………", bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.CENTER)

        else:
            # Bảng chuẩn 8 cột
            ind_info = None
            if i - 1 < len(indicator_list):
                ind_info = indicator_list[i - 1]
            elif indicator_list:
                ind_info = indicator_list[(i - 1) % len(indicator_list)]

            raw_name = ind_info.get("name", dots) if isinstance(ind_info, dict) else (str(ind_info) if ind_info else dots)
            # Làm sạch & chuẩn hóa: Ghi tên chỉ tiêu kèm chủ thể lấy mẫu/đo, loại bỏ mức đánh giá
            ind_name = clean_indicator_name(raw_name, category_key=category_key)

            if isinstance(ind_info, dict) and "is_measure" in ind_info and "is_sample" in ind_info:
                is_measure = ind_info["is_measure"]
                is_sample = ind_info["is_sample"]
            else:
                from .config import infer_method_type
                is_measure, is_sample = infer_method_type(ind_name)
            
            raw_loc = ind_info.get("location", "") if isinstance(ind_info, dict) else ""
            loc_val = clean_text_no_status_or_quantitative(raw_loc)
            if " - " in loc_val:
                loc_val = loc_val.split(" - ")[0].strip()
            if not loc_val or loc_val.startswith(".") or loc_val.startswith("…"):
                loc_val = dots
            
            # Cột 0: STT
            if num_cells > 0:
                set_cell_text_preserve_formatting(new_row.cells[0], str(i), bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.CENTER)
            # Cột 1: Vị trí (không chứa thông tin trạng thái hay định lượng)
            if num_cells > 1:
                set_cell_text_preserve_formatting(new_row.cells[1], loc_val, bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.LEFT)
            # Cột 2: Chỉ tiêu (chỉ ghi tên chỉ tiêu, font 13, không in đậm)
            if num_cells > 2:
                set_cell_text_preserve_formatting(new_row.cells[2], ind_name, bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.LEFT)
            # Cột 3: Đo
            if num_cells > 3:
                set_cell_text_preserve_formatting(new_row.cells[3], "x" if is_measure else "", bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.CENTER)
            # Cột 4: Lấy mẫu
            if num_cells > 4:
                set_cell_text_preserve_formatting(new_row.cells[4], "x" if is_sample else "", bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.CENTER)
            # Cột 5: Số lượng - để dòng chấm để điền tay (không để trống, vừa gọn 1 dòng trong ô 829 dxa)
            if num_cells > 5:
                qty_val = ind_info.get("quantity", "") if isinstance(ind_info, dict) else ""
                qty_str = str(qty_val) if qty_val else ("\u2026" * 2)
                set_cell_text_preserve_formatting(new_row.cells[5], qty_str, bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.CENTER)
            # Cột 6: Kí hiệu mẫu - để dòng chấm để điền tay (không để trống)
            if num_cells > 6:
                code_val = ind_info.get("sample_code", "") if isinstance(ind_info, dict) else ""
                code_str = str(code_val) if code_val else "…………"
                set_cell_text_preserve_formatting(new_row.cells[6], code_str, bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.CENTER)
            # Cột 7: Niêm phong mẫu - để dòng chấm để điền tay (không để trống)
            if num_cells > 7:
                seal_val = ind_info.get("seal", "") if isinstance(ind_info, dict) else ""
                seal_str = str(seal_val) if seal_val else "…………"
                set_cell_text_preserve_formatting(new_row.cells[7], seal_str, bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.CENTER)

        # Giãn dòng 1.5 để KTV có khoảng trống viết tay
        for c in new_row.cells:
            for p in c.paragraphs:
                p.paragraph_format.line_spacing = 1.5

    # Chống ngắt dòng qua 2 trang (thêm cantSplit cho tất cả các hàng của bảng)
    for row in tbl.rows:
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn('w:cantSplit')) is None:
            trPr.append(OxmlElement('w:cantSplit'))

    return True


def extract_quantities_from_indicators(indicators):
    """
    Trích xuất số lượng vị trí dự kiến cho từng loại chỉ tiêu từ danh sách indicators.
    Mặc định: dust: 6, diff: 1, temp_hum: 1, noise: 1, micro: 4
    """
    if not indicators:
        return {
            "dust": 6,
            "diff": 1,
            "temp_hum": 1,
            "noise": 1,
            "micro": 4
        }

    quantities = {
        "dust": 0,
        "diff": 0,
        "temp_hum": 0,
        "noise": 0,
        "micro": 0
    }

    for ind in indicators:
        if isinstance(ind, dict):
            name = ind.get("name", "").lower()
            qty_str = str(ind.get("quantity", ""))
        else:
            name = str(ind).lower()
            qty_str = ""

        # Tìm con số đầu tiên trong qty_str nếu có
        m = re.search(r"\d+", qty_str)
        qty = int(m.group(0)) if m else None

        if any(k in name for k in ["bụi", "particle", "hạt"]):
            quantities["dust"] = max(1, qty) if qty is not None else 6
        elif any(k in name for k in ["chênh áp", "áp suất", "differential pressure"]):
            quantities["diff"] = max(1, qty) if qty is not None else 1
        elif any(k in name for k in ["nhiệt độ", "độ ẩm", "nhiệt ẩm"]):
            quantities["temp_hum"] = max(1, qty) if qty is not None else 1
        elif any(k in name for k in ["độ ồn", "tiếng ồn", "ồn"]):
            quantities["noise"] = max(1, qty) if qty is not None else 1
        elif any(k in name for k in ["vi sinh", "micro", "khuẩn"]):
            quantities["micro"] = max(1, qty) if qty is not None else 4

    return quantities


def set_table_title_above(doc, tbl, title_text, bold=True, font_size=Pt(13)):
    """
    Ghi tên chỉ tiêu vào đoạn văn bản ngay phía trên bảng.
    Nếu phía trước bảng đã có đoạn văn thì cập nhật, nếu chưa có thì chèn thêm một đoạn văn mới.
    Đảm bảo tiêu đề được format lề trái (align left), không bị thụt đầu dòng.
    """
    prev = tbl._tbl.getprevious()
    if prev is not None and prev.tag.endswith('p'):
        p = docx.text.paragraph.Paragraph(prev, doc)
    else:
        new_p_elem = OxmlElement('w:p')
        tbl._tbl.addprevious(new_p_elem)
        p = docx.text.paragraph.Paragraph(new_p_elem, doc)

    # Xóa sạch các run cũ
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
        
    pPr = p._p.get_or_add_pPr()
    tabs = pPr.find(qn('w:tabs'))
    if tabs is not None:
        pPr.remove(tabs)
        
    ind = pPr.find(qn('w:ind'))
    if ind is not None:
        pPr.remove(ind)

    # Đảm bảo căn lề trái tuyệt đối
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    jc = pPr.find(qn('w:jc'))
    if jc is None:
        jc = OxmlElement('w:jc')
        pPr.append(jc)
    jc.set(qn('w:val'), 'left')
            
    p.paragraph_format.left_indent = Inches(0)
    p.paragraph_format.right_indent = Inches(0)
    p.paragraph_format.first_line_indent = Inches(0)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    
    run = p.add_run(title_text)
    run.font.name = "Times New Roman"
    run._r.get_or_add_rPr().set(qn('w:rFonts'), "Times New Roman")
    run.font.size = font_size
    run.bold = bold
    return p


def clean_measurement_tables_in_bm01(doc, category_id, location=None, indicators=None, num_samples=None, extra_rows=3):
    """
    Làm sạch các bảng ghi chép số liệu đo cũ trong BM3030.01.
    Ghi tên chỉ tiêu phía trên mỗi bảng đo.
    Giảm số lượng dòng trong các bảng đo theo số lượng dự kiến, thêm 2-3 dòng trắng dự phòng.
    Đồng thời loại bỏ hoàn toàn các thông tin trạng thái phòng sạch hoặc thông số định lượng.
    """
    if category_id == "co2_mtld":
        if len(doc.tables) >= 3:
            tbl = doc.tables[2]
            set_table_title_above(doc, tbl, "Đo CO2 và TVOC trong môi trường lao động", bold=True, font_size=Pt(13))
            num_data = int(num_samples) if num_samples else 4
            target_rows = 1 + num_data + extra_rows
            while len(tbl.rows) > target_rows:
                tbl._tbl.remove(tbl.rows[-1]._tr)
            for ri in range(1, len(tbl.rows)):
                row = tbl.rows[ri]
                for ci in range(len(row.cells)):
                    cell_text = row.cells[ci].text.strip()
                    if re.match(r"^\d+(\.\d+)?$", cell_text):
                        set_cell_text_preserve_formatting(row.cells[ci], "", bold=False, font_size=Pt(13))
                        
    elif category_id == "khi_nen":
        for tbl in doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    if "HDTN" in cell.text or "MTTN" in cell.text:
                        txt = cell.text
                        txt = re.sub(r"HDTN\w+", "………………………………", txt)
                        txt = re.sub(r"MTTN\w+", "………………………………", txt)
                        set_cell_text_preserve_formatting(cell, txt, bold=False, font_size=Pt(13))

        # Co giãn số lượng bảng ghi chép hiện trường theo đúng số lượng vị trí yêu cầu
        if num_samples and int(num_samples) > 0 and len(doc.tables) > 1:
            target_tables = int(num_samples)
            while len(doc.tables) - 1 < target_tables:
                last_tbl = doc.tables[-1]
                p = doc.add_paragraph()
                new_tbl = copy.deepcopy(last_tbl._tbl)
                p._p.addnext(new_tbl)
            while len(doc.tables) - 1 > target_tables and len(doc.tables) > 2:
                doc.tables[-1]._tbl.getparent().remove(doc.tables[-1]._tbl)

    elif category_id == "phong_sach":
        quantities = extract_quantities_from_indicators(indicators)
        num_dust = quantities["dust"]
        if num_samples and int(num_samples) > 0:
            if num_dust > 0:
                num_dust = max(num_dust, int(num_samples))
            elif not any([quantities["diff"], quantities["temp_hum"], quantities["noise"], quantities["micro"]]):
                num_dust = int(num_samples)
        num_diff = quantities["diff"]
        num_temp_hum = quantities["temp_hum"]
        num_noise = quantities["noise"]
        num_micro = quantities["micro"]

        # 1. Bảng đo hạt bụi (Table 2 nếu có)
        if len(doc.tables) >= 3:
            tbl_dust = doc.tables[2]
            if num_dust == 0:
                prev = tbl_dust._tbl.getprevious()
                if prev is not None and prev.tag.endswith('p'):
                    prev.getparent().remove(prev)
                tbl_dust._tbl.getparent().remove(tbl_dust._tbl)
            else:
                # Giữ nguyên 100% format chuẩn của biểu mẫu gốc:
                # Bật fixed layout để chống Word tự co giãn gây vỡ dòng các checkbox HVAC
                tblPr_dust = tbl_dust._tbl.tblPr
                if tblPr_dust.find(qn('w:tblLayout')) is None:
                    tblPr_dust.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))

                # - Row 0 Cell 0: Tên khu vực thực hiện kèm 3 dải chấm điền tay (định lượng chấm vừa vặn 1 dòng, không bị ngắt dòng rơi chữ/chấm)
                if len(tbl_dust.rows) > 0 and len(tbl_dust.rows[0].cells) > 0:
                    c0 = tbl_dust.rows[0].cells[0]
                    if len(c0.paragraphs) > 0:
                        p0 = c0.paragraphs[0]
                        if len(p0.runs) > 1:
                            p0.runs[1].text = "Tên khu vực thực hiện: ………"
                        else:
                            p0.text = "   Tên khu vực thực hiện: ………"
                        for r in p0.runs:
                            r.italic = True
                            r.font.name = 'Times New Roman'
                            r.font.size = Pt(11)
                    if len(c0.paragraphs) > 1:
                        p1 = c0.paragraphs[1]
                        p1.text = "   …………………………………"
                        for r in p1.runs:
                            r.italic = True
                            r.font.name = 'Times New Roman'
                            r.font.size = Pt(11)
                    if len(c0.paragraphs) > 2:
                        p2 = c0.paragraphs[2]
                        p2.text = "   …………………………………"
                        for r in p2.runs:
                            r.italic = True
                            r.font.name = 'Times New Roman'
                            r.font.size = Pt(11)

                # - Row 0 Cell 1: Giữ nguyên 100% (HVAC status + Wingdings checkboxes + tabs)
                # - Row 1 Cell 0: Giữ nguyên ô trống của biểu mẫu gốc (KHÔNG ghi 'CHỈ TIÊU THỰC HIỆN')
                # - Row 1 Cell 1: Giữ nguyên khối Particle, tối ưu các dòng chấm thời gian và File đo để nằm gọn trên 1 dòng
                if len(tbl_dust.rows) > 1 and len(tbl_dust.rows[1].cells) > 1:
                    c1_dust = tbl_dust.rows[1].cells[1]
                    for p in c1_dust.paragraphs:
                        if "Thời gian lấy" in p.text:
                            for r in p.runs:
                                if "…" in r.text or "..." in r.text:
                                    r.text = " ………………………………"
                        if "File zeroing" in p.text:
                            runs = p.runs
                            if len(runs) >= 3:
                                runs[2].text = " File zeroing ………………………………"
                                for r in runs[3:]:
                                    r.text = ""
                        if "File đo" in p.text:
                            runs = p.runs
                            if len(runs) >= 3:
                                runs[2].text = " File đo………………………………………"
                                for r in runs[3:]:
                                    r.text = ""

                # - Row 2: Giữ nguyên tiêu đề cột ['Ký hiệu vị trí đo', 'File ghi']
                
                # - Row 3 trở đi: Co giãn số dòng theo số lượng yêu cầu (target = 3 + num_dust + extra_rows)
                # Giữ nguyên STT tự động (numPr) ở Cột 0 và ô trống ở Cột 1 để KTV điền tay
                target_dust_rows = 3 + num_dust + extra_rows
                while len(tbl_dust.rows) > target_dust_rows:
                    tbl_dust._tbl.remove(tbl_dust.rows[-1]._tr)
                while len(tbl_dust.rows) < target_dust_rows:
                    new_tr = copy.deepcopy(tbl_dust.rows[-1]._tr)
                    tbl_dust._tbl.append(new_tr)

                for r in tbl_dust.rows:
                    trPr = r._tr.get_or_add_trPr()
                    if trPr.find(qn('w:cantSplit')) is None:
                        trPr.append(OxmlElement('w:cantSplit'))

        # Xử lý các bảng tiếp theo theo thứ tự ngược: Table 5 -> Table 4 -> Table 3
        # 4. Bảng vi sinh (Table 5 nếu có)
        if len(doc.tables) >= 6:
            tbl_micro = doc.tables[5]
            if num_micro == 0:
                prev = tbl_micro._tbl.getprevious()
                if prev is not None and prev.tag.endswith('p'):
                    prev.getparent().remove(prev)
                tbl_micro._tbl.getparent().remove(tbl_micro._tbl)
            else:
                set_table_title_above(doc, tbl_micro, "Lấy mẫu vi sinh trong không khí", bold=True, font_size=Pt(13))

                if len(tbl_micro.rows) > 0 and len(tbl_micro.rows[0].cells) > 0:
                    set_cell_text_preserve_formatting(tbl_micro.rows[0].cells[0], "Tên khu vực thực hiện: …………………………………………………………………………", bold=False, font_size=Pt(13))
                if len(tbl_micro.rows) > 1:
                    if len(tbl_micro.rows[1].cells) > 0:
                        set_cell_text_preserve_formatting(tbl_micro.rows[1].cells[0], "CHỈ TIÊU THỰC HIỆN", bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.CENTER)
                    if len(tbl_micro.rows[1].cells) > 1:
                        micro_spec = (
                            "Lấy mẫu vi sinh trong không khí\n"
                            "+ Thiết bị sử dụng: Buck – Bio Culture B30120\n"
                            "+ Lưu lượng lấy mẫu: ……………………\n"
                            "+ Thời gian lấy mẫu: …… phút (từ …… đến ……)\n"
                            "+ Ký hiệu đĩa:\n"
                            "- TS lấy mẫu: ……………………\n"
                            "- TS blank: ……………………"
                        )
                        set_cell_text_preserve_formatting(tbl_micro.rows[1].cells[1], micro_spec, bold=False, font_size=Pt(13))

                target_micro_rows = 3 + num_micro + extra_rows
                while len(tbl_micro.rows) < target_micro_rows:
                    new_tr = copy.deepcopy(tbl_micro.rows[-1]._tr)
                    tbl_micro._tbl.append(new_tr)
                while len(tbl_micro.rows) > target_micro_rows:
                    tbl_micro._tbl.remove(tbl_micro.rows[-1]._tr)

                for idx in range(1, num_micro + 1):
                    r_idx = 2 + idx
                    row_c = tbl_micro.rows[r_idx].cells
                    if len(row_c) > 0:
                        set_cell_text_preserve_formatting(row_c[0], "…………………………………………", bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.CENTER)
                    if len(row_c) > 1:
                        set_cell_text_preserve_formatting(row_c[1], str(idx), bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.CENTER)

                for idx in range(num_micro + 1, num_micro + extra_rows + 1):
                    r_idx = 2 + idx
                    row_c = tbl_micro.rows[r_idx].cells
                    if len(row_c) > 0:
                        set_cell_text_preserve_formatting(row_c[0], "…………………………………………", bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.CENTER)
                    if len(row_c) > 1:
                        set_cell_text_preserve_formatting(row_c[1], "…………………………………………", bold=False, font_size=Pt(13), alignment=WD_ALIGN_PARAGRAPH.CENTER)

                for r in tbl_micro.rows:
                    trPr = r._tr.get_or_add_trPr()
                    if trPr.find(qn('w:cantSplit')) is None:
                        trPr.append(OxmlElement('w:cantSplit'))

        # 3. Bảng đo nhiệt ẩm & độ ồn (Table 4 nếu có)
        if len(doc.tables) >= 5:
            tbl_env = doc.tables[4]
            if num_temp_hum == 0 and num_noise == 0:
                prev = tbl_env._tbl.getprevious()
                if prev is not None and prev.tag.endswith('p'):
                    prev.getparent().remove(prev)
                tbl_env._tbl.getparent().remove(tbl_env._tbl)
            else:
                title_env = "Đo nhiệt độ, độ ẩm trong không khí và đo độ ồn" if num_noise > 0 else "Đo nhiệt độ, độ ẩm trong không khí"
                set_table_title_above(doc, tbl_env, title_env, bold=True, font_size=Pt(13))

                if len(tbl_env.rows) > 0 and len(tbl_env.rows[0].cells) > 0:
                    env_head = f"{title_env}\n+ Thiết bị sử dụng đo nhiệt ẩm: Testo 625\n+ Thiết bị sử dụng đo độ ồn: Rion NL 53"
                    set_cell_text_preserve_formatting(tbl_env.rows[0].cells[0], env_head, bold=False, font_size=Pt(13))

                extra_noise = 2 if num_noise > 0 else 0
                target_env_rows = 2 + num_temp_hum + extra_rows + (num_noise + extra_noise if num_noise > 0 else 0)
                while len(tbl_env.rows) < target_env_rows:
                    new_tr = copy.deepcopy(tbl_env.rows[-1]._tr)
                    tbl_env._tbl.append(new_tr)
                while len(tbl_env.rows) > target_env_rows:
                    tbl_env._tbl.remove(tbl_env.rows[-1]._tr)

                current_r = 2
                # Dòng nhiệt ẩm dự kiến
                for idx in range(1, num_temp_hum + 1):
                    row_c = tbl_env.rows[current_r].cells
                    if len(row_c) >= 3:
                        set_cell_text_preserve_formatting(row_c[0], "…………………………………………", bold=False, font_size=Pt(13))
                        set_cell_text_preserve_formatting(row_c[1], "………… °C", bold=False, font_size=Pt(13))
                        set_cell_text_preserve_formatting(row_c[2], "………… %", bold=False, font_size=Pt(13))
                    current_r += 1

                # Các dòng trắng nhiệt ẩm dự phòng
                for _ in range(extra_rows):
                    row_c = tbl_env.rows[current_r].cells
                    if len(row_c) >= 3:
                        set_cell_text_preserve_formatting(row_c[0], "………………………………", bold=False, font_size=Pt(13))
                        set_cell_text_preserve_formatting(row_c[1], "………… °C", bold=False, font_size=Pt(13))
                        set_cell_text_preserve_formatting(row_c[2], "………… %", bold=False, font_size=Pt(13))
                    current_r += 1

                # Dòng độ ồn nếu có
                if num_noise > 0:
                    for idx in range(1, num_noise + 1):
                        row_c = tbl_env.rows[current_r].cells
                        if len(row_c) >= 3:
                            set_cell_text_preserve_formatting(row_c[0], "Độ ồn ……………………", bold=False, font_size=Pt(13))
                            set_cell_text_preserve_formatting(row_c[1], "Độ ồn nền: ……… dBA", bold=False, font_size=Pt(13))
                            set_cell_text_preserve_formatting(row_c[2], "Khi vận hành: ……… dBA", bold=False, font_size=Pt(13))
                        current_r += 1

                    for _ in range(extra_noise):
                        row_c = tbl_env.rows[current_r].cells
                        if len(row_c) >= 3:
                            set_cell_text_preserve_formatting(row_c[0], "Độ ồn ……………………", bold=False, font_size=Pt(13))
                            set_cell_text_preserve_formatting(row_c[1], "Độ ồn nền: ……… dBA", bold=False, font_size=Pt(13))
                            set_cell_text_preserve_formatting(row_c[2], "Khi vận hành: ……… dBA", bold=False, font_size=Pt(13))
                        current_r += 1

                for r in tbl_env.rows:
                    trPr = r._tr.get_or_add_trPr()
                    if trPr.find(qn('w:cantSplit')) is None:
                        trPr.append(OxmlElement('w:cantSplit'))

        # 2. Bảng đo chênh áp (Table 3 nếu có)
        if len(doc.tables) >= 4:
            tbl_diff = doc.tables[3]
            if num_diff == 0:
                prev = tbl_diff._tbl.getprevious()
                if prev is not None and prev.tag.endswith('p'):
                    prev.getparent().remove(prev)
                tbl_diff._tbl.getparent().remove(tbl_diff._tbl)
            else:
                set_table_title_above(doc, tbl_diff, "Đo chênh áp", bold=True, font_size=Pt(13))

                if len(tbl_diff.rows) > 0 and len(tbl_diff.rows[0].cells) > 0:
                    set_cell_text_preserve_formatting(tbl_diff.rows[0].cells[0], "Đo chênh áp\n+ Thiết bị sử dụng: TSI 8375", bold=False, font_size=Pt(13))

                target_diff_rows = 2 + num_diff + extra_rows
                while len(tbl_diff.rows) < target_diff_rows:
                    new_tr = copy.deepcopy(tbl_diff.rows[-1]._tr)
                    tbl_diff._tbl.append(new_tr)
                while len(tbl_diff.rows) > target_diff_rows:
                    tbl_diff._tbl.remove(tbl_diff.rows[-1]._tr)

                if len(tbl_diff.rows) > 2 and len(tbl_diff.rows[2].cells) >= 2:
                    set_cell_text_preserve_formatting(tbl_diff.rows[2].cells[0], "………………………………………… -> Hành lang", bold=False, font_size=Pt(13))
                    set_cell_text_preserve_formatting(tbl_diff.rows[2].cells[1], "……………… Pa", bold=False, font_size=Pt(13))
                    for ci in range(2, len(tbl_diff.rows[2].cells)):
                        set_cell_text_preserve_formatting(tbl_diff.rows[2].cells[ci], "", bold=False, font_size=Pt(13))

                for idx in range(2, num_diff + 1):
                    r_idx = 1 + idx
                    row_c = tbl_diff.rows[r_idx].cells
                    if len(row_c) > 0:
                        set_cell_text_preserve_formatting(row_c[0], "………………………………………… -> Hành lang", bold=False, font_size=Pt(13))
                    if len(row_c) > 1:
                        set_cell_text_preserve_formatting(row_c[1], "……………… Pa", bold=False, font_size=Pt(13))
                    for ci in range(2, len(row_c)):
                        set_cell_text_preserve_formatting(row_c[ci], "", bold=False, font_size=Pt(13))

                for idx in range(num_diff + 1, num_diff + extra_rows + 1):
                    r_idx = 1 + idx
                    row_c = tbl_diff.rows[r_idx].cells
                    if len(row_c) > 0:
                        set_cell_text_preserve_formatting(row_c[0], "…………………………………………", bold=False, font_size=Pt(13))
                    if len(row_c) > 1:
                        set_cell_text_preserve_formatting(row_c[1], "……………… Pa", bold=False, font_size=Pt(13))
                    for ci in range(2, len(row_c)):
                        set_cell_text_preserve_formatting(row_c[ci], "", bold=False, font_size=Pt(13))

                for r in tbl_diff.rows:
                    trPr = r._tr.get_or_add_trPr()
                    if trPr.find(qn('w:cantSplit')) is None:
                        trPr.append(OxmlElement('w:cantSplit'))

        # Cập nhật font chữ và căn lề các paragraph trong tài liệu
        for p in doc.paragraphs:
            txt = p.text.strip()
            if not txt:
                continue
            # Tiêu đề các bảng đo: format lề trái tuyệt đối
            if any(kw in txt for kw in ["Đo nồng độ", "Đo chênh áp", "Đo nhiệt độ", "Lấy mẫu vi sinh", "Đo CO2"]):
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                pPr = p._p.get_or_add_pPr()
                jc = pPr.find(qn('w:jc'))
                if jc is None:
                    jc = OxmlElement('w:jc')
                    pPr.append(jc)
                jc.set(qn('w:val'), 'left')
                ind = pPr.find(qn('w:ind'))
                if ind is not None:
                    pPr.remove(ind)
                tabs = pPr.find(qn('w:tabs'))
                if tabs is not None:
                    pPr.remove(tabs)
                p.paragraph_format.left_indent = Inches(0)
                p.paragraph_format.first_line_indent = Inches(0)
                for r in p.runs:
                    r.font.name = "Times New Roman"
                    r.font.size = Pt(13)
                    r.bold = True
            elif "Người thực hiện" not in txt and "Đại diện" not in txt:
                p.paragraph_format.left_indent = Inches(0)
                p.paragraph_format.first_line_indent = Inches(0)
                for r in p.runs:
                    r.font.name = "Times New Roman"
                    r.font.size = Pt(13)

        # Quét dọn toàn bộ các ô còn sót lại thông tin trạng thái phòng
        for tbl in doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    txt = cell.text
                    if "trạng thái" in txt.lower() or "at-rest" in txt.lower():
                        cleaned = clean_text_no_status_or_quantitative(txt)
                        if cleaned != txt:
                            set_cell_text_preserve_formatting(cell, cleaned, bold=False, font_size=Pt(13))

    return True
