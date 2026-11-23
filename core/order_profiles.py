"""
Module quản lý danh mục cấu hình hồ sơ mẫu (Order Profiles)
Đúc kết toàn bộ các mẫu hồ sơ thực tế đã thực hiện tại QUATEST 3:
- BSC (Tủ an toàn sinh học cấp II)
- Khí nén (1 hoặc nhiều vị trí)
- Phòng sạch (ISO 14644)
- Phòng sạch có Vi sinh không khí & Swab bề mặt
- Khí nén kết hợp Phòng sạch (Hybrid)
- Môi trường lao động (CO2, TVOC, Vi khí hậu)
- Kho lạnh (Mapping nhiệt ẩm)
- Môi trường nước (Nước uống, nước sinh hoạt, nước thải)
"""

ORDER_PROFILES = {
    # 1. Tủ an toàn sinh học cấp II (BSC)
    "bsc": {
        "id": "bsc",
        "name": "Giám định tủ an toàn sinh học cấp II (BSC)",
        "category_key": "bsc",
        "short_content": "BSC",
        "default_request": "Giám định tủ an toàn sinh học cấp 2.",
        "default_equipment": [
            "- Thiết bị đo lưu lượng TSI 8375/8380, USA;",
            "- Máy đo tốc độ gió Kanomax Climomaster 6501, Japan;",
            "- Máy tạo sol khí ATI aerosol generator, model 6D và quang kế aerosol photometer, model 2i, USA;",
            "- Máy tạo khói khí dung (Smoke pencil / Aerosol generator);",
            "- Máy đo độ rọi Testo 440, Germany;",
            "- Máy đo cường độ bức xạ UVC Extech SDL470, USA;",
            "- Máy đo độ ồn Rion NL-53, Japan;",
            "- Máy đếm hạt TSI AeroTrak 9350, USA."
        ],
        "default_procedure": "QUATEST 3 3030:2025; EN 12469:2000; N2.TC.BSC.02",
        "indicators": [
            {"name": "Đo tốc độ trung bình dòng khí qua cửa làm việc (inflow) tủ an toàn sinh học", "is_measure": True, "is_sample": False},
            {"name": "Đo tốc độ trung bình dòng khí cấp qua lọc HEPA (downflow) tủ an toàn sinh học", "is_measure": True, "is_sample": False},
            {"name": "Đo cường độ sáng trung bình bên trong khu vực làm việc tủ an toàn sinh học", "is_measure": True, "is_sample": False},
            {"name": "Đo độ ồn tủ an toàn sinh học", "is_measure": True, "is_sample": False},
            {"name": "Kiểm tra độ rò rỉ của các lọc HEPA tủ an toàn sinh học", "is_measure": True, "is_sample": False},
            {"name": "Kiểm tra hình thái dòng khí tủ an toàn sinh học", "is_measure": True, "is_sample": False},
            {"name": "Đo cường độ UVC tủ an toàn sinh học", "is_measure": True, "is_sample": False},
            {"name": "Đo nồng độ bụi trong tủ an toàn sinh học", "is_measure": True, "is_sample": False}
        ]
    },

    # 2. Khí nén & Khí có áp chuẩn
    "khi_nen": {
        "id": "khi_nen",
        "name": "Kiểm tra chất lượng hệ thống khí nén",
        "category_key": "khi_nen",
        "short_content": "Khí nén",
        "default_request": "Kiểm tra chất lượng hệ thống khí nén.",
        "default_equipment": [
            "- Máy đo nồng độ hạt AeroTrak 9350 và phụ kiện HPD III, USA;",
            "- Thiết bị đo nhiệt độ điểm sương;",
            "- Bộ dụng cụ lấy mẫu tổng dầu trong khí nén;",
            "- Giấy lọc thủy tinh Pall life Sciences, ống than hoạt tính ORBO®."
        ],
        "default_procedure": "QUATEST 3 3030:2025; N2.TC.KN.01",
        "indicators": [
            {"name": "Lấy mẫu tổng dầu trong khí nén", "is_measure": False, "is_sample": True},
            {"name": "Đo nhiệt độ điểm sương trong khí nén", "is_measure": True, "is_sample": False},
            {"name": "Đo nồng độ bụi trong khí nén", "is_measure": True, "is_sample": False}
        ]
    },

    # 3. Phòng sạch tiêu chuẩn
    "phong_sach": {
        "id": "phong_sach",
        "name": "Đo kiểm chất lượng không khí phòng sạch",
        "category_key": "phong_sach",
        "short_content": "Phòng sạch",
        "default_request": "Đo kiểm chất lượng không khí phòng sạch.",
        "default_equipment": [
            "- Thiết bị đo nồng độ hạt TSI AeroTrak 9350, USA;",
            "- Thiết bị đo nhiệt độ, độ ẩm trong không khí Testo 625, Germany;",
            "- Thiết bị đo chênh lệch áp suất TSI 8375, USA;",
            "- Thiết bị đo độ ồn Rion NL 53, Japan;",
            "- Thiết bị đo độ sáng Testo 440, Germany."
        ],
        "default_procedure": "QUATEST 3 3030:2025; N2.TC.PS.01",
        "indicators": [
            {"name": "Đo nồng độ bụi phòng sạch", "is_measure": True, "is_sample": False},
            {"name": "Đo nhiệt độ, độ ẩm phòng sạch", "is_measure": True, "is_sample": False},
            {"name": "Đo chênh áp phòng sạch", "is_measure": True, "is_sample": False},
            {"name": "Đo độ ồn phòng sạch", "is_measure": True, "is_sample": False},
            {"name": "Đo độ sáng phòng sạch", "is_measure": True, "is_sample": False}
        ]
    },

    # 4. Phòng sạch có Vi sinh không khí (VSKK) và Swab bề mặt
    "phong_sach_vskk_swab": {
        "id": "phong_sach_vskk_swab",
        "name": "Phòng sạch có Vi sinh không khí và Vi sinh bề mặt",
        "category_key": "phong_sach",
        "short_content": "vskk - swab",
        "default_request": "Quan trắc phòng sạch và lấy mẫu vi sinh không khí, vi sinh bề mặt.",
        "default_equipment": [
            "- Thiết bị đo nồng độ hạt TSI AeroTrak 9350, USA;",
            "- Thiết bị đo nhiệt độ, độ ẩm trong không khí Testo 625, Germany;",
            "- Thiết bị đo chênh lệch áp suất TSI 8375, USA;",
            "- Thiết bị lấy mẫu vi sinh trong không khí Buck – Bio Culture B30120, USA;",
            "- Dụng cụ lấy mẫu vi sinh bề mặt chuyên dụng."
        ],
        "default_procedure": "QUATEST 3 3030:2025; N2.TC.PS.01; N2.TC.09",
        "indicators": [
            {"name": "Đo nồng độ bụi phòng sạch", "is_measure": True, "is_sample": False},
            {"name": "Đo nhiệt độ, độ ẩm phòng sạch", "is_measure": True, "is_sample": False},
            {"name": "Đo chênh áp phòng sạch", "is_measure": True, "is_sample": False},
            {"name": "Lấy mẫu vi sinh trong không khí phòng sạch", "is_measure": False, "is_sample": True},
            {"name": "Lấy mẫu vi sinh bề mặt", "is_measure": False, "is_sample": True}
        ]
    },

    # 5. Khí nén kết hợp Phòng sạch (Hybrid)
    "khi_nen_phong_sach": {
        "id": "khi_nen_phong_sach",
        "name": "Khí nén kết hợp Phòng sạch (Hybrid)",
        "category_key": "khi_nen_phong_sach",
        "short_content": "Khí nén - Phòng sạch",
        "default_request": "Kiểm tra chất lượng hệ thống khí nén và đo nồng độ bụi phòng sạch.",
        "default_equipment": [
            "- Máy đo nồng độ hạt AeroTrak 9350 và phụ kiện HPD III, USA;",
            "- Thiết bị đo nhiệt độ điểm sương;",
            "- Bộ dụng cụ lấy mẫu tổng dầu trong khí nén;",
            "- Giấy lọc thủy tinh Pall life Sciences, ống than hoạt tính ORBO®;",
            "- Thiết bị đo nồng độ hạt TSI AeroTrak 9350, USA."
        ],
        "default_procedure": "QUATEST 3 3030:2025; N2.TC.KN.01; N2.TC.PS.01",
        "indicators": [
            {"name": "Lấy mẫu tổng dầu trong khí nén", "is_measure": False, "is_sample": True},
            {"name": "Đo nhiệt độ điểm sương trong khí nén", "is_measure": True, "is_sample": False},
            {"name": "Đo nồng độ bụi trong khí nén", "is_measure": True, "is_sample": False},
            {"name": "Đo nồng độ bụi phòng sạch", "is_measure": True, "is_sample": False}
        ]
    },

    # 6. Môi trường lao động (CO2, TVOC, Vi khí hậu)
    "co2_mtld": {
        "id": "co2_mtld",
        "name": "Môi trường lao động (CO2, TVOC, Vi khí hậu)",
        "category_key": "co2_mtld",
        "short_content": "MTLĐ",
        "default_request": "Quan trắc chất lượng không khí môi trường lao động: TVOC và CO2.",
        "default_equipment": [
            "- Thiết bị đo nồng độ CO2 trong không khí testo 440, Germany;",
            "- Thiết bị đo nồng độ TVOC trong không khí Drager Xam 8000, Germany."
        ],
        "default_procedure": "QUATEST 3 3030:2025; N2.TC.07",
        "indicators": [
            {"name": "Đo nồng độ TVOC trong không khí môi trường lao động", "is_measure": True, "is_sample": False},
            {"name": "Đo nồng độ CO2 trong không khí môi trường lao động", "is_measure": True, "is_sample": False}
        ]
    },

    # 7. Kho lạnh / Tủ bảo quản (Mapping nhiệt độ độ ẩm)
    "kho_lanh": {
        "id": "kho_lanh",
        "name": "Kho lạnh / Tủ bảo quản (Mapping nhiệt ẩm)",
        "category_key": "kho_lanh",
        "short_content": "Kho lạnh",
        "default_request": "Đánh giá độ đồng đều nhiệt độ kho bảo quản thuốc.",
        "default_equipment": [
            "- Thiết bị đo nhiệt độ độ ẩm Testo 625, Germany;",
            "- Thiết bị tự ghi nhiệt độ độ ẩm Haxo-8 (mapping)."
        ],
        "default_procedure": "QUATEST 3 3030:2025; N2.TC.MP.01",
        "indicators": [
            {"name": "Đánh giá độ đồng đều nhiệt độ kho bảo quản thuốc", "is_measure": True, "is_sample": False},
            {"name": "Đánh giá độ đồng đều độ ẩm kho bảo quản thuốc", "is_measure": True, "is_sample": False}
        ]
    },

    # 8. Lấy mẫu vi sinh bề mặt (Swab / VSBM)
    "vi_sinh_be_mat": {
        "id": "vi_sinh_be_mat",
        "name": "Lấy mẫu và xác định độ nhiễm bẩn vi sinh bề mặt",
        "category_key": "phong_sach",
        "short_content": "VSBM",
        "default_request": "Lấy mẫu và xác định độ nhiễm bẩn vi sinh bề mặt.",
        "default_equipment": [
            "- Dụng cụ lấy mẫu vi sinh bề mặt chuyên dụng."
        ],
        "default_procedure": "QUATEST 3 3030:2025; N2.TC.09",
        "indicators": [
            {"name": "Lấy mẫu vi sinh bề mặt", "is_measure": False, "is_sample": True}
        ]
    }
}


def get_profile(profile_id):
    """Lấy cấu hình profile theo id."""
    return ORDER_PROFILES.get(profile_id)


def detect_profile(order_data):
    """
    Tự động nhận diện profile phù hợp nhất dựa trên dữ liệu đơn hàng.
    """
    if not order_data:
        return "phong_sach"

    req = str(order_data.get("request_content", "")).lower()
    inds = " ".join(str(i) for i in order_data.get("indicators", [])).lower()
    cat = str(order_data.get("category", "")).lower()
    cname = str(order_data.get("customer_name", "")).lower()
    combined = f"{req} {inds} {cat} {cname}"

    # 1. BSC
    if any(k in combined for k in ["bsc", "an toàn sinh học", "tủ an toàn", "en 12469"]):
        return "bsc"

    # 2. Khí nén kết hợp Phòng sạch (Hybrid)
    has_kn = any(k in combined for k in ["khí nén", "khi nen", "tổng dầu", "điểm sương"])
    has_ps = any(k in combined for k in ["phòng sạch", "phong sach", "bụi phòng sạch", "cleanroom"])
    if has_kn and has_ps:
        return "khi_nen_phong_sach"

    # 3. Khí nén đơn
    if has_kn and not has_ps:
        return "khi_nen"

    # 4. Phòng sạch + Vi sinh không khí + Swab
    has_vskk = any(k in combined for k in ["vi sinh không khí", "vskk", "air micro"])
    has_swab = any(k in combined for k in ["vi sinh bề mặt", "swab", "vsbm", "rodac"])
    if has_ps and (has_vskk or has_swab):
        return "phong_sach_vskk_swab"

    # 4b. Vi sinh bề mặt đơn thuần (không có phòng sạch, không có khí nén)
    if has_swab and not has_ps and not has_kn:
        return "vi_sinh_be_mat"

    # 5. Phòng sạch đơn
    if has_ps:
        return "phong_sach"

    # 6. Kho lạnh
    if any(k in combined for k in ["kho lạnh", "kho lanh", "mapping", "tủ bảo quản", "haxo"]):
        return "kho_lanh"

    # 7. MTLĐ / CO2
    if any(k in combined for k in ["co2", "tvoc", "môi trường lao động", "mtld"]):
        return "co2_mtld"

    return "phong_sach"
