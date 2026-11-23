"""
Module cấu hình danh mục biểu mẫu, chỉ tiêu và thông số kỹ thuật chuẩn
Củng cố dữ liệu kỹ thuật từ Hệ thống Quy trình thao tác chuẩn (SOP / N2.TC) của QUATEST 3
Áp dụng cho các biểu mẫu BM3030.01 (Phiếu ghi chép) và BM3030.02 (Biên bản giám định / lấy mẫu)
Trung tâm Kỹ thuật Tiêu chuẩn Đo lường Chất lượng 3 (QUATEST 3)
"""
import os
import re

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
QUY_TRINH_DIR = os.path.join(BASE_DIR, "QuyTrinh")
# Đường dẫn lưu trữ / sao chép hồ sơ trên ổ đĩa Z của QUATEST 3 (Quy tắc 13)
Z_DRIVE_DIR = os.environ.get("QUATEST3_Z_DRIVE", r"Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026")

# ==============================================================================
# HỆ THỐNG CƠ SỞ DỮ LIỆU QUY TRÌNH THAO TÁC CHUẨN (SOP DATABASE - N2.TC)
# Trích xuất từ các tài liệu quy trình nội bộ của QUATEST 3 (Phòng Nghiệp vụ 2)
# ==============================================================================
SOP_DATABASE = {
    # --- PHÒNG SẠCH & VI SINH ---
    "N2.TC.PS.01": {
        "code": "N2.TC.PS.01",
        "name": "Đo phòng sạch",
        "standards": ["ISO 14644-1:2015", "ISO 14644-3:2019"],
        "equipment": {
            "dust_regular": "Thiết bị đo nồng độ hạt TSI AeroTrak 9350, USA;",
            "dust_fine": "Máy đo nồng độ hạt bụi PMS Lasair III 110, USA;",
            "temp_humidity": "Thiết bị đo nhiệt độ, độ ẩm trong không khí Testo 625, Germany;",
            "pressure_diff": "Thiết bị đo chênh lệch áp suất TSI 8375, USA;",
            "air_flow": "Thiết bị đo lưu lượng TSI 8375, USA;",
            "air_velocity": "Máy đo tốc độ gió Kanomax Climomaster 6501, Japan;",
            "lux": "Thiết bị đo độ sáng Testo 440, Germany;",
            "noise": "Thiết bị đo độ ồn Rion NL 53, Japan;"
        }
    },
    "N2.TC.09": {
        "code": "N2.TC.09",
        "name": "Quy trình lấy mẫu vi sinh",
        "standards": ["EN 17141:2020", "ISO 14644-1:2015", "ISO 18593:2018", "TCVN 8129:2019", "ISO 11133:2014"],
        "equipment": {
            "micro_air": "Thiết bị lấy mẫu vi sinh trong không khí Buck – Bio Culture B30120, USA.",
            "micro_compressed_air": "Thiết bị lấy mẫu vi sinh khí nén SAS Pinocchio Super II.",
            "micro_surface": "Dụng cụ lấy mẫu vi sinh bề mặt (đĩa tiếp xúc Rodac / tăm bông tiệt trùng kèm khuôn định hình)."
        }
    },
    # --- KHÍ NÉN & KHÍ CÓ ÁP ---
    "N2.TC.KN.01": {
        "code": "N2.TC.KN.01",
        "name": "Khí nén, khí có áp",
        "standards": ["ISO 8573-1:2010", "ISO 8573-3:1999", "ISO 8573-7:2003"],
        "equipment": {
            "dust_regular": "Máy đo nồng độ hạt AeroTrak 9350 và phụ kiện HPD III, USA;",
            "dust_fine": "Máy đo nồng độ hạt bụi Lasair III 110 và phụ kiện HPD II, USA;",
            "dew_point": "Thiết bị đo nhiệt độ điểm sương;",
            # [11/11/2026] dụng cụ tổng dầu mới (2 dòng thiết bị)
            "oil": "Bộ dụng cụ lấy mẫu tổng dầu trong khí nén;\n- Giấy lọc thủy tinh Pall life Sciences, ống than hoạt tính ORBO®.",
            "micro": "Thiết bị lấy mẫu vi sinh khí nén SAS Pinocchio Super II."
        }
    },
    # --- TỦ AN TOÀN SINH HỌC & TỦ SẠCH & TỦ HÚT ---
    "N2.TC.BSC.02": {
        "code": "N2.TC.BSC.02",
        "name": "Đo tủ an toàn sinh học cấp II",
        "standards": ["EN 12469:2000", "ISO 14644-1:2015", "ISO 12464-1:2011", "ISO 11201:1995", "ISO 3744"],
        "equipment": {
            "inflow": "Thiết bị đo lưu lượng TSI 8375/8380, USA;",
            "downflow": "Máy đo tốc độ gió Kanomax Climomaster 6501, Japan;",
            "hepa_leak": "Máy tạo sol khí ATI aerosol generator, model 6D và quang kế aerosol photometer, model 2i, USA;",
            "smoke": "Máy tạo khói khí dung (Smoke pencil / Aerosol generator);",
            "lux": "Máy đo độ rọi Testo 540 / Testo 440, Germany;",
            "noise": "Máy đo độ ồn Rion NL-53, Japan;",
            "vibration": "Máy đo độ rung RION VM-85, Japan;",
            "uvc": "Máy đo cường độ bức xạ UVC Extech SDL470, USA;",
            "particle": "Máy đếm hạt TSI AeroTrak 9350, USA."
        }
    },
    "N2.TC.BSC.01": {
        "code": "N2.TC.BSC.01",
        "name": "Đo tủ an toàn sinh học cấp I",
        "standards": ["EN 12469:2000", "ISO 12464-1:2011", "ISO 11201:1995", "ISO 3744"],
        "equipment": {
            "inflow": "Máy đo tốc độ gió Kanomax Climomaster 6501, Japan;",
            "lux": "Máy đo độ rọi Testo 540, Germany;",
            "noise": "Máy đo độ ồn Rion NL-53, Japan;"
        }
    },
    "N2.TC.LAF.01": {
        "code": "N2.TC.LAF.01",
        "name": "Đo tủ sạch",
        "standards": ["ISO 14644-1:2015", "TCVN 6915:2001", "EN 12469"],
        "equipment": {
            "velocity": "Máy đo tốc độ gió Kanomax Climomaster 6501, Japan;",
            "hepa_leak": "Máy tạo sol khí ATI model 6D và quang kế 2i, USA;",
            "particle": "Máy đếm hạt TSI AeroTrak 9350, USA;",
            "lux": "Máy đo độ rọi Testo 540, Germany;",
            "noise": "Máy đo độ ồn Rion NL-53, Japan."
        }
    },
    "N2.TC.T.01": {
        "code": "N2.TC.T.01",
        "name": "Đo tủ hút hóa chất",
        "standards": ["EN 14175", "TCVN 6914", "TCVN 6915"],
        "equipment": {
            "face_velocity": "Máy đo tốc độ gió Kanomax Climomaster 6501, Japan;",
            "air_flow": "Thiết bị đo lưu lượng TSI 8375/8380, USA;",
            "smoke": "Thiết bị tạo khói kiểm tra hình thái dòng khí;",
            "lux": "Máy đo độ sáng Testo 440, Germany;",
            "noise": "Máy đo độ ồn Rion NL-53, Japan."
        }
    },
    # --- MAPPING NHIỆT ĐỘ (KHO LẠNH, TỦ BẢO QUẢN) ---
    "N2.TC.MP.01": {
        "code": "N2.TC.MP.01",
        "name": "Mapping nhiệt độ, độ ẩm",
        "standards": ["QUATEST 3 3030:2025", "USP-NF (1079.4)"],
        "equipment": {
            "mapping_logger": "Thiết bị tự ghi nhiệt độ độ ẩm Haxo-8 (mapping);",
            "verify_meter": "Thiết bị đo nhiệt độ độ ẩm Testo 625, Germany."
        }
    },
    # --- LƯU LƯỢNG GIÓ ĐƯỜNG ỐNG ---
    "N2.TC.G.01": {
        "code": "N2.TC.G.01",
        "name": "Đo lưu lượng gió đường ống",
        "standards": ["ISO 10780"],
        "equipment": {
            "velocity_tube": "Thiết bị đo vận tốc gió VelociCal 8345, TSI – USA;"
        }
    },
    # --- VI KHÍ HẬU & MÔI TRƯỜNG LAO ĐỘNG (MTLĐ) ---
    "N2.TC.01": {
        "code": "N2.TC.01",
        "name": "Đo vi khí hậu và độ sáng",
        "standards": ["QCVN 26:2016/BYT", "QCVN 22:2016/BYT", "TCVN 5508:2009", "TCVN 5176:1990"],
        "equipment": {
            "temp_hum": "Máy đo nhiệt độ, độ ẩm Testo 625, Germany;",
            "air_speed": "Máy đo tốc độ gió Kanomax Climomaster 6533-2G, Japan;",
            "lux": "Máy đo độ sáng Testo 440 Kit Lux, Testo – Germany;",
            "thermal_rad": "Máy đo bức xạ nhiệt cầm tay skSATO SK-170GT, Japan."
        }
    },
    "N2.TC.02": {
        "code": "N2.TC.02",
        "name": "Đo tiếng ồn",
        "standards": ["QCVN 24:2016/BYT", "TCVN 9799:2013", "ISO 9612:2009"],
        "equipment": {
            "noise": "Thiết bị đo mức âm thanh Rion NL-53, Japan."
        }
    },
    "N2.TC.03": {
        "code": "N2.TC.03",
        "name": "Đo độ rung",
        "standards": ["QCVN 27:2016/BYT", "TCVN 5127:1990", "ISO 2631-1:1997"],
        "equipment": {
            "vibration": "Máy đo độ rung Rion VM-85, Japan."
        }
    },
    "N2.TC.04": {
        "code": "N2.TC.04",
        "name": "Đo bức xạ tử ngoại (UV)",
        "standards": ["QCVN 23:2016/BYT"],
        "equipment": {
            "uv": "Máy đo bức xạ tử ngoại UV Extech SDL470, USA."
        }
    },
    "N2.TC.05": {
        "code": "N2.TC.05",
        "name": "Đo điện từ trường",
        "standards": ["QCVN 25:2016/BYT"],
        "equipment": {
            "emf": "Thiết bị đo điện từ trường tần số công nghiệp HI-3604, Holaday Industries – USA."
        }
    },
    "N2.TC.06": {
        "code": "N2.TC.06",
        "name": "Lấy mẫu bụi môi trường lao động",
        "standards": ["QCVN 02:2019/BYT", "TCVN 5067:1995"],
        "equipment": {
            "dust_sampler": "Bơm lấy mẫu bụi cá nhân SKC / Buck Genie và cassette màng lọc;",
            "dust_direct": "Máy đo bụi trực tiếp Aerocet 831, USA."
        }
    },
    "N2.TC.07": {
        "code": "N2.TC.07",
        "name": "Lấy mẫu khí môi trường lao động",
        "standards": ["QCVN 03:2019/BYT"],
        "equipment": {
            "gas_sampler": "Bơm lấy mẫu khí cá nhân, ống hấp phụ than hoạt / impinger dung dịch hấp thụ;",
            "gas_co2": "Thiết bị đo nồng độ CO2 trong không khí Testo 440, Germany;",
            "gas_tvoc": "Thiết bị đo nồng độ TVOC Dräger X-am 8000, Germany."
        }
    },
    "N2.TC.08": {
        "code": "N2.TC.08",
        "name": "Phương pháp test hơi khí độc bằng ống phát hiện",
        "standards": ["QCVN 03:2019/BYT"],
        "equipment": {
            "detector_tube": "Bơm hút khí Gastec / Kitagawa và ống phát hiện nhanh chuyên dụng."
        }
    },
    # --- MÔI TRƯỜNG NƯỚC ---
    "N2.TC.N.01": {
        "code": "N2.TC.N.01",
        "name": "Hướng dẫn lấy mẫu nước cấp",
        "standards": ["QCVN 01-1:2024/BYT", "TCVN 6663-1:2011", "TCVN 6663-3:2016", "ISO 19458"],
        "equipment": {
            "water_sample": "Dụng cụ lấy mẫu và bảo quản mẫu nước vô trùng chuyên dụng;"
        }
    },
    "N2.TC.N.02": {
        "code": "N2.TC.N.02",
        "name": "Đo các thông số môi trường nước tại hiện trường",
        "standards": ["QCVN 08:2023/BTNMT", "QCVN 09:2023/BTNMT", "QCVN 14:2008/BTNMT", "QCVN 40:2011/BTNMT"],
        "equipment": {
            "water_onsite": "Thiết bị đo đa chỉ tiêu nước tại hiện trường (pH, DO, EC, TDS, nhiệt độ nước)."
        }
    }
}

# ==============================================================================
# DANH MỤC CẤU HÌNH BIỂU MẪU (CATEGORIES)
# Tương thích ngược hoàn toàn và bổ sung các danh mục mở rộng
# ==============================================================================
CATEGORIES = {
    "phong_sach": {
        "id": "phong_sach",
        "name": "Phòng sạch, Vi sinh",
        "folder_name": "Phong sach, Vi sinh",
        "default_request": "Quan trắc chất lượng phòng sạch và lấy mẫu vi sinh trong không khí",
        "standard_equipment": (
            "- Thiết bị đo nồng độ hạt TSI AeroTrak 9350, USA;\n"
            "- Thiết bị đo nhiệt độ, độ ẩm trong không khí Testo 625, Germany;\n"
            "- Thiết bị đo chênh lệch áp suất TSI 8375, USA;\n"
            "- Thiết bị đo độ ồn Rion NL 53, Japan;\n"
            "- Thiết bị lấy mẫu vi sinh trong không khí Buck – Bio Culture B30120, USA."
        ),
        "standard_procedure": "QUATEST 3 3030:2025; N2.TC.09, N2.TC.PS.01",
        "template_bm02_source": os.path.join(BASE_DIR, "Phong sach, Vi sinh", "BM3030.02- BBLM - Q326D108033.docx"),
        "template_bm01_source": os.path.join(BASE_DIR, "Phong sach, Vi sinh", "BM3030.01- Ghi chép - 116326.docx"),
        "indicators": [
            {"name": "Đo nồng độ bụi phòng sạch", "is_measure": True, "is_sample": False},
            {"name": "Đo nhiệt độ, độ ẩm phòng sạch", "is_measure": True, "is_sample": False},
            {"name": "Đo chênh áp phòng sạch", "is_measure": True, "is_sample": False},
            {"name": "Đo độ ồn phòng sạch", "is_measure": True, "is_sample": False},
            {"name": "Lấy mẫu vi sinh trong không khí phòng sạch", "is_measure": False, "is_sample": True},
            {"name": "Đo lưu lượng và tính toán số lần trao đổi khí phòng sạch", "is_measure": True, "is_sample": False},
        ]
    },
    "khi_nen": {
        "id": "khi_nen",
        "name": "Khí nén, Khí Nitơ",
        "folder_name": "Khi nen",
        "default_request": "Lấy mẫu tổng dầu, đo nồng độ bụi và nhiệt độ điểm sương trong khí nén",
        "standard_equipment": (
            "- Máy đo nồng độ hạt AeroTrak 9350 và phụ kiện HPD III, USA;\n"
            "- Thiết bị đo nhiệt độ điểm sương;\n"
            "- Bộ dụng cụ lấy mẫu tổng dầu trong khí nén;\n"
            "- Giấy lọc thủy tinh Pall life Sciences, ống than hoạt tính ORBO®."
        ),
        "standard_procedure": "QUATEST 3 3030:2025; N2.TC.KN.01",
        # [FIX 11/11/2026] Mẫu BBLM khí nén cũ (Khi nen/BM3030.02- BBLM - Q326D106978.docx) có các đoạn viền chấm
        # làm biên bản nhảy sang Trang 3 -> dùng mẫu BBLM phòng sạch (bố cục 2 trang chuẩn Quy tắc 7).
        "template_bm02_source": os.path.join(BASE_DIR, "Phong sach, Vi sinh", "BM3030.02- BBLM - Q326D108033.docx"),
        "template_bm01_source": os.path.join(BASE_DIR, "Khi nen", "BM3030.01- Ghi chép - Q326D122266.docx"),
        "indicators": [
            {"name": "Lấy mẫu tổng dầu trong khí nén", "is_measure": False, "is_sample": True},
            {"name": "Đo nhiệt độ điểm sương trong khí nén", "is_measure": True, "is_sample": False},
            {"name": "Đo nồng độ bụi trong khí nén", "is_measure": True, "is_sample": False},
            {"name": "Lấy mẫu vi sinh trong khí nén", "is_measure": False, "is_sample": True},
            {"name": "Lấy mẫu khí Nitơ", "is_measure": False, "is_sample": True},
        ]
    },
    "co2_mtld": {
        "id": "co2_mtld",
        "name": "CO2 - Môi trường lao động",
        "folder_name": "CO2 - MTLD",
        "default_request": "Quan trắc chất lượng không khí môi trường lao động: TVOC và CO2",
        "standard_equipment": (
            "- Thiết bị đo nồng độ CO2 trong không khí testo 440, CO2 probe, Germany;\n"
            "- Thiết bị đo nồng độ TVOC trong không khí Drager Xam 8000, Gemany."
        ),
        "standard_procedure": "QUATEST 3 3030:2025; N2.TC.07",
        "template_bm02_source": os.path.join(BASE_DIR, "CO2 - MTLD", "BM3030.02- BBLM - Q326D091792.docx"),
        "template_bm01_source": os.path.join(BASE_DIR, "CO2 - MTLD", "BM3030.01- Ghi chép - Q326D091792.docx"),
        "indicators": [
            {"name": "Đo nồng độ TVOC trong không khí môi trường lao động", "is_measure": True, "is_sample": False},
            {"name": "Đo nồng độ CO2 trong không khí môi trường lao động", "is_measure": True, "is_sample": False},
        ]
    },
    "kho_lanh": {
        "id": "kho_lanh",
        "name": "Kho lạnh (Mapping nhiệt độ/độ ẩm)",
        "folder_name": "Kho lanh",
        "default_request": "Đánh giá độ đồng đều nhiệt độ kho bảo quản thuốc",
        "standard_equipment": (
            "- Thiết bị đo nhiệt độ độ ẩm Testo 625, Germany;\n"
            "- Thiết bị tự ghi nhiệt độ độ ẩm Haxo-8 (mapping)."
        ),
        "standard_procedure": "QUATEST 3 3030:2025; N2.TC.MP.01",
        "template_bm02_source": os.path.join(BASE_DIR, "Kho lanh", "BM3030.02- BBLM - Q326D025841.docx"),
        "template_bm01_source": os.path.join(BASE_DIR, "Kho lanh", "BM3030.01- Ghi chép giám định - Q326D029160.docx"),
        "indicators": [
            {"name": "Đánh giá độ đồng đều nhiệt độ kho bảo quản thuốc", "is_measure": True, "is_sample": False},
            {"name": "Đánh giá độ đồng đều độ ẩm kho bảo quản thuốc", "is_measure": True, "is_sample": False},
        ]
    },
    "bsc": {
        "id": "bsc",
        "name": "Tủ an toàn sinh học (BSC)",
        "folder_name": "BSC",
        "default_request": "Giám định an toàn và hiệu năng tủ an toàn sinh học",
        "standard_equipment": (
            "- Thiết bị đo lưu lượng TSI 8375/8380, USA;\n"
            "- Máy đo tốc độ gió Kanomax Climomaster 6501, Japan;\n"
            "- Máy tạo sol khí ATI aerosol generator, model 6D và quang kế aerosol photometer, model 2i, USA;\n"
            "- Máy tạo khói khí dung (Smoke pencil / Aerosol generator);\n"
            "- Máy đo độ rọi Testo 540 / Testo 440, Germany;\n"
            "- Máy đo độ ồn Rion NL-53, Japan;\n"
            "- Máy đo độ rung RION VM-85, Japan;\n"
            "- Máy đo cường độ bức xạ UVC Extech SDL470, USA;\n"
            "- Máy đếm hạt TSI AeroTrak 9350, USA."
        ),
        "standard_procedure": "QUATEST 3 3030:2025; EN 12469:2000; N2.TC.BSC.02",
        "template_bm02_source": os.path.join(BASE_DIR, "BSC", "BM3030.02- Biên bản - 45837.docx"),
        "template_bm01_source": os.path.join(BASE_DIR, "BSC", "BM3030.01- Ghi chep - BSC.docx"),
        "indicators": [
            {"name": "Đo tốc độ trung bình dòng khí qua cửa làm việc (inflow) tủ an toàn sinh học", "is_measure": True, "is_sample": False},
            {"name": "Đo tốc độ trung bình dòng khí cấp qua lọc HEPA (downflow) tủ an toàn sinh học", "is_measure": True, "is_sample": False},
            {"name": "Đo cường độ sáng trung bình bên trong khu vực làm việc tủ an toàn sinh học", "is_measure": True, "is_sample": False},
            {"name": "Đo độ ồn tủ an toàn sinh học", "is_measure": True, "is_sample": False},
            {"name": "Kiểm tra độ rò rỉ của các lọc HEPA tủ an toàn sinh học", "is_measure": True, "is_sample": False},
            {"name": "Kiểm tra hình thái dòng khí tủ an toàn sinh học", "is_measure": True, "is_sample": False},
            {"name": "Đo cường độ UVC tủ an toàn sinh học", "is_measure": True, "is_sample": False},
            {"name": "Đo nồng độ bụi trong tủ an toàn sinh học", "is_measure": True, "is_sample": False},
        ]
    },
    "tu_laf": {
        "id": "tu_laf",
        "name": "Tủ sạch (LAF)",
        "folder_name": "BSC",
        "default_request": "Đo kiểm tra hiệu năng tủ sạch dòng khí tầng (LAF)",
        "standard_equipment": (
            "- Máy đo tốc độ gió Kanomax Climomaster 6501, Japan;\n"
            "- Máy tạo sol khí ATI model 6D và quang kế 2i, USA;\n"
            "- Máy đếm hạt TSI AeroTrak 9350, USA;\n"
            "- Máy đo độ rọi Testo 540, Germany;\n"
            "- Máy đo độ ồn Rion NL-53, Japan."
        ),
        "standard_procedure": "QUATEST 3 3030:2025; N2.TC.LAF.01; ISO 14644-1:2015",
        "template_bm02_source": os.path.join(BASE_DIR, "BSC", "BM3030.02- Biên bản - 45837.docx"),
        "template_bm01_source": os.path.join(BASE_DIR, "BSC", "BM3030.01- Ghi chep - BSC.docx"),
        "indicators": [
            {"name": "Đo vận tốc dòng khí cấp (downflow) tủ LAF", "is_measure": True, "is_sample": False},
            {"name": "Kiểm tra rò rỉ của các lọc HEPA tủ LAF", "is_measure": True, "is_sample": False},
            {"name": "Đo nồng độ bụi trong tủ LAF", "is_measure": True, "is_sample": False},
            {"name": "Đo cường độ sáng trung bình bên trong khu vực làm việc tủ LAF", "is_measure": True, "is_sample": False},
            {"name": "Đo độ ồn tủ LAF", "is_measure": True, "is_sample": False},
        ]
    },
    "tu_hut": {
        "id": "tu_hut",
        "name": "Tủ hút khí độc (Fume Hood)",
        "folder_name": "BSC",
        "default_request": "Đo kiểm tra các yếu tố môi trường của tủ hút khí độc",
        "standard_equipment": (
            "- Thiết bị đo độ sáng Testo 440, Germany;\n"
            "- Máy đo tốc độ gió Climomaster Model 6501, Japan;\n"
            "- Máy đo độ ồn Rion NL-53, Japan;\n"
            "- Thiết bị tạo khói kiểm tra hình thái dòng khí."
        ),
        "standard_procedure": "QUATEST 3 3030:2025; N2.TC.T.01; EN 14175",
        "template_bm02_source": os.path.join(BASE_DIR, "BSC", "BM3030.02- Biên bản - Q326D008933.docx"),
        "template_bm01_source": os.path.join(BASE_DIR, "BSC", "BM3030.01- Ghi chep - BSC.docx"),
        "indicators": [
            {"name": "Đo vận tốc dòng khí qua cửa làm việc tủ hút", "is_measure": True, "is_sample": False},
            {"name": "Kiểm tra hình thái dòng khí tủ hút", "is_measure": True, "is_sample": False},
            {"name": "Đo cường độ sáng bên trong khu vực làm việc tủ hút", "is_measure": True, "is_sample": False},
            {"name": "Đo độ ồn tủ hút", "is_measure": True, "is_sample": False},
        ]
    },
    "vi_khi_hau_mtld": {
        "id": "vi_khi_hau_mtld",
        "name": "Vi khí hậu & Môi trường lao động (MTLĐ)",
        "folder_name": "CO2 - MTLD",
        "default_request": "Quan trắc vi khí hậu, tiếng ồn, ánh sáng, bụi và các yếu tố môi trường lao động",
        "standard_equipment": (
            "- Máy đo nhiệt độ, độ ẩm Testo 625, Germany;\n"
            "- Máy đo tốc độ gió Kanomax Climomaster 6533-2G, Japan;\n"
            "- Máy đo độ sáng Testo 440 Kit Lux, Testo – Germany;\n"
            "- Thiết bị đo mức âm thanh Rion NL-53, Japan;\n"
            "- Bơm lấy mẫu bụi cá nhân SKC / Buck Genie và cassette màng lọc."
        ),
        "standard_procedure": "QUATEST 3 3030:2025; N2.TC.01, N2.TC.02, N2.TC.06",
        "template_bm02_source": os.path.join(BASE_DIR, "CO2 - MTLD", "BM3030.02- BBLM - Q326D091792.docx"),
        "template_bm01_source": os.path.join(BASE_DIR, "CO2 - MTLD", "BM3030.01- Ghi chép - Q326D091792.docx"),
        "indicators": [
            {"name": "Đo vi khí hậu môi trường lao động", "is_measure": True, "is_sample": False},
            {"name": "Đo độ ồn môi trường lao động", "is_measure": True, "is_sample": False},
            {"name": "Đo độ sáng môi trường lao động", "is_measure": True, "is_sample": False},
            {"name": "Lấy mẫu bụi toàn phần môi trường lao động", "is_measure": False, "is_sample": True},
            {"name": "Lấy mẫu bụi hô hấp môi trường lao động", "is_measure": False, "is_sample": True},
        ]
    }
}


def infer_method_type(indicator_name):
    """
    Suy luận phương pháp thực hiện (Đo vs Lấy mẫu) chuẩn xác dựa trên tên chỉ tiêu
    theo các quy trình kỹ thuật N2.TC của QUATEST 3.
    Trả về: (is_measure: bool, is_sample: bool)
    """
    name_low = indicator_name.lower().strip()
    
    # Danh sách các từ khóa đặc trưng cho "Lấy mẫu"
    sample_keywords = [
        "lấy mẫu", "thu mẫu", "vi sinh", "vsv", "khuẩn", "nấm",
        "tổng dầu", "hàm lượng dầu", "khí nito", "khí nén lấy mẫu",
        "màng lọc", "than hoạt", "impinger", "hấp thụ", "hấp phụ",
        "nước cấp", "nước mặt", "nước ngầm", "nước thải", "nước ro", "nước uống",
        "bùn", "trầm tích", "chất thải", "bụi toàn phần", "bụi hô hấp", "bụi cá nhân"
    ]
    
    # Các trường hợp đo trực tiếp rõ ràng (kể cả có từ gây nhiễu)
    measure_keywords = [
        "đo nồng độ", "đo nhiệt độ", "đo độ ẩm", "đo chênh áp", "đo áp suất",
        "đo độ ồn", "đo ồn", "đo rung", "đo độ rung", "đo vận tốc", "đo lưu lượng",
        "đo độ rọi", "đo độ sáng", "đo chiếu sáng", "kiểm tra rò rỉ", "quét rò rỉ", "rò rỉ", "kiểm tra hình thái",
        "đo uvc", "đo uv", "đo điện từ trường", "đo điểm sương", "đánh giá độ đồng đều",
        "mapping", "đo tốc độ", "inflow", "downflow", "face velocity"
    ]
    
    for mk in measure_keywords:
        if mk in name_low:
            return True, False
            
    for sk in sample_keywords:
        if sk in name_low:
            return False, True
            
    # Mặc định là Đo nếu không khớp mẫu lấy mẫu
    return True, False


def resolve_equipment_and_procedure(category_key, indicators=None, request_content=None, custom_equipment=None, custom_procedure=None):
    """
    Tự động xác định Thiết bị sử dụng và Quy trình áp dụng chuẩn xác
    dựa trên các chỉ tiêu thực tế và nội dung yêu cầu, tuân thủ nghiêm ngặt
    Hệ thống Quy tắc QUATEST 3 (Phần 3 & Phần 4).
    """
    if custom_equipment and custom_equipment.strip() and custom_procedure and custom_procedure.strip():
        return custom_equipment.strip(), custom_procedure.strip()
        
    category_cfg = CATEGORIES.get(category_key, CATEGORIES["phong_sach"])
    
    # 1. Tổng hợp text để phân tích từ indicators và request_content
    ind_texts = []
    if indicators:
        for ind in indicators:
            if isinstance(ind, dict):
                ind_texts.append(ind.get("name", ""))
            else:
                ind_texts.append(str(ind))
    all_ind_str = " ".join(ind_texts).lower()
    combined_info = f"{all_ind_str} {(request_content or '').lower()}"
    
    eq_list = []
    sop_codes = ["QUATEST 3 3030:2025"]
    
    # --- LOGIC XỬ LÝ THEO TỪNG LĨNH VỰC ---
    if category_key == "khi_nen":
        # Khí nén theo N2.TC.KN.01
        sop_codes.append("N2.TC.KN.01")
        
        has_dust = any(k in combined_info for k in ["bụi", "hạt", "particle", "tiểu phân"])
        has_dew = any(k in combined_info for k in ["điểm sương", "dew"])
        has_oil = any(k in combined_info for k in ["tổng dầu", "hàm lượng dầu", "dầu khoáng", "oil"])
        has_micro = any(k in combined_info for k in ["vi sinh", "micro", "vsv", "khuẩn", "nấm"])
        
        # Nếu không có chỉ tiêu cụ thể nào khớp, lấy theo default
        if not (has_dust or has_dew or has_oil or has_micro):
            has_dust = has_dew = has_oil = True
            
        if has_dust:
            if any(k in combined_info for k in ["0,1", "0.1"]):
                eq_list.append("- Máy đo nồng độ hạt bụi Lasair III 110 và phụ kiện HPD II, USA;")
            else:
                eq_list.append("- Máy đo nồng độ hạt AeroTrak 9350 và phụ kiện HPD III, USA;")
        if has_dew:
            eq_list.append("- Thiết bị đo nhiệt độ điểm sương;")
        if has_oil:
            eq_list.append("- Bộ dụng cụ lấy mẫu tổng dầu trong khí nén;")
            eq_list.append("- Giấy lọc thủy tinh Pall life Sciences, ống than hoạt tính ORBO®;")  # [11/11/2026]
        if has_micro:
            eq_list.append("- Thiết bị lấy mẫu vi sinh khí nén SAS Pinocchio Super II.")
            sop_codes.append("N2.TC.09")

    elif category_key == "phong_sach":
        # Phòng sạch theo N2.TC.PS.01
        sop_codes.append("N2.TC.PS.01")
        
        has_dust = any(k in combined_info for k in ["bụi", "hạt", "particle", "tiểu phân"])
        has_temp_hum = any(k in combined_info for k in ["nhiệt độ", "độ ẩm", "nhiệt ẩm", "vi khí hậu"])
        has_diff = any(k in combined_info for k in ["chênh áp", "áp suất", "chênh lệch áp"])
        has_noise = any(k in combined_info for k in ["độ ồn", "tiếng ồn", "mức âm", "noise"])
        has_lux = any(k in combined_info for k in ["độ sáng", "độ rọi", "chiếu sáng", "lux"])
        has_micro_surface = any(k in combined_info for k in ["bề mặt", "swab", "vsbm", "rodac"])
        has_micro_air = any(k in combined_info for k in ["vi sinh không khí", "vi sinh trong không khí", "vskk", "air micro"]) or (
            any(k in combined_info for k in ["vi sinh", "micro", "vsv", "khuẩn", "nấm"]) and not has_micro_surface
        )
        # [14/11/2026] Vận tốc gió tại miệng cấp gió -> máy đo tốc độ gió (Kanomax 6501),
        # lưu lượng / trao đổi khí -> TSI 8375
        has_flow = any(k in combined_info for k in ["lưu lượng", "trao đổi khí"])
        has_velocity = any(k in combined_info for k in ["vận tốc gió", "tốc độ gió"])
        # [11/11/2026] Quét rò rỉ màng lọc HEPA phòng sạch (ISO 14644-3) -> ATI 6D + 2i
        # [14/11/2026] Chỉ khi là rò rỉ; "chênh áp màng lọc HEPA" là đo chênh áp (TSI 8375), không phải quét rò rỉ
        has_hepa = any(k in combined_info for k in ["rò rỉ", "leak", "xuyên thấu"])
        # [14/11/2026] Vi sinh không khí phương pháp đặt đĩa (lắng) / vi sinh bề mặt đĩa tiếp xúc
        has_settle = any(k in combined_info for k in ["đặt đĩa", "đĩa lắng", "settle plate"])
        has_contact = any(k in combined_info for k in ["đĩa tiếp xúc", "contact plate", "rodac"])

        # Nếu không có chỉ tiêu cụ thể nào khớp, lấy danh mục chuẩn
        if not (has_dust or has_temp_hum or has_diff or has_noise or has_lux or has_micro_air or has_micro_surface or has_flow or has_velocity or has_hepa):
            has_dust = has_temp_hum = has_diff = has_noise = has_micro_air = True
            
        if has_dust:
            if any(k in combined_info for k in ["0,1", "0.1", "iso 1", "iso 2", "iso 3", "iso 4", "iso 5", "iso 6", "grade a", "grade b"]):
                eq_list.append("- Máy đo nồng độ hạt bụi PMS Lasair III 110, USA;")
            else:
                eq_list.append("- Thiết bị đo nồng độ hạt TSI AeroTrak 9350, USA;")
        if has_hepa:
            eq_list.append("- Máy tạo sol khí ATI aerosol generator, model 6D và quang kế aerosol photometer, model 2i, USA;")
        if has_temp_hum:
            eq_list.append("- Thiết bị đo nhiệt độ, độ ẩm trong không khí Testo 625, Germany;")
            if "N2.TC.01" not in sop_codes: sop_codes.append("N2.TC.01")
        if has_diff:
            eq_list.append("- Thiết bị đo chênh lệch áp suất TSI 8375, USA;")
        if has_noise:
            eq_list.append("- Thiết bị đo độ ồn Rion NL 53, Japan;")
            if "N2.TC.02" not in sop_codes: sop_codes.append("N2.TC.02")
        if has_lux:
            eq_list.append("- Thiết bị đo độ sáng Testo 440, Germany;")
            if "N2.TC.01" not in sop_codes: sop_codes.append("N2.TC.01")
        if has_flow:
            eq_list.append("- Thiết bị đo lưu lượng TSI 8375, USA;")
        if has_velocity:
            eq_list.append("- Máy đo tốc độ gió Kanomax Climomaster 6501, Japan;")
        if has_micro_air:
            if has_settle:
                eq_list.append("- Khay inox và đĩa môi trường lấy mẫu vi sinh không khí (phương pháp đặt đĩa);")
            else:
                eq_list.append("- Thiết bị lấy mẫu vi sinh trong không khí Buck – Bio Culture B30120, USA.")
            if "N2.TC.09" not in sop_codes: sop_codes.append("N2.TC.09")
        if has_micro_surface:
            if has_contact:
                eq_list.append("- Đĩa tiếp xúc Rodac lấy mẫu vi sinh bề mặt.")
            else:
                eq_list.append("- Dụng cụ lấy mẫu vi sinh bề mặt chuyên dụng.")
            if "N2.TC.09" not in sop_codes: sop_codes.append("N2.TC.09")

        # Nếu là hồ sơ chỉ thuần vi sinh bề mặt (không có đo phòng sạch dust/temp/diff...)
        if has_micro_surface and not (has_dust or has_temp_hum or has_diff or has_noise or has_lux or has_flow or has_velocity):
            sop_codes = ["QUATEST 3 3030:2025", "N2.TC.09"]

    elif category_key == "bsc":
        # Tủ an toàn sinh học theo EN 12469:2000; N2.TC.BSC.01 / N2.TC.BSC.02
        has_cap1 = bool(re.search(r'\bcấp\s*(?:1|i)\b|\bbsc[-\s]*(?:1|i)\b', combined_info))
        has_cap2 = bool(re.search(r'\bcấp\s*(?:2|ii)\b|\bbsc[-\s]*(?:2|ii)\b', combined_info))
        if has_cap1 and has_cap2:
            sop_codes = ["QUATEST 3 3030:2025", "EN 12469:2000", "N2.TC.BSC.01", "N2.TC.BSC.02"]
        elif has_cap1 and not has_cap2:
            sop_codes = ["QUATEST 3 3030:2025", "EN 12469:2000", "N2.TC.BSC.01"]
        else:
            sop_codes = ["QUATEST 3 3030:2025", "EN 12469:2000", "N2.TC.BSC.02"]
        
        has_inflow = any(k in combined_info for k in ["inflow", "hút", "cửa làm việc", "dòng khí (inflow)"])
        has_downflow = any(k in combined_info for k in ["downflow", "cấp", "thổi"])
        has_hepa = any(k in combined_info for k in ["hepa", "rò rỉ", "lọc"])
        has_smoke = any(k in combined_info for k in ["khói", "hình thái"])
        has_lux = any(k in combined_info for k in ["độ sáng", "độ rọi", "chiếu sáng", "cường độ sáng", "lux"])
        has_noise = any(k in combined_info for k in ["độ ồn", "tiếng ồn", "noise"])
        has_vib = any(k in combined_info for k in ["độ rung", "rung động", "vibration"])
        has_uv = any(k in combined_info for k in ["uv", "uvc", "tử ngoại"])
        has_dust = any(k in combined_info for k in ["hạt", "bụi", "particle", "tiểu phân"])
        
        if not (has_inflow or has_downflow or has_hepa or has_smoke or has_lux or has_noise or has_vib or has_uv or has_dust):
            has_inflow = has_downflow = has_hepa = has_smoke = has_lux = has_noise = has_vib = has_uv = has_dust = True
            
        if has_inflow:
            eq_list.append("- Thiết bị đo lưu lượng TSI 8375/8380, USA;")
        if has_downflow or (has_cap1 and has_inflow):
            eq_list.append("- Máy đo tốc độ gió Kanomax Climomaster 6501, Japan;")
        if has_hepa:
            eq_list.append("- Máy tạo sol khí ATI aerosol generator, model 6D và quang kế aerosol photometer, model 2i, USA;")
        if has_smoke:
            eq_list.append("- Máy tạo khói khí dung (Smoke pencil / Aerosol generator);")
        if has_lux:
            eq_list.append("- Máy đo độ rọi Testo 540 / Testo 440, Germany;")
        if has_noise:
            eq_list.append("- Máy đo độ ồn Rion NL-53, Japan;")
        if has_vib:
            eq_list.append("- Máy đo độ rung RION VM-85, Japan;")
        if has_uv:
            eq_list.append("- Máy đo cường độ bức xạ UVC Extech SDL470, USA;")
        if has_dust:
            eq_list.append("- Máy đếm hạt TSI AeroTrak 9350, USA.")

    elif category_key == "kho_lanh":
        sop_codes.append("N2.TC.MP.01")
        eq_list.append("- Thiết bị đo nhiệt độ độ ẩm Testo 625, Germany;")
        eq_list.append("- Thiết bị tự ghi nhiệt độ độ ẩm Haxo-8 (mapping).")

    elif category_key in ["co2_mtld", "vi_khi_hau_mtld"]:
        sop_codes.append("N2.TC.07")
        if any(k in combined_info for k in ["co2", "cacbon dioxit"]):
            eq_list.append("- Thiết bị đo nồng độ CO2 trong không khí testo 440, CO2 probe, Germany;")
        if any(k in combined_info for k in ["tvoc", "voc", "hữu cơ bay hơi"]):
            eq_list.append("- Thiết bị đo nồng độ TVOC trong không khí Drager Xam 8000, Gemany;")
        if any(k in combined_info for k in ["nhiệt độ", "độ ẩm", "vi khí hậu"]):
            eq_list.append("- Máy đo nhiệt độ, độ ẩm Testo 625, Germany;")
            if "N2.TC.01" not in sop_codes: sop_codes.append("N2.TC.01")
        if any(k in combined_info for k in ["ồn", "tiếng ồn"]):
            eq_list.append("- Thiết bị đo mức âm thanh Rion NL-53, Japan;")
            if "N2.TC.02" not in sop_codes: sop_codes.append("N2.TC.02")
        if any(k in combined_info for k in ["bụi"]):
            eq_list.append("- Bơm lấy mẫu bụi cá nhân SKC / Buck Genie và cassette màng lọc;")
            if "N2.TC.06" not in sop_codes: sop_codes.append("N2.TC.06")
        if not eq_list:
            eq_list.append("- Thiết bị đo nồng độ CO2 trong không khí testo 440, CO2 probe, Germany;")
            eq_list.append("- Thiết bị đo nồng độ TVOC trong không khí Drager Xam 8000, Gemany;")

    elif category_key == "khi_nen_phong_sach":
        # [FIX 27/10/2026] Hồ sơ Hybrid Khí nén + Phòng sạch: tách chỉ tiêu theo chủ thể
        # (bản cũ rơi vào fallback phòng sạch -> thiếu thiết bị khí nén và thêm thiết bị không được yêu cầu;
        #  'Đo nhiệt độ điểm sương' còn bị hiểu nhầm là chỉ tiêu nhiệt độ - độ ẩm phòng sạch)
        kn_texts = [t.lower() for t in ind_texts if any(k in t.lower() for k in ["khí nén", "khí có áp", "nitơ", "tổng dầu", "điểm sương"])]
        ps_texts = [t.lower() for t in ind_texts if t.lower() not in kn_texts]
        kn_str = " ".join(kn_texts)
        ps_str = " ".join(ps_texts)
        if not kn_texts and not ps_texts:
            kn_str = "bụi điểm sương tổng dầu"
            ps_str = "bụi"

        if kn_str:
            sop_codes.append("N2.TC.KN.01")
            if any(k in kn_str for k in ["bụi", "hạt", "particle", "tiểu phân"]):
                if any(k in kn_str for k in ["0,1", "0.1"]):
                    eq_list.append("- Máy đo nồng độ hạt bụi Lasair III 110 và phụ kiện HPD II, USA;")
                else:
                    eq_list.append("- Máy đo nồng độ hạt AeroTrak 9350 và phụ kiện HPD III, USA;")
            if any(k in kn_str for k in ["điểm sương", "dew"]):
                eq_list.append("- Thiết bị đo nhiệt độ điểm sương;")
            if any(k in kn_str for k in ["tổng dầu", "hàm lượng dầu", "oil"]):
                eq_list.append("- Bộ dụng cụ lấy mẫu tổng dầu trong khí nén;")
                eq_list.append("- Giấy lọc thủy tinh Pall life Sciences, ống than hoạt tính ORBO®;")  # [11/11/2026]
            if any(k in kn_str for k in ["vi sinh", "micro", "vsv"]):
                eq_list.append("- Thiết bị lấy mẫu vi sinh khí nén SAS Pinocchio Super II;")
                if "N2.TC.09" not in sop_codes: sop_codes.append("N2.TC.09")

        if ps_str:
            sop_codes.append("N2.TC.PS.01")
            if any(k in ps_str for k in ["bụi", "hạt", "particle", "tiểu phân"]):
                if any(k in ps_str for k in ["0,1", "0.1", "iso 5", "grade a", "grade b"]):
                    eq_list.append("- Máy đo nồng độ hạt bụi PMS Lasair III 110, USA;")
                else:
                    eq_list.append("- Thiết bị đo nồng độ hạt TSI AeroTrak 9350, USA;")
            if any(k in ps_str for k in ["nhiệt độ", "độ ẩm"]):
                eq_list.append("- Thiết bị đo nhiệt độ, độ ẩm trong không khí Testo 625, Germany;")
            if any(k in ps_str for k in ["chênh áp", "chênh lệch áp"]):
                eq_list.append("- Thiết bị đo chênh lệch áp suất TSI 8375, USA;")
            if any(k in ps_str for k in ["độ ồn", "tiếng ồn"]):
                eq_list.append("- Thiết bị đo độ ồn Rion NL 53, Japan;")
            if any(k in ps_str for k in ["độ sáng", "độ rọi"]):
                eq_list.append("- Thiết bị đo độ sáng Testo 440, Germany;")
            if any(k in ps_str for k in ["vi sinh", "vskk"]):
                eq_list.append("- Thiết bị lấy mẫu vi sinh trong không khí Buck – Bio Culture B30120, USA;")
                if "N2.TC.09" not in sop_codes: sop_codes.append("N2.TC.09")

        # Dấu câu chuẩn: ';' cuối mỗi dòng, '.' ở dòng cuối cùng
        eq_list = [re.sub(r'[;.]\s*$', '', e) + ';' for e in eq_list]
        if eq_list:
            eq_list[-1] = eq_list[-1][:-1] + '.'

    else:
        # Fallback to category standard equipment
        eq_str = category_cfg.get("standard_equipment", "")
        return (custom_equipment.strip() if custom_equipment else eq_str), (custom_procedure.strip() if custom_procedure else category_cfg.get("standard_procedure", "QUATEST 3 3030:2025"))

    # Kết luận danh sách thiết bị (chuẩn hóa dấu ';' cuối dòng và '.' ở dòng cuối cùng)
    if custom_equipment and custom_equipment.strip():
        final_equipment = custom_equipment.strip()
    elif eq_list:
        eq_list = [re.sub(r'[;.]\s*$', '', e) + ';' for e in eq_list]
        eq_list[-1] = eq_list[-1][:-1] + '.'
        final_equipment = "\n".join(eq_list)
    else:
        final_equipment = category_cfg.get("standard_equipment", "")

    # Kết luận quy trình
    if custom_procedure and custom_procedure.strip():
        final_procedure = custom_procedure.strip()
    else:
        # Giữ thứ tự và loại bỏ trùng lặp
        unique_sops = []
        for s in sop_codes:
            if s not in unique_sops:
                unique_sops.append(s)
        final_procedure = "; ".join(unique_sops)

    return final_equipment, final_procedure
