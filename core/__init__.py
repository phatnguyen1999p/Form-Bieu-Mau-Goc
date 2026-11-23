"""
Gói module lõi (Core Package) cho hệ thống tự động sinh biểu mẫu QUATEST 3
Tuân thủ nghiêm ngặt 11 Quy tắc soạn thảo chuẩn mực (GEMINI.md).
"""

from .config import (
    BASE_DIR,
    OUTPUT_DIR,
    Z_DRIVE_DIR,
    CATEGORIES,
    SOP_DATABASE,
    resolve_equipment_and_procedure,
    infer_method_type
)

from .normalizer import (
    normalize_customer_name,
    normalize_address,
    normalize_location,
    normalize_date_str,
    clean_indicator_name,
    clean_text_no_status_or_quantitative,
    extract_short_customer_name,
    extract_short_content,
    build_output_folder_name,
    auto_generate_doc_number,
    auto_generate_date,
    get_default_execution_date_dotted,
    format_equipment_multiline,
    update_conversation_db_title
)

from .bm02_builder import build_bm3030_02
from .bm01_builder import build_bm3030_01
from .pdf_exporter import convert_docx_to_pdf, export_order_pdfs, merge_bsc_vector_pdf
from .order_profiles import ORDER_PROFILES, get_profile, detect_profile
from .form_cloner import clone_form_set, process_order, copy_order_to_z_drive

__all__ = [
    "BASE_DIR",
    "OUTPUT_DIR",
    "Z_DRIVE_DIR",
    "CATEGORIES",
    "SOP_DATABASE",
    "resolve_equipment_and_procedure",
    "infer_method_type",
    "normalize_customer_name",
    "normalize_address",
    "normalize_location",
    "normalize_date_str",
    "clean_indicator_name",
    "clean_text_no_status_or_quantitative",
    "extract_short_customer_name",
    "extract_short_content",
    "build_output_folder_name",
    "auto_generate_doc_number",
    "auto_generate_date",
    "get_default_execution_date_dotted",
    "format_equipment_multiline",
    "update_conversation_db_title",
    "build_bm3030_02",
    "build_bm3030_01",
    "convert_docx_to_pdf",
    "export_order_pdfs",
    "merge_bsc_vector_pdf",
    "ORDER_PROFILES",
    "get_profile",
    "detect_profile",
    "clone_form_set",
    "process_order",
    "copy_order_to_z_drive"
]
