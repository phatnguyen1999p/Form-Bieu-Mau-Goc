"""
Module chuyên trách xuất PDF và xử lý ghép PDF vector chất lượng cao
Sử dụng Microsoft Word COM (win32com.client) và PyMuPDF (fitz).
"""
import os
import sys
import shutil
import docx
import fitz
import win32com.client

from .config import BASE_DIR
from .normalizer import doc_number_for_filename, sanitize_for_filename

def convert_docx_to_pdf(docx_path, pdf_path=None):
    """
    Chuyển đổi tệp .docx sang .pdf thông qua Microsoft Word COM.
    Đảm bảo 100% độ chính xác về layout, ngắt trang và font chữ.
    Trả về: (pdf_path, page_count)
    """
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"Không tìm thấy file docx: {docx_path}")

    if pdf_path is None:
        pdf_path = os.path.splitext(docx_path)[0] + ".pdf"

    abs_docx = os.path.abspath(docx_path)
    abs_pdf = os.path.abspath(pdf_path)

    word = None
    doc_word = None
    page_count = 0
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc_word = word.Documents.Open(abs_docx)
        doc_word.SaveAs2(abs_pdf, FileFormat=17) # 17 = wdFormatPDF
        try:
            page_count = doc_word.ComputeStatistics(2) # 2 = wdStatisticPages
        except Exception:
            page_count = 0
    finally:
        if doc_word is not None:
            try:
                doc_word.Close(SaveChanges=0)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass

    return abs_pdf, page_count


# [NEW 11/11/2026] Biên bản (BM3030.02) được lập thành 2 bản -> file PDF chứa sẵn 2 bản liên tiếp để in một lần.
BM02_PDF_COPIES = 2


def repeat_pdf_copies(pdf_path, copies=BM02_PDF_COPIES):
    """
    Lặp lại toàn bộ nội dung PDF `copies` lần trong cùng file (vector, không giảm chất lượng).
    Nếu mỗi bản có số trang lẻ thì chèn 1 trang trắng sau mỗi bản để bản sau luôn bắt đầu ở tờ mới khi in 2 mặt.
    Trả về tổng số trang của file sau khi lặp.
    """
    copies = max(1, int(copies or 1))
    src = fitz.open(pdf_path)
    n = len(src)
    if copies == 1 or n == 0:
        src.close()
        return n
    out = fitz.open()
    for i in range(copies):
        out.insert_pdf(src)
        if n % 2 == 1 and i < copies - 1:
            r = src[-1].rect
            out.new_page(width=r.width, height=r.height)
    src.close()
    tmp = pdf_path + ".tmp"
    out.save(tmp, garbage=3, deflate=True)
    total = len(out)
    out.close()
    os.replace(tmp, pdf_path)
    return total


def merge_bsc_vector_pdf(p1_pdf_path, bsc_source_pdf, output_pdf_path, num_positions=1):
    """
    Ghép trang 1 PDF với 3 trang vector gốc của Ghi Chep BSC.pdf (nhân theo số lượng tủ num_positions)
    tạo thành PDF hoàn chỉnh (1 + 3*N trang) với chất lượng vector nét cao 100%.
    """
    pos_count = max(1, int(num_positions or 1))
    merged_pdf = fitz.open()
    p1_fitz = fitz.open(p1_pdf_path)
    merged_pdf.insert_pdf(p1_fitz)
    bsc_fitz = fitz.open(bsc_source_pdf)
    for _ in range(pos_count):
        merged_pdf.insert_pdf(bsc_fitz)
    os.makedirs(os.path.dirname(output_pdf_path), exist_ok=True)
    merged_pdf.save(output_pdf_path)
    total_pages = len(merged_pdf)
    merged_pdf.close()
    p1_fitz.close()
    bsc_fitz.close()
    return output_pdf_path, total_pages


def export_order_pdfs(results_dict, customer_info=None):
    """
    Tự động xuất cả bộ hồ sơ (BM3030.01 và BM3030.02) sang PDF:
    - BM3030.02: Chuyển đổi trực tiếp qua Word COM.
    - BM3030.01:
      + Nếu là BSC: Soạn riêng Trang 1 vector, xuất PDF rồi ghép với 3 trang vector gốc Ghi Chep BSC.pdf × số tủ.
      + Nếu là loại khác: Chuyển đổi trực tiếp qua Word COM.
    """
    target_dir = results_dict.get("output_dir")
    doc_num_raw = results_dict.get("doc_number", "")
    doc_num = doc_number_for_filename(doc_num_raw)
    short_kh = sanitize_for_filename(results_dict.get("short_customer_name", ""))
    short_content = str(results_dict.get("short_content", "")).lower()
    category = str(results_dict.get("category", "")).lower()
    num_positions = int(
        results_dict.get("num_positions")
        or (customer_info.get("num_positions") if customer_info else 1)
        or 1
    )

    is_bsc = (short_content == "bsc" or category == "bsc" or "tủ an toàn" in category)

    # 1. Xuất PDF cho BM3030.02
    bm02_docx = results_dict.get("bm02_file")
    if bm02_docx and os.path.exists(bm02_docx):
        bm02_pdf = os.path.join(target_dir, f"{doc_num} - BM3030.02 - BBLM - {short_kh}.pdf")
        try:
            pdf_path, p_cnt = convert_docx_to_pdf(bm02_docx, bm02_pdf)
            results_dict["bm02_pdf"] = pdf_path
            results_dict["bm02_page_count"] = p_cnt  # số trang của 1 bản (kiểm tra chuẩn 2 trang)
            # [NEW 11/11/2026] PDF biên bản chứa sẵn 2 bản (file .docx vẫn giữ 1 bản để chỉnh sửa)
            results_dict["bm02_pdf_copies"] = BM02_PDF_COPIES
            results_dict["bm02_pdf_total_pages"] = repeat_pdf_copies(pdf_path, BM02_PDF_COPIES)
        except Exception as e:
            results_dict["errors"].append(f"Lỗi xuất PDF BM3030.02: {str(e)}")

    # 2. Xuất PDF cho BM3030.01
    bm01_docx = results_dict.get("bm01_file")
    if bm01_docx and os.path.exists(bm01_docx):
        bm01_pdf = os.path.join(target_dir, f"{doc_num} - BM3030.01 - Ghi chep - {short_kh}.pdf")
        if is_bsc and customer_info:
            # Xử lý ghép vector nét cao cho BSC
            from .bm01_builder import prepare_bsc_temp_p1
            temp_p1_docx = os.path.join(BASE_DIR, "scratch", f"temp_p1_bm01_{doc_num}.docx")
            temp_p1_pdf = os.path.join(BASE_DIR, "scratch", f"temp_p1_bm01_{doc_num}.pdf")
            pdf_bsc_source = os.path.join(BASE_DIR, "BSC", "Ghi Chep BSC.pdf")
            try:
                prepare_bsc_temp_p1(customer_info, temp_p1_docx)
                convert_docx_to_pdf(temp_p1_docx, temp_p1_pdf)
                _, total_bsc_pages = merge_bsc_vector_pdf(
                    temp_p1_pdf, pdf_bsc_source, bm01_pdf, num_positions=num_positions
                )
                results_dict["bm01_pdf"] = bm01_pdf
                results_dict["bm01_page_count"] = total_bsc_pages
            except Exception as e:
                # Fallback qua convert trực tiếp
                try:
                    pdf_path, p_cnt = convert_docx_to_pdf(bm01_docx, bm01_pdf)
                    results_dict["bm01_pdf"] = pdf_path
                    results_dict["bm01_page_count"] = p_cnt
                except Exception as e2:
                    results_dict["errors"].append(f"Lỗi xuất PDF BM3030.01 BSC: {str(e2)}")
        else:
            # Xuất trực tiếp cho các lĩnh vực khác
            try:
                pdf_path, p_cnt = convert_docx_to_pdf(bm01_docx, bm01_pdf)
                results_dict["bm01_pdf"] = pdf_path
                results_dict["bm01_page_count"] = p_cnt
            except Exception as e:
                results_dict["errors"].append(f"Lỗi xuất PDF BM3030.01: {str(e)}")

    # Đồng bộ các tệp PDF sang thư mục ổ Z nếu đã có bản sao lưu trước đó
    z_copy_dir = results_dict.get("z_drive_copy")
    if z_copy_dir and os.path.exists(z_copy_dir):
        for pdf_key in ("bm02_pdf", "bm01_pdf"):
            p_file = results_dict.get(pdf_key)
            if p_file and os.path.exists(p_file):
                try:
                    shutil.copy2(p_file, os.path.join(z_copy_dir, os.path.basename(p_file)))
                except Exception as e:
                    print(f"[WARN] Khong the dong bo PDF sang o Z ({z_copy_dir}): {e}")

    return results_dict
