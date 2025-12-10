import streamlit as st
import pandas as pd
import docx
import re
import pytesseract
from PIL import Image
import io

# ===========================================
# ========== MODULE XỬ LÝ KPCS ==============
# ===========================================

def extract_r2_from_heading(text):
    """
    Tự động nhận tiêu đề dạng 1.1 / 2.1 / 3.1 → Tên phát hiện (R2)
    """
    pattern = r"(\d+\.\d+)\s*[-:]?\s*(.*)"
    m = re.match(pattern, text.strip())
    if m:
        return m.group(2).strip()
    return text


def extract_word_paragraphs(doc):
    """
    Lấy toàn bộ đoạn từ Word (không OCR)
    """
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]


def extract_images_from_word(doc):
    """
    Trích ảnh từ file Word để đưa OCR
    """
    images = []
    rels = doc.part.rels

    for rel in rels:
        if "image" in rels[rel].target_ref:
            img = rels[rel]._target.blob
            images.append(Image.open(io.BytesIO(img)))

    return images


def run_ocr_on_images(images):
    """
    OCR toàn bộ ảnh → trả về text
    """
    text_blocks = []
    for img in images:
        text = pytesseract.image_to_string(img, lang="vie+eng")
        text_blocks.append(text)
    return text_blocks


def extract_4_regions(paragraphs, ocr_blocks):
    """
    Tách 4 vùng theo yêu cầu:
    1. R0 – R1
    2. R3
    3. Mô tả chi tiết
    4. Dẫn chiếu
    Các vùng còn lại lấy từ mark gạch chân / khoanh tròn trong OCR
    """

    r0_r1, r3, mo_ta, dan_chieu = "", "", "", ""

    # lấy các vùng từ Word
    for p in paragraphs:
        if "Nghiệp vụ" in p or "R0" in p:
            r0_r1 = p
        elif "Chi tiết phát hiện" in p or "R3" in p:
            r3 = p
        elif "Mô tả chi tiết" in p:
            mo_ta = p
        elif "Dẫn chiếu" in p:
            dan_chieu = p

    # OCR lấy thêm thông tin khoanh tròn / gạch chân
    ocr_text = "\n".join(ocr_blocks)

    return r0_r1, r3, mo_ta, dan_chieu, ocr_text


def build_kpcs_row(r0_r1, r3, mo_ta, dan_chieu, ocr_text, r2_title):
    """
    Mapping ĐỦ 43 cột KPCS
    """
    return {
        "STT": "",
        "Đối tượng được KT": "",
        "Số văn bản": "",
        "Ngày, tháng, năm ban hành (mm/dd/yyyy)": "",
        "Tên Đoàn kiểm toán": "",
        "Số hiệu rủi ro": "",
        "Số hiệu kiểm soát": "",
        "Nghiệp vụ (R0)": r0_r1,
        "Quy trình/hoạt động con (R1)": r0_r1,
        "Tên phát hiện (R2)": r2_title,
        "Chi tiết phát hiện (R3)": r3,
        "Dẫn chiếu": dan_chieu,
        "Mô tả chi tiết phát hiện": mo_ta,
        "CIF Khách hàng/bút toán": "",
        "Tên khách hàng": "",
        "Loại KH": "",
        "Số phát hiện/số mẫu chọn": "",
        "Dư nợ sai phạm": "",
        "Số tiền tổn thất": "",
        "Số tiền cần thu hồi": "",
        "Trách nhiệm trực tiếp": "",
        "Trách nhiệm quản lý": "",
        "Xếp hạng rủi ro": "",
        "Xếp hạng kiểm soát": "",
        "Nguyên nhân": ocr_text,
        "Ảnh hưởng": ocr_text,
        "Kiến nghị": ocr_text,
        "Loại/nhóm nguyên nhân": "",
        "Loại/nhóm ảnh hưởng": "",
        "Loại/nhóm kiến nghị": "",
        "Chủ thể kiến nghị": "",
        "Kế hoạch thực hiện": "",
        "Trách nhiệm thực hiện": "",
        "Đơn vị thực hiện KPCS": "",
        "ĐVKD, AMC, Hội sở": "",
        "Người phê duyệt": "",
        "Ý kiến của đơn vị": "",
        "Mức độ ưu tiên hành động": "",
        "Thời hạn hoàn thành": "",
        "Đã khắc phục": "",
        "Ngày đã KPCS": "",
        "CBKT (Mã CBKT-Họ tên)": ""
    }


def process_word_to_kpcs(doc_file):
    """
    Pipeline từ Word → OCR → Mapping 43 cột
    """
    doc = docx.Document(doc_file)

    paragraphs = extract_word_paragraphs(doc)
    images = extract_images_from_word(doc)
    ocr_blocks = run_ocr_on_images(images)

    r0_r1, r3, mo_ta, dan_chieu, ocr_text = extract_4_regions(paragraphs, ocr_blocks)

    # tìm tiêu đề dòng 1.1 / 2.1 / 3.1
    r2_title = ""
    for p in paragraphs:
        if re.match(r"\d+\.\d+", p):
            r2_title = extract_r2_from_heading(p)
            break

    row = build_kpcs_row(r0_r1, r3, mo_ta, dan_chieu, ocr_text, r2_title)

    return pd.DataFrame([row])


# ===========================================
# =========== STREAMLIT UI ==================
# ===========================================

st.title("📘 TRÍCH 4 VÙNG & MAPPING 43 CỘT KPCS – FULL FINAL VERSION")

uploaded = st.file_uploader("Tải file Word (.docx)", type=["docx"])

if uploaded:
    st.success("File đã tải. Nhấn xử lý.")

    if st.button("🔥 XỬ LÝ FILE WORD → EXCEL KPCS"):
        df = process_word_to_kpcs(uploaded)

        st.subheader("🎯 Bảng kết quả 43 cột KPCS")
        st.dataframe(df, use_container_width=True)

        # Xuất Excel
        output = io.BytesIO()
        df.to_excel(output, index=False, sheet_name="KPCS")
        st.download_button(
            label="📥 Tải xuống Excel KPCS",
            data=output.getvalue(),
            file_name="KPCS_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

