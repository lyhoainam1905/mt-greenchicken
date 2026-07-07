import streamlit as st
import io
import openpyxl
import re
import unicodedata
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

# Cấu hình giao diện Web Streamlit
st.set_page_config(page_title="Hệ Thống Xử Lý Hóa Đơn Green Chicken", page_icon="🍗", layout="centered")

# Nạp Font Tiếng Việt chuẩn để không bao giờ lỗi font
@st.cache_resource
def load_vietnamese_font():
    danh_sach_font = [
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Library/Fonts/Arial Bold.ttf",
        "/System/Library/Fonts/Supplemental/Times New Roman Bold.ttf",
        "/System/Library/Fonts/Supplemental/Tahoma.ttf",
    ]
    for path in danh_sach_font:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("FontTiengViet", path))
                return "FontTiengViet", False
            except Exception:
                continue
    return "Helvetica-Bold", True

def xoa_dau_tieng_viet(text):
    if not text:
        return ""
    text = unicodedata.normalize('NFD', str(text))
    text = ''.join(c for c in text if unicodedata.category(c) != 'Mn')
    return text.replace('Đ', 'D').replace('đ', 'd').strip()

def lay_5_so_cuoi(text_value):
    if not text_value:
        return None
    digits = "".join(re.findall(r'\d+', str(text_value)))
    return digits[-5:] if len(digits) >= 5 else None

# ==============================================================================
# BỘ LỌC TÊN SIÊU THỊ THÔNG MINH (LOẠI BỎ RÁC LOGISTICS)
# ==============================================================================
def loc_ten_sieu_thi(raw_note):
    if not raw_note:
        return ""
    
    # 1. Danh sách từ khóa SIÊU THỊ ưu tiên giữ lại (Anh muốn thêm siêu thị nào cứ gõ vào đây)
    tu_khoa_sieu_thi = ["FUJIMART", "BRG", "DELICA", "INTRACOM", "WINMART", "BIG C", "GO!", "AEON", "LOTTE"]
    
    # 2. Danh sách từ khóa RÁC LOGISTICS cần xóa bỏ
    tu_khoa_rac = [
        "giao xe máy", "giao xe may", "xe máy", "xe may",
        "giao ô tô", "giao oto", "ô tô", "oto",
        "giao hàng", "giao hang", "giao xe", "giao trước", "giao lúc", "giao"
    ]
    
    note_str = str(raw_note).strip()
    
    # Tách câu theo dấu phẩy, dấu gạch ngang, dấu chấm phẩy hoặc dấu ngoặc đơn
    # Ví dụ: "Giao xe máy, Fujimart Hoàng Cầu - trước 8h" -> ["Giao xe máy", "Fujimart Hoàng Cầu", "trước 8h"]
    cac_cum_tu = re.split(r'[,;\-\(\)]', note_str)
    
    # Tìm cụm từ nào chứa từ khóa siêu thị (FUJIMART, BRG, DELICA...)
    for cum in cac_cum_tu:
        cum_clean = cum.strip()
        cum_upper = cum_clean.upper()
        for tk in tu_khoa_sieu_thi:
            if tk in cum_upper:
                # Tìm thấy siêu thị! Tiến hành dọn dẹp từ rác nếu có dính kèm trong cụm này
                ket_qua = cum_clean
                for rac in tu_khoa_rac:
                    ket_qua = re.sub(re.escape(rac), "", ket_qua, flags=re.IGNORECASE).strip()
                # Xóa dấu ký tự thừa ở đầu/cuối (ví dụ: ": Fujimart" -> "Fujimart")
                return re.sub(r'^[:\-\s]+|[:\-\s]+$', '', ket_qua).strip()
                
    # Nếu trong câu không có tên siêu thị nào quen thuộc, tiến hành lọc rác trên toàn câu và lấy phần có ý nghĩa nhất
    ket_qua = note_str
    for rac in tu_khoa_rac:
        ket_qua = re.sub(re.escape(rac), "", ket_qua, flags=re.IGNORECASE).strip()
        
    cac_cum_sach = [c.strip() for c in re.split(r'[,;\-\(\)]', ket_qua) if c.strip()]
    return cac_cum_sach[0] if cac_cum_sach else ""
# ==============================================================================

# --- GIAO DIỆN TRANG WEB ---
st.title("🍗 CỔNG XỬ LÝ HÓA ĐƠN TỰ ĐỘNG")
st.markdown("---")
st.markdown("### 1️⃣ Tải dữ liệu lên")

col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("📊 Chọn file Excel từ Admin (.xlsx)", type=["xlsx"])
with col2:
    pdf_files = st.file_uploader("📄 Chọn các file Hóa Đơn Lẻ (.pdf)", type=["pdf"], accept_multiple_files=True)

st.markdown("---")

# Nút bấm xử lý
if st.button("🚀 TIẾN HÀNH ĐÓNG DẤU & GỘP HÓA ĐƠN", use_container_width=True, type="primary"):
    if not excel_file or not pdf_files:
        st.error("⚠️ Vui lòng tải lên đầy đủ file Excel và ít nhất 1 file Hóa Đơn PDF!")
    else:
        with st.spinner("⏳ Hệ thống đang lọc siêu thị (Fujimart, BRG, Delica...) & đóng dấu..."):
            try:
                # 1. Đọc Excel
                wb = openpyxl.load_workbook(io.BytesIO(excel_file.read()))
                sheet = wb.active
                
                header_row = [str(cell).strip() if cell is not None else "" for cell in next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))]
                note_col_idx = -1
                for idx, header in enumerate(header_row):
                    h_upper = header.upper()
                    if "NOTE" in h_upper or "GIAO" in h_upper or "ĐỊA CHỈ" in h_upper or "STORE" in h_upper:
                        note_col_idx = idx
                        break
                if note_col_idx == -1:
                    note_col_idx = 2 if len(header_row) >= 3 else 1
                
                ten_font, can_xoa_dau = load_vietnamese_font()
                so_mapping = {}
                for row in sheet.iter_rows(min_row=2, values_only=True):
                    if row[0] is not None and len(row) > note_col_idx and row[note_col_idx] is not None:
                        so_number = str(row[0]).strip()
                        raw_note = str(row[note_col_idx]).strip()
                        
                        # ÁP DỤNG BỘ LỌC TÊN SIÊU THỊ THÔNG MINH
                        clean_store_name = loc_ten_sieu_thi(raw_note)
                        
                        if can_xoa_dau:
                            clean_store_name = xoa_dau_tieng_viet(clean_store_name)
                        excel_5 = lay_5_so_cuoi(so_number)
                        if excel_5 and clean_store_name:
                            so_mapping[excel_5] = clean_store_name
                
                # 2. Quét và Đóng dấu PDF
                final_writer = PdfWriter()
                for pdf_file in pdf_files:
                    reader = PdfReader(io.BytesIO(pdf_file.read()))
                    for page in reader.pages:
                        page_text = page.extract_text() or ""
                        all_digits = re.findall(r'\d+', page_text)
                        matched_store = None
                        for num_str in all_digits:
                            if len(num_str) >= 5:
                                invoice_5 = num_str[-5:]
                                if invoice_5 in so_mapping:
                                    matched_store = so_mapping[invoice_5]
                                    break
                        
                        if matched_store:
                            mediabox = page.mediabox
                            width, height = float(mediabox.width), float(mediabox.height)
                            packet = io.BytesIO()
                            can = canvas.Canvas(packet, pagesize=(width, height))
                            can.setFillColorRGB(1, 0, 0) # Màu đỏ rực
                            can.setFont(ten_font, 18)
                            
                            # Căn chính giữa đỉnh đầu
                            can.drawCentredString(width / 2.0, height - 28, matched_store.upper())
                            can.save()
                            
                            packet.seek(0)
                            new_pdf = PdfReader(packet)
                            page.merge_page(new_pdf.pages[0])
                            
                        final_writer.add_page(page)
                
                # Xuất file ra bộ nhớ tạm của Web
                output_pdf_stream = io.BytesIO()
                final_writer.write(output_pdf_stream)
                output_pdf_stream.seek(0)
                
                # 3. Hiển thị kết quả thành công
                st.success(f"🎉 ĐÃ XỬ LÝ XONG {len(pdf_files)} FILE HÓA ĐƠN! CHỈ LẤY ĐÚNG TÊN SIÊU THỊ!")
                st.markdown("---")
                
                # Nút tải file về cho điện thoại / máy tính
                st.download_button(
                    label="📥 BẤM VÀO ĐÂY ĐỂ TẢI HÓA ĐƠN HOÀN CHỈNH VỀ MÁY",
                    data=output_pdf_stream,
                    file_name="TAT_CA_HOA_DON_DE_IN.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"⚠️ Có lỗi xảy ra trong quá trình xử lý: {str(e)}")