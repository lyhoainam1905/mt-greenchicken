import streamlit as st
import io, openpyxl, re, os, unicodedata, tempfile, urllib.request, ssl, zipfile
from datetime import datetime, timedelta
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# --- CẤU HÌNH GIAO DIỆN ---
st.set_page_config(page_title="Green Chicken - Kênh MT", page_icon="🍗", layout="centered")
st.markdown("""<style>.stButton>button {background-color: #2E8B57 !important; color: white !important; font-weight: bold;} h1 {color: #2E8B57 !important;}</style>""", unsafe_allow_html=True)

# --- LOGO & THÔNG TIN ---
st.title("CÔNG CỤ HỖ TRỢ KÊNH MT PRO")
st.markdown("""
### Hỗ trợ note địa chỉ vào hoá đơn (Tự động đặt tên file T+1)
* 📞 **Liên hệ hỗ trợ:** [0326.019.777](tel:0326019777)
* 🏭 **Email:** Torres.nam@deheus.vn
""")
st.divider()

# --- 1. CHUẨN HÓA UNICODE ---
def chuan_hoa_unicode(text):
    if not text: return ""
    return unicodedata.normalize('NFC', str(text)).replace('\xa0', ' ').strip()

# --- 2. NẠP FONT TIẾNG VIỆT ---
@st.cache_resource
def load_vietnamese_font():
    font_path = os.path.join(tempfile.gettempdir(), "Roboto-Bold.ttf")
    if not os.path.exists(font_path):
        try:
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
            url = "https://raw.githubusercontent.com/google/fonts/main/ofl/roboto/Roboto-Bold.ttf"
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, context=ctx, timeout=10) as response, open(font_path, 'wb') as out_file:
                out_file.write(response.read())
        except: pass
        
    if os.path.exists(font_path):
        try:
            pdfmetrics.registerFont(TTFont("FontTiengViet", font_path))
            return "FontTiengViet"
        except: pass

    local_fonts = ["Roboto-Bold.ttf", "Arial.ttf", "arialbd.ttf", "Arial Bold.ttf"]
    for lf in local_fonts:
        if os.path.exists(lf):
            try:
                pdfmetrics.registerFont(TTFont("FontTiengViet", lf))
                return "FontTiengViet"
            except: pass

    system_fonts = [
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
        "/Library/Fonts/Arial Bold.ttf",
        "C:/Windows/Fonts/arialbd.ttf"
    ]
    for path in system_fonts:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("FontTiengViet", path))
                return "FontTiengViet"
            except: continue
    return None

# --- 3. BỘ LỌC TÊN SIÊU THỊ / ĐỊA CHỈ CHUẨN (GIỮ NGUYÊN GỐC CHO CON DẤU ĐỎ) ---
def loc_ten_sieu_thi_pro(raw_note):
    text = chuan_hoa_unicode(raw_note)
    if not text: return ""
    
    # Không gộp chung ở đây nữa, giữ nguyên gốc để đóng dấu chính xác
    typo_map = {"JJIMART": "FUJIMART", "FUJI MART": "FUJIMART", "FUJI ": "FUJIMART ", "WIN MART": "WINMART", "WINMAT": "WINMART", "DELI ": "DELICA ", "THANH DO": "THÀNH ĐÔ", "BRG ": "BRG "}
    for wrong, right in typo_map.items():
        text = re.sub(re.escape(wrong), right, text, flags=re.IGNORECASE)
        
    text = re.sub(r'\d+h\d*(-\d+h\d*)?|\d+:\d+|trước\s*\d+h', ' ', text, flags=re.IGNORECASE)
    text = re.sub(r'Giao\s*xe\s*máy|Giao\s*xe|Giao\s*hàng|Giao|Xuất\s*kho|Xuất\s*cho|Xuất', ' ', text, flags=re.IGNORECASE)
    
    cac_cum = [c.strip() for c in re.split(r'[,;]', text) if c.strip()]
    if not cac_cum: return ""
    
    brands = ["FUJIMART", "BRG", "DELICA", "THÀNH ĐÔ", "INTRACOM", "WINMART", "AEON", "LOTTE"]
    for cum in cac_cum:
        cum_upper = cum.upper()
        for b in brands:
            if b in cum_upper:
                clean_loc = re.sub(r'\d+', ' ', cum_upper).replace(b, '').strip()
                clean_loc = re.sub(r'[\-\(\)\.]', ' ', clean_loc)
                clean_loc = re.sub(r'\s+', ' ', clean_loc).strip()
                words = clean_loc.split()
                seen, unique = set(), []
                for w in words:
                    if w not in seen:
                        unique.append(w)
                        seen.add(w)
                loc_final = " ".join(unique)
                return f"THÀNH ĐÔ - {loc_final}" if (b == "THÀNH ĐÔ" and loc_final) else (f"{b} {loc_final}".strip() if loc_final else b)
                    
    for cum in cac_cum:
        cum_clean = re.sub(r'\s+', ' ', cum).strip()
        if len(cum_clean) > 4 and not any(q in cum_clean.upper() for q in ["TP. HÀ NỘI", "HÀ NỘI", "VIỆT NAM", "Q. ĐỐNG ĐA"]):
            return cum_clean.upper()
            
    return cac_cum[0].upper().strip()

# --- 4. THUẬT TOÁN QUÉT EXCEL ĐA NĂNG ---
def quet_excel_da_nang(exc_file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(exc_file_bytes), data_only=True)
    so_mapping = {}
    total_rows = 0
    
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        
        so_col_idx, note_col_idx = -1, -1
        for row in sheet.iter_rows(min_row=1, max_row=15, values_only=True):
            for c_idx, val in enumerate(row):
                if not val: continue
                val_str = str(val).upper()
                if any(k in val_str for k in ["MÃ SỐ SO", "SO ĐƠN HÀNG", "MÃ SO", "ĐƠN HÀNG", "ORDER"]):
                    if so_col_idx == -1: so_col_idx = c_idx
                if any(k in val_str for k in ["ĐỊA CHỈ", "GHI CHÚ", "NOTE", "GIAO HÀNG", "STORE", "SIÊU THỊ", "TÊN KH"]):
                    if note_col_idx == -1 or any(x in val_str for x in ["ĐỊA CHỈ", "NOTE", "GHI CHÚ"]): 
                        note_col_idx = c_idx
        
        for row in sheet.iter_rows(min_row=1, values_only=True):
            so_val, note_val = None, None
            if so_col_idx != -1 and len(row) > so_col_idx: so_val = row[so_col_idx]
            if note_col_idx != -1 and len(row) > note_col_idx: note_val = row[note_col_idx]
                
            if not so_val or not any(char.isdigit() for char in str(so_val)):
                for cell in row:
                    if cell and ("GCN/SO" in str(cell).upper() or "SO/" in str(cell).upper()):
                        so_val = cell
                        break
            
            if so_val and not note_val:
                max_len = 0
                for cell in row:
                    if cell and cell != so_val and isinstance(cell, str) and len(cell) > max_len:
                        max_len = len(cell)
                        note_val = cell

            if so_val and note_val:
                digits = "".join(re.findall(r'\d+', str(so_val)))
                if len(digits) >= 6:
                    so_key = digits[-6:]
                    clean_store = loc_ten_sieu_thi_pro(note_val)
                    if clean_store and len(clean_store) > 1:
                        so_mapping[so_key] = clean_store
                        total_rows += 1

    return so_mapping, total_rows

# --- 5. GIAO DIỆN TẢI FILE ---
st.markdown("### 1️⃣ Tải dữ liệu lên (Hỗ trợ tải nhiều file)")
col1, col2 = st.columns(2)
with col1:
    excel_files = st.file_uploader("📊 Chọn file Excel (.xlsx)", type=["xlsx"], accept_multiple_files=True)
with col2:
    pdf_files = st.file_uploader("📄 Chọn file PDF, ZIP hoặc 7Z", type=["pdf", "zip", "7z"], accept_multiple_files=True)

st.markdown("---")

# --- 6. XỬ LÝ DỮ LIỆU ---
if st.button("🚀 Bấm Để Xử Lý Dữ Liệu", use_container_width=True, type="primary"):
    if not excel_files or not pdf_files:
        st.error("⚠️ Vui lòng tải lên ít nhất 1 file Excel và 1 file Hóa Đơn PDF (hoặc ZIP/7Z)!")
    else:
        ten_font = load_vietnamese_font()
        if not ten_font:
            st.error("🚨 LỖI CLOUD: Không tải được Font Tiếng Việt. Vui lòng up file Roboto-Bold.ttf lên Github.")
            st.stop()
            
        with st.spinner("⏳ Đang giải nén và xử lý dữ liệu..."):
            try:
                # Quét Excel
                so_mapping = {}
                total_rows = 0
                for exc_file in excel_files:
                    mapping_part, rows_count = quet_excel_da_nang(exc_file.read())
                    so_mapping.update(mapping_part)
                    total_rows += rows_count

                # Lọc và giải nén file PDF, ZIP, 7Z
                all_pdf_data = []
                for file_upload in pdf_files:
                    file_name_lower = file_upload.name.lower()
                    
                    if file_name_lower.endswith('.zip'):
                        with zipfile.ZipFile(io.BytesIO(file_upload.read())) as z:
                            for file_info in z.infolist():
                                if file_info.filename.lower().endswith('.pdf') and '__MACOSX' not in file_info.filename:
                                    all_pdf_data.append(z.read(file_info.filename))
                                    
                    elif file_name_lower.endswith('.7z'):
                        try:
                            import py7zr
                        except ImportError:
                            st.error("🚨 HỆ THỐNG THIẾU THƯ VIỆN ĐỌC FILE .7z")
                            st.warning("Anh Nam hãy lên GitHub tạo một file tên là `requirements.txt` và ghi chữ `py7zr` vào trong đó nhé!")
                            st.stop()
                            
                        with py7zr.SevenZipFile(io.BytesIO(file_upload.read()), mode='r') as z:
                            for filename, bio in z.readall().items():
                                if filename.lower().endswith('.pdf') and '__MACOSX' not in filename:
                                    all_pdf_data.append(bio.read())
                                    
                    elif file_name_lower.endswith('.pdf'):
                        all_pdf_data.append(file_upload.read())

                if not all_pdf_data:
                    st.error("⚠️ Không tìm thấy file PDF nào hợp lệ bên trong!")
                    st.stop()

                # Đóng dấu PDF
                final_writer = PdfWriter()
                stamped_count = 0
                
                for pdf_bytes in all_pdf_data:
                    reader = PdfReader(io.BytesIO(pdf_bytes))
                    for page in reader.pages:
                        page_text = chuan_hoa_unicode(page.extract_text() or "")
                        all_digits = re.findall(r'\d+', page_text)
                        
                        matched_store = None
                        for num_str in all_digits:
                            if len(num_str) >= 6 and num_str[-6:] in so_mapping:
                                matched_store = so_mapping[num_str[-6:]]
                                break
                        
                        if matched_store:
                            stamped_count += 1
                            mediabox = page.mediabox
                            width, height = float(mediabox.width), float(mediabox.height)
                            
                            packet = io.BytesIO()
                            can = canvas.Canvas(packet, pagesize=(width, height))
                            can.setFillColorRGB(1, 0, 0)
                            can.setFont(ten_font, 16) 
                            can.drawCentredString(width / 2.0, height - 25, matched_store)
                            can.save()
                            
                            packet.seek(0)
                            page.merge_page(PdfReader(packet).pages[0])
                            
                        final_writer.add_page(page)
                
                # -------------------------------------------------------------
                # TỰ ĐỘNG ĐẶT TÊN FILE T+1 (GOM CHUNG HỆ THỐNG FUJIMART TRÊN TÊN FILE)
                # -------------------------------------------------------------
                ngay_mai = datetime.now() + timedelta(days=1)
                str_ngay_mai = ngay_mai.strftime("%d.%m")
                
                str_all_stores = " ".join(so_mapping.values()).upper()
                danh_sach_brand = []
                
                # Gom chung nhóm Fuji, BRG, Delica vào tên "Fuji Mart" cho file xuất ra
                if any(x in str_all_stores for x in ["FUJIMART", "FUJI", "BRG", "DELICA"]):
                    danh_sach_brand.append("Fuji Mart")
                    
                if "THÀNH ĐÔ" in str_all_stores:
                    danh_sach_brand.append("Thành Đô")
                if "WINMART" in str_all_stores:
                    danh_sach_brand.append("Winmart")
                if "LOTTE" in str_all_stores:
                    danh_sach_brand.append("Lotte")
                if "AEON" in str_all_stores:
                    danh_sach_brand.append("Aeon")
                
                if danh_sach_brand:
                    # Gộp các siêu thị lại (VD: Fuji Mart - Thành Đô)
                    brand_str = " - ".join(danh_sach_brand)
                    ten_file_xuat = f"Hoá đơn {brand_str} ngày {str_ngay_mai}.pdf"
                else:
                    ten_file_xuat = f"Hoá đơn ngày {str_ngay_mai}.pdf"

                # Xuất file kết quả
                output_pdf_stream = io.BytesIO()
                final_writer.write(output_pdf_stream)
                output_pdf_stream.seek(0)
                
                st.success(f"🎉 HOÀN TẤT! Đã quét {len(so_mapping)} mã SO. Đóng dấu thành công {stamped_count} hóa đơn!")
                st.download_button(
                    label=f"📥 TẢI FILE: {ten_file_xuat.upper()}",
                    data=output_pdf_stream,
                    file_name=ten_file_xuat,
                    mime="application/pdf",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"⚠️ Có lỗi kỹ thuật xảy ra: {str(e)}")