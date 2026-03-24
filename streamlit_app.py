import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
import io
import unicodedata

# Hàm xóa dấu tiếng Việt
def remove_accents(input_str):
    if not isinstance(input_str, str): return str(input_str)
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).replace('đ', 'd').replace('Đ', 'D')

st.set_page_config(page_title="In Tem Team MT - Chuong Duong", page_icon="🥤")
st.title("🏷️ He Thong In Tem Tu Dong - Team MT")

uploaded_file = st.file_uploader("Tai file Excel du lieu", type=['xlsx'])

if uploaded_file:
    # Đọc file Excel
    df = pd.read_excel(uploaded_file)
    
    # 1. Lưu lại tên cột gốc để báo lỗi nếu cần
    original_cols = df.columns.tolist()
    
    # 2. Chuẩn hóa tên cột để máy dễ đọc
    df.columns = [remove_accents(str(col)).strip().upper() for col in df.columns]

    # 3. Hàm dò tìm cột linh hoạt (Chỉ cần CHỨA từ khóa là lấy)
    def find_col_flexible(keywords):
        for col in df.columns:
            if any(key in col for key in keywords):
                return col
        return None

    # Tìm kiếm các cột quan trọng
    c_po = find_col_flexible(['SO PO', 'PO'])
    c_kien_tong = find_col_flexible(['TONG SO KIEN', 'TONG KIEN', 'KIEN'])
    c_ncc = find_col_flexible(['NCC'])
    c_nhan = find_col_flexible(['NOI NHAN', 'NHAN'])
    c_ma_ncc = find_col_flexible(['MA NCC'])
    c_ma_st = find_col_flexible(['MA SIEU THI', 'MA ST'])
    c_ma_sp = find_col_flexible(['MA SAN PHAM', 'MA SP', 'MA HANG'])
    c_ten_sp = find_col_flexible(['TEN SAN PHAM', 'TEN SP', 'TEN HANG'])
    c_ngay = find_col_flexible(['NGAY'])

    # KIỂM TRA LỖI VÀ HIỂN THỊ HƯỚNG DẪN
    if not c_po or not c_kien_tong:
        st.error("❌ KHÔNG TÌM THẤY CỘT 'SỐ PO' HOẶC 'TỔNG SỐ KIỆN'!")
        st.warning(f"Cột máy nhận được: {original_cols}")
        st.info("💡 Mẹo: Bạn hãy kiểm tra lại dòng đầu tiên trong file Excel xem đã có tiêu đề chưa nhé.")
    else:
        try:
            # Thông báo cho người dùng biết máy đã nhận diện đúng
            st.success(f"✅ Đã tìm thấy: Cột PO ({c_po}) và Cột Số Kiện ({c_kien_tong})")

            # Xử lý ngày tháng
            df[c_ngay] = pd.to_datetime(df[c_ngay], errors='coerce').dt.strftime('%d/%m/%Y')
            
            # Xóa dấu nội dung
            df = df.astype(str).applymap(remove_accents)
            
            # Chuyển số kiện sang dạng số
            df[c_kien_tong] = pd.to_numeric(df[c_kien_tong], errors='coerce').fillna(1).astype(int)

            # LOGIC GỘP PO
            group_list = [c for c in [c_po, c_ma_ncc, c_ma_st, c_ncc, c_nhan, c_ngay] if c is not None]
            df_gop = df.groupby(group_list, as_index=False).agg({
                c_kien_tong: 'max', 
                c_ma_sp: 'first',
                c_ten_sp: 'first'
            })

            if st.button("🚀 XUẤT FILE PDF IN TEM"):
                buffer = io.BytesIO()
                c = canvas.Canvas(buffer, pagesize=(10*cm, 6*cm))

                for _, row in df_gop.iterrows():
                    tk = int(row[c_kien_tong])
                    po_display = f"{row.get(c_ma_ncc, '')}/{row.get(c_ma_st, '')}. {row.get(c_po, '')}"

                    for i in range(1, tk + 1):
                        c.setLineWidth(1)
                        c.rect(0.2*cm, 0.2*cm, 9.6*cm, 5.6*cm)
                        c.line(0.2*cm, 1.4*cm, 9.8*cm, 1.4*cm)
                        c.line(2.8*cm, 1.4*cm, 2.8*cm, 5.8*cm)
                        c.line(6.0*cm, 2.1*cm, 6.0*cm, 3.1*cm)

                        c.setFont("Helvetica-Bold", 8)
                        c.drawString(0.4*cm, 5.2*cm, "NCC")
                        c.drawString(0.4*cm, 4.4*cm, "NOI NHAN")
                        c.drawString(0.4*cm, 3.6*cm, "SO PO :")
                        c.drawString(0.4*cm, 2.8*cm, "KIEN SO :")
                        c.drawString(0.4*cm, 2.0*cm, "NGAY GIAO :")

                        c.setFont("Helvetica", 9)
                        c.drawCentredString(6.3*cm, 5.2*cm, str(row.get(c_ncc, '')))
                        c.setFont("Helvetica-Bold", 11)
                        c.drawCentredString(6.3*cm, 4.4*cm, str(row.get(c_nhan, '')))
                        c.setFont("Helvetica", 10)
                        c.drawCentredString(6.3*cm, 3.6*cm, po_display)
                        
                        c.setFont("Helvetica-Bold", 14)
                        c.drawCentredString(4.4*cm, 2.4*cm, str(i))
                        c.drawCentredString(8.0*cm, 2.4*cm, str(tk))
                        
                        c.setFont("Helvetica", 10)
                        c.drawCentredString(6.3*cm, 1.7*cm, str(row.get(c_ngay, '')))

                        c.setFont("Helvetica-Bold", 9)
                        c.drawString(0.6*cm, 0.6*cm, str(row.get(c_ma_sp, '')))
                        c.drawRightString(9.4*cm, 0.6*cm, str(row.get(c_ten_sp, '')))
                        c.showPage()

                c.save()
                st.download_button("📥 TẢI FILE PDF", buffer.getvalue(), "Tem_TeamMT_Update.pdf")

        except Exception as e:
            st.error(f"Lỗi xử lý: {e}")
