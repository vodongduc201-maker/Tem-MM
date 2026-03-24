import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
import io
import unicodedata

# Hàm xóa dấu tiếng Việt để in ấn không bị lỗi font
def remove_accents(input_str):
    if not isinstance(input_str, str): return str(input_str)
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).replace('đ', 'd').replace('Đ', 'D')

st.set_page_config(page_title="In Tem Chuong Duong MT", page_icon="🥤")
st.title("🏷️ He Thong In Tem Tu Dong - Team MT")

uploaded_file = st.file_uploader("Tai file Excel du lieu (nhu anh mau)", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # 1. Chuẩn hóa tên cột
    df.columns = [remove_accents(str(col)).strip().upper() for col in df.columns]
    
    # 2. Nhận diện cột thông minh theo ảnh mẫu mới
    def find_col(keywords):
        for col in df.columns:
            for key in keywords:
                if key in col: return col
        return None

    c_po = find_col(['SO PO', 'PO'])
    c_ma_ncc = find_col(['MA NCC'])
    c_ma_st = find_col(['MA SIEU THI', 'MA ST'])
    c_ncc = find_col(['NCC', 'NHA CUNG CAP'])
    c_nhan = find_col(['NOI NHAN'])
    c_ngay = find_col(['NGAY GIAO', 'NGAY'])
    c_kien_tong = find_col(['TONG SO KIEN'])
    c_ma_sp = find_col(['MA SAN PHAM', 'MA SP'])
    c_ten_sp = find_col(['TEN SAN PHAM', 'TEN SP'])

    if not c_po or not c_kien_tong:
        st.error("Khong tim thay cot 'SO PO' hoac 'TONG SO KIEN'!")
    else:
        try:
            # Định dạng ngày tháng (chỉ lấy ngày/tháng/năm)
            if c_ngay:
                df[c_ngay] = pd.to_datetime(df[c_ngay], errors='coerce').dt.strftime('%d/%m/%Y')

            # Xóa dấu toàn bộ nội dung
            df = df.astype(str).applymap(remove_accents)
            df[c_kien_tong] = pd.to_numeric(df[c_kien_tong], errors='coerce').fillna(1).astype(int)

            # GỘP PO: Gom các dòng trùng PO, Mã NCC, Mã ST
            group_list = [c for c in [c_po, c_ma_ncc, c_ma_st, c_ncc, c_nhan, c_ngay] if c in df.columns]
            df_gop = df.groupby(group_list, as_index=False).agg({
                c_kien_tong: 'sum',
                c_ma_sp: 'first',
                c_ten_sp: 'first'
            })

            st.success(f"Da gop thanh {len(df_gop)} PO duy nhat.")

            col1, col2 = st.columns(2)
            w_cm = col1.number_input("Rong tem (cm)", value=10.0)
            h_cm = col2.number_input("Cao tem (cm)", value=6.0)

            if st.button("🚀 XUAT FILE PDF IN TEM"):
                buffer = io.BytesIO()
                width, height = w_cm * cm, h_cm * cm
                c = canvas.Canvas(buffer, pagesize=(width, height))

                for _, row in df_gop.iterrows():
                    tong_kien = int(row[c_kien_tong])
                    # Tạo chuỗi PO: MA_NCC/MA_ST. SO_PO
                    po_display = f"{row.get(c_ma_ncc, '')}/{row.get(c_ma_st, '')}. {row.get(c_po, '')}"

                    for i in range(1, tong_kien + 1):
                        # --- VẼ KHUNG ---
                        c.setLineWidth(1)
                        c.rect(0.2*cm, 0.2*cm, width-0.4*cm, height-0.4*cm)
                        # Đường ngang trên Mã SP
                        c.line(0.2*cm, 1.4*cm, width-0.2*cm, 1.4*cm)
                        # Đường dọc chia nhãn và dữ liệu
                        c.line(2.8*cm, 1.4*cm, 2.8*cm, height-0.2*cm)
                        # Đường dọc trong ô KIỆN SỐ (chia i và tổng kiện)
                        c.line(6.0*cm, 2.2*cm, 6.0*cm, 3.0*cm)

                        # --- ĐIỀN NHÃN (Lables) ---
                        c.setFont("Helvetica-Bold", 8)
                        c.drawString(0.4*cm, 5.2*cm, "NCC")
                        c.drawString(0.4*cm, 4.4*cm, "NOI NHAN")
                        c.drawString(0.4*cm, 3.6*cm, "SO PO :")
                        c.drawString(0.4*cm, 2.8*cm, "KIEN SO :")
                        c.drawString(0.4*cm, 2.0*cm, "NGAY GIAO :")

                        # --- ĐIỀN DỮ LIỆU (Data) ---
                        c.setFont("Helvetica", 10)
                        # NCC
                        c.drawCentredString((width+2.8*cm)/2, 5.2*cm, str(row.get(c_ncc, '')))
                        # NƠI NHẬN
                        c.setFont("Helvetica-Bold", 12)
                        c.drawCentredString((width+2.8*cm)/2, 4.4*cm, str(row.get(c_nhan, '')))
                        # SỐ PO (đã ghép)
                        c.setFont("Helvetica-Bold", 11)
                        c.drawCentredString((width+2.8*cm)/2, 3.6*cm, po_display)
                        
                        # KIỆN SỐ (Nhảy số tự động)
                        c.setFont("Helvetica-Bold", 14)
                        c.drawCentredString(4.4*cm, 2.4*cm, str(i))
                        c.drawCentredString(8.0*cm, 2.4*cm, str(tong_kien))
                        
                        # NGÀY GIAO
                        c.setFont("Helvetica", 10)
                        c.drawCentredString((width+2.8*cm)/2, 1.8*cm, str(row.get(c_ngay, '')))

                        # DÒNG CUỐI: MÃ SP & TÊN SP
                        c.setFont("Helvetica-Bold", 10)
                        c.drawString(0.8*cm, 0.6*cm, str(row.get(c_ma_sp, '')))
                        c.drawRightString(width-0.8*cm, 0.6*cm, str(row.get(c_ten_sp, '')))

                        c.showPage()

                c.save()
                st.download_button("📥 TAI FILE PDF VE", buffer.getvalue(), "Tem_Giao_Hang_TeamMT.pdf")

        except Exception as e:
            st.error(f"Loi: {e}")
