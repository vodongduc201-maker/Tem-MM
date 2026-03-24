import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
import io
import unicodedata

def remove_accents(input_str):
    if not isinstance(input_str, str): return str(input_str)
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).replace('đ', 'd').replace('Đ', 'D')

st.set_page_config(page_title="In Tem Chuong Duong - Team MT", page_icon="🥤")
st.title("🏷️ He Thong In Tem Tu Dong - Team MT")

uploaded_file = st.file_uploader("Tai file Excel du lieu", type=['xlsx'])

if uploaded_file:
    # 1. Đọc file
    df = pd.read_excel(uploaded_file)
    
    # 2. Chuẩn hóa tên cột để dò tìm
    cols = [remove_accents(str(c)).strip().upper() for c in df.columns]
    df.columns = cols

    # 3. Logic tìm cột: Nếu không thấy tên thì lấy theo chỉ số (Index) dựa trên ảnh bạn gửi
    # A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9
    def get_col_idx(name_keys, default_idx):
        for i, col in enumerate(df.columns):
            if any(k in col for k in name_keys): return i
        return default_idx

    idx_ncc = get_col_idx(['NCC'], 0)
    idx_nhan = get_col_idx(['NOI NHAN', 'NHAN'], 1)
    idx_ma_ncc = get_col_idx(['MA NCC'], 2)
    idx_ma_st = get_col_idx(['MA SIEU THI', 'MA ST'], 3)
    idx_po = get_col_idx(['SO PO', 'PO'], 4)
    idx_ma_sp = get_col_idx(['MA SAN PHAM', 'MA SP'], 5)
    idx_ten_sp = get_col_idx(['TEN SAN PHAM', 'TEN SP'], 6)
    idx_tong_kien = get_col_idx(['TONG SO KIEN', 'TONG KIEN'], 8)
    idx_ngay = get_col_idx(['NGAY'], 9)

    try:
        # Làm sạch dữ liệu
        df_clean = pd.DataFrame()
        df_clean['NCC'] = df.iloc[:, idx_ncc].astype(str).apply(remove_accents)
        df_clean['NHAN'] = df.iloc[:, idx_nhan].astype(str).apply(remove_accents)
        df_clean['MA_NCC'] = df.iloc[:, idx_ma_ncc].astype(str).apply(remove_accents)
        df_clean['MA_ST'] = df.iloc[:, idx_ma_st].astype(str).apply(remove_accents)
        df_clean['PO'] = df.iloc[:, idx_po].astype(str).apply(remove_accents)
        df_clean['MA_SP'] = df.iloc[:, idx_ma_sp].astype(str).apply(remove_accents)
        df_clean['TEN_SP'] = df.iloc[:, idx_ten_sp].astype(str).apply(remove_accents)
        df_clean['TONG_KIEN'] = pd.to_numeric(df.iloc[:, idx_tong_kien], errors='coerce').fillna(1).astype(int)
        
        # Xử lý ngày
        raw_ngay = df.iloc[:, idx_ngay]
        df_clean['NGAY'] = pd.to_datetime(raw_ngay, errors='coerce').dt.strftime('%d/%m/%Y').fillna(str(raw_ngay))

        # Gộp PO (tránh in trùng nhiều mã hàng trong 1 PO)
        df_gop = df_clean.groupby(['PO', 'MA_NCC', 'MA_ST', 'NCC', 'NHAN', 'NGAY'], as_index=False).agg({
            'TONG_KIEN': 'max',
            'MA_SP': 'first',
            'TEN_SP': 'first'
        })

        st.success(f"✅ Da nhan dien thanh cong {len(df_gop)} PO.")
        st.write("Xem truoc du lieu gop:", df_gop[['PO', 'NHAN', 'TONG_KIEN']])

        if st.button("🚀 XUAT PDF IN TEM"):
            buffer = io.BytesIO()
            c = canvas.Canvas(buffer, pagesize=(10*cm, 6*cm))

            for _, row in df_gop.iterrows():
                tk = int(row['TONG_KIEN'])
                po_display = f"{row['MA_NCC']}/{row['MA_ST']}. {row['PO']}"

                for i in range(1, tk + 1):
                    # Khung tem
                    c.setLineWidth(1)
                    c.rect(0.2*cm, 0.2*cm, 9.6*cm, 5.6*cm)
                    c.line(0.2*cm, 1.4*cm, 9.8*cm, 1.4*cm)
                    c.line(2.8*cm, 1.4*cm, 2.8*cm, 5.8*cm)
                    c.line(6.0*cm, 2.1*cm, 6.0*cm, 3.1*cm)

                    # Nhãn
                    c.setFont("Helvetica-Bold", 8)
                    c.drawString(0.4*cm, 5.2*cm, "NCC")
                    c.drawString(0.4*cm, 4.4*cm, "NOI NHAN")
                    c.drawString(0.4*cm, 3.6*cm, "SO PO :")
                    c.drawString(0.4*cm, 2.8*cm, "KIEN SO :")
                    c.drawString(0.4*cm, 2.0*cm, "NGAY GIAO :")

                    # Dữ liệu
                    c.setFont("Helvetica", 9)
                    c.drawCentredString(6.3*cm, 5.2*cm, str(row['NCC']))
                    c.setFont("Helvetica-Bold", 11)
                    c.drawCentredString(6.3*cm, 4.4*cm, str(row['NHAN']))
                    c.setFont("Helvetica", 10)
                    c.drawCentredString(6.3*cm, 3.6*cm, po_display)
                    
                    c.setFont("Helvetica-Bold", 14)
                    c.drawCentredString(4.4*cm, 2.4*cm, str(i))
                    c.drawCentredString(8.0*cm, 2.4*cm, str(tk))
                    
                    c.setFont("Helvetica", 10)
                    c.drawCentredString(6.3*cm, 1.7*cm, str(row['NGAY']))

                    # Mã & Tên SP
                    c.setFont("Helvetica-Bold", 9)
                    c.drawString(0.6*cm, 0.6*cm, str(row['MA_SP']))
                    c.drawRightString(9.4*cm, 0.6*cm, str(row['TEN_SP']))
                    c.showPage()

            c.save()
            st.download_button("📥 TAI FILE PDF", buffer.getvalue(), "Tem_TeamMT.pdf")

    except Exception as e:
        st.error(f"Loi: {e}")
