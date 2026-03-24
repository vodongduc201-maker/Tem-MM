import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
import io
import unicodedata

# Ham xoa dau tieng Viet
def remove_accents(input_str):
    if not isinstance(input_str, str): return str(input_str)
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).replace('đ', 'd').replace('Đ', 'D')

st.set_page_config(page_title="In Tem Chuong Duong", page_icon="🥤")
st.title("🏷️ He Thong In Tem Tu Dong - Team MT")

uploaded_file = st.file_uploader("Tai file Excel du lieu", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # 1. Chuan hoa ten cot (Xoa dau + Viet hoa)
    df.columns = [remove_accents(str(col)).strip().upper() for col in df.columns]
    
    # 2. Tim cot thong minh dựa trên ảnh mẫu
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
    c_kien = find_col(['TONG SO KIEN', 'SO KIEN'])
    c_ma_sp = find_col(['MA SAN PHAM', 'MA SP'])
    c_ten_sp = find_col(['TEN SAN PHAM', 'TEN SP'])

    if not c_po or not c_kien:
        st.error("Khong tim thay cot 'SO PO' hoac 'TONG SO KIEN' trong file!")
    else:
        try:
            # Dinh dang ngay thang (Bo gio)
            if c_ngay:
                df[c_ngay] = pd.to_datetime(df[c_ngay], errors='coerce').dt.strftime('%d/%m/%Y')

            # Chuan hoa noi dung (Xoa dau)
            df = df.astype(str).applymap(remove_accents)
            df[c_kien] = pd.to_numeric(df[c_kien], errors='coerce').fillna(1)

            # GOP DU LIEU THEO PO
            group_list = [c for c in [c_po, c_ma_ncc, c_ma_st, c_ncc, c_nhan, c_ngay] if c is not None]
            df_gop = df.groupby(group_list, as_index=False).agg({
                c_kien: 'sum',
                c_ma_sp: 'first',
                c_ten_sp: 'first'
            })

            st.success(f"Da gop thanh {len(df_gop)} PO duy nhat.")

            # Tuy chinh kich thuoc
            col1, col2 = st.columns(2)
            w_cm = col1.number_input("Chieu rong (cm)", value=10.0)
            h_cm = col2.number_input("Chieu cao (cm)", value=6.0)

            if st.button("🚀 XUAT FILE PDF"):
                buffer = io.BytesIO()
                width, height = w_cm * cm, h_cm * cm
                c = canvas.Canvas(buffer, pagesize=(width, height))

                for _, row in df_gop.iterrows():
                    tk = int(row[c_kien])
                    # Tao chuoi PO gop: MA_NCC/MA_ST. SO_PO
                    ma_ncc = str(row.get(c_ma_ncc, ''))
                    ma_st = str(row.get(c_ma_st, ''))
                    so_po = str(row.get(c_po, ''))
                    po_display = f"{ma_ncc}/{ma_st}. {so_po}"

                    for i in range(1, tk + 1):
                        # Ve khung tem
                        c.setLineWidth(1)
                        c.rect(0.2*cm, 0.2*cm, width-0.4*cm, height-0.4*cm)
                        c.line(0.2*cm, 1.4*cm, width-0.2*cm, 1.4*cm) # Line ngang cuoi
                        c.line(2.8*cm, 1.4*cm,
