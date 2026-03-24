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

st.set_page_config(page_title="In Tem Tu Dong", page_icon="📦")
st.title("🏷️ He Thong In Tem Tu Dong")

uploaded_file = st.file_uploader("Tai file Excel", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # 1. Chuan hoa ten cot: Xoa dau va Viet hoa
    df.columns = [remove_accents(str(col)).strip().upper() for col in df.columns]
    # 2. Chuan hoa noi dung: Xoa dau
    df = df.astype(str).applymap(remove_accents)

    # 3. Tim cot thong minh
    def find_col(keywords):
        for col in df.columns:
            for key in keywords:
                if key in col: return col
        return None

    c_po = find_col(['SO PO', 'PO'])
    c_ncc = find_col(['NCC', 'NHA CUNG CAP'])
    c_nhan = find_col(['NOI NHAN', 'DON VI NHAN'])
    c_ngay = find_col(['NGAY GIAO', 'NGAY'])
    c_kien = find_col(['TONG SO KIEN', 'SO KIEN'])
    c_ma = find_col(['MA SAN PHAM', 'MA HANG', 'MA SP'])
    c_ten = find_col(['TEN SAN PHAM', 'TEN HANG', 'TEN SP'])

    if not c_po or not c_kien:
        st.error(f"Khong tim thay cot 'SO PO' hoac 'TONG SO KIEN'. Cac cot dang co: {list(df.columns)}")
    else:
        try:
            # Chuyen so kien sang dang so
            df[c_kien] = pd.to_numeric(df[c_kien], errors='coerce').fillna(1)
            
            # Tao danh sach cot de gop
            group_list = [c for c in [c_po, c_ncc, c_nhan, c_ngay] if c is not None]
            
            df_gop = df.groupby(group_list, as_index=False).agg({
                c_kien: 'sum',
                c_ma: 'first' if c_ma else lambda x: '',
                c_ten: 'first' if c_ten else lambda x: ''
            })

            st.success(f"Da gop thanh {len(df_gop)} PO duy nhat.")

            col1, col2 = st.columns(2)
            w_cm = col1.number_input("Chieu rong tem (cm)", value=10.0)
            h_cm = col2.number_input("Chieu cao tem (cm)", value=6.0)

            if st.button("🚀 XUAT FILE PDF"):
                buffer = io.BytesIO()
                width, height = w_cm * cm, h_cm * cm
                c = canvas.Canvas(buffer, pagesize=(width, height))

                for _, row in df_gop.iterrows():
                    tk = int(row[c_kien])
                    for i in range(1, tk + 1):
                        c.setLineWidth(1)
                        c.rect(0.2*cm, 0.2*cm, width-0.4*cm, height-0.4*cm)
                        c.line(0.2*cm, 1.4*cm, width-0.2*cm, 1.4*cm) 
                        c.line(2.8*cm, 1.4*cm, 2.8*cm, height-0.2*cm) 

                        c.setFont("Helvetica-Bold", 8)
                        c.drawString(0.4*cm, 5.2*cm, "NCC")
                        c.drawString(0.4*cm, 4.4*cm
