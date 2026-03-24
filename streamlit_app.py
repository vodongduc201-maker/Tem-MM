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

st.set_page_config(page_title="In Tem Tu Dong", page_icon="📦")
st.title("🏷️ He Thong In Tem Tu Dong")

uploaded_file = st.file_uploader("Tai file Excel", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # 1. Chuan hoa: Xoa dau va viet hoa toan bo ten cot
    df.columns = [remove_accents(str(col)).strip().upper() for col in df.columns]
    # 2. Chuan hoa: Xoa dau toan bo noi dung trong file
    df = df.astype(str).applymap(remove_accents)

    # 3. Ham do tim cot thong minh
    def find_col(keywords):
        for col in df.columns:
            for key in keywords:
                if key in col: return col
        return None

    # Tu dong nhan dien cac cot quan trong
    c_po = find_col(['SO PO', 'PO'])
    c_ncc = find_col(['NCC', 'NHA CUNG CAP'])
    c_nhan = find_col(['NOI NHAN', 'DON VI NHAN'])
    c_ngay = find_col(['NGAY GIAO', 'NGAY'])
    c_kien = find_col(['TONG SO KIEN', 'SO KIEN'])
    c_ma = find_col(['MA SAN PHAM', 'MA HANG', 'MA SP'])
    c_ten = find_col(['TEN SAN PHAM', 'TEN HANG', 'TEN SP'])

    # Kiem tra neu thieu cot quan trong nhat la SO PO va TONG SO KIEN
    if not c_po or not c_kien:
        st.error(f"Khong tim thay cot 'SO PO' hoac 'TONG SO KIEN'. Cac cot dang co: {list(df.columns)}")
    else:
        st.info(f"Da nhan dien: PO -> {c_po} | Kien -> {c_kien}")

        try:
            # Chuyen so kien sang dang so de cong don
            df[c_kien] = pd.to_numeric(df[c_kien], errors='coerce').fillna(1)
            
            # Tao danh sach cac cot de gop (loai bo nhung cot bi thieu)
            group_list = [c for c in [c_po, c_ncc, c_nhan, c_ngay] if c is not None]
            
            # GOP DU LIEU
            df_gop = df.groupby(group_list, as_index=False).agg({
                c_kien: 'sum',
                c_ma: 'first' if c_ma else lambda x: '',
                c_ten: 'first' if c_ten else lambda x: ''
            })

            st.success(f"Da gop thanh {len(df_gop)} don hang duy nhat.")

            # --- PHAN IN PDF ---
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
                        # Ve khung tem
                        c.setLineWidth(1)
                        c.rect(0.2*cm, 0.2*cm, width-0.4*cm, height-0.4*cm)
                        c.line(0.2*cm, 1.4*cm, width-0.2*cm, 1.4*cm) 
                        c.line(2.8*cm, 1.4*cm, 2.8*cm, height-0.2*cm) 

                        # Dien tieu de (Khong dau)
                        c.setFont("Helvetica-Bold", 8)
                        y_starts = [5.2, 4.4, 3.6, 2.8, 2.0]
                        labels = ["NCC", "NOI NHAN", "SO PO:", "KIEN SO:", "NGAY GIAO:"]
                        for lab, y in zip(labels, y_starts):
                            c.drawString(0.4*cm, y*cm, lab)

                        # Dien du lieu
                        c.setFont("Helvetica", 9)
                        c.drawString(3.0*cm, 5.2*cm, str(row.get(c_ncc, '')))
                        c.setFont("Helvetica-Bold", 10)
                        c.drawString(3.0*cm, 4.4*cm, str(row.get(c_nhan, '')))
                        c.setFont("Helvetica", 10)
                        c.drawString(3.0*cm, 3.6*cm, str(row.get(c_po, '')))
                        
                        # Kien so tu dong
                        c.setFont("Helvetica-Bold", 12)
                        c.drawString(3.0*cm, 2.8*cm, f"{i}  /  {tk}")
                        
                        c.setFont("Helvetica", 10)
                        c.drawString(3.0*
