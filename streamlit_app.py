import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
import io
import unicodedata
import re

# Hàm chuyển đổi tiếng Việt có dấu sang không dấu
def remove_accents(input_str):
    if not isinstance(input_str, str):
        return str(input_str)
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).replace('đ', 'd').replace('Đ', 'D')

st.set_page_config(page_title="In Tem Khong Dau", page_icon="📦")
st.title("🏷️ He Thong In Tem Tu Dong (Khong Dau)")

uploaded_file = st.file_uploader("Tai file Excel du lieu", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # 1. Chuyen tieu de cot sang KHONG DAU va VIET HOA
    df.columns = [remove_accents(str(col)).strip().upper() for col in df.columns]
    
    # 2. Chuyen toan bo noi dung trong bang sang KHONG DAU
    df = df.astype(str).applymap(remove_accents)
    
    # 3. Gom nhom theo SO PO va cong don SO KIEN
    # Luu y: Ten cot gio day la 'SO PO' va 'TONG SO KIEN' (da mat dau)
    col_po = 'SO PO :' if 'SO PO :' in df.columns else 'SO PO'
    col_kien = 'TONG SO KIEN' if 'TONG SO KIEN' in df.columns else 'TONG SO KIEN :'
    
    # Gom nhom de cong don so kien
    df_gop = df.groupby([col_po, 'NCC', 'NOI NHAN', 'NGAY GIAO :'], as_index=False).agg({
        col_kien: lambda x: pd.to_numeric(x).sum(),
        'MA SAN PHAM': 'first',
        'TEN SAN PHAM': 'first'
    })

    st.success(f"Da xu ly {len(df_gop)} don hang (PO).")
    st.write("Du lieu xem truoc:", df_gop.head())

    col1, col2 = st.columns(2)
    w_cm = col1.number_input("Rong (cm)", value=10.0)
    h_cm = col2.number_input("Cao (cm)", value=6.0)

    if st.button("🚀 XUAT FILE PDF"):
        buffer = io.BytesIO()
        width, height = w_cm * cm, h_cm * cm
        c = canvas.Canvas(buffer, pagesize=(width, height))

        for index, row in df_gop.iterrows():
            try:
                tong_kien = int(row[col_kien])
            except:
                tong_kien = 1
            
            for i in range(1, tong_kien + 1):
                # Ve khung
                c.setLineWidth(1)
                c.rect(0.2*cm, 0.2*cm, width-0.4*cm, height-0.4*cm)
                c.line(0.2*cm, 1.4*cm, width-0.2*cm, 1.4*cm) 
                c.line(2.8*cm, 1.4*cm, 2.8*cm, height-0.2*cm) 

                # Dung font mac dinh Helvetica (vi da khong con dau nen khong lo loi)
                c.setFont("Helvetica-Bold", 8)
                c.drawString(0.4*cm, 5.2*cm, "NCC")
                c.drawString(0.4*cm, 4.4*cm, "NOI NHAN")
                c.drawString(0.4*cm, 3.6*cm, "SO PO:")
                c.drawString(0.4*cm, 2.8*cm, "KIEN SO:")
                c.drawString(0.4*cm, 2.0*cm, "NGAY GIAO:")

                c.setFont("Helvetica", 9)
                c.drawString(3.0*cm, 5.2*cm, str(row.get('NCC', '')))
                c.setFont("Helvetica-Bold", 11)
                c.drawString(3.0*cm, 4.4*cm, str(row.get('NOI NHAN', '')))
                c.setFont("Helvetica", 10)
                c.drawString(3.0*cm, 3.6*cm, str(row.get(col_po, '')))
                
                c.setFont("Helvetica-Bold", 12)
                c.drawString(3.0*cm, 2.8*cm, f"{i}  /  {tong_kien}")
                
                c.setFont("Helvetica", 10)
                c.drawString(3.0*cm, 2.0*cm, str(row.get('NGAY GIAO :', '')))

                c.setFont("Helvetica-Bold", 10)
                c.drawString(0.4*cm, 0.6*cm, str(row.get('MA SAN PHAM', '')))
                c.drawRightString(width-0.4*cm, 0.6*cm, str(row.get('TEN SAN PHAM', '')))

                c.showPage()

        c.save()
        st.download_button("📥 TAI PDF", buffer.getvalue(), "Tem_Khong_Dau.pdf", "application/pdf")
