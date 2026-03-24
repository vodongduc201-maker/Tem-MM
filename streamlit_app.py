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

st.set_page_config(page_title="In Tem Team MT", page_icon="🥤")
st.title("🏷️ He Thong In Tem Team MT - Final Fix")

uploaded_file = st.file_uploader("Tai file Excel du lieu", type=['xlsx'])

if uploaded_file:
    # Đọc file KHÔNG lấy tiêu đề để lấy toàn bộ dữ liệu từ dòng đầu tiên
    df = pd.read_excel(uploaded_file, header=None)
    
    st.info(f"Da doc file voi {len(df.columns)} cot.")

    try:
        data_rows = []
        for i, row in df.iterrows():
            # Bỏ qua các dòng tiêu đề nếu người dùng để lại
            val_a = str(row[0]).upper()
            if "NCC" in val_a or "NHA CUNG CAP" in val_a:
                continue
            
            # Lấy dữ liệu theo vị trí cột (Index 0-9 tương ứng A-J)
            data_rows.append({
                'NCC': remove_accents(str(row[0])),       # Cột A
                'NHAN': remove_accents(str(row[1])),      # Cột B
                'MA_NCC': str(row[2]),                   # Cột C
                'MA_ST': str(row[3]),                    # Cột D
                'PO': str(row[4]),                       # Cột E
                'MA_SP': str(row[5]),                    # Cột F
                'TEN_SP': remove_accents(str(row[6])),   # Cột G
                'TONG_KIEN': int(float(row[8])) if pd.notnull(row[8]) else 1, # Cột I
                'NGAY': str(row[9])                      # Cột J
            })
        
        df_final = pd.DataFrame(data_rows)

        # Gộp PO: Nếu 1 PO có nhiều mã hàng, chỉ in 1 bộ tem dựa trên TỔNG SỐ KIỆN lớn nhất
        df_gop = df_final.groupby(['PO', 'MA_NCC', 'MA_ST', 'NCC', 'NHAN', 'NGAY'], as_index=False).agg({
            'TONG_KIEN': 'max',
            'MA_SP': 'first',
            'TEN_SP': 'first'
        })

        st.success(f"✅ San sang in {len(df_gop)} don hang PO.")
        st.dataframe(df_gop[['PO', 'NHAN', 'TONG_KIEN']])

        if st.button("🚀 XUAT PDF IN TEM"):
            buffer = io.BytesIO()
            c = canvas.Canvas(buffer, pagesize=(10*cm, 6*cm))
            
            for _, row in df_gop.iterrows():
                tong_kien = int(row['TONG_KIEN'])
                po_full = f"{row['MA_NCC']}/{row['MA_ST']}. {row['PO']}"
                
                for i in range(1, tong_kien + 1):
                    # Vẽ khung tem 10x6
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

                    c.setFont("Helvetica", 9); c.drawCentredString(6.3*cm, 5.2*cm, row['NCC'])
                    c.setFont("Helvetica-Bold", 11); c.drawCentredString(6.3*cm, 4.4*cm, row['NHAN'])
                    c.setFont("Helvetica", 10); c.drawCentredString(6.3*cm, 3.6*cm, po_full)
                    
                    c.setFont("Helvetica-Bold", 14)
                    c.drawCentredString(4.4*cm, 2.4*cm, str(i))
                    c.drawCentredString(8.0*cm, 2.4*cm, str(tong_kien))
                    
                    c.setFont("Helvetica", 10); c.drawCentredString(6.3*cm, 1.7*cm, row['NGAY'])
                    c.setFont("Helvetica-Bold", 9); c.drawString(0.6*cm, 0.6*cm, row['MA_SP'])
                    c.drawRightString(9.4*cm, 0.6*cm, row['TEN_SP'])
                    c.showPage()
            
            c.save()
            st.download_button("📥 TAI FILE PDF", buffer.getvalue(), "Tem_TeamMT_Chuan.pdf")

    except Exception as e:
        st.error(f"Loi doc file: {e}. Hay dam bao file co cac cot tu A den J.")
