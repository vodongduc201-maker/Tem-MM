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

st.title("🏷️ In Tem Team MT - Chuong Duong")

uploaded_file = st.file_uploader("Tai file Excel", type=['xlsx'])

if uploaded_file:
    try:
        # Đọc file và ép lấy 10 cột đầu tiên (A đến J)
        df = pd.read_excel(uploaded_file, header=None).iloc[:, :10]
        
        data_rows = []
        for _, row in df.iterrows():
            # Bỏ qua dòng tiêu đề
            if "NCC" in str(row[0]).upper() or pd.isna(row[4]):
                continue
            
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

        if data_rows:
            df_final = pd.DataFrame(data_rows)
            # Gộp các dòng trùng PO để in theo bộ kiện
            df_gop = df_final.groupby(['PO', 'MA_NCC', 'MA_ST', 'NCC', 'NHAN', 'NGAY'], as_index=False).agg({
                'TONG_KIEN': 'max', 'MA_SP': 'first', 'TEN_SP': 'first'
            })

            st.success(f"Da nhan dien {len(df_gop)} PO.")
            
            if st.button("🚀 XUAT PDF"):
                buffer = io.BytesIO()
                c = canvas.Canvas(buffer, pagesize=(10*cm, 6*cm))
                for _, row in df_gop.iterrows():
                    tk = int(row['TONG_KIEN'])
                    for i in range(1, tk + 1):
                        c.rect(0.2*cm, 0.2*cm, 9.6*cm, 5.6*cm)
                        c.setFont("Helvetica-Bold", 10)
                        c.drawString(0.5*cm, 5*cm, f"NCC: {row['NCC']}")
                        c.drawString(0.5*cm, 4*cm, f"NOI NHAN: {row['NHAN']}")
                        c.drawString(0.5*cm, 3*cm, f"PO: {row['MA_NCC']}/{row['MA_ST']}. {row['PO']}")
                        c.setFont("Helvetica-Bold", 16)
                        c.drawCentredString(5*cm, 2*cm, f"{i} / {tk}")
                        c.setFont("Helvetica", 8)
                        c.drawString(0.5*cm, 0.5*cm, f"{row['MA_SP']} - {row['TEN_SP']}")
                        c.showPage()
                c.save()
                st.download_button("📥 Tai file PDF", buffer.getvalue(), "Tem_MT.pdf")
    except Exception as e:
        st.error(f"Loi: {str(e)}")
