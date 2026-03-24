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
        # Đọc toàn bộ file, không lấy header
        df = pd.read_excel(uploaded_file, header=None)
        
        # SỬA LỖI 4: Nếu file thiếu cột, tự động bù thêm cột trống cho đủ 10 cột (A-J)
        while len(df.columns) < 10:
            df[len(df.columns)] = ""
        
        data_rows = []
        for _, row in df.iterrows():
            # Bỏ qua dòng tiêu đề dựa trên cột A hoặc E
            val_a = str(row[0]).upper()
            if "NCC" in val_a or "PO" in str(row[4]).upper() or pd.isna(row[4]):
                continue
            
            data_rows.append({
                'NCC': remove_accents(str(row[0])),       # Cột A
                'NHAN': remove_accents(str(row[1])),      # Cột B
                'MA_NCC': str(row[2]),                   # Cột C
                'MA_ST': str(row[3]),                    # Cột D
                'PO': str(row[4]),                       # Cột E (Index 4)
                'MA_SP': str(row[5]),                    # Cột F
                'TEN_SP': remove_accents(str(row[6])),   # Cột G
                'TONG_KIEN': int(float(row[8])) if str(row[8]).replace('.','').isdigit() else 1, # Cột I
                'NGAY': str(row[9])                      # Cột J
            })

        if data_rows:
            df_final = pd.DataFrame(data_rows)
            # Gộp PO: Nếu trùng PO thì lấy số kiện lớn nhất
            df_gop = df_final.groupby(['PO', 'MA_NCC', 'MA_ST', 'NCC', 'NHAN', 'NGAY'], as_index=False).agg({
                'TONG_KIEN': 'max', 'MA_SP': 'first', 'TEN_SP': 'first'
            })

            st.success(f"✅ Da tim thay {len(df_gop)} PO.")
            st.dataframe(df_gop[['PO', 'NHAN', 'TONG_KIEN']])
            
            if st.button("🚀 XUAT FILE PDF"):
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
                st.download_button("📥 Tai file PDF", buffer.getvalue(), "Tem_MT_Fixed.pdf")
        else:
            st.warning("Khong tim thay du lieu trong file. Hay kiem tra lai cot E (Số PO).")
            
    except Exception as e:
        st.error(f"Loi: {str(e)}")
