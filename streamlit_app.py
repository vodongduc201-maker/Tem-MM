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
st.title("🏷️ In Tem Tu Dong - Sheet 1")

uploaded_file = st.file_uploader("Tai file Excel", type=['xlsx'])

if uploaded_file:
    try:
        # Đọc Sheet 1, ép kiểu string toàn bộ để tránh lỗi định dạng
        df = pd.read_excel(uploaded_file, sheet_name=0, header=None).astype(str)
        
        # Bù cột nếu file bị thiếu
        while len(df.columns) < 10:
            df[len(df.columns)] = "nan"
        
        data_rows = []
        for _, row in df.iterrows():
            # Lấy giá trị cột E (Index 4)
            val_po = row[4].strip()
            
            # ĐIỀU KIỆN LỌC: Chỉ lấy dòng mà cột E có chứa số (Số PO thường là dãy số)
            if any(char.isdigit() for char in val_po) and val_po.upper() != "NAN":
                try:
                    data_rows.append({
                        'NCC': remove_accents(row[0]),       # Cột A
                        'NHAN': remove_accents(row[1]),      # Cột B
                        'MA_NCC': row[2],                   # Cột C
                        'MA_ST': row[3],                    # Cột D
                        'PO': val_po,                       # Cột E
                        'MA_SP': row[5],                    # Cột F
                        'TEN_SP': remove_accents(row[6]),   # Cột G
                        'TONG_KIEN': int(float(row[8])) if row[8].replace('.','').isdigit() else 1, # Cột I
                        'NGAY': row[9]                      # Cột J
                    })
                except:
                    continue

        if data_rows:
            df_final = pd.DataFrame(data_rows)
            # Gộp PO để in đúng số kiện lớn nhất
            df_gop = df_final.groupby(['PO', 'MA_NCC', 'MA_ST', 'NCC', 'NHAN', 'NGAY'], as_index=False).agg({
                'TONG_KIEN': 'max', 'MA_SP': 'first', 'TEN_SP': 'first'
            })

            st.success(f"✅ Da tim thay {len(df_gop)} PO trong Sheet 1.")
            st.dataframe(df_gop[['PO', 'NHAN', 'TONG_KIEN']])
            
            if st.button("🚀 XUAT PDF IN TEM"):
                buffer = io.BytesIO()
                c = canvas.Canvas(buffer, pagesize=(10*cm, 6*cm))
                for _, row in df_gop.iterrows():
                    tk = int(row['TONG_KIEN'])
                    for i in range(1, tk + 1):
                        c.setLineWidth(1); c.rect(0.2*cm, 0.2*cm, 9.6*cm, 5.6*cm)
                        c.line(0.2*cm, 1.4*cm, 9.8*cm, 1.4*cm)
                        c.line(2.8*cm, 1.4*cm, 2.8*cm, 5.8*cm)
                        c.setFont("Helvetica-Bold", 9); c.drawString(0.4*cm, 5*cm, f"NCC: {row['NCC']}")
                        c.setFont("Helvetica-Bold", 12); c.drawString(0.4*cm, 4*cm, f"NHAN: {row['NHAN']}")
                        c.setFont("Helvetica", 10); c.drawString(0.4*cm, 3*cm, f"PO: {row['MA_NCC']}/{row['MA_ST']}. {row['PO']}")
                        c.setFont("Helvetica-Bold", 20); c.drawCentredString(5*cm, 2.2*cm, f"{i} / {tk}")
                        c.setFont("Helvetica", 8); c.drawString(0.4*cm, 0.6*cm, f"SP: {row['MA_SP']} - {row['TEN_SP']}")
                        c.drawRightString(9.4*cm, 0.6*cm, f"Ngay: {row['NGAY']}")
                        c.showPage()
                c.save()
                st.download_button("📥 TAI FILE PDF", buffer.getvalue(), "Tem_MT_Sheet1.pdf")
        else:
            st.error("❌ Khong tim thay bat ky So PO nao tai Cot E (Cot thu 5).")
            st.write("Day la du lieu Sheet 1 may dang doc duoc:", df.head(5))
            
    except Exception as e:
        st.error(f"Loi: {str(e)}")
