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

st.set_page_config(page_title="In Tem Team MT - Chuong Duong", page_icon="🏷️")
st.title("🏷️ In Tem Tu Dong - Phien Ban Moi nhat")

uploaded_file = st.file_uploader("Tai file Excel cua Team MT", type=['xlsx'])

if uploaded_file:
    try:
        df_all = pd.read_excel(uploaded_file, sheet_name="TEM MM", header=None)
        while len(df_all.columns) < 10:
            df_all[len(df_all.columns)] = ""

        data_rows = []
        for _, row in df_all.iterrows():
            val_po = str(row[4]).strip().upper()
            if any(char.isdigit() for char in val_po) and val_po not in ["NAN", "SO PO", "PO"]:
                try:
                    data_rows.append({
                        'NCC': remove_accents(str(row[0])),
                        'NHAN': remove_accents(str(row[1])),
                        'MA_NCC': str(row[2]),
                        'MA_ST': str(row[3]),
                        'PO': str(row[4]).strip(),
                        'MA_SP': str(row[5]),
                        'TEN_SP': remove_accents(str(row[6])),
                        'TONG_KIEN': int(float(row[8])) if str(row[8]).replace('.','').isdigit() else 1,
                        'NGAY': str(row[9])
                    })
                except:
                    continue

        if data_rows:
            df_final = pd.DataFrame(data_rows)
            df_gop = df_final.groupby(['PO', 'MA_NCC', 'MA_ST', 'NCC', 'NHAN', 'NGAY'], as_index=False).agg({
                'TONG_KIEN': 'max', 'MA_SP': 'first', 'TEN_SP': 'first'
            })

            st.success(f"✅ Da san sang in {len(df_gop)} PO.")
            
            if st.button("🚀 XUAT PDF IN TEM"):
                buffer = io.BytesIO()
                c = canvas.Canvas(buffer, pagesize=(10*cm, 6*cm))
                for _, row in df_gop.iterrows():
                    tk = int(row['TONG_KIEN'])
                    po_hien_thi = f"{row['MA_NCC']}/{row['MA_ST']}. {row['PO']}"
                    
                    for i in range(1, tk + 1):
                        # Khung ngoai
                        c.setLineWidth(1); c.rect(0.2*cm, 0.2*cm, 9.6*cm, 5.6*cm)
                        # Duong ke ngang phan SP
                        c.line(0.2*cm, 1.2*cm, 9.8*cm, 1.2*cm)
                        
                        # Thiet ke noi dung phan chinh
                        c.setFont("Helvetica-Bold", 10)
                        c.drawString(0.4*cm, 5.1*cm, "NCC:")
                        c.drawString(0.4*cm, 4.3*cm, "NHAN:")
                        c.drawString(0.4*cm, 3.5*cm, "SO PO:")
                        c.drawString(0.4*cm, 2.7*cm, "KIEN SO:")
                        # --- DI CHUYEN NGAY GIAO LEN DAY ---
                        c.drawString(0.4*cm, 1.9*cm, "NGAY GIAO:") 
                        
                        # Du lieu chi tiet
                        c.setFont("Helvetica", 10)
                        c.drawString(1.7*cm, 5.1*cm, row['NCC'])
                        
                        c.setFont("Helvetica-Bold", 13) 
                        c.drawString(2.1*cm, 4.3*cm, row['NHAN'])
                        
                        c.setFont("Helvetica", 10)
                        c.drawString(2.2*cm, 3.5*cm, po_hien_thi)
                        
                        # So kien (font vua phai, can deu)
                        c.setFont("Helvetica-Bold", 12)
                        c.drawString(2.5*cm, 2.7*cm, f"{i} / {tk}")
                        
                        # Ngay giao (hien thi ngay duoi kien so)
                        c.setFont("Helvetica", 11)
                        c.drawString(3.0*cm, 1.9*cm, row['NGAY'])
                        
                        # Dong duoi cung (Chi de thong tin San pham cho rong rai)
                        c.setFont("Helvetica-Bold", 9)
                        c.drawString(0.4*cm, 0.5*cm, f"SP: {row['MA_SP']} - {row['TEN_SP']}")
                        
                        c.showPage()
                c.save()
                st.download_button("📥 TAI FILE PDF", buffer.getvalue(), "Tem_MT_CanDoi.pdf")
        else:
            st.error("❌ Khong tim thay du lieu PO hop le.")
            
    except Exception as e:
        st.error(f"Loi: {str(e)}")
