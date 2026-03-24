import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
import io
import unicodedata

def remove_accents(input_str):
    if not isinstance(input_str, str): return str(input_str)
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).replace('đ', 'd').replace('Đ', 'D')

st.set_page_config(page_title="In Tem 2x4 Inch - Team MT", page_icon="🏷️")
st.title("🏷️ In Tem: Tem MM in 1 lan")

uploaded_file = st.file_uploader("Tai file Excel (Sheet: TEM MM)", type=['xlsx'])

if uploaded_file:
    try:
        df_all = pd.read_excel(uploaded_file, sheet_name="TEM MM", header=None).astype(str)
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
                        'SO_KIEN_SP': int(float(row[8])) if str(row[8]).replace('.','').isdigit() else 1,
                        'NGAY': str(row[9])
                    })
                except: continue

        if data_rows:
            df_final = pd.DataFrame(data_rows)
            po_totals = df_final.groupby('PO')['SO_KIEN_SP'].sum().to_dict()

            st.success(f"✅ Da san sang in {len(df_final)} dong SP.")
            
            if st.button("🚀 XUAT PDF"):
                buffer = io.BytesIO()
                c = canvas.Canvas(buffer, pagesize=(4*inch, 2*inch))
                current_po_tracker = {} 

                for _, row in df_final.iterrows():
                    po_id = row['PO']
                    tong_kien_po = po_totals[po_id]
                    if po_id not in current_po_tracker:
                        current_po_tracker[po_id] = 1
                    
                    for _ in range(row['SO_KIEN_SP']):
                        c.setLineWidth(1)
                        c.rect(0.1*inch, 0.1*inch, 3.8*inch, 1.8*inch)
                        c.line(0.1*inch, 0.55*inch, 3.9*inch, 0.55*inch)
                        
                        c.setFont("Helvetica-Bold", 9)
                        c.drawString(0.2*inch, 1.65*inch, "NCC:")
                        c.drawString(0.2*inch, 1.38*inch, "NHAN:")
                        c.drawString(0.2*inch, 1.11*inch, "SO PO:")
                        c.drawString(0.2*inch, 0.84*inch, "KIEN:")
                        c.drawString(2.1*inch, 0.84*inch, "NGAY:")
                        
                        c.setFont("Helvetica", 9)
                        c.drawString(0.75*inch, 1.65*inch, row['NCC'])
                        c.setFont("Helvetica-Bold", 11)
                        c.drawString(0.85*inch, 1.38*inch, row['NHAN'])
                        c.setFont("Helvetica", 9)
                        c.drawString(0.95*inch, 1.11*inch, f"{row['MA_NCC']}/{row['MA_ST']}. {po_id}")
                        
                        c.setFont("Helvetica-Bold", 11)
                        c.drawString(0.75*inch, 0.84*inch, f"{current_po_tracker[po_id]} / {tong_kien_po}")
                        c.setFont("Helvetica", 9)
                        c.drawString(2.7*inch, 0.84*inch, row['NGAY'])
                        
                        c.setFont("Helvetica-Bold", 10) 
                        c.drawString(0.2*inch, 0.25*inch, f"SP: {row['MA_SP']} - {row['TEN_SP']}")
                        
                        c.showPage()
                        current_po_tracker[po_id] += 1 
                        
                c.save()
                # --- ĐỔI TÊN FILE TẠI ĐÂY ---
                st.download_button("📥 TAI PDF", buffer.getvalue(), "Tem MM in 1 lan.pdf")
        else:
            st.error("❌ Khong co du lieu.")
    except Exception as e:
        st.error(f"Loi: {str(e)}")
