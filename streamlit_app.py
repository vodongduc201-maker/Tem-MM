import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch # Chuyển sang dùng đơn vị inch
import io
import unicodedata

def remove_accents(input_str):
    if not isinstance(input_str, str): return str(input_str)
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).replace('đ', 'd').replace('Đ', 'D')

st.set_page_config(page_title="In Tem 2x4 Inch - Team MT", page_icon="🏷️")
st.title("🏷️ In Tem Nhiet: Kho 2x4 Inch")

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

            st.success(f"✅ San sang in {len(df_final)} dong SP tren kho 2x4 inch.")
            
            if st.button("🚀 XUAT FILE PDF 2x4 INCH"):
                buffer = io.BytesIO()
                # THIẾT LẬP KHỔ TRANG: Ngang 4 inch, Cao 2 inch
                c = canvas.Canvas(buffer, pagesize=(4*inch, 2*inch))
                
                current_po_tracker = {} 

                for _, row in df_final.iterrows():
                    po_id = row['PO']
                    tong_kien_po = po_totals[po_id]
                    if po_id not in current_po_tracker:
                        current_po_tracker[po_id] = 1
                    
                    for _ in range(row['SO_KIEN_SP']):
                        # Khung bao ngoài sát lề (cách lề 0.05 inch)
                        c.setLineWidth(1)
                        c.rect(0.05*inch, 0.05*inch, 3.9*inch, 1.9*inch)
                        # Đường kẻ ngang phía trên mã SP
                        c.line(0.05*inch, 0.45*inch, 3.95*inch, 0.45*inch)
                        
                        # Nội dung chính - Căn chỉnh lại tọa độ theo inch
                        c.setFont("Helvetica-Bold", 9)
                        c.drawString(0.15*inch, 1.7*inch, "NCC:")
                        c.drawString(0.15*inch, 1.4*inch, "NHAN:")
                        c.drawString(0.15*inch, 1.1*inch, "SO PO:")
                        c.drawString(0.15*inch, 0.8*inch, "KIEN SO:")
                        c.drawString(2.2*inch, 0.8*inch, "NGAY:") # Đưa ngày sang bên cạnh Kiện số cho gọn
                        
                        c.setFont("Helvetica", 9)
                        c.drawString(0.7*inch, 1.7*inch, row['NCC'])
                        
                        c.setFont("Helvetica-Bold", 11)
                        c.drawString(0.8*inch, 1.4*inch, row['NHAN'])
                        
                        c.setFont("Helvetica", 9)
                        c.drawString(0.9*inch, 1.1*inch, f"{row['MA_NCC']}/{row['MA_ST']}. {po_id}")
                        
                        # Kiện số và Ngày giao cùng hàng
                        c.setFont("Helvetica-Bold", 11)
                        c.drawString(1.0*inch, 0.8*inch, f"{current_po_tracker[po_id]} / {tong_kien_po}")
                        c.setFont("Helvetica", 9)
                        c.drawString(2.8*inch, 0.8*inch, row['NGAY'])
                        
                        # Thông tin SP ở dưới cùng
                        c.setFont("Helvetica-Bold", 8)
                        c.drawString(0.15*inch, 0.2*inch, f"SP: {row['MA_SP']} - {row['TEN_SP']}")
                        
                        c.showPage()
                        current_po_tracker[po_id] += 1 
                        
                c.save()
                st.download_button("📥 TAI PDF 2x4 INCH", buffer.getvalue(), "Tem_MT_2x4.pdf")
        else:
            st.error("❌ Khong tim thay du lieu.")
    except Exception as e:
        st.error(f"Loi: {str(e)}")
