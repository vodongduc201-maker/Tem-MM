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
st.title("🏷️ In Tem: Tach SP - Tong Kien theo PO")

uploaded_file = st.file_uploader("Tai file Excel cua Team MT", type=['xlsx'])

if uploaded_file:
    try:
        # Đọc Sheet "TEM MM"
        df_all = pd.read_excel(uploaded_file, sheet_name="TEM MM", header=None).astype(str)
        
        while len(df_all.columns) < 10:
            df_all[len(df_all.columns)] = ""

        data_rows = []
        for _, row in df_all.iterrows():
            val_po = str(row[4]).strip().upper()
            # Lọc các dòng có số PO
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
                except:
                    continue

        if data_rows:
            df_final = pd.DataFrame(data_rows)
            
            # TÍNH TỔNG SỐ KIỆN THEO PO (Gộp tất cả SP trong cùng 1 PO)
            po_totals = df_final.groupby('PO')['SO_KIEN_SP'].sum().to_dict()

            st.success(f"✅ Da san sang in {len(df_final)} dong san pham.")
            st.write("Bang kiem tra (Tach SP, hien tong kien PO):")
            
            # Hiển thị bảng để user kiểm tra
            df_view = df_final.copy()
            df_view['TONG_PO'] = df_view['PO'].map(po_totals)
            st.dataframe(df_view[['PO', 'MA_SP', 'SO_KIEN_SP', 'TONG_PO']])
            
            if st.button("🚀 XUAT PDF IN TEM"):
                buffer = io.BytesIO()
                c = canvas.Canvas(buffer, pagesize=(10*cm, 6*cm))
                
                # Biến để theo dõi số thứ tự kiện chạy liên tục trong 1 PO
                current_po_tracker = {} 

                for _, row in df_final.iterrows():
                    po_id = row['PO']
                    tong_kien_po = po_totals[po_id]
                    
                    # Khởi tạo số thứ tự kiện cho PO nếu chưa có
                    if po_id not in current_po_tracker:
                        current_po_tracker[po_id] = 1
                    
                    # In số lượng tem tương ứng với số kiện của sản phẩm đó
                    for _ in range(row['SO_KIEN_SP']):
                        # Khung ngoai
                        c.setLineWidth(1); c.rect(0.2*cm, 0.2*cm, 9.6*cm, 5.6*cm)
                        c.line(0.2*cm, 1.2*cm, 9.8*cm, 1.2*cm)
                        
                        # Noi dung
                        c.setFont("Helvetica-Bold", 10)
                        c.drawString(0.4*cm, 5.1*cm, "NCC:")
                        c.drawString(0.4*cm, 4.3*cm, "NHAN:")
                        c.drawString(0.4*cm, 3.5*cm, "SO PO:")
                        c.drawString(0.4*cm, 2.7*cm, "KIEN SO:")
                        c.drawString(0.4*cm, 1.9*cm, "NGAY GIAO:")
                        
                        c.setFont("Helvetica", 10)
                        c.drawString(1.7*cm, 5.1*cm, row['NCC'])
                        c.setFont("Helvetica-Bold", 13)
                        c.drawString(2.1*cm, 4.3*cm, row['NHAN'])
                        c.setFont("Helvetica", 10)
                        c.drawString(2.2*cm, 3.5*cm, f"{row['MA_NCC']}/{row['MA_ST']}. {po_id}")
                        
                        # KIỆN SỐ: Lấy số thứ tự đang chạy / Tổng PO
                        c.setFont("Helvetica-Bold", 12)
                        c.drawString(2.5*cm, 2.7*cm, f"{current_po_tracker[po_id]} / {tong_kien_po}")
                        
                        c.setFont("Helvetica", 11)
                        c.drawString(3.0*cm, 1.9*cm, row['NGAY'])
                        
                        # Dong duoi cung: Ma SP nao ra tem SP do
                        c.setFont("Helvetica-Bold", 9)
                        c.drawString(0.4*cm, 0.5*cm, f"SP: {row['MA_SP']} - {row['TEN_SP']}")
                        
                        c.showPage()
                        current_po_tracker[po_id] += 1 # Tăng số kiện tiếp theo
                        
                c.save()
                st.download_button("📥 TAI FILE PDF", buffer.getvalue(), "Tem_MT_TachSP.pdf")
        else:
            st.error("❌ Khong tim thay du lieu PO.")
    except Exception as e:
        st.error(f"Loi: {str(e)}")
