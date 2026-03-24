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

st.set_page_config(page_title="In Tem Team MT - Chuong Duong", page_icon="🥤")
st.title("🏷️ In Tem Tu Dong - Phien Ban Moi")

uploaded_file = st.file_uploader("Tai file Excel cua Team MT", type=['xlsx'])

if uploaded_file:
    try:
        # 1. Đọc dữ liệu từ Sheet "TEM MM"
        df_all = pd.read_excel(uploaded_file, sheet_name="TEM MM", header=None)
        
        # Đảm bảo có đủ cột A-J (Index 0-9)
        while len(df_all.columns) < 10:
            df_all[len(df_all.columns)] = ""

        data_rows = []
        
        # 2. Dò tìm dữ liệu từ Cột E (Số PO)
        for _, row in df_all.iterrows():
            val_po = str(row[4]).strip().upper()
            
            # Chỉ lấy dòng nếu cột E có chứa số (Số PO) và không phải là tiêu đề
            if any(char.isdigit() for char in val_po) and val_po not in ["NAN", "SO PO", "PO"]:
                try:
                    data_rows.append({
                        'NCC': remove_accents(str(row[0])),       # Cột A
                        'NHAN': remove_accents(str(row[1])),      # Cột B
                        'MA_NCC': str(row[2]),                   # Cột C
                        'MA_ST': str(row[3]),                    # Dọc D
                        'PO': str(row[4]).strip(),               # Cột E
                        'MA_SP': str(row[5]),                    # Cột F
                        'TEN_SP': remove_accents(str(row[6])),   # Cột G
                        'TONG_KIEN': int(float(row[8])) if str(row[8]).replace('.','').isdigit() else 1, # Cột I
                        'NGAY': str(row[9])                      # Cột J
                    })
                except:
                    continue

        if data_rows:
            df_final = pd.DataFrame(data_rows)
            # Gộp các dòng trùng PO để lấy Tổng số kiện lớn nhất
            df_gop = df_final.groupby(['PO', 'MA_NCC', 'MA_ST', 'NCC', 'NHAN', 'NGAY'], as_index=False).agg({
                'TONG_KIEN': 'max', 'MA_SP': 'first', 'TEN_SP': 'first'
            })

            st.success(f"✅ Đã tìm thấy {len(df_gop)} PO.")
            
            if st.button("🚀 XUAT PDF IN TEM"):
                buffer = io.BytesIO()
                c = canvas.Canvas(buffer, pagesize=(10*cm, 6*cm))
                for _, row in df_gop.iterrows():
                    tk = int(row['TONG_KIEN'])
                    po_hien_thi = f"{row['MA_NCC']}/{row['MA_ST']}. {row['PO']}"
                    
                    for i in range(1, tk + 1):
                        # Vẽ khung ngoài tem 10x6
                        c.setLineWidth(1); c.rect(0.2*cm, 0.2*cm, 9.6*cm, 5.6*cm)
                        # Vẽ đường kẻ ngang phía trên phần mã SP
                        c.line(0.2*cm, 1.4*cm, 9.8*cm, 1.4*cm)
                        
                        # --- Cập nhật giao diện ---
                        # Bỏ đường kẻ dọc đã từng nằm ở đây

                        # 1. Các nhãn cố định
                        c.setFont("Helvetica-Bold", 10)
                        c.drawString(0.4*cm, 5.1*cm, "NCC:")
                        c.drawString(0.4*cm, 4.3*cm, "NHAN:")
                        c.drawString(0.4*cm, 3.5*cm, "SO PO:")
                        c.drawString(0.4*cm, 2.7*cm, "KIEN SO:")
                        
                        # 2. Dữ liệu thay đổi
                        # Chỉnh font chữ NCC
                        c.setFont("Helvetica", 10)
                        c.drawString(1.7*cm, 5.1*cm, row['NCC'])
                        
                        # Chỉnh font Nơi Nhận (MM AN PHÚ)
                        c.setFont("Helvetica-Bold", 13) 
                        c.drawString(2.1*cm, 4.3*cm, row['NHAN'])
                        
                        # Chỉnh font Số PO
                        c.setFont("Helvetica", 10)
                        c.drawString(2.2*cm, 3.5*cm, po_hien_thi)
                        
                        # --- CẬP NHẬT: Số kiện to đã cho nhỏ lại ---
                        c.setFont("Helvetica-Bold", 12) # Chỉnh từ 22 xuống 12 cho đều
                        c.drawCentredString(6.3*cm, 2.7*cm, f"{i} / {tk}")
                        
                        # Dòng dưới cùng
                        c.setFont("Helvetica", 8)
                        c.drawString(0.4*cm, 0.6*cm, f"SP: {row['MA_SP']} - {row['TEN_SP']}")
                        c.drawRightString(9.4*cm, 0.6*cm, f"Ngay: {row['NGAY']}")
                        c.showPage()
                c.save()
                st.download_button("📥 TAI FILE PDF", buffer.getvalue(), "Tem_MT_PhienBanMoi.pdf")
        else:
            st.error("❌ Khong tim thay du lieu PO hop le.")
            
    except Exception as e:
        st.error(f"Loi: {str(e)}")
