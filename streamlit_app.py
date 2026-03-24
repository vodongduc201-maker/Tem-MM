import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
import io
import unicodedata

# Hàm xóa dấu để in ấn không lỗi
def remove_accents(input_str):
    if not isinstance(input_str, str): return str(input_str)
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).replace('đ', 'd').replace('Đ', 'D')

st.set_page_config(page_title="In Tem Team MT", page_icon="🥤")
st.title("🏷️ He Thong In Tem Team MT - Fixed")

uploaded_file = st.file_uploader("Tai file Excel (Dam bao du lieu bat dau tu Dong 1)", type=['xlsx'])

if uploaded_file:
    # Đọc file KHÔNG lấy tiêu đề để tránh lỗi "Out of bounds"
    df = pd.read_excel(uploaded_file, header=None)
    
    # Kiểm tra số lượng cột thực tế
    num_cols = len(df.columns)
    
    if num_cols < 9:
        st.error(f"❌ File Excel chi co {num_cols} cot, thieu du lieu (Can it nhat 10 cot tu A den J).")
        st.write("Du lieu hien tai may doc duoc:", df.head(5))
    else:
        try:
            # Ép kiểu dữ liệu về chuỗi và xóa dấu
            df = df.astype(str).applymap(remove_accents)

            # ĐỊNH NGHĨA VỊ TRÍ CỘT THEO ẢNH MẪU (A=0, B=1, C=2...)
            # Cột A: NCC, B: Noi Nhan, C: Ma NCC, D: Ma ST, E: So PO, F: Ma SP, G: Ten SP, I: Tong Kien, J: Ngay
            data = []
            for i, row in df.iterrows():
                # Bỏ qua dòng tiêu đề nếu dòng đó chứa chữ "NCC" hoặc "PO"
                if "NCC" in str(row[0]) or "PO" in str(row[4]):
                    continue
                
                data.append({
                    'NCC': row[0],
                    'NHAN': row[1],
                    'MA_NCC': row[2],
                    'MA_ST': row[3],
                    'PO': row[4],
                    'MA_SP': row[5],
                    'TEN_SP': row[6],
                    'TONG_KIEN': int(float(row[8])) if row[8].replace('.','').isdigit() else 1,
                    'NGAY': row[9]
                })
            
            df_final = pd.DataFrame(data)

            # Gộp các mã hàng trùng PO
            df_gop = df_final.groupby(['PO', 'MA_NCC', 'MA_ST', 'NCC', 'NHAN', 'NGAY'], as_index=False).agg({
                'TONG_KIEN': 'max',
                'MA_SP': 'first',
                'TEN_SP': 'first'
            })

            st.success(f"✅ Da san sang in {len(df_gop)} PO.")
            st.dataframe(df_gop[['PO', 'NHAN', 'TONG_KIEN']])

            if st.button("🚀 XUAT PDF IN TEM"):
                buffer = io.BytesIO()
                c = canvas.Canvas(buffer, pagesize=(10*cm, 6*cm))
                
                for _, row in df_gop.iterrows():
                    tk = int(row['TONG_KIEN'])
                    po_text = f"{row['MA_NCC']}/{row['MA_ST']}. {row['PO']}"
                    
                    for i in range(1, tk + 1):
                        # Vẽ khung tem
                        c.setLineWidth(1)
                        c.rect(0.2*cm, 0.2*cm, 9.6*cm, 5.6*cm)
                        c.line(0.2*cm, 1.4*cm, 9.8*cm, 1.4*cm)
                        c.line(2.8*cm, 1.4*cm, 2.8*cm, 5.8*cm)
                        c.line(6.0*cm, 2.1*cm, 6.0*cm, 3.1*cm)

                        # Điền nhãn
                        c.setFont("Helvetica-Bold", 8)
                        c.drawString(0.4*cm, 5.2*cm, "NCC")
                        c.drawString(0.4*cm, 4.4*cm, "NOI NHAN")
                        c.drawString(0.4*cm, 3.6*cm, "SO PO :")
                        c.drawString(0.4*cm, 2.8*cm, "KIEN SO :")
                        c.drawString(0.4*cm, 2.0*cm, "NGAY GIAO :")

                        # Điền dữ liệu
                        c.setFont("Helvetica", 9); c.drawCentredString(6.3*cm, 5.2*cm, str(row['NCC']))
                        c.setFont("Helvetica-Bold", 11); c.drawCentredString(6.3*cm, 4.4*cm, str(row['NHAN']))
                        c.setFont("Helvetica", 10); c.drawCentredString(6.3*cm, 3.6*cm, po_text)
                        
                        # Số kiện
                        c.setFont("Helvetica-Bold", 14)
                        c.drawCentredString(4.4*cm, 2.4*cm, str(i))
                        c.drawCentredString(8.0*cm, 2.4*cm, str(tk))
                        
                        c.setFont("Helvetica", 10); c.drawCentredString(6.3*cm, 1.7*cm, str(row['NGAY']))
                        c.setFont("Helvetica-Bold", 9); c.drawString(0.6*cm, 0.6*cm, str(row['MA_SP']))
                        c.drawRightString(9.4*cm, 0.6*cm, str(row['TEN_SP']))
                        c.showPage()
                
                c.save()
                st.download_button("📥 TAI FILE PDF", buffer.getvalue(), "Tem_TeamMT_Final.pdf")
        
        except Exception as e:
            st.error(f"Loi: {e}")
