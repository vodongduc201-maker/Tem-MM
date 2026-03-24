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
st.title("🏷️ He Thong In Tem Tu Dong - Team MT")

uploaded_file = st.file_uploader("Tai file Excel du lieu", type=['xlsx'])

if uploaded_file:
    # DOC FILE: Khong lay dong dau lam tieu de de tranh loi thieu Header
    df = pd.read_excel(uploaded_file, header=None)
    
    st.warning("⚠️ He thong dang doc file theo thu tu cot (vi file cua ban thieu tieu de).")
    
    # Hien thi thu tu cot de user kiem tra
    with st.expander("Xem thu tu cot du lieu"):
        st.write(df.head(3))

    # KIEM TRA SO LUONG COT (File anh cua ban co khoang 11 cot)
    if len(df.columns) < 5:
        st.error("❌ File Excel qua it cot, khong du du lieu de in tem!")
    else:
        try:
            # Gan ten cot theo thu tu (Dua tren anh mau ban gui)
            # Cot: 0:NCC, 1:Noi Nhan, 2:Ma NCC, 3:Ma ST, 4:PO, 5:Ma SP, 6:Ten SP, 7:Kien So, 8:Tong Kien, 9:Ngay
            
            # Chuan hoa du lieu (Xoa dau)
            df = df.astype(str).applymap(remove_accents)

            # Trich xuat cac cot can thiet theo INDEX (vi tri)
            df_clean = pd.DataFrame()
            df_clean['NCC'] = df[0]
            df_clean['NHAN'] = df[1]
            df_clean['MA_NCC'] = df[2]
            df_clean['MA_ST'] = df[3]
            df_clean['PO'] = df[4]
            df_clean['MA_SP'] = df[5]
            df_clean['TEN_SP'] = df[6]
            df_clean['TONG_KIEN'] = pd.to_numeric(df[8], errors='coerce').fillna(1).astype(int)
            df_clean['NGAY'] = df[9]

            # GOP PO: Neu nhieu dong cung PO thi chi in 1 bo tem voi so kien lon nhat
            df_gop = df_clean.groupby(['PO', 'MA_NCC', 'MA_ST', 'NCC', 'NHAN', 'NGAY'], as_index=False).agg({
                'TONG_KIEN': 'max',
                'MA_SP': 'first',
                'TEN_SP': 'first'
            })

            st.success(f"✅ Da nhan dien duoc {len(df_gop)} PO.")

            if st.button("🚀 XUAT FILE PDF IN TEM"):
                buffer = io.BytesIO()
                c = canvas.Canvas(buffer, pagesize=(10*cm, 6*cm))

                for _, row in df_gop.iterrows():
                    tk = int(row['TONG_KIEN'])
                    po_display = f"{row['MA_NCC']}/{row['MA_ST']}. {row['PO']}"

                    for i in range(1, tk + 1):
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

                        c.setFont("Helvetica", 9)
                        c.drawCentredString(6.3*cm, 5.2*cm, str(row['NCC']))
                        c.setFont("Helvetica-Bold", 11)
                        c.drawCentredString(6.3*cm, 4.4*cm, str(row['NHAN']))
                        c.setFont("Helvetica", 10)
                        c.drawCentredString(6.3*cm, 3.6*cm, po_display)
                        
                        c.setFont("Helvetica-Bold", 14)
                        c.drawCentredString(4.4*cm, 2.4*cm, str(i))
                        c.drawCentredString(8.0*cm, 2.4*cm, str(tk))
                        
                        c.setFont("Helvetica", 10)
                        c.drawCentredString(6.3*cm, 1.7*cm, str(row['NGAY']))

                        c.setFont("Helvetica-Bold", 9)
                        c.drawString(0.6*cm, 0.6*cm, str(row['MA_SP']))
                        c.drawRightString(9.4*cm, 0.6*cm, str(row['TEN_SP']))
                        c.showPage()

                c.save()
                st.download_button("📥 TAI FILE PDF", buffer.getvalue(), "Tem_TeamMT_Fix.pdf")

        except Exception as e:
            st.error(f"Loi: {e}")
