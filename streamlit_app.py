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
st.title("🏷️ He Thong In Tem Team MT - Final Version")

uploaded_file = st.file_uploader("Tai file Excel (A den J)", type=['xlsx'])

if uploaded_file:
    # DOC FILE: Su dung usecols de ep buoc may phai doc tu cot A den J (0 den 9)
    try:
        # Doc thu de kiem tra so cot thuc te
        df_check = pd.read_excel(uploaded_file, header=None)
        max_cols = len(df_check.columns)
        
        if max_cols < 10:
            st.error(f"⚠️ File hien tai chi co {max_cols} cot. Ban can mo file Excel, bam Save As va dam bao vung du lieu keo dai den tan cot J.")
        
        # Doc chinh thuc lay 10 cot dau tien
        df = pd.read_excel(uploaded_file, header=None, usecols=list(range(min(10, max_cols))))
        
        data_rows = []
        for i, row in df.iterrows():
            # Bo qua dong tieu de va dong trong
            val_a = str(row[0]).upper()
            if "NCC" in val_a or "NHA CUNG CAP" in val_a or val_a == "NAN":
                continue
            
            # Gan du lieu theo toa do anh mau (A=0, B=1, C=2, D=3, E=4, F=5, G=6, I=8, J=9)
            try:
                data_rows.append({
                    'NCC': remove_accents(str(row[0])),
                    'NHAN': remove_accents(str(row[1])),
                    'MA_NCC': str(row[2]),
                    'MA_ST': str(row[3]),
                    'PO': str(row[4]),
                    'MA_SP': str(row[5]),
                    'TEN_SP': remove_accents(str(row[6])),
                    'TONG_KIEN': int(float(row[8])) if pd.notnull(row[8]) else 1,
                    'NGAY': str(row[9])
                })
            except:
                continue # Bo qua neu dong do bi loi dinh dang

        if not data_rows:
            st.warning("Chua tim thay du lieu hop le. Hay kiem tra lai file.")
        else:
            df_final = pd.DataFrame(data_rows)
            # Gop PO de in tem theo bo
            df_gop = df_final.groupby(['PO', 'MA_NCC', 'MA_ST', 'NCC', 'NHAN', 'NGAY'], as_index=False).agg({
                'TONG_KIEN': 'max',
                'MA_SP': 'first',
                'TEN_SP': 'first'
            })

            st.success(f"✅ Da tim thay {len(df_gop)} PO hop le.")
            st.dataframe(df_gop[['PO', 'NHAN', 'TONG_KIEN']])

            if st.button("🚀 XUAT PDF IN TEM"):
                buffer = io.BytesIO()
                c = canvas.Canvas(buffer, pagesize=(10*cm, 6*cm))
                for _, row in df_gop.iterrows():
                    tk = int(row['TONG_KIEN'])
                    po_full = f"{row['MA_NCC']}/{row['MA_ST']}. {row['PO']}"
                    for i in range(1, tk + 1):
                        c.setLineWidth(1); c.rect(0.2*cm, 0.2*cm, 9.6*cm, 5.6*cm)
                        c.line(0.2*cm, 1.4*cm, 9.8*cm, 1.4*cm)
                        c.line(2.8*cm, 1.4*cm, 2.8*cm, 5.8*cm)
                        c.line(6.0*cm, 2.1*cm, 6.0*cm, 3.1*cm)
                        c.setFont("Helvetica-Bold", 8)
                        c.drawString(0.4*cm, 5.2*cm, "NCC"); c.drawString(0.4*cm, 4.4*cm, "NOI NHAN")
                        c.drawString(0.4*cm, 3.6*cm, "SO PO :"); c.drawString(0.4*cm, 2.8*cm, "KIEN SO :")
                        c.drawString(0.4*cm, 2.0*cm, "NGAY GIAO :")
                        c.setFont("Helvetica", 9); c.drawCentredString(6.3*cm, 5.2*cm, row['NCC'])
                        c.setFont("Helvetica-Bold", 11); c.drawCentredString(6.3*cm, 4.4*cm, row['NHAN'])
                        c.setFont("Helvetica", 10); c.drawCentredString(6.3*cm, 3.6*cm, po_full)
                        c.setFont("Helvetica-Bold", 14); c.drawCentredString(4.4*cm, 2.4*cm, str(i))
                        c.drawCentredString(8.0*cm, 2.4*cm, str(tk))
                        c.setFont("Helvetica", 10); c.drawCentredString(6.3*cm, 1.7*cm, row['NGAY'])
                        c.setFont("Helvetica-Bold", 9); c.drawString(0.6*cm, 0.6*cm, row['MA_SP'])
                        c.drawRightString(9.4*cm, 0.6*cm, row['TEN_SP']); c.showPage()
                c.save()
                st.download_button("📥 TAI FILE PDF", buffer.getvalue(), "Tem_TeamMT_Fixed.pdf")

    except Exception as e:
        st.error(f"Loi: {e}")
