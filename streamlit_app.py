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
st.title("🏷️ He Thong In Tem Tu Dong - Team MT")

uploaded_file = st.file_uploader("Tai file Excel du lieu", type=['xlsx'])

if uploaded_file:
    # DOC FILE: Thu doc co tieu de truoc
    df = pd.read_excel(uploaded_file)
    
    # Kiem tra xem dong dau co phai la du lieu khong (neu la so thi kha nang cao thieu header)
    first_col_val = str(df.columns[0])
    if first_col_val.isdigit() or "SAXI" in first_col_val.upper():
        # Doc lai file va tu dat ten cot la 0, 1, 2, 3...
        df = pd.read_excel(uploaded_file, header=None)
        st.warning("⚠️ File cua ban hinh nhu thieu dong tieu de (Header). He thong se tu danh so cot.")

    # Chuan hoa ten cot de de xu ly
    df.columns = [remove_accents(str(col)).strip().upper() for col in df.columns]
    
    st.write("Du lieu hien tai:", df.head(3))

    # CHO PHEP USER CHON COT NEU KHONG TU NHAN DIEN DUOC
    st.sidebar.header("Cấu hình cột dữ liệu")
    all_cols = list(df.columns)
    
    def auto_select(keys, default_idx):
        for i, col in enumerate(all_cols):
            if any(k in str(col) for k in keys): return col
        return all_cols[default_idx] if len(all_cols) > default_idx else all_cols[0]

    c_ncc = st.sidebar.selectbox("Cot Ten NCC", all_cols, index=all_cols.index(auto_select(['NCC'], 0)))
    c_nhan = st.sidebar.selectbox("Cot Noi Nhan", all_cols, index=all_cols.index(auto_select(['NHAN'], 1)))
    c_ma_ncc = st.sidebar.selectbox("Cot Ma NCC", all_cols, index=all_cols.index(auto_select(['MA NCC'], 2)))
    c_ma_st = st.sidebar.selectbox("Cot Ma Sieu Thi", all_cols, index=all_cols.index(auto_select(['SIEU THI', 'ST'], 3)))
    c_po = st.sidebar.selectbox("Cot So PO", all_cols, index=all_cols.index(auto_select(['PO'], 4)))
    c_ma_sp = st.sidebar.selectbox("Cot Ma SP", all_cols, index=all_cols.index(auto_select(['MA SAN PHAM', 'MA SP'], 5)))
    c_ten_sp = st.sidebar.selectbox("Cot Ten SP", all_cols, index=all_cols.index(auto_select(['TEN'], 6)))
    c_kien_tong = st.sidebar.selectbox("Cot Tong So Kien", all_cols, index=all_cols.index(auto_select(['TONG'], 8)))
    c_ngay = st.sidebar.selectbox("Cot Ngay Giao", all_cols, index=all_cols.index(auto_select(['NGAY'], 9)))

    try:
        # Xu ly ngay thang
        df[c_ngay] = pd.to_datetime(df[c_ngay], errors='coerce').dt.strftime('%d/%m/%Y')
        df = df.astype(str).applymap(remove_accents)
        df[c_kien_tong] = pd.to_numeric(df[c_kien_tong], errors='coerce').fillna(1).astype(int)

        # GOP DU LIEU
        group_list = [c_po, c_ma_ncc, c_ma_st, c_ncc, c_nhan, c_ngay]
        df_gop = df.groupby(group_list, as_index=False).agg({
            c_kien_tong: 'sum',
            c_ma_sp: 'first',
            c_ten_sp: 'first'
        })

        st.success(f"Da gop thanh {len(df_gop)} PO duy nhat.")

        if st.button("🚀 XUAT FILE PDF IN TEM"):
            buffer = io.BytesIO()
            c = canvas.Canvas(buffer, pagesize=(10*cm, 6*cm))

            for _, row in df_gop.iterrows():
                tk = int(row[c_kien_tong])
                po_display = f"{row[c_ma_ncc]}/{row[c_ma_st]}. {row[c_po]}"

                for i in range(1, tk + 1):
                    c.setLineWidth(1)
                    c.rect(0.2*cm, 0.2*cm, 9.6*cm, 5.6*cm)
                    c.line(0.2*cm, 1.4*cm, 9.8*cm, 1.4*cm)
                    c.line(2.8*cm, 1.4*cm, 2.8*cm, 5.8*cm)
                    c.line(6.0*cm, 2.2*cm, 6.0*cm, 3.0*cm)

                    c.setFont("Helvetica-Bold", 8)
                    c.drawString(0.4*cm, 5.2*cm, "NCC")
                    c.drawString(0.4*cm, 4.4*cm, "NOI NHAN")
                    c.drawString(0.4*cm, 3.6*cm, "SO PO :")
                    c.drawString(0.4*cm, 2.8*cm, "KIEN SO :")
                    c.drawString(0.4*cm, 2.0*cm, "NGAY GIAO :")

                    c.setFont("Helvetica", 10)
                    c.drawCentredString(6.3*cm, 5.2*cm, str(row[c_ncc]))
                    c.setFont("Helvetica-Bold", 11)
                    c.drawCentredString(6.3*cm, 4.4*cm, str(row[c_nhan]))
                    c.setFont("Helvetica", 10)
                    c.drawCentredString(6.3*cm, 3.6*cm, po_display)
                    
                    c.setFont("Helvetica-Bold", 14)
                    c.drawCentredString(4.4*cm, 2.4*cm, str(i))
                    c.drawCentredString(8.0*cm, 2.4*cm, str(tk))
                    
                    c.setFont("Helvetica", 10)
                    c.drawCentredString(6.3*cm, 1.7*cm, str(row[c_ngay]))

                    c.setFont("Helvetica-Bold", 9)
                    c.drawString(0.6*cm, 0.6*cm, str(row[c_ma_sp]))
                    c.drawRightString(9.4*cm, 0.6*cm, str(row[c_ten_sp]))
                    c.showPage()

            c.save()
            st.download_button("📥 TAI FILE PDF", buffer.getvalue(), "Tem_TeamMT.pdf")

    except Exception as e:
        st.error(f"Loi: {e}")
        st.info("Hay kiem tra xem cac cot o menu ben trai da chon dung chua.")
