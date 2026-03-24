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
    # 1. Đọc toàn bộ file để kiểm tra dòng tiêu đề
    # Thử đọc bỏ qua các dòng trống ở đầu (nếu có)
    df_raw = pd.read_excel(uploaded_file, header=None)
    
    st.info(f"So cot thuc te trong file: {len(df_raw.columns)}")
    
    # Tìm dòng chứa tiêu đề (dòng có chữ 'PO' hoặc 'NCC')
    header_row = 0
    for i, row in df_raw.head(10).iterrows():
        row_str = " ".join(row.astype(str).tolist()).upper()
        if 'PO' in row_str or 'NCC' in row_str:
            header_row = i
            break
            
    # Đọc lại file từ dòng tiêu đề đã tìm thấy
    df = pd.read_excel(uploaded_file, header=header_row)
    df.columns = [remove_accents(str(c)).strip().upper() for c in df.columns]
    
    st.write("Du lieu doc duoc:", df.head(3))

    # 2. Cấu hình cột thủ công (Sidebar) nếu máy nhận diện sai
    st.sidebar.header("Cau hinh cot")
    all_cols = list(df.columns)
    
    def find_default(keys, d_idx):
        for i, c in enumerate(all_cols):
            if any(k in c for k in keys): return i
        return d_idx if d_idx < len(all_cols) else 0

    col_po = st.sidebar.selectbox("Cot SO PO", all_cols, index=find_default(['PO'], 4))
    col_tk = st.sidebar.selectbox("Cot TONG SO KIEN", all_cols, index=find_default(['TONG', 'KIEN'], 8))
    col_ncc = st.sidebar.selectbox("Cot NCC", all_cols, index=find_default(['NCC'], 0))
    col_nhan = st.sidebar.selectbox("Cot NOI NHAN", all_cols, index=find_default(['NHAN'], 1))
    col_ma_ncc = st.sidebar.selectbox("Cot MA NCC", all_cols, index=find_default(['MA NCC'], 2))
    col_ma_st = st.sidebar.selectbox("Cot MA ST", all_cols, index=find_default(['SIEU THI', 'ST'], 3))
    col_ma_sp = st.sidebar.selectbox("Cot MA SP", all_cols, index=find_default(['MA SAN PHAM', 'MA SP'], 5))
    col_ten_sp = st.sidebar.selectbox("Cot TEN SP", all_cols, index=find_default(['TEN'], 6))
    col_ngay = st.sidebar.selectbox("Cot NGAY GIAO", all_cols, index=find_default(['NGAY'], 9))

    try:
        # Làm sạch dữ liệu
        df_clean = pd.DataFrame()
        df_clean['PO'] = df[col_po].astype(str).apply(remove_accents)
        df_clean['TONG_KIEN'] = pd.to_numeric(df[col_tk], errors='coerce').fillna(1).astype(int)
        df_clean['NCC'] = df[col_ncc].astype(str).apply(remove_accents)
        df_clean['NHAN'] = df[col_nhan].astype(str).apply(remove_accents)
        df_clean['MA_NCC'] = df[col_ma_ncc].astype(str).apply(remove_accents)
        df_clean['MA_ST'] = df[col_ma_st].astype(str).apply(remove_accents)
        df_clean['MA_SP'] = df[col_ma_sp].astype(str).apply(remove_accents)
        df_clean['TEN_SP'] = df[col_ten_sp].astype(str).apply(remove_accents)
        
        # Xử lý ngày
        df_clean['NGAY'] = pd.to_datetime(df[col_ngay], errors='coerce').dt.strftime('%d/%m/%Y').fillna(df[col_ngay].astype(str))

        # Gộp PO
        df_gop = df_clean.groupby(['PO', 'MA_NCC', 'MA_ST', 'NCC', 'NHAN', 'NGAY'], as_index=False).agg({
            'TONG_KIEN': 'max',
            'MA_SP': 'first',
            'TEN_SP': 'first'
        })

        st.success(f"✅ San sang in {len(df_gop)} PO.")

        if st.button("🚀 XUAT PDF IN TEM"):
            buffer = io.BytesIO()
            c = canvas.Canvas(buffer, pagesize=(10*cm, 6*cm))
            for _, row in df_gop.iterrows():
                tk = int(row['TONG_KIEN'])
                po_text = f"{row['MA_NCC']}/{row['MA_ST']}. {row['PO']}"
                for i in range(1, tk + 1):
                    c.setLineWidth(1)
                    c.rect(0.2*cm, 0.2*cm, 9.6*cm, 5.6*cm)
                    c.line(0.2*cm, 1.4*cm, 9.8*cm, 1.4*cm)
                    c.line(2.8*cm, 1.4*cm, 2.8*cm, 5.8*cm)
                    c.line(6.0*cm, 2.1*cm, 6.0*cm, 3.1*cm)
                    c.setFont("Helvetica-Bold", 8); c.drawString(0.4*cm, 5.2*cm, "NCC")
                    c.drawString(0.4*cm, 4.4*cm, "NOI NHAN"); c.drawString(0.4*cm, 3.6*cm, "SO PO :")
                    c.drawString(0.4*cm, 2.8*cm, "KIEN SO :"); c.drawString(0.4*cm, 2.0*cm, "NGAY GIAO :")
                    c.setFont("Helvetica", 9); c.drawCentredString(6.3*cm, 5.2*cm, str(row['NCC']))
                    c.setFont("Helvetica-Bold", 11); c.drawCentredString(6.3*cm, 4.4*cm, str(row['NHAN']))
                    c.setFont("Helvetica", 10); c.drawCentredString(6.3*cm, 3.6*cm, po_text)
                    c.setFont("Helvetica-Bold", 14); c.drawCentredString(4.4*cm, 2.4*cm, str(i)); c.drawCentredString(8.0*cm, 2.4*cm, str(tk))
                    c.setFont("Helvetica", 10); c.drawCentredString(6.3*cm, 1.7*cm, str(row['NGAY']))
                    c.setFont("Helvetica-Bold", 9); c.drawString(0.6*cm, 0.6*cm, str(row['MA_SP']))
                    c.drawRightString(9.4*cm, 0.6*cm, str(row['TEN_SP']))
                    c.showPage()
            c.save()
            st.download_button("📥 TAI FILE PDF", buffer.getvalue(), "Tem_TeamMT.pdf")
            
    except Exception as e:
        st.error(f"Loi: {e}. Hay kiem tra lai viec chon cot o menu ben trai.")
