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
    # 1. Đọc file
    df = pd.read_excel(uploaded_file)
    
    # Kiểm tra số lượng cột thực tế
    num_cols = len(df.columns)
    st.info(f"So cot may doc duoc tu file: {num_cols}")

    # 2. Chuẩn hóa tên cột
    df.columns = [remove_accents(str(c)).strip().upper() for c in df.columns]

    # 3. Hàm dò tìm cột an toàn (Kết hợp tìm tên và vị trí)
    def safe_get_col(df, keywords, default_idx):
        # Ưu tiên tìm theo tên trước
        for i, col in enumerate(df.columns):
            if any(k in col for k in keywords):
                return i
        # Nếu không thấy tên, kiểm tra xem vị trí mặc định có tồn tại không
        if default_idx < len(df.columns):
            return default_idx
        return None

    # Lấy vị trí các cột (Dựa trên sơ đồ A=0, B=1... J=9)
    i_ncc = safe_get_col(df, ['NCC'], 0)
    i_nhan = safe_get_col(df, ['NOI NHAN', 'NHAN'], 1)
    i_ma_ncc = safe_get_col(df, ['MA NCC'], 2)
    i_ma_st = safe_get_col(df, ['MA SIEU THI', 'MA ST'], 3)
    i_po = safe_get_col(df, ['SO PO', 'PO'], 4)
    i_ma_sp = safe_get_col(df, ['MA SAN PHAM', 'MA SP'], 5)
    i_ten_sp = safe_get_col(df, ['TEN SAN PHAM', 'TEN SP'], 6)
    i_tkien = safe_get_col(df, ['TONG SO KIEN', 'TONG KIEN'], 8)
    i_ngay = safe_get_col(df, ['NGAY'], 9)

    # Kiểm tra xem có đủ các cột tối thiểu (PO và Tổng kiện) không
    if i_po is None or i_tkien is None:
        st.error(f"❌ File Excel thieu cot quan trong (PO hoac Tong Kien). File chi co {num_cols} cot.")
        st.write("Cac cot may dang thay:", list(df.columns))
    else:
        try:
            # Tạo DataFrame sạch
            df_clean = pd.DataFrame()
            df_clean['NCC'] = df.iloc[:, i_ncc].astype(str).apply(remove_accents)
            df_clean['NHAN'] = df.iloc[:, i_nhan].astype(str).apply(remove_accents)
            df_clean['MA_NCC'] = df.iloc[:, i_ma_ncc].astype(str).apply(remove_accents)
            df_clean['MA_ST'] = df.iloc[:, i_ma_st].astype(str).apply(remove_accents)
            df_clean['PO'] = df.iloc[:, i_po].astype(str).apply(remove_accents)
            df_clean['MA_SP'] = df.iloc[:, i_ma_sp].astype(str).apply(remove_accents)
            df_clean['TEN_SP'] = df.iloc[:, i_ten_sp].astype(str).apply(remove_accents)
            df_clean['TONG_KIEN'] = pd.to_numeric(df.iloc[:, i_tkien], errors='coerce').fillna(1).astype(int)
            
            # Xử lý ngày (An toàn)
            if i_ngay is not None:
                raw_ngay = df.iloc[:, i_ngay]
                df_clean['NGAY'] = pd.to_datetime(raw_ngay, errors='coerce').dt.strftime('%d/%m/%Y').fillna(str(raw_ngay))
            else:
                df_clean['NGAY'] = ""

            # Gộp dữ liệu theo PO
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
            st.error(f"Loi xu ly du lieu: {e}")
