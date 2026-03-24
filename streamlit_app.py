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

st.set_page_config(page_title="In Tem Chuong Duong MT", page_icon="🥤")
st.title("🏷️ He Thong In Tem Tu Dong - Team MT")

uploaded_file = st.file_uploader("Tai file Excel du lieu", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # 1. Chuan hoa ten cot: Xoa dau, xoa khoang trang thua, viet hoa
    original_cols = df.columns.tolist()
    df.columns = [remove_accents(str(col)).strip().upper() for col in df.columns]
    
    # 2. Ham do tim cot linh hoat (Chua tu khoa la lay)
    def find_col_flexible(keywords):
        for col in df.columns:
            if any(key in col for key in keywords):
                return col
        return None

    # Tim kiem cot dua tren cac tu khoa xuat hien trong anh mau
    c_po = find_col_flexible(['SO PO', 'PO'])
    c_ma_ncc = find_col_flexible(['MA NCC'])
    c_ma_st = find_col_flexible(['MA SIEU THI', 'MA ST'])
    c_ncc = find_col_flexible(['NCC', 'NHA CUNG CAP'])
    c_nhan = find_col_flexible(['NOI NHAN'])
    c_ngay = find_col_flexible(['NGAY GIAO', 'NGAY'])
    c_kien_tong = find_col_flexible(['TONG SO KIEN', 'SO KIEN'])
    c_ma_sp = find_col_flexible(['MA SAN PHAM', 'MA SP', 'MA HANG'])
    c_ten_sp = find_col_flexible(['TEN SAN PHAM', 'TEN SP', 'TEN HANG'])

    # KIEM TRA VA BAO LOI CHI TIET
    if not c_po or not c_kien_tong:
        st.error("❌ KHONG TIM THAY COT QUAN TRONG!")
        st.warning(f"Cac cot dang co trong file cua ban: {original_cols}")
        st.info("Meo: Hay dam bao file Excel co cot ten la 'SO PO' va 'TONG SO KIEN'.")
    else:
        try:
            # Hien thi thong tin cot da tim thay de User yen tam
            st.info(f"✅ Da nhan dien: PO ({c_po}), Tong kien ({c_kien_tong})")

            # Xu ly ngay thang (Lay ngay/thang/nam)
            if c_ngay:
                df[c_ngay] = pd.to_datetime(df[c_ngay], errors='coerce').dt.strftime('%d/%m/%Y')

            # Chuan hoa noi dung
            df = df.astype(str).applymap(remove_accents)
            df[c_kien_tong] = pd.to_numeric(df[c_kien_tong], errors='coerce').fillna(1).astype(int)

            # GOP DU LIEU
            group_list = [c for c in [c_po, c_ma_ncc, c_ma_st, c_ncc, c_nhan, c_ngay] if c is not None]
            df_gop = df.groupby(group_list, as_index=False).agg({
                c_kien_tong: 'sum',
                c_ma_sp: 'first',
                c_ten_sp: 'first'
            })

            st.success(f"Da gop thanh {len(df_gop)} PO duy nhat.")

            col1, col2 = st.columns(2)
            w_cm = col1.number_input("Rong tem (cm)", value=10.0)
            h_cm = col2.number_input("Cao tem (cm)", value=6.0)

            if st.button("🚀 XUAT FILE PDF IN TEM"):
                buffer = io.BytesIO()
                width, height = w_cm * cm, h_cm * cm
                c = canvas.Canvas(buffer, pagesize=(width, height))

                for _, row in df_gop.iterrows():
                    tk = int(row[c_kien_tong])
                    po_display = f"{row.get(c_ma_ncc, '')}/{row.get(c_ma_st, '')}. {row.get(c_po, '')}"

                    for i in range(1, tk + 1):
                        c.setLineWidth(1)
                        c.rect(0.2*cm, 0.2*cm, width-0.4*cm, height-0.4*cm)
                        c.line(0.2*cm, 1.4*cm, width-0.2*cm, 1.4*cm)
                        c.line(2.8*cm, 1.4*cm, 2.8*cm, height-0.2*cm)
                        c.line(6.0*cm, 2.1*cm, 6.0*cm, 3.1*cm) # Vach chia kien so

                        c.setFont("Helvetica-Bold", 8)
                        c.drawString(0.4*cm, 5.2*cm, "NCC")
                        c.drawString(0.4*cm, 4.4*cm, "NOI NHAN")
                        c.drawString(0.4*cm, 3.6*cm, "SO PO :")
                        c.drawString(0.4*cm, 2.8*cm, "KIEN SO :")
                        c.drawString(0.4*cm, 2.0*cm, "NGAY GIAO :")

                        c.setFont("Helvetica", 10)
                        c.drawCentredString((width+2.8*cm)/2, 5.2*cm, str(row.get(c_ncc, '')))
                        c.setFont("Helvetica-Bold", 11)
                        c.drawCentredString((width+2.8*cm)/2, 4.4*cm, str(row.get(c_nhan, '')))
                        c.setFont("Helvetica", 10)
                        c.drawCentredString((width+2.8*cm)/2, 3.6*cm, po_display)
                        
                        c.setFont("Helvetica-Bold", 14)
                        c.drawCentredString(4.4*cm, 2.4*cm, str(i))
                        c.drawCentredString(8.0*cm, 2.4*cm, str(tk))
                        
                        c.setFont("Helvetica", 10)
                        c.drawCentredString((width+2.8*cm)/2, 1.7*cm, str(row.get(c_ngay, '')))

                        c.setFont("Helvetica-Bold", 9)
                        c.drawString(0.6*cm, 0.6*cm, str(row.get(c_ma_sp, '')))
                        c.drawRightString(width-0.6*cm, 0.6*cm, str(row.get(c_ten_sp, '')))
                        c.showPage()

                c.save()
                st.download_button("📥 TAI FILE PDF", buffer.getvalue(), "Tem_TeamMT.pdf")

        except Exception as e:
            st.error(f"Loi: {e}")
