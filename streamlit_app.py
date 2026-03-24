import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
import io
import unicodedata

# Hàm xóa dấu tiếng Việt
def remove_accents(input_str):
    if not isinstance(input_str, str): return str(input_str)
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).replace('đ', 'd').replace('Đ', 'D')

st.set_page_config(page_title="In Tem Team MT - Chuong Duong", page_icon="🥤")
st.title("🏷️ He Thong In Tem Tu Dong - Team MT")

uploaded_file = st.file_uploader("Tai file Excel du lieu", type=['xlsx'])

if uploaded_file:
    # Đọc file Excel
    df = pd.read_excel(uploaded_file)
    
    # 1. Chuẩn hóa tên cột (Xóa dấu, viết hoa)
    df.columns = [remove_accents(str(col)).strip().upper() for col in df.columns]

    # 2. Hàm dò tìm cột dựa trên ảnh mẫu
    def find_col(keywords):
        for col in df.columns:
            if any(key in col for key in keywords): return col
        return None

    c_ncc = find_col(['NCC'])
    c_nhan = find_col(['NOI NHAN'])
    c_ma_ncc = find_col(['MA NCC'])
    c_ma_st = find_col(['MA SIEU THI', 'MA ST'])
    c_po = find_col(['SO PO', 'PO'])
    c_ma_sp = find_col(['MA SAN PHAM', 'MA SP'])
    c_ten_sp = find_col(['TEN SAN PHAM', 'TEN SP'])
    c_kien_tong = find_col(['TONG SO KIEN'])
    c_ngay = find_col(['NGAY'])

    if not c_po or not c_kien_tong:
        st.error("❌ Khong tim thay cot 'SO PO' hoac 'TONG SO KIEN'. Hay kiem tra lai tieu de file Excel!")
    else:
        try:
            # Xử lý ngày tháng: Chỉ lấy Ngày/Tháng/Năm
            df[c_ngay] = pd.to_datetime(df[c_ngay], errors='coerce').dt.strftime('%d/%m/%Y')
            
            # Xóa dấu toàn bộ nội dung trong file để in không lỗi font
            df = df.astype(str).applymap(remove_accents)
            
            # Chuyển tổng số kiện sang dạng số để gộp
            df[c_kien_tong] = pd.to_numeric(df[c_kien_tong], errors='coerce').fillna(1).astype(int)

            # --- LOGIC GỘP PO ---
            # Gom các dòng có cùng Số PO, Mã NCC, Mã ST thành 1 tem duy nhất
            group_list = [c_po, c_ma_ncc, c_ma_st, c_ncc, c_nhan, c_ngay]
            df_gop = df.groupby(group_list, as_index=False).agg({
                c_kien_tong: 'max', # Lấy số kiện lớn nhất trong các dòng gộp
                c_ma_sp: 'first',
                c_ten_sp: 'first'
            })

            st.success(f"✅ Da nhan dien va gop thanh {len(df_gop)} PO duy nhat.")
            st.dataframe(df_gop[[c_po, c_nhan, c_kien_tong]])

            if st.button("🚀 XUAT FILE PDF IN TEM"):
                buffer = io.BytesIO()
                # Kích thước tem chuẩn 10cm x 6cm
                c = canvas.Canvas(buffer, pagesize=(10*cm, 6*cm))

                for _, row in df_gop.iterrows():
                    tk = int(row[c_kien_tong])
                    # Hiển thị PO theo định dạng: Mã NCC/Mã ST. Số PO
                    po_display = f"{row[c_ma_ncc]}/{row[c_ma_st]}. {row[c_po]}"

                    for i in range(1, tk + 1):
                        # Vẽ khung và đường kẻ
                        c.setLineWidth(1)
                        c.rect(0.2*cm, 0.2*cm, 9.6*cm, 5.6*cm) # Khung ngoài
                        c.line(0.2*cm, 1.4*cm, 9.8*cm, 1.4*cm) # Đường ngang Mã SP
                        c.line(2.8*cm, 1.4*cm, 2.8*cm, 5.8*cm) # Đường dọc tiêu đề
                        c.line(6.0*cm, 2.1*cm, 6.0*cm, 3.1*cm) # Vạch chia kiện số

                        # --- NHÃN CỐ ĐỊNH ---
                        c.setFont("Helvetica-Bold", 8)
                        c.drawString(0.4*cm, 5.2*cm, "NCC")
                        c.drawString(0.4*cm, 4.4*cm, "NOI NHAN")
                        c.drawString(0.4*cm, 3.6*cm, "SO PO :")
                        c.drawString(0.4*cm, 2.8*cm, "KIEN SO :")
                        c.drawString(0.4*cm, 2.0*cm, "NGAY GIAO :")

                        # --- DỮ LIỆU ---
                        c.setFont("Helvetica", 9)
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

                        # --- DÒNG CUỐI (MÃ & TÊN SP) ---
                        c.setFont("Helvetica-Bold", 9)
                        c.drawString(0.6*cm, 0.6*cm, str(row[c_ma_sp]))
                        c.drawRightString(9.4*cm, 0.6*cm, str(row[c_ten_sp]))
                        
                        c.showPage()

                c.save()
                st.download_button("📥 TAI FILE PDF", buffer.getvalue(), f"Tem_MT_{row[c_ngay].replace('/','_')}.pdf")

        except Exception as e:
            st.error(f"Loi: {e}")
