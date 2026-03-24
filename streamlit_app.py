import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io
import os

# 1. Cấu hình trang (Phải đặt ở đầu)
st.set_page_config(page_title="Tem Chương Dương MT", page_icon="🥤")
st.title("🏷️ Hệ Thống In Tem Tự Động - Team MT")

# 2. Cài đặt Font (Để tránh lỗi dấu tiếng Việt)
font_path = "Arial.ttf" 
if os.path.exists(font_path):
    pdfmetrics.registerFont(TTFont('Arial-Viet', font_path))
    font_main = "Arial-Viet"
else:
    font_main = "Helvetica"

# 3. Tạo ô tải file (Đây là nơi biến 'uploaded_file' được sinh ra)
uploaded_file = st.file_uploader("Tải file Excel dữ liệu tem", type=['xlsx'])

# 4. Kiểm tra nếu đã có file thì mới chạy tiếp
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # Chuẩn hóa tên cột để tránh lỗi KeyError (Viết hoa, xóa khoảng trắng thừa)
    df.columns = [str(col).strip().upper() for col in df.columns]
    
    st.success(f"Đã nhận dữ liệu từ {len(df)} dòng hàng.")
    
    # Cài đặt kích thước tem
    col1, col2 = st.columns(2)
    w_cm = col1.number_input("Rộng (cm)", value=10.0)
    h_cm = col2.number_input("Cao (cm)", value=6.0)

    if st.button("🔥 XUẤT FILE PDF IN TEM"):
        buffer = io.BytesIO()
        width, height = w_cm * cm, h_cm * cm
        c = canvas.Canvas(buffer, pagesize=(width, height))

        for index, row in df.iterrows():
            # Lấy số kiện an toàn
            try:
                # Tìm cột TỔNG SỐ KIỆN (hoặc TONG SO KIEN nếu file không dấu)
                val_tong_kien = row.get('TỔNG SỐ KIỆN', row.get('TONG SO KIEN', 1))
                tong_kien = int(val_tong_kien)
            except:
                tong_kien = 1
            
            for i in range(1, tong_kien + 1):
                # Vẽ khung
                c.setLineWidth(1)
                c.rect(0.2*cm, 0.2*cm, width-0.4*cm, height-0.4*cm)
                c.line(0.2*cm, 1.4*cm, width-0.2*cm, 1.4*cm) 
                c.line(2.8*cm, 1.4*cm, 2.8*cm, height-0.2*cm) 

                # Lấy dữ liệu an toàn từ các cột (Dùng .get để không bị lỗi nếu sai tên cột)
                ncc = str(row.get('NCC', 'CTY CP NGK CHƯƠNG DƯƠNG'))
                noi_nhan = str(row.get('NOI NHẬN', row.get('NƠI NHẬN', '')))
                so_po = str(row.get('SỐ PO :', row.get('SO PO', '')))
                ngay_giao = str(row.get('NGAY GIAO :', row.get('NGÀY GIAO', '')))
                ma_hang = str(row.get('MÃ SẢN PHẨM', row.get('MA SAN PHAM', '')))
                ten_hang = str(row.get('TÊN SẢN PHẨM', row.get('TEN SAN PHAM', '')))

                # Điền nội dung
                c.setFont(font_main, 8)
                c.drawString(0.4*cm, 5.2*cm, "NCC")
                c.drawString(0.4*cm, 4.4*cm, "NƠI NHẬN")
                c.drawString(0.4*cm, 3.6*cm, "SỐ PO:")
                c.drawString(0.4*cm, 2.8*cm, "KIỆN SỐ:")
                c.drawString(0.4*cm, 2.0*cm, "NGÀY GIAO:")

                c.setFont(font_main, 9)
                c.drawString(3.0*cm, 5.2*cm, ncc)
                c.setFont(font_main, 11)
                c.drawString(3.0*cm, 4.4*cm, noi_nhan)
                c.setFont(font_main, 10)
                c.drawString(3.0*cm, 3.6*cm, so_po)
                
                # Kiện số nhảy tự động
                c.setFont(font_main, 12)
                c.drawString(3.0*cm, 2.8*cm, f"{i}  /  {tong_kien}")
                
                c.setFont(font_main, 10)
                c.drawString(3.0*cm, 2.0*cm, ngay_giao)

                c.setFont(font_main, 10)
                c.drawString(0.4*cm, 0.6*cm, ma_hang)
                c.drawRightString(width-0.4*cm, 0.6*cm, ten_hang)

                c.showPage()

        c.save()
        st.download_button(
            label="📥 TẢI FILE PDF TEM",
            data=buffer.getvalue(),
            file_name="Tem_Giao_Hang.pdf",
            mime="application/pdf"
        )
