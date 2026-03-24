import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io
import os

st.set_page_config(page_title="Tem Chương Dương MT", page_icon="🥤")
st.title("🏷️ Hệ Thống In Tem Tự Động - Team MT")

# --- CẤU HÌNH FONT TIẾNG VIỆT ---
# Đảm bảo bạn đã upload file 'Arial.ttf' lên cùng thư mục trên GitHub
font_path = "Arial.ttf" 
if os.path.exists(font_path):
    pdfmetrics.registerFont(TTFont('Arial-Viet', font_path))
    font_main = "Arial-Viet"
else:
    st.warning("⚠️ Không tìm thấy file Arial.ttf trên GitHub. Hệ thống sẽ dùng font mặc định (lỗi dấu).")
    font_main = "Helvetica"

uploaded_file = st.file_uploader("Tải file Excel dữ liệu (Dòng hàng & Tổng kiện)", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success(f"Đã nhận dữ liệu từ {len(df)} dòng hàng.")

    # Tùy chỉnh kích thước tem (thường là 10x6 hoặc 10x15 cho máy in nhiệt)
    col1, col2 = st.columns(2)
    w_cm = col1.number_input("Rộng (cm)", value=10.0)
    h_cm = col2.number_input("Cao (cm)", value=6.0)

    if st.button("🔥 XUẤT FILE PDF IN TEM"):
        buffer = io.BytesIO()
        width, height = w_cm * cm, h_cm * cm
        c = canvas.Canvas(buffer, pagesize=(width, height))

        for index, row in df.iterrows():
            # Lấy số kiện từ cột 'TỔNG SỐ KIỆN'
            try:
                tong_kien = int(row['TỔNG SỐ KIỆN'])
            except:
                tong_kien = 1
            
            # VÒNG LẶP TẠO TEM THEO SỐ KIỆN
            for i in range(1, tong_kien + 1):
                # 1. Vẽ khung viền & đường kẻ
                c.setLineWidth(1)
                c.rect(0.2*cm, 0.2*cm, width-0.4*cm, height-0.4*cm)
                c.line(0.2*cm, 1.4*cm, width-0.2*cm, 1.4*cm) # Ngang cuối
                c.line(2.8*cm, 1.4*cm, 2.8*cm, height-0.2*cm) # Dọc trái

                # 2. Điền tiêu đề cố định
                c.setFont(font_main, 8)
                labels = [("NCC", 5.2), ("NƠI NHẬN", 4.4), ("SỐ PO:", 3.6), ("KIỆN SỐ:", 2.8), ("NGÀY GIAO:", 2.0)]
                for txt, y_pos in labels:
                    c.drawString(0.4*cm, y_pos*cm, txt)

                # 3. Điền dữ liệu từ Excel (Dùng font tiếng Việt)
                c.setFont(font_main, 9)
                c.drawString(3.2*cm, 5.2*cm, str(row['NCC']))
                
                c.setFont(font_main, 11) # Nơi nhận in đậm/to hơn
                c.drawString(3.2*cm, 4.4*cm, str(row['NOI NHẬN']))
                
                c.setFont(font_main, 9)
                c.drawString(3.2*cm, 3.6*cm, str(row['SỐ PO :']))
                
                # Kiện số tự động nhảy: 1/10, 2/10...
                c.setFont(font_main, 12)
                c.drawString(3.2*cm, 2.8*cm, f"{i}  /  {tong_kien}")
                
                c.setFont(font_main, 10)
                c.drawString(3.2*cm, 2.0*cm, str(row['NGAY GIAO :']))

                # 4. Dòng cuối: Mã & Tên sản phẩm
                c.setFont(font_main, 10)
                c.drawString(0.4*cm, 0.6*cm, str(row['MÃ SẢN PHẨM']))
                c.drawRightString(width-0.4*cm, 0.6*cm, str(row['TÊN SẢN PHẨM']))

                c.showPage() # Kết thúc 1 con tem

        c.save()
        st.balloons()
        st.download_button(
            label="📥 TẢI FILE PDF (SẴN SÀNG IN)",
            data=buffer.getvalue(),
            file_name="Tem_Giao_Hang_Chương_Dương.pdf",
            mime="application/pdf"
        )
