if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # CHỈNH SỬA TÊN CỘT: Xóa khoảng trắng thừa và viết hoa để khớp với code
    df.columns = [str(col).strip().upper() for col in df.columns]
    
    st.success(f"Đã nhận dữ liệu từ {len(df)} dòng hàng.")
    st.write("Các cột hiện có trong file của bạn:", list(df.columns))

    if st.button("🔥 XUẤT FILE PDF IN TEM"):
        buffer = io.BytesIO()
        width, height = w_cm * cm, h_cm * cm
        c = canvas.Canvas(buffer, pagesize=(width, height))

        for index, row in df.iterrows():
            # Tự động tìm số kiện (viết hoa cột để khớp)
            try:
                tong_kien = int(row.get('TỔNG SỐ KIỆN', row.get('TONG SO KIEN', 1)))
            except:
                tong_kien = 1
            
            for i in range(1, tong_kien + 1):
                # ... (giữ nguyên phần vẽ khung) ...
                
                # Sửa lại cách lấy dữ liệu để không bị lỗi KeyError
                # Sử dụng .get() để nếu không thấy cột thì trả về khoảng trắng thay vì báo lỗi
                ncc = str(row.get('NCC', ''))
                noi_nhan = str(row.get('NOI NHẬN', row.get('NƠI NHẬN', '')))
                so_po = str(row.get('SỐ PO :', row.get('SO PO', '')))
                ngay_giao = str(row.get('NGAY GIAO :', row.get('NGÀY GIAO', '')))
                ma_hang = str(row.get('MÃ SẢN PHẨM', row.get('MA SAN PHAM', '')))
                ten_hang = str(row.get('TÊN SẢN PHẨM', row.get('TEN SAN PHAM', '')))

                # Điền dữ liệu vào tem
                c.setFont(font_main, 9)
                c.drawString(3.2*cm, 5.2*cm, ncc)
                
                c.setFont(font_main, 11)
                c.drawString(3.2*cm, 4.4*cm, noi_nhan)
                
                c.setFont(font_main, 9)
                c.drawString(3.2*cm, 3.6*cm, so_po)
                
                c.setFont(font_main, 12)
                c.drawString(3.2*cm, 2.8*cm, f"{i}  /  {tong_kien}")
                
                c.setFont(font_main, 10)
                c.drawString(3.2*cm, 2.0*cm, ngay_giao)

                c.setFont(font_main, 10)
                c.drawString(0.4*cm, 0.6*cm, ma_hang)
                c.drawRightString(width-0.4*cm, 0.6*cm, ten_hang)

                c.showPage()
