import streamlit as st
import pandas as pd
import yagmail
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

def send_emails(lecturers_data, sender_email, app_password, test_mode=True):
    """Gửi email cho giảng viên - ĐÃ SỬA HIỂN THỊ EMAIL"""
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    results = st.container()
    
    total_lecturers = len(lecturers_data)
    success_count = 0
    error_count = 0
    
    try:
        # Kết nối email server (chỉ khi gửi thật)
        if not test_mode:
            try:
                yag = yagmail.SMTP(
                    user=sender_email,
                    password=app_password,
                    host='smtp.gmail.com',
                    port=587,
                    smtp_starttls=True,
                    smtp_ssl=False
                )
                status_text.success("✅ Kết nối email thành công!")
            except Exception as e:
                st.error(f"❌ Lỗi kết nối yagmail: {e}")
                return send_emails_direct_smtp(lecturers_data, sender_email, app_password)
        
        # Gửi email cho từng giảng viên
        for i, (email, data) in enumerate(lecturers_data.items()):
            progress = (i + 1) / total_lecturers
            progress_bar.progress(progress)
            
            status_text.text(f"Đang xử lý: {data['name']} ({i+1}/{total_lecturers})")
            
            # Tạo nội dung email
            subject = f"THÔNG BÁO LỊCH THI - {data['name'].upper()}"
            email_content = generate_email_content(data)
            
            try:
                if test_mode:
                    # Chế độ test
                    with results.expander(f"🧪 TEST: {data['name']}", expanded=False):
                        st.write(f"**Email:** {email}")
                        st.write(f"**Tiêu đề:** {subject}")
                        st.write(f"**Số lớp:** {len(data['classes'])}")
                        st.components.v1.html(email_content, height=500, scrolling=True)
                    success_count += 1
                else:
                    # Gửi email thật
                    yag.send(
                        to=email,
                        subject=subject,
                        contents=email_content
                    )
                    results.success(f"✅ Đã gửi: {data['name']}")
                    success_count += 1
                    
            except Exception as e:
                error_count += 1
                results.error(f"❌ Lỗi {data['name']}: {str(e)}")
        
        # Hiển thị kết quả tổng
        progress_bar.empty()
        status_text.empty()
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Thành công", success_count)
        with col2:
            st.metric("Lỗi", error_count)
        with col3:
            st.metric("Tổng", total_lecturers)
        
        if test_mode:
            st.info("🎯 Đây là chế độ kiểm tra. Để gửi email thật, hãy bỏ chọn 'Chế độ kiểm tra'")
        else:
            st.balloons()
            st.success(f"🎉 Đã gửi thành công {success_count}/{total_lecturers} email!")
            
    except Exception as e:
        st.error(f"❌ Lỗi kết nối email: {e}")

def clean_data(value):
    """Làm sạch dữ liệu - loại bỏ khoảng trắng thừa và giá trị rỗng"""
    if pd.isna(value) or value == '' or value is None:
        return ""
    cleaned = str(value).strip()
    cleaned = ' '.join(cleaned.split())
    return cleaned

def format_date(date_value):
    """Định dạng ngày tháng cho email"""
    if not date_value:
        return ""
    
    if isinstance(date_value, datetime):
        return date_value.strftime("%d/%m/%Y")
    elif isinstance(date_value, str):
        # Thử parse string thành date
        try:
            if '-' in date_value:
                date_obj = datetime.strptime(date_value.split()[0], "%Y-%m-%d")
                return date_obj.strftime("%d/%m/%Y")
        except:
            pass
    return str(date_value)

def generate_email_content(lecturer_data):
    """Tạo nội dung email HTML - FIX LỖI HIỂN THỊ KHÁC NHAU"""
    
    if not lecturer_data['classes']:
        return "<p>Không có lịch thi nào.</p>"
    
    # Lọc dữ liệu
    valid_classes = []
    for class_info in lecturer_data['classes']:
        cleaned_info = {key: clean_data(value) for key, value in class_info.items()}
        data_fields = sum(1 for value in cleaned_info.values() if value != "")
        if data_fields >= 2:
            valid_classes.append(cleaned_info)
    
    if not valid_classes:
        return "<p>Không có lịch thi nào.</p>"
    
    # Tạo bảng với CSS TỐI ƯU CHO EMAIL
    table_html = """
    <table width="600" cellpadding="2" cellspacing="0" border="1" bgcolor="#FFFFFF" style="border-collapse: collapse; font-family: Arial, Helvetica, sans-serif; font-size: 11px; line-height: 1.1; mso-cellspacing: 0px;">
        <tr style="background-color: #2E86AB; color: white;">
            <th width="30" style="padding: 4px; border: 1px solid #cccccc; text-align: center; mso-padding-alt: 4px;"><strong>STT</strong></th>
            <th width="150" style="padding: 4px; border: 1px solid #cccccc; mso-padding-alt: 4px;"><strong>Ngành</strong></th>
            <th width="60" style="padding: 4px; border: 1px solid #cccccc; mso-padding-alt: 4px;"><strong>Lớp</strong></th>
            <th width="120" style="padding: 4px; border: 1px solid #cccccc; mso-padding-alt: 4px;"><strong>Môn thi</strong></th>
            <th width="80" style="padding: 4px; border: 1px solid #cccccc; mso-padding-alt: 4px;"><strong>Hình thức thi</strong></th>
            <th width="80" style="padding: 4px; border: 1px solid #cccccc; mso-padding-alt: 4px;"><strong>Ngày thi</strong></th>
            <th width="60" style="padding: 4px; border: 1px solid #cccccc; mso-padding-alt: 4px;"><strong>Giờ thi</strong></th>
        </tr>
    """
    
    displayed_count = 0
    for i, class_info in enumerate(valid_classes, 1):
        nganh = class_info.get('Nganh', '')
        lop = class_info.get('Lop', '')
        mon_thi = class_info.get('Mon_thi', '')
        hinh_thuc = class_info.get('Hinh_thuc_thi', '')
        ngay_thi = format_date(class_info.get('Ngay_thi', ''))
        gio_thi = class_info.get('Gio_thi', '')
        
        non_empty_fields = [field for field in [nganh, lop, mon_thi, hinh_thuc, ngay_thi, gio_thi] if field]
        
        if len(non_empty_fields) >= 2:
            displayed_count += 1
            bg_color = "#F8F9FA" if displayed_count % 2 == 0 else "#FFFFFF"
            
            table_html += f"""
            <tr style="background-color: {bg_color};">
                <td style="padding: 3px; border: 1px solid #cccccc; text-align: center; vertical-align: top; mso-padding-alt: 3px;">{displayed_count}</td>
                <td style="padding: 3px; border: 1px solid #cccccc; vertical-align: top; mso-padding-alt: 3px;">{nganh}</td>
                <td style="padding: 3px; border: 1px solid #cccccc; vertical-align: top; mso-padding-alt: 3px;">{lop}</td>
                <td style="padding: 3px; border: 1px solid #cccccc; vertical-align: top; mso-padding-alt: 3px;">{mon_thi}</td>
                <td style="padding: 3px; border: 1px solid #cccccc; vertical-align: top; mso-padding-alt: 3px;">{hinh_thuc}</td>
                <td style="padding: 3px; border: 1px solid #cccccc; vertical-align: top; mso-padding-alt: 3px;">{ngay_thi}</td>
                <td style="padding: 3px; border: 1px solid #cccccc; vertical-align: top; mso-padding-alt: 3px;">{gio_thi}</td>
            </tr>
            """
    
    table_html += "</table>"
    
    # Container email với CSS TỐI ƯU
    return f"""
    <div style="font-family: Arial, Helvetica, sans-serif; font-size: 12px; line-height: 1.2; color: #333333; width: 600px; max-width: 600px;">
        <!-- Header -->
        <div style="background: #2E86AB; padding: 4px 15px; color: white;">
            <h2 style="margin: 0; font-size: 14px; font-weight: bold; line-height: 1.2;">THÔNG BÁO LỊCH THI</h2>
        </div>
        
        <!-- Content -->
        <div style="padding: 15px; background-color: #ffffff;">
            <p style="margin: 0 0 8px 0; line-height: 1.2;"><strong>Kính gửi:</strong> {lecturer_data['name']}</p>
            
            <p style="margin: 0 0 10px 0; line-height: 1.2;">Thông tin lịch thi các lớp giảng viên phụ trách:</p>
            
            {table_html}
            
            <div style="margin-top: 15px; padding: 8px 10px; background-color: #F1F5F9; border-left: 4px solid #2E86AB;">
                <p style="margin: 0 0 4px 0; font-weight: bold; color: #2E86AB; line-height: 1.2;">Lưu ý:</p>
                <ul style="margin: 0; padding-left: 15px; line-height: 1.2;">
                    <li style="margin-bottom: 2px;">Vui lòng kiểm tra kỹ thông tin lịch thi</li>
                    <li style="margin-bottom: 2px;">Liên hệ Phòng Đào tạo nếu có thắc mắc</li>
                    <li>Đảm bảo có mặt tại phòng thi trước 15 phút</li>
                </ul>
            </div>
            
            <div style="margin-top: 15px; padding-top: 12px; border-top: 1px solid #E2E8F0;">
                <p style="margin: 0; font-style: italic; color: #666666; line-height: 1.2;">
                    Trân trọng,<br>
                    <strong style="color: #2E86AB;">Phòng Đào tạo</strong>
                </p>
            </div>
        </div>
    </div>
    """

def generate_preview_content(lecturer_data):
    """Tạo preview - DÙNG CÙNG CODE VỚI EMAIL để đồng bộ"""
    return generate_email_content(lecturer_data)

def generate_preview_content(lecturer_data):
    """Tạo nội dung preview cho Streamlit - GIỐNG VỚI EMAIL"""
    # Sử dụng cùng hàm với email để đảm bảo giống nhau
    return generate_email_content(lecturer_data)

def main():
    """Hàm chính của ứng dụng"""
    st.title("📧 Hệ Thống Gửi Email Tự Động")
    st.markdown("---")
    
    # Sidebar - Cấu hình email
    with st.sidebar:
        st.header("⚙️ Cấu hình Email")
        
        email_sender = st.text_input("Email gửi", placeholder="your_email@gmail.com")
        app_password = st.text_input("Mật khẩu ứng dụng", type="password", placeholder="Nhập app password")
        
        st.markdown("---")
        st.info("""
        **Hướng dẫn cấu hình Gmail:**
        1. Bật xác thực 2 bước
        2. Tạo mật khẩu ứng dụng
        3. Nhập thông tin vào form bên trái
        """)
    
    # Tab chính
    tab1, tab2 = st.tabs(["📤 Tải file & Gửi email", "👀 Xem trước email"])
    
    with tab1:
        st.header("Tải file Excel lên")
        
        uploaded_file = st.file_uploader("Chọn file Excel", type=['xlsx', 'xls'])
        
        if uploaded_file is not None:
            try:
                # Đọc file Excel
                df = pd.read_excel(uploaded_file)
                st.success(f"✅ Đã tải file thành công! Tổng số dòng: {len(df)}")
                
                # Hiển thị preview
                with st.expander("👁️ Xem trước toàn bộ dữ liệu"):
                    st.dataframe(df)
                    st.write("**Tên các cột:**", list(df.columns))
                
                # Xử lý dữ liệu
                df_clean = df[df['Email'].notna() & (df['Email'] != '')].fillna('')
                
                # Nhóm dữ liệu theo giảng viên
                lecturers_data = {}
                for _, row in df_clean.iterrows():
                    email = row['Email']
                    if email and '@' in email:
                        if email not in lecturers_data:
                            lecturers_data[email] = {
                                'name': clean_data(row.get('Giang_vien', '')),
                                'classes': []
                            }
                        
                        class_info = {
                            'Nganh': clean_data(row.get('Nganh', '') or row.get('Ngành', '')),
                            'Lop': clean_data(row.get('Lop', '') or row.get('Lớp', '')),
                            'Mon_thi': clean_data(row.get('Hoc_phan', '') or row.get('Môn_thi', '') or row.get('Học_phần', '')),
                            'Hinh_thuc_thi': clean_data(row.get('Hinh_thuc_thi', '') or row.get('Hình_thức_thi', '')),
                            'Ngay_thi': row.get('Ngay', '') or row.get('Ngày', '') or row.get('Ngay_thi', '') or row.get('Ngày_thi', ''),
                            'Gio_thi': clean_data(row.get('Gio_thi', '') or row.get('Giờ_thi', ''))
                        }
                        lecturers_data[email]['classes'].append(class_info)
                
                # Hiển thị thống kê
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Tổng giảng viên", len(lecturers_data))
                with col2:
                    total_classes = sum(len(data['classes']) for data in lecturers_data.values())
                    st.metric("Tổng lớp học", total_classes)
                with col3:
                    st.metric("Email hợp lệ", len([email for email in lecturers_data.keys() if '@' in email]))
                
                # Nút gửi email
                st.markdown("---")
                st.subheader("Gửi email")
                
                col1, col2 = st.columns(2)
                with col1:
                    test_mode = st.checkbox("Chế độ kiểm tra (không gửi thật)", value=True)
                with col2:
                    send_button = st.button("🚀 Gửi email", type="primary", use_container_width=True)
                
                if send_button:
                    if not email_sender or not app_password:
                        st.error("❌ Vui lòng nhập đầy đủ thông tin email và mật khẩu!")
                    else:
                        send_emails(lecturers_data, email_sender, app_password, test_mode)
                        
            except Exception as e:
                st.error(f"❌ Lỗi khi đọc file: {e}")
    
    with tab2:
        st.header("Xem trước mẫu email")
        
        if 'lecturers_data' in locals() and lecturers_data:
            lecturer_emails = list(lecturers_data.keys())
            selected_email = st.selectbox("Chọn giảng viên", lecturer_emails)
            
            if selected_email:
                # Sử dụng cùng hàm với email thật
                preview_content = generate_preview_content(lecturers_data[selected_email])
                st.components.v1.html(preview_content, height=600, scrolling=True)
        else:
            st.info("📁 Vui lòng tải file Excel ở tab đầu tiên")

if __name__ == "__main__":
    main()
