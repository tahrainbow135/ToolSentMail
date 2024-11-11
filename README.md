# Email Sending Tool

## Tổng quan
Đây là một công cụ Python giúp tự động hóa việc gửi email cá nhân hóa bằng máy chủ SMTP của Gmail. Công cụ này đọc thông tin người nhận từ một tệp Excel và gửi email tùy chỉnh đến từng người nhận.

## Tính năng
- Đọc dữ liệu người nhận (tên, email, thời gian, địa điểm) từ tệp Excel.
- Gửi email qua Gmail với nội dung có thể tùy chỉnh.
- Tạo tệp tóm tắt ghi lại danh sách người nhận thành công.

## Yêu cầu
Đảm bảo Python đã được cài đặt trên hệ thống của bạn. Có thể tải Python từ [python.org](https://www.python.org/downloads/).

### Thư viện cần thiết
Trước khi chạy script, cài đặt các thư viện Python cần thiết:

```bash
pip install pandas openpyxl
```

## Hướng dẫn sử dụng

1. **Tải về hoặc sao chép mã nguồn chứa script này.**
2. **Cập nhật đường dẫn đến tệp Excel**:
   Sửa dòng đọc tệp Excel:
   ```python
   df = pd.read_excel("D://User//Downloads//cv2.xlsx")
   ```
   Thay đường dẫn này bằng vị trí tệp của bạn.

3. **Cập nhật thông tin tài khoản Gmail**:
   Đảm bảo các biến `GMAIL_USER` và `GMAIL_PASSWORD` được đặt đúng(Nhắn tin cho Tú Anh).
   ```python
   GMAIL_USER = "your-email@gmail.com"
   GMAIL_PASSWORD = "your-app-password"
   ```

   > **Lưu ý**: Nên sử dụng [Mật khẩu ứng dụng](https://support.google.com/accounts/answer/185833?hl=vi) để bảo mật tốt hơn.

4. **Chạy script**:
   Chạy script trong môi trường Python hoặc terminal:
   ```bash
   python email_tool.py
   ```

5. **Kiểm tra kết quả**:
   - Script sẽ in thông báo xác nhận khi email được gửi thành công.
   - Một tệp văn bản `email_recipients.txt` sẽ được tạo ở vị trí đã chỉ định (`D://User//Downloads//`), chứa danh sách người nhận thành công.

## Cấu hình
- **Định dạng tệp Excel**:
   Đảm bảo tệp Excel của bạn có các cột sau:
   - `Name`: Tên người nhận.
   - `Email`: Địa chỉ email của người nhận.
   - `Time`: Thời gian sự kiện (ví dụ: "9h00, chủ nhật, ngày 24, tháng 09, năm 2023").
   - `Location`: Địa điểm sự kiện.

- **Tùy chỉnh nội dung email**:
   Bạn có thể tùy chỉnh mẫu email trong biến `html_body` trong script.

## Lưu ý bảo mật
- Tránh việc ghi trực tiếp mật khẩu Gmail vào script. Sử dụng biến môi trường hoặc mã hóa bí mật để bảo mật tốt hơn.
- Đảm bảo tuân thủ các hướng dẫn bảo mật của Gmail khi truy cập SMTP.

## Khắc phục sự cố
- Nếu gặp lỗi liên quan đến máy chủ Gmail hoặc đăng nhập, đảm bảo rằng:
   - Đã bật quyền truy cập cho ứng dụng kém an toàn trên tài khoản Gmail của bạn.
   - Đã sử dụng [Mật khẩu ứng dụng](https://support.google.com/accounts/answer/185833?hl=vi).

## Giấy phép
Công cụ này được cung cấp nguyên trạng mà không có bất kỳ bảo đảm nào. Sử dụng một cách có trách nhiệm và tự chịu rủi ro.
```

Bạn có thể điều chỉnh các phần trong `README.md` để phù hợp hơn với nhu cầu của mình.