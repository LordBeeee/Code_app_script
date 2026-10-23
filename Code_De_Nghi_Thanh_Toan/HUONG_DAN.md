# Tải Minh Chứng Bill qua Web App

## Cài đặt một lần

1. Trong dự án Apps Script của bảng tính, cập nhật `code.gs` (tệp `Mã.gs`) và `UploadPDF.html` (tệp HTML tên `UploadPdf`). Không đưa tệp kiểm thử Node.js vào dự án Apps Script. Nếu còn `test_upload_flow.gs`, hãy xóa tệp đó.
2. Tài khoản quản lý thư mục Drive tại `CONFIG.FOLDER_ID` chạy `caiDatHeThong`. Hàm lưu ID bảng tính, tạo link tải riêng cho từng dòng ở cột N và lưu ID thư mục Bill trong cột AB ẩn. Cột V (`Ngày Chi tiền`) được giữ nguyên.
3. Triển khai Web App phiên bản mới với **Thực thi bằng: Tôi** bằng đúng tài khoản đã chạy `caiDatHeThong`. Chọn phạm vi người truy cập phù hợp (ví dụ: người dùng trong tổ chức). Cập nhật `CONFIG.WEB_APP_URL` nếu URL `/exec` đổi, rồi chạy lại `caiDatHeThong` để cập nhật link ở cột N.
4. Tải lại bảng tính. Người tải bấm **↑ TẢI BILL** ở dòng cần tải, chọn PDF/JPG/PNG tối đa 10 MB và bấm **Tải lên**.

## Quyền Drive và kết quả

Web App chạy bằng tài khoản triển khai nên tài khoản này tạo thư mục con `MINH_CHUNG_BILL`, tạo tệp và ghi link vào cột O:Q. Người tải không cần quyền Chỉnh sửa trên thư mục cha, thư mục STT hoặc thư mục Bill. Cột N:Q được bảo vệ để người dùng không sửa nút tải hoặc link đã ghi. Mỗi dòng nhận tối đa ba tệp. Link của từng tệp được đặt **Bất kỳ ai có đường liên kết: Người xem** ngay khi tạo.

Link tải Bill ở cột N có token riêng của dòng. Người có link đó có thể tải lên dòng tương ứng cho đến khi đủ ba tệp; chỉ chia sẻ bảng tính hoặc link tải cho người cần dùng. Nếu Web App công khai cho mọi người, người có được link tải cũng có thể dùng nó. Để thu hồi các link cũ, xóa Script Property `BILL_LINK_SECRET` rồi chạy lại `caiDatHeThong` và triển khai phiên bản mới.

Với thư mục Bill đã tạo ở luồng cũ, hãy kiểm tra và gỡ quyền Chỉnh sửa đã cấp cho người tải trên thư mục Bill, thư mục STT hoặc thư mục cha. Những tệp Bill cũ do người tải sở hữu không tự chuyển thành tệp chỉ xem khi cập nhật mã; cần xử lý quyền và chủ sở hữu riêng trong Drive. Các tệp mới do tài khoản triển khai tạo sẽ thuộc tài khoản đó.

Luồng **FILE CỨNG** vẫn dành cho chủ sở hữu bảng tính qua menu **Tải minh chứng → Tải FILE CỨNG**.
