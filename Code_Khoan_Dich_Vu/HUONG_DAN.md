# Cài đặt Apps Script tạo hợp đồng khoán dịch vụ

## 1. Chuẩn bị Google Docs mẫu

Trong tài liệu mẫu, thay nội dung bôi vàng bằng các biến sau:

| Nội dung trong mẫu | Biến thay thế |
|---|---|
| Phần số hợp đồng `141026-01/LASHE` | `{{SO_HD}}` |
| Ngày ký `14/07/2026` | `{{NGAY_LAM_DON}}` |
| Họ và tên Bên B | `{{HO_TEN}}` |
| Số CCCD | `{{SO_CCCD}}` |
| Mã số thuế | `{{MA_SO_THUE}}` |
| Địa chỉ | `{{DIA_CHI}}` |
| Số điện thoại | `{{SO_DIEN_THOAI}}` |
| Tên Brand | `{{TEN_BRAND}}` |
| Tên dự án | `{{TEN_DU_AN}}` |
| Số lượng tại câu “Sản xuất ... Video” | `{{SO_LUONG_CONG_VIEC}}` |
| Nội dung kịch bản | `{{NOI_DUNG_KICH_BAN}}` |
| Số lượng kịch bản | `{{SO_LUONG_KICH_BAN}}` |
| Số lượng video | `{{SO_LUONG_VIDEO}}` |
| Khung video | `{{KHUNG_VIDEO}}` |
| Chất lượng video | `{{CHAT_LUONG_VIDEO}}` |
| Thời lượng video | `{{THOI_LUONG_VIDEO}}` |
| Hiệu ứng video | `{{HIEU_UNG_VIDEO}}` |
| Nội dung diễn viên | `{{NOI_DUNG_DIEN_VIEN}}` |
| Số lượng diễn viên | `{{SO_LUONG_DIEN_VIEN}}` |
| Ngày hoàn thành | `{{NGAY_HOAN_THANH}}` |
| Tổng thù lao | `{{TONG_THU_LAO}}` |
| Tổng thù lao bằng chữ | `{{TONG_THU_LAO_BANG_CHU}}` |
| Số tài khoản (cột W) | `{{SO_TAI_KHOAN}}` |
| Mở tại (cột X) | `{{MO_TAI}}` |

Nếu ngày trong mẫu được viết tách riêng, có thể dùng:

- Ngày ký: `{{NGAY}}`, `{{THANG}}`, `{{NAM}}`
- Ngày hoàn thành: `{{NGAY_HT}}`, `{{THANG_HT}}`, `{{NAM_HT}}`

Giữ nguyên định dạng chữ của phần bôi vàng khi thay bằng biến. Không chèn kiểu chữ khác nhau vào giữa một biến.

Ở mẫu hiện tại, nếu hai dòng **Số tài khoản:** và **Mở tại:** chưa có biến, code sẽ thêm biến vào bản sao hợp đồng rồi điền dữ liệu. Có thể tự đặt hai biến trên ngay sau dấu hai chấm trong mẫu; code sẽ không thêm lặp lại.

## 2. Cài code

1. Mở **Tiện ích mở rộng → Apps Script** từ Google Sheet.
2. Xóa hàm `myFunction` và dán toàn bộ nội dung file `code.gs`.
3. Bấm **Lưu**.
4. Chọn hàm `caiDatHeThong` và bấm **Chạy** một lần.
5. Chấp nhận quyền truy cập Google Sheets, Docs và Drive.
6. Tải lại Google Sheet.

Sau khi cài đặt, cột W/X vẫn là dữ liệu ngân hàng, cột Y/Z/AA chứa ba link hợp đồng và cột AB có checkbox **TẠO WORD**. Cột W nên được định dạng **Văn bản thuần túy** trước khi nhập số tài khoản để giữ số 0 ở đầu.

Cột U nhận số tiền nguyên dương, kể cả dạng văn bản có dấu phân tách hàng nghìn như `2,160,000` hoặc `2.160.000`. Khi sửa U, cột V tự hiện số tiền bằng chữ, ví dụ `Hai triệu một trăm sáu mươi nghìn`. Nếu xóa U hoặc nhập số không hợp lệ, V được xóa. Lúc tạo hợp đồng, code cũng tính lại V từ U để tránh dùng chữ cũ. Nếu đã cài trigger bằng `caiDatHeThong` trước đó, chỉ cần cập nhật `code.gs` trong Apps Script; trigger hiện có sẽ dùng mã mới.

## 3. Sử dụng

Điền đầy đủ dữ liệu từ cột B đến X, gồm Tên Brand tại cột H, Số tài khoản tại W và Mở tại tại X, rồi tích checkbox tại cột AB. Thiếu bất kỳ trường nào thì script không cấp số hợp đồng và không tạo tài liệu. Cột A do Apps Script tự sinh theo mẫu `DDMMYY-NN/BRAND`; không đặt công thức tại cột A.

Ví dụ ngày làm đơn `14/07/2026`, Brand `LASHE`:

- Hợp đồng đầu tiên: `140726-01/LASHE`
- Hợp đồng tiếp theo cùng ngày và Brand: `140726-02/LASHE`
- Brand khác trong cùng ngày sẽ bắt đầu lại từ `01`.

Bộ đếm chỉ xét những dòng đã có ít nhất một link Word tại Y:AA. Dòng có số ở cột A nhưng chưa có link không được tính. Nếu tạo lại cùng một dòng đã có link, số hợp đồng được giữ nguyên.

Khi tạo lại lần 2 hoặc lần 3 trên cùng một dòng, số hợp đồng ở cột A được giữ nguyên. Link được ghi lần lượt vào cột Y, Z và AA.

Nếu code báo không tìm thấy thư mục, mở thư mục đích trên Drive, sao chép chuỗi nằm sau `/folders/` trong URL và thay giá trị `FOLDER_ID` ở đầu file `code.gs`.

Lưu ý phân biệt chữ `I` hoa và chữ `l` thường trong ID. ID lấy từ ảnh hiện tại là `1FCIPy1gl88pMOlUn4ZtgPIPYflzzmNg`.

## 4. Cột AC và AD – FILE CỨNG

- AC: nút **⬆ TẢI FILE CỨNG**.
- AD: link xem tệp đã tải.
- AE: Trạng thái, code không thay đổi cột này.
- Chấp nhận PDF, JPG hoặc PNG, tối đa 10 MB.
- Mỗi dòng chỉ có một FILE CỨNG và chỉ chủ sở hữu bảng tính được tải.
- Tệp được lưu chung thư mục với các bản hợp đồng của Số hợp đồng tại cột A.

Để bật nút tải tệp:

1. Trong Apps Script, tạo tệp HTML tên `UploadPDF` và dán nội dung file `UploadPDF.html` vào đó.
2. Chọn **Triển khai → Lần triển khai mới → Ứng dụng web**.
3. Chọn thực thi bằng tài khoản của bạn và giới hạn người truy cập theo chính sách công ty.
4. Sao chép URL kết thúc bằng `/exec`, dán vào `CONFIG.WEB_APP_URL` trong `code.gs`. Code chấp nhận cả URL thường `/macros/s/.../exec` và URL Google Workspace `/a/macros/tên-miền/s/.../exec`.
5. Lưu và chạy lại `caiDatHeThong` một lần. Việc này viết nhãn từ cột Y trở đi, không ghi đè dữ liệu ngân hàng ở W/X.

Phần upload này dùng lại cùng giao diện và quy tắc của Đề nghị thanh toán. Code hợp đồng tiếp tục dùng hàm `layHoacTaoThuMucStt_` hiện có để không tạo thêm cấu trúc thư mục khác.
