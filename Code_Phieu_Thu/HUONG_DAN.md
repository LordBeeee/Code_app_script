# Phiếu thu từ TH_PHIEU_THU

## Cấu trúc theo Sheet hiện tại

| Cột | Tiêu đề | Chức năng |
|---|---|---|
| A | Ngày lập phiếu | Ngày lập phiếu thu |
| B | Họ và Tên | Người nộp tiền |
| C | Địa chỉ | Địa chỉ người nộp |
| D | Số tiền | Số tiền nguyên dương |
| E | Số tiền bằng chữ | Tự cập nhật khi nhập/sửa/xóa D |
| F | Link_Word_1 | Link Google Docs lần 1 |
| G | Link_Word_2 | Link Google Docs lần 2 |
| H | Link_Word_3 | Link Google Docs lần 3 |
| I | Tạo_Word | Checkbox tạo Google Docs |

Mỗi lần tích I sẽ tạo một bản Google Docs mới và ghi URL vào ô link trống đầu tiên trong F, G, H cùng dòng. Đủ cả ba ô sẽ báo lỗi và không tạo thêm. Không ghi đè ô có dữ liệu hoặc công thức. Không xuất PDF hoặc Word .docx. Các cột được nhận diện theo tiêu đề nên vẫn dùng được khi thay đổi thứ tự; hàng tiêu đề mặc định là hàng 1.

Trạng thái đang tạo, đã tạo và lỗi được ghi dưới dạng **ghi chú ngay trên ô checkbox I**, không thêm cột trạng thái. Di chuột lên ô để xem. Checkbox tự bỏ tích sau khi xử lý.

## Cài đặt / cập nhật từ bản cũ

1. Thay mã phiếu thu trong Apps Script bằng toàn bộ `code.gs` trong thư mục này. Không dán thêm bản thứ hai và không đưa `test_phieu_thu.cjs` vào Apps Script.
2. Kiểm tra mẫu đã có các biến ở mục dưới và các tiêu đề A:I đúng như bảng.
3. Tài khoản quản lý có quyền sửa Sheet, đọc mẫu và tạo tệp trong thư mục đích chạy **ptCaiDat** một lần, cấp các quyền Google yêu cầu.
4. Trở về `TH_PHIEU_THU`, nhập dữ liệu A:D rồi tích checkbox **Tạo_Word** tại I.

`ptCaiDat` đặt checkbox tại cột Tạo_Word, sử dụng ba cột Link_Word_1/2/3 đã có và không reset link hay checkbox hiện có. Chỉ thêm cột cần thiết nếu thiếu tiêu đề. Những cột do bản cũ thêm (`TẠO PHIẾU THU`, `Link_Docs_Phieu_Thu`, `Link_PDF_Phieu_Thu`, `Trạng thái tạo phiếu thu`) sẽ được xóa nếu dữ liệu bên dưới trống hoàn toàn; riêng checkbox cũ được coi là trống khi chưa tích. Nếu có dữ liệu, link hoặc công thức, cột cũ được giữ lại để tra cứu. Cột checkbox cũ được gỡ quy tắc checkbox và không còn được xử lý; chỉ dùng Tạo_Word.

Chỉ một tài khoản và một dự án quản lý trigger phiếu thu. Chạy lại cài đặt sẽ thay trigger ptXuLySua/ptMoMenu của tài khoản hiện tại, không thay trigger chứng từ khác. Không cài thêm simple trigger onEdit để gọi ptXuLySua vì sẽ xử lý trùng.

Không cần triển khai Web App. Dự án độc lập vẫn dùng được checkbox và tự cập nhật bằng chữ. Nếu mã nằm trong dự án mở từ **Sheet → Tiện ích mở rộng → Apps Script**, tải lại Sheet sẽ có thêm menu **Phiếu thu**. Menu chỉ dùng với dự án gắn với Sheet. Nếu toast không được hỗ trợ, thông báo được ghi vào nhật ký thực thi, không làm thất bại cài đặt/tạo phiếu.

## Mẫu Google Docs

Đã cấu hình các ID theo liên kết bạn gửi:

- Sheet: `1ShgTg3vFVaQbIE3R1ziQNQrwMMeiKxG_4_2nCXBmeBQ`, tab `TH_PHIEU_THU`.
- Mẫu: `1mn7hcT2Z4jxynPZSPvAI2F8WuEF_hHBcxnrNxr79U4I`.
- Thư mục lưu: `14tXO4OE6HmjVfshVK0sVYE2MfxhLshxm`.

| Vị trí trong mẫu | Biến |
|---|---|
| Ngày, tháng, năm lập và ngày ký nếu cùng ngày lập | `{{NGAY}}`, `{{THANG}}`, `{{NAM}}` |
| Người nộp tiền | `{{NGUOI_NOP}}` |
| Địa chỉ | `{{DIA_CHI}}` |
| Số tiền bằng số | `{{SO_TIEN}}` |
| Viết bằng chữ và Đã nhận đủ số tiền bằng chữ | `{{SO_TIEN_BANG_CHU}}` (đặt ở cả hai vị trí) |

Không dùng `{{SO_PHIEU}}`; mục Số để trống hoặc giữ dấu chấm để điền tay. Lý do nộp, chứng từ kèm theo, Nợ/Có, quyển số, thông tin công ty và chữ ký giữ theo mẫu. Gõ liền mỗi biến và giữ cùng định dạng chữ bên trong biến. Mã xử lý nội dung, bảng, đầu/chân trang của các thẻ tài liệu. Mẫu gốc không bị sửa; mã không kiểm định nội dung nghiệp vụ hoặc căn cứ pháp lý trên mẫu.

## Số tiền, ngày và xử lý lỗi

Số tiền hỗ trợ số nguyên dương đến `999.999.999.999.999` đồng. Chấp nhận số trong ô hoặc chuỗi `1250000`, `1.250.000`, `1,250,000`, `1 250 000`. Ví dụ `1.250.000` thành **Một triệu hai trăm năm mươi nghìn đồng**. Không hỗ trợ số âm, tiền lẻ hoặc chuỗi kèm đơn vị tiền.

Khi nhập/sửa/dán nhiều dòng tại D, E được cập nhật. Xóa D sẽ xóa E; tiền không hợp lệ làm trống E và ghi chú lỗi tại E. Trigger chỉ chạy cho sửa trực tiếp của người dùng; thay đổi từ công thức/API/script khác không kích hoạt onEdit. Khi tạo phiếu, E luôn được tính lại từ giá trị D để khớp với tệp xuất.

Ngày lập là ô ngày Google Sheets hoặc văn bản dd/MM/yyyy. Khi tạo phải đủ ngày, họ tên, địa chỉ và số tiền hợp lệ. Tên tệp có ngày lập, họ tên, dòng, lần tạo và thời gian. Tệp lưu trực tiếp trong thư mục đích, theo quyền thư mục; không bật chia sẻ công khai.

Nếu lỗi sau khi tạo bản sao, link vẫn được lưu tại F/G/H, kèm ghi chú bản sao đang điền dữ liệu. Kiểm tra bản sao đó trước khi tạo lại; một bản sao lỗi vẫn chiếm một ô link để không làm mất dấu tệp. Không tự xóa tệp khi lỗi.

Tối đa 5 dòng tạo trong mỗi thao tác tích/dán. Các dòng vượt giới hạn được bỏ tích và có ghi chú yêu cầu tích lại. Nếu đang xử lý và không lấy được khóa, bỏ tích rồi tích lại sau khi tác vụ trước hoàn thành. Không sắp xếp, di chuyển hoặc xóa dòng đang tạo. Không tự xóa link cũ nếu muốn giữ lịch sử và giới hạn 3 lần.

Kiểm tra cục bộ: `node Code_Phieu_Thu/test_phieu_thu.cjs`. Chưa triển khai mã lên Google hoặc kiểm tra trực tiếp bản Docs sau khi xuất; sau cập nhật hãy thử một dòng và mở link F để kiểm tra bố cục.
