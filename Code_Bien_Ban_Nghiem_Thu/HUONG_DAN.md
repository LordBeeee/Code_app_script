# Cài đặt tự động tạo Biên bản nghiệm thu

## 1. Đặt biến trong Google Docs mẫu

Các vị trí màu xanh lá được dùng lặp lại nhiều lần:

- Số biên bản ở đầu tài liệu: `{{SO_BBNT}}/{{NAM_KY}}/HĐKV-GETVINI`
- Tất cả số hợp đồng trong nội dung: `{{SO_HOP_DONG}}/{{NAM_KY}}/HĐKV-GETVINI`
- Tất cả ngày ký hợp đồng: `{{KY_NGAY}}`

`{{SO_BBNT}}` và `{{SO_HOP_DONG}}` đều nhận cùng giá trị từ cột B. Code tự ghi giá trị đó vào cột A, vì vậy **Số Biên Bản Nghiệm Thu = Số Hợp đồng khoán**.

Các phần màu vàng dùng biến sau:

| Nội dung | Biến |
|---|---|
| Họ và tên | `{{HO_TEN}}` |
| Ngày lập biên bản dạng đầy đủ | `{{NGAY_LAM_DON}}` |
| Ngày, tháng, năm lập biên bản viết tách | `{{NGAY}}`, `{{THANG}}`, `{{NAM}}` |
| Số CCCD | `{{SO_CCCD}}` |
| Mã số thuế | `{{MA_SO_THUE}}` |
| Địa chỉ | `{{DIA_CHI}}` |
| Số điện thoại | `{{SO_DIEN_THOAI}}` |
| Nội dung kịch bản | `{{NOI_DUNG_KICH_BAN}}` |
| Số lượng kịch bản | `{{SO_LUONG_KICH_BAN}}` |
| Số lượng video | `{{SO_LUONG_VIDEO}}` |
| Khung video | `{{KHUNG_VIDEO}}` |
| Chất lượng video | `{{CHAT_LUONG_VIDEO}}` |
| Thời lượng video | `{{THOI_LUONG_VIDEO}}` |
| Hiệu ứng video | `{{HIEU_UNG_VIDEO}}` |
| Nội dung diễn viên | `{{NOI_DUNG_DIEN_VIEN}}` |
| Số lượng diễn viên | `{{SO_LUONG_DIEN_VIEN}}` |

Giữ cùng một kiểu định dạng cho toàn bộ ký tự của mỗi biến; không đổi font hoặc màu ở giữa biến.

## 2. Cấu trúc Sheet

- A: Số Biên Bản Nghiệm Thu, code tự ghi bằng cột B.
- B:R: dữ liệu bắt buộc phải nhập đầy đủ.
- S:U: Link Word lần 01, 02, 03.
- V: checkbox TẠO WORD.

## 3. Cài đặt

1. Thêm một tệp `.gs` mới vào dự án Apps Script hiện tại, hoặc dùng một dự án Apps Script độc lập.
2. Dán toàn bộ nội dung `code.gs` vào tệp đó.
3. Lưu và chạy hàm `caiDatBienBanNghiemThu` một lần.
4. Chấp nhận quyền truy cập Google Docs, Sheets và Drive.
5. Tải lại Google Sheet.

Script dùng `SPREADSHEET_URL` để mở đúng Google Sheet nên chạy được cả trong dự án Apps Script độc lập. Tên hàm riêng cũng cho phép đặt cùng dự án với code hợp đồng và giấy đề nghị thanh toán.

## 4. Sử dụng

Khi chạy cài đặt, code tự xóa hai cột cũ **J – Tên dự án** và **K – Số lượng Video (Công việc giao khoán)** nếu chúng vẫn còn.

Điền đủ B:R rồi tích checkbox ở V. Thiếu trường, ngày sai định dạng hoặc số lượng không lớn hơn 0 thì biên bản không được tạo. Mỗi dòng được tạo tối đa ba lần và link được ghi lần lượt vào S, T, U.

Nếu mẫu còn câu dùng số lượng công việc giao khoán, đổi biến ở đó thành `{{SO_LUONG_VIDEO}}`. Xóa mọi vị trí `{{TEN_DU_AN}}` khỏi mẫu.
