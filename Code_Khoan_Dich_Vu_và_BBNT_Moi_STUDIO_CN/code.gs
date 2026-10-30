/**
 * TỰ ĐỘNG TẠO HỢP ĐỒNG KHOÁN STUDIO CN
 *
 * A:N - Dữ liệu hợp đồng (M là Số Tài khoản, N là Mở Tại)
 * O:R - Link Word và nút tạo hợp đồng khoán
 * S:T - FILE CỨNG hợp đồng khoán
 * U   - Đánh giá BBNT
 * V:Y - Link Word và nút tạo BBNT
 * Z:AA - FILE CỨNG BBNT
 * AB:AD - Trạng thái, Thuế TNCN, Ngày Thanh toán
 *
 * Đặt các biến {{...}} trong Google Docs mẫu theo TEMPLATE_FIELDS ở cuối file.
 * Riêng các số mẫu ở mục 3.2–3.3 được đổi thành biến trên bản sao hợp đồng.
 */

const CONFIG = {
  SHEET_NAME: 'HỢP ĐỒNG KHOÁN và BBNT STUDIO CN',
  LEGACY_BBNT_SHEET_NAME: 'BBNT STUDIO CN',
  SPREADSHEET_ID: '1uFp7McsEUjFQzMsQrW0ZSiqxQe7324Wy6KMzsHP_Mq0',

  TEMPLATE_URL:
    'https://docs.google.com/document/d/1FJusbfQnw4a6dzH-mVq5JB2JK4fAbpozqaQ11gHzTds/edit?tab=t.0',
  BBNT_TEMPLATE_URL:
    'https://docs.google.com/document/d/18i8Y3PeB8wFEzbC0Yapw_UFVgg_3yJR04I7S7zDaq00/edit?tab=t.0',

  // Thư mục "HỢP ĐỒNG KHOÁN DỊCH VỤ" trong ảnh.
  // Nếu Drive báo không tìm thấy thư mục, hãy dán lại ID thư mục vào đây.
  FOLDER_ID: '1FCIPy1gl88pMOlUnn4ZtgPIPYflzzmNg',
  BBNT_FOLDER_ID: '1sP4YJ8EVSNlq2D22M_J5ST2LXDjZ9ZKv',

  // Để trống để dùng URL Web App của chính dự án STUDIO CN sau khi triển khai.
  WEB_APP_URL: 'https://script.google.com/a/macros/getvini.com/s/AKfycbypX-CAZ3ykjnUJ62SLbUh0Eq5Li9OBeQm-ExZ_DJP9PwMDNw7_Bq8FjwzWDzo5gIUC/exec',
  CENTRAL_NUMBER_URL: 'https://script.google.com/a/macros/getvini.com/s/AKfycbwkG4U2MgM9aywiEcV7yKikOko-Wh0VRqE9OXGg3QC292L9AAxZ3u1mzb9UVrbDERO2hg/exec',

  FIRST_LINK_COLUMN: 15, // O
  LINK_COUNT: 3,         // O, P, Q
  ACTION_COLUMN: 18,     // R
  HARD_UPLOAD_COLUMN: 19, // S
  HARD_LINK_COLUMN: 20,   // T
  BBNT_EVALUATION_COLUMN: 21, // U
  BBNT_FIRST_LINK_COLUMN: 22, // V
  BBNT_ACTION_COLUMN: 25, // Y
  BBNT_HARD_UPLOAD_COLUMN: 26, // Z
  BBNT_HARD_LINK_COLUMN: 27, // AA
  LINK_PROTECTION_DESCRIPTION: 'KHOA_LINK_HOP_DONG_TU_DONG',
  BBNT_LINK_PROTECTION_DESCRIPTION: 'KHOA_LINK_BBNT_GOP_STUDIO_CN',
  HARD_FILE_PROTECTION_DESCRIPTION: 'KHOA_FILE_CUNG_HOP_DONG_TU_DONG',
  BBNT_HARD_FILE_PROTECTION_DESCRIPTION: 'KHOA_FILE_CUNG_BBNT_GOP_STUDIO_CN',
  AMOUNT_COLUMN: 11,       // K: Số tiền thực nhận sau thuế
  AMOUNT_WORDS_COLUMN: 12, // L: Số tiền thực nhận sau thuế bằng chữ (tự điền)
  UPLOAD_MAX_SIZE_MB: 10
};


// ========================================
// 0. TRANG WEB TẢI FILE CỨNG
// Tái sử dụng cùng cơ chế với Đề nghị thanh toán.
// ========================================

function doGet(e) {
  const loai = String(e && e.parameter && e.parameter.loai || '');
  const row = Number(e && e.parameter && e.parameter.row || 0);

  if (!['FILE_CUNG', 'FILE_CUNG_KHOAN', 'FILE_CUNG_BBNT'].includes(loai) ||
      !Number.isInteger(row) || row < 2) {
    return HtmlService.createHtmlOutput(
      '<h3>Đường link upload không hợp lệ.</h3>'
    );
  }

  const ss = laySpreadsheet_();
  const activeEmail = layEmailNguoiDung_();
  const ownerEmail = layEmailChuSoHuu_(ss);

  if (activeEmail !== ownerEmail) {
    return HtmlService.createHtmlOutput(
      '<h3>Bạn không có quyền tải FILE CỨNG.</h3>' +
      '<p>Chỉ chủ sở hữu bảng tính được thực hiện thao tác này.</p>'
    );
  }

  const template = HtmlService.createTemplateFromFile('UploadPDF');
  template.loai = loai;
  template.row = row;
  template.tieuDe = loai === 'FILE_CUNG_BBNT' ? 'FILE CỨNG BBNT' : 'FILE CỨNG HỢP ĐỒNG';
  template.maxSizeMb = CONFIG.UPLOAD_MAX_SIZE_MB;
  template.webMode = true;

  return template.evaluate()
    .setTitle('Tải tệp - FILE CỨNG - Dòng ' + row);
}


// ========================================
// 1. MENU VÀ CÀI ĐẶT HỆ THỐNG
// ========================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tạo hợp đồng')
    .addItem('Cài đặt hệ thống', 'caiDatHeThong')
    .addItem('Cài kết nối cấp số chung', 'caiDatKetNoiCapSoChung')
    .addItem('Kiểm tra kết nối cấp số chung', 'kiemTraKetNoiCapSoChung')
    .addItem('Tạo hợp đồng cho dòng đang chọn', 'taoHopDongDongDangChon')
    .addItem('Tạo BBNT cho dòng đang chọn', 'taoBbntDongDangChon')
    .addSeparator()
    .addItem('Chuyển dữ liệu BBNT STUDIO CN cũ', 'chuyenDuLieuBbntCu')
    .addSeparator()
    .addItem('Tải FILE CỨNG hợp đồng', 'moTaiFileCung')
    .addItem('Tải FILE CỨNG BBNT', 'moTaiFileCungBbnt')
    .addToUi();
}

/**
 * Chạy hàm này đúng một lần sau khi dán code.
 * Kiểm tra bố cục A:AD, cài hai nút Word và hai nút FILE CỨNG.
 */
function caiDatHeThong() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error('Không tìm thấy tab: ' + CONFIG.SHEET_NAME);
  }
  kiemTraBoCucHopDong_(sheet);

  const templateId = layIdGoogleDoc_(CONFIG.TEMPLATE_URL);
  const template = DriveApp.getFileById(templateId);
  const bbntTemplate = DriveApp.getFileById(layIdGoogleDoc_(CONFIG.BBNT_TEMPLATE_URL));

  if (template.getMimeType() !== MimeType.GOOGLE_DOCS ||
      bbntTemplate.getMimeType() !== MimeType.GOOGLE_DOCS) {
    throw new Error('Hai file mẫu phải là Google Docs.');
  }

  // Kiểm tra quyền truy cập trước khi tạo trigger.
  DocumentApp.openById(templateId).getBody();
  DocumentApp.openById(bbntTemplate.getId()).getBody();
  DriveApp.getFolderById(CONFIG.FOLDER_ID).getName();
  DriveApp.getFolderById(CONFIG.BBNT_FOLDER_ID).getName();

  PropertiesService.getScriptProperties().setProperty(
    'SPREADSHEET_ID',
    ss.getId()
  );

  xoaCongThucSttCu_(sheet);

  const maxRows = sheet.getMaxRows();

  if (maxRows >= 2) {
    const checkboxRange = sheet.getRange(
      2,
      CONFIG.ACTION_COLUMN,
      maxRows - 1,
      1
    );

    checkboxRange
      .insertCheckboxes()
      .setHorizontalAlignment('center')
      .setBackground('#E8F5E9');

    sheet.getRange(2, CONFIG.BBNT_ACTION_COLUMN, maxRows - 1, 1)
      .insertCheckboxes()
      .setHorizontalAlignment('center')
      .setBackground('#E8F5E9');

    khoaCotLink_(sheet, maxRows);
    taoNutUploadPdf_(sheet, maxRows);
  }

  sheet.setColumnWidths(CONFIG.FIRST_LINK_COLUMN, CONFIG.LINK_COUNT, 180);
  sheet.setColumnWidth(CONFIG.ACTION_COLUMN, 120);
  sheet.setColumnWidths(CONFIG.BBNT_FIRST_LINK_COLUMN, CONFIG.LINK_COUNT, 180);
  sheet.setColumnWidth(CONFIG.BBNT_ACTION_COLUMN, 120);
  sheet.setColumnWidth(CONFIG.HARD_UPLOAD_COLUMN, 180);
  sheet.setColumnWidth(CONFIG.BBNT_HARD_UPLOAD_COLUMN, 180);

  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'xuLyTaoHopDong')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('xuLyTaoHopDong')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  thongBaoHopDong_(
    ss,
    coLinkWebAppHopLe_()
      ? 'Đã tạo nút Word R/Y và FILE CỨNG S/Z.'
      : 'Đã cài Word. Hãy triển khai Web App và điền WEB_APP_URL để tạo nút S/Z.',
    'Cài đặt thành công',
    8
  );
}

/** Khi chạy từ Apps Script editor, Spreadsheet không hiển thị toast được. */
function thongBaoHopDong_(spreadsheet, message, title, seconds) {
  try {
    spreadsheet.toast(message, title, seconds);
  } catch (error) {
    console.log(title + ': ' + message);
  }
}

function coLinkWebAppHopLe_() {
  return /^https:\/\/script\.google\.com\/(?:macros\/s|a\/macros\/[^/]+\/s)\/[^/?#]+\/exec(?:[?#].*)?$/
    .test(layUrlWebApp_());
}

function layUrlWebApp_() {
  return String(CONFIG.WEB_APP_URL || ScriptApp.getService().getUrl() || '').trim();
}

const MERGED_HEADERS = [
  'Số hợp đồng', 'Họ và tên người đề nghị', 'Ngày làm đơn', 'số CCCD',
  'Mã số thuế', 'Địa chỉ', 'Số điện thoại', 'Tên dự án',
  'Nội dung công việc', 'Ngày hoàn thành công việc',
  'Số tiền thực nhận sau thuế', 'Số tiền thực nhận sau thuế(Bằng chữ)',
  'Số Tài khoản', 'Mở Tại',
  'Link_Word_01_Khoan', 'Link_Word_02_Khoan', 'Link_Word_03_Khoan',
  'TẠO WORD_Khoan', 'FILE CỨNG_Khoan', 'Link_FILE_Cung_Khoan',
  'Đánh giá', 'Link_Word_01_BBNT', 'Link_Word_02_BBNT',
  'Link_Word_03_BBNT', 'TẠO WORD_BBNT', 'FILE CỨNG_BBNT',
  'Link_FILE_Cung_BBNT', 'Trạng Thái', 'Thuế TNCN', 'Ngày Thanh toán'
];

function chuanHoaTieuDe_(value) {
  return String(value || '').replace(/\u00a0|&#x20;/gi, ' ').trim()
    .replace(/\s+/g, ' ').replace(/\s*\(/g, '(')
    .toLocaleLowerCase('vi-VN');
}

function coBoCucHopDongMoi_(sheet) {
  const headers = sheet.getRange(1, 1, 1, MERGED_HEADERS.length)
    .getDisplayValues()[0];
  return MERGED_HEADERS.every((expected, index) =>
    chuanHoaTieuDe_(headers[index]) === chuanHoaTieuDe_(expected)
  );
}

function kiemTraBoCucHopDong_(sheet) {
  if (!coBoCucHopDongMoi_(sheet)) {
    throw new Error('Bố cục tab gộp STUDIO CN A:AD chưa đúng với danh sách cột đã thống nhất.');
  }
}

/** Hai nút FILE CỨNG trên cùng dòng, dùng URL Web App của STUDIO CN. */
function taoNutUploadPdf_(sheet, maxRows) {
  if (!coLinkWebAppHopLe_() || maxRows < 2) return;

  const baseUrl = layUrlWebApp_();
  [
    { column: CONFIG.HARD_UPLOAD_COLUMN, type: 'FILE_CUNG_KHOAN' },
    { column: CONFIG.BBNT_HARD_UPLOAD_COLUMN, type: 'FILE_CUNG_BBNT' }
  ].forEach(item => {
    const buttons = [];
    for (let row = 2; row <= maxRows; row++) {
      buttons.push([SpreadsheetApp.newRichTextValue()
        .setText('⬆ TẢI FILE CỨNG')
        .setLinkUrl(baseUrl + '?loai=' + item.type + '&row=' + row)
        .build()]);
    }
    sheet.getRange(2, item.column, maxRows - 1, 1)
      .setRichTextValues(buttons)
      .setHorizontalAlignment('center')
      .setFontWeight('bold')
      .setFontColor('#1155CC')
      .setBackground('#E8F0FE');
  });
}

function moTaiFileCung() {
  moHopThoaiTaiPdf_('FILE_CUNG_KHOAN');
}

function moTaiFileCungBbnt() {
  moHopThoaiTaiPdf_('FILE_CUNG_BBNT');
}

/** Dùng lại hộp thoại upload của Đề nghị thanh toán. */
function moHopThoaiTaiPdf_(loai) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  if (layEmailNguoiDung_() !== layEmailChuSoHuu_(ss)) {
    ui.alert('Chỉ chủ sở hữu bảng tính được tải FILE CỨNG.');
    return;
  }

  if (sheet.getName() !== CONFIG.SHEET_NAME) {
    ui.alert('Hãy mở tab ' + CONFIG.SHEET_NAME + ' trước.');
    return;
  }

  const activeRange = sheet.getActiveRange();
  const row = activeRange ? activeRange.getRow() : 0;

  if (row < 2) {
    ui.alert('Hãy chọn một ô tại dòng dữ liệu cần tải tệp.');
    return;
  }

  const template = HtmlService.createTemplateFromFile('UploadPDF');
  template.loai = loai;
  template.row = row;
  template.tieuDe = loai === 'FILE_CUNG_BBNT' ? 'FILE CỨNG BBNT' : 'FILE CỨNG HỢP ĐỒNG';
  template.maxSizeMb = CONFIG.UPLOAD_MAX_SIZE_MB;
  template.webMode = false;

  const html = template.evaluate()
    .setWidth(440)
    .setHeight(330);

  ui.showModalDialog(html, 'Tải FILE CỨNG - Dòng ' + row);
}

/**
 * Hàm được gọi trực tiếp từ UploadPDF.html.
 * FILE CỨNG được lưu chung thư mục với các bản hợp đồng của dòng đó.
 */
function luuPdfTuMayTinh(payload) {
  if (!payload || !payload.base64) {
    throw new Error('Không nhận được dữ liệu tệp.');
  }

  const row = Number(payload.row);
  const loai = String(payload.loai || '');
  const originalName = String(payload.fileName || '').trim();
  const mimeType = String(payload.mimeType || '').toLowerCase();

  if (!Number.isInteger(row) || row < 2) {
    throw new Error('Dòng tải tệp không hợp lệ.');
  }

  if (!['FILE_CUNG', 'FILE_CUNG_KHOAN', 'FILE_CUNG_BBNT'].includes(loai)) {
    throw new Error('Loại tệp không hợp lệ.');
  }

  const allowedMimeTypes = [
    'application/pdf',
    'image/jpeg',
    'image/png'
  ];
  const allowedExtensions = ['.pdf', '.jpg', '.jpeg', '.png'];
  const lowerName = originalName.toLowerCase();

  if (
    !allowedMimeTypes.includes(mimeType) ||
    !allowedExtensions.some(extension => lowerName.endsWith(extension))
  ) {
    throw new Error('Chỉ được tải lên PDF, JPG hoặc PNG.');
  }

  const bytes = Utilities.base64Decode(payload.base64);
  const maxBytes = CONFIG.UPLOAD_MAX_SIZE_MB * 1024 * 1024;

  if (bytes.length > maxBytes) {
    throw new Error(
      'Tệp vượt quá ' + CONFIG.UPLOAD_MAX_SIZE_MB + ' MB.'
    );
  }

  const ss = laySpreadsheet_();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error('Không tìm thấy tab: ' + CONFIG.SHEET_NAME);
  }

  if (layEmailNguoiDung_() !== layEmailChuSoHuu_(ss)) {
    throw new Error('Chỉ chủ sở hữu bảng tính được tải FILE CỨNG.');
  }

  kiemTraBoCucHopDong_(sheet);
  const isBbnt = loai === 'FILE_CUNG_BBNT';
  const hardLinkColumn = isBbnt
    ? CONFIG.BBNT_HARD_LINK_COLUMN : CONFIG.HARD_LINK_COLUMN;
  const currentLink = String(
    sheet.getRange(row, hardLinkColumn).getDisplayValue() || ''
  ).trim();

  if (currentLink) {
    throw new Error('Dòng này đã có FILE CỨNG.');
  }

  const soHopDong = String(
    sheet.getRange(row, 1).getDisplayValue() || ''
  ).trim();

  if (!soHopDong) {
    throw new Error('Dòng ' + row + ' chưa có Số hợp đồng ở cột A.');
  }

  const mainFolder = DriveApp.getFolderById(
    isBbnt ? CONFIG.BBNT_FOLDER_ID : CONFIG.FOLDER_ID
  );
  const targetFolder = layHoacTaoThuMucStt_(mainFolder, soHopDong);
  const safeName = lamSachTenFile_(originalName);
  const blob = Utilities.newBlob(bytes, mimeType, safeName);
  const file = targetFolder.createFile(blob);

  file.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  const viewUrl =
    'https://drive.google.com/file/d/' + file.getId() + '/view';

  sheet.getRange(row, isBbnt ? CONFIG.BBNT_HARD_UPLOAD_COLUMN : CONFIG.HARD_UPLOAD_COLUMN)
    .setNote('Đã tải: ' + safeName);
  sheet.getRange(row, hardLinkColumn).setValue(viewUrl);
  SpreadsheetApp.flush();

  return {
    fileName: safeName,
    url: viewUrl,
    row: row,
    lanTai: '01'
  };
}

function laySpreadsheet_() {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (activeSpreadsheet) return activeSpreadsheet;

  const spreadsheetId = PropertiesService.getScriptProperties()
    .getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error(
      'Chưa lưu Spreadsheet ID. Chủ sở hữu hãy chạy caiDatHeThong một lần.'
    );
  }

  return SpreadsheetApp.openById(spreadsheetId);
}

function layEmailNguoiDung_() {
  const email = String(
    Session.getActiveUser().getEmail() || ''
  ).trim().toLowerCase();

  if (!email) {
    throw new Error(
      'Google không cung cấp email đăng nhập. ' +
      'Hãy dùng tài khoản Google Workspace cùng công ty và cấp quyền cho script.'
    );
  }

  return email;
}

function layEmailChuSoHuu_(spreadsheet) {
  const owner = DriveApp.getFileById(spreadsheet.getId()).getOwner();
  return owner
    ? String(owner.getEmail() || '').trim().toLowerCase()
    : '';
}


// ========================================
// 2. TỰ ĐIỀN CỘT L VÀ XỬ LÝ CHECKBOX Ở CỘT R
// ========================================

/** Nhận số Sheets hoặc số tiền nhập dạng 2,160,000 / 2.160.000. */
function docSoTien_(value) {
  if (typeof value === 'number') {
    return Number.isSafeInteger(value) && value > 0 ? value : null;
  }

  const input = String(value ?? '').trim().replace(/\u00a0/g, ' ');
  if (!input) return null;

  const valid = /^\d+$/.test(input) ||
    /^\d{1,3}(?:,\d{3})+$/.test(input) ||
    /^\d{1,3}(?:\.\d{3})+$/.test(input) ||
    /^\d{1,3}(?: \d{3})+$/.test(input);
  if (!valid) return null;

  const amount = Number(input.replace(/[., ]/g, ''));
  return Number.isSafeInteger(amount) && amount > 0 ? amount : null;
}

function vietBaChuSo_(number, docDuBaChuSo) {
  const digits = ['không', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín'];
  const tram = Math.floor(number / 100);
  const chuc = Math.floor(number / 10) % 10;
  const donVi = number % 10;
  const words = [];

  if (tram || docDuBaChuSo) words.push(digits[tram], 'trăm');
  if (chuc > 1) {
    words.push(digits[chuc], 'mươi');
    if (donVi) words.push(donVi === 1 ? 'mốt' : donVi === 5 ? 'lăm' : digits[donVi]);
  } else if (chuc === 1) {
    words.push('mười');
    if (donVi) words.push(donVi === 5 ? 'lăm' : digits[donVi]);
  } else if (donVi) {
    if (tram || docDuBaChuSo) words.push('linh');
    words.push(digits[donVi]);
  }

  return words.join(' ');
}

/** Viết số tiền đồng bằng chữ, không thêm "đồng" vì mẫu đã có từ này. */
function vietSoTienBangChu_(amount) {
  const units = ['', 'nghìn', 'triệu', 'tỷ', 'nghìn tỷ'];
  const groups = [];
  let remaining = amount;

  while (remaining > 0) {
    groups.push(remaining % 1000);
    remaining = Math.floor(remaining / 1000);
  }
  if (groups.length > units.length) {
    throw new Error('Số tiền ở cột K quá lớn để chuyển thành chữ.');
  }

  const words = [];
  for (let i = groups.length - 1; i >= 0; i--) {
    if (!groups[i]) continue;
    const groupWords = vietBaChuSo_(groups[i], words.length > 0 && groups[i] < 100);
    words.push(groupWords + (units[i] ? ' ' + units[i] : ''));
  }

  const result = words.join(' ');
  return result.charAt(0).toUpperCase() + result.slice(1);
}

function capNhatTienThucNhanBangChu_(sheet, row) {
  const amount = docSoTien_(sheet.getRange(row, CONFIG.AMOUNT_COLUMN).getValue());
  sheet.getRange(row, CONFIG.AMOUNT_WORDS_COLUMN)
    .setValue(amount === null ? '' : vietSoTienBangChu_(amount));
}

/** Tính thù lao trước thuế từ số thực nhận, làm tròn đến đồng. */
function tinhThuLaoTruocThue_(soTienThucNhan) {
  const tongThuLao = Math.round(soTienThucNhan / 0.9);
  if (!Number.isSafeInteger(tongThuLao)) {
    throw new Error('Số tiền thực nhận ở cột K quá lớn để tính thù lao trước thuế.');
  }
  return {
    tongThuLao: tongThuLao,
    thueTncn: tongThuLao - soTienThucNhan
  };
}

function dinhDangSoTien_(amount) {
  return String(amount).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

function xuLyTaoHopDong(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();

  if (sheet.getName() !== CONFIG.SHEET_NAME) return;
  if (!coBoCucHopDongMoi_(sheet)) {
    if ([CONFIG.ACTION_COLUMN, CONFIG.BBNT_ACTION_COLUMN].includes(range.getColumn()) &&
        e.value === 'TRUE') {
      range.setValue(false);
      range.setNote('Bố cục A:AD chưa đúng. Hãy kiểm tra tiêu đề cột trước khi tạo Word.');
    }
    return;
  }

  // Hỗ trợ cả khi dán nhiều dòng hoặc nhiều cột có chứa K.
  if (
    range.getColumn() <= CONFIG.AMOUNT_COLUMN &&
    range.getLastColumn() >= CONFIG.AMOUNT_COLUMN
  ) {
    for (let row = Math.max(2, range.getRow()); row <= range.getLastRow(); row++) {
      capNhatTienThucNhanBangChu_(sheet, row);
    }
  }

  if (
    ![CONFIG.ACTION_COLUMN, CONFIG.BBNT_ACTION_COLUMN].includes(range.getColumn()) ||
    range.getRow() < 2 ||
    range.getNumRows() !== 1 ||
    range.getNumColumns() !== 1 ||
    e.value !== 'TRUE'
  ) {
    return;
  }

  const actionCell = range;
  const row = range.getRow();
  const isBbnt = range.getColumn() === CONFIG.BBNT_ACTION_COLUMN;
  const title = isBbnt ? 'BBNT' : 'hợp đồng';
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();

  if (!lock.tryLock(15000)) {
    actionCell.setValue(false);
    actionCell.setNote('Hệ thống đang xử lý yêu cầu khác. Hãy thử lại.');
    return;
  }

  try {
    actionCell.setNote('Đang tạo ' + title + '...');
    const result = isBbnt ? taoBbntChoDong_(sheet, row) : taoHopDongChoDong_(sheet, row);

    actionCell.setNote(
      'Đã tạo ' + title + ' lần ' + result.lanTao + ': ' + result.url
    );

    thongBaoHopDong_(
      sheet.getParent(),
      'Đã tạo ' + title + ' lần ' + result.lanTao + ' cho dòng ' + row,
      'Thành công',
      6
    );
  } catch (error) {
    actionCell.setNote('Không thể tạo ' + title + ': ' + error.message);
    console.error('Dòng ' + row + ': ' + error.stack);

    thongBaoHopDong_(
      sheet.getParent(),
      error.message,
      'Không thể tạo ' + title + ' - Dòng ' + row,
      10
    );
  } finally {
    actionCell.setValue(false);
    lock.releaseLock();
  }
}

/** Tạo bằng menu cho dòng hiện đang chọn. */
function taoHopDongDongDangChon() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== CONFIG.SHEET_NAME) {
    throw new Error('Hãy mở tab ' + CONFIG.SHEET_NAME + '.');
  }

  const row = sheet.getActiveRange().getRow();

  if (row < 2) {
    throw new Error('Hãy chọn một ô trong dòng dữ liệu cần tạo hợp đồng.');
  }

  const lock = LockService.getDocumentLock() || LockService.getScriptLock();

  if (!lock.tryLock(15000)) {
    throw new Error('Hệ thống đang xử lý yêu cầu khác. Hãy thử lại.');
  }

  try {
    const result = taoHopDongChoDong_(sheet, row);

    thongBaoHopDong_(
      ss,
      'Đã tạo hợp đồng lần ' + result.lanTao + ' cho dòng ' + row,
      'Thành công',
      6
    );
  } finally {
    lock.releaseLock();
  }
}

function taoBbntDongDangChon() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  if (sheet.getName() !== CONFIG.SHEET_NAME) {
    throw new Error('Hãy mở tab ' + CONFIG.SHEET_NAME + '.');
  }
  const row = sheet.getActiveRange().getRow();
  if (row < 2) throw new Error('Hãy chọn một dòng dữ liệu để tạo BBNT.');
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Hệ thống đang xử lý yêu cầu khác.');
  try {
    const result = taoBbntChoDong_(sheet, row);
    thongBaoHopDong_(ss,
      'Đã tạo BBNT lần ' + result.lanTao + ' cho dòng ' + row,
      'Thành công', 6);
  } finally {
    lock.releaseLock();
  }
}


// ========================================
// 3. ĐỌC DỮ LIỆU VÀ TẠO HỢP ĐỒNG
// ========================================

function taoHopDongChoDong_(sheet, row) {
  kiemTraBoCucHopDong_(sheet);
  xoaCongThucSttCu_(sheet);

  // A:N là dữ liệu, O:Q là ba link đã tạo.
  const displayValues = sheet.getRange(row, 1, 1, CONFIG.FIRST_LINK_COLUMN + CONFIG.LINK_COUNT - 1)
    .getDisplayValues()[0];
  const rawValues = sheet.getRange(row, 1, 1, CONFIG.FIRST_LINK_COLUMN - 1).getValues()[0];

  const ngayLamDonHienThi = String(displayValues[2] || '').trim();

  if (!ngayLamDonHienThi) {
    throw new Error('Thiếu Ngày làm đơn ở cột C.');
  }

  const ngayKy = tachNgay_(ngayLamDonHienThi);

  const headers = [
    'Số hợp đồng',
    'Họ và tên người đề nghị',
    'Ngày làm đơn',
    'Số CCCD',
    'Mã số thuế',
    'Địa chỉ',
    'Số điện thoại',
    'Tên dự án',
    'Nội dung công việc',
    'Ngày hoàn thành công việc',
    'Số tiền thực nhận sau thuế',
    'Số tiền thực nhận sau thuế (Bằng chữ)',
    'Số Tài khoản',
    'Mở Tại'
  ];

  const errorValues = [
    '#ERROR!', '#REF!', '#VALUE!', '#N/A',
    '#DIV/0!', '#NUM!', '#NAME?'
  ];
  const missing = [];

  const soTienThucNhan = docSoTien_(rawValues[CONFIG.AMOUNT_COLUMN - 1]);
  if (soTienThucNhan === null) {
    throw new Error('Số tiền thực nhận ở cột K phải là số nguyên lớn hơn 0 (ví dụ: 6.750.000).');
  }
  const soTienThucNhanBangChu = vietSoTienBangChu_(soTienThucNhan);
  sheet.getRange(row, CONFIG.AMOUNT_WORDS_COLUMN).setValue(soTienThucNhanBangChu);
  displayValues[CONFIG.AMOUNT_WORDS_COLUMN - 1] = soTienThucNhanBangChu;
  const { tongThuLao, thueTncn } = tinhThuLaoTruocThue_(soTienThucNhan);

  // Cột A do code tự sinh; bắt buộc người dùng nhập đầy đủ B:N.
  for (let i = 1; i < headers.length; i++) {
    const value = String(displayValues[i] ?? '').trim();

    if (!value || errorValues.some(errorValue => value.startsWith(errorValue))) {
      missing.push(headers[i]);
    }
  }

  if (missing.length > 0) {
    throw new Error(
      'Chưa thể tạo hợp đồng. Thiếu hoặc lỗi dữ liệu: ' + missing.join(', ')
    );
  }

  const links = displayValues.slice(
    CONFIG.FIRST_LINK_COLUMN - 1,
    CONFIG.FIRST_LINK_COLUMN - 1 + CONFIG.LINK_COUNT
  );
  const emptyLinkIndex = links.findIndex(link => String(link).trim() === '');

  if (emptyLinkIndex === -1) {
    throw new Error('Dòng này đã tạo đủ 3 hợp đồng.');
  }

  // Kiểm tra ngày hoàn thành trước khi cấp số hợp đồng.
  const ngayHoanTat = tachNgay_(displayValues[9]);
  const soHopDong = taoSoHopDongMoi_(sheet, row);
  displayValues[0] = soHopDong;

  const [
    stt,
    hoTen,
    ngayLamDon,
    soCccd,
    maSoThue,
    diaChi,
    soDienThoai,
    tenDuAn,
    noiDungCongViec,
    ngayHoanThanh,
    soTienThucNhanHienThi,
    soTienThucNhanBangChuHienThi,
    soTaiKhoan,
    moTai
  ] = displayValues;

  const lanTao = String(emptyLinkIndex + 1).padStart(2, '0');
  const targetLinkColumn = CONFIG.FIRST_LINK_COLUMN + emptyLinkIndex;

  const templateId = layIdGoogleDoc_(CONFIG.TEMPLATE_URL);
  const template = DriveApp.getFileById(templateId);
  const mainFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  const sttFolder = layHoacTaoThuMucStt_(mainFolder, stt);
  const fileName = lamSachTenFile_(
    stt + ' - ' + hoTen + ' - HĐ khoán STUDIO CN - ' + lanTao
  );
  const copy = template.makeCopy(fileName, sttFolder);

  try {
    const doc = DocumentApp.openById(copy.getId());
    dienBienNganHangVaoMau_(doc);
    dienBienThuLaoVaoMau_(doc);
    const replacements = {
      '{{SO_HD}}': stt,
      '{{STT}}': stt,
      '{{HO_TEN}}': hoTen,
      '{{HO_TEN_IN_HOA}}': hoTen.toLocaleUpperCase('vi-VN'),
      '{{NGAY_LAM_DON}}': ngayLamDon,
      '{{NGAY}}': ngayKy.ngay,
      '{{THANG}}': ngayKy.thang,
      '{{NAM}}': ngayKy.nam,
      '{{SO_CCCD}}': soCccd,
      '{{MA_SO_THUE}}': maSoThue,
      '{{DIA_CHI}}': diaChi,
      '{{SO_DIEN_THOAI}}': soDienThoai,
      '{{TEN_BRAND}}': '', // Xóa biến Brand còn sót trong mẫu cũ.
      '{{TEN_DU_AN}}': tenDuAn,
      '{{CONG_VIEC_GIAO_KHOAN}}': noiDungCongViec,
      '{{NGAY_HOAN_THANH}}': ngayHoanThanh,
      '{{NGAY_HT}}': ngayHoanTat.ngay,
      '{{THANG_HT}}': ngayHoanTat.thang,
      '{{NAM_HT}}': ngayHoanTat.nam,
      '{{TONG_THU_LAO}}': dinhDangSoTien_(tongThuLao),
      '{{TONG_THU_LAO_BANG_CHU}}': vietSoTienBangChu_(tongThuLao),
      '{{THUE_TNCN}}': dinhDangSoTien_(thueTncn),
      '{{SO_TIEN_THUC_NHAN}}': dinhDangSoTien_(soTienThucNhan),
      '{{SO_TIEN_THUC_NHAN_BANG_CHU}}': soTienThucNhanBangChuHienThi,
      '{{SO_TAI_KHOAN}}': soTaiKhoan,
      '{{MO_TAI}}': moTai
    };

    thayBienTrongTaiLieu_(doc, replacements);
    doc.saveAndClose();

    const remaining = timBienConLai_(DocumentApp.openById(copy.getId()));

    if (remaining.length > 0) {
      throw new Error(
        'Tài liệu mẫu còn biến chưa được thay: ' + remaining.join(', ')
      );
    }

    copy.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );

    const viewUrl =
      'https://docs.google.com/document/d/' + copy.getId() + '/view';

    sheet.getRange(row, targetLinkColumn).setValue(viewUrl);
    SpreadsheetApp.flush();

    return {
      lanTao: lanTao,
      url: viewUrl,
      fileId: copy.getId()
    };
  } catch (error) {
    throw new Error(
      error.message +
      ' Bản sao đã tạo để kiểm tra: https://docs.google.com/document/d/' +
      copy.getId() + '/edit'
    );
  }
}

/** BBNT STUDIO CN dùng Ký ngày=C, Ngày làm đơn=J trên cùng dòng. */
function taoBbntChoDong_(sheet, row) {
  kiemTraBoCucHopDong_(sheet);
  const values = sheet.getRange(row, 1, 1, CONFIG.BBNT_HARD_LINK_COLUMN)
    .getDisplayValues()[0];
  const soHopDong = String(values[0] || '').trim();
  if (!soHopDong) throw new Error('Dòng này chưa có Số hợp đồng ở cột A.');

  const requiredColumns = [2, 3, 4, 5, 6, 7, 8, 9, 10, CONFIG.BBNT_EVALUATION_COLUMN];
  const missing = requiredColumns.filter(column => {
    const value = String(values[column - 1] || '').trim();
    return !value || laLoiSheet_(value);
  }).map(cotThanhChu_);
  if (missing.length) {
    throw new Error('Hãy điền đủ dữ liệu BBNT ở cột: ' + missing.join(', '));
  }

  const links = values.slice(CONFIG.BBNT_FIRST_LINK_COLUMN - 1,
    CONFIG.BBNT_FIRST_LINK_COLUMN - 1 + CONFIG.LINK_COUNT);
  const emptyIndex = links.findIndex(link => !String(link || '').trim());
  if (emptyIndex < 0) throw new Error('Dòng này đã tạo đủ 3 BBNT.');

  const kyNgay = String(values[2]).trim();
  const ngayLamDonBbnt = String(values[9]).trim();
  const kyParts = tachNgay_(kyNgay);
  const donParts = tachNgay_(ngayLamDonBbnt);
  const hoTen = String(values[1]).trim();
  const lanTao = String(emptyIndex + 1).padStart(2, '0');
  const targetColumn = CONFIG.BBNT_FIRST_LINK_COLUMN + emptyIndex;
  const template = DriveApp.getFileById(layIdGoogleDoc_(CONFIG.BBNT_TEMPLATE_URL));
  const folder = layHoacTaoThuMucStt_(
    DriveApp.getFolderById(CONFIG.BBNT_FOLDER_ID), soHopDong
  );
  const fileName = lamSachTenFile_(
    soHopDong + ' - ' + hoTen + ' - BBNT STUDIO CN - ' + lanTao
  );
  const copy = template.makeCopy(fileName, folder);

  try {
    const doc = DocumentApp.openById(copy.getId());
    const replacements = {
      '{{SO_BBNT}}': soHopDong,
      '{{SO_BIEN_BAN_NGHIEM_THU}}': soHopDong,
      '{{SO_HOP_DONG}}': soHopDong,
      '{{SO_HD}}': soHopDong,
      '{{KY_NGAY}}': kyNgay,
      '{{NGAY_KY}}': kyNgay,
      '{{NGAY_KY_NGAY}}': kyParts.ngay,
      '{{THANG_KY}}': kyParts.thang,
      '{{NAM_KY}}': kyParts.nam,
      '{{HO_TEN}}': hoTen,
      '{{HO_TEN_IN_HOA}}': hoTen.toLocaleUpperCase('vi-VN'),
      '{{NGAY_LAM_DON}}': ngayLamDonBbnt,
      '{{NGAY}}': donParts.ngay,
      '{{THANG}}': donParts.thang,
      '{{NAM}}': donParts.nam,
      '{{SO_CCCD}}': values[3],
      '{{MA_SO_THUE}}': values[4],
      '{{DIA_CHI}}': values[5],
      '{{SO_DIEN_THOAI}}': values[6],
      '{{TEN_DU_AN}}': values[7],
      '{{HANG_MUC_CONG_VIEC}}': values[8],
      '{{DANH_GIA}}': values[20]
    };
    thayBienTrongTaiLieu_(doc, replacements);
    doc.saveAndClose();
    const remaining = timBienConLai_(DocumentApp.openById(copy.getId()));
    if (remaining.length) {
      throw new Error('Tài liệu mẫu còn biến chưa được thay: ' + remaining.join(', '));
    }
    copy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const viewUrl = 'https://docs.google.com/document/d/' + copy.getId() + '/view';
    sheet.getRange(row, targetColumn).setValue(viewUrl);
    SpreadsheetApp.flush();
    return { lanTao: lanTao, url: viewUrl, fileId: copy.getId() };
  } catch (error) {
    throw new Error(error.message +
      ' Bản sao đã tạo để kiểm tra: https://docs.google.com/document/d/' +
      copy.getId() + '/edit');
  }
}

function laLoiSheet_(value) {
  return ['#ERROR!', '#REF!', '#VALUE!', '#N/A', '#DIV/0!', '#NUM!', '#NAME?']
    .some(error => String(value || '').startsWith(error));
}

function cotThanhChu_(column) {
  let letters = '';
  while (column > 0) {
    column--;
    letters = String.fromCharCode(65 + column % 26) + letters;
    column = Math.floor(column / 26);
  }
  return letters;
}

/** Chuyển đánh giá và các link từ tab BBNT STUDIO CN cũ theo số hợp đồng. */
function chuyenDuLieuBbntCu() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const merged = ss.getSheetByName(CONFIG.SHEET_NAME);
  const legacy = ss.getSheetByName(CONFIG.LEGACY_BBNT_SHEET_NAME);
  if (!merged || !legacy) {
    throw new Error('Cần cả tab gộp và tab BBNT STUDIO CN cũ để chuyển dữ liệu.');
  }
  kiemTraBoCucHopDong_(merged);
  if (chuanHoaTieuDe_(legacy.getRange('B1').getDisplayValue()) !==
        chuanHoaTieuDe_('Số Hợp đồng khoán') ||
      chuanHoaTieuDe_(legacy.getRange('M1').getDisplayValue()) !==
        chuanHoaTieuDe_('Link_Word_01') ||
      chuanHoaTieuDe_(legacy.getRange('R1').getDisplayValue()) !==
        chuanHoaTieuDe_('Link_FILE_Cung')) {
    throw new Error('Bố cục tab BBNT cũ chưa đúng ở cột B, M hoặc R.');
  }
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Hệ thống đang xử lý yêu cầu khác.');

  try {
    const oldLastRow = legacy.getLastRow();
    const newLastRow = merged.getLastRow();
    const oldRows = oldLastRow < 2 ? [] : legacy
      .getRange(2, 1, oldLastRow - 1, 18).getDisplayValues();
    const newRows = newLastRow < 2 ? [] : merged
      .getRange(2, 1, newLastRow - 1, CONFIG.BBNT_HARD_LINK_COLUMN)
      .getDisplayValues();
    const targetByCode = new Map();
    newRows.forEach((values, index) => {
      const code = String(values[0] || '').trim();
      if (!code) return;
      if (!targetByCode.has(code)) targetByCode.set(code, []);
      targetByCode.get(code).push(index + 2);
    });
    const oldCounts = new Map();
    oldRows.forEach(values => {
      const code = String(values[1] || '').trim();
      if (code) oldCounts.set(code, (oldCounts.get(code) || 0) + 1);
    });
    const fields = [
      { oldIndex: 11, newColumn: 21 }, // L -> U: Đánh giá
      { oldIndex: 12, newColumn: 22 }, // M -> V: Word BBNT 01
      { oldIndex: 13, newColumn: 23 }, // N -> W: Word BBNT 02
      { oldIndex: 14, newColumn: 24 }, // O -> X: Word BBNT 03
      { oldIndex: 17, newColumn: 27 }  // R -> AA: Link FILE CỨNG BBNT
    ];
    const issues = [];
    let transferredRows = 0;
    let transferredCells = 0;
    oldRows.forEach((oldValues, index) => {
      const code = String(oldValues[1] || '').trim();
      if (!code) return;
      const targets = targetByCode.get(code) || [];
      if (oldCounts.get(code) > 1 || targets.length !== 1) {
        issues.push(code + ' (BBNT cũ dòng ' + (index + 2) + ': mã trùng hoặc không có đúng một dòng hợp đồng)');
        return;
      }
      const targetRow = targets[0];
      const existing = newRows[targetRow - 2];
      const conflict = fields.some(field => {
        const source = String(oldValues[field.oldIndex] || '').trim();
        const current = String(existing[field.newColumn - 1] || '').trim();
        return source && current && source !== current;
      });
      if (conflict) {
        issues.push(code + ' (BBNT cũ dòng ' + (index + 2) + ': dữ liệu khác tab gộp)');
        return;
      }
      let count = 0;
      fields.forEach(field => {
        const source = String(oldValues[field.oldIndex] || '').trim();
        if (!source || String(existing[field.newColumn - 1] || '').trim()) return;
        merged.getRange(targetRow, field.newColumn).setValue(source);
        existing[field.newColumn - 1] = source;
        count++;
      });
      if (count) transferredRows++;
      transferredCells += count;
    });
    SpreadsheetApp.flush();
    const summary = 'Đã chuyển ' + transferredCells + ' ô của ' + transferredRows +
      ' hợp đồng. Cần kiểm tra ' + issues.length + ' dòng BBNT cũ.';
    console.log(summary + (issues.length ? '\n' + issues.join('\n') : ''));
    SpreadsheetApp.getUi().alert(summary +
      (issues.length ? '\n' + issues.slice(0, 10).join('\n') : ''));
    return { transferredRows: transferredRows, transferredCells: transferredCells, issues: issues };
  } finally {
    lock.releaseLock();
  }
}


// ========================================
// 4. SINH SỐ HỢP ĐỒNG
// ========================================

/** Gọi Web App của dự án cá nhân để cấp/kiểm tra mã dưới cùng một khóa. */
function taoSoHopDongMoi_(sheet, currentRow) {
  const result = goiDichVuCapSoChung_('allocate', sheet.getName(), currentRow);
  const number = String(result.number || '').trim();
  if (!/^\d+$/.test(number)) throw new Error('Dịch vụ cấp số chung trả mã không hợp lệ.');
  return number;
}

function goiDichVuCapSoChung_(action, sheetName, row) {
  const token = PropertiesService.getScriptProperties()
    .getProperty('SHARED_CONTRACT_NUMBER_TOKEN');
  if (!token) {
    throw new Error('Chưa cài khóa cấp số chung. Hãy dùng menu Cài kết nối cấp số chung.');
  }
  const url = String(CONFIG.CENTRAL_NUMBER_URL || '').trim();
  if (!/^https:\/\/script\.google\.com\/(?:macros\/s|a\/macros\/[^/]+\/s)\/[^/?#]+\/exec$/.test(url)) {
    throw new Error('CENTRAL_NUMBER_URL chưa phải URL Web App hợp lệ.');
  }
  let response;
  try {
    response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        action: action, token: token,
        sheetName: sheetName, row: row
      }),
      muteHttpExceptions: true,
      followRedirects: true,
      timeoutSeconds: 30
    });
  } catch (error) {
    throw new Error('Không gọi được dịch vụ cấp số chung: ' + error.message);
  }
  let result;
  try {
    result = JSON.parse(response.getContentText());
  } catch (error) {
    throw new Error('Dịch vụ cấp số chung chưa trả JSON. Kiểm tra triển khai Web App và quyền truy cập.');
  }
  if (!result.ok) throw new Error(result.error || 'Dịch vụ cấp số chung từ chối yêu cầu.');
  return result;
}

function kiemTraKetNoiCapSoChung() {
  const result = goiDichVuCapSoChung_('ping', CONFIG.SHEET_NAME, 0);
  if (!result.ready) throw new Error('Dịch vụ cấp số chung chưa sẵn sàng.');
  const message = 'Kết nối cấp số chung đã sẵn sàng.';
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (error) {
    if (!/Cannot call SpreadsheetApp\.getUi\(\) from this context/.test(String(error))) throw error;
    Logger.log(message);
  }
  return true;
}

function caiDatKetNoiCapSoChung() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  if (layEmailNguoiDung_() !== layEmailChuSoHuu_(ss)) {
    throw new Error('Chỉ chủ sở hữu bảng tính được cài kết nối cấp số chung.');
  }
  let ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (error) {
    if (!/Cannot call SpreadsheetApp\.getUi\(\) from this context/.test(String(error))) throw error;
    throw new Error('Dự án Apps Script độc lập: vào Cài đặt dự án > Thuộc tính tập lệnh, ' +
      'thêm SHARED_CONTRACT_NUMBER_TOKEN và dán khóa lấy từ dự án cá nhân.');
  }
  const answer = ui.prompt(
    'Cài kết nối cấp số chung',
    'Dán khóa lấy từ menu Lấy khóa cấp số chung của dự án cá nhân:',
    ui.ButtonSet.OK_CANCEL
  );
  if (answer.getSelectedButton() !== ui.Button.OK) return;
  const token = String(answer.getResponseText() || '').trim();
  if (!/^[a-fA-F0-9]{64}$/.test(token)) {
    throw new Error('Khóa cấp số chung phải gồm 64 ký tự hex.');
  }
  PropertiesService.getScriptProperties()
    .setProperty('SHARED_CONTRACT_NUMBER_TOKEN', token);
  ui.alert('Đã lưu khóa. Hãy cập nhật triển khai Web App dự án cá nhân trước khi tạo Word.');
}

/** Chỉ xóa công thức ARRAYFORMULA tạo STT cũ; không xóa số hợp đồng đã tạo. */
function xoaCongThucSttCu_(sheet) {
  sheet.getRange('A1').setValue('Số hợp đồng');

  const formula = String(sheet.getRange('A2').getFormula() || '');

  if (/ARRAYFORMULA/i.test(formula)) {
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, 1).clearContent();
  }
}


// ========================================
// 5. HÀM HỖ TRỢ GOOGLE DOCS
// ========================================

// Mẫu hiện có hai nhãn ngân hàng nhưng chưa đặt biến. Chỉ thêm biến trên
// bản sao để không phải sửa quyền hoặc bố cục của Google Docs mẫu.
function dienBienNganHangVaoMau_(doc) {
  const body = doc.getBody();
  const text = body.getText();
  const fields = [
    { label: 'Số tài khoản:', token: '{{SO_TAI_KHOAN}}' },
    { label: 'Mở tại:', token: '{{MO_TAI}}' }
  ];

  fields.forEach(field => {
    if (text.includes(field.token)) return;
    if (!text.includes(field.label)) {
      throw new Error('Mẫu hợp đồng thiếu vị trí "' + field.label + '".');
    }
    body.replaceText(field.label, field.label + ' ' + field.token);
  });
}

/** Đổi ví dụ cố định ở mục 3.2–3.3 thành biến trên bản sao mẫu. */
function dienBienThuLaoVaoMau_(doc) {
  const body = doc.getBody();
  const fields = [
    { sample: '750.000', token: '{{THUE_TNCN}}' },
    { sample: '6.750.000', token: '{{SO_TIEN_THUC_NHAN}}' },
    {
      sample: 'Sáu triệu bảy trăm năm mươi nghìn',
      token: '{{SO_TIEN_THUC_NHAN_BANG_CHU}}'
    }
  ];

  // Thay số dài trước để "750.000" không khớp một phần của "6.750.000".
  fields.sort((a, b) => b.sample.length - a.sample.length);
  fields.forEach(field => {
    if (body.getText().includes(field.token)) return;
    if (!body.getText().includes(field.sample)) {
      throw new Error('Mẫu hợp đồng thiếu biến hoặc giá trị mẫu: ' + field.token);
    }
    const escapedSample = field.sample.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    body.replaceText(escapedSample, field.token);
    if (!body.getText().includes(field.token)) {
      throw new Error('Không thể thay giá trị mẫu bằng biến: ' + field.token);
    }
  });
}

function thayBienTrongTaiLieu_(doc, replacements) {
  const containers = [doc.getBody()];
  const header = doc.getHeader();
  const footer = doc.getFooter();

  if (header) containers.push(header);
  if (footer) containers.push(footer);

  Object.entries(replacements).forEach(([key, value]) => {
    const escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

    containers.forEach(container => {
      container.replaceText(escapedKey, String(value ?? ''));
    });
  });
}

function timBienConLai_(doc) {
  const texts = [doc.getBody().getText()];
  const header = doc.getHeader();
  const footer = doc.getFooter();

  if (header) texts.push(header.getText());
  if (footer) texts.push(footer.getText());

  const matches = texts.join('\n').match(/\{\{[^{}]+\}\}/g) || [];
  return [...new Set(matches)];
}

function layIdGoogleDoc_(url) {
  const match = String(url || '').trim().match(
    /docs\.google\.com\/document\/d\/([a-zA-Z0-9_-]+)/
  );

  if (!match) {
    throw new Error('TEMPLATE_URL chưa đúng định dạng Google Docs.');
  }

  return match[1];
}

function tachNgay_(dateText) {
  const text = String(dateText || '').trim();
  const match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);

  if (!match) {
    throw new Error('Ngày phải có định dạng DD/MM/YYYY: ' + text);
  }

  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3]);
  const date = new Date(year, month - 1, day);

  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    throw new Error('Ngày không hợp lệ: ' + text);
  }

  return {
    ngay: String(day).padStart(2, '0'),
    thang: String(month).padStart(2, '0'),
    nam: String(year)
  };
}

function layHoacTaoThuMucStt_(parentFolder, stt) {
  const folderName = 'STUDIO CN - ' + dinhDangStt_(stt);
  const folders = parentFolder.getFoldersByName(folderName);

  return folders.hasNext()
    ? folders.next()
    : parentFolder.createFolder(folderName);
}

function dinhDangStt_(stt) {
  const text = String(stt || '').trim();

  if (!text) {
    throw new Error('STT không được để trống.');
  }

  return /^\d+$/.test(text)
    ? text.padStart(2, '0')
    : lamSachTenFile_(text);
}

function lamSachTenFile_(name) {
  return String(name || '')
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Khóa link Word O:Q, V:X và hai cặp FILE CỨNG. */
function khoaCotLink_(sheet, maxRows) {
  const protectionDescriptions = [
    CONFIG.LINK_PROTECTION_DESCRIPTION,
    CONFIG.BBNT_LINK_PROTECTION_DESCRIPTION,
    CONFIG.HARD_FILE_PROTECTION_DESCRIPTION,
    CONFIG.BBNT_HARD_FILE_PROTECTION_DESCRIPTION
  ];

  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter(protection =>
      protectionDescriptions.includes(protection.getDescription())
    )
    .forEach(protection => protection.remove());

  const linkRange = sheet.getRange(
    2,
    CONFIG.FIRST_LINK_COLUMN,
    maxRows - 1,
    CONFIG.LINK_COUNT
  );

  linkRange.clearDataValidations();
  const bbntLinkRange = sheet.getRange(
    2, CONFIG.BBNT_FIRST_LINK_COLUMN, maxRows - 1, CONFIG.LINK_COUNT
  );
  bbntLinkRange.clearDataValidations();

  const protectedRanges = [
    {
      range: linkRange,
      description: CONFIG.LINK_PROTECTION_DESCRIPTION
    },
    {
      range: bbntLinkRange,
      description: CONFIG.BBNT_LINK_PROTECTION_DESCRIPTION
    },
    {
      range: sheet.getRange(
        2,
        CONFIG.HARD_UPLOAD_COLUMN,
        maxRows - 1,
        2
      ),
      description: CONFIG.HARD_FILE_PROTECTION_DESCRIPTION
    },
    {
      range: sheet.getRange(
        2, CONFIG.BBNT_HARD_UPLOAD_COLUMN, maxRows - 1, 2
      ),
      description: CONFIG.BBNT_HARD_FILE_PROTECTION_DESCRIPTION
    }
  ];
  const effectiveUser = Session.getEffectiveUser();
  const effectiveEmail = effectiveUser.getEmail();

  protectedRanges.forEach(item => {
    const protection = item.range
      .protect()
      .setDescription(item.description);

    protection.addEditor(effectiveUser);

    const otherEditors = protection.getEditors().filter(editor =>
      editor.getEmail() !== effectiveEmail
    );

    if (otherEditors.length > 0) {
      protection.removeEditors(otherEditors);
    }

    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
  });
}


// ========================================
// 6. BIẾN CẦN ĐẶT TRONG GOOGLE DOCS MẪU
// ========================================

const TEMPLATE_FIELDS = [
  '{{SO_HD}}',
  '{{HO_TEN}}',
  '{{HO_TEN_IN_HOA}}',
  '{{NGAY_LAM_DON}}',
  '{{NGAY}}', '{{THANG}}', '{{NAM}}',
  '{{SO_CCCD}}',
  '{{MA_SO_THUE}}',
  '{{DIA_CHI}}',
  '{{SO_DIEN_THOAI}}',
  '{{TEN_DU_AN}}',
  '{{CONG_VIEC_GIAO_KHOAN}}',
  '{{NGAY_HOAN_THANH}}',
  '{{TONG_THU_LAO}}',
  '{{TONG_THU_LAO_BANG_CHU}}',
  '{{THUE_TNCN}}',
  '{{SO_TIEN_THUC_NHAN}}',
  '{{SO_TIEN_THUC_NHAN_BANG_CHU}}',
  '{{SO_TAI_KHOAN}}',
  '{{MO_TAI}}'
];
