/**
 * TỰ ĐỘNG TẠO HỢP ĐỒNG KHOÁN DỊCH VỤ
 *
 * A:X - Dữ liệu hợp đồng (H là Tên Brand, W là Số Tài khoản, X là Mở Tại)
 * Y:AA - Link Word lần 01, 02, 03
 * AB   - Checkbox tạo Word
 * AC   - Tải PDF/JPG/PNG FILE CỨNG
 * AD   - Link FILE CỨNG
 * AE   - Trạng thái nghiệp vụ (không do code này xử lý)
 *
 * Trước khi sử dụng, thay các phần bôi vàng trong Google Docs mẫu bằng
 * các biến {{...}} được liệt kê trong TEMPLATE_FIELDS ở cuối file.
 */

const CONFIG = {
  SHEET_NAME: 'TH_HỢP ĐỒNG KHOÁN DỊCH VỤ',

  TEMPLATE_URL:
    'https://docs.google.com/document/d/1M9TPokSCytSej9y_xq6QQupABJJLNvcP5IMfw9P9Qt8/edit?tab=t.0',

  // Thư mục "HỢP ĐỒNG KHOÁN DỊCH VỤ" trong ảnh.
  // Nếu Drive báo không tìm thấy thư mục, hãy dán lại ID thư mục vào đây.
  FOLDER_ID: '1FCIPy1gl88pMOlUnn4ZtgPIPYflzzmNg',

  // Sau khi triển khai Web App, dán URL kết thúc bằng /exec vào đây.
  WEB_APP_URL: 'https://script.google.com/a/macros/getvini.com/s/AKfycby5CqdBQn8biRP4iWmNywuF3lxUIwQDSvgWF7uvMx6BMoiUOGCquB2371RjpsV8lTqb/exec',

  FIRST_LINK_COLUMN: 25, // Y
  LINK_COUNT: 3,         // Y, Z, AA
  ACTION_COLUMN: 28,     // AB
  LINK_PROTECTION_DESCRIPTION: 'KHOA_LINK_HOP_DONG_TU_DONG',
  HARD_FILE_PROTECTION_DESCRIPTION: 'KHOA_FILE_CUNG_HOP_DONG_TU_DONG',
  HARD_UPLOAD_COLUMN: 29, // AC
  HARD_LINK_COLUMN: 30,   // AD
  AMOUNT_COLUMN: 21,       // U
  AMOUNT_WORDS_COLUMN: 22, // V
  UPLOAD_MAX_SIZE_MB: 10
};


// ========================================
// 0. TRANG WEB TẢI FILE CỨNG
// Tái sử dụng cùng cơ chế với Đề nghị thanh toán.
// ========================================

function doGet(e) {
  const loai = String(e && e.parameter && e.parameter.loai || '');
  const row = Number(e && e.parameter && e.parameter.row || 0);

  if (loai !== 'FILE_CUNG' || !Number.isInteger(row) || row < 2) {
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
  template.tieuDe = 'FILE CỨNG';
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
    .addItem('Tạo hợp đồng cho dòng đang chọn', 'taoHopDongDongDangChon')
    .addSeparator()
    .addItem('Tải FILE CỨNG cho dòng đang chọn', 'moTaiFileCung')
    .addToUi();
}

/**
 * Chạy hàm này đúng một lần sau khi dán code.
 * Hàm tạo Y:AB, FILE CỨNG tại AC:AD và installable onEdit trigger.
 */
function caiDatHeThong() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error('Không tìm thấy tab: ' + CONFIG.SHEET_NAME);
  }

  const templateId = layIdGoogleDoc_(CONFIG.TEMPLATE_URL);
  const template = DriveApp.getFileById(templateId);

  if (template.getMimeType() !== MimeType.GOOGLE_DOCS) {
    throw new Error('File mẫu phải là Google Docs.');
  }

  // Kiểm tra quyền truy cập trước khi tạo trigger.
  DocumentApp.openById(templateId).getBody();
  DriveApp.getFolderById(CONFIG.FOLDER_ID).getName();

  PropertiesService.getScriptProperties().setProperty(
    'SPREADSHEET_ID',
    ss.getId()
  );

  xoaCongThucSttCu_(sheet);

  sheet.getRange('Y1:AB1').setValues([[
    'Link_Word_01',
    'Link_Word_02',
    'Link_Word_03',
    'TẠO WORD'
  ]]);
  sheet.getRange('AC1:AD1').setValues([[
    'FILE CỨNG',
    'Link_FILE_Cung'
  ]]);

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

    khoaCotLink_(sheet, maxRows);
    taoNutUploadPdf_(sheet, maxRows);
  }

  sheet.setColumnWidths(CONFIG.FIRST_LINK_COLUMN, CONFIG.LINK_COUNT, 180);
  sheet.setColumnWidth(CONFIG.ACTION_COLUMN, 120);
  sheet.setColumnWidths(CONFIG.HARD_UPLOAD_COLUMN, 2, 180);

  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'xuLyTaoHopDong')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('xuLyTaoHopDong')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ss.toast(
    coLinkWebAppHopLe_()
      ? 'Đã tạo checkbox AB và nút tải FILE CỨNG tại AC.'
      : 'Đã cài hợp đồng. Hãy triển khai Web App và điền WEB_APP_URL để tạo nút AC.',
    'Cài đặt thành công',
    8
  );
}

function coLinkWebAppHopLe_() {
  return /^https:\/\/script\.google\.com\/(?:macros\/s|a\/macros\/[^/]+\/s)\/[^/?#]+\/exec(?:[?#].*)?$/
    .test(String(CONFIG.WEB_APP_URL || '').trim());
}

/** Dùng lại mẫu nút upload của Đề nghị thanh toán, chỉ tạo FILE CỨNG. */
function taoNutUploadPdf_(sheet, maxRows) {
  if (!coLinkWebAppHopLe_() || maxRows < 2) return;

  const buttons = [];
  const baseUrl = String(CONFIG.WEB_APP_URL).trim();

  for (let row = 2; row <= maxRows; row++) {
    const uploadUrl = baseUrl + '?loai=FILE_CUNG&row=' + row;

    buttons.push([
      SpreadsheetApp.newRichTextValue()
        .setText('⬆ TẢI FILE CỨNG')
        .setLinkUrl(uploadUrl)
        .build()
    ]);
  }

  sheet.getRange(2, CONFIG.HARD_UPLOAD_COLUMN, maxRows - 1, 1)
    .setRichTextValues(buttons)
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setFontColor('#1155CC')
    .setBackground('#E8F0FE');
}

function moTaiFileCung() {
  moHopThoaiTaiPdf_('FILE_CUNG');
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
  template.tieuDe = 'FILE CỨNG';
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

  if (loai !== 'FILE_CUNG') {
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

  const currentLink = String(
    sheet.getRange(row, CONFIG.HARD_LINK_COLUMN).getDisplayValue() || ''
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

  const mainFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
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

  sheet.getRange(row, CONFIG.HARD_UPLOAD_COLUMN)
    .setNote('Đã tải: ' + safeName);
  sheet.getRange(row, CONFIG.HARD_LINK_COLUMN).setValue(viewUrl);
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
// 2. TỰ ĐIỀN CỘT V VÀ XỬ LÝ CHECKBOX Ở CỘT AB
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

/** Viết số tiền đồng bằng chữ, giữ cách hiển thị của cột V hiện tại (không thêm "đồng"). */
function vietSoTienBangChu_(amount) {
  const units = ['', 'nghìn', 'triệu', 'tỷ', 'nghìn tỷ'];
  const groups = [];
  let remaining = amount;

  while (remaining > 0) {
    groups.push(remaining % 1000);
    remaining = Math.floor(remaining / 1000);
  }
  if (groups.length > units.length) {
    throw new Error('Số tiền ở cột U quá lớn để chuyển thành chữ.');
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

function capNhatThuLaoBangChu_(sheet, row) {
  const amount = docSoTien_(sheet.getRange(row, CONFIG.AMOUNT_COLUMN).getValue());
  sheet.getRange(row, CONFIG.AMOUNT_WORDS_COLUMN)
    .setValue(amount === null ? '' : vietSoTienBangChu_(amount));
}

function xuLyTaoHopDong(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();

  if (sheet.getName() !== CONFIG.SHEET_NAME) return;

  // Hỗ trợ cả khi dán nhiều dòng hoặc nhiều cột có chứa U.
  if (
    range.getColumn() <= CONFIG.AMOUNT_COLUMN &&
    range.getLastColumn() >= CONFIG.AMOUNT_COLUMN
  ) {
    for (let row = Math.max(2, range.getRow()); row <= range.getLastRow(); row++) {
      capNhatThuLaoBangChu_(sheet, row);
    }
  }

  if (
    range.getColumn() !== CONFIG.ACTION_COLUMN ||
    range.getRow() < 2 ||
    range.getNumRows() !== 1 ||
    range.getNumColumns() !== 1 ||
    e.value !== 'TRUE'
  ) {
    return;
  }

  const actionCell = range;
  const row = range.getRow();
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();

  if (!lock.tryLock(15000)) {
    actionCell.setValue(false);
    actionCell.setNote('Hệ thống đang xử lý yêu cầu khác. Hãy thử lại.');
    return;
  }

  try {
    actionCell.setNote('Đang tạo hợp đồng...');
    const result = taoHopDongChoDong_(sheet, row);

    actionCell.setNote(
      'Đã tạo hợp đồng lần ' + result.lanTao + ': ' + result.url
    );

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Đã tạo hợp đồng lần ' + result.lanTao + ' cho dòng ' + row,
      'Thành công',
      6
    );
  } catch (error) {
    actionCell.setNote('Không thể tạo hợp đồng: ' + error.message);
    console.error('Dòng ' + row + ': ' + error.stack);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      error.message,
      'Không thể tạo hợp đồng - Dòng ' + row,
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

    ss.toast(
      'Đã tạo hợp đồng lần ' + result.lanTao + ' cho dòng ' + row,
      'Thành công',
      6
    );
  } finally {
    lock.releaseLock();
  }
}


// ========================================
// 3. ĐỌC DỮ LIỆU VÀ TẠO HỢP ĐỒNG
// ========================================

function taoHopDongChoDong_(sheet, row) {
  xoaCongThucSttCu_(sheet);

  // A:X là dữ liệu, Y:AA là ba link đã tạo.
  const displayValues = sheet.getRange(row, 1, 1, CONFIG.FIRST_LINK_COLUMN + CONFIG.LINK_COUNT - 1)
    .getDisplayValues()[0];
  const rawValues = sheet.getRange(row, 1, 1, CONFIG.FIRST_LINK_COLUMN - 1).getValues()[0];

  const ngayLamDonHienThi = String(displayValues[2] || '').trim();
  const tenBrandHienThi = String(displayValues[7] || '').trim();

  if (!ngayLamDonHienThi) {
    throw new Error('Thiếu Ngày làm đơn ở cột C.');
  }

  if (!tenBrandHienThi) {
    throw new Error('Thiếu Tên Brand ở cột H.');
  }

  const ngayKy = tachNgay_(ngayLamDonHienThi);
  const brandCode = chuanHoaBrand_(tenBrandHienThi);

  const headers = [
    'Số hợp đồng',
    'Họ và tên người đề nghị',
    'Ngày làm đơn',
    'Số CCCD',
    'Mã số thuế',
    'Địa chỉ',
    'Số điện thoại',
    'Tên Brand',
    'Tên dự án',
    'Số lượng Video (Công việc giao khoán)',
    'Nội dung Kịch bản',
    'Số lượng Kịch bản',
    'Số lượng Video',
    'Khung Video',
    'Chất lượng Video',
    'Thời lượng Video',
    'Hiệu ứng Video',
    'Nội dung Diễn Viên',
    'Số lượng Diễn viên',
    'Ngày hoàn thành công việc',
    'Tổng giá trị thù lao giao khoán',
    'Tổng giá trị thù lao giao khoán (Bằng chữ)',
    'Số Tài khoản',
    'Mở Tại'
  ];

  const errorValues = [
    '#ERROR!', '#REF!', '#VALUE!', '#N/A',
    '#DIV/0!', '#NUM!', '#NAME?'
  ];
  const missing = [];

  const amount = docSoTien_(rawValues[CONFIG.AMOUNT_COLUMN - 1]);
  if (amount === null) {
    throw new Error('Tổng giá trị thù lao ở cột U phải là số nguyên lớn hơn 0 (ví dụ: 2,160,000).');
  }
  const amountWords = vietSoTienBangChu_(amount);
  sheet.getRange(row, CONFIG.AMOUNT_WORDS_COLUMN).setValue(amountWords);
  displayValues[CONFIG.AMOUNT_WORDS_COLUMN - 1] = amountWords;

  // Cột A do code tự sinh; bắt buộc người dùng nhập đầy đủ B:X.
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
  const currentRowHasLink = links.some(link => String(link).trim() !== '');

  if (emptyLinkIndex === -1) {
    throw new Error('Dòng này đã tạo đủ 3 hợp đồng.');
  }

  // Kiểm tra ngày hoàn thành trước khi cấp số hợp đồng.
  const ngayHoanTat = tachNgay_(displayValues[19]);
  let soHopDong = String(displayValues[0] || '').trim();

  // Chỉ giữ số cũ khi dòng này đã có link Word. Nếu chưa có link,
  // luôn tính lại từ các dòng thực sự có link để không giữ số "ảo".
  if (!currentRowHasLink || !laSoHopDongMoi_(soHopDong)) {
    soHopDong = taoSoHopDongMoi_(sheet, row, ngayKy, brandCode);
    sheet.getRange(row, 1)
      .setNumberFormat('@')
      .setValue(soHopDong);
    displayValues[0] = soHopDong;
    SpreadsheetApp.flush();
  }

  const [
    stt,
    hoTen,
    ngayLamDon,
    soCccd,
    maSoThue,
    diaChi,
    soDienThoai,
    tenBrand,
    tenDuAn,
    soLuongCongViec,
    noiDungKichBan,
    soLuongKichBan,
    soLuongVideo,
    khungVideo,
    chatLuongVideo,
    thoiLuongVideo,
    hieuUngVideo,
    noiDungDienVien,
    soLuongDienVien,
    ngayHoanThanh,
    tongThuLao,
    tongThuLaoBangChu,
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
    stt + ' - ' + hoTen + ' - HĐ khoán việc - ' + lanTao
  );
  const copy = template.makeCopy(fileName, sttFolder);

  try {
    const doc = DocumentApp.openById(copy.getId());
    dienBienNganHangVaoMau_(doc);
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
      '{{TEN_BRAND}}': tenBrand,
      '{{TEN_DU_AN}}': tenDuAn,
      '{{SO_LUONG_CONG_VIEC}}': soLuongCongViec,
      '{{NOI_DUNG_KICH_BAN}}': noiDungKichBan,
      '{{SO_LUONG_KICH_BAN}}': soLuongKichBan,
      '{{SO_LUONG_VIDEO}}': soLuongVideo,
      '{{KHUNG_VIDEO}}': khungVideo,
      '{{CHAT_LUONG_VIDEO}}': chatLuongVideo,
      '{{THOI_LUONG_VIDEO}}': thoiLuongVideo,
      '{{HIEU_UNG_VIDEO}}': hieuUngVideo,
      '{{NOI_DUNG_DIEN_VIEN}}': noiDungDienVien,
      '{{SO_LUONG_DIEN_VIEN}}': soLuongDienVien,
      '{{NGAY_HOAN_THANH}}': ngayHoanThanh,
      '{{NGAY_HT}}': ngayHoanTat.ngay,
      '{{THANG_HT}}': ngayHoanTat.thang,
      '{{NAM_HT}}': ngayHoanTat.nam,
      '{{TONG_THU_LAO}}': tongThuLao,
      '{{TONG_THU_LAO_BANG_CHU}}': tongThuLaoBangChu,
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


// ========================================
// 4. SINH SỐ HỢP ĐỒNG
// ========================================

/**
 * Ví dụ: 14/07/2026 + LASHE ->
 * 140726-01/LASHE
 *
 * Bộ đếm tăng riêng cho từng cặp Ngày làm đơn + Brand.
 * Chỉ các dòng có ít nhất một link Word tại Y:AA mới được tính.
 */
function taoSoHopDongMoi_(sheet, currentRow, ngayKy, brandCode) {
  const dateCode = ngayKy.ngay + ngayKy.thang + ngayKy.nam.slice(-2);
  const prefix = dateCode + '-';
  const suffix = '/' + brandCode;
  const pattern = new RegExp(
    '^' + escapeRegExp_(prefix) + '(\\d+)' + escapeRegExp_(suffix) +
    '(?:\\/\\d{4}\\/HĐKV-GETVINI)?$'
  );
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const existingRows = sheet.getRange(
    2, 1, lastRow - 1, CONFIG.FIRST_LINK_COLUMN + CONFIG.LINK_COUNT - 1
  )
    .getDisplayValues();
  let maxSequence = 0;

  existingRows.forEach((values, index) => {
    const sheetRow = index + 2;

    if (sheetRow === currentRow) return;

    const hasWordLink = values
      .slice(CONFIG.FIRST_LINK_COLUMN - 1, CONFIG.FIRST_LINK_COLUMN - 1 + CONFIG.LINK_COUNT)
      .some(link => String(link || '').trim() !== '');

    if (!hasWordLink) return;

    const match = String(values[0] || '').trim().match(pattern);

    if (match) {
      maxSequence = Math.max(maxSequence, Number(match[1]));
    }
  });

  const nextSequence = String(maxSequence + 1).padStart(2, '0');
  return prefix + nextSequence + suffix;
}

function laSoHopDongMoi_(value) {
  return /^\d{6}-\d{2,}\/[^/]+$/
    .test(String(value || '').trim());
}

function chuanHoaBrand_(brand) {
  const normalized = String(brand || '')
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/Đ/g, 'D')
    .replace(/[^A-Z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .replace(/-+/g, '-');

  if (!normalized) {
    throw new Error('Tên Brand ở cột H không hợp lệ.');
  }

  return normalized;
}

function escapeRegExp_(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
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
  const folderName = dinhDangStt_(stt);
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

/** Khóa vùng link Word Y:AA và FILE CỨNG AC:AD. */
function khoaCotLink_(sheet, maxRows) {
  const protectionDescriptions = [
    CONFIG.LINK_PROTECTION_DESCRIPTION,
    CONFIG.HARD_FILE_PROTECTION_DESCRIPTION
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

  const protectedRanges = [
    {
      range: linkRange,
      description: CONFIG.LINK_PROTECTION_DESCRIPTION
    },
    {
      range: sheet.getRange(
        2,
        CONFIG.HARD_UPLOAD_COLUMN,
        maxRows - 1,
        2
      ),
      description: CONFIG.HARD_FILE_PROTECTION_DESCRIPTION
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
  '{{TEN_BRAND}}',
  '{{TEN_DU_AN}}',
  '{{SO_LUONG_CONG_VIEC}}',
  '{{NOI_DUNG_KICH_BAN}}',
  '{{SO_LUONG_KICH_BAN}}',
  '{{SO_LUONG_VIDEO}}',
  '{{KHUNG_VIDEO}}',
  '{{CHAT_LUONG_VIDEO}}',
  '{{THOI_LUONG_VIDEO}}',
  '{{HIEU_UNG_VIDEO}}',
  '{{NOI_DUNG_DIEN_VIEN}}',
  '{{SO_LUONG_DIEN_VIEN}}',
  '{{NGAY_HOAN_THANH}}',
  '{{TONG_THU_LAO}}',
  '{{TONG_THU_LAO_BANG_CHU}}',
  '{{SO_TAI_KHOAN}}',
  '{{MO_TAI}}'
];
