/**
 * TỰ ĐỘNG TẠO HỢP ĐỒNG KHOÁN DỊCH VỤ
 *
 * A:W - Dữ liệu hợp đồng (đã xóa cột Tên Brand; V là Số Tài khoản, W là Mở Tại)
 * X:Z - Link Word lần 01, 02, 03
 * AA  - Checkbox tạo Word
 * AB  - Checkbox tự động điền BBNT
 * AC  - Tải PDF/JPG/PNG FILE CỨNG
 * Link_FILE_Cung được tìm theo tiêu đề cột để không ghi đè các cột nghiệp vụ.
 *
 * Đặt các biến {{...}} trong Google Docs mẫu theo TEMPLATE_FIELDS ở cuối file.
 * Riêng các số mẫu ở mục 3.2–3.3 được đổi thành biến trên bản sao hợp đồng.
 */

const CONFIG = {
  SHEET_NAME: 'HỢP ĐỒNG KHOÁN MẪU CÁ NHÂN',
  SPREADSHEET_ID: '1uFp7McsEUjFQzMsQrW0ZSiqxQe7324Wy6KMzsHP_Mq0',

  TEMPLATE_URL:
    'https://docs.google.com/document/d/1LnQFgP1Rpzvi30xW2OjSG7E1JguffbLlRe24Js738Mw/edit?tab=t.0',

  // Thư mục "HỢP ĐỒNG KHOÁN DỊCH VỤ" trong ảnh.
  // Nếu Drive báo không tìm thấy thư mục, hãy dán lại ID thư mục vào đây.
  FOLDER_ID: '1FCIPy1gl88pMOlUnn4ZtgPIPYflzzmNg',

  // Sau khi triển khai Web App, dán URL kết thúc bằng /exec vào đây.
  WEB_APP_URL: 'https://script.google.com/a/macros/getvini.com/s/AKfycbwkG4U2MgM9aywiEcV7yKikOko-Wh0VRqE9OXGg3QC292L9AAxZ3u1mzb9UVrbDERO2hg/exec',

  FIRST_LINK_COLUMN: 24, // X
  LINK_COUNT: 3,         // X, Y, Z
  ACTION_COLUMN: 27,     // AA
  BBNT_ACTION_COLUMN: 28, // AB
  BBNT_SHEET_NAME: 'BBNT MẪU CÁ NHÂN',
  // URL Web App của Code_Bien_Ban_Nghiem_Thu_Moi; dùng khi cột Z không còn nút mẫu.
  BBNT_WEB_APP_URL: 'https://script.google.com/a/macros/getvini.com/s/AKfycbxp_ehCPWJIEmjhPaLm7gzFA-63Roi3jnI7yENdLC6XCUUG0Ous8JW8PU804BTO5bl2/exec',
  LINK_PROTECTION_DESCRIPTION: 'KHOA_LINK_HOP_DONG_TU_DONG',
  HARD_FILE_PROTECTION_DESCRIPTION: 'KHOA_FILE_CUNG_HOP_DONG_TU_DONG',
  HARD_UPLOAD_COLUMN: 29, // AC
  APPLICATION_DATE_COLUMN: 3, // C: Ngày làm đơn
  COMPLETION_DATE_COLUMN: 19, // S: Ngày hoàn thành công việc
  AMOUNT_COLUMN: 20,       // T: Số tiền thực nhận sau thuế
  AMOUNT_WORDS_COLUMN: 21, // U: Số tiền thực nhận sau thuế bằng chữ (tự điền)
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
    .addItem('Khôi phục nút FILE CỨNG BBNT', 'khoiPhucNutFileCungBbnt')
    .addSeparator()
    .addItem('Tải FILE CỨNG cho dòng đang chọn', 'moTaiFileCung')
    .addToUi();
}

/**
 * Chạy hàm này đúng một lần sau khi dán code.
 * Hàm tạo X:AA, checkbox BBNT ở AB, FILE CỨNG ở AC và onEdit trigger.
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

  sheet.getRange('X1:AA1').setValues([[
    'Link_Word_01',
    'Link_Word_02',
    'Link_Word_03',
    'TẠO WORD'
  ]]);
  const bbntHeader = String(sheet.getRange('AB1').getDisplayValue() || '').trim();
  const hardUploadHeader = String(sheet.getRange('AC1').getDisplayValue() || '').trim();
  if (bbntHeader && bbntHeader !== 'Tự động điền BBNT') {
    throw new Error('Cột AB phải là Tự động điền BBNT. Hãy kiểm tra bố cục Sheet.');
  }
  if (hardUploadHeader && hardUploadHeader !== 'FILE CỨNG') {
    throw new Error('Cột AC phải là FILE CỨNG. Hãy kiểm tra bố cục Sheet.');
  }
  sheet.getRange('AB1').setValue('Tự động điền BBNT');
  sheet.getRange('AC1').setValue('FILE CỨNG');
  layCotLinkFileCung_(sheet, true);

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
  sheet.setColumnWidth(CONFIG.BBNT_ACTION_COLUMN, 170);
  sheet.setColumnWidth(CONFIG.HARD_UPLOAD_COLUMN, 180);

  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'xuLyTaoHopDong')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('xuLyTaoHopDong')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ss.toast(
    coLinkWebAppHopLe_()
      ? 'Đã tạo checkbox AA, AB và nút tải FILE CỨNG tại AC.'
      : 'Đã cài hợp đồng. Hãy triển khai Web App và điền WEB_APP_URL để tạo nút AC.',
    'Cài đặt thành công',
    8
  );
}

function coLinkWebAppHopLe_() {
  return /^https:\/\/script\.google\.com\/(?:macros\/s|a\/macros\/[^/]+\/s)\/[^/?#]+\/exec(?:[?#].*)?$/
    .test(String(CONFIG.WEB_APP_URL || '').trim());
}

/** Ngăn ghi nhầm cột nếu Sheet vẫn còn cột Tên Brand ở H. */
function coBoCucHopDongMoi_(sheet) {
  const projectHeader = String(sheet.getRange('H1').getDisplayValue() || '').trim();
  const firstLinkHeader = String(sheet.getRange('X1').getDisplayValue() || '').trim();
  return projectHeader === 'Tên dự án' && firstLinkHeader === 'Link_Word_01';
}

function kiemTraBoCucHopDong_(sheet) {
  if (!coBoCucHopDongMoi_(sheet)) {
    throw new Error('Bố cục Sheet chưa đúng: hãy xóa hẳn cột H (Tên Brand) trước khi chạy code mới.');
  }
}

/** Link FILE CỨNG có thể nằm sau các cột nghiệp vụ đã thêm. */
function layCotLinkFileCung_(sheet, createIfMissing) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];
  const index = headers.findIndex(header =>
    String(header || '').trim() === 'Link_FILE_Cung'
  );
  if (index >= 0) return index + 1;
  if (!createIfMissing) {
    throw new Error('Chưa có cột Link_FILE_Cung. Hãy chạy Cài đặt hệ thống.');
  }

  const column = sheet.getLastColumn() + 1;
  if (column > sheet.getMaxColumns()) sheet.insertColumnAfter(column - 1);
  sheet.getRange(1, column).setValue('Link_FILE_Cung');
  return column;
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

  const hardLinkColumn = layCotLinkFileCung_(sheet, false);
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
// 2. TỰ ĐIỀN CỘT U VÀ XỬ LÝ CHECKBOX Ở CỘT AA
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
    throw new Error('Số tiền ở cột T quá lớn để chuyển thành chữ.');
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
    throw new Error('Số tiền thực nhận ở cột T quá lớn để tính thù lao trước thuế.');
  }
  return {
    tongThuLao: tongThuLao,
    thueTncn: tongThuLao - soTienThucNhan
  };
}

function dinhDangSoTien_(amount) {
  return String(amount).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

function capNhatNgayHoanThanh_(sheet, row) {
  const ngayLamDon = String(
    sheet.getRange(row, CONFIG.APPLICATION_DATE_COLUMN).getDisplayValue() || ''
  ).trim();
  const cell = sheet.getRange(row, CONFIG.COMPLETION_DATE_COLUMN);

  if (!ngayLamDon) {
    cell.clearContent();
    return;
  }

  let parts;
  try {
    parts = tachNgay_(ngayLamDon);
  } catch (error) {
    cell.clearContent();
    return;
  }

  const ngayHoanThanh = new Date(
    Number(parts.nam), Number(parts.thang) - 1, Number(parts.ngay) + 20
  );
  cell.setNumberFormat('dd/MM/yyyy').setValue(ngayHoanThanh);
}

function dienBbntTuHopDong_(sheet, row) {
  const values = sheet.getRange(row, 1, 1, CONFIG.FIRST_LINK_COLUMN + CONFIG.LINK_COUNT - 1)
    .getDisplayValues()[0];
  const rawValues = sheet.getRange(row, 1, 1, CONFIG.FIRST_LINK_COLUMN - 1)
    .getValues()[0];
  const soHopDong = String(values[0] || '').trim();
  if (!soHopDong) {
    throw new Error('Dòng này chưa có Số hợp đồng ở cột A.');
  }

  const missing = [];
  for (let column = 2; column <= 23; column++) {
    const value = String(values[column - 1] || '').trim();
    if (!value || /^#(?:ERROR!|REF!|VALUE!|N\/A|DIV\/0!|NUM!|NAME\?)/.test(value)) {
      missing.push(columnToLetter_(column));
    }
  }
  if (missing.length) {
    throw new Error('Hãy điền đủ dữ liệu hợp đồng ở cột: ' + missing.join(', '));
  }

  const wordLinks = values.slice(CONFIG.FIRST_LINK_COLUMN - 1,
    CONFIG.FIRST_LINK_COLUMN - 1 + CONFIG.LINK_COUNT);
  if (!wordLinks.some(link =>
    /^https:\/\/docs\.google\.com\/document\/d\/[a-zA-Z0-9_-]+(?:[/?#]|$)/.test(String(link).trim())
  )) {
    throw new Error('Hãy tạo ít nhất một link Word hợp lệ ở cột X:Z trước.');
  }

  const kyNgay = tachNgay_(values[CONFIG.APPLICATION_DATE_COLUMN - 1]);
  const ngayLamDonBbnt = tachNgay_(values[CONFIG.COMPLETION_DATE_COLUMN - 1]);
  const dateValue = parts => new Date(
    Number(parts.nam), Number(parts.thang) - 1, Number(parts.ngay)
  );
  const ss = sheet.getParent();
  const bbntSheet = ss.getSheetByName(CONFIG.BBNT_SHEET_NAME);
  if (!bbntSheet) throw new Error('Không tìm thấy tab ' + CONFIG.BBNT_SHEET_NAME + '.');
  const bbntHeaders = bbntSheet.getRange(1, 2, 1, 4).getDisplayValues()[0]
    .map(value => String(value || '').trim());
  if (
    bbntHeaders[0] !== 'Số Hợp đồng khoán' ||
    bbntHeaders[1] !== 'Ký ngày' ||
    bbntHeaders[3] !== 'Ngày làm đơn'
  ) {
    throw new Error('Bố cục tab BBNT chưa đúng ở các cột B, C hoặc E.');
  }
  if (String(bbntSheet.getRange('V1').getDisplayValue()).trim() !== 'Link_Word_01') {
    throw new Error('Bố cục tab BBNT chưa đúng: cột V phải là Link_Word_01.');
  }

  const lastRow = bbntSheet.getLastRow();
  const contractNumbers = lastRow < 2 ? [] : bbntSheet
    .getRange(2, 2, lastRow - 1, 1).getDisplayValues()
    .map(item => String(item[0] || '').trim());
  const matches = contractNumbers
    .map((number, index) => number === soHopDong ? index + 2 : null)
    .filter(value => value !== null);
  if (matches.length > 1) {
    throw new Error('Tab BBNT có nhiều dòng cùng Số hợp đồng ' + soHopDong + '.');
  }
  let targetRow;
  if (matches.length) {
    targetRow = matches[0];
  } else {
    // Chọn vị trí theo mã hợp đồng ở cột B, bỏ qua các dòng trống do
    // checkbox và nút tải FILE CỨNG tạo ra đến tận cuối Sheet.
    let previousRow = 1;
    let nextRow = Infinity;
    contractNumbers.forEach((number, index) => {
      if (!number) return;
      const comparison = number.localeCompare(
        soHopDong, 'vi', { numeric: true, sensitivity: 'base' }
      );
      if (comparison < 0) previousRow = Math.max(previousRow, index + 2);
      if (comparison > 0) nextRow = Math.min(nextRow, index + 2);
    });
    if (previousRow >= nextRow) {
      throw new Error('Thứ tự mã hợp đồng ở cột B của tab BBNT chưa nhất quán.');
    }
    targetRow = previousRow + 1;

    if (targetRow > bbntSheet.getMaxRows()) {
      bbntSheet.insertRowAfter(bbntSheet.getMaxRows());
    } else {
      const rowHasData = bbntSheet.getRange(targetRow, 1, 1, 24)
        .getDisplayValues()[0]
        .some(value => String(value || '').trim() !== '') ||
        String(bbntSheet.getRange(targetRow, 27).getDisplayValue() || '').trim() !== '';
      if (rowHasData) {
        bbntSheet.insertRowsBefore(targetRow, 1);
      }
    }
    bbntSheet.getRange(targetRow, 25).insertCheckboxes(); // Y: TẠO WORD BBNT
  }

  // A:U của BBNT. Các cột đánh giá L, R, U không có dữ liệu nguồn.
  const mapped = [
    soHopDong, soHopDong, dateValue(kyNgay), values[1],
    dateValue(ngayLamDonBbnt), values[3], values[4], values[5], values[6],
    values[9], rawValues[10], '', rawValues[11], values[12], values[13],
    values[14], values[15], '', values[16], rawValues[17], ''
  ];
  const existing = bbntSheet.getRange(targetRow, 1, 1, mapped.length)
    .getDisplayValues()[0];
  let filled = 0;
  mapped.forEach((value, index) => {
    if (value === '' || String(existing[index] || '').trim()) return;
    const cell = bbntSheet.getRange(targetRow, index + 1);
    if (index === 2 || index === 4) cell.setNumberFormat('dd/MM/yyyy');
    if (index === 0 || index === 1) cell.setNumberFormat('@');
    cell.setValue(value);
    filled++;
  });
  capNhatNutTaiBbnt_(bbntSheet);
  SpreadsheetApp.flush();
  return { row: targetRow, created: matches.length === 0, filled: filled };
}

/** Tạo lại mọi nút Z với tham số row đúng, kể cả khi Z2 đang trống. */
function capNhatNutTaiBbnt_(sheet) {
  if (String(sheet.getRange('Z1').getDisplayValue() || '').trim() !== 'FILE CỨNG') {
    throw new Error('Bố cục tab BBNT chưa đúng: cột Z phải là FILE CỨNG.');
  }
  const maxRows = sheet.getMaxRows();
  if (maxRows < 2) return 0;

  const richValues = sheet.getRange(2, 26, maxRows - 1, 1).getRichTextValues();
  let baseUrl = '';
  for (const [richValue] of richValues) {
    const url = richValue && richValue.getLinkUrl();
    const match = String(url || '').match(
      /^(https:\/\/script\.google\.com\/(?:macros\/s|a\/macros\/[^/]+\/s)\/[^/?#]+\/exec)\?loai=FILE_CUNG&row=\d+$/
    );
    if (match) {
      baseUrl = match[1];
      break;
    }
  }
  if (!baseUrl) baseUrl = String(CONFIG.BBNT_WEB_APP_URL || '').trim();
  if (!/^https:\/\/script\.google\.com\/(?:macros\/s|a\/macros\/[^/]+\/s)\/[^/?#]+\/exec$/.test(baseUrl)) {
    throw new Error('Chưa có URL Web App BBNT hợp lệ để tạo nút FILE CỨNG.');
  }

  const buttons = [];
  for (let row = 2; row <= maxRows; row++) {
    buttons.push([
      SpreadsheetApp.newRichTextValue()
        .setText('⬆ TẢI FILE CỨNG')
        .setLinkUrl(baseUrl + '?loai=FILE_CUNG&row=' + row)
        .build()
    ]);
  }
  sheet.getRange(2, 26, buttons.length, 1)
    .setRichTextValues(buttons)
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setFontColor('#1155CC')
    .setBackground('#E8F0FE');
  return buttons.length;
}

/** Chạy qua menu để khôi phục ngay các nút BBNT đã mất. */
function khoiPhucNutFileCungBbnt() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.BBNT_SHEET_NAME);
  if (!sheet) throw new Error('Không tìm thấy tab ' + CONFIG.BBNT_SHEET_NAME + '.');
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    throw new Error('Hệ thống đang xử lý yêu cầu khác. Hãy thử lại.');
  }
  try {
    const count = capNhatNutTaiBbnt_(sheet);
    SpreadsheetApp.flush();
    ss.toast('Đã khôi phục ' + count + ' nút FILE CỨNG ở tab BBNT.', 'Thành công', 6);
  } finally {
    lock.releaseLock();
  }
}

function columnToLetter_(column) {
  let letters = '';
  while (column > 0) {
    column--;
    letters = String.fromCharCode(65 + column % 26) + letters;
    column = Math.floor(column / 26);
  }
  return letters;
}

function xuLyTuDongDienBbnt_(sheet, actionCell) {
  const row = actionCell.getRow();
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    actionCell.setValue(false);
    actionCell.setNote('Hệ thống đang xử lý yêu cầu khác. Hãy thử lại.');
    return;
  }
  try {
    actionCell.setNote('Đang tự động điền BBNT...');
    const result = dienBbntTuHopDong_(sheet, row);
    const message = result.created
      ? 'Đã tạo dòng ' + result.row + ' ở tab BBNT.'
      : 'Đã bổ sung ' + result.filled + ' ô ở dòng ' + result.row + ' của tab BBNT.';
    actionCell.setNote(message);
    sheet.getParent().toast(message, 'Tự động điền BBNT', 6);
  } catch (error) {
    actionCell.setNote('Không thể tự động điền BBNT: ' + error.message);
    console.error('Dòng ' + row + ': ' + error.stack);
    sheet.getParent().toast(error.message, 'Không thể điền BBNT', 10);
  } finally {
    actionCell.setValue(false);
    lock.releaseLock();
  }
}

function xuLyTaoHopDong(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();

  if (sheet.getName() !== CONFIG.SHEET_NAME) return;
  if (!coBoCucHopDongMoi_(sheet)) return;

  // Hỗ trợ sửa hoặc dán nhiều dòng có chứa cột C.
  if (
    range.getColumn() <= CONFIG.APPLICATION_DATE_COLUMN &&
    range.getLastColumn() >= CONFIG.APPLICATION_DATE_COLUMN
  ) {
    for (let row = Math.max(2, range.getRow()); row <= range.getLastRow(); row++) {
      capNhatNgayHoanThanh_(sheet, row);
    }
  }

  // Hỗ trợ cả khi dán nhiều dòng hoặc nhiều cột có chứa T.
  if (
    range.getColumn() <= CONFIG.AMOUNT_COLUMN &&
    range.getLastColumn() >= CONFIG.AMOUNT_COLUMN
  ) {
    for (let row = Math.max(2, range.getRow()); row <= range.getLastRow(); row++) {
      capNhatTienThucNhanBangChu_(sheet, row);
    }
  }

  if (
    range.getColumn() === CONFIG.BBNT_ACTION_COLUMN &&
    range.getRow() >= 2 &&
    range.getNumRows() === 1 &&
    range.getNumColumns() === 1 &&
    e.value === 'TRUE'
  ) {
    xuLyTuDongDienBbnt_(sheet, range);
    return;
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
  kiemTraBoCucHopDong_(sheet);
  xoaCongThucSttCu_(sheet);

  // A:W là dữ liệu, X:Z là ba link đã tạo.
  const displayValues = sheet.getRange(row, 1, 1, CONFIG.FIRST_LINK_COLUMN + CONFIG.LINK_COUNT - 1)
    .getDisplayValues()[0];
  const rawValues = sheet.getRange(row, 1, 1, CONFIG.FIRST_LINK_COLUMN - 1).getValues()[0];

  const ngayLamDonHienThi = String(displayValues[2] || '').trim();

  if (!ngayLamDonHienThi) {
    throw new Error('Thiếu Ngày làm đơn ở cột C.');
  }

  const ngayKy = tachNgay_(ngayLamDonHienThi);
  capNhatNgayHoanThanh_(sheet, row);
  displayValues[CONFIG.COMPLETION_DATE_COLUMN - 1] = sheet
    .getRange(row, CONFIG.COMPLETION_DATE_COLUMN).getDisplayValue();

  const headers = [
    'Số hợp đồng',
    'Họ và tên người đề nghị',
    'Ngày làm đơn',
    'Số CCCD',
    'Mã số thuế',
    'Địa chỉ',
    'Số điện thoại',
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
    throw new Error('Số tiền thực nhận ở cột T phải là số nguyên lớn hơn 0 (ví dụ: 6.750.000).');
  }
  const soTienThucNhanBangChu = vietSoTienBangChu_(soTienThucNhan);
  sheet.getRange(row, CONFIG.AMOUNT_WORDS_COLUMN).setValue(soTienThucNhanBangChu);
  displayValues[CONFIG.AMOUNT_WORDS_COLUMN - 1] = soTienThucNhanBangChu;
  const { tongThuLao, thueTncn } = tinhThuLaoTruocThue_(soTienThucNhan);

  // Cột A do code tự sinh; bắt buộc người dùng nhập đầy đủ B:W.
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
  const ngayHoanTat = tachNgay_(displayValues[18]);
  let soHopDong = String(displayValues[0] || '').trim();

  // Giữ số đã dùng trong Word của dòng này, kể cả số theo định dạng cũ.
  // Dòng chưa có link luôn được cấp số từ các hợp đồng đã tạo.
  if (!currentRowHasLink || !soHopDong) {
    soHopDong = taoSoHopDongMoi_(sheet, row);
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
    stt + ' - ' + hoTen + ' - HĐ khoán việc - ' + lanTao
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


// ========================================
// 4. SINH SỐ HỢP ĐỒNG
// ========================================

/**
 * Cấp số nguyên liên tiếp từ 3000.
 * Chỉ các dòng có ít nhất một link Word tại X:Z mới được tính.
 */
function taoSoHopDongMoi_(sheet, currentRow) {
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const existingRows = sheet.getRange(
    2, 1, lastRow - 1, CONFIG.FIRST_LINK_COLUMN + CONFIG.LINK_COUNT - 1
  )
    .getDisplayValues();
  let maxSequence = 2999;

  existingRows.forEach((values, index) => {
    const sheetRow = index + 2;

    if (sheetRow === currentRow) return;

    const hasWordLink = values
      .slice(CONFIG.FIRST_LINK_COLUMN - 1, CONFIG.FIRST_LINK_COLUMN - 1 + CONFIG.LINK_COUNT)
      .some(link => String(link || '').trim() !== '');

    if (!hasWordLink) return;

    const soHopDong = String(values[0] || '').trim();

    if (/^\d+$/.test(soHopDong)) {
      maxSequence = Math.max(maxSequence, Number(soHopDong));
    }
  });

  return String(maxSequence + 1);
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

/** Khóa link Word, nút FILE CỨNG AC và cột link theo tiêu đề. */
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
      range: sheet.getRange(2, CONFIG.HARD_UPLOAD_COLUMN, maxRows - 1, 1),
      description: CONFIG.HARD_FILE_PROTECTION_DESCRIPTION
    },
    {
      range: sheet.getRange(2, layCotLinkFileCung_(sheet, false), maxRows - 1, 1),
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
  '{{THUE_TNCN}}',
  '{{SO_TIEN_THUC_NHAN}}',
  '{{SO_TIEN_THUC_NHAN_BANG_CHU}}',
  '{{SO_TAI_KHOAN}}',
  '{{MO_TAI}}'
];
