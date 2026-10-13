/**
 * TỰ ĐỘNG TẠO GIẤY ĐỀ NGHỊ THANH TOÁN
 *
 * A:I - Dữ liệu
 * J:L - Link Word lần 01, 02, 03
 * M   - Checkbox tạo Word
 * N   - Tải PDF Minh Chứng Bill
 * O:Q - Link Bill lần 01, 02, 03
 * R   - Tải PDF FILE CỨNG
 * S   - Link PDF FILE CỨNG
 * T   - Trạng thái nghiệp vụ (không do code này xử lý)
 * U   - Email người tải Bill đầu tiên (cột ẩn)
 */

const CONFIG = {
  SHEET_NAME: 'TH_GIẤY ĐỀ NGHỊ THANH TOÁN',

  TEMPLATE_URL: 'https://docs.google.com/document/d/1SImIGq8QvIl-k5wIdiNsB7LvTbZw43_WHqgcLf1LzkE/edit?pli=1&tab=t.0',

  FOLDER_ID: '1tOWPSQzZY1EOfvsR-VQr_AHCjHp2igdr',

  // Sau khi triển khai Web App, dán URL kết thúc bằng /exec vào đây.
  WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbwtpKgb-wKnZ9cYUX5w-kfnAoP0wgjY_SGDv2UgPSCI3wJUC6GehXL3eIOwgv-Wjso/exec',

  FIRST_LINK_COLUMN: 10, // J
  LINK_COUNT: 3,
  ACTION_COLUMN: 13,    // M
  LINK_PROTECTION_DESCRIPTION: 'KHOA_LINK_WORD_TU_DONG',
  HARD_FILE_PROTECTION_DESCRIPTION: 'KHOA_FILE_CUNG_TU_DONG',

  BILL_UPLOAD_COLUMN: 14,          // N
  BILL_FIRST_LINK_COLUMN: 15,      // O
  BILL_LINK_COUNT: 3,              // O, P, Q
  HARD_UPLOAD_COLUMN: 18,          // R
  HARD_LINK_COLUMN: 19,            // S
  BILL_UPLOADER_EMAIL_COLUMN: 21,  // U - cột ẩn
  UPLOAD_MAX_SIZE_MB: 10
};


// ========================================
// 0. TRANG WEB UPLOAD KHI BẤM NÚT Ở N/R
// ========================================

function doGet(e) {
  const loai = String(e && e.parameter && e.parameter.loai || '');
  const row = Number(e && e.parameter && e.parameter.row || 0);

  if (!['BILL', 'FILE_CUNG'].includes(loai) || !Number.isInteger(row) || row < 2) {
    return HtmlService.createHtmlOutput(
      '<h3>Đường link upload không hợp lệ.</h3>'
    );
  }

  // FILE CỨNG chỉ chủ sở hữu bảng tính được mở trang tải lên.
  if (loai === 'FILE_CUNG') {
    const ss = laySpreadsheet_();
    const activeEmail = layEmailNguoiDung_();
    const ownerEmail = layEmailChuSoHuu_(ss);

    if (activeEmail !== ownerEmail) {
      return HtmlService.createHtmlOutput(
        '<h3>Bạn không có quyền tải FILE CỨNG.</h3>' +
        '<p>Chỉ chủ sở hữu bảng tính được thực hiện thao tác này.</p>'
      );
    }
  }

  const template = HtmlService.createTemplateFromFile('UploadPdf');
  template.loai = loai;
  template.row = row;
  template.tieuDe = loai === 'BILL'
    ? 'Minh Chứng Bill'
    : 'FILE CỨNG';
  template.maxSizeMb = CONFIG.UPLOAD_MAX_SIZE_MB;
  template.webMode = true;

  return template.evaluate()
    .setTitle('Tải tệp - ' + template.tieuDe + ' - Dòng ' + row);
}


// ========================================
// 1. TẠO MENU
// ========================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tạo giấy thanh toán')
    .addItem('Cài đặt nút tạo Word', 'caiDatHeThong')
    .addToUi();

  SpreadsheetApp.getUi()
    .createMenu('Tải minh chứng')
    .addItem('Tải Minh Chứng Bill', 'moTaiMinhChungBill')
    .addItem('Tải FILE CỨNG', 'moTaiFileCung')
    .addToUi();
}


// ========================================
// 2. CÀI ĐẶT CỘT LINK, CHECKBOX VÀ TRIGGER
// Chạy lại một lần sau khi thay code
// ========================================

function caiDatHeThong() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error('Không tìm thấy tab: ' + CONFIG.SHEET_NAME);
  }

  const templateId = layIdTaiLieu(CONFIG.TEMPLATE_URL);
  const template = DriveApp.getFileById(templateId);

  if (template.getMimeType() !== 'application/vnd.google-apps.document') {
    throw new Error('File mẫu phải là Google Docs.');
  }

  DocumentApp.openById(templateId).getBody();
  DriveApp.getFolderById(CONFIG.FOLDER_ID).getName();

  PropertiesService.getScriptProperties().setProperty(
    'SPREADSHEET_ID',
    ss.getId()
  );

  // J, K, L là ba lần tạo; M là nút tạo.
  sheet.getRange('J1:L1').setValues([[
    'Link_Word_01',
    'Link_Word_02',
    'Link_Word_03'
  ]]);
  sheet.getRange('M1').setValue('TẠO WORD');
  sheet.getRange('N1:Q1').setValues([[
    'Minh Chứng Bill',
    'Link_Bill_01',
    'Link_Bill_02',
    'Link_Bill_03'
  ]]);
  sheet.getRange('R1:S1').setValues([[
    'FILE CỨNG',
    'Link_FILE_Cung'
  ]]);
  sheet.getRange('U1').setValue('Email_Upload_Bill');
  sheet.hideColumns(CONFIG.BILL_UPLOADER_EMAIL_COLUMN);

  const maxRows = sheet.getMaxRows();

  if (maxRows >= 2) {
    const checkboxRange = sheet.getRange(
      2,
      CONFIG.ACTION_COLUMN,
      maxRows - 1,
      1
    );

    checkboxRange.insertCheckboxes();
    checkboxRange.setHorizontalAlignment('center');
    checkboxRange.setBackground('#E8F5E9');

    khoaVungLink(sheet, maxRows);
    taoNutUploadPdf_(sheet, maxRows);
  }

  sheet.setColumnWidths(CONFIG.FIRST_LINK_COLUMN, CONFIG.LINK_COUNT, 180);
  sheet.setColumnWidth(CONFIG.ACTION_COLUMN, 120);
  sheet.setColumnWidths(CONFIG.BILL_UPLOAD_COLUMN, 6, 180);

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'xuLyTaoWord') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('xuLyTaoWord')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ss.toast(
    coLinkWebAppHopLe_()
      ? 'Đã tạo nút upload PDF tại cột N và R.'
      : 'Chưa có WEB_APP_URL. Hãy triển khai Web App rồi chạy cài đặt lại.',
    'CÀI ĐẶT THÀNH CÔNG',
    8
  );
}

function coLinkWebAppHopLe_() {
  return /^https:\/\/script\.google\.com\/macros\/s\/.+\/exec$/
    .test(String(CONFIG.WEB_APP_URL || '').trim());
}

function taoNutUploadPdf_(sheet, maxRows) {
  if (!coLinkWebAppHopLe_() || maxRows < 2) return;

  const billButtons = [];
  const hardFileButtons = [];
  const baseUrl = String(CONFIG.WEB_APP_URL).trim();

  for (let row = 2; row <= maxRows; row++) {
    const billUrl = baseUrl + '?loai=BILL&row=' + row;
    const hardFileUrl = baseUrl + '?loai=FILE_CUNG&row=' + row;

    billButtons.push([
      SpreadsheetApp.newRichTextValue()
        .setText('⬆ TẢI BILL')
        .setLinkUrl(billUrl)
        .build()
    ]);

    hardFileButtons.push([
      SpreadsheetApp.newRichTextValue()
        .setText('⬆ TẢI FILE CỨNG')
        .setLinkUrl(hardFileUrl)
        .build()
    ]);
  }

  const billRange = sheet.getRange(
    2,
    CONFIG.BILL_UPLOAD_COLUMN,
    maxRows - 1,
    1
  );
  const hardFileRange = sheet.getRange(
    2,
    CONFIG.HARD_UPLOAD_COLUMN,
    maxRows - 1,
    1
  );

  billRange.setRichTextValues(billButtons);
  hardFileRange.setRichTextValues(hardFileButtons);

  [billRange, hardFileRange].forEach(range => {
    range
      .setHorizontalAlignment('center')
      .setFontWeight('bold')
      .setFontColor('#1155CC')
      .setBackground('#E8F0FE');
  });
}


// ========================================
// 3. CỬA SỔ TẢI PDF / JPG / PNG TỪ MÁY TÍNH HOẶC ĐIỆN THOẠI
// N và O:Q: Minh Chứng Bill, tối đa ba lần
// R và S: FILE CỨNG, một lần
// Word và PDF của cùng một dòng được lưu chung trong thư mục STT.
// ========================================

function moTaiMinhChungBill() {
  moHopThoaiTaiPdf_('BILL');
}

function moTaiFileCung() {
  moHopThoaiTaiPdf_('FILE_CUNG');
}

function moHopThoaiTaiPdf_(loai) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  if (
    loai === 'FILE_CUNG' &&
    layEmailNguoiDung_() !== layEmailChuSoHuu_(ss)
  ) {
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
    ui.alert('Hãy chọn một ô tại dòng dữ liệu cần tải PDF.');
    return;
  }

  const template = HtmlService.createTemplateFromFile('UploadPdf');
  template.loai = loai;
  template.row = row;
  template.tieuDe = loai === 'BILL'
    ? 'Minh Chứng Bill'
    : 'FILE CỨNG';
  template.maxSizeMb = CONFIG.UPLOAD_MAX_SIZE_MB;
  template.webMode = false;

  const html = template.evaluate()
    .setWidth(440)
    .setHeight(330);

  ui.showModalDialog(
    html,
    'Tải tệp - ' + template.tieuDe + ' - Dòng ' + row
  );
}

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

  if (!['BILL', 'FILE_CUNG'].includes(loai)) {
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

  const activeEmail = layEmailNguoiDung_();
  const ownerEmail = layEmailChuSoHuu_(ss);
  let linkColumn;
  let lanTai;
  let billUploaderEmailToSave = '';

  if (loai === 'BILL') {
    const billLinks = sheet.getRange(
      row,
      CONFIG.BILL_FIRST_LINK_COLUMN,
      1,
      CONFIG.BILL_LINK_COUNT
    ).getDisplayValues()[0];
    const emptyBillIndex = billLinks.findIndex(
      link => String(link || '').trim() === ''
    );

    if (emptyBillIndex === -1) {
      throw new Error('Dòng này đã tải đủ 3 Minh Chứng Bill.');
    }

    const emailCell = sheet.getRange(
      row,
      CONFIG.BILL_UPLOADER_EMAIL_COLUMN
    );
    const firstUploaderEmail = String(
      emailCell.getDisplayValue() || ''
    ).trim().toLowerCase();

    if (
      firstUploaderEmail &&
      activeEmail !== firstUploaderEmail &&
      activeEmail !== ownerEmail
    ) {
      throw new Error(
        'Chỉ ' + firstUploaderEmail +
        ' được tải Bill lần tiếp theo cho dòng này.'
      );
    }

    if (!firstUploaderEmail) {
      billUploaderEmailToSave = activeEmail;
    }

    linkColumn = CONFIG.BILL_FIRST_LINK_COLUMN + emptyBillIndex;
    lanTai = String(emptyBillIndex + 1).padStart(2, '0');
  } else {
    if (activeEmail !== ownerEmail) {
      throw new Error(
        'Chỉ chủ sở hữu bảng tính được tải FILE CỨNG.'
      );
    }

    const currentHardFileLink = String(
      sheet.getRange(row, CONFIG.HARD_LINK_COLUMN).getDisplayValue() || ''
    ).trim();

    if (currentHardFileLink) {
      throw new Error('Dòng này đã có FILE CỨNG.');
    }

    linkColumn = CONFIG.HARD_LINK_COLUMN;
    lanTai = '01';
  }

  const stt = String(sheet.getRange(row, 1).getDisplayValue() || '').trim();

  if (!stt) {
    throw new Error('Dòng ' + row + ' chưa có STT ở cột A.');
  }

  const mainFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  const targetFolder = layHoacTaoThuMucStt_(mainFolder, stt);
  const safeName = originalName.replace(/[\\/:*?"<>|]/g, '_');
  const blob = Utilities.newBlob(bytes, mimeType, safeName);
  const file = targetFolder.createFile(blob);

  file.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  const viewUrl =
    'https://drive.google.com/file/d/' + file.getId() + '/view';

  if (loai === 'BILL') {
    sheet.getRange(row, CONFIG.BILL_UPLOAD_COLUMN).setNote(
      'Đã tải ' + lanTai + '/03'
    );

    if (billUploaderEmailToSave) {
      sheet.getRange(
        row,
        CONFIG.BILL_UPLOADER_EMAIL_COLUMN
      ).setValue(billUploaderEmailToSave);
    }
  } else {
    sheet.getRange(row, CONFIG.HARD_UPLOAD_COLUMN).setNote(
      'Đã tải: ' + safeName
    );
  }
  sheet.getRange(row, linkColumn).setValue(viewUrl);
  SpreadsheetApp.flush();

  return {
    fileName: safeName,
    url: viewUrl,
    row: row,
    lanTai: lanTai
  };
}

function laySpreadsheet_() {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (activeSpreadsheet) {
    return activeSpreadsheet;
  }

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

  if (/^\d+$/.test(text)) {
    return text.padStart(2, '0');
  }

  return text.replace(/[\\/:*?"<>|]/g, '_');
}


// ========================================
// 4. KHÓA VÙNG LINK J:L VÀ FILE CỨNG R:S
// Chủ sở hữu bảng tính vẫn có thể chỉnh sửa.
// ========================================

function khoaVungLink(sheet, maxRows) {
  const protectionDescriptions = [
    CONFIG.LINK_PROTECTION_DESCRIPTION,
    CONFIG.HARD_FILE_PROTECTION_DESCRIPTION
  ];

  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter(protection =>
      protectionDescriptions.includes(protection.getDescription())
    )
    .forEach(protection => protection.remove());

  const wordLinkRange = sheet.getRange(
    2,
    CONFIG.FIRST_LINK_COLUMN,
    maxRows - 1,
    CONFIG.LINK_COUNT
  );

  // Xóa quy tắc checkbox cũ nếu trước đây cột L từng là nút tạo.
  wordLinkRange.clearDataValidations();

  const protectedRanges = [
    {
      range: wordLinkRange,
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
// 5. LẤY ID TỪ LINK GOOGLE DOCS
// ========================================

function layIdTaiLieu(url) {
  const text = String(url || '').trim();
  const match = text.match(
    /docs\.google\.com\/document\/d\/([a-zA-Z0-9_-]+)/
  );

  if (!match) {
    throw new Error(
      'TEMPLATE_URL chưa đúng. Hãy dán toàn bộ URL Google Docs mẫu.'
    );
  }

  return match[1];
}


// ========================================
// 6. TÁCH NGÀY / THÁNG / NĂM
// ========================================

function tachNgay(dateText) {
  const text = String(dateText || '').trim();
  const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);

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


// ========================================
// 7. XỬ LÝ KHI TÍCH CHECKBOX Ở CỘT M
// ========================================

function xuLyTaoWord(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();

  if (sheet.getName() !== CONFIG.SHEET_NAME) return;

  if (
    e.range.getColumn() !== CONFIG.ACTION_COLUMN ||
    e.range.getNumRows() !== 1 ||
    e.range.getNumColumns() !== 1 ||
    e.range.getRow() < 2 ||
    e.value !== 'TRUE'
  ) {
    return;
  }

  const row = e.range.getRow();
  const actionCell = e.range;
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    actionCell.setValue(false);
    actionCell.setNote('Hệ thống đang bận. Hãy thử lại.');
    return;
  }

  try {
    actionCell.setNote('Đang xử lý...');

    const result = taoTaiLieuChoDong(sheet, row);

    actionCell.setNote(
      'Đã tạo tài liệu lần ' + result.lanTao + ' thành công.'
    );

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Đã tạo tài liệu lần ' + result.lanTao + ' cho dòng ' + row,
      'Thành công',
      5
    );
  } catch (error) {
    actionCell.setNote('Không thể tạo Word: ' + error.message);
    console.error('Dòng ' + row + ': ' + error.stack);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      error.message,
      'Không thể tạo Word - Dòng ' + row,
      10
    );
  } finally {
    actionCell.setValue(false);
    lock.releaseLock();
  }
}


// ========================================
// 8. KIỂM TRA VÀ TẠO TÀI LIỆU
// ========================================

function taoTaiLieuChoDong(sheet, row) {
  // A:I là dữ liệu; J:L là ba link hiện có.
  const values = sheet.getRange(row, 1, 1, 12).getDisplayValues()[0];

  const headers = [
    'STT',
    'Họ và tên người đề nghị',
    'Ngày làm đơn',
    'Chức vụ',
    'Bộ phận (Phòng/Ban)',
    'Số tiền thanh toán',
    'Số tiền thanh toán (Bằng Chữ)',
    'Lý do',
    'Thời hạn thanh toán tạm ứng'
  ];

  const errorValues = [
    '#ERROR!', '#REF!', '#VALUE!',
    '#N/A', '#DIV/0!', '#NUM!', '#NAME?'
  ];
  const missing = [];

  for (let i = 0; i < 9; i++) {
    const value = String(values[i] ?? '').trim();

    if (!value || errorValues.some(errorValue => value.startsWith(errorValue))) {
      missing.push(headers[i]);
    }
  }

  if (missing.length > 0) {
    throw new Error('Thiếu hoặc lỗi dữ liệu: ' + missing.join(', '));
  }

  const [
    stt,
    hoTen,
    ngayLamDon,
    chucVu,
    boPhan,
    soTien,
    bangChu,
    lyDo,
    thoiHan
  ] = values;

  const links = values.slice(
    CONFIG.FIRST_LINK_COLUMN - 1,
    CONFIG.FIRST_LINK_COLUMN - 1 + CONFIG.LINK_COUNT
  );
  const emptyLinkIndex = links.findIndex(link => String(link).trim() === '');

  if (emptyLinkIndex === -1) {
    throw new Error('Dòng này đã tạo đủ 3 tài liệu.');
  }

  // Lần tạo 1, 2, 3 tương ứng 01, 02, 03.
  const lanTao = String(emptyLinkIndex + 1).padStart(2, '0');
  const targetLinkColumn = CONFIG.FIRST_LINK_COLUMN + emptyLinkIndex;

  const amount = sheet.getRange(row, 6).getValue();

  if (
    typeof amount !== 'number' ||
    !Number.isFinite(amount) ||
    amount <= 0
  ) {
    throw new Error('Số tiền thanh toán phải lớn hơn 0.');
  }

  const ngayDon = tachNgay(ngayLamDon);
  const ngayHan = tachNgay(thoiHan);
  const templateId = layIdTaiLieu(CONFIG.TEMPLATE_URL);
  const template = DriveApp.getFileById(templateId);

  const fileName =
    'Giay de nghi thanh toan - ' +
    stt + ' - ' + hoTen + ' - ' + lanTao;

  const mainFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  const sttFolder = layHoacTaoThuMucStt_(mainFolder, stt);

  const copy = template.makeCopy(fileName, sttFolder);

  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  const replacements = {
    '{{STT}}': stt,
    '{{HO_TEN}}': hoTen,
    '{{NGAY}}': ngayDon.ngay,
    '{{THANG}}': ngayDon.thang,
    '{{NAM}}': ngayDon.nam,
    '{{NGAY_LAM_DON}}': ngayLamDon,
    '{{CHUC_VU}}': chucVu,
    '{{BO_PHAN}}': boPhan,
    '{{SO_TIEN}}': soTien,
    '{{BANG_CHU}}': bangChu,
    '{{LY_DO}}': lyDo,
    '{{THOI_HAN}}': thoiHan,
    '{{NGAY_HAN}}': ngayHan.ngay,
    '{{THANG_HAN}}': ngayHan.thang,
    '{{NAM_HAN}}': ngayHan.nam
  };

  Object.entries(replacements).forEach(([key, value]) => {
    const escaped = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    body.replaceText(escaped, String(value));
  });

  doc.saveAndClose();

  const finishedDoc = DocumentApp.openById(copy.getId());
  const remaining = finishedDoc.getBody()
    .getText()
    .match(/\{\{[^{}]+\}\}/g);

  if (remaining && remaining.length > 0) {
    throw new Error(
      'Tài liệu còn biến chưa được thay: ' +
      [...new Set(remaining)].join(', ') +
      '. Bản sao đã được tạo nhưng chưa cập nhật Sheet.'
    );
  }

  // Ai có link cũng chỉ được xem.
  copy.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  const viewUrl =
    'https://docs.google.com/document/d/' + copy.getId() + '/view';

  // Ghi link vào đúng J/K/L.
  sheet.getRange(row, targetLinkColumn).setValue(viewUrl);

  SpreadsheetApp.flush();

  return {
    lanTao: lanTao,
    url: viewUrl
  };
}