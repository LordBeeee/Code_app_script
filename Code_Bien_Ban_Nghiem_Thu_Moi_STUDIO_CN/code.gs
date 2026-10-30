/**
 * TỰ ĐỘNG TẠO BIÊN BẢN NGHIỆM THU
 *
 * A   - Số Biên Bản Nghiệm Thu (tự lấy bằng cột B)
 * B:L - Dữ liệu bắt buộc
 * M:O - Link Word lần 01, 02, 03
 * P   - Checkbox TẠO WORD
 * Q   - Tải FILE CỨNG
 * R   - Link_FILE_Cung
 *
 * Tên hàm và cấu hình dùng tiền tố BBNT để không xung đột với các
 * script Hợp đồng khoán và Giấy đề nghị thanh toán trong cùng dự án.
 */

const BBNT_CONFIG = {
  SHEET_NAME: 'BBNT STUDIO CN',

  SPREADSHEET_URL:
    'https://docs.google.com/spreadsheets/d/1uFp7McsEUjFQzMsQrW0ZSiqxQe7324Wy6KMzsHP_Mq0/edit',

  TEMPLATE_URL:
    'https://docs.google.com/document/d/18i8Y3PeB8wFEzbC0Yapw_UFVgg_3yJR04I7S7zDaq00/edit?tab=t.0',

  FOLDER_URL:
    'https://drive.google.com/drive/folders/1sP4YJ8EVSNlq2D22M_J5ST2LXDjZ9ZKv?hl=vi',

  // Có thể để trống: code tự lấy URL Web App đã triển khai.
  WEB_APP_URL: 'https://script.google.com/macros/library/d/14AByJdWt0ypnpW2_JHazZ-S74d8iLr4jpi0MMSxjjgHZMrOvk-cODP3B/1',

  FIRST_LINK_COLUMN: 13, // M
  LINK_COUNT: 3,         // M, N, O
  ACTION_COLUMN: 16,     // P
  LINK_PROTECTION_DESCRIPTION: 'KHOA_LINK_BBNT_TU_DONG',
  HARD_FILE_PROTECTION_DESCRIPTION: 'KHOA_FILE_CUNG_BBNT_TU_DONG',
  HARD_UPLOAD_COLUMN: 17, // Q
  HARD_LINK_COLUMN: 18,   // R
  UPLOAD_MAX_SIZE_MB: 10
};


// ========================================
// 0. FILE CỨNG: CÙNG GIAO DIỆN VỚI ĐỀ NGHỊ THANH TOÁN
// ========================================

function bbntLayUrlWebApp_() {
  const url = String(
    BBNT_CONFIG.WEB_APP_URL || ScriptApp.getService().getUrl() || ''
  ).trim();

  return /^https:\/\/script\.google\.com\/(?:macros\/s|a\/macros\/[^/]+\/s)\/[^/?#]+\/exec(?:[?#].*)?$/
    .test(url) ? url : '';
}

function bbntTaoNutUpload_(sheet, maxRows) {
  const baseUrl = bbntLayUrlWebApp_();
  if (!baseUrl || maxRows < 2) return;

  const buttons = [];
  for (let row = 2; row <= maxRows; row++) {
    buttons.push([
      SpreadsheetApp.newRichTextValue()
        .setText('⬆ TẢI FILE CỨNG')
        .setLinkUrl(baseUrl + '?loai=FILE_CUNG&row=' + row)
        .build()
    ]);
  }

  sheet.getRange(2, BBNT_CONFIG.HARD_UPLOAD_COLUMN, maxRows - 1, 1)
    .setRichTextValues(buttons)
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setFontColor('#1155CC')
    .setBackground('#E8F0FE');
}

function doGet(e) {
  const loai = String(e && e.parameter && e.parameter.loai || '');
  const row = Number(e && e.parameter && e.parameter.row || 0);

  if (loai !== 'FILE_CUNG' || !Number.isInteger(row) || row < 2) {
    return HtmlService.createHtmlOutput('<h3>Đường link upload không hợp lệ.</h3>');
  }

  const ss = bbntLaySpreadsheet_();
  const owner = DriveApp.getFileById(ss.getId()).getOwner();
  const ownerEmail = owner ? String(owner.getEmail()).trim().toLowerCase() : '';
  const activeEmail = String(Session.getActiveUser().getEmail() || '')
    .trim().toLowerCase();

  if (!activeEmail || activeEmail !== ownerEmail) {
    return HtmlService.createHtmlOutput(
      '<h3>Chỉ chủ sở hữu bảng tính được tải FILE CỨNG.</h3>'
    );
  }

  const template = HtmlService.createTemplateFromFile('UploadPDF');
  template.loai = loai;
  template.row = row;
  template.tieuDe = 'FILE CỨNG';
  template.maxSizeMb = BBNT_CONFIG.UPLOAD_MAX_SIZE_MB;
  template.webMode = true;

  return template.evaluate()
    .setTitle('Tải FILE CỨNG - Dòng ' + row);
}

/** Hàm được gọi từ UploadPDF.html. */
function luuPdfTuMayTinh(payload) {
  if (!payload || !payload.base64) {
    throw new Error('Không nhận được dữ liệu tệp.');
  }

  const row = Number(payload.row);
  const originalName = String(payload.fileName || '').trim();
  const mimeType = String(payload.mimeType || '').toLowerCase();
  const lowerName = originalName.toLowerCase();
  const validTypes = ['application/pdf', 'image/jpeg', 'image/png'];
  const validExtensions = ['.pdf', '.jpg', '.jpeg', '.png'];

  if (payload.loai !== 'FILE_CUNG' || !Number.isInteger(row) || row < 2) {
    throw new Error('Yêu cầu tải tệp không hợp lệ.');
  }

  if (!validTypes.includes(mimeType) ||
      !validExtensions.some(extension => lowerName.endsWith(extension))) {
    throw new Error('Chỉ được tải lên PDF, JPG hoặc PNG.');
  }

  const bytes = Utilities.base64Decode(payload.base64);
  if (bytes.length > BBNT_CONFIG.UPLOAD_MAX_SIZE_MB * 1024 * 1024) {
    throw new Error('Tệp vượt quá ' + BBNT_CONFIG.UPLOAD_MAX_SIZE_MB + ' MB.');
  }

  const ss = bbntLaySpreadsheet_();
  const owner = DriveApp.getFileById(ss.getId()).getOwner();
  const ownerEmail = owner ? String(owner.getEmail()).trim().toLowerCase() : '';
  const activeEmail = String(Session.getActiveUser().getEmail() || '')
    .trim().toLowerCase();

  if (!activeEmail || activeEmail !== ownerEmail) {
    throw new Error('Chỉ chủ sở hữu bảng tính được tải FILE CỨNG.');
  }

  const sheet = ss.getSheetByName(BBNT_CONFIG.SHEET_NAME);
  if (!sheet || row > sheet.getMaxRows()) {
    throw new Error('Không tìm thấy dòng biên bản nghiệm thu.');
  }

  const linkCell = sheet.getRange(row, BBNT_CONFIG.HARD_LINK_COLUMN);
  if (String(linkCell.getDisplayValue() || '').trim()) {
    throw new Error('Dòng này đã có FILE CỨNG.');
  }

  const soBienBan = String(sheet.getRange(row, 1).getDisplayValue() || '').trim();
  if (!soBienBan) {
    throw new Error('Dòng ' + row + ' chưa có Số Biên Bản Nghiệm Thu ở cột A.');
  }

  const folderId = bbntLayIdFolder_(BBNT_CONFIG.FOLDER_URL);
  const mainFolder = DriveApp.getFolderById(folderId);
  const targetFolder = bbntLayHoacTaoThuMuc_(mainFolder, soBienBan);
  const safeName = bbntLamSachTenFile_(originalName);
  const blob = Utilities.newBlob(bytes, mimeType, safeName);
  const file = targetFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const viewUrl = 'https://drive.google.com/file/d/' + file.getId() + '/view';
  linkCell.setValue(viewUrl);
  sheet.getRange(row, BBNT_CONFIG.HARD_UPLOAD_COLUMN)
    .setNote('Đã tải: ' + safeName);
  SpreadsheetApp.flush();

  return { fileName: safeName, url: viewUrl, row: row, lanTai: '01' };
}


// ========================================
// 1. MENU VÀ CÀI ĐẶT
// ========================================

function bbntOnOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tạo biên bản nghiệm thu')
    .addItem('Cài đặt hệ thống', 'caiDatBienBanNghiemThu')
    .addItem(
      'Tạo biên bản cho dòng đang chọn',
      'taoBienBanNghiemThuDongDangChon'
    )
    .addToUi();
}

/** Chạy hàm này một lần sau khi dán code. */
function caiDatBienBanNghiemThu() {
  const ss = bbntLaySpreadsheet_();
  const sheet = ss.getSheetByName(BBNT_CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error('Không tìm thấy tab: ' + BBNT_CONFIG.SHEET_NAME);
  }
  bbntKiemTraBoCuc_(sheet);

  const templateId = bbntLayIdGoogleDoc_(BBNT_CONFIG.TEMPLATE_URL);
  const folderId = bbntLayIdFolder_(BBNT_CONFIG.FOLDER_URL);
  const template = DriveApp.getFileById(templateId);

  if (template.getMimeType() !== MimeType.GOOGLE_DOCS) {
    throw new Error('File mẫu phải là Google Docs.');
  }

  // Kiểm tra quyền truy cập trước khi tạo trigger.
  DocumentApp.openById(templateId).getBody();
  DriveApp.getFolderById(folderId).getName();

  sheet.getRange('A1').setValue('Số Biên Bản Nghiệm Thu');
  sheet.getRange('M1:P1').setValues([[
    'Link_Word_01',
    'Link_Word_02',
    'Link_Word_03',
    'TẠO WORD'
  ]]);
  sheet.getRange('Q1:R1').setValues([[
    'FILE CỨNG',
    'Link_FILE_Cung'
  ]]);

  const maxRows = sheet.getMaxRows();

  if (maxRows >= 2) {
    sheet.getRange(2, BBNT_CONFIG.ACTION_COLUMN, maxRows - 1, 1)
      .insertCheckboxes()
      .setHorizontalAlignment('center')
      .setBackground('#E8F5E9');

    bbntKhoaCotLink_(sheet, maxRows);
    bbntTaoNutUpload_(sheet, maxRows);
  }

  sheet.setColumnWidths(
    BBNT_CONFIG.FIRST_LINK_COLUMN,
    BBNT_CONFIG.LINK_COUNT,
    180
  );
  sheet.setColumnWidth(BBNT_CONFIG.ACTION_COLUMN, 120);
  sheet.setColumnWidths(BBNT_CONFIG.HARD_UPLOAD_COLUMN, 2, 180);

  ScriptApp.getProjectTriggers()
    .filter(trigger =>
      trigger.getHandlerFunction() === 'xuLyTaoBienBanNghiemThu'
    )
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('xuLyTaoBienBanNghiemThu')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  bbntThongBao_(
    ss,
    bbntLayUrlWebApp_()
       ? 'Đã tạo checkbox P và nút FILE CỨNG ở Q.'
       : 'Đã cài biên bản. Triển khai Web App rồi chạy cài đặt lại để hiện nút Q.',
    'Cài đặt biên bản nghiệm thu thành công',
    7
  );
}

/** Ngăn ghi nhầm cột khi tab chưa có cấu trúc STUDIO CN. */
function bbntCoBoCucMoi_(sheet) {
  const headers = sheet.getRange('J1:M1').getDisplayValues()[0]
    .map(value => String(value || '').trim().toLowerCase());
  return headers[0] === 'tên dự án' &&
    headers[1] === 'hạng mục công việc' &&
    headers[2] === 'đánh giá' &&
    headers[3] === 'link_word_01';
}

function bbntKiemTraBoCuc_(sheet) {
  if (!bbntCoBoCucMoi_(sheet)) {
    throw new Error('Bố cục tab BBNT STUDIO CN chưa đúng: J=Tên Dự án, K=Hạng mục công việc, L=Đánh giá, M=Link_Word_01.');
  }
}

// ========================================
// 2. XỬ LÝ CHECKBOX CỘT P
// ========================================

function xuLyTaoBienBanNghiemThu(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();

  if (
    sheet.getName() !== BBNT_CONFIG.SHEET_NAME ||
    !bbntCoBoCucMoi_(sheet) ||
    range.getColumn() !== BBNT_CONFIG.ACTION_COLUMN ||
    range.getRow() < 2 ||
    range.getNumRows() !== 1 ||
    range.getNumColumns() !== 1 ||
    e.value !== 'TRUE'
  ) {
    return;
  }

  const row = range.getRow();
  const actionCell = range;
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();

  if (!lock.tryLock(15000)) {
    actionCell.setValue(false);
    actionCell.setNote('Hệ thống đang xử lý yêu cầu khác. Hãy thử lại.');
    return;
  }

  try {
    actionCell.setNote('Đang tạo biên bản nghiệm thu...');
    const result = bbntTaoTaiLieuChoDong_(sheet, row);

    actionCell.setNote(
      'Đã tạo biên bản lần ' + result.lanTao + ': ' + result.url
    );

    bbntThongBao_(
      sheet.getParent(),
      'Đã tạo biên bản lần ' + result.lanTao + ' cho dòng ' + row,
      'Thành công',
      6
    );
  } catch (error) {
    actionCell.setNote('Không thể tạo biên bản: ' + error.message);
    console.error('Dòng ' + row + ': ' + error.stack);

    bbntThongBao_(
      sheet.getParent(),
      error.message,
      'Không thể tạo biên bản - Dòng ' + row,
      10
    );
  } finally {
    actionCell.setValue(false);
    lock.releaseLock();
  }
}

function taoBienBanNghiemThuDongDangChon() {
  const ss = bbntLaySpreadsheet_();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== BBNT_CONFIG.SHEET_NAME) {
    throw new Error('Hãy mở tab ' + BBNT_CONFIG.SHEET_NAME + '.');
  }

  const row = sheet.getActiveRange().getRow();

  if (row < 2) {
    throw new Error('Hãy chọn một ô trong dòng dữ liệu cần tạo biên bản.');
  }

  const lock = LockService.getDocumentLock() || LockService.getScriptLock();

  if (!lock.tryLock(15000)) {
    throw new Error('Hệ thống đang xử lý yêu cầu khác. Hãy thử lại.');
  }

  try {
    const result = bbntTaoTaiLieuChoDong_(sheet, row);

    bbntThongBao_(
      ss,
      'Đã tạo biên bản lần ' + result.lanTao + ' cho dòng ' + row,
      'Thành công',
      6
    );
  } finally {
    lock.releaseLock();
  }
}


// ========================================
// 3. KIỂM TRA DỮ LIỆU VÀ TẠO BIÊN BẢN
// ========================================

function bbntTaoTaiLieuChoDong_(sheet, row) {
  bbntKiemTraBoCuc_(sheet);
  // A:L là dữ liệu; M:O là ba link hiện có.
  const displayValues = sheet.getRange(
    row, 1, 1, BBNT_CONFIG.FIRST_LINK_COLUMN + BBNT_CONFIG.LINK_COUNT - 1
  )
    .getDisplayValues()[0];

  const headers = [
    'Số Biên Bản Nghiệm Thu',
    'Số Hợp đồng khoán',
    'Ký ngày',
    'Họ và tên người đề nghị',
    'Ngày làm đơn',
    'Số CCCD',
    'Mã số thuế',
    'Địa chỉ',
    'Số điện thoại',
    'Tên Dự án',
    'Hạng mục công việc',
    'Đánh giá'
  ];

  const errorValues = [
    '#ERROR!', '#REF!', '#VALUE!', '#N/A',
    '#DIV/0!', '#NUM!', '#NAME?'
  ];
  const missing = [];

  // A do code tự ghi bằng B; người dùng bắt buộc nhập đủ B:L.
  for (let i = 1; i < headers.length; i++) {
    const value = String(displayValues[i] ?? '').trim();

    if (!value || errorValues.some(errorValue => value.startsWith(errorValue))) {
      missing.push(headers[i]);
    }
  }

  if (missing.length > 0) {
    throw new Error(
      'Chưa thể tạo biên bản. Thiếu hoặc lỗi dữ liệu: ' + missing.join(', ')
    );
  }

  const links = displayValues.slice(
    BBNT_CONFIG.FIRST_LINK_COLUMN - 1,
    BBNT_CONFIG.FIRST_LINK_COLUMN - 1 + BBNT_CONFIG.LINK_COUNT
  );
  const emptyLinkIndex = links.findIndex(link => String(link).trim() === '');

  if (emptyLinkIndex === -1) {
    throw new Error('Dòng này đã tạo đủ 3 biên bản nghiệm thu.');
  }

  const [
    ,
    soHopDong,
    kyNgay,
    hoTen,
    ngayLamDon,
    soCccd,
    maSoThue,
    diaChi,
    soDienThoai,
    tenDuAn,
    hangMucCongViec,
    danhGia
  ] = displayValues;

  const ngayKyParts = bbntTachNgay_(kyNgay);
  const ngayLapParts = bbntTachNgay_(ngayLamDon);
  const soBienBan = String(soHopDong).trim();
  const lanTao = String(emptyLinkIndex + 1).padStart(2, '0');
  const targetLinkColumn = BBNT_CONFIG.FIRST_LINK_COLUMN + emptyLinkIndex;

  // Chỉ ghi A sau khi toàn bộ B:L đã vượt qua kiểm tra.
  sheet.getRange(row, 1)
    .setNumberFormat('@')
    .setValue(soBienBan);

  const templateId = bbntLayIdGoogleDoc_(BBNT_CONFIG.TEMPLATE_URL);
  const folderId = bbntLayIdFolder_(BBNT_CONFIG.FOLDER_URL);
  const template = DriveApp.getFileById(templateId);
  const mainFolder = DriveApp.getFolderById(folderId);
  const targetFolder = bbntLayHoacTaoThuMuc_(mainFolder, soHopDong);
  const fileName = bbntLamSachTenFile_(
    soBienBan + ' - ' + hoTen + ' - BBNT STUDIO CN - ' + lanTao
  );
  const copy = template.makeCopy(fileName, targetFolder);

  try {
    const doc = DocumentApp.openById(copy.getId());
    const replacements = {
      '{{SO_BBNT}}': soBienBan,
      '{{SO_BIEN_BAN_NGHIEM_THU}}': soBienBan,
      '{{SO_HOP_DONG}}': soHopDong,
      '{{SO_HD}}': soHopDong,
      '{{KY_NGAY}}': kyNgay,
      '{{NGAY_KY}}': kyNgay,
      '{{NGAY_KY_NGAY}}': ngayKyParts.ngay,
      '{{THANG_KY}}': ngayKyParts.thang,
      '{{NAM_KY}}': ngayKyParts.nam,
      '{{HO_TEN}}': hoTen,
      '{{HO_TEN_IN_HOA}}': hoTen.toLocaleUpperCase('vi-VN'),
      '{{NGAY_LAM_DON}}': ngayLamDon,
      '{{NGAY}}': ngayLapParts.ngay,
      '{{THANG}}': ngayLapParts.thang,
      '{{NAM}}': ngayLapParts.nam,
      '{{SO_CCCD}}': soCccd,
      '{{MA_SO_THUE}}': maSoThue,
      '{{DIA_CHI}}': diaChi,
      '{{SO_DIEN_THOAI}}': soDienThoai,
      '{{TEN_DU_AN}}': tenDuAn,
      '{{HANG_MUC_CONG_VIEC}}': hangMucCongViec,
      '{{DANH_GIA}}': danhGia
    };

    bbntThayBienTrongTaiLieu_(doc, replacements);
    doc.saveAndClose();

    const remaining = bbntTimBienConLai_(
      DocumentApp.openById(copy.getId())
    );

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
// 4. HÀM HỖ TRỢ
// ========================================

function bbntLaySpreadsheet_() {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (activeSpreadsheet) {
    return activeSpreadsheet;
  }

  return SpreadsheetApp.openByUrl(BBNT_CONFIG.SPREADSHEET_URL);
}

/**
 * Dự án Apps Script độc lập có thể không được phép hiện toast.
 * Khi đó chỉ ghi nhật ký để không làm hỏng quy trình chính.
 */
function bbntThongBao_(spreadsheet, message, title, seconds) {
  try {
    spreadsheet.toast(message, title, seconds);
  } catch (error) {
    console.log(title + ': ' + message);
  }
}

function bbntThayBienTrongTaiLieu_(doc, replacements) {
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

function bbntTimBienConLai_(doc) {
  const texts = [doc.getBody().getText()];
  const header = doc.getHeader();
  const footer = doc.getFooter();

  if (header) texts.push(header.getText());
  if (footer) texts.push(footer.getText());

  const matches = texts.join('\n').match(/\{\{[^{}]+\}\}/g) || [];
  return [...new Set(matches)];
}

function bbntTachNgay_(dateText) {
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

function bbntLayIdGoogleDoc_(url) {
  const match = String(url || '').trim().match(
    /docs\.google\.com\/document\/d\/([a-zA-Z0-9_-]+)/
  );

  if (!match) {
    throw new Error('TEMPLATE_URL chưa đúng định dạng Google Docs.');
  }

  return match[1];
}

function bbntLayIdFolder_(urlOrId) {
  const text = String(urlOrId || '').trim();
  const urlMatch = text.match(/drive\.google\.com\/drive\/folders\/([a-zA-Z0-9_-]+)/);

  if (urlMatch) return urlMatch[1];

  if (/^[a-zA-Z0-9_-]+$/.test(text)) return text;

  throw new Error('FOLDER_URL chưa đúng định dạng thư mục Google Drive.');
}

function bbntLayHoacTaoThuMuc_(parentFolder, soHopDong) {
  const folderName = 'STUDIO CN - ' + bbntLamSachTenFile_(soHopDong);
  const folders = parentFolder.getFoldersByName(folderName);

  return folders.hasNext()
    ? folders.next()
    : parentFolder.createFolder(folderName);
}

function bbntLamSachTenFile_(name) {
  return String(name || '')
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, ' ')
    .trim();
}

function bbntKhoaCotLink_(sheet, maxRows) {
  const protectionDescriptions = [
    BBNT_CONFIG.LINK_PROTECTION_DESCRIPTION,
    BBNT_CONFIG.HARD_FILE_PROTECTION_DESCRIPTION
  ];

  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter(protection =>
      protectionDescriptions.includes(protection.getDescription())
    )
    .forEach(protection => protection.remove());

  const linkRange = sheet.getRange(
    2,
    BBNT_CONFIG.FIRST_LINK_COLUMN,
    maxRows - 1,
    BBNT_CONFIG.LINK_COUNT
  );

  linkRange.clearDataValidations();

  const effectiveUser = Session.getEffectiveUser();
  const effectiveEmail = effectiveUser.getEmail();

  [
    { range: linkRange, description: BBNT_CONFIG.LINK_PROTECTION_DESCRIPTION },
    {
      range: sheet.getRange(2, BBNT_CONFIG.HARD_UPLOAD_COLUMN, maxRows - 1, 2),
      description: BBNT_CONFIG.HARD_FILE_PROTECTION_DESCRIPTION
    }
  ].forEach(item => {
    const protection = item.range.protect().setDescription(item.description);
    protection.addEditor(effectiveUser);

    const otherEditors = protection.getEditors().filter(editor =>
      editor.getEmail() !== effectiveEmail
    );

    if (otherEditors.length > 0) protection.removeEditors(otherEditors);
    if (protection.canDomainEdit()) protection.setDomainEdit(false);
  });
}


// ========================================
// 5. BIẾN CẦN DÙNG TRONG GOOGLE DOCS MẪU
// ========================================

const BBNT_TEMPLATE_FIELDS = [
  '{{SO_BBNT}}',
  '{{SO_HOP_DONG}}',
  '{{KY_NGAY}}',
  '{{HO_TEN}}',
  '{{HO_TEN_IN_HOA}}',
  '{{NGAY_LAM_DON}}',
  '{{NGAY}}', '{{THANG}}', '{{NAM}}',
  '{{SO_CCCD}}',
  '{{MA_SO_THUE}}',
  '{{DIA_CHI}}',
  '{{SO_DIEN_THOAI}}',
  '{{TEN_DU_AN}}',
  '{{HANG_MUC_CONG_VIEC}}',
  '{{DANH_GIA}}'
];
