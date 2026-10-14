/**
 * TỰ ĐỘNG TẠO BIÊN BẢN NGHIỆM THU
 *
 * A   - Số Biên Bản Nghiệm Thu (tự lấy bằng cột B)
 * B:R - Dữ liệu bắt buộc
 * S:U - Link Word lần 01, 02, 03
 * V   - Checkbox TẠO WORD
 *
 * Tên hàm và cấu hình dùng tiền tố BBNT để không xung đột với các
 * script Hợp đồng khoán và Giấy đề nghị thanh toán trong cùng dự án.
 */

const BBNT_CONFIG = {
  SHEET_NAME: 'TH_BIÊN BẢN NGHIỆM THU',

  SPREADSHEET_URL:
    'https://docs.google.com/spreadsheets/d/1ShgTg3vFVaQbIE3R1ziQNQrwMMeiKxG_4_2nCXBmeBQ/edit?gid=422970170#gid=422970170',

  TEMPLATE_URL:
    'https://docs.google.com/document/d/1WO2V0MroDxffnLnTOaMBRn0W7baZgflzhqMuSqkT8n4/edit?tab=t.0',

  FOLDER_URL:
    'https://drive.google.com/drive/folders/1sP4YJ8EVSNlq2D22M_J5ST2LXDjZ9ZKv?hl=vi',

  FIRST_LINK_COLUMN: 19, // S
  LINK_COUNT: 3,         // S, T, U
  ACTION_COLUMN: 22,     // V
  LINK_PROTECTION_DESCRIPTION: 'KHOA_LINK_BBNT_TU_DONG'
};


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

  const templateId = bbntLayIdGoogleDoc_(BBNT_CONFIG.TEMPLATE_URL);
  const folderId = bbntLayIdFolder_(BBNT_CONFIG.FOLDER_URL);
  const template = DriveApp.getFileById(templateId);

  if (template.getMimeType() !== MimeType.GOOGLE_DOCS) {
    throw new Error('File mẫu phải là Google Docs.');
  }

  // Kiểm tra quyền truy cập trước khi tạo trigger.
  DocumentApp.openById(templateId).getBody();
  DriveApp.getFolderById(folderId).getName();

  bbntXoaHaiCotTrungLap_(sheet);
  bbntChuanHoaCotLink_(sheet);
  sheet.getRange('A1').setValue('Số Biên Bản Nghiệm Thu');
  sheet.getRange('S1:V1').setValues([[
    'Link_Word_01',
    'Link_Word_02',
    'Link_Word_03',
    'TẠO WORD'
  ]]);

  const maxRows = sheet.getMaxRows();

  if (maxRows >= 2) {
    sheet.getRange(2, BBNT_CONFIG.ACTION_COLUMN, maxRows - 1, 1)
      .insertCheckboxes()
      .setHorizontalAlignment('center')
      .setBackground('#E8F5E9');

    bbntKhoaCotLink_(sheet, maxRows);
  }

  sheet.setColumnWidths(
    BBNT_CONFIG.FIRST_LINK_COLUMN,
    BBNT_CONFIG.LINK_COUNT,
    180
  );
  sheet.setColumnWidth(BBNT_CONFIG.ACTION_COLUMN, 120);

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
    'Đã tạo checkbox cột V và trigger tự động.',
    'Cài đặt biên bản nghiệm thu thành công',
    7
  );
}

/**
 * Xóa hai cột cũ J:K theo tên tiêu đề. Hàm an toàn khi chạy lại:
 * nếu tiêu đề không còn tồn tại thì không xóa thêm cột nào.
 */
function bbntXoaHaiCotTrungLap_(sheet) {
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(value => String(value || '').trim().toLowerCase());
  const targets = [
    'tên dự án',
    'số lượng video(công việc giao khoán)'
  ];
  const columns = targets
    .map(target => headers.indexOf(target) + 1)
    .filter(column => column > 0)
    .sort((a, b) => b - a);

  columns.forEach(column => sheet.deleteColumn(column));
}

/**
 * Dọn cấu trúc link bị trùng do code cũ từng tạo U:X sau khi J:K đã bị xóa.
 * Các URL tài liệu hiện có trong S:X được gom lại và giữ tối đa ba link ở S:U.
 */
function bbntChuanHoaCotLink_(sheet) {
  const maxRows = sheet.getMaxRows();
  const firstColumn = 19; // S
  const width = 6;        // S:X
  const oldRange = sheet.getRange(1, firstColumn, maxRows, width);
  const oldValues = oldRange.getDisplayValues();
  const preservedLinks = [];

  for (let rowIndex = 1; rowIndex < oldValues.length; rowIndex++) {
    const links = [];

    oldValues[rowIndex].forEach(value => {
      const text = String(value || '').trim();

      if (
        /https:\/\/docs\.google\.com\/document\//i.test(text) &&
        !links.includes(text)
      ) {
        links.push(text);
      }
    });

    preservedLinks.push([
      links[0] || '',
      links[1] || '',
      links[2] || ''
    ]);
  }

  oldRange.clearContent().clearDataValidations().clearNote();
  sheet.getRange(1, firstColumn, maxRows, width).setBackground(null);

  if (preservedLinks.length > 0) {
    sheet.getRange(2, firstColumn, preservedLinks.length, 3)
      .setValues(preservedLinks);
  }
}


// ========================================
// 2. XỬ LÝ CHECKBOX CỘT V
// ========================================

function xuLyTaoBienBanNghiemThu(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();

  if (
    sheet.getName() !== BBNT_CONFIG.SHEET_NAME ||
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
  // A:R là dữ liệu; S:U là ba link hiện có.
  const displayValues = sheet.getRange(row, 1, 1, 21)
    .getDisplayValues()[0];
  const rawValues = sheet.getRange(row, 1, 1, 18)
    .getValues()[0];

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
    'Nội dung Kịch bản',
    'Số lượng Kịch bản',
    'Số lượng Video',
    'Khung Video',
    'Chất lượng Video',
    'Thời lượng Video',
    'Hiệu ứng Video',
    'Nội dung Diễn Viên',
    'Số lượng Diễn viên'
  ];

  const errorValues = [
    '#ERROR!', '#REF!', '#VALUE!', '#N/A',
    '#DIV/0!', '#NUM!', '#NAME?'
  ];
  const missing = [];

  // A do code tự ghi bằng B; người dùng bắt buộc nhập đủ B:R.
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

  const links = displayValues.slice(18, 21);
  const emptyLinkIndex = links.findIndex(link => String(link).trim() === '');

  if (emptyLinkIndex === -1) {
    throw new Error('Dòng này đã tạo đủ 3 biên bản nghiệm thu.');
  }

  const positiveNumberColumns = [
    { index: 10, name: headers[10] }, // K
    { index: 11, name: headers[11] }, // L
    { index: 17, name: headers[17] }  // R
  ];
  const invalidNumbers = positiveNumberColumns
    .filter(item => {
      return !bbntLaSoDuong_(
        rawValues[item.index],
        displayValues[item.index]
      );
    })
    .map(item => item.name);

  if (invalidNumbers.length > 0) {
    throw new Error(
      'Các trường phải là số lớn hơn 0: ' + invalidNumbers.join(', ')
    );
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
    noiDungKichBan,
    soLuongKichBan,
    soLuongVideo,
    khungVideo,
    chatLuongVideo,
    thoiLuongVideo,
    hieuUngVideo,
    noiDungDienVien,
    soLuongDienVien
  ] = displayValues;

  const ngayKyParts = bbntTachNgay_(kyNgay);
  const ngayLapParts = bbntTachNgay_(ngayLamDon);
  const soBienBan = String(soHopDong).trim();
  const lanTao = String(emptyLinkIndex + 1).padStart(2, '0');
  const targetLinkColumn = BBNT_CONFIG.FIRST_LINK_COLUMN + emptyLinkIndex;

  // Chỉ ghi A sau khi toàn bộ B:R đã vượt qua kiểm tra.
  sheet.getRange(row, 1)
    .setNumberFormat('@')
    .setValue(soBienBan);

  const templateId = bbntLayIdGoogleDoc_(BBNT_CONFIG.TEMPLATE_URL);
  const folderId = bbntLayIdFolder_(BBNT_CONFIG.FOLDER_URL);
  const template = DriveApp.getFileById(templateId);
  const mainFolder = DriveApp.getFolderById(folderId);
  const targetFolder = bbntLayHoacTaoThuMuc_(mainFolder, soHopDong);
  const fileName = bbntLamSachTenFile_(
    soBienBan + ' - ' + hoTen + ' - BBNT HĐ khoán việc - ' + lanTao
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
      '{{NGAY_LAM_DON}}': ngayLamDon,
      '{{NGAY}}': ngayLapParts.ngay,
      '{{THANG}}': ngayLapParts.thang,
      '{{NAM}}': ngayLapParts.nam,
      '{{SO_CCCD}}': soCccd,
      '{{MA_SO_THUE}}': maSoThue,
      '{{DIA_CHI}}': diaChi,
      '{{SO_DIEN_THOAI}}': soDienThoai,
      // Tương thích mẫu cũ sau khi bỏ hai cột J:K.
      '{{TEN_DU_AN}}': '',
      '{{SO_LUONG_CONG_VIEC}}': soLuongVideo,
      '{{NOI_DUNG_KICH_BAN}}': noiDungKichBan,
      '{{SO_LUONG_KICH_BAN}}': soLuongKichBan,
      '{{SO_LUONG_VIDEO}}': soLuongVideo,
      '{{KHUNG_VIDEO}}': khungVideo,
      '{{CHAT_LUONG_VIDEO}}': chatLuongVideo,
      '{{THOI_LUONG_VIDEO}}': thoiLuongVideo,
      '{{HIEU_UNG_VIDEO}}': hieuUngVideo,
      '{{NOI_DUNG_DIEN_VIEN}}': noiDungDienVien,
      '{{SO_LUONG_DIEN_VIEN}}': soLuongDienVien
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

/** Chấp nhận số thật hoặc chuỗi số do công thức/Sheet trả về. */
function bbntLaSoDuong_(rawValue, displayValue) {
  if (typeof rawValue === 'number') {
    return Number.isFinite(rawValue) && rawValue > 0;
  }

  const text = String(
    rawValue !== '' && rawValue !== null && rawValue !== undefined
      ? rawValue
      : displayValue
  )
    .trim()
    .replace(/\s+/g, '')
    .replace(',', '.');

  if (!/^\d+(?:\.\d+)?$/.test(text)) {
    return false;
  }

  const number = Number(text);
  return Number.isFinite(number) && number > 0;
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
  const folderName = bbntLamSachTenFile_(soHopDong);
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
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter(protection =>
      protection.getDescription() ===
      BBNT_CONFIG.LINK_PROTECTION_DESCRIPTION
    )
    .forEach(protection => protection.remove());

  const linkRange = sheet.getRange(
    2,
    BBNT_CONFIG.FIRST_LINK_COLUMN,
    maxRows - 1,
    BBNT_CONFIG.LINK_COUNT
  );

  linkRange.clearDataValidations();

  const protection = linkRange
    .protect()
    .setDescription(BBNT_CONFIG.LINK_PROTECTION_DESCRIPTION);
  const effectiveUser = Session.getEffectiveUser();
  const effectiveEmail = effectiveUser.getEmail();

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
}


// ========================================
// 5. BIẾN CẦN DÙNG TRONG GOOGLE DOCS MẪU
// ========================================

const BBNT_TEMPLATE_FIELDS = [
  '{{SO_BBNT}}',
  '{{SO_HOP_DONG}}',
  '{{KY_NGAY}}',
  '{{HO_TEN}}',
  '{{NGAY_LAM_DON}}',
  '{{NGAY}}', '{{THANG}}', '{{NAM}}',
  '{{SO_CCCD}}',
  '{{MA_SO_THUE}}',
  '{{DIA_CHI}}',
  '{{SO_DIEN_THOAI}}',
  '{{NOI_DUNG_KICH_BAN}}',
  '{{SO_LUONG_KICH_BAN}}',
  '{{SO_LUONG_VIDEO}}',
  '{{KHUNG_VIDEO}}',
  '{{CHAT_LUONG_VIDEO}}',
  '{{THOI_LUONG_VIDEO}}',
  '{{HIEU_UNG_VIDEO}}',
  '{{NOI_DUNG_DIEN_VIEN}}',
  '{{SO_LUONG_DIEN_VIEN}}'
];
