/**
 * TỰ ĐỘNG TẠO HỢP ĐỒNG KHOÁN DỊCH VỤ
 *
 * A:V - Dữ liệu hợp đồng (H là Tên Brand)
 * W:Y - Link Word lần 01, 02, 03
 * Z   - Checkbox tạo Word
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

  FIRST_LINK_COLUMN: 23, // W
  LINK_COUNT: 3,         // W, X, Y
  ACTION_COLUMN: 26,     // Z
  LINK_PROTECTION_DESCRIPTION: 'KHOA_LINK_HOP_DONG_TU_DONG'
};


// ========================================
// 1. MENU VÀ CÀI ĐẶT HỆ THỐNG
// ========================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tạo hợp đồng')
    .addItem('Cài đặt hệ thống', 'caiDatHeThong')
    .addItem('Tạo hợp đồng cho dòng đang chọn', 'taoHopDongDongDangChon')
    .addToUi();
}

/**
 * Chạy hàm này đúng một lần sau khi dán code.
 * Hàm sẽ tạo tiêu đề W:Z, checkbox ở Z và installable onEdit trigger.
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

  xoaCongThucSttCu_(sheet);

  sheet.getRange('W1:Z1').setValues([[
    'Link_Word_01',
    'Link_Word_02',
    'Link_Word_03',
    'TẠO WORD'
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
  }

  sheet.setColumnWidths(CONFIG.FIRST_LINK_COLUMN, CONFIG.LINK_COUNT, 180);
  sheet.setColumnWidth(CONFIG.ACTION_COLUMN, 120);

  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'xuLyTaoHopDong')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('xuLyTaoHopDong')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ss.toast(
    'Đã tạo checkbox cột Z và trigger tự động.',
    'Cài đặt thành công',
    6
  );
}


// ========================================
// 2. XỬ LÝ CHECKBOX Ở CỘT Y
// ========================================

function xuLyTaoHopDong(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();

  if (
    sheet.getName() !== CONFIG.SHEET_NAME ||
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

  // A:V là dữ liệu, W:Y là ba link đã tạo.
  const displayValues = sheet.getRange(row, 1, 1, 25)
    .getDisplayValues()[0];
  const rawValues = sheet.getRange(row, 1, 1, 22).getValues()[0];

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
    'Tổng giá trị thù lao giao khoán (Bằng chữ)'
  ];

  const errorValues = [
    '#ERROR!', '#REF!', '#VALUE!', '#N/A',
    '#DIV/0!', '#NUM!', '#NAME?'
  ];
  const missing = [];

  // Cột A do code tự sinh; bắt buộc người dùng nhập đầy đủ B:V.
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

  const links = displayValues.slice(22, 25);
  const emptyLinkIndex = links.findIndex(link => String(link).trim() === '');
  const currentRowHasLink = links.some(link => String(link).trim() !== '');

  if (emptyLinkIndex === -1) {
    throw new Error('Dòng này đã tạo đủ 3 hợp đồng.');
  }

  const amount = rawValues[20]; // Cột U

  if (typeof amount !== 'number' || !Number.isFinite(amount) || amount <= 0) {
    throw new Error('Tổng giá trị thù lao ở cột U phải là số lớn hơn 0.');
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
    tongThuLaoBangChu
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
    const replacements = {
      '{{SO_HD}}': stt,
      '{{STT}}': stt,
      '{{HO_TEN}}': hoTen,
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
      '{{TONG_THU_LAO_BANG_CHU}}': tongThuLaoBangChu
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
 * Chỉ các dòng có ít nhất một link Word tại W:Y mới được tính.
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
  const existingRows = sheet.getRange(2, 1, lastRow - 1, 25)
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

function khoaCotLink_(sheet, maxRows) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter(protection =>
      protection.getDescription() === CONFIG.LINK_PROTECTION_DESCRIPTION
    )
    .forEach(protection => protection.remove());

  const linkRange = sheet.getRange(
    2,
    CONFIG.FIRST_LINK_COLUMN,
    maxRows - 1,
    CONFIG.LINK_COUNT
  );

  linkRange.clearDataValidations();

  const protection = linkRange
    .protect()
    .setDescription(CONFIG.LINK_PROTECTION_DESCRIPTION);
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
// 6. BIẾN CẦN ĐẶT TRONG GOOGLE DOCS MẪU
// ========================================

const TEMPLATE_FIELDS = [
  '{{SO_HD}}',
  '{{HO_TEN}}',
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
  '{{TONG_THU_LAO_BANG_CHU}}'
];
