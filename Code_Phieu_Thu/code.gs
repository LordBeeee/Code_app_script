/** PHIẾU THU: chạy ptCaiDat một lần từ dự án Apps Script gắn với Sheet.
 * Không khai báo onOpen/onEdit chung để tránh đè các luồng chứng từ khác.
 * Chỉ các tiêu đề trong PT_CONFIG mới thuộc phạm vi xử lý của module này.
 */
const PT_CONFIG = {
  SPREADSHEET_ID: '1ShgTg3vFVaQbIE3R1ziQNQrwMMeiKxG_4_2nCXBmeBQ',
  SHEET_NAME: 'TH_PHIEU_THU',
  TEMPLATE_ID: '1mn7hcT2Z4jxynPZSPvAI2F8WuEF_hHBcxnrNxr79U4I',
  FOLDER_ID: '14tXO4OE6HmjVfshVK0sVYE2MfxhLshxm',
  HEADER_ROW: 1,
  // Sửa danh sách tên nếu tiêu đề thực tế khác. Không phụ thuộc thứ tự cột.
  HEADERS: {
    date: ['Ngày lập phiếu', 'Ngày thu', 'Ngày lập', 'Ngày'],
    payer: ['Người nộp tiền', 'Họ và tên người nộp tiền', 'Họ tên người nộp', 'Người nộp', 'Họ và tên', 'Họ tên'],
    address: ['Địa chỉ người nộp', 'Địa chỉ'],
    amount: ['Số tiền', 'Số tiền thu'],
    words: ['Số tiền bằng chữ', 'Bằng chữ'],
    docs1: ['Link_Word_1'],
    docs2: ['Link_Word_2'],
    docs3: ['Link_Word_3'],
    action: ['Tạo_Word']
  },
  REQUIRED_TOKENS: ['{{NGAY}}', '{{THANG}}', '{{NAM}}', '{{NGUOI_NOP}}',
    '{{DIA_CHI}}', '{{SO_TIEN}}', '{{SO_TIEN_BANG_CHU}}'],
  MAX_AMOUNT: 999999999999999,
  // Giới hạn số dòng tạo tệp trong một lần tích/dán để tránh hết thời gian trigger.
  MAX_CREATE_PER_EDIT: 5
};

function ptMoMenu() {
  let ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (error) {
    console.log('[Phiếu thu] Menu chỉ dùng trong dự án gắn với Sheet. Có thể tạo phiếu bằng checkbox.');
    return;
  }
  ui.createMenu('Phiếu thu')
    .addItem('Cài đặt / cập nhật', 'ptCaiDat')
    .addItem('Tạo phiếu dòng đang chọn', 'ptTaoDongDangChon')
    .addToUi();
}

function ptCaiDat() {
  const ss = SpreadsheetApp.openById(PT_CONFIG.SPREADSHEET_ID);
  const sheet = ptLaySheet_(ss);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // Kiểm tra trước khi thay đổi Sheet. Không sửa tài liệu mẫu.
    const template = DriveApp.getFileById(PT_CONFIG.TEMPLATE_ID);
    if (template.getMimeType() !== 'application/vnd.google-apps.document') {
      throw new Error('Mẫu phải là Google Docs.');
    }
    DriveApp.getFolderById(PT_CONFIG.FOLDER_ID).getName();
    ptKiemTraMau_(DocumentApp.openById(PT_CONFIG.TEMPLATE_ID));
    ptCot_(sheet, ['date', 'payer', 'address', 'amount']);
    const outputKeys = ['words', 'docs1', 'docs2', 'docs3', 'action'];
    outputKeys.forEach(key => {
      const existing = ptCot_(sheet, [], true);
      if (existing[key]) return;
      const column = sheet.getLastColumn() + 1;
      if (column > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), 1);
      sheet.getRange(PT_CONFIG.HEADER_ROW, column).setValue(PT_CONFIG.HEADERS[key][0]);
    });
    // Dọn các cột trống do bản cũ thêm; giữ cột có dữ liệu/link thực tế.
    ptDonCotCu_(sheet);
    const cols = ptCot_(sheet);
    const firstRow = PT_CONFIG.HEADER_ROW + 1;
    const count = sheet.getMaxRows() - PT_CONFIG.HEADER_ROW;
    if (count > 0) {
      // Không reset checkbox hay lịch sử link khi chạy cài đặt lại.
      sheet.getRange(firstRow, cols.action, count, 1).setDataValidation(
        SpreadsheetApp.newDataValidation().requireCheckbox().build());
    }
    if (sheet.getLastRow() >= firstRow) {
      ptCapNhatBangChu_(sheet, cols, firstRow, sheet.getLastRow());
    }
    outputKeys.forEach(key => sheet.setColumnWidth(cols[key], key === 'action' ? 145 : 260));
    // Chỉ quản lý trigger của module phiếu thu.
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (['ptXuLySua', 'ptMoMenu'].includes(trigger.getHandlerFunction())) ScriptApp.deleteTrigger(trigger);
    });
    ScriptApp.newTrigger('ptXuLySua').forSpreadsheet(ss).onEdit().create();
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active && active.getId() === ss.getId()) {
      ScriptApp.newTrigger('ptMoMenu').forSpreadsheet(ss).onOpen().create();
    }
    ptThongBao_(ss, 'Đã cài đặt. Nhập số tiền và tích Tạo_Word; link Docs vào Link_Word_1, 2, 3.', 8);
  } finally {
    lock.releaseLock();
  }
}

function ptXuLySua(e) {
  if (!e || !e.range || !e.source || e.source.getId() !== PT_CONFIG.SPREADSHEET_ID) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== PT_CONFIG.SHEET_NAME) return;
  const firstRow = Math.max(PT_CONFIG.HEADER_ROW + 1, e.range.getRow());
  const lastRow = e.range.getLastRow();
  if (firstRow > lastRow) return;
  const cols = ptCot_(sheet);
  const touches = column => column >= e.range.getColumn() && column <= e.range.getLastColumn();
  const amountEdited = touches(cols.amount);
  const actionEdited = touches(cols.action);
  if (!amountEdited && !actionEdited) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) {
    ptThongBao_(e.source, 'Đang xử lý phiếu khác. Hãy sửa lại số tiền hoặc bỏ tích rồi tích lại.', 10);
    return;
  }
  try {
    if (amountEdited) ptCapNhatBangChu_(sheet, cols, firstRow, lastRow);
    if (actionEdited) {
      let processed = 0;
      for (let row = firstRow; row <= lastRow; row++) {
        if (sheet.getRange(row, cols.action).getValue() !== true) continue;
        if (processed >= PT_CONFIG.MAX_CREATE_PER_EDIT) {
          sheet.getRange(row, cols.action).setValue(false);
          sheet.getRange(row, cols.action).setNote('Chưa tạo: tối đa ' + PT_CONFIG.MAX_CREATE_PER_EDIT + ' dòng/lần. Hãy tích lại dòng này.');
          continue;
        }
        processed++;
        ptTaoAnToan_(sheet, cols, row, e.source);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function ptTaoDongDangChon() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss || ss.getId() !== PT_CONFIG.SPREADSHEET_ID) throw new Error('Hãy mở đúng bảng tính phiếu thu.');
  const sheet = ss.getActiveSheet();
  if (sheet.getName() !== PT_CONFIG.SHEET_NAME) throw new Error('Hãy chọn tab ' + PT_CONFIG.SHEET_NAME);
  const row = sheet.getActiveRange().getRow();
  if (row <= PT_CONFIG.HEADER_ROW) throw new Error('Hãy chọn một dòng dữ liệu.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    ptTaoAnToan_(sheet, ptCot_(sheet), row, ss);
  } finally {
    lock.releaseLock();
  }
}

function ptTaoAnToan_(sheet, cols, row, ss) {
  const action = sheet.getRange(row, cols.action);
  action.setNote('Đang tạo phiếu thu…');
  try {
    const name = ptTaoPhieu_(sheet, cols, row, ss.getSpreadsheetTimeZone());
    action.setNote('Đã tạo: ' + name);
    ptThongBao_(ss, 'Đã tạo phiếu thu dòng ' + row, 5);
  } catch (error) {
    action.setNote('Lỗi: ' + String(error.message || error));
    ptThongBao_(ss, 'Dòng ' + row + ': ' + String(error.message || error), 10);
  } finally {
    action.setValue(false);
  }
}

// Thông báo UI có thể không được hỗ trợ khi chạy từ editor/dự án độc lập.
// Không để lỗi thông báo làm hỏng cài đặt hoặc đổi trạng thái phiếu đã tạo thành lỗi.
function ptThongBao_(ss, message, seconds) {
  try {
    ss.toast(message, 'Phiếu thu', seconds || 5);
  } catch (error) {
    console.log('[Phiếu thu] ' + message + ' (không hiện được toast: ' + String(error.message || error) + ')');
  }
}

function ptTaoPhieu_(sheet, cols, row, timezone) {
  const slots = [cols.docs1, cols.docs2, cols.docs3];
  const index = slots.findIndex(column => {
    const cell = sheet.getRange(row, column);
    return cell.getValue() === '' && !cell.getFormula();
  });
  if (index < 0) throw new Error('Đã đủ 3 link Word. Không tạo thêm hoặc ghi đè link cũ.');
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const payer = String(values[cols.payer - 1] || '').trim();
  const address = String(values[cols.address - 1] || '').trim();
  const amount = ptDocSoTien_(values[cols.amount - 1]);
  if (!payer || !address) throw new Error('Thiếu họ tên hoặc địa chỉ người nộp tiền.');
  if (amount === null) throw new Error('Số tiền phải là số nguyên dương, tối đa 999.999.999.999.999 đồng.');
  const date = ptTachNgay_(values[cols.date - 1], timezone);
  const words = ptBangChu_(amount);
  sheet.getRange(row, cols.words).setValue(words).clearNote();
  const replacements = {
    '{{NGAY}}': date.day, '{{THANG}}': date.month, '{{NAM}}': date.year,
    '{{NGUOI_NOP}}': payer, '{{DIA_CHI}}': address,
    '{{SO_TIEN}}': String(amount).replace(/\B(?=(\d{3})+(?!\d))/g, '.'),
    '{{SO_TIEN_BANG_CHU}}': words
  };
  // Kiểm tra mẫu mỗi lần tạo vì mẫu có thể đã được người dùng sửa.
  ptKiemTraMau_(DocumentApp.openById(PT_CONFIG.TEMPLATE_ID));
  const folder = DriveApp.getFolderById(PT_CONFIG.FOLDER_ID);
  const docsCell = sheet.getRange(row, slots[index]);
  const version = index + 1;
  const stamp = Utilities.formatDate(new Date(), timezone, 'yyyyMMdd_HHmmss');
  const name = ptTenFile_('PhieuThu_' + date.year + date.month + date.day + '_' + payer +
    '_Dong' + row + '_Lan' + String(version).padStart(2, '0') + '_' + stamp);
  const copy = DriveApp.getFileById(PT_CONFIG.TEMPLATE_ID).makeCopy(name, folder);
  // Lưu link ngay để giữ dấu bản sao nếu bước điền mẫu gặp lỗi.
  docsCell.setValue(copy.getUrl()).setNote('Đang điền dữ liệu. Nếu tạo lỗi, kiểm tra bản sao này.');
  const doc = DocumentApp.openById(copy.getId());
  try {
    const containers = ptVungTaiLieu_(doc);
    Object.keys(replacements).forEach(key => {
      const escaped = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      containers.forEach(container => container.replaceText(escaped, replacements[key]));
    });
    const remaining = ptBien_(doc);
    if (remaining.length) throw new Error('Còn biến chưa điền: ' + remaining.join(', '));
  } finally {
    doc.saveAndClose();
  }
  docsCell.clearNote();
  return name;
}

function ptDonCotCu_(sheet) {
  const legacy = ['TẠO PHIẾU THU', 'Link_Docs_Phieu_Thu', 'Link_PDF_Phieu_Thu', 'Trạng thái tạo phiếu thu'].map(ptChuanHoa_);
  const headers = sheet.getRange(PT_CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  // Từ phải qua trái để không lệch chỉ số khi xóa cột trống.
  for (let i = headers.length - 1; i >= 0; i--) {
    if (!legacy.includes(ptChuanHoa_(headers[i]))) continue;
    const oldAction = ptChuanHoa_(headers[i]) === ptChuanHoa_('TẠO PHIẾU THU');
    const count = sheet.getMaxRows() - PT_CONFIG.HEADER_ROW;
    let empty = true;
    if (count > 0) {
      const range = sheet.getRange(PT_CONFIG.HEADER_ROW + 1, i + 1, count, 1);
      empty = range.getValues().every(([value]) => value === '' || (oldAction && value === false)) &&
        range.getFormulas().every(([formula]) => !formula);
      if (oldAction) range.clearDataValidations();
    }
    if (empty) sheet.deleteColumns(i + 1, 1);
  }
}

function ptLaySheet_(ss) {
  const sheet = ss.getSheetByName(PT_CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Không tìm thấy tab ' + PT_CONFIG.SHEET_NAME);
  return sheet;
}

function ptChuanHoa_(value) {
  return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd').replace(/Đ/g, 'D').replace(/[_\s]+/g, ' ').trim().toLowerCase();
}

function ptCot_(sheet, required, allowMissing) {
  const headers = sheet.getRange(PT_CONFIG.HEADER_ROW, 1, 1, Math.max(1, sheet.getLastColumn()))
    .getDisplayValues()[0].map(ptChuanHoa_);
  const cols = {};
  Object.keys(PT_CONFIG.HEADERS).forEach(key => {
    const aliases = PT_CONFIG.HEADERS[key].map(ptChuanHoa_);
    const matches = headers.map((header, index) => aliases.includes(header) ? index + 1 : 0).filter(Boolean);
    if (matches.length > 1) throw new Error('Có nhiều cột cùng nghĩa: ' + PT_CONFIG.HEADERS[key][0]);
    if (matches.length) cols[key] = matches[0];
  });
  const keys = required || (allowMissing ? [] : Object.keys(PT_CONFIG.HEADERS));
  const missing = keys.filter(key => !cols[key]);
  if (missing.length) throw new Error('Thiếu tiêu đề hàng ' + PT_CONFIG.HEADER_ROW + ': ' +
    missing.map(key => PT_CONFIG.HEADERS[key][0]).join(', ') + '. Xem PT_CONFIG.HEADERS.');
  return cols;
}

function ptCapNhatBangChu_(sheet, cols, firstRow, lastRow) {
  const count = lastRow - firstRow + 1;
  const values = sheet.getRange(firstRow, cols.amount, count, 1).getValues();
  const notes = [];
  const words = values.map(([value]) => {
    const amount = ptDocSoTien_(value);
    const empty = value === '' || value === null || String(value).trim() === '';
    notes.push([amount === null && !empty ? 'Số tiền không hợp lệ. Nhập số nguyên dương; ví dụ 1250000 hoặc 1.250.000.' : '']);
    return [amount === null ? '' : ptBangChu_(amount)];
  });
  sheet.getRange(firstRow, cols.words, count, 1).setValues(words).setNotes(notes);
}

function ptDocSoTien_(raw) {
  if (typeof raw === 'number') {
    return Number.isSafeInteger(raw) && raw > 0 && raw <= PT_CONFIG.MAX_AMOUNT ? raw : null;
  }
  const input = String(raw == null ? '' : raw).trim().replace(/\u00a0/g, ' ');
  if (!(/^\d+$/.test(input) || /^\d{1,3}(?:\.\d{3})+$/.test(input) ||
    /^\d{1,3}(?:,\d{3})+$/.test(input) || /^\d{1,3}(?: \d{3})+$/.test(input))) return null;
  const amount = Number(input.replace(/[., ]/g, ''));
  return Number.isSafeInteger(amount) && amount > 0 && amount <= PT_CONFIG.MAX_AMOUNT ? amount : null;
}

function ptBaChuSo_(number, full) {
  const digits = ['không', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín'];
  const hundred = Math.floor(number / 100);
  const ten = Math.floor(number / 10) % 10;
  const one = number % 10;
  const words = [];
  if (hundred || full) words.push(digits[hundred], 'trăm');
  if (ten > 1) {
    words.push(digits[ten], 'mươi');
    if (one) words.push(one === 1 ? 'mốt' : one === 5 ? 'lăm' : digits[one]);
  } else if (ten === 1) {
    words.push('mười');
    if (one) words.push(one === 5 ? 'lăm' : digits[one]);
  } else if (one) {
    if (hundred || full) words.push('lẻ');
    words.push(digits[one]);
  }
  return words.join(' ');
}

function ptBangChu_(amount) {
  if (ptDocSoTien_(amount) === null) throw new Error('Số tiền không hợp lệ.');
  const units = ['', 'nghìn', 'triệu', 'tỷ', 'nghìn tỷ'];
  const groups = [];
  for (let rest = amount; rest > 0; rest = Math.floor(rest / 1000)) groups.push(rest % 1000);
  const words = [];
  for (let i = groups.length - 1; i >= 0; i--) {
    if (groups[i]) words.push(ptBaChuSo_(groups[i], words.length > 0 && groups[i] < 100) + (units[i] ? ' ' + units[i] : ''));
  }
  const result = words.join(' ') + ' đồng';
  return result.charAt(0).toUpperCase() + result.slice(1);
}

function ptTachNgay_(raw, timezone) {
  let text;
  if (raw instanceof Date && !isNaN(raw.getTime())) text = Utilities.formatDate(raw, timezone, 'dd/MM/yyyy');
  else text = String(raw || '').trim();
  const match = /^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/.exec(text);
  if (!match) throw new Error('Ngày lập phiếu phải là ô ngày hoặc dd/MM/yyyy.');
  const day = Number(match[1]), month = Number(match[2]), year = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (year < 1900 || date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
    throw new Error('Ngày lập phiếu không hợp lệ.');
  }
  return { day: String(day).padStart(2, '0'), month: String(month).padStart(2, '0'), year: String(year) };
}

function ptVungTaiLieu_(doc) {
  const containers = [];
  const add = documentTab => {
    containers.push(documentTab.getBody());
    const header = documentTab.getHeader(), footer = documentTab.getFooter();
    if (header) containers.push(header);
    if (footer) containers.push(footer);
  };
  if (typeof doc.getTabs === 'function') {
    const visit = tabs => tabs.forEach(tab => {
      if (tab.getType() === DocumentApp.TabType.DOCUMENT_TAB) add(tab.asDocumentTab());
      visit(tab.getChildTabs());
    });
    visit(doc.getTabs());
  } else add(doc);
  return containers;
}

function ptBien_(doc) {
  return [...new Set(ptVungTaiLieu_(doc).map(container => container.getText()).join('\n')
    .match(/\{\{[^{}]+\}\}/g) || [])];
}

function ptKiemTraMau_(doc) {
  const tokens = ptBien_(doc);
  const missing = PT_CONFIG.REQUIRED_TOKENS.filter(token => !tokens.includes(token));
  const unknown = tokens.filter(token => !PT_CONFIG.REQUIRED_TOKENS.includes(token));
  if (missing.length || unknown.length) throw new Error('Cần chỉnh mẫu Docs. ' +
    (missing.length ? 'Thiếu: ' + missing.join(', ') + '. ' : '') +
    (unknown.length ? 'Biến không hỗ trợ (hãy bỏ): ' + unknown.join(', ') : ''));
}

function ptTenFile_(name) {
  return name.replace(/[\\/:*?"<>|\r\n]/g, '_').slice(0, 220);
}
