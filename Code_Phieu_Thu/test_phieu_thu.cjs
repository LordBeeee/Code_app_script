// Kiểm tra cục bộ; không tải tệp này vào Google Apps Script.
const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const source = fs.readFileSync(require('node:path').join(__dirname, 'code.gs'), 'utf8');
let count = 0;
function test(name, run) { run(); count++; console.log('OK ' + name); }
const ctx = vm.createContext({ console });
vm.runInContext(source, ctx, { filename: 'code.gs' });

test('Số tiền: phân cách hàng nghìn, giới hạn và dữ liệu lỗi', () => {
  for (const raw of [1250000, '1250000', '1.250.000', '1,250,000', '1 250 000']) {
    assert.equal(ctx.ptDocSoTien_(raw), 1250000);
  }
  for (const raw of ['', null, 0, -1, 1.5, '1.25', '1,000.000', '10 đồng', 'abc', 1e15, true]) {
    assert.equal(ctx.ptDocSoTien_(raw), null, String(raw));
  }
  assert.equal(ctx.ptDocSoTien_('999.999.999.999.999'), 999999999999999);
});

test('Đọc tiền bằng chữ: lẻ, mốt, lăm và các nhóm trống', () => {
  const cases = [[1, 'Một đồng'], [15, 'Mười lăm đồng'], [21, 'Hai mươi mốt đồng'],
    [105, 'Một trăm lẻ năm đồng'], [1250000, 'Một triệu hai trăm năm mươi nghìn đồng'],
    [1000001, 'Một triệu không trăm lẻ một đồng'],
    [1001000, 'Một triệu không trăm lẻ một nghìn đồng'],
    [1000000000000, 'Một nghìn tỷ đồng']];
  cases.forEach(([raw, expected]) => assert.equal(ctx.ptBangChu_(raw), expected));
});

test('Ngày tháng: năm nhuận và ngày không tồn tại', () => {
  assert.equal(ctx.ptTachNgay_('29/02/2024').day, '29');
  assert.equal(ctx.ptTachNgay_('3/11/2026').month, '11');
  for (const raw of ['29/02/2025', '31/04/2026', '', '2026-11-03']) {
    assert.throws(() => ctx.ptTachNgay_(raw));
  }
});

class Range {
  constructor(sheet, row, col, rows = 1, cols = 1) { Object.assign(this, { sheet, row, col, rows, cols }); }
  getValues() { return Array.from({ length: this.rows }, (_, i) => Array.from({ length: this.cols }, (_, j) => this.sheet.data[this.row + i - 1]?.[this.col + j - 1] ?? '')); }
  getDisplayValues() { return this.getValues().map(row => row.map(String)); }
  getValue() { return this.getValues()[0][0]; }
  getFormula() { return this.sheet.formulas?.[this.row + ':' + this.col] || ''; }
  getFormulas() { return Array.from({ length: this.rows }, (_, i) => [this.sheet.formulas?.[(this.row + i) + ':' + this.col] || '']); }
  setValue(value) { return this.setValues([[value]]); }
  setValues(values) { values.forEach((row, i) => row.forEach((value, j) => { this.sheet.data[this.row + i - 1] ||= []; this.sheet.data[this.row + i - 1][this.col + j - 1] = value; })); return this; }
  clearNote() { return this.setNotes([['']]); }
  setNote(note) { return this.setNotes([[note]]); }
  clearDataValidations() { this.sheet.clearedValidations ||= []; this.sheet.clearedValidations.push(this.col); return this; }
  setNotes(notes) { this.sheet.notes[this.row + ':' + this.col] = notes; return this; }
  setWrap() { return this; }
  getSheet() { return this.sheet; }
  getRow() { return this.row; }
  getColumn() { return this.col; }
  getLastRow() { return this.row + this.rows - 1; }
  getLastColumn() { return this.col + this.cols - 1; }
}
function sheet() {
  return { data: [ ['Ngày lập phiếu', 'Họ và Tên', 'Địa chỉ', 'Số tiền', 'Số tiền bằng chữ', 'Link_Word_1', 'Link_Word_2', 'Link_Word_3', 'Tạo_Word'],
    ['03/11/2026', 'Nguyễn A', 'TP HCM', 1250000, '', '', '', '', false] ], notes: {},
    getRange(...args) { return new Range(this, ...args); }, getLastColumn() { return this.data[0].length; },
    getMaxRows() { return this.data.length; }, deleteColumns(col, count) { this.data.forEach(row => row.splice(col - 1, count)); },
    getName() { return 'TH_PHIEU_THU'; } };
}
test('Nhận diện theo tiêu đề, bỏ dấu và phát hiện trùng', () => {
  const s = sheet();
  s.data[0][3] = 'SO_TIEN';
  assert.equal(ctx.ptCot_(s).amount, 4);
  s.data[0].push('Số tiền thu');
  assert.throws(() => ctx.ptCot_(s), /nhiều cột/);
  s.data[0] = ['Người nộp tiền'];
  assert.throws(() => ctx.ptCot_(s), /Thiếu tiêu đề/);
});
test('Dán nhiều số tiền; xóa tiền và báo dữ liệu sai', () => {
  const s = sheet();
  s.data.push(['', '', '', ''], ['', '', '', '1.25']);
  ctx.ptCapNhatBangChu_(s, ctx.ptCot_(s), 2, 4);
  assert.equal(s.data[1][4], 'Một triệu hai trăm năm mươi nghìn đồng');
  assert.equal(s.data[2][4], '');
  assert.equal(s.data[3][4], '');
  assert.match(s.notes['2:5'][2][0], /không hợp lệ/);
});

const tokens = ['{{NGAY}}', '{{THANG}}', '{{NAM}}', '{{NGUOI_NOP}}', '{{DIA_CHI}}', '{{SO_TIEN}}', '{{SO_TIEN_BANG_CHU}}'];
let templateText = tokens.join(' ') + ' {{SO_TIEN_BANG_CHU}}';
let files = [], failFill = false;
function document(text) {
  const body = { text, getText() { return this.text; }, replaceText(pattern, replacement) { if (failFill) throw Error('Fill failed'); this.text = this.text.replace(new RegExp(pattern, 'g'), () => replacement); } };
  return { getBody: () => body, getHeader: () => null, getFooter: () => null, saveAndClose() {} };
}
ctx.DocumentApp = { openById(id) { return id === '1mn7hcT2Z4jxynPZSPvAI2F8WuEF_hHBcxnrNxr79U4I' ? document(templateText) : files.find(f => f.id === id).doc; } };
const folder = { createFile() { throw Error('Không được tạo PDF'); } };
ctx.DriveApp = { getFolderById: () => folder, getFileById() { return { makeCopy(name) {
  const file = { id: String(files.length + 1), name, doc: document(templateText),
    getId() { return this.id; }, getUrl() { return 'https://docs/' + this.id; }, getAs() { return { setName() { return this; } }; } };
  files.push(file); return file;
} }; } };
ctx.MimeType = { PDF: 'application/pdf' };
ctx.Utilities = { formatDate: () => '20261103_110000' };

test('Tạo bản sao, điền hai vị trí bằng chữ và lưu nhiều phiên bản', () => {
  const s = sheet(), cols = ctx.ptCot_(s);
  ctx.ptTaoPhieu_(s, cols, 2, 'Asia/Ho_Chi_Minh');
  assert.equal(files.length, 1);
  assert.equal((files[0].doc.getBody().text.match(/Một triệu/g) || []).length, 2);
  assert.equal(files[0].doc.getBody().text.includes('{{'), false);
  ctx.ptTaoPhieu_(s, cols, 2, 'Asia/Ho_Chi_Minh');
  assert.match(files[1].name, /Lan02/);
  assert.equal(s.data[1][5], 'https://docs/1');
  assert.equal(s.data[1][6], 'https://docs/2');
  ctx.ptTaoPhieu_(s, cols, 2, 'Asia/Ho_Chi_Minh');
  assert.equal(s.data[1][7], 'https://docs/3');
  const before = files.length;
  assert.throws(() => ctx.ptTaoPhieu_(s, cols, 2, 'Asia/Ho_Chi_Minh'), /Đã đủ 3/);
  assert.equal(files.length, before);
});
test('Mẫu sai hoặc dữ liệu thiếu không tạo tệp', () => {
  const s = sheet(), cols = ctx.ptCot_(s), before = files.length;
  templateText += ' {{SO_PHIEU}}';
  assert.throws(() => ctx.ptTaoPhieu_(s, cols, 2, 'Asia/Ho_Chi_Minh'), /không hỗ trợ/);
  templateText = tokens.join(' ');
  s.data[1][1] = '';
  assert.throws(() => ctx.ptTaoPhieu_(s, cols, 2, 'Asia/Ho_Chi_Minh'), /Thiếu/);
  assert.equal(files.length, before);
});
test('Điền mẫu lỗi: giữ link bản sao, báo lỗi và bỏ checkbox', () => {
  const s = sheet(); s.data[1][8] = true; failFill = true;
  ctx.ptTaoAnToan_(s, ctx.ptCot_(s), 2, { getSpreadsheetTimeZone: () => 'Asia/Ho_Chi_Minh', toast() {} });
  assert.match(s.data[1][5], /https:\/\/docs\//);
  assert.equal(s.data[1][7], '');
  assert.match(s.notes['2:9'][0][0], /Fill failed/);
  assert.equal(s.data[1][8], false);
  failFill = false;
});
test('onEdit dán nhiều dòng, bỏ qua tab khác và giới hạn số dòng xuất', () => {
  const s = sheet();
  for (let i = 0; i < 6; i++) s.data.push(['03/11/2026', 'B', 'HCM', 21000, '', '', '', '', true]);
  s.data[1][8] = true;
  const id = '1ShgTg3vFVaQbIE3R1ziQNQrwMMeiKxG_4_2nCXBmeBQ';
  const ss = { getId: () => id, getSpreadsheetTimeZone: () => 'Asia/Ho_Chi_Minh', toast() {} };
  let released = false;
  ctx.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock() { released = true; } }) };
  const before = files.length;
  ctx.ptXuLySua({ source: ss, range: s.getRange(2, 4, 7, 6) });
  assert.equal(files.length - before, 5);
  assert.equal(released, true);
  assert.equal(s.data[7][8], false);
  assert.match(s.notes['8:9'][0][0], /tối đa 5/);
  assert.equal(s.data[7][4], 'Hai mươi mốt nghìn đồng');
  s.getName = () => 'OTHER';
  ctx.ptXuLySua({ source: ss, range: s.getRange(2, 6) });
  assert.equal(files.length - before, 5);
});
test('Lỗi toast không đổi kết quả tạo phiếu thành lỗi', () => {
  const s = sheet();
  ctx.ptTaoAnToan_(s, ctx.ptCot_(s), 2, {
    getSpreadsheetTimeZone: () => 'Asia/Ho_Chi_Minh',
    toast() { throw Error('Cannot call SpreadsheetApp.showNotification() from this context.'); }
  });
  assert.match(s.notes['2:9'][0][0], /^Đã tạo:/);
  assert.match(s.data[1][5], /https:\/\/docs\//);
  assert.equal(s.data[1][8], false);
});
test('Cài đặt độc lập: không có UI vẫn cài được trigger chỉnh sửa', () => {
  const s = sheet(), installed = [];
  s.getMaxRows = () => 2;
  s.getLastRow = () => 2;
  s.setColumnWidth = () => {};
  Range.prototype.setDataValidation = function() { return this; };
  const ss = { getSheetByName: () => s, getId: () => 'test',
    toast() { throw Error('Cannot call SpreadsheetApp.showNotification() from this context.'); } };
  ctx.SpreadsheetApp = {
    openById: () => ss, getActiveSpreadsheet: () => null,
    getUi() { throw Error('Cannot call SpreadsheetApp.getUi() from this context.'); },
    newDataValidation: () => ({ requireCheckbox() { return this; }, build() { return {}; } })
  };
  const oldDrive = ctx.DriveApp;
  ctx.DriveApp = {
    getFileById: () => ({ getMimeType: () => 'application/vnd.google-apps.document' }),
    getFolderById: () => ({ getName: () => 'Phiếu thu' })
  };
  ctx.ScriptApp = { getProjectTriggers: () => [], newTrigger: handler => ({
    forSpreadsheet() { return this; }, onEdit() { return this; }, onOpen() { return this; },
    create() { installed.push(handler); }
  }) };
  let released = false;
  ctx.LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() { released = true; } }) };
  try {
    ctx.ptCaiDat();
    ctx.ptMoMenu();
    assert.deepEqual(installed, ['ptXuLySua']);
    assert.equal(released, true);
  } finally { ctx.DriveApp = oldDrive; }
});
test('Cập nhật: dọn cột cũ trống và giữ nguyên F:I', () => {
  const s = sheet();
  s.data[0].push('TẠO PHIẾU THU', 'Link_Docs_Phieu_Thu', 'Link_PDF_Phieu_Thu', 'Trạng thái tạo phiếu thu');
  s.data[1].push(false, '', '', '');
  ctx.ptDonCotCu_(s);
  assert.equal(s.data[0].length, 9);
  assert.equal(ctx.ptCot_(s).action, 9);
  assert.equal(ctx.ptCot_(s).docs1, 6);
  assert.equal(s.data[1][8], false);
});
test('Cập nhật: giữ cột cũ có link hoặc công thức', () => {
  const s = sheet();
  s.data[0].push('Link_Docs_Phieu_Thu', 'Link_PDF_Phieu_Thu');
  s.data[1].push('https://old/doc', '');
  s.formulas = { '2:11': '=IF(TRUE,"","")' };
  ctx.ptDonCotCu_(s);
  assert.equal(s.data[0].length, 11);
  assert.equal(s.data[1][9], 'https://old/doc');
});
test('Ô link có công thức không bị ghi đè dù hiển thị trống', () => {
  const s = sheet();
  s.formulas = { '2:6': '=IF(TRUE,"","")' };
  ctx.ptTaoPhieu_(s, ctx.ptCot_(s), 2, 'Asia/Ho_Chi_Minh');
  assert.equal(s.data[1][5], '');
  assert.match(s.data[1][6], /https:\/\/docs\//);
  assert.match(files.at(-1).name, /Lan02/);
});
console.log('Passed ' + count + ' tests.');
