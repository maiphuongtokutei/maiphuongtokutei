// ============================================================
// TokuteiJob — Quản lý đơn hàng ứng viên + CTV
// ─────────────────────────────────────────────────────────────
// Cài đặt:
//   1. Tạo Google Spreadsheet mới
//   2. Paste file này vào Apps Script (Code.gs)
//   3. Tạo file HTML tên "DonHangSidebar", paste sidebar vào
//   4. Reload sheet → cho phép quyền
//   5. Menu → ⚙️ Cài đặt sheet (chạy 1 lần)
//   6. Thêm CTV vào tab "CTV" trước khi nhập đơn hàng
// ============================================================

const SHEET_DON_HANG  = 'Đơn hàng';
const SHEET_CTV       = 'CTV';
const SHEET_TIEM_NANG = 'Tiềm năng';

// ── CỘT A-Z (index 1-26) ─────────────────────────────────────
// A:  STT           B: Họ tên         C: SĐT          D: Zalo
// E:  Ngày sinh     F: Giới tính      G: Tỉnh/TP(VN)  H: Tiếng Nhật
// I:  Loại đơn      J: Đối tác        K: Ngành nghề   L: Tên công việc
// M:  Trạng thái    N: Ngày đăng ký
// O:  Hoa hồng ĐV   P: T.thái thu ĐV  Q: Ngày thu ĐV
// R:  Tiền back ĐN  S: T.thái nhận ĐN T: Ngày nhận ĐN
// U:  CTV giới thiệu V: Tiền trả CTV  W: T.thái trả CTV X: Ngày trả CTV
// Y:  Ghi chú       Z: Ngày cập nhật

const HEADERS = [
  'STT', 'Họ tên', 'SĐT', 'Zalo', 'Ngày sinh', 'Giới tính', 'Tỉnh/TP (VN)',
  'Tiếng Nhật', 'Loại đơn hàng', 'Đối tác', 'Ngành nghề', 'Tên công việc',
  'Trạng thái', 'Ngày đăng ký',
  'Hoa hồng (Đ.Việt)', 'Thu tiền (Đ.Việt)', 'Ngày thu',
  'Tiền back (Đ.Nhật)', 'Nhận tiền (Đ.Nhật)', 'Ngày nhận',
  'CTV giới thiệu', 'Tiền trả CTV', 'Đã trả CTV', 'Ngày trả CTV',
  'Ghi chú', 'Ngày cập nhật',
];

const COL_WIDTHS = [
  45, 155, 110, 110, 90, 80, 120,
  85, 170, 135, 155, 190,
  150, 90,
  135, 115, 90,
  135, 115, 90,
  135, 115, 100, 90,
  240, 140,
];

// Màu header theo nhóm cột
const HEADER_SECTION_COLORS = [
  // [startCol, endCol, bg, fg]
  [1,  1,  '#455a64', '#fff'],  // STT
  [2,  8,  '#1a3c5e', '#fff'],  // Ứng viên
  [9,  14, '#2e7d32', '#fff'],  // Đơn hàng
  [15, 17, '#bf360c', '#fff'],  // Tài chính đầu Việt
  [18, 20, '#1565c0', '#fff'],  // Tài chính đầu Nhật
  [21, 24, '#6a1b9a', '#fff'],  // CTV
  [25, 26, '#546e7a', '#fff'],  // Meta
];

const TRANG_THAI_DON = [
  'Mới đăng ký',
  'Đang chuẩn bị hồ sơ',
  'Đã nộp hồ sơ',
  'Phỏng vấn',
  'Đậu - Chờ xuất cảnh ✅',
  'Đã xuất cảnh 🎉',
  'Rớt ❌',
  'Tự hủy',
  'Hủy do công ty',
];

const TRANG_THAI_COLORS = {
  'Mới đăng ký':             '#f8f9fa',
  'Đang chuẩn bị hồ sơ':    '#fff8e1',
  'Đã nộp hồ sơ':            '#e8f5e9',
  'Phỏng vấn':               '#f3e5f5',
  'Đậu - Chờ xuất cảnh ✅':  '#dcedc8',
  'Đã xuất cảnh 🎉':         '#c8e6c9',
  'Rớt ❌':                  '#ffcdd2',
  'Tự hủy':                  '#eeeeee',
  'Hủy do công ty':          '#eeeeee',
};

const THANH_TOAN_LIST = ['Chưa', 'Một phần', 'Đã xong ✓'];

const TOKUTEI_CATS = [
  'Chế biến thực phẩm','Nhóm ngành 1 (Cơ khí)','Nhóm ngành 2 (Điện tử)',
  'Xây dựng','Điều dưỡng','Nhà hàng','Vệ sinh tòa nhà','Bảo dưỡng ô tô',
  'Nông nghiệp','Lưu trú / Khách sạn','Ngư nghiệp','Đóng tàu','Hàng không',
  'Vận tải','Lâm nghiệp','May','In ấn',
];

const KYSIS_CATS = [
  'Lập trình viên','Kỹ sư hệ thống / Mạng','AI & Data Science',
  'Cơ khí','Điện - Điện tử','Hóa học & Vật liệu','Ô tô',
  'Phiên dịch / Thông dịch','Xuất nhập khẩu','Quản trị kinh doanh / Marketing',
];

// ── AUTO-UPDATE KHI SỬA ──────────────────────────────────────
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_DON_HANG) return;

  const startRow = e.range.getRow();
  const endRow   = startRow + e.range.getNumRows() - 1;
  const col      = e.range.getColumn();
  if (endRow < 2 || col === 26) return;

  const now = new Date();
  for (let r = Math.max(startRow, 2); r <= endRow; r++) {
    sheet.getRange(r, 26).setValue(now).setNumberFormat('dd/MM/yyyy HH:mm');
    if (col === 13) {
      const status = sheet.getRange(r, 13).getValue();
      _colorRow(sheet, r, status);
    }
  }
}

// ── MENU ─────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 Quản lý')
    .addItem('➕ Thêm đơn hàng mới',          'showDonHangSidebar')
    .addItem('🔖 Lưu ứng viên tiềm năng',     'showTiemNangSidebar')
    .addSeparator()
    .addItem('💰 Thống kê tài chính',          'showStats')
    .addSeparator()
    .addItem('⚙️ Cài đặt sheet (chạy 1 lần)', 'setupSheets')
    .addToUi();
}

// ── SIDEBAR ──────────────────────────────────────────────────
function showDonHangSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('DonHangSidebar')
    .setTitle('Thêm đơn hàng')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showTiemNangSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('TiemNangSidebar')
    .setTitle('Ứng viên tiềm năng')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── CÀI ĐẶT SHEET ────────────────────────────────────────────
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Tab Đơn hàng ──
  let sh = ss.getSheetByName(SHEET_DON_HANG) || ss.insertSheet(SHEET_DON_HANG);

  const hdrRange = sh.getRange(1, 1, 1, HEADERS.length);
  hdrRange.setValues([HEADERS]).setFontWeight('bold').setHorizontalAlignment('center');

  // Màu header theo nhóm
  HEADER_SECTION_COLORS.forEach(([s, e, bg, fg]) => {
    sh.getRange(1, s, 1, e - s + 1).setBackground(bg).setFontColor(fg);
  });

  sh.setFrozenRows(1);
  COL_WIDTHS.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // Data validation
  const N = 5000;
  sh.getRange(2, 13, N, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(TRANG_THAI_DON).setAllowInvalid(false).build()
  );
  sh.getRange(2, 9, N, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Tokutei Đầu Việt (VN→JP)', 'Tokutei Đầu Nhật (đang ở JP)', 'Kỹ sư IT'])
      .setAllowInvalid(true).build()
  );
  [16, 19, 23].forEach(c => {
    sh.getRange(2, c, N, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(THANH_TOAN_LIST).setAllowInvalid(false).build()
    );
  });

  // Format ngày
  sh.getRange(2, 5,  N, 1).setNumberFormat('dd/MM/yyyy');
  sh.getRange(2, 14, N, 1).setNumberFormat('dd/MM/yyyy');
  sh.getRange(2, 17, N, 1).setNumberFormat('dd/MM/yyyy');
  sh.getRange(2, 20, N, 1).setNumberFormat('dd/MM/yyyy');
  sh.getRange(2, 24, N, 1).setNumberFormat('dd/MM/yyyy');
  sh.getRange(2, 26, N, 1).setNumberFormat('dd/MM/yyyy HH:mm');

  // ── Tab CTV ──
  let ctv = ss.getSheetByName(SHEET_CTV) || ss.insertSheet(SHEET_CTV);
  ctv.getRange(1, 1, 1, 6)
    .setValues([['Tên CTV', 'SĐT', 'Zalo', 'Tỉnh/TP', 'Ghi chú', 'Ngày tạo']])
    .setBackground('#6a1b9a').setFontColor('#fff')
    .setFontWeight('bold').setHorizontalAlignment('center');
  ctv.setFrozenRows(1);
  [160, 110, 110, 120, 240, 120].forEach((w, i) => ctv.setColumnWidth(i + 1, w));
  ctv.getRange(2, 6, 5000, 1).setNumberFormat('dd/MM/yyyy');

  // ── Tab Tiềm năng ──
  let tn = ss.getSheetByName(SHEET_TIEM_NANG) || ss.insertSheet(SHEET_TIEM_NANG);
  const tnHeaders = [
    'Họ tên', 'SĐT', 'Zalo', 'Ngày sinh', 'Giới tính', 'Tỉnh/TP (VN)',
    'Tiếng Nhật', 'Loại đơn quan tâm', 'Ngành nghề quan tâm',
    'Yêu cầu lương', 'Ghi chú / Kinh nghiệm', 'Ngày liên hệ', 'Ngày cập nhật',
  ];
  const tnWidths = [150, 110, 110, 90, 80, 120, 85, 170, 200, 130, 270, 100, 140];
  tn.getRange(1, 1, 1, tnHeaders.length)
    .setValues([tnHeaders])
    .setBackground('#37474f').setFontColor('#fff')
    .setFontWeight('bold').setHorizontalAlignment('center');
  tn.setFrozenRows(1);
  tnWidths.forEach((w, i) => tn.setColumnWidth(i + 1, w));
  tn.getRange(2, 4,  5000, 1).setNumberFormat('dd/MM/yyyy');
  tn.getRange(2, 12, 5000, 1).setNumberFormat('dd/MM/yyyy');
  tn.getRange(2, 13, 5000, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  tn.getRange(2, 8, 5000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Tokutei Đầu Việt (VN→JP)', 'Tokutei Đầu Nhật (đang ở JP)', 'Kỹ sư IT', 'Chưa rõ'])
      .setAllowInvalid(true).build()
  );

  SpreadsheetApp.getUi().alert(
    '✅ Cài đặt hoàn tất!\n\n' +
    '📌 Bước tiếp theo:\n' +
    '  1. Vào tab "CTV" → thêm tên cộng tác viên\n' +
    '  2. Menu → ➕ Thêm đơn hàng mới để nhập ứng viên có đơn\n' +
    '  3. Menu → 🔖 Lưu ứng viên tiềm năng để lưu người chưa có job'
  );
}

// ── LƯU ĐƠN HÀNG (gọi từ Sidebar) ───────────────────────────
function saveDonHang(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DON_HANG);
  if (!sheet) throw new Error('Chưa cài đặt sheet. Vào menu → ⚙️ Cài đặt sheet.');

  let dob = '';
  if (data.ngaySinh) {
    const [y, m, d] = data.ngaySinh.split('-').map(Number);
    dob = new Date(y, m - 1, d);
  }

  const now    = new Date();
  const status = data.trangThai || 'Mới đăng ký';
  const stt    = Math.max(1, sheet.getLastRow()); // hàng 1 = header, hàng 2 = entry 1

  const row = [
    stt,                                          // A  STT
    data.hoTen,                                   // B  Họ tên
    data.sdt,                                     // C  SĐT
    data.zalo          || '',                     // D  Zalo
    dob,                                          // E  Ngày sinh
    data.gioiTinh,                                // F  Giới tính
    data.tinh          || '',                     // G  Tỉnh/TP
    data.tiengNhat,                               // H  Tiếng Nhật
    data.loaiDon,                                 // I  Loại đơn
    data.doiTac        || '',                     // J  Đối tác
    data.nganh,                                   // K  Ngành nghề
    data.tenCv         || '',                     // L  Tên công việc
    status,                                       // M  Trạng thái
    now,                                          // N  Ngày đăng ký
    data.hoaHongViet   || '',                     // O  Hoa hồng ĐV
    data.ttThuViet     || 'Chưa',                 // P  T.thái thu ĐV
    '',                                           // Q  Ngày thu ĐV
    data.tienBackNhat  || '',                     // R  Tiền back ĐN
    data.ttNhanNhat    || 'Chưa',                 // S  T.thái nhận ĐN
    '',                                           // T  Ngày nhận ĐN
    data.ctvGioiThieu  || '',                     // U  CTV
    data.tienTraCtv    || '',                     // V  Tiền trả CTV
    data.ctvGioiThieu ? (data.ttTraCtv || 'Chưa') : '', // W  T.thái trả CTV
    '',                                           // X  Ngày trả CTV
    data.ghiChu        || '',                     // Y  Ghi chú
    now,                                          // Z  Ngày cập nhật
  ];

  const insertAt = sheet.getLastRow() + 1;
  sheet.getRange(insertAt, 1, 1, row.length).setValues([row]);
  sheet.getRange(insertAt, 5).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(insertAt, 14).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(insertAt, 26).setNumberFormat('dd/MM/yyyy HH:mm');

  _colorRow(sheet, insertAt, status);
}

function _colorRow(sheet, row, status) {
  const color = TRANG_THAI_COLORS[status] || '#ffffff';
  sheet.getRange(row, 1, 1, HEADERS.length).setBackground(color);
}

// ── LƯU ỨNG VIÊN TIỀM NĂNG ───────────────────────────────────
function saveTiemNang(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TIEM_NANG);
  if (!sheet) throw new Error('Chưa cài đặt sheet. Vào menu → ⚙️ Cài đặt sheet.');

  let dob = '';
  if (data.ngaySinh) {
    const [y, m, d] = data.ngaySinh.split('-').map(Number);
    dob = new Date(y, m - 1, d);
  }

  const now = new Date();
  const row = [
    data.hoTen,
    data.sdt,
    data.zalo         || '',
    dob,
    data.gioiTinh,
    data.tinh         || '',
    data.tiengNhat,
    data.loaiDon      || 'Chưa rõ',
    data.nganh        || '',
    data.ycLuong      || '',
    data.ghiChu       || '',
    now,  // ngày liên hệ
    now,  // ngày cập nhật
  ];

  const insertAt = sheet.getLastRow() + 1;
  sheet.getRange(insertAt, 1, 1, row.length).setValues([row]);
  sheet.getRange(insertAt, 4).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(insertAt, 12).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(insertAt, 13).setNumberFormat('dd/MM/yyyy HH:mm');

  // Highlight nhạt để phân biệt với đơn hàng đã có
  sheet.getRange(insertAt, 1, 1, row.length).setBackground('#fafafa');
}

// ── TRẢ VỀ DANH SÁCH CTV CHO SIDEBAR ────────────────────────
function getCtvList() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CTV);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 2)
    .getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({ ten: String(r[0]).trim(), sdt: String(r[1]).trim() }));
}

// ── THỐNG KÊ TÀI CHÍNH ───────────────────────────────────────
function showStats() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DON_HANG);

  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Chưa có dữ liệu đơn hàng.');
    return;
  }

  const numRows = sheet.getLastRow() - 1;
  const data    = sheet.getRange(2, 1, numRows, 26).getValues();

  const statusCount = {};
  TRANG_THAI_DON.forEach(s => (statusCount[s] = 0));

  let totalHH = 0, chuaThuHH = 0;
  let totalBack = 0, chuaNhanBack = 0;
  let totalCtv = 0, chuaTraCtv = 0;

  const toNum = v => parseFloat(String(v).replace(/[^0-9.]/g, '')) || 0;

  data.forEach(r => {
    const st = r[12]; // M trạng thái
    if (st && statusCount[st] !== undefined) statusCount[st]++;

    const hh = toNum(r[14]); // O hoa hồng ĐV
    if (hh) { totalHH += hh; if (r[15] !== 'Đã xong ✓') chuaThuHH += hh; }

    const back = toNum(r[17]); // R tiền back ĐN
    if (back) { totalBack += back; if (r[18] !== 'Đã xong ✓') chuaNhanBack += back; }

    const ctv = toNum(r[21]); // V tiền trả CTV
    if (ctv) { totalCtv += ctv; if (r[22] !== 'Đã xong ✓') chuaTraCtv += ctv; }
  });

  const fmt = n => n > 0 ? n.toLocaleString('vi-VN') + 'đ' : '—';
  const statusLines = TRANG_THAI_DON.map(s => `  ${s}: ${statusCount[s]}`).join('\n');

  SpreadsheetApp.getUi().alert(
    `💰 Thống kê tài chính\n\n` +
    `── Trạng thái đơn hàng ──\n${statusLines}\n\n` +
    `── Tài chính đầu Việt (hoa hồng) ──\n` +
    `  Tổng:    ${fmt(totalHH)}\n` +
    `  Chưa thu: ${fmt(chuaThuHH)}\n\n` +
    `── Tài chính đầu Nhật (tiền back) ──\n` +
    `  Tổng:     ${fmt(totalBack)}\n` +
    `  Chưa nhận: ${fmt(chuaNhanBack)}\n\n` +
    `── CTV ──\n` +
    `  Tổng phải trả: ${fmt(totalCtv)}\n` +
    `  Chưa trả:      ${fmt(chuaTraCtv)}`
  );
}
