// ============================================================
// TokuteiJob — Quản lý ứng viên
// ─────────────────────────────────────────────────────────────
// Cài đặt:
//   1. Tạo Google Spreadsheet mới (riêng với sheet job)
//   2. Paste file này vào Apps Script (Code.gs)
//   3. Tạo file HTML tên "CandidateSidebar", paste nội dung vào
//   4. Reload Sheet → cho phép quyền
//   5. Menu → ⚙️ Cài đặt sheet (chạy 1 lần)
// ============================================================

const SHEET_NAME = 'Ứng viên';

const STATUS_LIST = [
  'Mới nộp',
  'Đang liên hệ',
  'Đã nộp hồ sơ',
  'Phỏng vấn',
  'Đậu ✅',
  'Rớt ❌',
  'Hủy',
];

const STATUS_COLORS = {
  'Mới nộp':       '#e8f0fe',
  'Đang liên hệ':  '#fff8e1',
  'Đã nộp hồ sơ':  '#e8f5e9',
  'Phỏng vấn':     '#f3e5f5',
  'Đậu ✅':        '#c8e6c9',
  'Rớt ❌':        '#ffcdd2',
  'Hủy':           '#eeeeee',
};

const NGUON_LIST = [
  'Facebook', 'Zalo Group', 'Website', 'Giới thiệu', 'TikTok', 'YouTube', 'Khác',
];

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

//  A  Họ tên
//  B  SĐT
//  C  Zalo
//  D  Email
//  E  Ngày sinh
//  F  Giới tính
//  G  Tỉnh / TP (Việt Nam)
//  H  Tiếng Nhật
//  I  Loại đơn hàng
//  J  Ngành nghề
//  K  Tên công việc
//  L  Nguồn
//  M  Ngày nộp hồ sơ
//  N  Trạng thái
//  O  Ghi chú
//  P  Ngày cập nhật  ← tự động

const HEADERS = [
  'Họ tên', 'SĐT', 'Zalo', 'Email', 'Ngày sinh', 'Giới tính', 'Tỉnh / TP (VN)',
  'Tiếng Nhật', 'Loại đơn hàng', 'Ngành nghề', 'Tên công việc',
  'Nguồn', 'Ngày nộp', 'Trạng thái', 'Ghi chú', 'Ngày cập nhật',
];

const COL_WIDTHS = [150, 110, 110, 170, 95, 90, 130, 90, 130, 160, 190, 100, 100, 120, 260, 140];

// ── TỰ ĐỘNG CẬP NHẬT CỘT P KHI SỬA ─────────────────────────
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  const startRow = e.range.getRow();
  const endRow   = startRow + e.range.getNumRows() - 1;
  const col      = e.range.getColumn();

  if (endRow < 2 || col === 16) return;

  const now = new Date();
  for (let r = Math.max(startRow, 2); r <= endRow; r++) {
    sheet.getRange(r, 16).setValue(now).setNumberFormat('dd/MM/yyyy HH:mm');
  }

  // Đổi màu hàng khi trạng thái thay đổi (cột N = 14)
  if (col === 14) {
    for (let r = Math.max(startRow, 2); r <= endRow; r++) {
      const status = sheet.getRange(r, 14).getValue();
      _applyRowColor(sheet, r, status);
    }
  }
}

// ── MENU ─────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('👥 Ứng viên')
    .addItem('➕ Thêm ứng viên mới',          'showCandidateSidebar')
    .addSeparator()
    .addItem('📊 Thống kê nhanh',             'showStats')
    .addSeparator()
    .addItem('⚙️ Cài đặt sheet (chạy 1 lần)', 'setupCandidateSheet')
    .addToUi();
}

// ── SIDEBAR ──────────────────────────────────────────────────
function showCandidateSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('CandidateSidebar')
    .setTitle('Thêm ứng viên')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── CÀI ĐẶT SHEET ────────────────────────────────────────────
function setupCandidateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  // Header
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setValues([HEADERS])
    .setBackground('#1a3c5e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);

  // Column widths
  COL_WIDTHS.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // Data validation: Trạng thái (col N = 14)
  sheet.getRange(2, 14, 5000, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(STATUS_LIST)
        .setAllowInvalid(false)
        .build()
    );

  // Data validation: Loại đơn hàng (col I = 9)
  sheet.getRange(2, 9, 5000, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Tokutei Đầu Việt (VN→JP)', 'Tokutei Đầu Nhật (đang ở JP)', 'Kỹ sư IT'])
        .setAllowInvalid(true)
        .build()
    );

  // Data validation: Tiếng Nhật (col H = 8)
  sheet.getRange(2, 8, 5000, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Không có', 'N5', 'N4', 'N3', 'N2', 'N1'])
        .setAllowInvalid(true)
        .build()
    );

  // Format ngày
  sheet.getRange(2, 5,  5000, 1).setNumberFormat('dd/MM/yyyy');        // Ngày sinh
  sheet.getRange(2, 13, 5000, 1).setNumberFormat('dd/MM/yyyy');        // Ngày nộp
  sheet.getRange(2, 16, 5000, 1).setNumberFormat('dd/MM/yyyy HH:mm'); // Ngày cập nhật

  SpreadsheetApp.getUi().alert(
    '✅ Cài đặt hoàn tất!\n\n' +
    'Menu → ➕ Thêm ứng viên mới để bắt đầu nhập liệu.'
  );
}

// ── LƯU ỨNG VIÊN (gọi từ Sidebar) ───────────────────────────
function saveCandidate(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Chưa cài đặt sheet. Vào menu → ⚙️ Cài đặt sheet.');

  // Parse ngày sinh từ chuỗi YYYY-MM-DD
  let dob = '';
  if (data.ngaySinh) {
    const [y, m, d] = data.ngaySinh.split('-').map(Number);
    dob = new Date(y, m - 1, d);
  }

  const now    = new Date();
  const status = data.trangThai || 'Mới nộp';

  const row = [
    data.hoTen,          // A
    data.sdt,            // B
    data.zalo    || '',  // C
    data.email   || '',  // D
    dob,                 // E
    data.gioiTinh,       // F
    data.tinh    || '',  // G
    data.tiengNhat,      // H
    data.loaiDon,        // I
    data.nganh,          // J
    data.tenCv   || '',  // K
    data.nguon   || '',  // L
    now,                 // M — ngày nộp hồ sơ
    status,              // N
    data.ghiChu  || '',  // O
    now,                 // P — ngày cập nhật
  ];

  const insertAt = sheet.getLastRow() + 1;
  sheet.getRange(insertAt, 1, 1, row.length).setValues([row]);

  // Format ngày
  sheet.getRange(insertAt, 5).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(insertAt, 13).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(insertAt, 16).setNumberFormat('dd/MM/yyyy HH:mm');

  // Màu theo trạng thái
  _applyRowColor(sheet, insertAt, status);
}

function _applyRowColor(sheet, row, status) {
  const color = STATUS_COLORS[status] || '#ffffff';
  sheet.getRange(row, 1, 1, HEADERS.length).setBackground(color);
}

// ── THỐNG KÊ ─────────────────────────────────────────────────
function showStats() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Chưa có dữ liệu ứng viên.');
    return;
  }

  const lastRow  = sheet.getLastRow();
  const statuses = sheet.getRange(2, 14, lastRow - 1, 1).getValues().flat();
  const loais    = sheet.getRange(2,  9, lastRow - 1, 1).getValues().flat();

  const statusCount = {};
  STATUS_LIST.forEach(s => statusCount[s] = 0);
  statuses.forEach(s => { if (s && statusCount[s] !== undefined) statusCount[s]++; });

  const loaiCount = {};
  loais.forEach(l => { if (l) loaiCount[l] = (loaiCount[l] || 0) + 1; });

  const total      = statuses.filter(s => s).length;
  const statusLines = STATUS_LIST.map(s => `  ${s}: ${statusCount[s]}`).join('\n');
  const loaiLines   = Object.entries(loaiCount).map(([k, v]) => `  ${k}: ${v}`).join('\n');

  SpreadsheetApp.getUi().alert(
    `📊 Thống kê ứng viên\n\n` +
    `Tổng: ${total} ứng viên\n\n` +
    `── Theo trạng thái ──\n${statusLines}\n\n` +
    `── Theo loại đơn hàng ──\n${loaiLines}`
  );
}

// ── TRẢ VỀ DỮ LIỆU CHO SIDEBAR ──────────────────────────────
function getCategoryList(type) {
  if (type === 'kysis') return KYSIS_CATS;
  if (type === 'nhat' || type === 'viet') return TOKUTEI_CATS;
  return [];
}
