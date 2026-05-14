// ============================================================
// TokuteiJob — Quản lý Đăng ký Phỏng vấn & Thu nhập
// ─────────────────────────────────────────────────────────────
// Cột:
//   A  HỌ VÀ TÊN
//   B  NGUỒN / CTV
//   C  CHỨNG CHỈ
//   D  TIẾNG NHẬT
//   E  ĐƠN HÀNG
//   F  NGUỒN (kênh quảng cáo)
//   G  BACK      (= I − H,  tự động)
//   H  BACK CTV  (nhập tay)
//   I  BACK GỐC  (nhập tay)
//   J  HOA HỒNG  (= K − L,  tự động)
//   K  NET CTV/UV (nhập tay)
//   L  NET GỐC    (nhập tay)
//   M  GHI CHÚ
//   N  TRẠNG THÁI
//   O  NGÀY ĐĂNG KÝ
//   P  NGÀY PHỎNG VẤN
//   Q  NGÀY CẬP NHẬT  ← tự động
// ─────────────────────────────────────────────────────────────
// Cài đặt:
//   1. Mở Google Spreadsheet mới (hoặc dùng chung với sheet job)
//   2. Extensions → Apps Script → paste toàn bộ file này vào Code.gs
//   3. Tạo thêm file HTML tên "InterviewSidebar", paste nội dung phía dưới
//   4. Lưu → Reload Sheet → Cho phép quyền
//   5. Menu → ⚙️ Cài đặt sheet (chạy 1 lần)
// ============================================================

const IV_SHEET = 'Đăng ký PV & Thu nhập';

const IV_HEADERS = [
  'HỌ VÀ TÊN',       // A  1
  'NGUỒN / CTV',      // B  2
  'CHỨNG CHỈ',        // C  3
  'TIẾNG NHẬT',       // D  4
  'ĐƠN HÀNG',         // E  5
  'NGUỒN',            // F  6
  'BACK',             // G  7  ← = I − H (tự động)
  'BACK CTV',         // H  8
  'BACK GỐC',         // I  9
  'HOA HỒNG',         // J  10  ← = K − L (tự động)
  'NET CTV/UV',        // K  11
  'NET GỐC',          // L  12
  'GHI CHÚ',          // M  13
  'TRẠNG THÁI',       // N  14
  'NGÀY ĐĂNG KÝ',     // O  15
  'NGÀY PHỎNG VẤN',   // P  16
  'NGÀY CẬP NHẬT',    // Q  17
];

const IV_COL_WIDTHS = [160, 140, 140, 100, 180, 110, 110, 110, 110, 110, 110, 110, 220, 130, 110, 110, 130];

const IV_STATUS_LIST = [
  'Mới đăng ký',
  'Đang liên hệ',
  'Đã nộp hồ sơ',
  'Lịch phỏng vấn',
  'Đậu PV ✅',
  'Rớt PV ❌',
  'Đang chờ xuất cảnh',
  'Đang làm việc',
  'Hoàn thành',
  'Hủy',
];

const IV_STATUS_COLORS = {
  'Mới đăng ký':         '#e8f0fe',
  'Đang liên hệ':        '#fff8e1',
  'Đã nộp hồ sơ':        '#e3f2fd',
  'Lịch phỏng vấn':      '#f3e5f5',
  'Đậu PV ✅':           '#c8e6c9',
  'Rớt PV ❌':           '#ffcdd2',
  'Đang chờ xuất cảnh':  '#fff9c4',
  'Đang làm việc':       '#a5d6a7',
  'Hoàn thành':          '#b2dfdb',
  'Hủy':                 '#eeeeee',
};

const IV_NGUON_LIST = [
  'Facebook', 'Zalo Group', 'TikTok', 'YouTube',
  'Website', 'Giới thiệu', 'CTV', 'Khác',
];

const IV_TIENG_NHAT = ['Không có', 'N5', 'N4', 'N3', 'N2', 'N1', 'Đang học'];

const IV_CHUNG_CHI = [
  'JLPT N5', 'JLPT N4', 'JLPT N3', 'JLPT N2', 'JLPT N1',
  'NAT-TEST 5Q', 'NAT-TEST 4Q', 'NAT-TEST 3Q',
  'JLCT', 'JFT-Basic',
  'Chứng chỉ kỹ năng Tokutei',
  'Chứng chỉ kỹ năng đặc định ngành điều dưỡng',
  'Chứng chỉ kỹ năng ngành chế biến thực phẩm',
  'Chưa có',
];

const IV_DON_HANG_LIST = [
  'Tokutei Đầu Việt',
  'Tokutei Đầu Nhật',
  'Kỹ sư',
];

// ── TỰ ĐỘNG: CẬP NHẬT NGÀY & HOA HỒNG ───────────────────────
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== IV_SHEET) return;

  const startRow = e.range.getRow();
  const endRow   = startRow + e.range.getNumRows() - 1;
  const col      = e.range.getColumn();

  if (endRow < 2) return;

  const now = new Date();

  for (let r = Math.max(startRow, 2); r <= endRow; r++) {
    // Cột Q (17) = ngày cập nhật tự động
    if (col !== 17) {
      sheet.getRange(r, 17).setValue(now).setNumberFormat('dd/MM/yyyy HH:mm');
    }

    // G = I − H: khi sửa BACK CTV (H=8) hoặc BACK GỐC (I=9)
    if (col === 8 || col === 9) {
      const backCtv = Number(sheet.getRange(r, 8).getValue()) || 0;
      const backGoc = Number(sheet.getRange(r, 9).getValue()) || 0;
      sheet.getRange(r, 7).setValue(backGoc - backCtv);
    }

    // J = K − L: khi sửa NET CTV/UV (K=11) hoặc NET GỐC (L=12)
    if (col === 11 || col === 12) {
      const netCtv = Number(sheet.getRange(r, 11).getValue()) || 0;
      const netGoc = Number(sheet.getRange(r, 12).getValue()) || 0;
      sheet.getRange(r, 10).setValue(netCtv - netGoc);
    }

    // Màu hàng theo trạng thái (cột N=14)
    if (col === 14) {
      const status = sheet.getRange(r, 14).getValue();
      _applyRowColor(sheet, r, status);
    }
  }
}


function _applyRowColor(sheet, row, status) {
  const color = IV_STATUS_COLORS[status] || '#ffffff';
  sheet.getRange(row, 1, 1, IV_HEADERS.length).setBackground(color);
}

// ── MENU ─────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 PV & Thu nhập')
    .addItem('➕ Thêm ứng viên mới',             'showInterviewSidebar')
    .addSeparator()
    .addItem('📊 Thống kê nhanh',                'showIvStats')
    .addItem('💰 Báo cáo thu nhập',              'showIncomeReport')
    .addSeparator()
    .addItem('⚙️ Cài đặt sheet (chạy 1 lần)',   'setupInterviewSheet')
    .addToUi();
}

// ── CÀI ĐẶT SHEET (chạy 1 lần) ───────────────────────────────
function setupInterviewSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(IV_SHEET);
  if (!sheet) sheet = ss.insertSheet(IV_SHEET);

  // ── Header row ──
  const hdr = sheet.getRange(1, 1, 1, IV_HEADERS.length);
  hdr.setValues([IV_HEADERS])
     .setBackground('#1a3c5e')
     .setFontColor('#ffffff')
     .setFontWeight('bold')
     .setHorizontalAlignment('center')
     .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  // ── Màu nền cho nhóm cột tài chính (G-L) ──
  sheet.getRange(1, 7, 1, 6)
       .setBackground('#0d3d2c')
       .setFontColor('#a8f0c6');

  // ── Độ rộng cột ──
  IV_COL_WIDTHS.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  const MAX_ROWS = 2000;

  // ── Data validation ──
  _setDropdown(sheet, 2, 14, MAX_ROWS, IV_STATUS_LIST, false);      // TRẠNG THÁI
  _setDropdown(sheet, 2, 4,  MAX_ROWS, IV_TIENG_NHAT,  true);       // TIẾNG NHẬT
  _setDropdown(sheet, 2, 3,  MAX_ROWS, IV_CHUNG_CHI,   true);       // CHỨNG CHỈ
  _setDropdown(sheet, 2, 5,  MAX_ROWS, IV_DON_HANG_LIST, true);     // ĐƠN HÀNG
  _setDropdown(sheet, 2, 6,  MAX_ROWS, IV_NGUON_LIST,  true);       // NGUỒN

  // ── Định dạng ngày ──
  sheet.getRange(2, 15, MAX_ROWS, 1).setNumberFormat('dd/MM/yyyy');        // NGÀY ĐĂNG KÝ
  sheet.getRange(2, 16, MAX_ROWS, 1).setNumberFormat('dd/MM/yyyy');        // NGÀY PHỎNG VẤN
  sheet.getRange(2, 17, MAX_ROWS, 1).setNumberFormat('dd/MM/yyyy HH:mm'); // NGÀY CẬP NHẬT

  // ── Định dạng tiền (G-L) ──
  sheet.getRange(2,  7, MAX_ROWS, 6).setNumberFormat('#,##0');

  // ── Ghi chú header ──
  sheet.getRange(1, 10).setNote('HOA HỒNG = BACK − BACK CTV − BACK GỐC (tự động tính khi nhập số)');

  SpreadsheetApp.getUi().alert(
    '✅ Cài đặt hoàn tất!\n\n' +
    'Tab "' + IV_SHEET + '" đã sẵn sàng.\n\n' +
    'Menu → ➕ Thêm ứng viên mới để bắt đầu nhập liệu.'
  );
}

function _setDropdown(sheet, startRow, col, numRows, list, allowInvalid) {
  sheet.getRange(startRow, col, numRows, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(list)
        .setAllowInvalid(allowInvalid)
        .build()
    );
}

// ── SIDEBAR ──────────────────────────────────────────────────
function showInterviewSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('InterviewSidebar')
    .setTitle('Thêm ứng viên PV')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── LƯU ỨNG VIÊN (gọi từ Sidebar) ───────────────────────────
function saveInterviewCandidate(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(IV_SHEET);
  if (!sheet) throw new Error('Chưa cài đặt sheet. Vào menu → ⚙️ Cài đặt sheet.');

  const backCtv = Number(data.backCtv) || 0;
  const backGoc = Number(data.backGoc) || 0;
  const back    = backGoc - backCtv;           // G = I − H

  const netCtv  = Number(data.netCtv)  || 0;
  const netGoc  = Number(data.netGoc)  || 0;
  const hoaHong = netCtv - netGoc;             // J = K − L

  const now    = new Date();
  const ngayDk = data.ngayDangKy ? _parseDate(data.ngayDangKy) : now;
  const ngayPv = data.ngayPv     ? _parseDate(data.ngayPv)     : '';

  const status = data.trangThai || 'Mới đăng ký';

  const row = [
    data.hoTen      || '',   // A
    data.nguonCtv   || '',   // B
    data.chungChi   || '',   // C
    data.tiengNhat  || '',   // D
    data.donHang    || '',   // E
    data.nguon      || '',   // F
    back,                    // G
    backCtv,                 // H
    backGoc,                 // I
    hoaHong,                 // J
    Number(data.netCtv)  || 0, // K
    Number(data.netGoc)  || 0, // L
    data.ghiChu     || '',   // M
    status,                  // N
    ngayDk,                  // O
    ngayPv,                  // P
    now,                     // Q
  ];

  const insertAt = sheet.getLastRow() + 1;
  sheet.getRange(insertAt, 1, 1, row.length).setValues([row]);

  // Format ngày
  sheet.getRange(insertAt, 15).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(insertAt, 16).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(insertAt, 17).setNumberFormat('dd/MM/yyyy HH:mm');

  // Format tiền
  sheet.getRange(insertAt, 7, 1, 6).setNumberFormat('#,##0');

  // Màu theo trạng thái
  _applyRowColor(sheet, insertAt, status);

  return { success: true, row: insertAt };
}

function _parseDate(str) {
  const [y, m, d] = String(str).split('-').map(Number);
  return new Date(y, m - 1, d);
}

// ── THỐNG KÊ NHANH ───────────────────────────────────────────
function showIvStats() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(IV_SHEET);

  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Chưa có dữ liệu.');
    return;
  }

  const lastRow  = sheet.getLastRow();
  const data     = sheet.getRange(2, 1, lastRow - 1, IV_HEADERS.length).getValues()
                        .filter(r => r[0]);

  const total       = data.length;
  const statusCount = {};
  IV_STATUS_LIST.forEach(s => statusCount[s] = 0);
  data.forEach(r => { const s = r[13]; if (statusCount[s] !== undefined) statusCount[s]++; });

  const statusLines = IV_STATUS_LIST
    .filter(s => statusCount[s] > 0)
    .map(s => `  ${s}: ${statusCount[s]}`)
    .join('\n');

  SpreadsheetApp.getUi().alert(
    `📊 Thống kê đăng ký phỏng vấn\n\n` +
    `Tổng ứng viên: ${total}\n\n` +
    `── Theo trạng thái ──\n${statusLines || '  (Chưa có dữ liệu)'}`
  );
}

// ── BÁO CÁO THU NHẬP ─────────────────────────────────────────
function showIncomeReport() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(IV_SHEET);

  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Chưa có dữ liệu.');
    return;
  }

  const lastRow = sheet.getLastRow();
  const data    = sheet.getRange(2, 1, lastRow - 1, IV_HEADERS.length).getValues()
                       .filter(r => r[0]);

  let totalBack    = 0;
  let totalBackCtv = 0;
  let totalBackGoc = 0;
  let totalHoaHong = 0;
  let totalNetCtv  = 0;
  let totalNetGoc  = 0;
  let countDone    = 0;

  data.forEach(r => {
    totalBack    += Number(r[6])  || 0;
    totalBackCtv += Number(r[7])  || 0;
    totalBackGoc += Number(r[8])  || 0;
    totalHoaHong += Number(r[9])  || 0;
    totalNetCtv  += Number(r[10]) || 0;
    totalNetGoc  += Number(r[11]) || 0;
    const s = String(r[13]);
    if (s === 'Đang làm việc' || s === 'Hoàn thành') countDone++;
  });

  function fmt(n) {
    return n.toLocaleString('vi-VN') + ' đ';
  }

  SpreadsheetApp.getUi().alert(
    `💰 Báo cáo thu nhập\n\n` +
    `Tổng ứng viên có dữ liệu tài chính: ${data.filter(r => r[6]).length}\n` +
    `Đang làm việc / Hoàn thành: ${countDone}\n\n` +
    `── Tổng cộng ──\n` +
    `  BACK:       ${fmt(totalBack)}\n` +
    `  BACK CTV:   ${fmt(totalBackCtv)}\n` +
    `  BACK GỐC:   ${fmt(totalBackGoc)}\n` +
    `  HOA HỒNG:   ${fmt(totalHoaHong)}\n` +
    `  NET CTV/UV: ${fmt(totalNetCtv)}\n` +
    `  NET GỐC:    ${fmt(totalNetGoc)}`
  );
}
