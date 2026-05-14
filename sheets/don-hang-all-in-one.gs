// ============================================================
// TokuteiJob — Quản lý đơn hàng + CTV + Tiềm năng
// ─────────────────────────────────────────────────────────────
// Cài đặt (CHỈ 1 FILE DUY NHẤT):
//   1. Tạo Google Spreadsheet mới
//   2. Extensions → Apps Script → paste toàn bộ vào Code.gs
//   3. Ctrl+S → Reload sheet → cho phép quyền
//   4. Menu → ⚙️ Cài đặt sheet (chạy 1 lần)
// ============================================================

const SHEET_DON_HANG  = 'Đơn hàng';
const SHEET_CTV       = 'CTV';
const SHEET_TIEM_NANG = 'Tiềm năng';

// Cột A-X (1-24)
// A: STT          B: Họ tên        C: Kênh        D: Liên lạc
// E: Giới tính    F: Tiếng Nhật
// G: Loại đơn     H: Đối tác       I: Ngành        J: Tên CV
// K: Trạng thái   L: Ngày đăng ký
// M: Hoa hồng ĐV  N: Thu tiền ĐV   O: Ngày thu ĐV
// P: Tiền back ĐN Q: Nhận tiền ĐN  R: Ngày nhận ĐN
// S: CTV          T: Tiền trả CTV  U: Đã trả CTV   V: Ngày trả CTV
// W: Ghi chú      X: Ngày cập nhật

const HEADERS = [
  'STT', 'Họ tên', 'Kênh', 'Liên lạc (Zalo/Facebook)',
  'Giới tính', 'Tiếng Nhật',
  'Loại đơn hàng', 'Đối tác', 'Ngành nghề', 'Tên công việc',
  'Trạng thái', 'Ngày đăng ký',
  'Hoa hồng (Đ.Việt)', 'Thu tiền (Đ.Việt)', 'Ngày thu',
  'Tiền back (Đ.Nhật)', 'Nhận tiền (Đ.Nhật)', 'Ngày nhận',
  'CTV giới thiệu', 'Tiền trả CTV', 'Đã trả CTV', 'Ngày trả CTV',
  'Ghi chú', 'Ngày cập nhật',
];

const COL_WIDTHS = [
  45, 155, 90, 200,
  80, 85,
  170, 135, 155, 200,
  150, 90,
  135, 115, 90,
  135, 115, 90,
  135, 115, 100, 90,
  250, 140,
];

const HEADER_COLORS = [
  [1,  1,  '#455a64'],  // STT
  [2,  6,  '#1a3c5e'],  // Ứng viên
  [7,  12, '#2e7d32'],  // Đơn hàng
  [13, 15, '#bf360c'],  // Tài chính ĐV
  [16, 18, '#1565c0'],  // Tài chính ĐN
  [19, 22, '#6a1b9a'],  // CTV
  [23, 24, '#546e7a'],  // Meta
];

const TRANG_THAI_DON = [
  'Mới đăng ký', 'Đang chuẩn bị hồ sơ', 'Đã nộp hồ sơ',
  'Phỏng vấn', 'Đậu - Chờ xuất cảnh ✅', 'Đã xuất cảnh 🎉',
  'Rớt ❌', 'Tự hủy', 'Hủy do công ty',
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

// ── AUTO-UPDATE KHI SỬA ──────────────────────────────────────
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_DON_HANG) return;

  var startRow = e.range.getRow();
  var endRow   = startRow + e.range.getNumRows() - 1;
  var col      = e.range.getColumn();
  if (endRow < 2 || col === 24) return;

  var now = new Date();
  for (var r = Math.max(startRow, 2); r <= endRow; r++) {
    sheet.getRange(r, 24).setValue(now).setNumberFormat('dd/MM/yyyy HH:mm');
    if (col === 11) {
      var status = sheet.getRange(r, 11).getValue();
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
  var html = HtmlService.createHtmlOutput(HTML_DON_HANG)
    .setTitle('Thêm đơn hàng').setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showTiemNangSidebar() {
  var html = HtmlService.createHtmlOutput(HTML_TIEM_NANG)
    .setTitle('Ứng viên tiềm năng').setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── CÀI ĐẶT SHEET ────────────────────────────────────────────
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var N  = 5000;

  // ── Tab Đơn hàng ──
  var sh = ss.getSheetByName(SHEET_DON_HANG) || ss.insertSheet(SHEET_DON_HANG);
  sh.getRange(1, 1, 1, HEADERS.length)
    .setValues([HEADERS]).setFontWeight('bold').setFontColor('#fff').setHorizontalAlignment('center');
  HEADER_COLORS.forEach(function(c) {
    sh.getRange(1, c[0], 1, c[1] - c[0] + 1).setBackground(c[2]);
  });
  sh.setFrozenRows(1);
  COL_WIDTHS.forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });

  // Validations
  sh.getRange(2, 3, N, 1).setDataValidation(   // Kênh
    SpreadsheetApp.newDataValidation().requireValueInList(['Zalo', 'Facebook']).setAllowInvalid(false).build());
  sh.getRange(2, 7, N, 1).setDataValidation(   // Loại đơn
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Tokutei Đầu Việt (VN→JP)', 'Tokutei Đầu Nhật (đang ở JP)', 'Kỹ sư IT'])
      .setAllowInvalid(true).build());
  sh.getRange(2, 11, N, 1).setDataValidation(  // Trạng thái
    SpreadsheetApp.newDataValidation().requireValueInList(TRANG_THAI_DON).setAllowInvalid(false).build());
  [14, 17, 21].forEach(function(c) {           // Thanh toán
    sh.getRange(2, c, N, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(THANH_TOAN_LIST).setAllowInvalid(false).build());
  });

  // Date formats
  sh.getRange(2, 12, N, 1).setNumberFormat('dd/MM/yyyy');       // Ngày đăng ký
  sh.getRange(2, 15, N, 1).setNumberFormat('dd/MM/yyyy');       // Ngày thu ĐV
  sh.getRange(2, 18, N, 1).setNumberFormat('dd/MM/yyyy');       // Ngày nhận ĐN
  sh.getRange(2, 22, N, 1).setNumberFormat('dd/MM/yyyy');       // Ngày trả CTV
  sh.getRange(2, 24, N, 1).setNumberFormat('dd/MM/yyyy HH:mm'); // Ngày cập nhật

  // ── Tab CTV ──
  var ctv = ss.getSheetByName(SHEET_CTV) || ss.insertSheet(SHEET_CTV);
  ctv.getRange(1, 1, 1, 5)
    .setValues([['Tên CTV', 'Zalo / Facebook', 'Liên lạc', 'Ghi chú', 'Ngày tạo']])
    .setBackground('#6a1b9a').setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
  ctv.setFrozenRows(1);
  [160, 100, 200, 240, 120].forEach(function(w, i) { ctv.setColumnWidth(i + 1, w); });
  ctv.getRange(2, 5, N, 1).setNumberFormat('dd/MM/yyyy');

  // ── Tab Tiềm năng ──
  var tn  = ss.getSheetByName(SHEET_TIEM_NANG) || ss.insertSheet(SHEET_TIEM_NANG);
  var tnH = ['Họ tên', 'Kênh', 'Liên lạc', 'Giới tính', 'Tiếng Nhật',
             'Loại đơn quan tâm', 'Ngành nghề quan tâm',
             'Yêu cầu lương', 'Ghi chú / Kinh nghiệm', 'Ngày liên hệ', 'Ngày cập nhật'];
  tn.getRange(1, 1, 1, tnH.length)
    .setValues([tnH]).setBackground('#37474f').setFontColor('#fff')
    .setFontWeight('bold').setHorizontalAlignment('center');
  tn.setFrozenRows(1);
  [155, 85, 200, 80, 85, 170, 200, 130, 270, 100, 140]
    .forEach(function(w, i) { tn.setColumnWidth(i + 1, w); });
  tn.getRange(2, 2, N, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Zalo', 'Facebook']).setAllowInvalid(false).build());
  tn.getRange(2, 10, N, 1).setNumberFormat('dd/MM/yyyy');
  tn.getRange(2, 11, N, 1).setNumberFormat('dd/MM/yyyy HH:mm');

  SpreadsheetApp.getUi().alert(
    '✅ Cài đặt hoàn tất!\n\n' +
    '📌 Bước tiếp theo:\n' +
    '  1. Vào tab "CTV" → thêm cộng tác viên\n' +
    '  2. Menu → ➕ Thêm đơn hàng mới\n' +
    '  3. Menu → 🔖 Lưu ứng viên tiềm năng'
  );
}

// ── LƯU ĐƠN HÀNG ─────────────────────────────────────────────
function saveDonHang(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_DON_HANG);
  if (!sheet) throw new Error('Chưa cài đặt sheet. Vào menu → ⚙️ Cài đặt sheet.');

  var now    = new Date();
  var status = data.trangThai || 'Mới đăng ký';
  var stt    = Math.max(1, sheet.getLastRow());

  var row = [
    stt,                                              // A  STT
    data.hoTen,                                       // B  Họ tên
    data.kenh,                                        // C  Kênh
    data.lienLac,                                     // D  Liên lạc
    data.gioiTinh,                                    // E  Giới tính
    data.tiengNhat,                                   // F  Tiếng Nhật
    data.loaiDon,                                     // G  Loại đơn
    data.doiTac        || '',                         // H  Đối tác
    data.nganh,                                       // I  Ngành
    data.tenCv         || '',                         // J  Tên CV
    status,                                           // K  Trạng thái
    now,                                              // L  Ngày đăng ký
    data.hoaHongViet   || '',                         // M  Hoa hồng ĐV
    data.ttThuViet     || 'Chưa',                     // N  Thu tiền ĐV
    '',                                               // O  Ngày thu
    data.tienBackNhat  || '',                         // P  Tiền back ĐN
    data.ttNhanNhat    || 'Chưa',                     // Q  Nhận tiền ĐN
    '',                                               // R  Ngày nhận
    data.ctvGioiThieu  || '',                         // S  CTV
    data.tienTraCtv    || '',                         // T  Tiền trả CTV
    data.ctvGioiThieu ? (data.ttTraCtv || 'Chưa') : '', // U  Đã trả CTV
    '',                                               // V  Ngày trả CTV
    data.ghiChu        || '',                         // W  Ghi chú
    now,                                              // X  Ngày cập nhật
  ];

  var at = sheet.getLastRow() + 1;
  sheet.getRange(at, 1, 1, row.length).setValues([row]);
  sheet.getRange(at, 12).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(at, 24).setNumberFormat('dd/MM/yyyy HH:mm');
  _colorRow(sheet, at, status);
}

function _colorRow(sheet, row, status) {
  sheet.getRange(row, 1, 1, HEADERS.length)
    .setBackground(TRANG_THAI_COLORS[status] || '#ffffff');
}

// ── LƯU TIỀM NĂNG ────────────────────────────────────────────
function saveTiemNang(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_TIEM_NANG);
  if (!sheet) throw new Error('Chưa cài đặt sheet. Vào menu → ⚙️ Cài đặt sheet.');

  var now = new Date();
  var row = [
    data.hoTen, data.kenh, data.lienLac,
    data.gioiTinh, data.tiengNhat,
    data.loaiDon || 'Chưa rõ', data.nganh || '',
    data.ycLuong || '', data.ghiChu || '',
    now, now,
  ];

  var at = sheet.getLastRow() + 1;
  sheet.getRange(at, 1, 1, row.length).setValues([row]);
  sheet.getRange(at, 10).setNumberFormat('dd/MM/yyyy');
  sheet.getRange(at, 11).setNumberFormat('dd/MM/yyyy HH:mm');
}

// ── DANH SÁCH CTV ────────────────────────────────────────────
function getCtvList() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CTV);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues()
    .filter(function(r) { return String(r[0]).trim(); })
    .map(function(r) { return { ten: String(r[0]).trim(), lienLac: String(r[2]).trim() }; });
}

// ── THỐNG KÊ ─────────────────────────────────────────────────
function showStats() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_DON_HANG);
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Chưa có dữ liệu đơn hàng.'); return;
  }

  var data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 24).getValues();
  var sc    = {}; TRANG_THAI_DON.forEach(function(s) { sc[s] = 0; });
  var totalHH = 0, chuaThuHH = 0, totalBK = 0, chuaNhanBK = 0, totalCtv = 0, chuaTraCtv = 0;
  var toN   = function(v) { return parseFloat(String(v).replace(/[^0-9.]/g, '')) || 0; };

  data.forEach(function(r) {
    if (r[10] && sc[r[10]] !== undefined) sc[r[10]]++;    // K trạng thái
    var hh = toN(r[12]); if (hh) { totalHH  += hh;  if (r[13] !== 'Đã xong ✓') chuaThuHH  += hh; }  // M,N
    var bk = toN(r[15]); if (bk) { totalBK  += bk;  if (r[16] !== 'Đã xong ✓') chuaNhanBK += bk; }  // P,Q
    var cv = toN(r[19]); if (cv) { totalCtv += cv;  if (r[20] !== 'Đã xong ✓') chuaTraCtv += cv; }  // T,U
  });

  var fmt = function(n) { return n > 0 ? n.toLocaleString('vi-VN') + 'đ' : '—'; };
  var sl  = TRANG_THAI_DON.map(function(s) { return '  ' + s + ': ' + sc[s]; }).join('\n');

  SpreadsheetApp.getUi().alert(
    '💰 Thống kê tài chính\n\n' +
    '── Trạng thái đơn hàng ──\n' + sl + '\n\n' +
    '── Tài chính đầu Việt ──\n' +
    '  Tổng hoa hồng: ' + fmt(totalHH)  + '\n  Chưa thu:       ' + fmt(chuaThuHH)  + '\n\n' +
    '── Tài chính đầu Nhật ──\n' +
    '  Tổng tiền back: ' + fmt(totalBK) + '\n  Chưa nhận:      ' + fmt(chuaNhanBK) + '\n\n' +
    '── CTV ──\n' +
    '  Tổng phải trả: ' + fmt(totalCtv) + '\n  Chưa trả:       ' + fmt(chuaTraCtv)
  );
}

// ════════════════════════════════════════════════════════════
// HTML — SIDEBAR ĐƠN HÀNG
// ════════════════════════════════════════════════════════════
var HTML_DON_HANG = '<!DOCTYPE html><html><head><base target="_top"><style>' +
'*{box-sizing:border-box;margin:0;padding:0}body{font-family:\'Google Sans\',Arial,sans-serif;font-size:13px;color:#202124;background:#f8f9fa}' +
'.hdr{background:#1a3c5e;color:#fff;padding:13px 16px;font-size:15px;font-weight:600;position:sticky;top:0;z-index:10}' +
'.fb{padding:12px 16px 80px}' +
'.sl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;margin:16px 0 8px;padding:5px 8px;border-radius:4px}' +
'.sl-uv{background:#e8f0fe;color:#1a3c5e}.sl-don{background:#e8f5e9;color:#2e7d32}' +
'.sl-viet{background:#fbe9e7;color:#bf360c}.sl-nhat{background:#e3f2fd;color:#1565c0}' +
'.sl-ctv{background:#ede7f6;color:#6a1b9a}.sl-note{background:#eceff1;color:#546e7a}' +
'.f{margin-bottom:10px}' +
'label{display:block;font-size:11px;font-weight:600;color:#5f6368;text-transform:uppercase;letter-spacing:.04em;margin-bottom:4px}' +
'.req{color:#e8305a;margin-left:2px}' +
'input,select,textarea{width:100%;border:1px solid #dadce0;border-radius:6px;padding:7px 10px;font-size:13px;color:#202124;background:#fff;outline:none;font-family:inherit}' +
'input:focus,select:focus,textarea:focus{border-color:#2D2CDB;box-shadow:0 0 0 2px rgba(45,44,219,.12)}' +
'textarea{resize:vertical;min-height:68px;line-height:1.5}' +
'.r2{display:grid;grid-template-columns:1fr 1fr;gap:10px}' +
'.sb{display:flex;gap:4px;flex-wrap:wrap}' +
'.chip{padding:4px 10px;border-radius:99px;font-size:11px;font-weight:600;cursor:pointer;border:2px solid transparent;white-space:nowrap}' +
'.chip.active{border-color:#2D2CDB!important}' +
'.c-moi{background:#f8f9fa;color:#546e7a}.c-cbhs{background:#fff8e1;color:#e65100}' +
'.c-dhs{background:#e8f5e9;color:#2e7d32}.c-pv{background:#f3e5f5;color:#7b1fa2}' +
'.c-dau{background:#dcedc8;color:#33691e}.c-xc{background:#c8e6c9;color:#1b5e20}' +
'.c-rot{background:#ffcdd2;color:#b71c1c}.c-huy{background:#eee;color:#616161}' +
'.tr{display:flex;align-items:center;gap:8px;padding:6px 0;cursor:pointer}' +
'.tr input[type=checkbox]{width:auto;accent-color:#6a1b9a}' +
'.tr span{font-size:12px;font-weight:600;color:#6a1b9a}' +
'.fblock,.cblock{display:none}.fblock.show,.cblock.show{display:block}' +
'.ft{position:fixed;bottom:0;left:0;right:0;padding:10px 16px;background:#fff;border-top:1px solid #e8eaed;display:flex;gap:8px}' +
'.bs{flex:1;background:#2D2CDB;color:#fff;border:none;border-radius:8px;padding:10px;font-size:13px;font-weight:600;cursor:pointer}' +
'.bs:hover{background:#1e1db8}.bs:disabled{background:#9aa0a6;cursor:not-allowed}' +
'.bc{background:transparent;color:#5f6368;border:1px solid #dadce0;border-radius:8px;padding:10px 14px;font-size:13px;cursor:pointer}' +
'.bc:hover{background:#f1f3f4}' +
'.toast{position:fixed;bottom:68px;left:16px;right:16px;padding:10px 14px;border-radius:8px;font-size:13px;font-weight:500;display:none;z-index:100;text-align:center}' +
'.toast.success{background:#e6f4ea;color:#137333;border:1px solid #b7dfbf}' +
'.toast.error{background:#fce8e6;color:#c5221f;border:1px solid #f5c6c4}' +
'.nh{font-size:11px;color:#e65100;margin-top:4px;display:none}' +
'</style></head><body>' +
'<div class="hdr">📋 Thêm đơn hàng mới</div>' +
'<div class="fb">' +

'<div class="sl sl-uv">Thông tin ứng viên</div>' +
'<div class="f"><label>Họ và tên <span class="req">*</span></label><input type="text" id="hoTen" placeholder="VD: Nguyễn Văn An"></div>' +
'<div class="r2">' +
  '<div class="f"><label>Kênh liên lạc <span class="req">*</span></label><select id="kenh" onchange="onKenh()"><option value="Zalo">Zalo</option><option value="Facebook">Facebook</option></select></div>' +
  '<div class="f"><label id="lbl-ll">Tên Zalo <span class="req">*</span></label><input type="text" id="lienLac" placeholder="@tên hoặc số điện thoại"></div>' +
'</div>' +
'<div class="r2">' +
  '<div class="f"><label>Giới tính</label><select id="gioiTinh"><option>Nam</option><option>Nữ</option></select></div>' +
  '<div class="f"><label>Tiếng Nhật</label><select id="tiengNhat"><option value="Không có">Không có</option><option>N5</option><option>N4</option><option>N3</option><option>N2</option><option>N1</option></select></div>' +
'</div>' +

'<div class="sl sl-don">Thông tin đơn hàng</div>' +
'<div class="r2">' +
  '<div class="f"><label>Loại đơn <span class="req">*</span></label><select id="loaiDon" onchange="onLD()"><option value="">-- Chọn --</option><option value="viet">Tokutei Đầu Việt (VN→JP)</option><option value="nhat">Tokutei Đầu Nhật (đang ở JP)</option><option value="kysis">Kỹ sư IT</option></select></div>' +
  '<div class="f"><label>Đối tác</label><input type="text" id="doiTac" placeholder="VD: Công ty ABC Nhật"></div>' +
'</div>' +
'<div class="f"><label>Ngành nghề <span class="req">*</span></label><select id="nganh"><option value="">-- Chọn loại đơn trước --</option></select></div>' +
'<div class="f"><label>Tên công việc cụ thể</label><input type="text" id="tenCv" placeholder="VD: Công nhân chế biến thực phẩm tại Aichi"></div>' +
'<div class="f"><label>Trạng thái</label><div class="sb">' +
'<span class="chip c-moi active" data-val="Mới đăng ký" onclick="selSt(this)">Mới</span>' +
'<span class="chip c-cbhs" data-val="Đang chuẩn bị hồ sơ" onclick="selSt(this)">Chuẩn bị HS</span>' +
'<span class="chip c-dhs" data-val="Đã nộp hồ sơ" onclick="selSt(this)">Đã nộp HS</span>' +
'<span class="chip c-pv" data-val="Phỏng vấn" onclick="selSt(this)">Phỏng vấn</span>' +
'<span class="chip c-dau" data-val="Đậu - Chờ xuất cảnh ✅" onclick="selSt(this)">Đậu ✅</span>' +
'<span class="chip c-xc" data-val="Đã xuất cảnh 🎉" onclick="selSt(this)">Xuất cảnh 🎉</span>' +
'<span class="chip c-rot" data-val="Rớt ❌" onclick="selSt(this)">Rớt ❌</span>' +
'<span class="chip c-huy" data-val="Tự hủy" onclick="selSt(this)">Hủy</span>' +
'</div><input type="hidden" id="trangThai" value="Mới đăng ký"></div>' +

'<div id="bv" class="fblock"><div class="sl sl-viet">Tài chính — Đầu Việt</div><div class="r2"><div class="f"><label>Hoa hồng</label><input type="text" id="hoaHongViet" placeholder="VD: 5.000.000đ"></div><div class="f"><label>Đã thu chưa</label><select id="ttThuViet"><option value="Chưa">Chưa</option><option value="Một phần">Một phần</option><option value="Đã xong ✓">Đã xong ✓</option></select></div></div></div>' +
'<div id="bn" class="fblock"><div class="sl sl-nhat">Tài chính — Đầu Nhật</div><div class="r2"><div class="f"><label>Tiền back từ ĐT</label><input type="text" id="tienBackNhat" placeholder="VD: 500 USD"></div><div class="f"><label>Đã nhận chưa</label><select id="ttNhanNhat"><option value="Chưa">Chưa</option><option value="Một phần">Một phần</option><option value="Đã xong ✓">Đã xong ✓</option></select></div></div></div>' +

'<div class="sl sl-ctv">CTV giới thiệu</div>' +
'<label class="tr"><input type="checkbox" id="hasCTV" onchange="onCtv()"><span>Ứng viên này được CTV giới thiệu</span></label>' +
'<div id="bc" class="cblock"><div class="f" style="margin-top:8px"><label>Tên CTV</label><select id="ctvGioiThieu"><option value="">-- Chọn CTV --</option></select><div class="nh" id="nh">⚠️ Chưa có CTV. Thêm vào tab "CTV" trong sheet.</div></div><div class="r2"><div class="f"><label>Tiền trả CTV</label><input type="text" id="tienTraCtv" placeholder="VD: 500.000đ"></div><div class="f"><label>Đã trả chưa</label><select id="ttTraCtv"><option value="Chưa">Chưa</option><option value="Một phần">Một phần</option><option value="Đã xong ✓">Đã xong ✓</option></select></div></div></div>' +

'<div class="sl sl-note">Ghi chú</div>' +
'<div class="f"><textarea id="ghiChu" placeholder="Ghi chú về hồ sơ, tình trạng, lưu ý..."></textarea></div>' +
'</div>' +
'<div class="toast" id="toast"></div>' +
'<div class="ft"><button class="bc" onclick="clr()">↺</button><button class="bs" id="btn" onclick="sub()">Lưu đơn hàng</button></div>' +
'<script>' +
'var TC=["Chế biến thực phẩm","Nhóm ngành 1 (Cơ khí)","Nhóm ngành 2 (Điện tử)","Xây dựng","Điều dưỡng","Nhà hàng","Vệ sinh tòa nhà","Bảo dưỡng ô tô","Nông nghiệp","Lưu trú / Khách sạn","Ngư nghiệp","Đóng tàu","Hàng không","Vận tải","Lâm nghiệp","May","In ấn"];' +
'var KC=["Lập trình viên","Kỹ sư hệ thống / Mạng","AI & Data Science","Cơ khí","Điện - Điện tử","Hóa học & Vật liệu","Ô tô","Phiên dịch / Thông dịch","Xuất nhập khẩu","Quản trị kinh doanh / Marketing"];' +
'google.script.run.withSuccessHandler(function(l){var s=document.getElementById("ctvGioiThieu");if(!l||!l.length){document.getElementById("nh").style.display="block";return;}l.forEach(function(c){var o=document.createElement("option");o.value=c.ten;o.textContent=c.lienLac?c.ten+" ("+c.lienLac+")":c.ten;s.appendChild(o);});}).getCtvList();' +
'function onKenh(){var fb=document.getElementById("kenh").value==="Facebook";document.getElementById("lbl-ll").innerHTML=fb?"Link Facebook <span class=\'req\'>*</span>":"Tên Zalo <span class=\'req\'>*</span>";document.getElementById("lienLac").placeholder=fb?"https://facebook.com/...":"@tên hoặc số điện thoại";}' +
'function onLD(){var t=document.getElementById("loaiDon").value,s=document.getElementById("nganh");s.innerHTML="<option value=\'\'>-- Chọn ngành --</option>";(t==="kysis"?KC:(t?TC:[])).forEach(function(c){var o=document.createElement("option");o.value=o.textContent=c;s.appendChild(o);});var bv=document.getElementById("bv"),bn=document.getElementById("bn");bv.classList.toggle("show",t==="viet"||t==="kysis");bn.classList.toggle("show",t==="nhat"||t==="kysis");}' +
'function onCtv(){document.getElementById("bc").classList.toggle("show",document.getElementById("hasCTV").checked);}' +
'function selSt(el){document.querySelectorAll(".chip").forEach(function(c){c.classList.remove("active");});el.classList.add("active");document.getElementById("trangThai").value=el.getAttribute("data-val");}' +
'function g(id){return document.getElementById(id).value.trim();}' +
'function sub(){if(!g("hoTen")){toast("Vui lòng nhập họ tên","error");return;}var isFB=g("kenh")==="Facebook";if(!g("lienLac")){toast("Vui lòng nhập "+(isFB?"link Facebook":"tên Zalo"),"error");return;}if(!g("loaiDon")){toast("Vui lòng chọn loại đơn","error");return;}if(!g("nganh")){toast("Vui lòng chọn ngành nghề","error");return;}var hc=document.getElementById("hasCTV").checked,b=document.getElementById("btn");b.disabled=true;b.textContent="Đang lưu...";google.script.run.withSuccessHandler(function(){toast("✅ Đã lưu!","success");clr();b.disabled=false;b.textContent="Lưu đơn hàng";}).withFailureHandler(function(e){toast("❌ "+e.message,"error");b.disabled=false;b.textContent="Lưu đơn hàng";}).saveDonHang({hoTen:g("hoTen"),kenh:g("kenh"),lienLac:g("lienLac"),gioiTinh:g("gioiTinh"),tiengNhat:g("tiengNhat"),loaiDon:g("loaiDon"),doiTac:g("doiTac"),nganh:g("nganh"),tenCv:g("tenCv"),trangThai:g("trangThai"),hoaHongViet:g("hoaHongViet"),ttThuViet:g("ttThuViet"),tienBackNhat:g("tienBackNhat"),ttNhanNhat:g("ttNhanNhat"),ctvGioiThieu:hc?g("ctvGioiThieu"):"",tienTraCtv:hc?g("tienTraCtv"):"",ttTraCtv:hc?g("ttTraCtv"):"",ghiChu:g("ghiChu")});}' +
'function clr(){["hoTen","lienLac","doiTac","tenCv","hoaHongViet","tienBackNhat","tienTraCtv","ghiChu"].forEach(function(id){document.getElementById(id).value="";});document.getElementById("kenh").value="Zalo";document.getElementById("gioiTinh").value="Nam";document.getElementById("tiengNhat").value="Không có";document.getElementById("loaiDon").value="";document.getElementById("ttThuViet").value="Chưa";document.getElementById("ttNhanNhat").value="Chưa";document.getElementById("ttTraCtv").value="Chưa";document.getElementById("hasCTV").checked=false;document.getElementById("trangThai").value="Mới đăng ký";document.querySelectorAll(".chip").forEach(function(c){c.classList.remove("active");});document.querySelector(".c-moi").classList.add("active");document.getElementById("bc").classList.remove("show");onKenh();onLD();}' +
'function toast(m,t){var el=document.getElementById("toast");el.textContent=m;el.className="toast "+t;el.style.display="block";clearTimeout(el._t);el._t=setTimeout(function(){el.style.display="none";},3200);}' +
'<\/script></body></html>';

// ════════════════════════════════════════════════════════════
// HTML — SIDEBAR TIỀM NĂNG
// ════════════════════════════════════════════════════════════
var HTML_TIEM_NANG = '<!DOCTYPE html><html><head><base target="_top"><style>' +
'*{box-sizing:border-box;margin:0;padding:0}body{font-family:\'Google Sans\',Arial,sans-serif;font-size:13px;color:#202124;background:#f8f9fa}' +
'.hdr{background:#37474f;color:#fff;padding:13px 16px;font-size:15px;font-weight:600;position:sticky;top:0;z-index:10}' +
'.hdr small{display:block;font-size:11px;font-weight:400;opacity:.8;margin-top:2px}' +
'.fb{padding:12px 16px 80px}' +
'.sl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;margin:16px 0 8px;padding:5px 8px;border-radius:4px}' +
'.sl-uv{background:#e8f0fe;color:#1a3c5e}.sl-yc{background:#e8f5e9;color:#2e7d32}.sl-note{background:#eceff1;color:#546e7a}' +
'.f{margin-bottom:10px}' +
'label{display:block;font-size:11px;font-weight:600;color:#5f6368;text-transform:uppercase;letter-spacing:.04em;margin-bottom:4px}' +
'.req{color:#e8305a;margin-left:2px}' +
'input,select,textarea{width:100%;border:1px solid #dadce0;border-radius:6px;padding:7px 10px;font-size:13px;color:#202124;background:#fff;outline:none;font-family:inherit}' +
'input:focus,select:focus,textarea:focus{border-color:#2D2CDB;box-shadow:0 0 0 2px rgba(45,44,219,.12)}' +
'textarea{resize:vertical;min-height:80px;line-height:1.5}' +
'.r2{display:grid;grid-template-columns:1fr 1fr;gap:10px}' +
'.ft{position:fixed;bottom:0;left:0;right:0;padding:10px 16px;background:#fff;border-top:1px solid #e8eaed;display:flex;gap:8px}' +
'.bs{flex:1;background:#37474f;color:#fff;border:none;border-radius:8px;padding:10px;font-size:13px;font-weight:600;cursor:pointer}' +
'.bs:hover{background:#263238}.bs:disabled{background:#9aa0a6;cursor:not-allowed}' +
'.bc{background:transparent;color:#5f6368;border:1px solid #dadce0;border-radius:8px;padding:10px 14px;font-size:13px;cursor:pointer}' +
'.bc:hover{background:#f1f3f4}' +
'.toast{position:fixed;bottom:68px;left:16px;right:16px;padding:10px 14px;border-radius:8px;font-size:13px;font-weight:500;display:none;z-index:100;text-align:center}' +
'.toast.success{background:#e6f4ea;color:#137333;border:1px solid #b7dfbf}' +
'.toast.error{background:#fce8e6;color:#c5221f;border:1px solid #f5c6c4}' +
'</style></head><body>' +
'<div class="hdr">🔖 Ứng viên tiềm năng<small>Chưa có job phù hợp — lưu lại để tư vấn sau</small></div>' +
'<div class="fb">' +
'<div class="sl sl-uv">Thông tin cá nhân</div>' +
'<div class="f"><label>Họ và tên <span class="req">*</span></label><input type="text" id="hoTen" placeholder="VD: Nguyễn Thị Bình"></div>' +
'<div class="r2">' +
  '<div class="f"><label>Kênh liên lạc <span class="req">*</span></label><select id="kenh" onchange="onKenh()"><option value="Zalo">Zalo</option><option value="Facebook">Facebook</option></select></div>' +
  '<div class="f"><label id="lbl-ll">Tên Zalo <span class="req">*</span></label><input type="text" id="lienLac" placeholder="@tên hoặc số điện thoại"></div>' +
'</div>' +
'<div class="r2">' +
  '<div class="f"><label>Giới tính</label><select id="gioiTinh"><option>Nam</option><option>Nữ</option></select></div>' +
  '<div class="f"><label>Tiếng Nhật</label><select id="tiengNhat"><option value="Không có">Không có</option><option>N5</option><option>N4</option><option>N3</option><option>N2</option><option>N1</option></select></div>' +
'</div>' +
'<div class="sl sl-yc">Mong muốn & yêu cầu</div>' +
'<div class="r2">' +
  '<div class="f"><label>Loại đơn quan tâm</label><select id="loaiDon" onchange="onLD()"><option value="">Chưa rõ</option><option value="viet">Tokutei Đầu Việt (VN→JP)</option><option value="nhat">Tokutei Đầu Nhật (đang ở JP)</option><option value="kysis">Kỹ sư IT</option></select></div>' +
  '<div class="f"><label>Yêu cầu lương</label><input type="text" id="ycLuong" placeholder="VD: 180.000 yên+"></div>' +
'</div>' +
'<div class="f"><label>Ngành nghề quan tâm</label><select id="nganh"><option value="">-- Tất cả ngành --</option></select></div>' +
'<div class="sl sl-note">Ghi chú & kinh nghiệm</div>' +
'<div class="f"><textarea id="ghiChu" placeholder="Kinh nghiệm, chứng chỉ, nguyện vọng, lý do muốn đi Nhật..."></textarea></div>' +
'</div>' +
'<div class="toast" id="toast"></div>' +
'<div class="ft"><button class="bc" onclick="clr()">↺</button><button class="bs" id="btn" onclick="sub()">Lưu tiềm năng</button></div>' +
'<script>' +
'var TC=["Chế biến thực phẩm","Nhóm ngành 1 (Cơ khí)","Nhóm ngành 2 (Điện tử)","Xây dựng","Điều dưỡng","Nhà hàng","Vệ sinh tòa nhà","Bảo dưỡng ô tô","Nông nghiệp","Lưu trú / Khách sạn","Ngư nghiệp","Đóng tàu","Hàng không","Vận tải","Lâm nghiệp","May","In ấn"];' +
'var KC=["Lập trình viên","Kỹ sư hệ thống / Mạng","AI & Data Science","Cơ khí","Điện - Điện tử","Hóa học & Vật liệu","Ô tô","Phiên dịch / Thông dịch","Xuất nhập khẩu","Quản trị kinh doanh / Marketing"];' +
'function onKenh(){var fb=document.getElementById("kenh").value==="Facebook";document.getElementById("lbl-ll").innerHTML=fb?"Link Facebook <span class=\'req\'>*</span>":"Tên Zalo <span class=\'req\'>*</span>";document.getElementById("lienLac").placeholder=fb?"https://facebook.com/...":"@tên hoặc số điện thoại";}' +
'function onLD(){var t=document.getElementById("loaiDon").value,s=document.getElementById("nganh");s.innerHTML="<option value=\'\'>-- Tất cả ngành --</option>";(t==="kysis"?KC:(t?TC:[])).forEach(function(c){var o=document.createElement("option");o.value=o.textContent=c;s.appendChild(o);});}' +
'function g(id){return document.getElementById(id).value.trim();}' +
'function sub(){if(!g("hoTen")){toast("Vui lòng nhập họ tên","error");return;}var isFB=g("kenh")==="Facebook";if(!g("lienLac")){toast("Vui lòng nhập "+(isFB?"link Facebook":"tên Zalo"),"error");return;}var b=document.getElementById("btn");b.disabled=true;b.textContent="Đang lưu...";google.script.run.withSuccessHandler(function(){toast("✅ Đã lưu!","success");clr();b.disabled=false;b.textContent="Lưu tiềm năng";}).withFailureHandler(function(e){toast("❌ "+e.message,"error");b.disabled=false;b.textContent="Lưu tiềm năng";}).saveTiemNang({hoTen:g("hoTen"),kenh:g("kenh"),lienLac:g("lienLac"),gioiTinh:g("gioiTinh"),tiengNhat:g("tiengNhat"),loaiDon:g("loaiDon"),nganh:g("nganh"),ycLuong:g("ycLuong"),ghiChu:g("ghiChu")});}' +
'function clr(){["hoTen","lienLac","ycLuong","ghiChu"].forEach(function(id){document.getElementById(id).value="";});document.getElementById("kenh").value="Zalo";document.getElementById("gioiTinh").value="Nam";document.getElementById("tiengNhat").value="Không có";document.getElementById("loaiDon").value="";onKenh();onLD();}' +
'function toast(m,t){var el=document.getElementById("toast");el.textContent=m;el.className="toast "+t;el.style.display="block";clearTimeout(el._t);el._t=setTimeout(function(){el.style.display="none";},3200);}' +
'<\/script></body></html>';
