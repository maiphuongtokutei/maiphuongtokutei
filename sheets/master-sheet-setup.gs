// ============================================================
// TokuteiJob — Script tổng hợp (1 tab duy nhất)
// ─────────────────────────────────────────────────────────────
// Cài đặt:
//   1. Paste file này vào Apps Script (file Code.gs)
//   2. Tạo thêm file HTML tên "Sidebar", paste Sidebar.html vào
//   3. Reload Sheet → cho phép quyền khi được hỏi
//   4. Menu → ⚙️ Thiết lập sheet (chạy 1 lần)
//   5. Menu → 🌐 Tạo Sheet Công Khai (chạy 1 lần)
//      → copy link hiện ra, share cho Viewer
// ============================================================

// ── SUPABASE ─────────────────────────────────────────────────
const SUPABASE_URL = 'https://mfwxijsrrtfqmcmscocj.supabase.co';
const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1md3hpanNycnRmcW1jbXNjb2NqIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg2MDk3MzgsImV4cCI6MjA5NDE4NTczOH0.2tGumJw0n1nVidFah-xp7WC96KYlla3BmrTVGGadsuo';

// ── 1 TAB DUY NHẤT ────────────────────────────────────────────
// Cột A–O (hiển thị), P–S (nội bộ, ẩn):
//   A: type (nhat/viet/kysis)   B: title        C: category
//   D: city                     E: salary        F: japanese
//   G: gender                   H: workHours     I: daysOff
//   J: overtime                 K: raise         L: bonus
//   M: housing                  N: desc          O: status
//   P: nguon (nội bộ)           Q: hoaHong       R: phiNet    S: updatedAt
const SHEET_NAME = 'Công Việc';

// ── TỰ ĐỘNG CẬP NHẬT NGÀY KHI SỬA ──────────────────────────
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  const startRow = e.range.getRow();
  const endRow   = startRow + e.range.getNumRows() - 1;
  const col      = e.range.getColumn();

  if (endRow < 2 || col === 19) return;

  const now = new Date();
  for (let r = Math.max(startRow, 2); r <= endRow; r++) {
    sheet.getRange(r, 19).setValue(now).setNumberFormat('dd/MM/yyyy HH:mm');
  }
}

// ── MENU ─────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔄 TokuteiJob')
    .addItem('➕ Thêm job mới',                     'showSidebar')
    .addSeparator()
    .addItem('📋 Đồng bộ → Sheet Công Khai',        'syncToPublicSheet')
    .addItem('🚀 Sync tất cả lên web',               'syncToSupabase')
    .addSeparator()
    .addItem('👁️ Hiện / Ẩn cột nội bộ',            'toggleInternalColumns')
    .addSeparator()
    .addItem('⚙️ Thiết lập sheet (chạy 1 lần)',     'setupSourceSheet')
    .addItem('🌐 Tạo Sheet Công Khai (chạy 1 lần)', 'setupPublicSheet')
    .addToUi();
}

// ── QUYỀN TRUY CẬP ───────────────────────────────────────────
function isOwner() {
  const user  = Session.getActiveUser().getEmail();
  const owner = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId())
                        .getOwner().getEmail();
  return user === owner;
}

// ── SHEET CÔNG KHAI ──────────────────────────────────────────
function getPublicSheetId() {
  return PropertiesService.getScriptProperties().getProperty('PUBLIC_SHEET_ID');
}

function getPublicSpreadsheet() {
  const id = getPublicSheetId();
  if (!id) throw new Error('Chưa tạo Sheet Công Khai. Vào menu → 🌐 Tạo Sheet Công Khai.');
  return SpreadsheetApp.openById(id);
}

function setupPublicSheet() {
  const existing = getPublicSheetId();
  if (existing) {
    const btn = SpreadsheetApp.getUi().alert(
      'Sheet Công Khai đã tồn tại. Tạo lại sẽ xóa sheet cũ. Tiếp tục?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if (btn !== SpreadsheetApp.getUi().Button.YES) return;
    DriveApp.getFileById(existing).setTrashed(true);
  }

  const pub = SpreadsheetApp.create('TokuteiJob — Công Khai');
  const HEADERS = [
    'Loại đơn','Tên công việc','Ngành nghề','Tỉnh/TP','Lương','Tiếng Nhật',
    'Giới tính','Giờ làm','Ngày nghỉ','Tăng ca',
    'Tăng lương','Thưởng','Nhà ở','Mô tả','Trạng thái',
  ];
  const COL_WIDTHS = [100,160,160,120,200,110,110,130,130,100,100,110,120,300,90];

  const sheet = pub.getSheets()[0];
  sheet.setName(SHEET_NAME);
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setValues([HEADERS])
    .setBackground('#1a3c5e').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  COL_WIDTHS.forEach((w, ci) => sheet.setColumnWidth(ci + 1, w));

  PropertiesService.getScriptProperties().setProperty('PUBLIC_SHEET_ID', pub.getId());
  _doSyncToPublic(pub);

  SpreadsheetApp.getUi().alert(
    '✅ Đã tạo Sheet Công Khai!\n\n' +
    'Link: ' + pub.getUrl() + '\n\n' +
    'Share link này cho Viewer.\n' +
    '(Chỉ share quyền View — không share Edit)'
  );
}

// Đồng bộ thủ công từ menu
function syncToPublicSheet() {
  try {
    _doSyncToPublic(getPublicSpreadsheet());
    SpreadsheetApp.getUi().alert('✅ Đã đồng bộ sang Sheet Công Khai!');
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ ' + e.message);
  }
}

// Xóa sạch rồi copy cột A–O từ sheet nội bộ → công khai
function _doSyncToPublic(pubSS) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const sheet    = ss.getSheetByName(SHEET_NAME);
  const pubSheet = pubSS.getSheetByName(SHEET_NAME);
  if (!sheet || !pubSheet) return;

  const pubLast = pubSheet.getLastRow();
  if (pubLast > 1) pubSheet.getRange(2, 1, pubLast - 1, 15).clearContent();

  const last = sheet.getLastRow();
  if (last < 2) return;

  const data = sheet.getRange(2, 1, last - 1, 15).getValues()
                    .filter(r => r[1]); // r[1] = title (cột B)
  if (data.length) pubSheet.getRange(2, 1, data.length, 15).setValues(data);
}

// ── TOGGLE CỘT NỘI BỘ ────────────────────────────────────────
function toggleInternalColumns() {
  if (!isOwner()) {
    SpreadsheetApp.getUi().alert('⛔ Chỉ owner mới có quyền xem cột nội bộ.');
    return;
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const hidden = sheet.isColumnHiddenByUser(16);
  for (let c = 16; c <= 19; c++) {
    hidden ? sheet.showColumns(c) : sheet.hideColumns(c);
  }
}

// ── SIDEBAR ──────────────────────────────────────────────────
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Thêm job mới')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

function saveJob(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Không tìm thấy tab: ' + SHEET_NAME);

  const row = [
    data.type,               // A — loại đơn
    data.title,              // B
    data.category,           // C
    data.city,               // D
    data.salary,             // E
    data.japanese,           // F
    data.gender,             // G
    data.workHours,          // H
    data.daysOff,            // I
    data.overtime,           // J
    data.raise,              // K
    data.bonus,              // L
    data.housing,            // M
    data.desc,               // N
    data.status || 'active', // O
    data.nguon   || '',      // P — nội bộ
    data.hoaHong || '',      // Q — nội bộ
    data.phiNet  || '',      // R — nội bộ
    new Date(),              // S — nội bộ
  ];

  const insertAt = sheet.getLastRow() + 1;
  sheet.getRange(insertAt, 1, 1, row.length).setValues([row]);
  sheet.getRange(insertAt, 16, 1, 4).setBackground('#f3f0ff').setFontColor('#5f4b8b');
  sheet.getRange(insertAt, 19).setNumberFormat('dd/MM/yyyy HH:mm');

  // Sắp xếp theo ngành (cột C = column 3)
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet.getRange(2, 1, lastRow - 1, row.length)
      .sort([{ column: 3, ascending: true }]);
  }

  try {
    _doSyncToPublic(getPublicSpreadsheet());
  } catch (e) {
    Logger.log('Public sheet chưa setup: ' + e.message);
  }
}

// ── SUPABASE HELPERS ─────────────────────────────────────────
function getSupabaseHeaders() {
  return {
    'apikey':        SUPABASE_KEY,
    'Authorization': 'Bearer ' + SUPABASE_KEY,
    'Content-Type':  'application/json',
    'Prefer':        'return=minimal',
  };
}

function buildJob(r, i) {
  return {
    id:        String(r[0] || '').trim() + '-' + i,
    type:      String(r[0] || '').trim(),  // A
    title:     String(r[1] || '').trim(),  // B
    category:  String(r[2] || '').trim(),  // C
    city:      String(r[3] || '').trim(),  // D
    salary:    String(r[4] || '').trim(),  // E
    japanese:  String(r[5] || '').trim(),  // F
    gender:    String(r[6] || '').trim(),  // G
    workHours: String(r[7] || '').trim(),  // H
    daysOff:   String(r[8] || '').trim(),  // I
    overtime:  String(r[9] || '').trim(),  // J
    raise:     String(r[10]|| '').trim(),  // K
    bonus:     String(r[11]|| '').trim(),  // L
    housing:   String(r[12]|| '').trim(),  // M
    desc:      String(r[13]|| '').trim(),  // N
    status:    'active',
  };
}

// ── SYNC LÊN SUPABASE ────────────────────────────────────────
function syncToSupabase() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Không tìm thấy tab: "' + SHEET_NAME + '"');
    return;
  }

  const rows = sheet.getDataRange().getValues()
    .slice(1)
    .filter(r => r[1] && String(r[14] || '').toLowerCase().trim() === 'active');

  // Xóa toàn bộ jobs cũ trên Supabase
  const delRes = UrlFetchApp.fetch(
    `${SUPABASE_URL}/rest/v1/jobs?id=not.is.null`,
    { method: 'DELETE', headers: getSupabaseHeaders(), muteHttpExceptions: true }
  );
  if (delRes.getResponseCode() >= 300) {
    SpreadsheetApp.getUi().alert('❌ Lỗi xóa jobs cũ: ' + delRes.getContentText());
    return;
  }

  if (!rows.length) {
    SpreadsheetApp.getUi().alert('⚠️ Không có job active nào để sync.');
    return;
  }

  const payload = rows.map((r, i) => buildJob(r, i));
  const insRes  = UrlFetchApp.fetch(`${SUPABASE_URL}/rest/v1/jobs`, {
    method:             'POST',
    headers:            getSupabaseHeaders(),
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const code = insRes.getResponseCode();
  if (code >= 300) {
    SpreadsheetApp.getUi().alert('❌ Lỗi insert (' + code + '): ' + insRes.getContentText());
  } else {
    SpreadsheetApp.getUi().alert('✅ Sync thành công!\nĐã đẩy ' + rows.length + ' jobs lên web.');
  }
}

// ── SETUP: THIẾT LẬP SHEET (chạy 1 lần) ─────────────────────
function setupSourceSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);

  // Tạo tab nếu chưa có
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Header hàng 1
  const HEADERS = [
    'Loại đơn','Tên công việc','Ngành nghề','Tỉnh/TP','Lương','Tiếng Nhật',
    'Giới tính','Giờ làm','Ngày nghỉ','Tăng ca','Tăng lương','Thưởng','Nhà ở','Mô tả','Trạng thái',
    'Nguồn','Hoa hồng','Phí net','Ngày cập nhật',
  ];
  sheet.getRange(1, 1, 1, 15)
    .setValues([HEADERS.slice(0, 15)])
    .setBackground('#1a3c5e').setFontColor('#fff')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(1, 16, 1, 4)
    .setValues([HEADERS.slice(15)])
    .setBackground('#4a3580').setFontColor('#fff')
    .setFontWeight('bold').setHorizontalAlignment('center');

  sheet.setFrozenRows(1);

  // Độ rộng cột
  const widths = [100,160,160,120,200,110,110,130,130,100,100,110,120,300,90,150,110,110,140];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // Format ngày cập nhật (cột S)
  const lastRow = Math.max(sheet.getLastRow(), 2);
  sheet.getRange(2, 19, lastRow - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');

  // Ẩn cột nội bộ (P–S)
  for (let c = 16; c <= 19; c++) sheet.hideColumns(c);

  SpreadsheetApp.getUi().alert(
    '✅ Hoàn tất!\n\n' +
    'Tab "' + SHEET_NAME + '" đã sẵn sàng.\n\n' +
    'Tiếp theo: Menu → 🌐 Tạo Sheet Công Khai'
  );
}
