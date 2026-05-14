// ============================================================
// TokuteiJob — Script tổng hợp
// ─────────────────────────────────────────────────────────────
// Cài đặt:
//   1. Paste file này vào Apps Script (file Code.gs)
//   2. Tạo thêm file HTML tên "Sidebar", paste Sidebar.html vào
//   3. Reload Sheet → cho phép quyền khi được hỏi
//   4. Menu → ⚙️ Thêm cột P-S vào 3 tab   (chạy 1 lần)
//   5. Menu → 🌐 Tạo Sheet Công Khai        (chạy 1 lần)
//      → copy link hiện ra, share cho Viewer
// ============================================================

// ── SUPABASE ─────────────────────────────────────────────────
const SUPABASE_URL = 'https://mfwxijsrrtfqmcmscocj.supabase.co';
const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1md3hpanNycnRmcW1jbXNjb2NqIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg2MDk3MzgsImV4cCI6MjA5NDE4NTczOH0.2tGumJw0n1nVidFah-xp7WC96KYlla3BmrTVGGadsuo';

// ── CẤU TRÚC 3 TAB NGUỒN ─────────────────────────────────────
const SOURCES = [
  { sheet: 'Tokutei Đầu Nhật', type: 'nhat'  },
  { sheet: 'Tokutei Đầu Việt', type: 'viet'  },
  { sheet: 'Kỹ sư',            type: 'kysis' },
];

// ── TỰ ĐỘNG CẬP NHẬT NGÀY KHI SỬA ──────────────────────────
function onEdit(e) {
  if (!e || !e.range) return;

  const sheetNames = SOURCES.map(s => s.sheet);
  const sheet = e.range.getSheet();
  if (!sheetNames.includes(sheet.getName())) return;

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
    .addItem('➕ Thêm job mới',                       'showSidebar')
    .addSeparator()
    .addItem('📋 Đồng bộ → Sheet Công Khai',          'syncToPublicSheet')
    .addItem('🚀 Sync tất cả lên web',                 'syncToSupabase')
    .addSeparator()
    .addItem('👁️ Hiện / Ẩn cột nội bộ',              'toggleInternalColumns')
    .addSeparator()
    .addItem('⚙️ Thêm cột P-S vào 3 tab (chạy 1 lần)', 'setupSourceSheets')
    .addItem('🌐 Tạo Sheet Công Khai (chạy 1 lần)',    'setupPublicSheet')
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

// Tạo Sheet Công Khai mới — chạy 1 lần
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
    'Tên công việc','Ngành nghề','Tỉnh/TP','Lương','Tiếng Nhật',
    'Số lượng','Giới tính','Giờ làm','Ngày nghỉ','Tăng ca',
    'Tăng lương','Thưởng','Nhà ở','Mô tả','Trạng thái',
  ];
  const COL_WIDTHS = [160,160,120,200,110,70,110,130,130,100,100,110,120,300,90];

  const defaultSheet = pub.getSheets()[0];
  SOURCES.forEach((src, i) => {
    const sheet = i === 0 ? defaultSheet : pub.insertSheet(src.sheet);
    if (i === 0) sheet.setName(src.sheet);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setValues([HEADERS])
      .setBackground('#1a3c5e').setFontColor('#ffffff')
      .setFontWeight('bold').setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    COL_WIDTHS.forEach((w, ci) => sheet.setColumnWidth(ci + 1, w));
  });

  PropertiesService.getScriptProperties().setProperty('PUBLIC_SHEET_ID', pub.getId());

  // Sync data ngay lần đầu
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

// Hàm nội bộ: xóa sạch rồi copy A-O từ sheet nội bộ → công khai
function _doSyncToPublic(pubSS) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SOURCES.forEach(src => {
    const sheet    = ss.getSheetByName(src.sheet);
    const pubSheet = pubSS.getSheetByName(src.sheet);
    if (!sheet || !pubSheet) return;

    const pubLast = pubSheet.getLastRow();
    if (pubLast > 1) pubSheet.getRange(2, 1, pubLast - 1, 15).clearContent();

    const last = sheet.getLastRow();
    if (last < 2) return;

    const data = sheet.getRange(2, 1, last - 1, 15).getValues()
                      .filter(r => r[0]);
    if (data.length) pubSheet.getRange(2, 1, data.length, 15).setValues(data);
  });
}

// ── TOGGLE CỘT NỘI BỘ ────────────────────────────────────────
function toggleInternalColumns() {
  if (!isOwner()) {
    SpreadsheetApp.getUi().alert('⛔ Chỉ owner mới có quyền xem cột nội bộ.');
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SOURCES.forEach(src => {
    const sheet = ss.getSheetByName(src.sheet);
    if (!sheet) return;
    const hidden = sheet.isColumnHiddenByUser(16);
    for (let c = 16; c <= 19; c++) {
      hidden ? sheet.showColumns(c) : sheet.hideColumns(c);
    }
  });
}

// ── SIDEBAR ──────────────────────────────────────────────────
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Thêm job mới')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

function saveJob(data) {
  const TAB_MAP = {
    nhat:  'Tokutei Đầu Nhật',
    viet:  'Tokutei Đầu Việt',
    kysis: 'Kỹ sư',
  };
  const sheetName = TAB_MAP[data.type];
  if (!sheetName) throw new Error('Loại việc không hợp lệ: ' + data.type);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Không tìm thấy tab: ' + sheetName);

  const row = [
    data.title,                 // A
    data.category,              // B
    data.city,                  // C
    data.salary,                // D
    data.japanese,              // E
    '',                         // F (không dùng)
    data.gender,                // G
    data.workHours,             // H
    data.daysOff,               // I
    data.overtime,              // J
    data.raise,                 // K
    data.bonus,                 // L
    data.housing,               // M
    data.desc,                  // N
    data.status || 'active',    // O
    data.nguon   || '',         // P — nội bộ
    data.hoaHong || '',         // Q — nội bộ
    data.phiNet  || '',         // R — nội bộ
    new Date(),                 // S — nội bộ
  ];

  const insertAt = sheet.getLastRow() + 1;
  sheet.getRange(insertAt, 1, 1, row.length).setValues([row]);
  sheet.getRange(insertAt, 16, 1, 4).setBackground('#f3f0ff').setFontColor('#5f4b8b');
  sheet.getRange(insertAt, 19).setNumberFormat('dd/MM/yyyy HH:mm');

  // Sắp xếp theo ngành (cột B)
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet.getRange(2, 1, lastRow - 1, row.length)
      .sort([{ column: 2, ascending: true }]);
  }

  // Tự đồng bộ sang Sheet Công Khai (bỏ qua nếu chưa setup)
  try {
    _doSyncToPublic(getPublicSpreadsheet());
  } catch (e) {
    Logger.log('Public sheet chưa setup: ' + e.message);
  }
}

// ── SUPABASE HELPERS ─────────────────────────────────────────
function getHeaders() {
  return {
    'apikey':        SUPABASE_KEY,
    'Authorization': 'Bearer ' + SUPABASE_KEY,
    'Content-Type':  'application/json',
    'Prefer':        'return=minimal',
  };
}

function buildJob(r, src, i) {
  return {
    id:        src.type + '-' + i,
    type:      src.type,
    title:     String(r[0]  || '').trim(),
    category:  String(r[1]  || '').trim(),
    city:      String(r[2]  || '').trim(),
    salary:    String(r[3]  || '').trim(),
    japanese:  String(r[4]  || '').trim(),
    gender:    String(r[6]  || '').trim(),
    workHours: String(r[7]  || '').trim(),
    daysOff:   String(r[8]  || '').trim(),
    overtime:  String(r[9]  || '').trim(),
    raise:     String(r[10] || '').trim(),
    bonus:     String(r[11] || '').trim(),
    housing:   String(r[12] || '').trim(),
    desc:      String(r[13] || '').trim(),
    status:    'active',
  };
}

// ── SYNC LÊN SUPABASE ────────────────────────────────────────
function syncToSupabase() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const errors = [];
  let   total  = 0;

  for (const src of SOURCES) {
    const sheet = ss.getSheetByName(src.sheet);
    if (!sheet) {
      errors.push(`Không tìm thấy tab: "${src.sheet}"`);
      continue;
    }

    const rows = sheet.getDataRange().getValues()
      .slice(1)
      .filter(r => r[0] && String(r[14] || '').toLowerCase().trim() === 'active');

    const delRes = UrlFetchApp.fetch(
      `${SUPABASE_URL}/rest/v1/jobs?type=eq.${src.type}`,
      { method: 'DELETE', headers: getHeaders(), muteHttpExceptions: true }
    );
    if (delRes.getResponseCode() >= 300) {
      errors.push(`❌ Lỗi xóa [${src.sheet}]: ${delRes.getContentText()}`);
      continue;
    }

    if (!rows.length) {
      Logger.log(`${src.sheet}: 0 job active — bỏ qua`);
      continue;
    }

    const payload = rows.map((r, i) => buildJob(r, src, i));
    const insRes  = UrlFetchApp.fetch(`${SUPABASE_URL}/rest/v1/jobs`, {
      method:             'POST',
      headers:            getHeaders(),
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    const code = insRes.getResponseCode();
    if (code >= 300) {
      errors.push(`❌ Lỗi insert [${src.sheet}] (${code}): ${insRes.getContentText()}`);
    } else {
      total += rows.length;
      Logger.log(`✅ ${src.sheet}: ${rows.length} jobs`);
    }
  }

  const msg = errors.length
    ? `⚠️ Sync hoàn tất nhưng có lỗi:\n\n${errors.join('\n')}`
    : `✅ Sync thành công!\nĐã đẩy ${total} jobs lên web.`;
  SpreadsheetApp.getUi().alert(msg);
}

// ── SETUP: THÊM CỘT P-S (chạy 1 lần) ────────────────────────
function setupSourceSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  SOURCES.forEach(src => {
    const sheet = ss.getSheetByName(src.sheet);
    if (!sheet) { Logger.log(`Không tìm thấy: ${src.sheet}`); return; }

    sheet.getRange(1, 16, 1, 4)
      .setValues([['Nguồn', 'Hoa hồng', 'Phí net', 'Ngày cập nhật']])
      .setBackground('#4a3580').setFontColor('#fff')
      .setFontWeight('bold').setHorizontalAlignment('center');

    const lastRow = Math.max(sheet.getLastRow(), 2);
    sheet.getRange(2, 19, lastRow - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');

    [150, 110, 110, 140].forEach((w, i) => sheet.setColumnWidth(16 + i, w));

    // Ẩn cột nội bộ mặc định
    for (let c = 16; c <= 19; c++) sheet.hideColumns(c);

    Logger.log(`✅ ${src.sheet}: đã thêm cột P-S`);
  });

  SpreadsheetApp.getUi().alert(
    '✅ Hoàn tất!\n\n' +
    'Đã thêm 4 cột P-S vào cuối mỗi tab (đã ẩn).\n\n' +
    'Tiếp theo: Menu → 🌐 Tạo Sheet Công Khai'
  );
}
