// ============================================================
// TokuteiJob — All-in-one (không cần file HTML riêng)
// ─────────────────────────────────────────────────────────────
// Cài đặt (CHỈ CẦN 1 FILE NÀY):
//   1. Paste toàn bộ file này vào Apps Script (Code.gs)
//   2. Save → Reload Sheet → cho phép quyền
//   3. Menu → ⚙️ Thiết lập sheet   (chạy 1 lần)
//   4. Menu → 🌐 Tạo Sheet Công Khai (chạy 1 lần)
// ============================================================

// ── SUPABASE ─────────────────────────────────────────────────
const SUPABASE_URL = 'https://guiyhpaeackfdsptjbol.supabase.co';
const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imd1aXlocGFlYWNrZmRzcHRqYm9sIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg5NTcwMDIsImV4cCI6MjA5NDUzMzAwMn0.jYXm2e1Wwkl8lLXOmT6JZkUxjZUKIdAPc6NdyEyFgeU';

// ── 1 TAB DUY NHẤT ───────────────────────────────────────────
// A: type   B: title     C: category  D: category_sub  E: city   F: salary
// G: japanese  H: gender  I: workHours  J: daysOff  K: overtime
// L: raise  M: bonus  N: housing  O: desc  P: status
// Q: nguon (nội bộ)  R: hoaHong  S: phiNet  T: updatedAt
const SHEET_NAME = 'Công Việc';

// ── HTML SIDEBAR (nhúng trực tiếp) ───────────────────────────
const SIDEBAR_HTML = '<!DOCTYPE html>' +
'<html><head><base target="_top"><style>' +
'* { box-sizing: border-box; margin: 0; padding: 0; }' +
'body { font-family: "Google Sans", Arial, sans-serif; font-size: 13px; color: #202124; background: #f8f9fa; }' +
'.header { background: #1a3c5e; color: white; padding: 14px 16px; font-size: 15px; font-weight: 600; position: sticky; top: 0; z-index: 10; }' +
'.form-body { padding: 12px 16px 80px; }' +
'.section-label { font-size: 10px; font-weight: 700; color: #1a3c5e; text-transform: uppercase; letter-spacing: .07em; margin: 16px 0 8px; padding-bottom: 5px; border-bottom: 2px solid #e8f0fe; }' +
'.field { margin-bottom: 10px; }' +
'label { display: block; font-size: 11px; font-weight: 600; color: #5f6368; text-transform: uppercase; letter-spacing: .04em; margin-bottom: 4px; }' +
'label .req { color: #e8305a; margin-left: 2px; }' +
'input, select, textarea { width: 100%; border: 1px solid #dadce0; border-radius: 6px; padding: 7px 10px; font-size: 13px; color: #202124; background: #fff; outline: none; transition: border-color .15s, box-shadow .15s; font-family: inherit; }' +
'input:focus, select:focus, textarea:focus { border-color: #2D2CDB; box-shadow: 0 0 0 2px rgba(45,44,219,.12); }' +
'textarea { resize: vertical; min-height: 80px; line-height: 1.5; }' +
'.row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }' +
'.footer { position: fixed; bottom: 0; left: 0; right: 0; padding: 10px 16px; background: #fff; border-top: 1px solid #e8eaed; display: flex; gap: 8px; }' +
'.btn-save { flex: 1; background: #2D2CDB; color: white; border: none; border-radius: 8px; padding: 10px; font-size: 13px; font-weight: 600; cursor: pointer; }' +
'.btn-save:hover { background: #1e1db8; } .btn-save:disabled { background: #9aa0a6; cursor: not-allowed; }' +
'.btn-clear { background: transparent; color: #5f6368; border: 1px solid #dadce0; border-radius: 8px; padding: 10px 14px; font-size: 13px; cursor: pointer; }' +
'.btn-clear:hover { background: #f1f3f4; }' +
'.toast { position: fixed; bottom: 68px; left: 16px; right: 16px; padding: 10px 14px; border-radius: 8px; font-size: 13px; font-weight: 500; display: none; z-index: 100; text-align: center; }' +
'.toast.success { background: #e6f4ea; color: #137333; border: 1px solid #b7dfbf; }' +
'.toast.error { background: #fce8e6; color: #c5221f; border: 1px solid #f5c6c4; }' +
'</style></head><body>' +
'<div class="header">+ Thêm job mới</div>' +
'<div class="form-body">' +
'<div class="section-label">Thông tin cơ bản</div>' +
'<div class="row2">' +
'<div class="field"><label>Loại việc <span class="req">*</span></label>' +
'<select id="type" onchange="onTypeChange()">' +
'<option value="">-- Chọn --</option>' +
'<option value="Việt">Tokutei Đầu Việt (VN → JP)</option>' +
'<option value="Nhật">Tokutei Đầu Nhật (đang ở JP)</option>' +
'<option value="Nhật-Việt">Nhật - Việt (cả hai)</option>' +
'<option value="Kỹ sư">Kỹ sư IT</option>' +
'</select></div>' +
'<div class="field"><label>Trạng thái</label>' +
'<select id="status"><option value="active">active</option><option value="inactive">inactive</option></select>' +
'</div></div>' +
'<div class="field"><label>Ngành nghề <span class="req">*</span></label>' +
'<select id="category"><option value="">-- Chọn loại việc trước --</option></select></div>' +
'<div class="field"><label>Tên công việc <span class="req">*</span></label>' +
'<input type="text" id="title" placeholder="VD: Công nhân chế biến thực phẩm"></div>' +
'<div class="row2">' +
'<div class="field"><label>Tỉnh / TP</label><input type="text" id="city" placeholder="VD: Aichi"></div>' +
'<div class="field"><label>Tiếng Nhật</label>' +
'<select id="japanese">' +
'<option value="Không yêu cầu">Không yêu cầu</option>' +
'<option value="N5">N5</option><option value="N4">N4</option>' +
'<option value="N3">N3</option><option value="N2">N2</option><option value="N1">N1</option>' +
'</select></div></div>' +
'<div class="field"><label>Giới tính</label>' +
'<select id="gender">' +
'<option value="Nam">Nam</option>' +
'<option value="Nữ">Nữ</option>' +
'<option value="Nam Nữ">Nam Nữ</option>' +
'</select></div>' +
'<div class="field"><label>Lương</label>' +
'<input type="text" id="salary" placeholder="VD: 180.000 - 220.000 yên/tháng"></div>' +
'<div class="section-label">Điều kiện làm việc</div>' +
'<div class="row2">' +
'<div class="field"><label>Giờ làm</label><input type="text" id="workHours" placeholder="VD: 8:00 - 17:00"></div>' +
'<div class="field"><label>Ngày nghỉ</label><input type="text" id="daysOff" placeholder="VD: Thứ 7 &amp; CN"></div>' +
'</div>' +
'<div class="row2">' +
'<div class="field"><label>Tăng ca</label><input type="text" id="overtime" placeholder="VD: 20h/tháng"></div>' +
'<div class="field"><label>Tăng lương</label><input type="text" id="raise" placeholder="VD: 6 tháng/lần"></div>' +
'</div>' +
'<div class="row2">' +
'<div class="field"><label>Thưởng</label><input type="text" id="bonus" placeholder="VD: Thưởng cuối năm"></div>' +
'<div class="field"><label>Nhà ở</label><input type="text" id="housing" placeholder="VD: Có (miễn phí)"></div>' +
'</div>' +
'<div class="section-label">Mô tả</div>' +
'<div class="field"><textarea id="desc" placeholder="Chi tiết công việc, yêu cầu, phúc lợi..."></textarea></div>' +
'<div id="mgmt-section" style="display:none">' +
'<div class="section-label">Quản lý nội bộ</div>' +
'<div class="field"><label>Nguồn</label><input type="text" id="nguon" placeholder="VD: Đối tác ABC"></div>' +
'<div class="row2">' +
'<div class="field"><label>Hoa hồng</label><input type="text" id="hoaHong" placeholder="VD: 500.000đ"></div>' +
'<div class="field"><label>Phí net</label><input type="text" id="phiNet" placeholder="VD: 300.000đ"></div>' +
'</div></div>' +
'</div>' +
'<div class="toast" id="toast"></div>' +
'<div class="footer">' +
'<button class="btn-clear" onclick="clearForm()" title="Xóa form">↺</button>' +
'<button class="btn-save" id="btnSave" onclick="doSave()">Lưu vào sheet</button>' +
'</div>' +
'<script>' +
'var TOKUTEI_CATS=["Chế biến thực phẩm","Nhóm 1","Nhóm 2","Nhóm 1+2","Đúc nóng","Đúc lạnh","Dập kim loại","Kim loại tấm","Tekko","Rèn","Gia công cơ khí","Hoàn thiện sản phẩm","Đúc nhựa","Hàn","Sơn kim loại","Lắp ráp thiết bị điện","Lắp ráp thiết bị điện tử","Sản xuất bảng mạch in (PCB)","Kiểm tra máy móc","Bảo trì máy móc","Đóng gói công nghiệp","Mạ điện","Xử lý oxy hóa nhôm (Anodize)","Đường ống","Cốp pha","Phá dỡ","Cốt thép","Sơn xây dựng","Chống thấm","Giàn giáo","Bê tông","Xây trát","San lấp","Ốp lát","Lái máy xây dựng","Hoàn thiện nội thất","Mộc xây dựng","Hoàn thiện ngoại thất","Hàn xây dựng","Điện","Thi công cách nhiệt","Lợp mái","Kim loại tấm xây dựng","Điều dưỡng","Nhà hàng","Vệ sinh tòa nhà","Bảo dưỡng ô tô","Nông nghiệp","Lưu trú / Khách sạn","Ngư nghiệp","Đóng tàu","Hàng không","Vận tải","Lâm nghiệp","May","In ấn"];' +
'var KYSIS_CATS=["Kỹ thuật & Sản xuất","Xây dựng & Công trình","IT & Công nghệ","Kinh tế & Văn phòng"];' +
'function onTypeChange(){' +
'var type=document.getElementById("type").value;' +
'var sel=document.getElementById("category");' +
'sel.innerHTML="<option value=\\"\\">-- Chọn ngành --</option>";' +
'var list=(type==="kysis"||type==="Kỹ sư")?KYSIS_CATS:(type?TOKUTEI_CATS:[]);' +
'list.forEach(function(c){var o=document.createElement("option");o.value=o.textContent=c;sel.appendChild(o);});}' +
'function doSave(){' +
'var get=function(id){return document.getElementById(id).value.trim();};' +
'if(!get("type")){showToast("Vui lòng chọn loại việc","error");return;}' +
'if(!get("category")){showToast("Vui lòng chọn ngành nghề","error");return;}' +
'if(!get("title")){showToast("Vui lòng nhập tên công việc","error");return;}' +
'var data={type:get("type"),category:get("category"),title:get("title"),city:get("city"),' +
'salary:get("salary"),japanese:get("japanese"),gender:get("gender"),' +
'workHours:get("workHours"),daysOff:get("daysOff"),overtime:get("overtime"),raise:get("raise"),' +
'bonus:get("bonus"),housing:get("housing"),desc:get("desc"),status:get("status"),' +
'nguon:get("nguon"),hoaHong:get("hoaHong"),phiNet:get("phiNet")};' +
'var btn=document.getElementById("btnSave");' +
'btn.disabled=true;btn.textContent="Đang lưu...";' +
'google.script.run' +
'.withSuccessHandler(function(){showToast("✅ Đã lưu!","success");clearForm();btn.disabled=false;btn.textContent="Lưu vào sheet";})' +
'.withFailureHandler(function(err){showToast("❌ "+err.message,"error");btn.disabled=false;btn.textContent="Lưu vào sheet";})' +
'.saveJob(data);}' +
'function clearForm(){' +
'["title","city","salary","workHours","daysOff","overtime","raise","bonus","housing","desc","nguon","hoaHong","phiNet"].forEach(function(id){document.getElementById(id).value="";});' +
'document.getElementById("type").value="";' +
'document.getElementById("status").value="active";' +
'document.getElementById("japanese").value="Không yêu cầu";' +
'document.getElementById("gender").value="Nam Nữ";' +
'onTypeChange();}' +
'function showToast(msg,type){var t=document.getElementById("toast");t.textContent=msg;t.className="toast "+type;t.style.display="block";clearTimeout(t._timer);t._timer=setTimeout(function(){t.style.display="none";},3000);}' +
'google.script.run.withSuccessHandler(function(owner){if(owner)document.getElementById("mgmt-section").style.display="block";}).isOwner();' +
'<\/script></body></html>';

// ── TỰ ĐỘNG CẬP NHẬT NGÀY KHI SỬA ──────────────────────────
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  const startRow = e.range.getRow();
  const endRow   = startRow + e.range.getNumRows() - 1;
  const col      = e.range.getColumn();

  if (endRow < 2 || col === 20) return;

  const now = new Date();
  for (let r = Math.max(startRow, 2); r <= endRow; r++) {
    sheet.getRange(r, 20).setValue(now).setNumberFormat('dd/MM/yyyy HH:mm');
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
    .addItem('↕️ Sắp xếp theo Ngành nghề',           'sortByCategory')
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
    'Loại đơn','Tên công việc','Ngành chính','Ngành phụ','Tỉnh/TP','Lương','Tiếng Nhật',
    'Giới tính','Giờ làm','Ngày nghỉ','Tăng ca',
    'Tăng lương','Thưởng','Nhà ở','Mô tả','Trạng thái',
  ];
  const COL_WIDTHS = [100,160,160,130,120,200,110,110,130,130,100,100,110,120,300,90];

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

function syncToPublicSheet() {
  try {
    _doSyncToPublic(getPublicSpreadsheet());
    SpreadsheetApp.getUi().alert('✅ Đã đồng bộ sang Sheet Công Khai!');
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ ' + e.message);
  }
}

function _doSyncToPublic(pubSS) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const sheet    = ss.getSheetByName(SHEET_NAME);
  const pubSheet = pubSS.getSheetByName(SHEET_NAME);
  if (!sheet || !pubSheet) return;

  const pubLast = pubSheet.getLastRow();
  if (pubLast > 1) pubSheet.getRange(2, 1, pubLast - 1, 16).clearContent();

  const last = sheet.getLastRow();
  if (last < 2) return;

  const data = sheet.getRange(2, 1, last - 1, 16).getValues()
                    .filter(r => r[1]);
  if (data.length) pubSheet.getRange(2, 1, data.length, 16).setValues(data);
}

// ── TOGGLE CỘT NỘI BỘ ────────────────────────────────────────
function toggleInternalColumns() {
  if (!isOwner()) {
    SpreadsheetApp.getUi().alert('⛔ Chỉ owner mới có quyền xem cột nội bộ.');
    return;
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const hidden = sheet.isColumnHiddenByUser(17);
  for (let c = 17; c <= 20; c++) {
    hidden ? sheet.showColumns(c) : sheet.hideColumns(c);
  }
}

// ── SIDEBAR ──────────────────────────────────────────────────
function showSidebar() {
  const html = HtmlService.createHtmlOutput(SIDEBAR_HTML)
    .setTitle('Thêm job mới')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

function saveJob(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Không tìm thấy tab: ' + SHEET_NAME);

  const row = [
    data.type,               // A
    data.title,              // B
    data.category,           // C
    '',                      // D — category_sub (nhập tay trong sheet)
    data.city,               // E
    data.salary,             // F
    data.japanese,           // G
    data.gender,             // H
    data.workHours,          // I
    data.daysOff,            // J
    data.overtime,           // K
    data.raise,              // L
    data.bonus,              // M
    data.housing,            // N
    data.desc,               // O
    data.status || 'active', // P
    data.nguon   || '',      // Q — nội bộ
    data.hoaHong || '',      // R — nội bộ
    data.phiNet  || '',      // S — nội bộ
    new Date(),              // T — nội bộ
  ];

  const insertAt = sheet.getLastRow() + 1;
  sheet.getRange(insertAt, 1, 1, row.length).setValues([row]);
  sheet.getRange(insertAt, 17, 1, 4).setBackground('#f3f0ff').setFontColor('#5f4b8b');
  sheet.getRange(insertAt, 20).setNumberFormat('dd/MM/yyyy HH:mm');

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

function sortByCategory() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert('Không tìm thấy tab: ' + SHEET_NAME); return; }
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) { SpreadsheetApp.getUi().alert('Chưa có dữ liệu để sắp xếp.'); return; }
  sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .sort([
      { column: 3, ascending: true },  // C: ngành nghề
      { column: 2, ascending: true },  // B: tên công việc
    ]);
  SpreadsheetApp.getActiveSpreadsheet().toast('Đã sắp xếp theo ngành nghề ✅', 'TokuteiJob', 3);
}

// ── SUPABASE ─────────────────────────────────────────────────
// Map tên hiển thị trong sheet → code gửi lên Supabase/web
function normalizeType(raw) {
  const s = String(raw || '').trim();
  const TABLE = {
    'viet': 'viet', 'nhat': 'nhat', 'both': 'both', 'kysis': 'kysis',
    'Việt': 'viet', 'Nhật': 'nhat', 'Nhật-Việt': 'both', 'Kỹ sư': 'kysis',
    'việt': 'viet', 'nhật': 'nhat', 'nhật-việt': 'both', 'kỹ sư': 'kysis',
  };
  if (TABLE[s]) return TABLE[s];
  // fallback: ghép chứa cả hai thì = both
  const lo = s.toLowerCase();
  if (lo.includes('k')) return 'kysis';
  if (lo.includes('nh') && lo.includes('vi')) return 'both';
  if (lo.includes('nh')) return 'nhat';
  if (lo.includes('vi')) return 'viet';
  return 'viet';
}

function getSupabaseHeaders() {
  return {
    'apikey':        SUPABASE_KEY,
    'Authorization': 'Bearer ' + SUPABASE_KEY,
    'Content-Type':  'application/json',
    'Prefer':        'return=minimal',
  };
}

function buildJob(r, i) {
  const type    = normalizeType(r[0]);
  const catMain = String(r[2] || '').trim();
  const catSub  = String(r[3] || '').trim();
  return {
    id:        type + '-' + i,
    type:      type,
    title:     String(r[1] || '').trim(),
    category:  [catMain, catSub].filter(Boolean).join(','),
    city:      String(r[4] || '').trim(),
    salary:    String(r[5] || '').trim(),
    japanese:  String(r[6] || '').trim(),
    gender:    String(r[7] || '').trim(),
    workHours: String(r[8] || '').trim(),
    daysOff:   String(r[9] || '').trim(),
    overtime:  String(r[10]|| '').trim(),
    raise:     String(r[11]|| '').trim(),
    bonus:     String(r[12]|| '').trim(),
    housing:   String(r[13]|| '').trim(),
    desc:      String(r[14]|| '').trim(),
    status:    'active',
    updatedAt: r[19] ? (r[19] instanceof Date ? r[19].toISOString() : new Date(r[19]).toISOString()) : null,
  };
}

function syncToSupabase() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Không tìm thấy tab: "' + SHEET_NAME + '"');
    return;
  }

  const rows = sheet.getDataRange().getValues()
    .slice(1)
    .filter(r => r[1] && String(r[15] || '').toLowerCase().trim() !== 'inactive');

  const delRes = UrlFetchApp.fetch(
    SUPABASE_URL + '/rest/v1/jobs?id=not.is.null',
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

  const payload = rows.map(function(r, i) { return buildJob(r, i); });
  const insRes  = UrlFetchApp.fetch(SUPABASE_URL + '/rest/v1/jobs', {
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

  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  const PUB_HEADERS  = ['Loại đơn','Tên công việc','Ngành chính','Ngành phụ','Tỉnh/TP','Lương','Tiếng Nhật','Giới tính','Giờ làm','Ngày nghỉ','Tăng ca','Tăng lương','Thưởng','Nhà ở','Mô tả','Trạng thái'];
  const PRIV_HEADERS = ['Nguồn','Hoa hồng','Phí net','Ngày cập nhật'];

  sheet.getRange(1, 1, 1, 16)
    .setValues([PUB_HEADERS])
    .setBackground('#1a3c5e').setFontColor('#fff')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(1, 17, 1, 4)
    .setValues([PRIV_HEADERS])
    .setBackground('#4a3580').setFontColor('#fff')
    .setFontWeight('bold').setHorizontalAlignment('center');

  sheet.setFrozenRows(1);

  [100,160,160,130,120,200,110,110,130,130,100,100,110,120,300,90,150,110,110,140]
    .forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });

  const lastRow = Math.max(sheet.getLastRow(), 2);
  sheet.getRange(2, 20, lastRow - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');

  for (let c = 17; c <= 20; c++) sheet.hideColumns(c);

  SpreadsheetApp.getUi().alert(
    '✅ Hoàn tất!\n\n' +
    'Tab "' + SHEET_NAME + '" đã sẵn sàng.\n\n' +
    'Tiếp theo: Menu → 🌐 Tạo Sheet Công Khai'
  );
}

// ── HIGHLIGHT ROW & COLUMN KHI CHỌN Ô ───────────────────────
const HL_COLOR = '#DBEAFE'; // xanh nhạt

function onSelectionChange(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  const row     = e.range.getRow();
  const col     = e.range.getColumn();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  // Xóa highlight cũ (bỏ qua hàng tiêu đề - row 1)
  sheet.getRange(2, 1, lastRow - 1, lastCol).setBackground(null);

  // Highlight toàn dòng (bỏ qua header)
  if (row >= 2) sheet.getRange(row, 1, 1, lastCol).setBackground(HL_COLOR);

  // Highlight toàn cột (bỏ qua header)
  sheet.getRange(2, col, lastRow - 1, 1).setBackground(HL_COLOR);
}
