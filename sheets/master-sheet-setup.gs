// ============================================================
// TokuteiJob — Script quản lý Sheet JOB
// Sync đồng thời lên 2 web: maiphuongtokutei.com + nippontokutei.com
// Kiến trúc giống Sheet_JOB_AppScript_FULL.js (nippontokutei)
// ============================================================

// ── SUPABASE: maiphuongtokutei.com (guiyh) ───────────────────
// Script này chỉ dùng cho sheet phụ/backup của Web-MP.
// Sync nippontokutei.com được xử lý bởi Sheet_JOB_AppScript_FULL.js
const SB_MP_URL = 'https://guiyhpaeackfdsptjbol.supabase.co';
const SB_MP_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imd1aXlocGFlYWNrZmRzcHRqYm9sIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg5NTcwMDIsImV4cCI6MjA5NDUzMzAwMn0.jYXm2e1Wwkl8lLXOmT6JZkUxjZUKIdAPc6NdyEyFgeU';

// ── TAB NAMES ─────────────────────────────────────────────────
const SHEET_NAME  = 'Công Việc';
const INACTIVE_TAB = 'Inactive';

// ── HEADERS cho sheet phụ (A–P = 16 cột công khai + Q = Link ảnh) ──
const PUB_HEADERS = [
  'Loại đơn','Tên công việc','Ngành chính','Ngành phụ','Tỉnh/TP','Lương','Tiếng Nhật',
  'Giới tính','Giờ làm','Ngày nghỉ','Tăng ca',
  'Tăng lương','Thưởng','Nhà ở','Mô tả','Trạng thái','Link ảnh',
];
const PUB_COL_WIDTHS = [100,160,160,130,120,200,110,110,130,130,100,100,110,120,300,90,240];
const PUB_NCOLS = PUB_HEADERS.length;  // 17

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
'.np{position:relative}' +
'.np-trg{width:100%;border:1px solid #dadce0;border-radius:6px;padding:7px 10px;font-size:13px;background:#fff;text-align:left;cursor:pointer;display:flex;justify-content:space-between;align-items:center;font-family:inherit;color:#202124}' +
'.np-trg:hover{border-color:#2D2CDB}' +
'.np-trg.ph{color:#9aa0a6}' +
'.np-trg .cv{color:#5f6368;font-size:10px;transition:transform .15s}' +
'.np-trg.op .cv{transform:rotate(180deg)}' +
'.np-pn{position:absolute;top:calc(100% + 4px);left:0;right:0;background:#fff;border:1px solid #dadce0;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.1);max-height:320px;overflow-y:auto;z-index:50;padding:4px}' +
'.np-gh,.np-fl,.np-it{width:100%;border:none;background:transparent;cursor:pointer;font-family:inherit;text-align:left;border-radius:5px}' +
'.np-gh{display:flex;justify-content:space-between;align-items:center;padding:8px 10px;font-size:13px;font-weight:600;color:#1a3c5e}' +
'.np-gh:hover{background:#f1f3f4}' +
'.np-gh .cv{font-size:10px;color:#5f6368;transition:transform .15s}' +
'.np-gh.op .cv{transform:rotate(180deg)}' +
'.np-gi{display:none;padding:2px 0 4px 8px;border-left:2px solid #e8eaed;margin-left:14px}' +
'.np-gi.op{display:block}' +
'.np-fl{padding:8px 10px;font-size:13px;color:#202124;display:block}' +
'.np-it{padding:6px 10px;font-size:12px;color:#3c4043;display:block}' +
'.np-fl:hover,.np-it:hover{background:#e8f0fe;color:#1a3c5e}' +
'.np-fl.sel,.np-it.sel{background:#2D2CDB;color:#fff}' +
'.big-group-hd{font-size:12px;font-weight:700;color:#5f6368;text-transform:uppercase;padding:12px 10px 4px;border-bottom:1px solid #e8eaed;margin-bottom:4px;letter-spacing:0.05em}' +
'.cv svg{width:12px;height:12px;stroke-width:2.5px;stroke:currentColor;fill:none;display:block}' +
'</style></head><body>' +
'<div class="header">+ Thêm job mới</div>' +
'<div class="form-body">' +
'<div class="section-label">Thông tin cơ bản</div>' +
'<div class="row2">' +
'<div class="field"><label>Loại việc <span class="req">*</span></label>' +
'<select id="type" onchange="onTypeChange()">' +
'<option value="">-- Chọn --</option>' +
'<option value="Nhật">Tokutei Đầu Nhật (đang ở JP)</option>' +
'<option value="Việt">Tokutei Đầu Việt (VN → JP)</option>' +
'<option value="Nhật-Việt">Nhật - Việt (cả hai)</option>' +
'<option value="Indonesia">Indonesia (đầu Nhật)</option>' +
'<option value="Kỹ sư đầu Nhật">Kỹ sư đầu Nhật</option>' +
'<option value="Kỹ sư đầu Việt">Kỹ sư đầu Việt</option>' +
'</select></div>' +
'<div class="field"><label>Trạng thái</label>' +
'<select id="status"><option value="active">active</option><option value="inactive">inactive</option></select>' +
'</div></div>' +
'<div class="field"><label>Ngành nghề <span class="req">*</span></label>' +
'<div class="np" id="npPicker">' +
'<button type="button" class="np-trg ph" id="npTrg" onclick="tNP(event)"><span id="npTxt">-- Chọn loại việc trước --</span><span class="cv">▾</span></button>' +
'<div class="np-pn" id="npPn" hidden></div>' +
'</div><input type="hidden" id="category"></div>' +
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
'<div class="field"><label>Link ảnh (Google Drive)</label>' +
'<input type="text" id="imageUrl" placeholder="VD: https://drive.google.com/file/d/..."></div>' +
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
'var TG=[' +
  '{g:"Chế biến thực phẩm",items:["Chế biến thực phẩm"]},' +
  '{g:"① Gia công cơ khí - kim loại",items:["Đúc nóng","Đúc lạnh","Dập kim loại","Kim loại tấm","Tekko","Rèn","Gia công cơ khí","Hoàn thiện sản phẩm","Đúc nhựa","Hàn","Sơn kim loại","Lắp ráp thiết bị điện","Kiểm tra máy móc","Bảo trì máy móc","Đóng gói công nghiệp"]},' +
  '{g:"② Lắp ráp điện - điện tử",items:["Gia công cơ khí","Hoàn thiện sản phẩm","Đúc nhựa","Lắp ráp thiết bị điện","Lắp ráp thiết bị điện tử","Sản xuất bảng mạch in (PCB)","Kiểm tra máy móc","Bảo trì máy móc","Đóng gói công nghiệp"]},' +
  '{g:"③ Xử lý bề mặt kim loại",items:["Mạ điện","Xử lý oxy hóa nhôm (Anodize)"]},' +
  '{g:"Xây dựng",items:["Nhóm ngành xây dựng","Đường ống","Cốp pha","Phá dỡ","Cốt thép","Sơn xây dựng","Chống thấm","Giàn giáo","Bê tông","Xây trát","San lấp","Ốp lát","Lái máy xây dựng","Hoàn thiện nội thất","Mộc xây dựng","Hoàn thiện ngoại thất","Hàn xây dựng","Điện","Thi công cách nhiệt","Lợp mái","Kim loại tấm xây dựng"]},' +
  '{g:"Điều dưỡng",items:["Điều dưỡng"]},' +
  '{g:"Nhà hàng",items:["Nhà hàng"]},' +
  '{g:"Nông nghiệp",items:["Nông nghiệp"]},' +
  '{g:"Vệ sinh tòa nhà",items:["Vệ sinh tòa nhà"]},' +
  '{g:"Bảo dưỡng ô tô",items:["Bảo dưỡng ô tô"]},' +
  '{g:"Lưu trú / Khách sạn",items:["Lưu trú / Khách sạn"]},' +
  '{g:"Ngư nghiệp",items:["Ngư nghiệp"]},' +
  '{g:"Đóng tàu",items:["Đóng tàu"]},' +
  '{g:"Hàng không",items:["Hàng không"]},' +
  '{g:"Vận tải",items:["Vận tải"]},' +
  '{g:"Lâm nghiệp",items:["Lâm nghiệp"]},' +
  '{g:"May",items:["May"]},' +
  '{g:"In ấn",items:["In ấn"]}' +
'];' +
'var KC=["Kỹ thuật & Sản xuất","Xây dựng & Công trình","IT & Công nghệ","Kinh tế & Văn phòng"];' +
'function escA(s){return String(s).replace(/"/g,"&quot;");}' +
'function buildNP(t){' +
  'var p=document.getElementById("npPn");' +
  'if(!t){p.innerHTML="";return;}' +
  'if(t.indexOf("Kỹ")!==-1){' +
    'p.innerHTML=KC.map(function(v){return "<button type=\\"button\\" class=\\"np-fl\\" data-v=\\""+escA(v)+"\\" onclick=\\"pickN(this)\\">"+v+"</button>";}).join("");' +
  '}else{' +
    'p.innerHTML=TG.map(function(o){' +
      'if(o.items.length===1)return "<button type=\\"button\\" class=\\"np-fl\\" data-v=\\""+escA(o.items[0])+"\\" onclick=\\"pickN(this)\\">"+o.items[0]+"</button>";' +
      'return "<button type=\\"button\\" class=\\"np-gh\\" onclick=\\"tNG(this)\\"><span>"+o.g+"</span><span class=\\"cv\\">▾</span></button><div class=\\"np-gi\\">"+o.items.map(function(v){return "<button type=\\"button\\" class=\\"np-it\\" data-v=\\""+escA(v)+"\\" onclick=\\"pickN(this)\\">"+v+"</button>";}).join("")+"</div>";' +
    '}).join("");' +
  '}' +
'}' +
'function tNP(e){if(e)e.stopPropagation();var tr=document.getElementById("npTrg"),pn=document.getElementById("npPn");if(!document.getElementById("type").value){alert("Vui lòng chọn Loại việc trước");return;}var op=pn.hidden;pn.hidden=!op;tr.classList.toggle("op",op);}' +
'function tNG(el){el.classList.toggle("op");el.nextElementSibling.classList.toggle("op");}' +
'function pickN(el){var v=el.getAttribute("data-v");document.getElementById("category").value=v;document.getElementById("npTxt").textContent=v;var tr=document.getElementById("npTrg");tr.classList.remove("op","ph");document.getElementById("npPn").hidden=true;document.querySelectorAll("#npPn .np-fl,#npPn .np-it").forEach(function(b){b.classList.remove("sel");});el.classList.add("sel");}' +
'document.addEventListener("click",function(e){var p=document.getElementById("npPicker");if(p&&!p.contains(e.target)){document.getElementById("npPn").hidden=true;document.getElementById("npTrg").classList.remove("op");}});' +
'function onTypeChange(){' +
'var type=document.getElementById("type").value;' +
'document.getElementById("category").value="";' +
'document.getElementById("npTxt").textContent=type?"-- Chọn ngành --":"-- Chọn loại việc trước --";' +
'var tr=document.getElementById("npTrg");tr.classList.add("ph");tr.classList.remove("op");' +
'document.getElementById("npPn").hidden=true;' +
'buildNP(type);}' +
'function doSave(){' +
'var get=function(id){return document.getElementById(id).value.trim();};' +
'if(!get("type")){showToast("Vui lòng chọn loại việc","error");return;}' +
'if(!get("category")){showToast("Vui lòng chọn ngành nghề","error");return;}' +
'if(!get("title")){showToast("Vui lòng nhập tên công việc","error");return;}' +
'var data={type:get("type"),category:get("category"),title:get("title"),city:get("city"),' +
'salary:get("salary"),japanese:get("japanese"),gender:get("gender"),' +
'workHours:get("workHours"),daysOff:get("daysOff"),overtime:get("overtime"),raise:get("raise"),' +
'bonus:get("bonus"),housing:get("housing"),desc:get("desc"),status:get("status"),' +
'imageUrl:get("imageUrl"),' +
'nguon:get("nguon"),hoaHong:get("hoaHong"),phiNet:get("phiNet")};' +
'var btn=document.getElementById("btnSave");' +
'btn.disabled=true;btn.textContent="Đang lưu...";' +
'google.script.run' +
'.withSuccessHandler(function(){showToast("✅ Đã lưu và sync lên web!","success");clearForm();btn.disabled=false;btn.textContent="Lưu vào sheet";})' +
'.withFailureHandler(function(err){showToast("❌ "+err.message,"error");btn.disabled=false;btn.textContent="Lưu vào sheet";})' +
'.saveJob(data);}' +
'function clearForm(){' +
'["title","city","salary","workHours","daysOff","overtime","raise","bonus","housing","desc","imageUrl","nguon","hoaHong","phiNet"].forEach(function(id){document.getElementById(id).value="";});' +
'document.getElementById("type").value="";' +
'document.getElementById("status").value="active";' +
'document.getElementById("japanese").value="Không yêu cầu";' +
'document.getElementById("gender").value="Nam Nữ";' +
'onTypeChange();}' +
'function showToast(msg,type){var t=document.getElementById("toast");t.textContent=msg;t.className="toast "+type;t.style.display="block";clearTimeout(t._timer);t._timer=setTimeout(function(){t.style.display="none";},3000);}' +
'google.script.run.withSuccessHandler(function(owner){if(owner)document.getElementById("mgmt-section").style.display="block";}).isOwner();' +
'<\/script></body></html>';

// ── onEdit: cập nhật ngày + move giữa Công Việc ↔ Inactive ──
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const name  = sheet.getName();
  if (name !== SHEET_NAME && name !== INACTIVE_TAB) return;

  const startRow = e.range.getRow();
  const endRow   = startRow + e.range.getNumRows() - 1;
  const col      = e.range.getColumn();

  if (endRow < 2 || col === 20) return;

  // Cập nhật ngày sửa (cột T = 20)
  const now = new Date();
  for (let r = Math.max(startRow, 2); r <= endRow; r++) {
    sheet.getRange(r, 20).setValue(now);
  }

  // Nếu sửa cột P (status = 16) → tự move dòng theo trạng thái
  if (col === 16) {
    let moved = 0;
    for (let r = endRow; r >= Math.max(startRow, 2); r--) {
      const val = String(sheet.getRange(r, 16).getValue()).trim().toLowerCase();
      if (name === SHEET_NAME && val === 'inactive') {
        _moveRowToInactive(sheet, r); moved++;
      } else if (name === INACTIVE_TAB && val === 'active') {
        _moveRowToActive(sheet, r); moved++;
      }
    }
    if (moved) {
      const target = name === SHEET_NAME ? 'Inactive' : SHEET_NAME;
      SpreadsheetApp.getActiveSpreadsheet()
        .toast('✅ Đã chuyển ' + moved + ' job sang ' + target, 'TokuteiJob', 3);
    }
  }
}

// ── MENU ─────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 TokuteiJob')
    .addItem('➕ Thêm job mới',                        'showSidebar')
    .addSeparator()
    .addItem('🌐 Đẩy job lên 2 web (maiphuong + nippon)', 'syncEverything')
    .addItem('⏰ Cài tự đẩy web (15 phút)',             'setupAutoSync')
    .addSeparator()
    .addItem('📋 Đồng bộ → Sheet Công Khai (Nhật)',    'syncToPublicSheet')
    .addItem('📂 Đồng bộ → Sheet Đầu Việt',            'syncDauVietSheet')
    .addSeparator()
    .addItem('👁️ Hiện / Ẩn cột nội bộ',               'toggleInternalColumns')
    .addItem('⚙️ Sắp xếp theo Ngành nghề',              'sortByCategory')
    .addItem('🗂️ Move tất cả job Inactive',             'moveAllInactive')
    .addSeparator()
    .addItem('🔽 Cập nhật dropdown "Loại đơn"',         'setupLoaiDonDropdown')
    .addItem('⚙️ Thiết lập sheet (chạy 1 lần)',         'setupSourceSheet')
    .addItem('🌐 Tạo Sheet Công Khai (chạy 1 lần)',     'setupPublicSheet')
    .addItem('🌍 Tạo Sheet Đầu Việt (chạy 1 lần)',      'setupDauVietSheet')
    .addToUi();
}

// ── QUYỀN TRUY CẬP ───────────────────────────────────────────
function isOwner() {
  const user  = Session.getActiveUser().getEmail();
  const owner = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId())
                        .getOwner().getEmail();
  return user === owner;
}

// ══════════════════════════════════════════════════════════════
//  PHÂN LOẠI loaiDon → nhánh
//  nhat | viet | both | kysis_nhat | kysis_viet
// ══════════════════════════════════════════════════════════════
function normalizeType(raw) {
  const s = String(raw || '').trim().toLowerCase();
  if (!s) return 'nhat';
  const isK    = s.indexOf('kỹ sư') !== -1 || s.indexOf('ky su') !== -1 || s.indexOf('kysis') !== -1;
  const dauViet = s.indexOf('đầu việt') !== -1 || s.indexOf('dau viet') !== -1;
  if (isK) return dauViet ? 'kysis_viet' : 'kysis_nhat';
  if (s === 'viet' || s === 'việt') return 'viet';
  if (s.indexOf('nhật-việt') !== -1 || s.indexOf('nhat-viet') !== -1) return 'both';
  if (s.indexOf('indonesia') !== -1) return 'nhat';
  if (s.indexOf('nhật') !== -1 || s.indexOf('nhat') !== -1) return 'nhat';
  if (s.indexOf('việt') !== -1 || s.indexOf('viet') !== -1) return 'viet';
  return 'nhat';
}

// Job lên web ĐẦU NHẬT? → Nhật / Nhật-Việt / Indonesia / Kỹ sư đầu Nhật
function _isForNhat(t) { return t === 'nhat' || t === 'both' || t === 'kysis_nhat'; }
// Job lên web ĐẦU VIỆT (sheet phụ)? → Việt / Nhật-Việt / Kỹ sư đầu Việt
function _isForViet(t) { return t === 'viet' || t === 'both' || t === 'kysis_viet'; }

// ── SYNC LÊN CẢ 2 WEB + SHEET PHỤ (menu) ────────────────────
function syncEverything() {
  const log = [];
  try {
    const n = _pushSupabase();
    log.push('✅ Supabase maiphuong + nippon: ' + n + ' job');
  } catch (e) {
    log.push('❌ Supabase: ' + e.message);
  }
  try {
    const n = _syncToPublic(getPublicSheetId(), _isForNhat, '#1a3c5e');
    log.push('✅ Sheet Công Khai (Nhật): ' + n + ' job');
  } catch (e) {
    log.push('⚠️ Sheet Nhật: ' + e.message);
  }
  try {
    const n = _syncToPublic(getDauVietSheetId(), _isForViet, '#B5426A');
    log.push('✅ Sheet Đầu Việt: ' + n + ' job');
  } catch (e) {
    log.push('⚠️ Sheet Việt: ' + e.message);
  }
  SpreadsheetApp.getUi().alert('Kết quả đẩy job lên web:\n\n' + log.join('\n'));
}

// Bản chạy ngầm cho trigger / saveJob (không hiện popup)
function autoSyncAll() {
  try { _pushSupabase(); } catch (e) { Logger.log('Supabase: ' + e.message); }
  try {
    const pubId = getPublicSheetId();
    if (pubId) _syncToPublic(pubId, _isForNhat, '#1a3c5e');
  } catch (e) { Logger.log('Sheet Nhật: ' + e.message); }
  try {
    const vietId = getDauVietSheetId();
    if (vietId) _syncToPublic(vietId, _isForViet, '#B5426A');
  } catch (e) { Logger.log('Sheet Việt: ' + e.message); }
}

// Cài trigger tự đẩy web mỗi 15 phút
function setupAutoSync() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'autoSyncAll') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('autoSyncAll').timeBased().everyMinutes(15).create();
  SpreadsheetApp.getActiveSpreadsheet()
    .toast('⏰ Đã cài tự đẩy web mỗi 15 phút.', 'TokuteiJob', 4);
}

// ── HÀM SYNC RA SHEET PHỤ ────────────────────────────────────
function _syncToPublic(sheetId, filterFn, headerColor) {
  if (!sheetId) throw new Error('Chưa có ID sheet phụ.');
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Không tìm thấy tab "' + SHEET_NAME + '"');

  const pubSS = SpreadsheetApp.openById(sheetId);
  let pubSheet = pubSS.getSheetByName(SHEET_NAME);
  if (!pubSheet) {
    pubSheet = pubSS.getSheets()[0];
    pubSheet.setName(SHEET_NAME);
    _writePublicHeader(pubSheet, headerColor || '#1a3c5e');
  }

  const pubLast = pubSheet.getLastRow();
  if (pubLast > 1) pubSheet.getRange(2, 1, pubLast - 1, PUB_NCOLS).clearContent();

  const last = sheet.getLastRow();
  if (last < 2) return 0;

  const src = sheet.getRange(2, 1, last - 1, 22).getValues();
  const data = src
    .filter(function(r) {
      if (!r[1]) return false;
      if (String(r[15] || '').toLowerCase().trim() === 'inactive') return false;
      return filterFn(normalizeType(r[0]));
    })
    .map(function(r) {
      return r.slice(0, 16).concat([r[21] || '']);  // A–P + cột V (imageUrl) → Q sheet phụ
    });

  if (data.length) pubSheet.getRange(2, 1, data.length, PUB_NCOLS).setValues(data);
  return data.length;
}

function _writePublicHeader(pubSheet, color) {
  pubSheet.getRange(1, 1, 1, PUB_HEADERS.length)
    .setValues([PUB_HEADERS])
    .setBackground(color || '#1a3c5e').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');
  pubSheet.setFrozenRows(1);
  PUB_COL_WIDTHS.forEach(function(w, ci) { pubSheet.setColumnWidth(ci + 1, w); });
}

// ── WRAPPERS CÓ UI (backward compatible) ─────────────────────
function syncToPublicSheet() {
  try {
    const n = _syncToPublic(getPublicSheetId(), _isForNhat, '#1a3c5e');
    SpreadsheetApp.getUi().alert('✅ Đã đồng bộ ' + n + ' job sang Sheet Công Khai!');
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ ' + e.message);
  }
}

function syncDauVietSheet() {
  try {
    const n = _syncToPublic(getDauVietSheetId(), _isForViet, '#B5426A');
    SpreadsheetApp.getUi().alert('✅ Đã đồng bộ ' + n + ' job sang Sheet Đầu Việt!');
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ ' + e.message);
  }
}

// ── SHEET CÔNG KHAI (Đầu Nhật) ───────────────────────────────
function getPublicSheetId() {
  return PropertiesService.getScriptProperties().getProperty('PUBLIC_SHEET_ID');
}

function setupPublicSheet() {
  const existing = getPublicSheetId();
  if (existing) {
    const btn = SpreadsheetApp.getUi().alert(
      'Sheet Công Khai đã tồn tại. Tạo lại sẽ xóa sheet cũ. Tiếp tục?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if (btn !== SpreadsheetApp.getUi().Button.YES) return;
    try { DriveApp.getFileById(existing).setTrashed(true); } catch(e) {}
  }

  const pub = SpreadsheetApp.create('TokuteiJob — Công Khai (Đầu Nhật)');
  _writePublicHeader(pub.getSheets()[0].setName(SHEET_NAME), '#1a3c5e');
  pub.getSheets()[0].setFrozenRows(1);

  PropertiesService.getScriptProperties().setProperty('PUBLIC_SHEET_ID', pub.getId());
  _syncToPublic(pub.getId(), _isForNhat, '#1a3c5e');

  SpreadsheetApp.getUi().alert(
    '✅ Đã tạo Sheet Công Khai!\n\nLink: ' + pub.getUrl() +
    '\n\nShare quyền Viewer cho người cần xem.'
  );
}

// ── SHEET ĐẦU VIỆT ───────────────────────────────────────────
function getDauVietSheetId() {
  return PropertiesService.getScriptProperties().getProperty('DAU_VIET_SHEET_ID');
}

function setupDauVietSheet() {
  const existing = getDauVietSheetId();
  if (existing) {
    const btn = SpreadsheetApp.getUi().alert(
      'Sheet Đầu Việt đã tồn tại. Tạo lại sẽ xóa sheet cũ. Tiếp tục?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if (btn !== SpreadsheetApp.getUi().Button.YES) return;
    try { DriveApp.getFileById(existing).setTrashed(true); } catch(e) {}
  }

  const pub = SpreadsheetApp.create('TokuteiJob — Đầu Việt');
  _writePublicHeader(pub.getSheets()[0].setName(SHEET_NAME), '#B5426A');
  pub.getSheets()[0].setFrozenRows(1);

  PropertiesService.getScriptProperties().setProperty('DAU_VIET_SHEET_ID', pub.getId());
  _syncToPublic(pub.getId(), _isForViet, '#B5426A');

  SpreadsheetApp.getUi().alert(
    '✅ Đã tạo Sheet Đầu Việt!\n\nLink: ' + pub.getUrl() +
    '\n\nShare quyền Viewer cho cộng tác viên Đầu Việt.'
  );
}

// ── TOGGLE CỘT NỘI BỘ ────────────────────────────────────────
function toggleInternalColumns() {
  if (!isOwner()) {
    SpreadsheetApp.getUi().alert('❌ Chỉ owner mới có quyền xem cột nội bộ.');
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
    data.type,               // A  — loại đơn
    data.title,              // B  — tên công việc
    data.category,           // C  — ngành chính
    '',                      // D  — ngành phụ (nhập tay trong sheet)
    data.city,               // E  — tỉnh/TP
    data.salary,             // F  — lương
    data.japanese,           // G  — tiếng Nhật
    data.gender,             // H  — giới tính
    data.workHours,          // I  — giờ làm
    data.daysOff,            // J  — ngày nghỉ
    data.overtime,           // K  — tăng ca
    data.raise,              // L  — tăng lương
    data.bonus,              // M  — thưởng
    data.housing,            // N  — nhà ở
    data.desc,               // O  — mô tả
    data.status || 'active', // P  — trạng thái
    data.nguon   || '',      // Q  — nguồn (nội bộ)
    data.hoaHong || '',      // R  — hoa hồng (nội bộ)
    data.phiNet  || '',      // S  — phí net (nội bộ)
    new Date(),              // T  — ngày cập nhật (nội bộ)
    '',                      // U  — (dự phòng)
    data.imageUrl || '',     // V  — link ảnh
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

  autoSyncAll();
}

function sortByCategory() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert('Không tìm thấy tab: ' + SHEET_NAME); return; }
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) { SpreadsheetApp.getUi().alert('Chưa có dữ liệu để sắp xếp.'); return; }
  sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .sort([
      { column: 3, ascending: true },
      { column: 2, ascending: true },
    ]);
  SpreadsheetApp.getActiveSpreadsheet().toast('Đã sắp xếp theo ngành nghề ✅', 'TokuteiJob', 3);
}

// ── MOVE INACTIVE ↔ ACTIVE ───────────────────────────────────
function _moveRowToInactive(sheet, row) {
  const ss = sheet.getParent();
  let dest = ss.getSheetByName(INACTIVE_TAB);
  const lastCol = sheet.getLastColumn();
  if (!dest) {
    dest = ss.insertSheet(INACTIVE_TAB);
    sheet.getRange(1, 1, 1, lastCol).copyTo(dest.getRange(1, 1, 1, lastCol));
  }
  const destRow = dest.getLastRow() + 1;
  sheet.getRange(row, 1, 1, lastCol).copyTo(dest.getRange(destRow, 1, 1, lastCol));
  sheet.deleteRow(row);
}

function _moveRowToActive(sheet, row) {
  const ss = sheet.getParent();
  const dest = ss.getSheetByName(SHEET_NAME);
  if (!dest) return;
  const lastCol = sheet.getLastColumn();
  const destRow = dest.getLastRow() + 1;
  sheet.getRange(row, 1, 1, lastCol).copyTo(dest.getRange(destRow, 1, 1, lastCol));
  sheet.deleteRow(row);
}

function moveAllInactive() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    ss.toast('Không có dữ liệu', 'TokuteiJob', 3);
    return;
  }
  const data = sheet.getRange(2, 16, sheet.getLastRow() - 1, 1).getValues();
  const rows = [];
  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]).trim().toLowerCase() === 'inactive') rows.push(i + 2);
  }
  rows.forEach(r => _moveRowToInactive(sheet, r));
  ss.toast('✅ Đã chuyển ' + rows.length + ' job sang Inactive', 'TokuteiJob', 3);
}

// ══════════════════════════════════════════════════════════════
//  SUPABASE — đẩy lên đồng thời 2 web
//  maiphuongtokutei.com (SB_MP) + nippontokutei.com (SB_NT)
// ══════════════════════════════════════════════════════════════
function _sbHeaders(key) {
  return {
    'apikey':        key,
    'Authorization': 'Bearer ' + key,
    'Content-Type':  'application/json',
    'Prefer':        'return=minimal',
  };
}

// Chuyển ngày (cột T) sang ISO an toàn
function _safeISO(v) {
  try {
    if (v instanceof Date) return isNaN(v.getTime()) ? new Date().toISOString() : v.toISOString();
    const s = String(v || '').trim();
    if (!s) return new Date().toISOString();
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[ T](\d{1,2}):(\d{2}))?/);
    const d = m ? new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0)) : new Date(s);
    return isNaN(d.getTime()) ? new Date().toISOString() : d.toISOString();
  } catch (e) {
    return new Date().toISOString();
  }
}

// 1 dòng → object Supabase
// type: 'nhat' | 'kysis' | 'both' | 'viet' | 'kysis_viet'
// loai_don: giữ nguyên giá trị gốc từ sheet để web lọc đúng
function buildJob(r, i) {
  const t = normalizeType(r[0]);
  const webType =
    t === 'kysis_nhat' ? 'kysis'
    : t === 'kysis_viet' ? 'kysis_viet'
    : t === 'viet'      ? 'viet'
    : t === 'both'      ? 'both'
    : 'nhat';
  const catMain = String(r[2] || '').trim();
  const catSub  = String(r[3] || '').trim();
  return {
    id:        'gs_' + i,
    loai_don:  String(r[0] || '').trim(),   // ← bắt buộc để TokuteiViet lọc đúng
    type:      webType,
    title:     String(r[1] || '').trim(),
    category:  [catMain, catSub].filter(Boolean).join(','),
    city:      String(r[4]  || '').trim(),
    salary:    String(r[5]  || '').trim(),
    japanese:  String(r[6]  || '').trim(),
    gender:    String(r[7]  || '').trim(),
    workHours: String(r[8]  || '').trim(),
    daysOff:   String(r[9]  || '').trim(),
    overtime:  String(r[10] || '').trim(),
    raise:     String(r[11] || '').trim(),
    bonus:     String(r[12] || '').trim(),
    housing:   String(r[13] || '').trim(),
    desc:      String(r[14] || '').trim(),
    imageUrl:  String(r[21] || '').trim(),   // cột V
    status:    'active',
    updatedAt: _safeISO(r[19]),
  };
}

// Push vào 1 Supabase (xóa job sync cũ → insert mới)
// Giữ lại job admin thêm tay trên web (id bắt đầu 'job_')
function _pushToOne(sbUrl, sbKey, payload) {
  const headers = _sbHeaders(sbKey);
  const delRes = UrlFetchApp.fetch(
    sbUrl + '/rest/v1/jobs?id=not.like.job_*',
    { method: 'DELETE', headers: headers, muteHttpExceptions: true }
  );
  if (delRes.getResponseCode() >= 300) throw new Error('DELETE ' + sbUrl + ': ' + delRes.getContentText());
  // Dọn job thêm tay có type không còn dùng trên web (viet/both cũ)
  UrlFetchApp.fetch(
    sbUrl + '/rest/v1/jobs?type=not.in.' + encodeURIComponent('("nhat","kysis")'),
    { method: 'DELETE', headers: headers, muteHttpExceptions: true }
  );

  if (!payload.length) return;
  const insRes = UrlFetchApp.fetch(sbUrl + '/rest/v1/jobs', {
    method:             'POST',
    headers:            headers,
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  if (insRes.getResponseCode() >= 300) throw new Error('INSERT ' + sbUrl + ': ' + insRes.getContentText());
}

// Lõi: đọc sheet → build payload → push lên CẢ 2 Supabase
function _pushSupabase() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Không tìm thấy tab: "' + SHEET_NAME + '"');

  const last = sheet.getLastRow();
  const rows = (last < 2 ? [] : sheet.getRange(2, 1, last - 1, 22).getValues())
    .filter(function(r) {
      if (!r[1]) return false;  // bỏ dòng không có tên việc
      if (String(r[15] || '').toLowerCase().trim() === 'inactive') return false;
      return true;  // push TẤT CẢ loại: Nhật, Việt, Nhật-Việt, Kỹ sư đầu Nhật/Việt
    });

  const payload = rows.map(function(r, i) { return buildJob(r, i); });

  _pushToOne(SB_MP_URL, SB_MP_KEY, payload);
  return payload.length;
}

// Wrapper với UI
function syncToSupabase() {
  try {
    const n = _pushSupabase();
    SpreadsheetApp.getUi().alert('✅ Sync thành công!\nĐã đẩy ' + n + ' jobs lên maiphuongtokutei.com');
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Lỗi: ' + e.message);
  }
}

// ── DROPDOWN cột "Loại đơn" ───────────────────────────────────
const LOAI_DON_OPTIONS = ['Nhật', 'Việt', 'Nhật-Việt', 'Indonesia', 'Kỹ sư đầu Nhật', 'Kỹ sư đầu Việt'];

function setupLoaiDonDropdown() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(LOAI_DON_OPTIONS, true)
    .setAllowInvalid(true)
    .build();
  [SHEET_NAME, INACTIVE_TAB].forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const rows = sheet.getMaxRows() - 1;
    if (rows > 0) sheet.getRange(2, 1, rows, 1).setDataValidation(rule);
  });
  ss.toast('✅ Đã cập nhật dropdown "Loại đơn".', 'TokuteiJob', 4);
}

// ── SETUP: THIẾT LẬP SHEET (chạy 1 lần) ─────────────────────
function setupSourceSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  const PUB_HDR  = ['Loại đơn','Tên công việc','Ngành chính','Ngành phụ','Tỉnh/TP','Lương',
                    'Tiếng Nhật','Giới tính','Giờ làm','Ngày nghỉ','Tăng ca',
                    'Tăng lương','Thưởng','Nhà ở','Mô tả','Trạng thái'];
  const PRIV_HDR = ['Nguồn','Hoa hồng','Phí net','Ngày cập nhật'];

  sheet.getRange(1, 1, 1, 16)
    .setValues([PUB_HDR])
    .setBackground('#1a3c5e').setFontColor('#fff')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(1, 17, 1, 4)
    .setValues([PRIV_HDR])
    .setBackground('#4a3580').setFontColor('#fff')
    .setFontWeight('bold').setHorizontalAlignment('center');
  // Cột U (21) = dự phòng, cột V (22) = Link ảnh
  sheet.getRange(1, 21).setValue('(dự phòng)').setBackground('#e8eaed').setFontColor('#666');
  sheet.getRange(1, 22).setValue('Link ảnh').setBackground('#1a3c5e').setFontColor('#fff')
    .setFontWeight('bold').setHorizontalAlignment('center');

  sheet.setFrozenRows(1);

  [100,160,160,130,120,200,110,110,130,130,100,100,110,120,300,90, // A–P
   150,110,110,140,  // Q–T
   80, 240,          // U (dự phòng), V (Link ảnh)
  ].forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });

  const lastRow = Math.max(sheet.getLastRow(), 2);
  sheet.getRange(2, 20, lastRow - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');

  // Ẩn cột nội bộ Q–T
  for (let c = 17; c <= 20; c++) sheet.hideColumns(c);

  SpreadsheetApp.getUi().alert(
    '✅ Sheet đã sẵn sàng!\n\n' +
    'Cột A–P: dữ liệu công khai\n' +
    'Cột Q–T: nội bộ (ẩn)\n' +
    'Cột V: Link ảnh\n\n' +
    'Tiếp theo: Menu → 🌐 Tạo Sheet Công Khai'
  );
}

// ── HIGHLIGHT ROW & COLUMN KHI CHỌN Ô ───────────────────────
const HL_COLOR = '#DBEAFE';

function onSelectionChange(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  const row     = e.range.getRow();
  const col     = e.range.getColumn();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  sheet.getRange(2, 1, lastRow - 1, lastCol).setBackground(null);
  // Khôi phục nền tím cột nội bộ Q–T sau khi xóa highlight
  if (lastCol >= 20) sheet.getRange(2, 17, lastRow - 1, 4).setBackground('#f3f0ff');
  if (row >= 2) sheet.getRange(row, 1, 1, lastCol).setBackground(HL_COLOR);
  sheet.getRange(2, col, lastRow - 1, 1).setBackground(HL_COLOR);
}
