const XLSX = require('xlsx');

const wb = XLSX.utils.book_new();

// ── SHEET 1: TỔNG HỢP WEB (chỉ header, dữ liệu do QUERY tự fill) ──
const s1 = XLSX.utils.aoa_to_sheet([
  ['type','title','category','city','salary','japanese','count','gender','workHours','daysOff','overtime','raise','bonus','housing','status','id'],
  ['← Dòng này do công thức QUERY tự điền từ 3 sheets bên dưới. Không nhập tay vào đây.'],
]);
s1['!cols'] = [8,22,22,18,18,14,8,12,16,16,12,10,10,16,10,14].map(w=>({wch:w}));
XLSX.utils.book_append_sheet(wb, s1, 'TỔNG HỢP WEB');

// ── SHEET 2: Tokutei Đầu Nhật ──
const hdTokutei = [
  'Tên đơn hàng','Ngành nghề','Tỉnh / Thành phố','Mức lương (¥/tháng)',
  'Tiếng Nhật yêu cầu','Số lượng tuyển','Giới tính','Giờ làm việc',
  'Ngày nghỉ','Tăng ca','Tăng lương','Thưởng','Hỗ trợ nhà ở',
  'Mô tả công việc','Trạng thái'
];
const s2 = XLSX.utils.aoa_to_sheet([
  hdTokutei,
  ['Bánh gạo - Mỳ ăn liền','Chế biến thực phẩm','Tochigi - Oyama','1080y/h','N4',5,'Nam','8:00 - 17:00','Thứ 7 & Chủ nhật','20-40h/tháng','Có','Có','Miễn phí','Sản xuất bánh gạo và nguyên liệu khô cho mỳ ăn liền','active'],
  ['Làm bánh Pizza','Chế biến thực phẩm','Yamagata - Yonezawa','1100y/h','N4',5,'Nam','8:00 - 17:00','Theo lịch công ty','40h/tháng','Có','~10man','2man/tháng','Làm việc tại dây chuyền chế biến thực phẩm đông lạnh. Bao ký túc xá 2 bữa/ngày','active'],
  ['Công nhân xây dựng','Xây dựng','Osaka','200.000 - 240.000','N4',3,'Nam','8:00 - 17:00','Chủ nhật','Có','Có','Có','Hỗ trợ','Gia công cốp pha lắp dựng khung bê tông. OT được trả 125%','active'],
  ['Điều dưỡng viên','Điều dưỡng','Hokkaido','210.000 - 250.000','N3',2,'Không yêu cầu','Ca ngày','2 ngày/tuần','Có','6 tháng/lần','Có','Gần nơi làm','Chăm sóc người cao tuổi hỗ trợ sinh hoạt hàng ngày','active'],
]);
s2['!cols'] = [24,18,18,16,14,8,12,16,16,14,10,10,14,36,10].map(w=>({wch:w}));
XLSX.utils.book_append_sheet(wb, s2, 'Tokutei Đầu Nhật');

// ── SHEET 3: Tokutei Đầu Việt ──
const s3 = XLSX.utils.aoa_to_sheet([
  hdTokutei,
  ['Làm bánh mỳ','Chế biến thực phẩm','Osaka','20man','N4',5,'Nam','8:00 - 17:00','Thứ 7 & Chủ nhật','Có','Có','Cuối năm','Miễn phí','Sản xuất các loại bánh mỳ theo dây chuyền công nghiệp','active'],
  ['Làm bánh mochi','Chế biến thực phẩm','Tokyo','','N5',6,'Nam','8:00 - 17:00','Thứ 7 & Chủ nhật','Có','Có','Không','Hỗ trợ','Sản xuất bánh mochi truyền thống tại nhà máy','active'],
  ['Công nhân nông nghiệp','Nông nghiệp','Aichi','160.000 - 190.000','N5',10,'Không yêu cầu','8 tiếng/ngày','Thứ 7 & Chủ nhật','Có','Theo năng lực','Không','Miễn phí','Trồng và thu hoạch rau củ tại trang trại. Phù hợp người mới không cần kinh nghiệm','active'],
  ['Hộ lý tại viện dưỡng lão','Điều dưỡng','Kanagawa','195.000 - 230.000','N3',4,'Nữ ưu tiên','8 tiếng/ngày (ca)','2 ngày/tuần','Có','6 tháng/lần','Có','Gần nơi làm','Chăm sóc người cao tuổi tại viện dưỡng lão. Được đào tạo nghiệp vụ tại chỗ','active'],
]);
s3['!cols'] = [24,18,18,16,14,8,12,16,16,14,10,10,14,36,10].map(w=>({wch:w}));
XLSX.utils.book_append_sheet(wb, s3, 'Tokutei Đầu Việt');

// ── SHEET 4: Kỹ sư ──
const s4 = XLSX.utils.aoa_to_sheet([
  ['Vị trí tuyển dụng','Chuyên ngành','Địa điểm','Mức lương','Tiếng Nhật / Tiếng Anh',
   'Số lượng','Giới tính','Giờ làm việc','Ngày nghỉ','Tăng ca',
   'Tăng lương','Thưởng','Hỗ trợ nhà ở','Mô tả / Yêu cầu','Trạng thái'],
  ['Kỹ sư Vận hành máy - NVCT','Cơ khí','Yamanashi','Nenshu 350-600man','N2',2,'Nam','8:00 - 17:00','Theo lịch công ty','10-45h/tháng','Có','Không','2man/tháng','Vận hành và bảo trì máy móc công nghiệp. Công ty có nhiều trợ cấp','active'],
  ['Kỹ sư xây dựng','Xây dựng','Mie - Komono','1300y/h','N4',2,'Nam','8:00 - 17:00','Theo lịch công ty','10-45h/tháng','Có','Không','2man/tháng','Vệ sinh sửa chữa giàn giáo tại công trình','active'],
  ['Lập trình viên Web/Mobile','Lập trình viên','Osaka','250.000 - 350.000','N2',2,'Không yêu cầu','Linh hoạt (remote/hybrid)','Thứ 7 & Chủ nhật','Không bắt buộc','Hàng năm','Thưởng hiệu suất','Hỗ trợ tìm nhà','Phát triển ứng dụng web/mobile. Tech stack: React, Node.js hoặc Java. Môi trường quốc tế','active'],
]);
s4['!cols'] = [26,16,16,18,16,8,12,20,16,14,10,14,14,36,10].map(w=>({wch:w}));
XLSX.utils.book_append_sheet(wb, s4, 'Kỹ sư');

// Xuất file
XLSX.writeFile(wb, 'NipponTokutei-JobSheets.xlsx');
console.log('✅ Tạo xong: NipponTokutei-JobSheets.xlsx');
