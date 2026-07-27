/*************************************************************
 *  HỆ THỐNG BIỂU MẪU NỘI BỘ — PCCC PHÚ QUÝ
 *  File: Code.gs  (backend Apps Script)
 *
 *  Kiến trúc nhiều trang: mỗi tính năng là 1 file .html riêng.
 *  Router doGet() chọn trang theo tham số ?page=...
 *  Tính năng mới sau này: tạo file html mới + thêm tên vào TRANG_HOP_LE.
 *************************************************************/

/*==================== CẤU HÌNH (SỬA Ở ĐÂY) ====================*/
// 1) ID thư mục Drive để lưu file Word (lấy từ URL thư mục, xem hướng dẫn)
var FOLDER_ID = 'DAN_ID_THU_MUC_DRIVE_VAO_DAY';

// 2) Email nhận tin nội bộ (tạm thay Zalo). Để trống '' nếu không muốn gửi mail.
var EMAIL_NHAN_MAC_DINH = 'ndanha23006@cusc.ctu.edu.vn';

// 3) true = gửi kèm email khi tạo phiếu | false = chỉ tạo file Word, không gửi mail
var GUI_EMAIL = true;

// 4) ID thư mục Drive chứa các file Google Sheets "Báo cáo công việc tuần"
var BAOCAO_FOLDER_ID = '1W-F_GHTIf--zuY0v79D0YfkB6KNE9EXC';

// Danh sách trang hợp lệ (thêm trang mới vào đây khi phát triển tính năng)
var TRANG_HOP_LE = ['index', 'phanhoi', 'baocaohub', 'baocao'];
/*=============================================================*/


/** Router: chọn trang theo ?page= */
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'index';
  if (TRANG_HOP_LE.indexOf(page) < 0) page = 'index';

  var t = HtmlService.createTemplateFromFile(page);
  t.baseUrl = ScriptApp.getService().getUrl(); // URL gốc web app, để điều hướng giữa các trang
  t.fileId = (e && e.parameter && e.parameter.file) ? e.parameter.file : ''; // ID file tuần (trang baocao)
  return t.evaluate()
    .setTitle('PCCC Phú Quý · Hệ thống nội bộ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Cho phép nhúng file html khác (dùng cho CSS dùng chung: styles.html) */
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}


/**
 * XỬ LÝ KHI GỬI BIỂU MẪU "XỬ LÝ PHẢN HỒI TỔ ĐỘI"
 * - Chốt chặn trường bắt buộc phía server
 * - Tạo 1 file Word (.docx) trong thư mục Drive
 * - (tuỳ chọn) Gửi email nội bộ kèm file Word
 * Trả về: {ok, ma, fileUrl, fileName}  hoặc  {ok:false, thieu:[...]}
 */
function luuPhanHoi(d) {
  // 1) Kiểm tra bắt buộc
  var buoc = [
    ['nguoiTiepNhan', 'Người tiếp nhận'], ['tenToDoi', 'Tên tổ đội'], ['kenh', 'Kênh phản hồi'],
    ['loai', 'Loại phản hồi'], ['noiDung', 'Nội dung phản hồi'], ['uuTien', 'Mức độ ưu tiên'],
    ['nguoiPhuTrach', 'Người phụ trách xử lý'], ['hanPhanHoi', 'Hạn phản hồi lại']
  ];
  var thieu = buoc.filter(function (b) { return !d[b[0]] || !String(d[b[0]]).trim(); })
                  .map(function (b) { return b[1]; });
  if (thieu.length) return { ok: false, thieu: thieu };

  // 2) Mã phiếu
  var now = new Date();
  var ma = 'PH-' + Utilities.formatDate(now, 'GMT+7', 'yyyyMMdd-HHmmss');
  var som = (d.uuTien === 'Cao' || d.uuTien === 'Khẩn cấp');

  // 3) Tạo Google Doc rồi xuất .docx vào thư mục Drive
  var fileName = ma + ' - ' + d.tenToDoi;
  var doc = DocumentApp.create(fileName);
  var body = doc.getBody();

  body.appendParagraph('PHIẾU XỬ LÝ PHẢN HỒI TỔ ĐỘI')
      .setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('CÔNG TY TNHH XD TM PHÚ QUÝ — Bộ phận quản lý dự án & thi công')
      .setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
  body.appendParagraph('Mã phiếu: ' + ma + '     |     Mức độ ưu tiên: ' + d.uuTien + (som ? '  (⚠ CẦN XỬ LÝ SỚM)' : ''));

  var rows = [
    ['Người tiếp nhận', d.nguoiTiepNhan],
    ['Thời gian tiếp nhận', fmt(d.thoiGian)],
    ['Tổ đội', d.tenToDoi + (d.maToDoi ? ' (' + d.maToDoi + ')' : '')],
    ['Kênh phản hồi', d.kenh],
    ['Tình huống liên quan', orDash(d.tinhHuong)],
    ['Loại phản hồi', d.loai],
    ['Mức độ ưu tiên', d.uuTien],
    ['Nội dung phản hồi', d.noiDung],
    ['Người phụ trách xử lý', d.nguoiPhuTrach],
    ['Hạn tổ đội phản hồi lại', fmt(d.hanPhanHoi)],
    ['Hướng xử lý ban đầu', orDash(d.huongXuLy)],
    ['Ghi chú nội bộ', orDash(d.ghiChu)]
  ];
  var table = body.appendTable(rows);
  for (var i = 0; i < rows.length; i++) {
    table.getCell(i, 0).setWidth(170);
    table.getCell(i, 0).getChild(0).asParagraph().editAsText().setBold(true);
  }

  body.appendParagraph('');
  body.appendParagraph('— Phiếu được tạo tự động từ hệ thống nội bộ PCCC Phú Quý lúc ' + fmt2(now))
      .editAsText().setItalic(true).setForegroundColor('#7a6f5c');

  doc.saveAndClose();

  // Xuất .docx (Word thật) và lưu vào thư mục
  var id = doc.getId();
  var exportUrl = 'https://docs.google.com/feeds/download/documents/export/Export?id=' + id + '&exportFormat=docx';
  var blob = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  }).getBlob().setName(fileName + '.docx');

  var folder = DriveApp.getFolderById(FOLDER_ID);
  var wordFile = folder.createFile(blob);
  DriveApp.getFileById(id).setTrashed(true); // xoá bản Google Doc tạm, chỉ giữ .docx

  // 4) (tuỳ chọn) Gửi email nội bộ kèm file Word
  var emailTo = (d.emailNhan && String(d.emailNhan).trim()) ? String(d.emailNhan).trim() : EMAIL_NHAN_MAC_DINH;
  if (GUI_EMAIL && emailTo) {
    var flag = d.uuTien === 'Khẩn cấp' ? '🔴 KHẨN CẤP'
             : d.uuTien === 'Cao' ? '🟠 ƯU TIÊN CAO'
             : d.uuTien === 'Trung bình' ? '🟡 TRUNG BÌNH' : '🟢 THẤP';
    var msg =
      '[PHẢN HỒI TỔ ĐỘI — CẦN XỬ LÝ] ' + flag + '\n' +
      '• Tổ đội: ' + d.tenToDoi + (d.maToDoi ? ' (' + d.maToDoi + ')' : '') + '\n' +
      '• Loại phản hồi: ' + d.loai + '\n' +
      (d.tinhHuong ? '• Tình huống: ' + d.tinhHuong + '\n' : '') +
      '• Nội dung: ' + d.noiDung + '\n' +
      '• Người phụ trách xử lý: ' + d.nguoiPhuTrach + '\n' +
      '• Hạn tổ đội phản hồi lại: ' + fmt(d.hanPhanHoi) + '\n' +
      (d.huongXuLy ? '• Hướng xử lý ban đầu: ' + d.huongXuLy + '\n' : '') +
      '—\nTiếp nhận bởi ' + d.nguoiTiepNhan + ' · ' + fmt(d.thoiGian) + '\n\n' +
      'File phiếu (Word): ' + wordFile.getUrl();
    MailApp.sendEmail({
      to: emailTo,
      subject: '[Phản hồi tổ đội] ' + d.tenToDoi + ' - ' + d.loai + ' (' + d.uuTien + ')',
      body: msg,
      attachments: [wordFile.getAs('application/vnd.openxmlformats-officedocument.wordprocessingml.document')]
    });
  }

  return { ok: true, ma: ma, fileUrl: wordFile.getUrl(), fileName: wordFile.getName() };
}

/* ---------- Helper ---------- */
// datetime-local "2026-07-22T14:30" -> "22/07/2026 14:30" (không đổi múi giờ)
function fmt(v) {
  if (!v) return '(chưa nhập)';
  var m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(v);
  if (!m) return v;
  return m[3] + '/' + m[2] + '/' + m[1] + ' ' + m[4] + ':' + m[5];
}
function fmt2(dateObj) { return Utilities.formatDate(dateObj, 'GMT+7', 'dd/MM/yyyy HH:mm'); }
function orDash(s) { return (s && String(s).trim()) ? s : '(không nhập)'; }

/** Chạy 1 lần để kiểm tra cấu hình thư mục (tuỳ chọn) */
function kiemTraCauHinh() {
  var f = DriveApp.getFolderById(FOLDER_ID);
  Logger.log('OK - Thư mục lưu trữ (phiếu Word): ' + f.getName());
  var g = DriveApp.getFolderById(BAOCAO_FOLDER_ID);
  Logger.log('OK - Thư mục báo cáo tuần: ' + g.getName());
}


/*******************************************************
 *  BÁO CÁO CÔNG VIỆC TUẦN
 *  - Mỗi tuần = 1 file Google Sheets trong BAOCAO_FOLDER_ID
 *  - Mỗi tổ đội = 1 sheet (tab) riêng trong file tuần
 *******************************************************/

var BC_HEADER = ['STT', 'Tên công trình', 'Hạng mục', 'Khối lượng thực hiện',
                 'Khó khăn', 'Tiến độ cam kết hoàn thành', 'Yêu cầu khác',
                 'Người nhập', 'Thời gian nhập'];

/** Sheet không phải tổ đội (bỏ qua khi liệt kê đội) */
function _isMetaSheet(n) { return n === 'HƯỚNG DẪN' || String(n).charAt(0) === '_'; }

/** Liệt kê các file tuần trong thư mục */
function listWeeks() {
  var folder = DriveApp.getFolderById(BAOCAO_FOLDER_ID);
  var it = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var out = [];
  while (it.hasNext()) {
    var f = it.next();
    var ss = SpreadsheetApp.openById(f.getId());
    var teams = ss.getSheets().filter(function (s) { return !_isMetaSheet(s.getName()); }).length;
    out.push({
      id: f.getId(), name: f.getName(), url: f.getUrl(), teams: teams,
      updated: Utilities.formatDate(f.getLastUpdated(), 'GMT+7', 'dd/MM/yyyy HH:mm')
    });
  }
  out.sort(function (a, b) { return b.name.localeCompare(a.name, 'vi', { numeric: true }); });
  return out;
}

/** Tạo file tuần mới */
function createWeek(soTuan, tuNgay, denNgay) {
  soTuan = String(soTuan || '').trim();
  if (!soTuan) return { ok: false, msg: 'Chưa nhập số tuần' };
  var name = _weekFileName(soTuan, tuNgay, denNgay);
  var ss = SpreadsheetApp.create(name);
  // chuyển file vào đúng thư mục
  var file = DriveApp.getFileById(ss.getId());
  DriveApp.getFolderById(BAOCAO_FOLDER_ID).addFile(file);
  try { DriveApp.getRootFolder().removeFile(file); } catch (e) {}
  // sheet mặc định -> trang hướng dẫn
  var sh0 = ss.getSheets()[0];
  sh0.setName('HƯỚNG DẪN');
  sh0.getRange(1, 1).setValue('Báo cáo công việc tuần ' + soTuan +
    '. Mỗi tổ đội là một sheet (tab) riêng — tạo tự động khi đội gửi báo cáo. Hệ thống nội bộ PCCC Phú Quý.');
  return { ok: true, id: ss.getId(), name: name, url: ss.getUrl() };
}

/** Thông tin 1 file tuần + danh sách đội đã có */
function getWeekInfo(fileId) {
  if (!fileId) return { ok: false, msg: 'Thiếu mã file tuần' };
  var f = DriveApp.getFileById(fileId);
  var ss = SpreadsheetApp.openById(fileId);
  var teams = ss.getSheets()
    .filter(function (s) { return !_isMetaSheet(s.getName()); })
    .map(function (s) { return s.getName(); });
  return { ok: true, name: f.getName(), url: f.getUrl(), teams: teams };
}

/** Danh sách tổ đội (sheet) trong 1 file tuần */
function listTeams(fileId) {
  if (!fileId) return [];
  var ss = SpreadsheetApp.openById(fileId);
  return ss.getSheets()
    .filter(function (s) { return !_isMetaSheet(s.getName()); })
    .map(function (s) { return s.getName(); });
}

/** Ghi báo cáo: append các dòng công trình vào sheet của tổ đội (tạo mới nếu chưa có) */
function addReport(fileId, tenDoi, rows) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!fileId) return { ok: false, msg: 'Thiếu mã file tuần' };
    tenDoi = String(tenDoi || '').trim();
    if (!tenDoi) return { ok: false, msg: 'Chưa chọn/nhập tổ đội', field: 'tenDoi' };
    if (!rows || !rows.length) return { ok: false, msg: 'Chưa có dòng công trình nào' };

    // Kiểm tra bắt buộc phía server: Công trình, Hạng mục, Khối lượng
    var thieu = [];
    rows.forEach(function (r, i) {
      var n = i + 1;
      if (!r.congTrinh || !String(r.congTrinh).trim()) thieu.push('Dòng ' + n + ': thiếu Tên công trình');
      if (!r.hangMuc || !String(r.hangMuc).trim()) thieu.push('Dòng ' + n + ': thiếu Hạng mục');
      if (!r.khoiLuong || !String(r.khoiLuong).trim()) thieu.push('Dòng ' + n + ': thiếu Khối lượng thực hiện');
    });
    if (thieu.length) return { ok: false, msg: thieu.join('\n'), thieu: thieu };

    var ss = SpreadsheetApp.openById(fileId);
    var sh = _teamSheet(ss, tenDoi);
    var email = Session.getActiveUser().getEmail() || '';
    var now = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm');
    var soCu = Math.max(sh.getLastRow() - 2, 0); // số dòng dữ liệu đã có (2 dòng đầu = tiêu đề)
    var values = rows.map(function (r, i) {
      return [soCu + i + 1, String(r.congTrinh).trim(), String(r.hangMuc).trim(),
              String(r.khoiLuong).trim(), (r.khoKhan || '').trim(), (r.tienDo || '').trim(),
              (r.yeuCau || '').trim(), email, now];
    });
    sh.getRange(sh.getLastRow() + 1, 1, values.length, BC_HEADER.length).setValues(values);
    return { ok: true, added: values.length, tenDoi: tenDoi, url: ss.getUrl() + '#gid=' + sh.getSheetId() };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Helper báo cáo tuần ---------- */
function _teamSheet(ss, team) {
  var name = _sanitizeSheetName(team);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, BC_HEADER.length).merge()
      .setValue('BÁO CÁO CÔNG VIỆC TUẦN — ĐỘI: ' + team)
      .setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
    sh.getRange(2, 1, 1, BC_HEADER.length).setValues([BC_HEADER])
      .setFontWeight('bold').setBackground('#f2eade');
    sh.setFrozenRows(2);
    sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 160); sh.setColumnWidth(4, 150);
    sh.setColumnWidth(5, 180); sh.setColumnWidth(6, 170); sh.setColumnWidth(7, 160);
  }
  return sh;
}
function _sanitizeSheetName(s) {
  return String(s).replace(/[\[\]\*\/\\\?\:]/g, ' ').replace(/\s+/g, ' ').trim().substring(0, 90) || 'Tổ đội';
}
function _weekFileName(soTuan, bd, kt) {
  var dm = function (v) { var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(v || ''); return m ? (m[3] + '.' + m[2]) : ''; };
  var dmy = function (v) { var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(v || ''); return m ? (m[3] + '.' + m[2] + '.' + m[1]) : ''; };
  var range = (bd && kt) ? ' (' + dm(bd) + '–' + dmy(kt) + ')' : '';
  return 'Báo cáo tuần ' + soTuan + range;
}
