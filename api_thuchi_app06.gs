/* ============================================================================
 * API BÁO CÁO THU CHI — PHI LONG / CARDIY (bản cho app PhiLong Auto 06)
 * ----------------------------------------------------------------------------
 * Nguồn : file "CARDIY MINI_CRM_KQKD_2026 - APP 06"  — tab LIVE_TC
 *         (LIVE_TC là IMPORTRANGE CHỈ ĐỌC từ file philongauto, không ghi gì vào app cũ)
 * Dùng cho trang: philonggroup.github.io/CARDIY5WAY/philongauto.baocaothuchi.html
 *
 * LIVE_TC KHÔNG có dòng tiêu đề (dữ liệu bắt đầu từ dòng 2). Cột cố định:
 *   A id | B THU/CHI | C mã phiếu | D phân nhóm | E đối tác | F đối tác đầy đủ
 *   G id đơn hàng | H số tiền | I lý do | J hình thức TT | K người tạo
 *   L ngày | M chi nhánh
 * Nếu sau này tab có dòng tiêu đề thì script tự nhận theo tên cột.
 *
 * GỌI:
 *   ?tungay=18/10/2026&denngay=24/10/2026&chinhanh=CN1
 *   ?ngay=18/10/26 - 24/10/26&chinhanh=CN1 , CN2 , CN3     (dạng AppSheet gửi)
 *   ?debug=1                                               (in 3 dòng đầu để tra soát)
 *
 * TRẢ: giữ NGUYÊN cấu trúc JSON của API cũ để trang HTML chạy không cần sửa.
 * ==========================================================================*/

var FILE_ID = '1EOQBrj9atyqZa4LMe2HSf28MFFCO4QtvkgL5mw0cqYo';
var TAB     = 'LIVE_TC';

/* cột cố định khi tab không có dòng tiêu đề (0-based) */
var C_LOAI = 1, C_SO = 2, C_NHOM = 3, C_DT = 4, C_DT2 = 5,
    C_DH = 6, C_TIEN = 7, C_LYDO = 8, C_HT = 9, C_NGUOI = 10,
    C_NGAY = 11, C_CN = 12;

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    if (String(p.debug || '') === '1') return json(debugTab_());
    return json(apiThuChi(p));
  } catch (err) {
    return json({ ok: false, loi: String(err && err.message ? err.message : err) });
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ───────────────────────────── SỐ & NGÀY ─────────────────────────────── */
function parseNum(raw) {
  if (raw === null || raw === undefined || raw === '') return 0;
  if (typeof raw === 'number') return raw;
  var t = String(raw).trim();
  if (!t || t === '-' || t === '—') return 0;
  var am = /^\(.*\)$/.test(t) || /^-/.test(t);
  t = t.replace(/[()\-\s]/g, '').replace(/[^0-9.,]/g, '');
  if (!t) return 0;
  var coC = t.indexOf(',') > -1, coD = t.indexOf('.') > -1, v;
  if (coC && coD) {
    if (t.lastIndexOf(',') > t.lastIndexOf('.')) v = parseFloat(t.replace(/\./g, '').replace(',', '.'));
    else v = parseFloat(t.replace(/,/g, ''));
  } else if (coC || coD) {
    var d = coC ? ',' : '.';
    var k = t.split(d);
    if (k.length > 2 || k[k.length - 1].length === 3) v = parseFloat(k.join(''));
    else v = parseFloat(k.join('.'));
  } else v = parseFloat(t);
  return isFinite(v) ? (am ? -v : v) : 0;
}

function p2(n) { return (n < 10 ? '0' : '') + n; }
function d2s(d) { return p2(d.getDate()) + '/' + p2(d.getMonth() + 1) + '/' + d.getFullYear(); }

function toDate(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  var s = String(v == null ? '' : v).trim();
  if (!s) return null;
  var m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (m) { var y = +m[3]; if (y < 100) y += 2000; return new Date(y, +m[2] - 1, +m[1]); }
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  return null;
}

/* ───────────────────── nhận diện cột (có/không tiêu đề) ───────────────── */
function mapCot_(vals) {
  var hdr = (vals[0] || []).map(function (h) { return String(h || '').trim().toLowerCase(); });
  function idx(name) { return hdr.indexOf(name); }
  if (idx('loai_phieu') > -1 && idx('thanh_tien') > -1 && idx('ngay_tao') > -1) {
    return {
      batDau: 1,
      loai: idx('loai_phieu'), so: idx('id_tc'), nhom: idx('phan_nhom'),
      dt: idx('doi_tac'), dt2: idx('ten_kh'), dh: idx('id_dh'),
      tien: idx('thanh_tien'), lydo: idx('ly_do'), ht: idx('hinh_thuc_tt'),
      nguoi: idx('nguoi_tao'), ngay: idx('ngay_tao'), cn: idx('chi_nhanh'),
      kieu: 'theo tiêu đề'
    };
  }
  return {
    batDau: 0,
    loai: C_LOAI, so: C_SO, nhom: C_NHOM, dt: C_DT, dt2: C_DT2, dh: C_DH,
    tien: C_TIEN, lydo: C_LYDO, ht: C_HT, nguoi: C_NGUOI, ngay: C_NGAY, cn: C_CN,
    kieu: 'cột cố định'
  };
}

/* ───────────────────────────── API CHÍNH ─────────────────────────────── */
function apiThuChi(p) {
  /* 1. Kỳ báo cáo */
  var tu = toDate(p.tungay), den = toDate(p.denngay);
  var ngay = String(p.ngay || '').trim();
  if ((!tu || !den) && ngay) {
    var ph = ngay.split(/\s+-\s+|\s-\s/);
    if (ph.length >= 2) { tu = tu || toDate(ph[0]); den = den || toDate(ph[1]); }
    else { tu = tu || toDate(ngay); den = den || toDate(ngay); }
  }
  var hnay = new Date();
  if (!den) den = new Date(hnay.getFullYear(), hnay.getMonth(), hnay.getDate());
  if (!tu) tu = den;
  if (tu.getTime() > den.getTime()) { var x = tu; tu = den; den = x; }

  /* 2. Chi nhánh */
  var cnList = String(p.chinhanh || '').toUpperCase()
    .split(/[,;|]+/).map(function (s) { return s.trim(); }).filter(function (s) { return s; });
  var locCN = {}, coLocCN = cnList.length > 0;
  cnList.forEach(function (c) { locCN[c] = 1; });

  /* 3. Đọc tab */
  var sh = SpreadsheetApp.openById(FILE_ID).getSheetByName(TAB);
  if (!sh) return { ok: false, loi: 'Không tìm thấy tab ' + TAB + '.' };
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return { ok: false, loi: 'Tab ' + TAB + ' không có dữ liệu.' };
  var c = mapCot_(vals);

  var rows = [], boQua = 0, tongDong = 0;
  var tTu = tu.getTime(), tDen = den.getTime();

  for (var r = c.batDau; r < vals.length; r++) {
    var row = vals[r];
    var loai = String(row[c.loai] || '').trim().toUpperCase();
    if (loai !== 'THU' && loai !== 'CHI') continue;
    tongDong++;

    var d = toDate(row[c.ngay]);
    if (!d) { boQua++; continue; }
    var t = d.getTime();
    if (t < tTu || t > tDen) continue;

    var cn = String(row[c.cn] || '').trim().toUpperCase();
    if (coLocCN && !locCN[cn]) continue;

    var ten = c.dt >= 0 ? String(row[c.dt] || '').trim() : '';
    if (!ten && c.dt2 >= 0) ten = String(row[c.dt2] || '').trim();

    rows.push({
      ngay: d2s(d),
      ts: t,
      loai: loai,
      sophieu: c.so >= 0 ? String(row[c.so] || '').trim() : '',
      nhom: c.nhom >= 0 ? String(row[c.nhom] || '').trim() : '',
      doitac: ten,
      ncc: '',
      tien: parseNum(row[c.tien]),
      lydo: c.lydo >= 0 ? String(row[c.lydo] || '').trim() : '',
      ht: (c.ht >= 0 ? String(row[c.ht] || '').trim().toUpperCase() : '') || 'TM',
      cn: cn,
      iddh: c.dh >= 0 ? String(row[c.dh] || '').trim() : '',
      nguoi: c.nguoi >= 0 ? String(row[c.nguoi] || '').trim() : ''
    });
  }

  rows.sort(function (a, b) { return a.ts - b.ts || (a.sophieu < b.sophieu ? -1 : 1); });

  return {
    ok: true,
    ky: { tu: d2s(tu), den: d2s(den) },
    chinhanh: coLocCN ? cnList.join(', ') : '',
    nguon: 'CARDIY MINI_CRM_KQKD_2026 - APP 06 / tab LIVE_TC (' + c.kieu + ')',
    tinhLuc: Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm'),
    tongDong: tongDong,
    soDong: rows.length,
    boQuaKhongNgay: boQua,
    rows: rows
  };
}

/* ──────────────────── tra soát nhanh: ?debug=1 ───────────────────────── */
function debugTab_() {
  var sh = SpreadsheetApp.openById(FILE_ID).getSheetByName(TAB);
  if (!sh) return { ok: false, loi: 'Không tìm thấy tab ' + TAB + '.' };
  var vals = sh.getRange(1, 1, Math.min(4, sh.getLastRow()), Math.min(15, sh.getLastColumn())).getValues();
  return { ok: true, tab: TAB, soDong: sh.getLastRow(), soCot: sh.getLastColumn(), mau: vals };
}

/* Chạy tay trong editor để kiểm tra */
function test_() {
  var kq = apiThuChi({ tungay: '01/10/2026', denngay: '31/10/2026', chinhanh: 'CN1 , CN2 , CN3' });
  var thu = 0, chi = 0;
  kq.rows.forEach(function (x) { if (x.loai === 'THU') thu += x.tien; else chi += x.tien; });
  Logger.log('%s dòng | THU %s | CHI %s | bỏ qua %s | nguồn %s',
    kq.soDong, thu.toLocaleString('vi-VN'), chi.toLocaleString('vi-VN'), kq.boQuaKhongNgay, kq.nguon);
}
