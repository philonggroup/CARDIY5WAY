/* ============================================================================
 * PLG-EXCEL  —  CHUẨN XUẤT FILE EXCEL PHI LONG AUTO
 * ----------------------------------------------------------------------------
 * Sinh file .xlsx THẬT (OOXML + ZIP) ngay trên trình duyệt, KHÔNG cần internet,
 * KHÔNG cần thư viện ngoài. Dùng chung cho toàn bộ biểu mẫu CARDIY5WAY.
 *
 * CHUẨN ĐỊNH DẠNG (chốt 22/10/2026):
 *   - Phông chữ      : Times New Roman
 *   - Cỡ chữ         : nội dung 12 | tiêu đề 16 đậm | ghi chú 11 nghiêng
 *   - Màu thương hiệu: #0B5394 (nền tiêu đề bảng, chữ trắng đậm)
 *   - Dòng tổng cộng : nền #DCE6F1, chữ đậm
 *   - Khung viền     : thin, màu #000000, toàn bộ ô dữ liệu
 *   - Định dạng số   : #,##0 (tiền) | 0.0% (phần trăm) | ngày giữ dạng chữ
 *   - Trang in       : A4, fit 1 trang ngang, lặp dòng tiêu đề bảng
 *
 * CÁCH DÙNG: chèn trước thẻ đóng body
 *   <script src="plg-excel.js"></script>
 * Tuỳ biến (không bắt buộc):
 *   <script>window.PLG_EXCEL_OPTIONS = { tenBaoCao:'BÁO GIÁ', tenFile:'BaoGia' };</script>
 * Đánh dấu thủ công trong HTML (không bắt buộc):
 *   <table data-plg-sheet="CÔNG NỢ">   -> đặt tên sheet
 *   <table data-plg-skip>              -> bỏ qua bảng này
 *   <div data-plg-info="Biển số">51A-12345</div>
 * ==========================================================================*/
(function (global) {
  'use strict';

  /* ---------------------------------------------------------------- CHUẨN -- */
  var CHUAN = {
    font: 'Times New Roman',
    coChu: 12,
    coChuTieuDe: 16,
    mauThuongHieu: '124A95',   // xanh PLG - nền header bảng
    mauChuHeader: 'FFFFFF',
    mauDongTong: 'DCE6F1',
    mauVien: '000000',
    tenCongTy: 'CÔNG TY TNHH TM DỊCH VỤ Ô TÔ PHI LONG',
    rongCotMin: 8,
    rongCotMax: 46
  };

  /* ====================================================== 1. TIỆN ÍCH ZIP == */
  var CRC_TABLE = (function () {
    var t = new Int32Array(256), c, n, k;
    for (n = 0; n < 256; n++) {
      c = n;
      for (k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c;
    }
    return t;
  })();

  function crc32(buf) {
    var c = -1;
    for (var i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ -1) >>> 0;
  }

  function utf8(str) { return new TextEncoder().encode(str); }

  /* ZIP "stored" (không nén) — Excel đọc bình thường, file nhỏ nên không cần nén */
  function zip(files) {
    var parts = [], central = [], offset = 0, i;

    function u16(v) { return [v & 0xFF, (v >>> 8) & 0xFF]; }
    function u32(v) { return [v & 0xFF, (v >>> 8) & 0xFF, (v >>> 16) & 0xFF, (v >>> 24) & 0xFF]; }

    for (i = 0; i < files.length; i++) {
      var name = utf8(files[i].name);
      var data = files[i].data;
      var crc = crc32(data);
      var local = [].concat(
        u32(0x04034b50), u16(20), u16(0x0800), u16(0), u16(0), u16(0),
        u32(crc), u32(data.length), u32(data.length),
        u16(name.length), u16(0)
      );
      parts.push(new Uint8Array(local), name, data);
      central.push([].concat(
        u32(0x02014b50), u16(20), u16(20), u16(0x0800), u16(0), u16(0), u16(0),
        u32(crc), u32(data.length), u32(data.length),
        u16(name.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(offset)
      ), name);
      offset += local.length + name.length + data.length;
    }

    var cdStart = offset, cdSize = 0, flat = [];
    for (i = 0; i < central.length; i += 2) {
      var h = new Uint8Array(central[i]);
      flat.push(h, central[i + 1]);
      cdSize += h.length + central[i + 1].length;
    }
    var eocd = new Uint8Array([].concat(
      u32(0x06054b50), u16(0), u16(0),
      u16(files.length), u16(files.length), u32(cdSize), u32(cdStart), u16(0)
    ));

    var all = parts.concat(flat, [eocd]), total = 0;
    for (i = 0; i < all.length; i++) total += all[i].length;
    var out = new Uint8Array(total), p = 0;
    for (i = 0; i < all.length; i++) { out.set(all[i], p); p += all[i].length; }
    return out;
  }

  /* ====================================================== 2. TIỆN ÍCH XML == */
  function esc(s) {
    return String(s === null || s === undefined ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&apos;')
      /* loại ký tự điều khiển Excel không chấp nhận */
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
  }

  function colName(n) { // 1 -> A
    var s = '';
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
    return s;
  }

  /* ================================================ 3. STYLES.XML (CHUẨN) == */
  /* Chỉ số style dùng trong code:
     0 mặc định | 1 tiêu đề lớn | 2 tên công ty | 3 HEADER bảng | 4 ô chữ
     5 ô số #,##0 | 6 ô canh giữa | 7 tổng-chữ | 8 tổng-số | 9 nhãn thông tin
     10 giá trị thông tin | 11 tiêu đề mục | 12 ghi chú | 13 ô tiền có "đ"
     14 ô phần trăm                                                            */
  function stylesXml() {
    var F = CHUAN.font;
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0&quot; đ&quot;"/></numFmts>'
      + '<fonts count="6">'
      + '<font><sz val="' + CHUAN.coChu + '"/><color rgb="FF000000"/><name val="' + F + '"/></font>'
      + '<font><b/><sz val="' + CHUAN.coChuTieuDe + '"/><color rgb="FF' + CHUAN.mauThuongHieu + '"/><name val="' + F + '"/></font>'
      + '<font><b/><sz val="' + CHUAN.coChu + '"/><color rgb="FF000000"/><name val="' + F + '"/></font>'
      + '<font><b/><sz val="' + CHUAN.coChu + '"/><color rgb="FF' + CHUAN.mauChuHeader + '"/><name val="' + F + '"/></font>'
      + '<font><i/><sz val="11"/><color rgb="FF595959"/><name val="' + F + '"/></font>'
      + '<font><b/><sz val="13"/><color rgb="FF' + CHUAN.mauThuongHieu + '"/><name val="' + F + '"/></font>'
      + '</fonts>'
      + '<fills count="5">'
      + '<fill><patternFill patternType="none"/></fill>'
      + '<fill><patternFill patternType="gray125"/></fill>'
      + '<fill><patternFill patternType="solid"><fgColor rgb="FF' + CHUAN.mauThuongHieu + '"/><bgColor indexed="64"/></patternFill></fill>'
      + '<fill><patternFill patternType="solid"><fgColor rgb="FF' + CHUAN.mauDongTong + '"/><bgColor indexed="64"/></patternFill></fill>'
      + '<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill>'
      + '</fills>'
      + '<borders count="2">'
      + '<border><left/><right/><top/><bottom/><diagonal/></border>'
      + '<border>'
      + '<left style="thin"><color rgb="FF' + CHUAN.mauVien + '"/></left>'
      + '<right style="thin"><color rgb="FF' + CHUAN.mauVien + '"/></right>'
      + '<top style="thin"><color rgb="FF' + CHUAN.mauVien + '"/></top>'
      + '<bottom style="thin"><color rgb="FF' + CHUAN.mauVien + '"/></bottom>'
      + '<diagonal/></border>'
      + '</borders>'
      + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
      + '<cellXfs count="15">'
      /* 0 */ + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
      /* 1 */ + '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
      /* 2 */ + '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
      /* 3 */ + '<xf numFmtId="0" fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
      /* 4 */ + '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>'
      /* 5 */ + '<xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>'
      /* 6 */ + '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
      /* 7 */ + '<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>'
      /* 8 */ + '<xf numFmtId="3" fontId="2" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>'
      /* 9 */ + '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>'
      /*10 */ + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>'
      /*11 */ + '<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>'
      /*12 */ + '<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>'
      /*13 */ + '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>'
      /*14 */ + '<xf numFmtId="10" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>'
      + '</cellXfs>'
      + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
      + '</styleSheet>';
  }

  /* ==================================================== 4. DỰNG WORKSHEET == */
  function sheetXml(sheet) {
    var rows = sheet.rows, i, j, out = [], maxCol = 1;
    for (i = 0; i < rows.length; i++) if (rows[i].cells.length > maxCol) maxCol = rows[i].cells.length;

    /* cols */
    var cols = '<cols>';
    for (j = 0; j < maxCol; j++) {
      var w = sheet.widths[j] || CHUAN.rongCotMin;
      cols += '<col min="' + (j + 1) + '" max="' + (j + 1) + '" width="' + w.toFixed(2) + '" customWidth="1"/>';
    }
    cols += '</cols>';

    /* sheetData */
    for (i = 0; i < rows.length; i++) {
      var r = rows[i], rn = i + 1, cs = '';
      for (j = 0; j < r.cells.length; j++) {
        var c = r.cells[j];
        if (c === null || c === undefined) continue;
        var ref = colName(j + 1) + rn;
        var st = c.s || 0;
        if (c.t === 'n' && c.v !== null && c.v !== '' && isFinite(c.v)) {
          cs += '<c r="' + ref + '" s="' + st + '"><v>' + c.v + '</v></c>';
        } else if (c.v === '' || c.v === null || c.v === undefined) {
          cs += '<c r="' + ref + '" s="' + st + '"/>';
        } else {
          cs += '<c r="' + ref + '" s="' + st + '" t="inlineStr"><is><t xml:space="preserve">'
              + esc(c.v) + '</t></is></c>';
        }
      }
      out.push('<row r="' + rn + '"' + (r.h ? ' ht="' + r.h + '" customHeight="1"' : '') + '>' + cs + '</row>');
    }

    /* merges */
    var merges = '';
    if (sheet.merges.length) {
      merges = '<mergeCells count="' + sheet.merges.length + '">';
      for (i = 0; i < sheet.merges.length; i++) merges += '<mergeCell ref="' + sheet.merges[i] + '"/>';
      merges += '</mergeCells>';
    }

    var freeze = '';
    if (sheet.freezeRow) {
      freeze = '<pane ySplit="' + sheet.freezeRow + '" topLeftCell="A' + (sheet.freezeRow + 1)
             + '" activePane="bottomLeft" state="frozen"/>';
    }
    var af = sheet.autoFilter ? '<autoFilter ref="' + sheet.autoFilter + '"/>' : '';
    var landscape = maxCol > 6 ? 'landscape' : 'portrait';

    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
      + '<dimension ref="A1:' + colName(maxCol) + Math.max(rows.length, 1) + '"/>'
      + '<sheetViews><sheetView showGridLines="0" tabSelected="' + (sheet.first ? 1 : 0)
      + '" workbookViewId="0">' + freeze + '</sheetView></sheetViews>'
      + '<sheetFormatPr defaultRowHeight="18"/>'
      + cols
      + '<sheetData>' + out.join('') + '</sheetData>'
      + af
      + merges
      + '<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>'
      + '<pageSetup paperSize="9" orientation="' + landscape + '" fitToWidth="1" fitToHeight="0"/>'
      + '</worksheet>';
  }

  function buildXlsx(sheets) {
    var files = [], i;
    var ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
    var wbSheets = '', wbRels = '';
    for (i = 0; i < sheets.length; i++) {
      ct += '<Override PartName="/xl/worksheets/sheet' + (i + 1)
          + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
      wbSheets += '<sheet name="' + esc(sheets[i].name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>';
      wbRels += '<Relationship Id="rId' + (i + 1)
             + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'
             + ' Target="worksheets/sheet' + (i + 1) + '.xml"/>';
      files.push({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: utf8(sheetXml(sheets[i])) });
    }
    ct += '</Types>';
    wbRels += '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';

    files.unshift(
      { name: '[Content_Types].xml', data: utf8(ct) },
      { name: '_rels/.rels', data: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
          + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
          + '</Relationships>') },
      { name: 'xl/workbook.xml', data: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
          + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
          + '<sheets>' + wbSheets + '</sheets></workbook>') },
      { name: 'xl/_rels/workbook.xml.rels', data: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
          + wbRels + '</Relationships>') },
      { name: 'xl/styles.xml', data: utf8(stylesXml()) }
    );
    return zip(files);
  }

  /* ============================================== 5. ĐỌC SỐ KIỂU VIỆT NAM == */
  function parseSo(raw) {
    if (raw === null || raw === undefined) return null;
    var s = String(raw).trim();
    if (!s) return null;

    var phanTram = /%$/.test(s);
    var t = s.replace(/[\s ]/g, '').replace(/%$/, '')
             .replace(/(đ|₫|vnđ|vnd)$/i, '');
    var am = false;
    if (/^\(.*\)$/.test(t)) { am = true; t = t.slice(1, -1); }      // (1.000) = âm
    if (/^-/.test(t)) { am = true; t = t.slice(1); }
    if (!/^[0-9]+([.,][0-9]+)*$/.test(t)) return null;

    var coP = t.indexOf(',') >= 0, coC = t.indexOf('.') >= 0, v;
    if (coP && coC) {                                   // dấu xuất hiện sau cùng = thập phân
      var thapPhan = t.lastIndexOf(',') > t.lastIndexOf('.') ? ',' : '.';
      var ngan = thapPhan === ',' ? '.' : ',';
      v = parseFloat(t.split(ngan).join('').replace(thapPhan, '.'));
    } else if (coP || coC) {
      var d = coP ? ',' : '.';
      var khuc = t.split(d);
      // nhiều dấu, hoặc 1 dấu + đúng 3 chữ số phía sau => dấu ngăn hàng nghìn
      if (khuc.length > 2 || khuc[khuc.length - 1].length === 3) v = parseFloat(khuc.join(''));
      else v = parseFloat(khuc.join('.'));
    } else {
      v = parseFloat(t);
    }
    if (!isFinite(v)) return null;
    if (am) v = -v;
    return { v: phanTram ? v / 100 : v, phanTram: phanTram };
  }

  /* ========================================== 6. THU THẬP DỮ LIỆU TỪ TRANG == */
  function hien(el) {
    if (!el) return false;
    var st = global.getComputedStyle ? global.getComputedStyle(el) : null;
    if (st && (st.display === 'none' || st.visibility === 'hidden')) return false;
    return true;
  }

  function chu(el) {
    return (el.innerText || el.textContent || '').replace(/\s+/g, ' ').trim();
  }

  /* Dựng lưới có xử lý colspan / rowspan */
  function luoiBang(table) {
    var trs = [], i, j;
    var all = table.rows;
    for (i = 0; i < all.length; i++) if (hien(all[i])) trs.push(all[i]);

    var grid = [], merges = [];
    for (i = 0; i < trs.length; i++) {
      grid[i] = grid[i] || [];
      var c = 0;
      for (j = 0; j < trs[i].cells.length; j++) {
        while (grid[i][c] !== undefined) c++;
        var td = trs[i].cells[j];
        var cs = td.colSpan || 1, rs = td.rowSpan || 1;
        grid[i][c] = { text: chu(td), th: td.tagName === 'TH', raw: td };
        if (cs > 1 || rs > 1) merges.push({ r: i, c: c, cs: cs, rs: rs });
        for (var a = 0; a < rs; a++) {
          grid[i + a] = grid[i + a] || [];
          for (var b = 0; b < cs; b++) if (!(a === 0 && b === 0)) grid[i + a][c + b] = { text: '', th: false, spanned: true };
        }
        c += cs;
      }
    }
    for (i = 0; i < grid.length; i++) { grid[i] = grid[i] || []; for (j = 0; j < grid[i].length; j++) if (!grid[i][j]) grid[i][j] = { text: '', th: false }; }
    return { grid: grid, merges: merges, headRows: (table.tHead ? table.tHead.rows.length : 0) };
  }

  function laBangDuLieu(t) {
    if (t.hasAttribute('data-plg-skip')) return false;
    if (!hien(t)) return false;
    if (t.rows.length < 2) return false;
    var ths = t.querySelectorAll('th');
    var cells = t.rows[0] ? t.rows[0].cells.length : 0;
    return (ths.length >= 2 || cells >= 3);
  }

  var tieuDeMacDinh = '';
  function tenSheet(t, idx, dung) {
    var n = t.getAttribute('data-plg-sheet');
    if (!n) {
      var p = t.previousElementSibling, b = 0;
      while (p && b < 4) {
        var cls = String(p.className || '');
        var laTieuDe = /^H[1-6]$/.test(p.tagName)
          || /(^|[\s_-])(doc-title|section-title|sub-hdr|tieu-de|title|heading)([\s_-]|$)/i.test(cls);
        if (laTieuDe) {
          var txt = chu(p);
          /* bỏ qua thứ trông như "Nhãn: giá trị", ngày tháng, hoặc quá dài */
          if (txt && txt.indexOf(':') < 0 && !/\d{2}[\/-]\d{2}/.test(txt) && txt.length <= 40) { n = txt; break; }
        }
        p = p.previousElementSibling; b++;
      }
    }
    if (!n) n = (tieuDeMacDinh || '') || ('Bảng ' + (idx + 1));
    n = n.replace(/[\\\/\?\*\[\]:]/g, ' ').trim().slice(0, 28) || ('Bảng ' + (idx + 1));
    var g = n, k = 2;
    while (dung[g]) { g = n.slice(0, 26) + ' ' + k; k++; }
    dung[g] = 1;
    return g;
  }

  function thongTinPhieu(doc) {
    var ra = [], i, seen = {};
    function them(l, v) {
      l = String(l || '').replace(/[:：]\s*$/, '').trim();
      v = String(v || '').trim();
      if (!l || !v || v === '——' || v === '--' || seen[l]) return;
      seen[l] = 1; ra.push([l, v]);
    }
    var mark = doc.querySelectorAll('[data-plg-info]');
    for (i = 0; i < mark.length; i++) them(mark[i].getAttribute('data-plg-info'), chu(mark[i]));
    if (ra.length) return ra;

    /* cặp nhãn/giá trị tường minh */
    var items = doc.querySelectorAll('.info-item, .info-row, .meta-field, .meta-row, .thong-tin-item, .ir tr');
    for (i = 0; i < items.length && ra.length < 24; i++) {
      var el = items[i];
      if (!hien(el)) continue;
      var l = el.querySelector('.info-label, .meta-label, .label, th, td:first-child');
      var v = el.querySelector('.info-value, .meta-value, .value, td:last-child');
      if (l && v && l !== v) { them(chu(l), chu(v)); continue; }
      them.apply(null, tachHaiVe(chu(el)));
    }
    /* dòng meta kiểu "Chi nhánh: CN3" */
    var metas = doc.querySelectorAll('.meta span, .meta div, .header-info .meta > *');
    for (i = 0; i < metas.length && ra.length < 24; i++) {
      if (!hien(metas[i])) continue;
      var t = chu(metas[i]);
      if (t.indexOf(':') > 0) them.apply(null, tachHaiVe(t));
    }
    return ra;
  }

  function tachHaiVe(t) {
    var k = String(t || '').indexOf(':');
    return k > 0 ? [t.slice(0, k), t.slice(k + 1)] : ['', ''];
  }

  /* Khối tổng kết KPI (.sum-item / .sum-row / .sum-box) */
  function tongKet(doc) {
    var ra = [], i, seen = {};
    var items = doc.querySelectorAll('.sum-item, .sum-row, .sum-box');
    for (i = 0; i < items.length; i++) {
      var el = items[i];
      if (!hien(el)) continue;
      var l = el.querySelector('.sum-label, .label');
      var v = el.querySelector('.sum-value, .value, input');
      var lt = l ? chu(l) : '', vt = '';
      if (v) vt = (v.tagName === 'INPUT') ? (v.value || '') : chu(v);
      if (!l || !v) {
        var t = chu(el);
        if (t.indexOf(':') > 0) { var p2 = tachHaiVe(t); lt = p2[0]; vt = p2[1]; }
      }
      lt = String(lt).replace(/[:：]\s*$/, '').trim(); vt = String(vt).trim();
      if (!lt || !vt || vt === '——' || vt === '--' || seen[lt]) continue;
      seen[lt] = 1; ra.push([lt, vt]);
    }
    return ra;
  }

  /* =========================================== 7. DỰNG MÔ HÌNH WORKBOOK ==== */
  function kieuO(text, laTong, laCot1) {
    var so = parseSo(text);
    if (so !== null && !laCot1) {
      if (so.phanTram) return { v: so.v, t: 'n', s: 14 };
      return { v: so.v, t: 'n', s: laTong ? 8 : 5 };
    }
    if (so !== null && laCot1 && /^\d{1,3}$/.test(String(text).trim())) {
      return { v: so.v, t: 'n', s: 6 };                      // cột STT
    }
    return { v: text, t: 's', s: laTong ? 7 : 4 };
  }

  function dungSheet(name, tieuDe, info, bang, first, tk) {
    var rows = [], merges = [], i, j;
    var soCot = 1;
    for (i = 0; i < bang.grid.length; i++) if (bang.grid[i].length > soCot) soCot = bang.grid[i].length;
    if (soCot < 3) soCot = 3;

    function dong(cells, h) { rows.push({ cells: cells, h: h }); return rows.length; }
    function trai(txt, s, h) { var c = [{ v: txt, t: 's', s: s }]; return dong(c, h); }

    /* Khối đầu trang: công ty + tiêu đề */
    var r = trai(CHUAN.tenCongTy, 2, 20);
    merges.push('A' + r + ':' + colName(soCot) + r);
    r = trai(tieuDe, 1, 26);
    merges.push('A' + r + ':' + colName(soCot) + r);
    dong([], 6);

    /* Thông tin phiếu: 2 cặp nhãn/giá trị mỗi dòng */
    for (i = 0; i < info.length; i += 2) {
      var cells = [
        { v: info[i][0] + ':', t: 's', s: 9 },
        { v: info[i][1], t: 's', s: 10 }
      ];
      if (info[i + 1]) {
        while (cells.length < Math.max(2, Math.floor(soCot / 2))) cells.push({ v: '', s: 0 });
        cells.push({ v: info[i + 1][0] + ':', t: 's', s: 9 });
        cells.push({ v: info[i + 1][1], t: 's', s: 10 });
      }
      dong(cells, 18);
    }
    if (info.length) dong([], 6);

    /* Khối TỔNG KẾT (nếu trang có) */
    if (tk && tk.length) {
      trai('TỔNG KẾT', 11, 20);
      for (i = 0; i < tk.length; i++) {
        var so = parseSo(tk[i][1]);
        var oGiaTri = so !== null
          ? { v: so.phanTram ? so.v : so.v, t: 'n', s: so.phanTram ? 14 : 5 }
          : { v: tk[i][1], t: 's', s: 4 };
        var cTK = [{ v: tk[i][0], t: 's', s: 7 }, oGiaTri];
        dong(cTK, 18);
        merges.push('B' + rows.length + ':' + colName(Math.min(3, soCot)) + rows.length);
      }
      dong([], 6);
    }

    /* Bảng dữ liệu */
    var headRows = bang.headRows || 1;
    var startRow = rows.length + 1;
    for (i = 0; i < bang.grid.length; i++) {
      var g = bang.grid[i], cells = [], laHead = i < headRows;
      var dauDong = g[0] ? g[0].text : '';
      var laTong = !laHead && /tổng|cộng|total|thành tiền cuối/i.test(dauDong);
      for (j = 0; j < soCot; j++) {
        var cell = g[j];
        var txt = cell ? cell.text : '';
        if (laHead) cells.push({ v: txt, t: 's', s: 3 });
        else cells.push(kieuO(txt, laTong, j === 0));
      }
      dong(cells, laHead ? 26 : 0);
    }
    var endRow = rows.length;

    /* Merge trong bảng */
    for (i = 0; i < bang.merges.length; i++) {
      var m = bang.merges[i];
      if (m.cs <= 1 && m.rs <= 1) continue;
      var r1 = startRow + m.r, r2 = r1 + m.rs - 1;
      var c1 = m.c + 1, c2 = Math.min(m.c + m.cs, soCot);
      if (c2 > c1 || r2 > r1) merges.push(colName(c1) + r1 + ':' + colName(c2) + r2);
    }

    /* Ghi chú chân trang */
    dong([], 6);
    r = trai('Xuất từ hệ thống Phi Long Auto — ' + new Date().toLocaleString('vi-VN'), 12, 16);

    /* Độ rộng cột theo nội dung */
    var widths = [];
    for (j = 0; j < soCot; j++) {
      var max = CHUAN.rongCotMin;
      for (i = 0; i < rows.length; i++) {
        var c = rows[i].cells[j];
        if (!c || c.v === undefined || c.v === null) continue;
        var len = String(c.v).length;
        if (rows[i].cells.length === 1) continue;            // bỏ qua dòng tiêu đề merge
        if (len > max) max = len;
      }
      widths[j] = Math.min(CHUAN.rongCotMax, Math.max(CHUAN.rongCotMin, max * 1.15 + 2));
    }

    var headerRowIdx = startRow + headRows - 1;
    var soDongDL = endRow - headerRowIdx;

    return {
      name: name, rows: rows, merges: merges, widths: widths, first: !!first,
      freezeRow: headerRowIdx,
      autoFilter: soDongDL >= 3 ? ('A' + headerRowIdx + ':' + colName(soCot) + endRow) : null
    };
  }

  /* ====================================================== 8. XUẤT FILE ===== */
  function tenFileAnToan(s) {
    return String(s || 'BaoCao')
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/đ/g, 'd').replace(/Đ/g, 'D')
      .replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 60) || 'BaoCao';
  }

  function xuat(opt) {
    opt = opt || {};
    var doc = global.document;
    var cfg = global.PLG_EXCEL_OPTIONS || {};
    var tieuDe = opt.tenBaoCao || cfg.tenBaoCao
      || chu(doc.querySelector('h1, .title, .header-title, .doc-title') || {})
      || (doc.title || 'BÁO CÁO').split('-')[0].trim();
    tieuDe = tieuDe.toUpperCase();

    tieuDeMacDinh = tieuDe;
    var info = thongTinPhieu(doc);
    var tk = tongKet(doc);
    var tables = doc.querySelectorAll('table'), ds = [], i;
    for (i = 0; i < tables.length; i++) if (laBangDuLieu(tables[i])) ds.push(tables[i]);

    if (!ds.length) {
      alert('Không tìm thấy bảng dữ liệu nào trong phiếu này để xuất Excel.');
      return null;
    }

    var dung = {}, sheets = [];
    for (i = 0; i < ds.length; i++) {
      sheets.push(dungSheet(tenSheet(ds[i], i, dung), tieuDe, info, luoiBang(ds[i]), i === 0, i === 0 ? tk : null));
    }

    var bytes = buildXlsx(sheets);
    var d = new Date(), p2 = function (n) { return (n < 10 ? '0' : '') + n; };
    var stamp = d.getFullYear() + p2(d.getMonth() + 1) + p2(d.getDate()) + '_' + p2(d.getHours()) + p2(d.getMinutes());
    var ten = tenFileAnToan(opt.tenFile || cfg.tenFile || tieuDe) + '_' + stamp + '.xlsx';

    var blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var url = URL.createObjectURL(blob);
    var a = doc.createElement('a');
    a.href = url; a.download = ten; doc.body.appendChild(a); a.click();
    setTimeout(function () { doc.body.removeChild(a); URL.revokeObjectURL(url); }, 1500);
    return ten;
  }

  /* ====================================================== 9. NÚT BẤM ======= */
  function ganNut() {
    var doc = global.document;
    if (doc.getElementById('plg-excel-btn')) return;
    var css = doc.createElement('style');
    css.textContent =
      '#plg-excel-btn{position:fixed;right:20px;bottom:20px;z-index:99999;'
      + 'background:#' + CHUAN.mauThuongHieu + ';color:#fff;border:none;border-radius:8px;'
      + 'padding:12px 20px;font:600 15px/1.2 "Segoe UI",Arial,sans-serif;cursor:pointer;'
      + 'box-shadow:0 4px 14px rgba(0,0,0,.25);display:flex;align-items:center;gap:8px}'
      + '#plg-excel-btn:hover{background:#09406f}'
      + '#plg-excel-btn:active{transform:translateY(1px)}'
      + '@media print{#plg-excel-btn{display:none!important}}';
    doc.head.appendChild(css);

    var b = doc.createElement('button');
    b.id = 'plg-excel-btn';
    b.type = 'button';
    b.innerHTML = '<span style="font-size:17px">⬇</span> Xuất Excel';
    b.onclick = function () {
      b.disabled = true;
      var cu = b.innerHTML;
      b.innerHTML = 'Đang xuất...';
      setTimeout(function () {
        try { xuat(); }
        catch (e) { alert('Lỗi xuất Excel: ' + e.message); }
        b.innerHTML = cu; b.disabled = false;
      }, 30);
    };
    doc.body.appendChild(b);
  }

  var API = { xuat: xuat, tongKet: tongKet, thongTinPhieu: thongTinPhieu, buildXlsx: buildXlsx, dungSheet: dungSheet, parseSo: parseSo, CHUAN: CHUAN };
  global.PLGExcel = API;
  if (typeof module !== 'undefined' && module.exports) module.exports = API;

  if (typeof global.document !== 'undefined' && global.document.createElement) {
    if (global.document.readyState === 'loading')
      global.document.addEventListener('DOMContentLoaded', ganNut);
    else ganNut();
  }
})(typeof window !== 'undefined' ? window : globalThis);
