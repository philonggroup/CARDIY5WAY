# -*- coding: utf-8 -*-
"""Nhúng plg-excel.js vào từng biểu mẫu HTML.
Chạy lại file này mỗi khi sửa plg-excel.js:  python3 build-inline.py"""
import io, re, glob, sys

BEGIN = '<!-- ===== BEGIN PLG-EXCEL (tu dong tao - khong sua tay) ===== -->'
END   = '<!-- ===== END PLG-EXCEL ===== -->'
OLD_MARK = '<!-- ===== CHUẨN XUẤT EXCEL PHI LONG (plg-excel.js) ===== -->'

NAMES = {
 'baogia.html':                             ('PHIẾU BÁO GIÁ & NGHIỆM THU','PhieuBaoGia_NghiemThu'),
 'phieukiemtra.html':                       ('PHIẾU KIỂM TRA XE','PhieuKiemTraXe'),
 'philongauto.baocaothuchi.html':           ('BÁO CÁO THU CHI NGÀY','BaoCaoThuChiNgay'),
 'philongauto.congnokh.html':               ('BÁO CÁO CÔNG NỢ KHÁCH HÀNG','CongNoKhachHang'),
 'philongauto.congnoncc.html':              ('BÁO CÁO CÔNG NỢ NHÀ CUNG CẤP','CongNoNhaCungCap'),
 'philongauto.theodoicongnokhachhang.html': ('THEO DÕI CÔNG NỢ KHÁCH HÀNG','TheoDoiCongNoKH'),
 'philongauto_baogia.html':                 ('PHIẾU DỊCH VỤ KIÊM BẢO HÀNH','PhieuDichVu_BaoHanh'),
 'philongauto_baogialop_gam.html':          ('BÁO GIÁ LỐP & GẦM','BaoGiaLop_Gam'),
 'philongautobaogialop.html':               ('PHIẾU DỊCH VỤ LỐP','PhieuDichVuLop'),
 'philongautohoadonthanhtoan.html':         ('HÓA ĐƠN THANH TOÁN','HoaDonThanhToan'),
 'philongautolenhsuachua.html':             ('LỆNH SỬA CHỮA','LenhSuaChua'),
 'prinbillthanhtoan.html':                  ('HÓA ĐƠN THANH TOÁN','HoaDonThanhToan_Bill'),
}

js = io.open('plg-excel.js', encoding='utf-8').read()
# thoát chuỗi </script trong chú thích để không làm vỡ thẻ <script> bao ngoài
js = re.sub(r'</(script)', r'<\\/\1', js, flags=re.I)

for f,(ten,tenfile) in sorted(NAMES.items()):
    s = io.open(f, encoding='utf-8').read()
    # gỡ bản cũ (nếu có) để chạy lại được nhiều lần
    s = re.sub(re.escape(BEGIN)+r'.*?'+re.escape(END), '', s, flags=re.S)
    s = re.sub(re.escape(OLD_MARK)+r'.*?<script src="plg-excel\.js"></script>\s*', '', s, flags=re.S)
    block = (BEGIN + '\n'
             + '<script>window.PLG_EXCEL_OPTIONS={tenBaoCao:%r,tenFile:%r};</script>\n' % (ten, tenfile)
             + '<script>\n' + js + '\n</script>\n' + END + '\n')
    block = block.replace("u'", "'")   # repr python2-style an toàn
    assert '</body>' in s, f
    s = s.replace('</body>', block + '</body>', 1)
    io.open(f,'w',encoding='utf-8').write(s)
    print('nhúng OK:', f)
