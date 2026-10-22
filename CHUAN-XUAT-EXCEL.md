# CHUẨN XUẤT FILE EXCEL — PHI LONG AUTO
_Chốt ngày 22/10/2026_

## 1. Định dạng chuẩn (áp cho TẤT CẢ biểu mẫu)

| Hạng mục | Chuẩn |
|---|---|
| Phông chữ | **Times New Roman** |
| Cỡ chữ | Nội dung **12** · Tiêu đề **16 đậm** · Ghi chú **11 nghiêng** |
| Màu thương hiệu | **#124A95** (lấy đúng từ logo PLG) |
| Tiêu đề bảng | Nền #124A95, chữ **trắng đậm**, canh giữa, tự xuống dòng |
| Dòng tổng cộng | Nền #DCE6F1, chữ đậm |
| Khung viền | Thin, màu đen, toàn bộ ô dữ liệu |
| Định dạng số | Tiền `#,##0` · Phần trăm `0.00%` · Ngày giữ dạng chữ |
| Cố định dòng | Freeze ngay dưới dòng tiêu đề bảng |
| Lọc tự động | Bật khi bảng có từ 3 dòng dữ liệu |
| Trang in | A4, fit 1 trang ngang, dọc/ngang tự chọn theo số cột |
| Tên file | `TenBaoCao_YYYYMMDD_HHMM.xlsx` |

## 2. Cấu trúc file xuất ra
1. Dòng 1 — Tên công ty
2. Dòng 2 — Tên báo cáo (merge, chữ xanh 16 đậm)
3. Khối **thông tin phiếu** (Chi nhánh, Ngày, Biển số, Số phiếu…)
4. Khối **TỔNG KẾT** (các chỉ số tổng của trang, nếu có)
5. **Bảng dữ liệu** — mỗi bảng trên trang là một sheet riêng
6. Dòng cuối — thời điểm xuất file

## 3. Cách dùng
Mở biểu mẫu → bấm nút **⬇ Xuất Excel** (góc dưới phải). Nút tự ẩn khi in.

## 4. Bảo trì
- Nguồn chuẩn duy nhất: `plg-excel.js`
- Sau khi sửa `plg-excel.js`, chạy: `python3 build-inline.py`
  → tự nhúng lại vào cả 12 biểu mẫu (phần nằm giữa `BEGIN PLG-EXCEL` và `END PLG-EXCEL`).
- **Không sửa tay** phần nhúng trong file HTML.

## 5. Logo
Toàn bộ biểu mẫu dùng `PLG-logo-print.png` (bản gốc Tracy cung cấp 22/10/2026,
1024×1024, giữ nguyên không tách nền). 4 file logo trong repo đều là bản này.
Thẻ ảnh ưu tiên file cạnh bên, tự dự phòng sang GitHub nếu thiếu.
