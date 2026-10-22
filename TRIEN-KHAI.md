# TRIỂN KHAI — ĐẨY LÊN GITHUB (5 phút)

Repo: **philonggroup/CARDIY5WAY** · Trang chạy thật: **philonggroup.github.io/CARDIY5WAY/**

## Bước 1 — Giải nén
Giải nén file zip ra 1 thư mục. Bên trong có **21 file** — đẩy hết, không cần chọn lọc.

## Bước 2 — Mở repo trên GitHub
1. Vào `github.com/philonggroup/CARDIY5WAY`
2. Đăng nhập tài khoản có quyền ghi repo
3. Bấm **Add file → Upload files**

## Bước 3 — Kéo thả
Chọn **TẤT CẢ 21 file** trong thư mục vừa giải nén → kéo thả vào khung upload.
GitHub sẽ tự ghi đè file trùng tên. Không cần xoá file cũ.

| Nhóm | Số file | Ghi chú |
|---|---|---|
| Biểu mẫu HTML | 12 | đã nhúng nút Xuất Excel |
| Logo PLG | 4 | bản gốc Tracy gửi 22/10/2026 |
| `plg-excel.js` | 1 | nguồn chuẩn duy nhất |
| `build-inline.py` | 1 | công cụ nhúng lại khi sửa chuẩn |
| `CHUAN-XUAT-EXCEL.md` | 1 | tài liệu chuẩn |
| `TRIEN-KHAI.md` + `README.md` | 2 | hướng dẫn |

## Bước 4 — Ghi chú commit
Ô **Commit changes**, gõ:

```
Chuan xuat Excel + doi logo PLG moi (22/10/2026)
```

Bấm **Commit changes**.

## Bước 5 — Chờ và kiểm tra
Đợi **2–10 phút** để GitHub Pages cập nhật, rồi mở:

```
https://philonggroup.github.io/CARDIY5WAY/philongauto.congnokh.html
```

Bấm **Ctrl + F5** (xoá cache trình duyệt) trước khi xem.

---

# CHECKLIST NGHIỆM THU

Mở lần lượt vài biểu mẫu trên GitHub Pages, kiểm 6 điểm:

- [ ] **1. Logo mới** — logo PLG hiện đúng ở góc trái đầu phiếu
- [ ] **2. Nút Xuất Excel** — nút xanh ở góc dưới phải màn hình
- [ ] **3. Bấm nút** — tải về file `.xlsx`, mở được bằng Excel không báo lỗi
- [ ] **4. Phông chữ** — mở file, bấm vào ô bất kỳ: phải là **Times New Roman 12**
- [ ] **5. Màu header** — dòng tiêu đề bảng nền **xanh**, chữ **trắng đậm**
- [ ] **6. Số là số thật** — bôi đen cột tiền, góc dưới Excel phải hiện **Sum**
      (nếu không hiện Sum nghĩa là đang lưu dạng chữ → báo lại ngay)

- [ ] **7. In thử** — bấm In báo cáo: nút Xuất Excel **phải tự ẩn**, layout in giữ nguyên như cũ

## Nếu có sự cố

| Hiện tượng | Nguyên nhân | Xử lý |
|---|---|---|
| Không thấy nút Xuất Excel | Trình duyệt còn cache | Ctrl + F5 |
| Logo vẫn là logo cũ | Cache ảnh | Ctrl + F5, hoặc đợi thêm 10 phút |
| Trang trắng / lỗi | Upload thiếu file | Đẩy lại đủ 21 file |
| File Excel mở báo lỗi | Tải chưa xong | Tải lại |

## Lùi về bản cũ (nếu cần)
Vào tab **Commits** của repo → chọn commit trước đó → **Revert**.
Toàn bộ thay đổi được gói trong 1 commit nên lùi về rất gọn.
