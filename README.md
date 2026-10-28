# Quản lý Trung tâm Đào tạo Lái xe HỒNG PHÚC

Ứng dụng quản lý cho **TRUNG TÂM ĐÀO TẠO LÁI XE HỒNG PHÚC**
404 QL1A, P. Thới An, Q12, TP.HCM · ☎ 0949 155 955

- 🌐 **GitHub Pages (địa chỉ đích sau khi đẩy repo):** https://hanglex9x-cmyk.github.io/tt-hongphuc/
  · Repo: `hanglex9x-cmyk/tt-hongphuc` — xem mục "Trạng thái GitHub" bên dưới.
- 📊 **Google Sheet (database):** [QuanLy_TrungTamLaiXe_HongPhuc](https://docs.google.com/spreadsheets/d/1_YsxPojVSBm9rJ2u4x-maKBddtozyfflO2ff-wRpD9A/edit)
  · Sheet ID: `1_YsxPojVSBm9rJ2u4x-maKBddtozyfflO2ff-wRpD9A`
- ⚙️ **Apps Script Web App (API):** xem mục "Trạng thái API" bên dưới.
- 💻 **Frontend:** `index.html` trong repo này (1 file duy nhất, gọi API ở trên).

## Kiến trúc

```
GitHub Pages (index.html)  →  Apps Script Web App (/exec)  →  Google Sheet (17 tab)
```

## Trạng thái API (cập nhật 28/10/2026)

- `apps-script/Code.gs` đã viết xong toàn bộ API (đọc + ghi + ghi log LichSu +
  sao lưu tự động 2h sáng) và đã kiểm chứng offline trên dữ liệu Sheet thật:
  **42/42 số liệu chốt khớp bản gốc** (165 HV, 2.640 buổi, DAT 70.981 km,
  đã thu 2.940,1tr, nợ 403,9tr, đậu 34, hoa hồng TT 18,45tr…).
- Bước tạo project + deploy Web App qua API đang **bị chặn quyền OAuth**
  (Apps Script API trả 401; CLI bị chặn service `script`) nên chưa có URL `/exec`.
- Deploy thủ công 5 phút theo `apps-script/DEPLOY.md` (Cách A), sau đó dán URL
  `/exec` vào trang web: mở trang → mục **Cảnh báo & Zalo → Cấu hình API**, hoặc
  mở kèm `?api=<URL /exec>`. URL được lưu vào trình duyệt, các lần sau không cần nhập lại.

## Cấu trúc database (17 tab trong Sheet)

| Tab | Nội dung |
|---|---|
| KhoaHoc | 6 khóa (K38, K39, K40, K41, C1-K120, BK220): hạng, học phí, ngày khai giảng/bế giảng/bắt đầu DAT, sĩ số |
| HocVien | 165 học viên: thông tin cá nhân, khóa, GV chính, xe quen, khám SK, cabin, DAT, giảm tiền, trạng thái, tốt nghiệp, bằng lái |
| LichHoc | 2.421 dòng lịch: ngày, khung giờ, loại buổi (DAT/Sa hình/Lý thuyết/Cabin/Thi TN), HV, GV, xe, km kế hoạch/thực tế, điểm danh |
| ThanhToan | 251 phiếu thu: học phí, lệ phí thi, phí thi lại, hoàn phí |
| GiamTien | Các khoản giảm học phí theo học viên |
| GiaoVien | 13 giáo viên: tên ngắn, màu nhãn, loại, đơn giá giờ, giờ/tháng |
| Xe | 14 xe: biển số, hạng, GV chủ xe, khóa phụ trách, km đồng hồ, bảo dưỡng, hạn đăng kiểm/bảo hiểm/phù hiệu |
| TrungTamLienKet | 5 trung tâm: tỷ lệ & tiền hoa hồng |
| BoTucThueXe | Dịch vụ bổ túc tay lái & thuê xe (thành tiền = giờ × đơn giá) |
| TuyenSinh | 14 hồ sơ tuyển sinh + tình trạng từng mục giấy tờ |
| DotSathach | Các đợt thi sát hạch |
| DangKyThi | Học viên đăng ký theo đợt + số báo danh |
| KetQuaThi | Kết quả Lý thuyết / Sa hình / Đường trường, phí thi lại, cấp bằng |
| TotNghiep | Kết quả thi tốt nghiệp + số/ngày cấp chứng chỉ |
| NhatKyNhienLieu | Nhật ký đổ xăng theo xe (tính định mức lít/100km) |
| CauHinh | Tên trung tâm, học phí theo hạng, mục tiêu DAT, phí thi lại, mốc "ngày hiện tại", mẫu thông báo, số liệu lãi/lỗ |
| LichSu | Nhật ký mọi thao tác ghi (thời gian, hành động, đối tượng, người thực hiện, chi tiết) |

## Cách deploy GitHub Pages

Thư mục này đã là 1 git repo (đã commit sẵn). Khi GitHub được kết nối lại, đẩy lên bằng:

```bash
cd ~/workspace/hongphuc-github
git remote add origin https://github.com/hanglex9x-cmyk/tt-hongphuc.git
git push -u origin main
```

Sau đó vào repo → **Settings → Pages** → Source: branch `main`, thư mục `/ (root)` → Save.
Pages tự cập nhật sau mỗi lần push lên `main` (~1 phút).

## Trạng thái GitHub (cập nhật 28/10/2026)

Chưa tạo được repo `tt-hongphuc` và chưa bật Pages vì mọi đường truy cập GitHub
từ môi trường này đều bị chặn ở bước xác thực: connector GitHub trả 401
(Bad credentials), GitHub MCP chưa kết nối OAuth, `gh` chưa đăng nhập,
SSH tới github.com bị reset. Chỉ cần kết nối lại GitHub (hoặc chạy 2 lệnh push
ở trên từ máy đã đăng nhập) là trang lên ngay tại địa chỉ đích ở đầu file.

## Quy tắc nghiệp vụ chính

- Sát hạch theo Thông tư 108/2026: thứ tự Lý thuyết → Sa hình → Đường trường (bỏ phần thi mô phỏng).
- Đủ điều kiện thi khi: đạt Lý thuyết + Cabin đủ 3h + DAT đủ km (B1 710, B2 810, C1/C 825) + đã học Sa hình + đóng đủ học phí.
- Học phí: B1 18tr · B2 20tr · C1 25tr · C 24tr · Nâng dấu 13tr · Bổ túc 350k/giờ · Thuê xe 300k/giờ.
- Sao lưu Sheet tự động 2h sáng mỗi ngày (hàm `autoBackup` + trigger trong Apps Script, giữ 14 bản).
