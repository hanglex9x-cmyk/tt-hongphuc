# Deploy Apps Script API — HỒNG PHÚC

Trạng thái: Code.gs đã viết xong + đã kiểm chứng logic offline trên dữ liệu Sheet thật.
Bước tạo project/deploy tự động qua API đang bị chặn quyền (xem mục "Điểm chặn" cuối file),
nên deploy thủ công theo cách A (5 phút) hoặc chờ cấp lại credential để chạy cách B.

- Sheet database: `QuanLy_TrungTamLaiXe_HongPhuc`
  ID: `1_YsxPojVSBm9rJ2u4x-maKBddtozyfflO2ff-wRpD9A`
  Link: https://docs.google.com/spreadsheets/d/1_YsxPojVSBm9rJ2u4x-maKBddtozyfflO2ff-wRpD9A/edit
- File nguồn: `Code.gs` (cùng thư mục này) — dòng đầu đã gán đúng SHEET_ID.
- File manifest: `appsscript.json` (Web App: Anyone with link, Execute as Me).

## Cách A — Deploy thủ công (khuyên dùng)

1. Mở Sheet ở trên → menu **Tiện ích mở rộng (Extensions) → Apps Script**.
   (Hoặc vào script.google.com → Dự án mới, Code.gs tự đọc Sheet theo SHEET_ID.)
2. Xóa nội dung mặc định, dán toàn bộ `Code.gs` vào.
3. (Tùy chọn) Bật hiển thị file manifest: Cài đặt dự án → "Show appsscript.json manifest file",
   rồi dán nội dung `appsscript.json` đè lên.
4. Chạy thử hàm `caiTriggerBackup` 1 lần (tạo trigger sao lưu Sheet lúc 2h sáng) và cấp quyền khi được hỏi.
5. **Deploy → New deployment → Web app**:
   - Execute as: **Me (hanglex9x@gmail.com)**
   - Who has access: **Anyone**
6. Copy URL `/exec` → dán vào biến `API_URL` ở đầu file `index.html` (frontend)
   hoặc mở frontend kèm `?api=<URL /exec>`.
7. Mỗi lần sửa Code.gs: **Deploy → Manage deployments → Edit → Version: New version → Deploy**
   (URL /exec giữ nguyên).

## Cách B — Qua API (khi credential Apps Script được cấp lại dạng OAuth)

```bash
python3 ~/workspace/custom-connectors/custom.google-apps-script/bin/gas.py POST /v1/projects --data '{"title":"QuanLy_TrungTamLaiXe_HongPhuc_API"}'
# lấy scriptId, sau đó PUT content gồm Code.gs + appsscript.json, tạo version rồi tạo deployment
```

## API

- GET: `<URL /exec>?action=<tên>&callback=<cb>` (JSONP) hoặc không callback (JSON).
- POST: `<URL /exec>` body JSON `{"action":"<tên>","params":{...}}`.
- Action đọc: getdashboard, getdatsummary, liststudents, getstudent, listschedules,
  listschedulesrange, listconflicts, listpayments, listcenters, getfleet, listfuellogs,
  listadmissions, listexamrounds, listexams, getreports, getgovreports,
  listnotifications, getlookups.
- Action ghi (đều ghi log vào tab LichSu): savestudent, setstudentstatus,
  transferstudentcourse, createschedule, updateschedule, deleteschedule, addpayment,
  saveteacher, savevehicle, savecenter, deletecenter, savebotuc, deletebotuc,
  addfuellog, deletefuellog, saveadmission, deleteadmission, setadmissiondoc,
  saveexamround, deleteexamround, addroundregistration, removeroundregistration,
  updateexamresult, updategraduation, updatelicensestatus, sendnotification,
  enrolladmission, checkcccd.

## Điểm chặn (ghi nhận 28/10/2026)

- `gas.py` (connector custom.google-apps-script): POST /v1/projects trả **HTTP 401
  UNAUTHENTICATED** — API này bắt buộc OAuth 2 access token của user, credential
  api_key đã đăng ký không dùng được.
- `hatch_gws_cli script ...`: wrapper chặn service `script`
  ("unsupported gws top-level command for hatch_gws_cli: script").
→ Chưa có Apps Script project ID và URL /exec. Không bịa URL; deploy theo Cách A ở trên.
