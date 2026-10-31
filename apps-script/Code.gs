// ============================================================
// QUẢN LÝ TRUNG TÂM ĐÀO TẠO LÁI XE HỒNG PHÚC — Apps Script API
// Database: Google Sheet QuanLy_TrungTamLaiXe_HongPhuc
// Web App: Execute as Me (hanglex9x@gmail.com), Anyone with link
// ============================================================
// BẢN TỐI ƯU TỐC ĐỘ (28/10/2026) — hợp đồng API giữ nguyên 100%:
//  1) ss_() chỉ openById 1 lần / lần chạy (trước đây: 1 lần / tab).
//  2) Các tập dữ liệu dẫn xuất (allStudents_, allSchedules_, conflicts_...)
//     được memoize theo request: dựng 1 lần, dùng chung; ghi xong tự xóa đệm.
//  3) [v3] CacheService cho MỌI action đọc (kể cả có tham số): khóa
//     hp2:<gen>:<action>:<tham số đã sắp xếp>, TTL 1800s. Mọi action ghi
//     thành công tăng <gen> lên 1 ở route_() -> khóa thế hệ cũ không còn ai
//     đọc (tự hết hạn). Trigger warmCacheHp mỗi 10 phút giữ cache luôn nóng.
// ============================================================
const SHEET_ID = '1_YsxPojVSBm9rJ2u4x-maKBddtozyfflO2ff-wRpD9A';

const TARGET_KM = { 'B1': 710, 'B2': 810, 'C1': 825, 'C': 825 };
const TUITION = { 'B1': 18000000, 'B2': 20000000, 'C1': 25000000, 'C': 24000000, 'Nâng dấu': 13000000 };

// ---------- Hạ tầng ----------
let _ssObj = null;
function ss_() { if (!_ssObj) _ssObj = SpreadsheetApp.openById(SHEET_ID); return _ssObj; }

const _cache = {};
function tab_(name) {
  if (_cache[name]) return _cache[name];
  const sh = ss_().getSheetByName(name);
  if (!sh) return { headers: [], rows: [], objs: [], sheet: null };
  const vals = sh.getDataRange().getValues();
  const headers = vals[0].map(String);
  const objs = [];
  for (let i = 1; i < vals.length; i++) {
    const o = { _row: i + 1 };
    headers.forEach((h, j) => { o[h] = vals[i][j]; });
    objs.push(o);
  }
  const t = { headers, rows: vals, objs, sheet: sh };
  _cache[name] = t;
  return t;
}
// Đệm dữ liệu dẫn xuất theo request: mỗi tập dữ liệu chỉ dựng 1 lần / lần chạy.
// Mọi thao tác ghi đều đi qua dropTab_() -> invalidateDerived_() để xóa đệm,
// nên trong cùng 1 request, dữ liệu đọc sau khi ghi luôn là dữ liệu mới.
const _memo = {};
function memo_(key, fn) { if (!(key in _memo)) _memo[key] = fn(); return _memo[key]; }
function invalidateDerived_() { for (const k in _memo) delete _memo[k]; }
function dropTab_(name) { delete _cache[name]; invalidateDerived_(); }
function cfg_(key, def) {
  return memo_('cfg:' + key, () => {
    const t = tab_('CauHinh');
    for (const o of t.objs) if (String(o['Khoa']) === key) return o['GiaTri'];
    return def;
  });
}
function today_() { return String(cfg_('NgayHienTai', '2026-10-27')).slice(0, 10); }
function num_(v) { const n = Number(v); return isNaN(n) ? 0 : n; }
function str_(v) { return v === null || v === undefined ? '' : String(v); }
function dstr_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
  if (typeof v === 'number' && v > 20000 && v < 80000) {
    // So serial ngay Excel/Sheet (serial 25569 = 01/01/1970) — o ngay bi luu dang so.
    return Utilities.formatDate(new Date(Math.round((v - 25569) * 86400000)), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[1]).padStart(2, '0');
  return s.slice(0, 10);
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function logHistory_(action, code, detail) {
  try {
    const sh = ss_().getSheetByName('LichSu');
    sh.appendRow([new Date().toISOString(), action, code || '', 'Admin', detail || '']);
  } catch (e) { /* không chặn luồng chính */ }
}
function findRow_(tabName, id) {
  const t = tab_(tabName);
  for (const o of t.objs) if (num_(o['ID']) === num_(id)) return o;
  return null;
}
function nextId_(tabName) {
  const t = tab_(tabName);
  let mx = 0;
  for (const o of t.objs) mx = Math.max(mx, num_(o['ID']));
  return mx + 1;
}
function writeRow_(tabName, id, fields) {
  const t = tab_(tabName);
  const o = findRow_(tabName, id);
  if (!o) return false;
  const headers = t.headers;
  const vals = headers.map(h => (fields[h] !== undefined ? fields[h] : o[h]));
  t.sheet.getRange(o._row, 1, 1, headers.length).setValues([vals]);
  dropTab_(tabName);
  return true;
}
function appendObj_(tabName, fields) {
  const t = tab_(tabName);
  const vals = t.headers.map(h => (fields[h] !== undefined ? fields[h] : ''));
  t.sheet.appendRow(vals);
  dropTab_(tabName);
  return true;
}
function deleteRow_(tabName, id) {
  const t = tab_(tabName);
  const o = findRow_(tabName, id);
  if (!o) return false;
  t.sheet.deleteRow(o._row);
  dropTab_(tabName);
  return true;
}

// ---------- Bổ trợ nghiệp vụ ----------
function courses_() {
  return memo_('courses', () => tab_('KhoaHoc').objs.map(o => ({
    code: str_(o['MaKhoa']), licenseClass: str_(o['Hang']), tuition: num_(o['HocPhi']),
    openDate: dstr_(o['NgayKhaiGiang']), closeDate: dstr_(o['NgayBeGiang']),
    datStart: dstr_(o['NgayBatDauDAT']), size: num_(o['SiSo'])
  })));
}
function courseMap_() {
  return memo_('courseMap', () => { const m = {}; courses_().forEach(c => m[c.code] = c); return m; });
}

function paymentsByStudent_() {
  return memo_('payMap', () => {
    const m = {};
    for (const p of tab_('ThanhToan').objs) {
      const sid = num_(p['IDHocVien']);
      if (!m[sid]) m[sid] = { paid: 0, rows: [] };
      const loai = str_(p['Loai']);
      if (loai === 'Học phí' || loai === 'Hoàn phí') m[sid].paid += num_(p['SoTien']);
      m[sid].rows.push(p);
    }
    return m;
  });
}
function datKmByStudent_() {
  return memo_('kmMap', () => {
    const T = today_();
    const m = {};
    for (const s of tab_('LichHoc').objs) {
      if (str_(s['LoaiBuoi']) !== 'DAT' || !s['IDHocVien']) continue;
      if (dstr_(s['Ngay']) > T) continue;
      if (str_(s['DiemDanh']) !== 'Đã học') continue;
      const km = s['KmThucTe'] !== '' && s['KmThucTe'] !== null && s['KmThucTe'] !== undefined ? num_(s['KmThucTe']) : num_(s['KmKeHoach']);
      const sid = num_(s['IDHocVien']);
      m[sid] = (m[sid] || 0) + km;
    }
    return m;
  });
}
function studentObj_(o, payMap, kmMap) {
  const cm = courseMap_();
  const course = cm[str_(o['MaKhoa'])] || {};
  const sid = num_(o['ID']);
  const paid = payMap[sid] ? payMap[sid].paid : 0;
  const ownFee = num_(o['HocPhi']);
  const total = num_(o['ID']) ? (ownFee > 0 ? ownFee : (course.tuition || 0)) : 0;
  const discount = num_(o['GiamTien']);
  const payable = total - discount;
  const cls = str_(o['Hang']);
  const kmDone = kmMap[sid] !== undefined ? kmMap[sid] : num_(o['KmDATDaChay']);
  const target = TARGET_KM[cls] || 810;
  const isMLL = num_(o['MuaLL']) === 1;
  const crit = {
    theory: str_(o['DatLyThuyet']) === 'Đạt',
    cabin: num_(o['GioCabin']) >= 3,
    dat: kmDone >= target,
    sahinh: str_(o['DaHocSaHinh']) === 'Đã học',
    tuition: isMLL ? true : (payable - paid) <= 0
  };
  return {
    id: sid, code: str_(o['MaHV']), maHv: str_(o['MaHVSach']), name: str_(o['HoTen']),
    dob: dstr_(o['NgaySinh']), cccd: str_(o['CCCD']), phone: str_(o['SDT']), address: str_(o['DiaChi']),
    courseCode: str_(o['MaKhoa']), licenseClass: cls, centerName: str_(o['TrungTamLienKet']),
    homeTeacherName: str_(o['GVChinh']), homePlate: str_(o['XeQuen']), plateKg: str_(o['BienSoKG']),
    regDate: dstr_(o['NgayDangKy']), healthExamDate: dstr_(o['NgayKhamSK']), healthExpiry: dstr_(o['HanKhamSK']),
    cabinHours: num_(o['GioCabin']), cabinSessions: num_(o['BuoiCabin']),
    theoryScore: o['DiemLyThuyet'] === '' ? null : num_(o['DiemLyThuyet']), theoryPassed: crit.theory,
    sahinhDone: crit.sahinh, datKmDone: kmDone, datKmTotal: target, datMarkerKm: num_(o['MocKmDAT']),
    datStartDate: course.datStart || '', datDeadline: course.closeDate || '',
    openDate: course.openDate || '', closeDate: course.closeDate || '',
    tuitionTotal: total, tuitionDiscount: discount, discountReason: str_(o['LyDoGiam']),
    tuitionPayable: payable, tuitionPaid: paid, tuitionDebt: Math.max(0, payable - paid),
    muaLL: isMLL,
    status: str_(o['TrangThai']), gradExamDate: dstr_(o['NgayThiTotNghiep']), gradResult: str_(o['KetQuaTotNghiep']),
    certNo: str_(o['SoChungChi']), certDate: dstr_(o['NgayCapChungChi']),
    licenseDate: dstr_(o['NgayCoBang']), hasLicense: str_(o['CoBang']) === 'Có',
    licenseScore: o['DiemBangLai'] === '' ? null : num_(o['DiemBangLai']),
    criteria: crit, eligible: crit.theory && crit.cabin && crit.dat && crit.sahinh && crit.tuition
  };
}
function allStudents_() {
  return memo_('students', () => {
    const payMap = paymentsByStudent_();
    const kmMap = datKmByStudent_();
    return tab_('HocVien').objs.map(o => studentObj_(o, payMap, kmMap));
  });
}
function teacherLook_() {
  // Map ID giáo viên -> màu/tên ngắn, dựng 1 lần (giữ ngữ nghĩa cũ: trùng ID thì bản ghi sau thắng).
  return memo_('teacherLook', () => {
    const m = {};
    for (const t of tab_('GiaoVien').objs) m[num_(t['ID'])] = { color: str_(t['Mau']), short: str_(t['TenNgan']) };
    return m;
  });
}
function scheduleObj_(s) {
  const tl = teacherLook_()[num_(s['IDGiaoVien'])];
  const color = tl ? tl.color : '#CFD8DC', short = tl ? tl.short : '';
  const sid = s['IDHocVien'] === '' || s['IDHocVien'] === null ? null : num_(s['IDHocVien']);
  return {
    id: num_(s['ID']), date: dstr_(s['Ngay']), slot: str_(s['KhungGio']), startMinutes: num_(s['PhutBatDau']),
    type: str_(s['LoaiBuoi']), courseCode: str_(s['MaKhoa']), courseSize: 0,
    studentId: sid, studentName: str_(s['HoTenHV']),
    teacherId: s['IDGiaoVien'] === '' || s['IDGiaoVien'] === null ? null : num_(s['IDGiaoVien']),
    teacherName: str_(s['TenGiaoVien']), teacherShort: short, teacherColor: color,
    vehicleId: s['IDXe'] === '' || s['IDXe'] === null ? null : num_(s['IDXe']), plate: str_(s['BienSo']),
    centerName: str_(s['TrungTamLienKet']), km: num_(s['KmKeHoach']),
    actualKm: s['KmThucTe'] === '' || s['KmThucTe'] === null || s['KmThucTe'] === undefined ? null : num_(s['KmThucTe']),
    attendance: str_(s['DiemDanh']), groupLabel: str_(s['NhanNhom']), note: str_(s['GhiChu']),
    editable: !!sid
  };
}
function allSchedules_() {
  return memo_('schedules', () => {
    const cs = courseMap_();
    return tab_('LichHoc').objs.map(o => {
      const s = scheduleObj_(o);
      s.courseSize = cs[s.courseCode] ? cs[s.courseCode].size : 0;
      return s;
    });
  });
}
function conflicts_() {
  return memo_('conflicts', () => {
  const rows = allSchedules_().filter(s => s.studentId);
  const out = [];
  // Giờ học tự do: trùng = hai buổi cùng ngày, cùng xe/HV/GV và khoảng giờ CHỒNG LẤN
  // (không còn so khớp chuỗi khung giờ như hồi khung cố định).
  function overlaps_(a, b) {
    const ra = slotRangeMin_(a.slot), rb = slotRangeMin_(b.slot);
    if (!ra || !rb) return a.slot === b.slot;
    return ra[0] < rb[1] && rb[0] < ra[1];
  }
  function pairCheck_(list, keyFn, kind, severity) {
    for (let i = 0; i < list.length; i++) for (let j = i + 1; j < list.length; j++) {
      const a = list[i], b = list[j];
      const ka = keyFn(a), kb = keyFn(b);
      if (!ka || ka !== kb || !overlaps_(a, b)) continue;
      out.push({ kind, severity, key: a.date + '|' + ka, date: a.date, slot: a.slot, sessions: [a, b] });
    }
  }
  const byDate = {};
  rows.forEach(s => { (byDate[s.date] = byDate[s.date] || []).push(s); });
  Object.keys(byDate).forEach(d => {
    const list = byDate[d];
    pairCheck_(list, s => s.vehicleId ? 'xe' + s.vehicleId : null, 'Trùng xe', 'error');
    pairCheck_(list, s => 'hv' + s.studentId, 'Trùng học viên', 'error');
    pairCheck_(list, s => s.teacherId ? 'gv' + s.teacherId : null, 'Trùng giáo viên', 'warning');
  });
  return out;
  });
}
function slotHours_(slot) {
  const r = slotRangeMin_(slot);
  return r ? Math.max(1, (r[1] - r[0]) / 60) : 2;
}
function weekHoursByVehicle_() {
  return memo_('weekHours', () => {
  const T = today_();
  const d = new Date(T + 'T00:00:00');
  const dow = (d.getDay() + 6) % 7; // Mon=0
  const mon = new Date(d); mon.setDate(d.getDate() - dow);
  const sun = new Date(mon); sun.setDate(mon.getDate() + 6);
  const f = x => Utilities.formatDate(x, 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
  const seen = {}, hrs = {};
  for (const s of allSchedules_()) {
    if (!s.vehicleId || s.type === 'Lý thuyết') continue;
    if (s.date < f(mon) || s.date > f(sun)) continue;
    const k = s.date + '|' + s.slot + '|' + s.vehicleId;
    if (seen[k]) continue;
    seen[k] = 1;
    hrs[s.plate] = (hrs[s.plate] || 0) + slotHours_(s.slot);
  }
  return hrs;
  });
}

// ---------- API: Đọc ----------
function apiGetDashboard() {
  const T = today_();
  const students = allStudents_();
  const rows = allSchedules_();
  const byClass = {};
  students.forEach(s => { byClass[s.licenseClass] = (byClass[s.licenseClass] || 0) + 1; });
  const indiv = rows.filter(s => s.studentId);
  const le = indiv.filter(s => s.date <= T);
  const cnt = (arr, t) => arr.filter(s => s.type === t).length;
  const ltWeighted = rows.filter(s => s.type === 'Lý thuyết').reduce((a, s) => a + (s.courseSize || 0), 0);
  const scheduled = cnt(le, 'DAT') + cnt(le, 'Sa hình') + cnt(le, 'Thi tốt nghiệp') + ltWeighted;
  const total = students.length * 16;
  const days = new Set(indiv.map(s => s.date)).size;
  const heat = {};
  indiv.forEach(s => { heat[s.date] = (heat[s.date] || 0) + 1; });
  rows.filter(s => !s.studentId && s.type === 'Lý thuyết').forEach(s => { heat[s.date] = (heat[s.date] || 0) + (s.courseSize || 0); });
  const exams = tab_('KetQuaThi').objs;
  const passed = exams.filter(e => str_(e['KetQua']) === 'Đậu');
  const examined = exams.filter(e => str_(e['KetQua']) === 'Đậu' || str_(e['KetQua']) === 'Rớt');
  const licensed = students.filter(s => s.hasLicense);
  const mllIds = {};
  students.filter(s => s.muaLL).forEach(s => { mllIds[s.id] = 1; });
  let revMonth = 0;
  for (const p of tab_('ThanhToan').objs) {
    if (str_(p['Loai']) === 'Học phí' && !mllIds[num_(p['IDHocVien'])] && dstr_(p['Ngay']).slice(0, 7) === T.slice(0, 7)) revMonth += num_(p['SoTien']);
  }
  const pl = JSON.parse(cfg_('PLSeries', '[]'));
  const cur = pl.length ? pl[pl.length - 1] : { profit: 0, revenue: 0 };
  const conflicts = conflicts_();
  const todayRows = rows.filter(s => s.date === T);
  const todaySessions = todayRows.reduce((a, s) => a + (s.studentId ? 1 : (s.type === 'Lý thuyết' ? (s.courseSize || 0) : 5)), 0);
  return {
    kpis: {
      totalStudents: students.length,
      byClass: Object.keys(byClass).map(k => ({ licenseClass: k, count: byClass[k] })),
      todaySessions, thToday: todayRows.filter(s => s.studentId).length,
      ltToday: todayRows.filter(s => s.type === 'Lý thuyết').length,
      revenueMonth: revMonth, profitMonth: num_(cur.profit) * 1000000,
      profitPct: cur.revenue ? Math.round(cur.profit / cur.revenue * 1000) / 10 : 0,
      passRate: examined.length ? Math.round(passed.length / examined.length * 100) : 0,
      licensedCount: licensed.length,
      teachersActive: tab_('GiaoVien').objs.filter(t => str_(t['Loai']) === 'chinh').length,
      vehiclesTotal: tab_('Xe').objs.length,
      vehiclesInUse: tab_('Xe').objs.filter(v => str_(v['TrangThai']) === 'Đang sử dụng').length,
      conflicts: conflicts.length
    },
    sessionsProgress: {
      total, scheduled, remaining: total - scheduled, gridCells: total,
      datSpecial: cnt(le, 'DAT'), sahinh: cnt(le, 'Sa hình') + cnt(le, 'Thi tốt nghiệp'), regular: ltWeighted,
      off: 0, days, avgPerDay: days ? Math.round(scheduled / days * 10) / 10 : 0,
      from: '2026-06-07', to: '2026-12-31'
    },
    cabinStats: {
      fullCount: students.filter(s => s.cabinHours >= 3).length,
      total: students.length,
      pct: students.length ? Math.round(students.filter(s => s.cabinHours >= 3).length / students.length * 100) : 0
    },
    gantt: courses_().map(c => ({
      code: c.code, licenseClass: c.licenseClass, openDate: c.openDate, closeDate: c.closeDate,
      ltStart: c.openDate, ltEnd: '', thStart: c.datStart, thEnd: c.closeDate,
      studentCount: students.filter(s => s.courseCode === c.code).length
    })),
    heatmap: Object.keys(heat).sort().map(d => ({ date: d, sessions: heat[d] })),
    licensed: licensed.map(s => ({ name: s.name, licenseClass: s.licenseClass, licenseDate: s.licenseDate, licenseScore: s.licenseScore })),
    alerts: apiNotifications(),
    muaLLStudents: students.filter(s => s.muaLL).map(s => ({ name: s.name, code: s.code, courseCode: s.courseCode, licenseClass: s.licenseClass, homeTeacherName: s.homeTeacherName }))
  };
}
function apiGetDatSummary() {
  const students = allStudents_();
  const T = today_();
  // Đếm số buổi DAT đã học theo học viên trong 1 lượt quét LichHoc
  // (trước đây: lọc lại toàn bộ 2.640 dòng cho từng học viên -> 165 x 2.640 vòng lặp).
  const datDoneCount = {};
  for (const x of tab_('LichHoc').objs) {
    if (str_(x['LoaiBuoi']) !== 'DAT' || str_(x['DiemDanh']) !== 'Đã học') continue;
    if (dstr_(x['Ngay']) > T) continue;
    const sid = num_(x['IDHocVien']);
    datDoneCount[sid] = (datDoneCount[sid] || 0) + 1;
  }
  const rows = students.map(s => {
    const course = courseMap_()[s.courseCode] || {};
    const phaseEnd = new Date(course.closeDate); phaseEnd.setDate(phaseEnd.getDate() - 30);
    const phaseMs = (phaseEnd - new Date(course.datStart)) || 1;
    const elMs = Math.min(Math.max(new Date(T) - new Date(course.datStart), 0), Math.max(phaseMs, 0));
    const g2Pct = phaseMs <= 0 ? 100 : Math.round(elMs / phaseMs * 1000) / 10;
    const pct = s.datKmTotal ? Math.round(s.datKmDone / s.datKmTotal * 1000) / 10 : 0;
    let status;
    if (s.datKmDone >= s.datKmTotal) status = 'Đạt chuẩn';
    else if (s.datKmDone <= 0) status = 'Chưa chạy';
    else if (pct + 10 >= g2Pct) status = 'Đúng tiến độ';
    else status = 'Chậm tiến độ';
    const hoursDone = (datDoneCount[s.id] || 0) * 2;
    const avg = hoursDone ? Math.round(s.datKmDone / hoursDone * 10) / 10 : 40;
    const rem = Math.max(0, s.datKmTotal - s.datKmDone);
    return {
      studentId: s.id, code: s.code, maHv: s.maHv, name: s.name, courseCode: s.courseCode,
      licenseClass: s.licenseClass, openDate: s.openDate, closeDate: s.closeDate,
      doneKm: s.datKmDone, targetKm: s.datKmTotal, remainingKm: rem, pct,
      g2Pct, status, below500: s.datKmDone < 500 && s.datKmDone < s.datKmTotal,
      hoursDone, hoursNeeded: rem ? Math.round(rem / (avg || 40) * 2) / 2 : 0,
      avgSpeedKmh: avg, homePlate: s.homePlate, homeTeacherName: s.homeTeacherName
    };
  });
  const sum = { totalStudents: rows.length, totalKmDone: rows.reduce((a, r) => a + r.doneKm, 0),
    totalTargetKm: rows.reduce((a, r) => a + r.targetKm, 0),
    passedCount: rows.filter(r => r.status === 'Đạt chuẩn').length,
    onTrackCount: rows.filter(r => r.status === 'Đúng tiến độ').length,
    slowCount: rows.filter(r => r.status === 'Chậm tiến độ').length,
    notStartedCount: rows.filter(r => r.status === 'Chưa chạy').length };
  return { rows, summary: sum };
}
function normSearch_(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd');
}
function apiListStudents(p) {
  let rows = allStudents_();
  if (p.search) { const q = normSearch_(p.search); rows = rows.filter(s => normSearch_(s.name).includes(q) || normSearch_(s.code).includes(q) || normSearch_(s.maHv).includes(q)); }
  if (p.courseCode) rows = rows.filter(s => s.courseCode === p.courseCode);
  if (p.status) rows = rows.filter(s => s.status === p.status);
  if (p.centerName) rows = rows.filter(s => s.centerName === p.centerName);
  if (p.onlyEligible) rows = rows.filter(s => s.eligible);
  const total = rows.length;
  const page = num_(p.page) || 1, size = num_(p.pageSize) || 20;
  return { students: rows.slice((page - 1) * size, page * size), total, page, pageSize: size };
}
function apiGetStudent(p) {
  const s = allStudents_().find(x => x.id === num_(p.id));
  if (!s) return { ok: false, error: 'Không tìm thấy học viên' };
  const pays = tab_('ThanhToan').objs.filter(x => num_(x['IDHocVien']) === s.id)
    .map(x => ({ id: num_(x['ID']), date: dstr_(x['Ngay']), amount: num_(x['SoTien']), kind: str_(x['Loai']), method: str_(x['PhuongThuc']), note: str_(x['GhiChu']) }));
  const scheds = allSchedules_().filter(x => x.studentId === s.id);
  const exams = tab_('KetQuaThi').objs.filter(x => num_(x['IDHocVien']) === s.id);
  const hist = tab_('LichSu').objs.filter(x => str_(x['MaDoiTuong']) === s.code)
    .map(x => ({ time: str_(x['ThoiGian']), action: str_(x['HanhDong']), detail: str_(x['ChiTiet']) }));
  return { student: s, payments: pays, schedules: scheds, exams, history: hist };
}
function apiListSchedules(p) {
  const rows = allSchedules_().filter(s => s.date === p.date);
  return { schedules: rows, date: p.date };
}
function apiListSchedulesRange(p) {
  const rows = allSchedules_().filter(s => s.date >= p.from && s.date <= p.to);
  const dayTotals = {};
  rows.forEach(s => { dayTotals[s.date] = (dayTotals[s.date] || 0) + (s.studentId ? 1 : (s.type === 'Lý thuyết' ? (s.courseSize || 0) : 5)); });
  const markers = [];
  courses_().forEach(c => markers.push({ date: c.closeDate, kind: 'close', label: 'Bế giảng khóa ' + c.code }));
  const byDate = {};
  tab_('KetQuaThi').objs.forEach(e => {
    const d = dstr_(e['NgayThi']);
    if (d >= p.from && d <= p.to) byDate[d] = (byDate[d] || 0) + 1;
  });
  Object.keys(byDate).forEach(d => markers.push({ date: d, kind: 'exam', label: 'Kỳ thi sát hạch (' + byDate[d] + ' HV)' }));
  const cabin = {};
  allStudents_().forEach(s => {
    if (!cabin[s.courseCode]) cabin[s.courseCode] = { courseCode: s.courseCode, students: 0, done3h: 0, sum: 0 };
    cabin[s.courseCode].students++;
    cabin[s.courseCode].sum += s.cabinHours;
    if (s.cabinHours >= 3) cabin[s.courseCode].done3h++;
  });
  return {
    schedules: rows,
    dayTotals: Object.keys(dayTotals).map(d => ({ date: d, sessions: dayTotals[d] })),
    markers: markers.filter(m => m.date >= p.from && m.date <= p.to),
    cabinByCourse: Object.keys(cabin).map(k => ({ courseCode: k, students: cabin[k].students, done3h: cabin[k].done3h, avgHours: Math.round(cabin[k].sum / cabin[k].students * 10) / 10, pct: Math.round(cabin[k].done3h / cabin[k].students * 100) }))
  };
}
function apiListConflicts() { return { conflicts: conflicts_() }; }
function apiListPayments(p) {
  const month = p.month || today_().slice(0, 7);
  const rows = tab_('ThanhToan').objs.map(x => {
    const st = bystId_(num_(x['IDHocVien']));
    return { id: num_(x['ID']), studentId: num_(x['IDHocVien']), studentCode: st ? st.code : '', studentName: st ? st.name : str_(x['HoTen']),
      courseCode: st ? st.courseCode : str_(x['MaKhoa']), date: dstr_(x['Ngay']), amount: num_(x['SoTien']),
      kind: str_(x['Loai']), method: str_(x['PhuongThuc']), note: str_(x['GhiChu']), kmDeducted: num_(x['KmTru']) };
  }).sort((a, b) => b.date.localeCompare(a.date) || b.id - a.id);
  const students = allStudents_();
  const core = students.filter(s => !s.muaLL);
  const mllIds = {};
  students.forEach(s => { if (s.muaLL) mllIds[s.id] = 1; });
  return {
    payments: rows.filter(r => r.date.slice(0, 7) === month),
    totalMonth: rows.filter(r => r.date.slice(0, 7) === month && r.kind === 'Học phí' && !mllIds[r.studentId]).reduce((a, r) => a + r.amount, 0),
    totalPaidAll: core.reduce((a, s) => a + s.tuitionPaid, 0),
    totalGrossAll: core.reduce((a, s) => a + s.tuitionTotal, 0),
    totalPayableAll: core.reduce((a, s) => a + s.tuitionPayable, 0),
    totalDiscountAll: core.reduce((a, s) => a + s.tuitionDiscount, 0),
    totalDebtAll: core.reduce((a, s) => a + s.tuitionPayable, 0) - core.reduce((a, s) => a + s.tuitionPaid, 0)
  };
}
function studentsById_() {
  // Map ID -> học viên, dựng 1 lần / request. Giữ bản ghi ĐẦU tiên khi trùng ID,
  // khớp đúng ngữ nghĩa Array.find() của bystId_() trước đây.
  return memo_('stuById', () => {
    const m = {};
    for (const s of allStudents_()) if (!(s.id in m)) m[s.id] = s;
    return m;
  });
}
function bystId_(id) {
  return studentsById_()[id];
}
function apiListCenters() {
  const centers = tab_('TrungTamLienKet').objs.map(o => ({
    id: num_(o['ID']), name: str_(o['Ten']), commissionRate: num_(o['TyLeHoaHong']),
    studentsCount: num_(o['SoHV']), revenue: num_(o['DoanhThu']), commission: num_(o['HoaHong'])
  }));
  return { centers, commissionTotal: centers.reduce((a, c) => a + c.commission, 0) };
}
function apiGetFleet() {
  const wk = weekHoursByVehicle_();
  const teachers = tab_('GiaoVien').objs.map(o => ({
    id: num_(o['ID']), name: str_(o['HoTen']), shortName: str_(o['TenNgan']), color: str_(o['Mau']),
    kind: str_(o['Loai']), phone: str_(o['SDT']), hourlyRate: num_(o['DonGiaGio']),
    hoursMonth: num_(o['GioThang']), salary: num_(o['DonGiaGio']) * num_(o['GioThang'])
  }));
  const threshold = num_(cfg_('NguongXeQuaGioTuan', 56)) || 56;
  const vehicles = tab_('Xe').objs.map(o => {
    const plate = str_(o['BienSo']);
    const rem = num_(o['KmBaoDuongCuoi']) + num_(o['ChuKyBaoDuongKm']) - num_(o['KmDongHo']);
    return {
      id: num_(o['ID']), plate, licenseClass: str_(o['HangXe']), ownerTeacher: str_(o['GVChuXe']),
      poolCourses: str_(o['KhoaPhuTrach']).split(',').filter(Boolean), status: str_(o['TrangThai']),
      odometerKm: num_(o['KmDongHo']), lastServiceKm: num_(o['KmBaoDuongCuoi']), serviceIntervalKm: num_(o['ChuKyBaoDuongKm']),
      serviceRemainingKm: rem, maintenanceDate: dstr_(o['NgayBaoDuong']), inspectionDate: dstr_(o['HanDangKiem']),
      insuranceExpiry: dstr_(o['HanBaoHiemTNDS']), badgeExpiry: dstr_(o['HanPhuHieu']),
      fuelMonth: num_(o['XangThang']), hoursWeek: wk[plate] || 0, overuse: (wk[plate] || 0) > threshold
    };
  });
  const botuc = tab_('BoTucThueXe').objs.map(o => ({
    id: num_(o['ID']), date: dstr_(o['Ngay']), customerName: str_(o['Khach']), phone: str_(o['SDT']),
    service: str_(o['DichVu']), teacherName: str_(o['GiaoVien']), plate: str_(o['BienSo']),
    slot: str_(o['KhungGio']), hours: num_(o['SoGio']), rate: num_(o['DonGia']), total: num_(o['ThanhTien']), note: str_(o['GhiChu'])
  })).sort((a, b) => b.date.localeCompare(a.date));
  const centers = apiListCenters();
  return { teachers, vehicles, botuc, botucTotal: botuc.reduce((a, b) => a + b.total, 0),
    centers: centers.centers, commissionTotal: centers.commissionTotal,
    fuelTotal: vehicles.reduce((a, v) => a + v.fuelMonth, 0) };
}
function apiListFuelLogs() {
  const logs = tab_('NhatKyNhienLieu').objs.map(o => ({
    id: num_(o['ID']), vehicleId: num_(o['IDXe']), plate: str_(o['BienSo']), date: dstr_(o['Ngay']),
    liters: num_(o['SoLit']), odometerKm: num_(o['KmDongHo']), amount: num_(o['SoTien'])
  })).sort((a, b) => a.plate.localeCompare(b.plate) || a.odometerKm - b.odometerKm);
  const per = {};
  logs.forEach(l => {
    if (!per[l.plate]) per[l.plate] = { plate: l.plate, vehicleId: l.vehicleId, logs: 0, totalLiters: 0, totalAmount: 0, first: null, last: null };
    const p = per[l.plate];
    p.logs++; p.totalAmount += l.amount;
    if (p.first === null) { p.first = l; } else { p.totalLiters += l.liters; }
    p.last = l;
  });
  const perVehicle = Object.keys(per).map(k => {
    const p = per[k];
    const kmDriven = p.last && p.first ? p.last.odometerKm - p.first.odometerKm : 0;
    return { plate: p.plate, vehicleId: p.vehicleId, logs: p.logs, totalLiters: p.totalLiters + (p.first ? p.first.liters : 0),
      totalAmount: p.totalAmount, kmDriven, litersPer100: kmDriven ? Math.round(p.totalLiters / kmDriven * 1000) / 10 : 0 };
  });
  return { logs: logs.slice().sort((a, b) => b.date.localeCompare(a.date)), perVehicle };
}
function admissionObj_(o) {
  const docs = { cccd: str_(o['CCCD']), health: str_(o['KhamSK']), photo: str_(o['Anh34']), form: str_(o['DonDeNghi']), resume: str_(o['LyLich']) };
  const labels = { cccd: 'CCCD', health: 'Khám SK', photo: 'Ảnh 3×4', form: 'Đơn đề nghị', resume: 'Lý lịch' };
  const missing = Object.keys(docs).filter(k => docs[k] !== 'Đã nộp');
  const T = today_();
  let status;
  if (!missing.length) status = 'Đủ giấy tờ';
  else if (dstr_(o['HanBoSung']) < T) status = 'Quá hạn bổ sung';
  else status = 'Còn thiếu ' + missing.length + ' mục';
  const he = dstr_(o['NgayKhamSK']);
  return {
    id: num_(o['ID']), expectedCode: str_(o['MaDuKien']), name: str_(o['HoTen']), phone: str_(o['SDT']),
    licenseClass: str_(o['HangGPLX']), source: str_(o['Nguon']), firstSubmitDate: dstr_(o['NgayNopDau']),
    deadline: dstr_(o['HanBoSung']), docs, docsStatus: status, missingCount: missing.length,
    missingLabels: missing.map(k => labels[k]), healthExamDate: he,
    healthExpiry: he ? String(Number(he.slice(0, 4)) + 1) + he.slice(4) : '', note: str_(o['GhiChu'])
  };
}
function apiListAdmissions() {
  const rows = tab_('TuyenSinh').objs.map(admissionObj_);
  return { rows, summary: {
    total: rows.length,
    enough: rows.filter(r => r.docsStatus === 'Đủ giấy tờ').length,
    missing: rows.filter(r => r.docsStatus.indexOf('Còn thiếu') === 0).length,
    overdue: rows.filter(r => r.docsStatus === 'Quá hạn bổ sung').length } };
}
function apiListExamRounds() {
  const students = allStudents_();
  const regs = tab_('DangKyThi').objs;
  const rounds = tab_('DotSathach').objs.map(o => ({
    id: num_(o['ID']), name: str_(o['TenDot']), examDate: dstr_(o['NgayThi']),
    location: str_(o['NoiThi']), dossierStatus: str_(o['TrangThaiHoSo']),
    registrations: regs.filter(r => num_(r['IDDot']) === num_(o['ID'])).map(r => {
      const st = studentsById_()[num_(r['IDHocVien'])] || {};
      return { id: num_(r['ID']), studentId: num_(r['IDHocVien']), sbd: str_(r['SBD']),
        code: st.code || str_(r['MaHV']), name: st.name || str_(r['HoTen']), courseCode: st.courseCode || '',
        licenseClass: st.licenseClass || '', dob: st.dob || '', cccd: st.cccd || '' };
    })
  }));
  const passedIds = new Set(tab_('KetQuaThi').objs.filter(e => str_(e['KetQua']) === 'Đậu').map(e => num_(e['IDHocVien'])));
  const regIds = new Set(regs.map(r => num_(r['IDHocVien'])));
  const registrable = students.filter(s => s.eligible && !passedIds.has(s.id) && !regIds.has(s.id));
  return { rounds, registrable };
}
function apiListExams() {
  const students = allStudents_();
  const rows = tab_('KetQuaThi').objs.map(o => {
    const st = studentsById_()[num_(o['IDHocVien'])] || {};
    return { id: num_(o['ID']), studentId: num_(o['IDHocVien']), code: st.code || str_(o['MaHV']),
      name: st.name || str_(o['HoTen']), courseCode: st.courseCode || str_(o['MaKhoa']), licenseClass: st.licenseClass || '',
      examDate: dstr_(o['NgayThi']), sbd: str_(o['SBD']), theory: str_(o['LyThuyet']),
      theoryScore: o['DiemLT'] === '' ? null : num_(o['DiemLT']), sahinh: str_(o['SaHinh']),
      sahinhScore: o['DiemSH'] === '' ? null : num_(o['DiemSH']), road: str_(o['DuongTruong']),
      roadScore: o['DiemDT'] === '' ? null : num_(o['DiemDT']), passed: str_(o['KetQua']) === 'Đậu',
      retakeFee: num_(o['PhiThiLai']), retakeFeeNote: str_(o['GhiChuPhi']), note: str_(o['GhiChu']),
      hasLicense: str_(o['CoBang']) === 'Có', licenseDate: dstr_(o['NgayCoBang']) };
  });
  const examined = rows.filter(r => r.theory !== 'Chưa thi');
  const passed = rows.filter(r => r.passed);
  const byCourse = {};
  examined.forEach(r => {
    if (!byCourse[r.courseCode]) byCourse[r.courseCode] = { courseCode: r.courseCode, examined: 0, passed: 0 };
    byCourse[r.courseCode].examined++;
    if (r.passed) byCourse[r.courseCode].passed++;
  });
  const eligible = students.filter(s => s.eligible && !passed.some(p => p.studentId === s.id));
  return {
    results: rows, eligible,
    eligibleTotal: eligible.length, passedTotal: passed.length,
    licensedCount: students.filter(s => s.hasLicense).length,
    passedNoLicense: passed.filter(p => !p.hasLicense).length,
    passRate: examined.length ? Math.round(passed.length / examined.length * 100) : 0,
    passByCourse: Object.keys(byCourse).map(k => ({ courseCode: k, examined: byCourse[k].examined, passed: byCourse[k].passed, rate: Math.round(byCourse[k].passed / byCourse[k].examined * 1000) / 10 })),
    nextExamDate: tab_('DotSathach').objs.map(o => dstr_(o['NgayThi'])).filter(d => d && d >= today_()).sort()[0] || ''
  };
}
function apiGetReports() {
  const students = allStudents_();
  const core = students.filter(s => !s.muaLL);
  const rows = allSchedules_();
  const T = today_();
  const pl = JSON.parse(cfg_('PLSeries', '[]'));
  const cost = JSON.parse(cfg_('CostBreakdown', '[]'));
  const revByCourse = {};
  tab_('ThanhToan').objs.forEach(p => {
    if (str_(p['Loai']) !== 'Học phí') return;
    const st = bystId_(num_(p['IDHocVien']));
    if (!st || st.muaLL) return;
    revByCourse[st.courseCode] = (revByCourse[st.courseCode] || 0) + num_(p['SoTien']);
  });
  const byVehicle = {}, bySlot = {};
  rows.filter(s => s.studentId).forEach(s => {
    if (s.type === 'DAT') {
      if (!byVehicle[s.plate]) byVehicle[s.plate] = { plate: s.plate, km: 0, sessions: 0 };
      byVehicle[s.plate].km += (s.actualKm !== null ? s.actualKm : s.km);
      byVehicle[s.plate].sessions++;
    }
    const key = s.slot;
    if (!bySlot[key]) bySlot[key] = { slot: key, dat: 0, sahinh: 0 };
    if (s.type === 'DAT') bySlot[key].dat++;
    if (s.type === 'Sa hình') bySlot[key].sahinh++;
  });
  // (Ban cu co dong apiListSchedulesRange 7 thang ma ket qua khong dung o dau — da bo.)
  // sessionsProgress tinh tai cho (cong thuc y het apiGetDashboard) — tranh goi
  // lai apiGetDashboard() (keo theo notifications + conflicts) va apiListExams() 2 lan.
  const indivR = rows.filter(s => s.studentId);
  const leR = indivR.filter(s => s.date <= T);
  const cntR = (arr, t) => arr.filter(s => s.type === t).length;
  const ltWeightedR = rows.filter(s => s.type === 'Lý thuyết').reduce((a, s) => a + (s.courseSize || 0), 0);
  const scheduledR = cntR(leR, 'DAT') + cntR(leR, 'Sa hình') + cntR(leR, 'Thi tốt nghiệp') + ltWeightedR;
  const totalR = students.length * 16;
  const daysR = new Set(indivR.map(s => s.date)).size;
  const examsR = apiListExams();
  return {
    totalStudents: students.length,
    studentsByCourse: courses_().map(c => ({ courseCode: c.code, licenseClass: c.licenseClass, count: students.filter(s => s.courseCode === c.code).length })),
    receivable: {
      totalGross: core.reduce((a, s) => a + s.tuitionTotal, 0) / 1000000,
      totalDiscount: core.reduce((a, s) => a + s.tuitionDiscount, 0) / 1000000,
      totalDue: core.reduce((a, s) => a + s.tuitionPayable, 0) / 1000000,
      collected: core.reduce((a, s) => a + s.tuitionPaid, 0) / 1000000,
      debt: (core.reduce((a, s) => a + s.tuitionPayable, 0) - core.reduce((a, s) => a + s.tuitionPaid, 0)) / 1000000
    },
    revenueByCourse: Object.keys(revByCourse).map(k => {
      const c = courseMap_()[k];
      return { label: k, students: students.filter(s => s.courseCode === k).length, perCourse: c ? c.tuition / 1000000 : 0, total: Math.round(revByCourse[k] / 1000000) };
    }),
    plSeries: pl, costBreakdown: cost,
    datByVehicle: Object.keys(byVehicle).map(k => byVehicle[k]).sort((a, b) => b.km - a.km),
    datBySlot: Object.keys(bySlot).sort().map(k => bySlot[k]),
    sessionsProgress: {
      total: totalR, scheduled: scheduledR, remaining: totalR - scheduledR, gridCells: totalR,
      datSpecial: cntR(leR, 'DAT'), sahinh: cntR(leR, 'Sa hình') + cntR(leR, 'Thi tốt nghiệp'), regular: ltWeightedR,
      off: 0, days: daysR, avgPerDay: daysR ? Math.round(scheduledR / daysR * 10) / 10 : 0,
      from: '2026-06-07', to: '2026-12-31'
    },
    eligibleCount: examsR.eligibleTotal,
    passRate: examsR.passRate
  };
}
function apiGetGovReports() {
  const students = allStudents_();
  const exams = tab_('KetQuaThi').objs;
  const byCourse = courses_().map(c => {
    const sts = students.filter(s => s.courseCode === c.code);
    const ex = exams.filter(e => {
      const st = studentsById_()[num_(e['IDHocVien'])];
      return st && st.courseCode === c.code && str_(e['KetQua']) !== 'Chưa thi';
    });
    return { courseCode: c.code, licenseClass: c.licenseClass, openDate: c.openDate, closeDate: c.closeDate,
      total: sts.length, graduated: sts.filter(s => s.gradResult === 'Đậu').length,
      examined: ex.length, passed: ex.filter(e => str_(e['KetQua']) === 'Đậu').length };
  });
  const enrollments = students.map(s => ({ code: s.code, name: s.name, cccd: s.cccd, dob: s.dob,
    licenseClass: s.licenseClass, courseCode: s.courseCode, regDate: s.regDate, openDate: s.openDate }));
  const byDate = {};
  exams.forEach(e => {
    const d = dstr_(e['NgayThi']);
    if (!byDate[d]) byDate[d] = { examDate: d, total: 0, passed: 0, students: [] };
    const st = studentsById_()[num_(e['IDHocVien'])] || {};
    byDate[d].total++;
    if (str_(e['KetQua']) === 'Đậu') byDate[d].passed++;
    byDate[d].students.push({ code: st.code || str_(e['MaHV']), name: st.name || str_(e['HoTen']),
      courseCode: st.courseCode || '', sbd: str_(e['SBD']), theory: str_(e['LyThuyet']),
      sahinh: str_(e['SaHinh']), road: str_(e['DuongTruong']), result: str_(e['KetQua']) });
  });
  return { byCourse, enrollments, examByDate: Object.keys(byDate).sort().map(k => byDate[k]) };
}
function apiNotifications() {
  const T = today_();
  const students = allStudents_();
  const alerts = [];
  const overdue = students.filter(s => s.closeDate && s.closeDate < T && s.datKmDone < s.datKmTotal);
  if (overdue.length) alerts.push({ kind: 'DAT quá hạn', level: 'error', message: overdue.length + ' HV chưa đủ km DAT, đã quá hạn hoàn thành' });
  const soon = students.filter(s => {
    if (!s.closeDate || s.datKmDone >= s.datKmTotal) return false;
    const days = (new Date(s.closeDate) - new Date(T)) / 86400000;
    return days >= 0 && days <= 5;
  });
  if (soon.length) alerts.push({ kind: 'DAT sắp hết hạn', level: 'warning', message: soon.length + ' HV sẽ hết hạn DAT trong 5 ngày tới' });
  students.filter(s => s.datMarkerKm !== 0).forEach(s => {
    alerts.push({ kind: 'Mốc DAT', level: s.datMarkerKm < 0 ? 'warning' : 'info',
      message: s.datMarkerKm < 0
        ? s.name.split(' ').slice(-2).join(' ') + ' (MHV ' + s.maHv + ') — ' + s.courseCode + ': mốc DAT ' + s.datMarkerKm + 'KM (còn thiếu ' + (s.datKmTotal - s.datKmDone) + 'km)'
        : s.name.split(' ').slice(-2).join(' ') + ' (MHV ' + s.maHv + ') — ' + s.courseCode + ': đã chạy DAT ' + s.datKmDone + 'KM' });
  });
  const docAlerts = [], svcAlerts = [];
  tab_('Xe').objs.forEach(o => {
    const plate = str_(o['BienSo']);
    [['HanBaoHiemTNDS', 'Bảo hiểm TNDS'], ['HanDangKiem', 'Đăng kiểm'], ['HanPhuHieu', 'Phù hiệu xe tập lái']].forEach(([col, label]) => {
      const exp = dstr_(o[col]);
      if (!exp) return;
      const days = Math.round((new Date(exp) - new Date(T)) / 86400000);
      if (days <= 30) docAlerts.push({ plate, docLabel: label, expiryDate: exp, daysLeft: days, state: days < 0 ? 'Đã hết hạn' : 'Sắp hết hạn' });
    });
    const rem = num_(o['KmBaoDuongCuoi']) + num_(o['ChuKyBaoDuongKm']) - num_(o['KmDongHo']);
    if (rem <= 500) svcAlerts.push({ plate, odometerKm: num_(o['KmDongHo']), remainingKm: rem });
  });
  if (docAlerts.length) alerts.push({ kind: 'Giấy tờ xe', level: 'warning', message: docAlerts.length + ' giấy tờ xe hết hạn / sắp hết hạn (< 30 ngày) — xem danh sách chi tiết ở mục Giấy tờ xe bên dưới' });
  if (svcAlerts.length) alerts.push({ kind: 'Bảo dưỡng theo km', level: 'warning', message: svcAlerts.length + ' xe tới hạn bảo dưỡng theo km (còn <= 500km) — xem danh sách chi tiết ở mục Bảo dưỡng theo km bên dưới' });
  const health = students.filter(s => {
    if (!s.healthExpiry) return false;
    const days = (new Date(s.healthExpiry) - new Date(T)) / 86400000;
    return days <= 30;
  }).map(s => ({ code: s.code, name: s.name, healthExpiry: s.healthExpiry }));
  const cabinNS = students.filter(s => s.cabinHours === 0 && s.openDate && s.openDate <= T)
    .map(s => ({ code: s.code, name: s.name, courseCode: s.courseCode, homeTeacherName: s.homeTeacherName }));
  /* Nhắc việc theo yêu cầu chị Hồng (28/10): khóa sắp bế giảng ≤3 ngày + HV sắp xuống DAT ≤2 ngày,
     kèm cờ thiếu DAT/cabin cho từng HV — ngày lấy theo khóa (NgayBeGiang / NgayBatDauDAT), HV thuộc khóa nào tính theo khóa đó.
     LƯU Ý: riêng phần nhắc này tính theo NGÀY THẬT (Asia/Ho_Chi_Minh), không theo mốc neo NgayHienTai của dữ liệu —
     vì bế giảng/xuống DAT là việc ngoài đời, mốc neo đứng yên sẽ làm nhắc không bao giờ kích hoạt. */
  const realToday_ = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
  const daysTo_ = (ds) => Math.round((new Date(ds) - new Date(realToday_)) / 86400000);
  const closingSoon = students.filter(s => s.closeDate && daysTo_(s.closeDate) >= 0 && daysTo_(s.closeDate) <= 3)
    .map(s => ({ id: s.id, code: s.code, name: s.name, courseCode: s.courseCode, closeDate: s.closeDate, daysLeft: daysTo_(s.closeDate), datOk: s.datKmDone >= s.datKmTotal, cabinOk: s.cabinHours >= 3, done: s.gradResult === 'Đậu' || s.hasLicense }));
  const datStartSoon = students.filter(s => s.datStartDate && daysTo_(s.datStartDate) >= 0 && daysTo_(s.datStartDate) <= 2)
    .map(s => ({ id: s.id, code: s.code, name: s.name, courseCode: s.courseCode, datStartDate: s.datStartDate, daysLeft: daysTo_(s.datStartDate), cabinOk: s.cabinHours >= 3, theoryOk: !!(s.criteria && s.criteria.theory) }));
  if (closingSoon.length) alerts.push({ kind: 'Sắp bế giảng', level: 'error', message: closingSoon.length + ' HV thuộc khóa bế giảng trong 3 ngày tới — xem card nhắc đỏ ở trang Tổng quan' });
  if (datStartSoon.length) alerts.push({ kind: 'Sắp xuống DAT', level: 'warning', message: datStartSoon.length + ' HV bắt đầu giai đoạn DAT trong 2 ngày tới — xem card nhắc đỏ ở trang Tổng quan' });
  const notifs = JSON.parse(cfg_('ThongBaoMau', '[]'));
  const wk = weekHoursByVehicle_();
  const threshold = num_(cfg_('NguongXeQuaGioTuan', 56)) || 56;
  const overuse = Object.keys(wk).filter(k => wk[k] > threshold).map(k => ({ plate: k, hoursWeek: wk[k] }));
  const datRow_ = (s) => ({ id: s.id, code: s.code, name: s.name, courseCode: s.courseCode, closeDate: s.closeDate, daysLeft: Math.round((new Date(s.closeDate) - new Date(T)) / 86400000), datKmDone: s.datKmDone, datKmTotal: s.datKmTotal, cabinOk: s.cabinHours >= 3, homeTeacherName: s.homeTeacherName });
  return { alerts, notifications: notifs,
    datOverdue: { count: overdue.length, rows: overdue.map(datRow_) },
    datDueSoon: { count: soon.length, rows: soon.map(datRow_) },
    reminders: { closingSoon: { count: closingSoon.length, rows: closingSoon }, datStartSoon: { count: datStartSoon.length, rows: datStartSoon } },
    vehicleDocAlerts: { count: docAlerts.length, rows: docAlerts },
    vehicleServiceAlerts: { count: svcAlerts.length, rows: svcAlerts },
    healthAlerts: { count: health.length, rows: health },
    cabinNotStarted: { count: cabinNS.length, rows: cabinNS }, vehicleOveruse: overuse };
}
function apiGetLookups() {
  return {
    courses: courses_().map(c => ({ code: c.code, licenseClass: c.licenseClass, tuition: c.tuition })),
    teachers: tab_('GiaoVien').objs.map(o => ({ id: num_(o['ID']), name: str_(o['HoTen']), shortName: str_(o['TenNgan']), color: str_(o['Mau']), kind: str_(o['Loai']) })),
    vehicles: tab_('Xe').objs.map(o => ({ id: num_(o['ID']), plate: str_(o['BienSo']), status: str_(o['TrangThai']) })),
    students: allStudents_().map(s => ({ id: s.id, code: s.code, maHv: s.maHv, name: s.name, courseCode: s.courseCode, status: s.status, datKmDone: s.datKmDone, datKmTotal: s.datKmTotal, plateKg: s.plateKg }))
  };
}


// ---------- API: Ghi ----------
function apiSaveStudent(p) {
  let id = num_(p.id);
  if (!id) {
    id = nextId_('HocVien');
    const cm = courseMap_();
    const c = cm[p.courseCode] || {};
    const hvObjs = tab_('HocVien').objs;
    const maxSach = hvObjs.reduce((a, o) => Math.max(a, Number(String(o['MaHVSach']).replace(/\D+/g, '')) || 0), 0);
    const maxCode = hvObjs.reduce((a, o) => Math.max(a, Number(String(o['MaHV']).replace(/\D+/g, '')) || 0), 0);
    appendObj_('HocVien', {
      ID: id, MaHV: 'HV' + String(maxCode + 1).padStart(3, '0'), MaHVSach: String(maxSach + 1),
      HoTen: p.name || '', NgaySinh: p.dob || '', CCCD: p.cccd || '', SDT: p.phone || '', DiaChi: p.address || '',
      MaKhoa: p.courseCode || '', Hang: p.licenseClass || c.licenseClass || '', TrungTamLienKet: p.centerName || '',
      GVChinh: '', XeQuen: '', BienSoKG: '', NgayDangKy: today_(), NgayKhamSK: p.healthExamDate || '',
      HanKhamSK: p.healthExamDate ? String(Number(String(p.healthExamDate).slice(0, 4)) + 1) + String(p.healthExamDate).slice(4) : '',
      GioCabin: 0, BuoiCabin: 0, DiemLyThuyet: '', DatLyThuyet: 'Chưa đạt', DaHocSaHinh: 'Chưa học',
      KmDATDaChay: 0, KmDATMucTieu: TARGET_KM[p.licenseClass || c.licenseClass] || 810, MocKmDAT: 0,
      GiamTien: num_(p.discount), MuaLL: p.muaLL ? 1 : '', HocPhi: num_(p.tuition) > 0 ? num_(p.tuition) : priceFor_(p.licenseClass || c.licenseClass || ''), LyDoGiam: p.discountReason || '', TrangThai: p.status || 'Đang học',
      NgayThiTotNghiep: '', KetQuaTotNghiep: '', SoChungChi: '', NgayCapChungChi: '',
      NgayCoBang: '', CoBang: 'Chưa', DiemBangLai: ''
    });
  }
  const patch = {};
  if (p.name !== undefined) patch['HoTen'] = p.name;
  if (p.dob !== undefined) patch['NgaySinh'] = p.dob;
  if (p.cccd !== undefined) patch['CCCD'] = p.cccd;
  if (p.phone !== undefined) patch['SDT'] = p.phone;
  if (p.address !== undefined) patch['DiaChi'] = p.address;
  if (p.status !== undefined) patch['TrangThai'] = p.status;
  if (p.centerName !== undefined) patch['TrungTamLienKet'] = p.centerName;
  if (p.healthExamDate !== undefined) {
    patch['NgayKhamSK'] = p.healthExamDate;
    patch['HanKhamSK'] = p.healthExamDate ? String(Number(String(p.healthExamDate).slice(0, 4)) + 1) + String(p.healthExamDate).slice(4) : '';
  }
  if (p.discount !== undefined) patch['GiamTien'] = num_(p.discount);
  if (p.discountReason !== undefined) patch['LyDoGiam'] = p.discountReason;
  if (p.muaLL !== undefined) patch['MuaLL'] = p.muaLL ? 1 : '';
  if (p.tuition !== undefined) patch['HocPhi'] = num_(p.tuition);
  if (p.homeTeacherId !== undefined) {
    const t = tab_('GiaoVien').objs.find(o => num_(o['ID']) === num_(p.homeTeacherId));
    if (t) patch['GVChinh'] = str_(t['HoTen']);
  } else if (p.homeTeacherName !== undefined) patch['GVChinh'] = p.homeTeacherName;
  if (p.courseCode !== undefined) {
    patch['MaKhoa'] = p.courseCode;
    const c = courseMap_()[p.courseCode];
    if (c) patch['Hang'] = c.licenseClass;
  }
  if (p.licenseClass !== undefined) patch['Hang'] = p.licenseClass;
  writeRow_('HocVien', id, patch);
  const st = findRow_('HocVien', id);
  logHistory_('saveStudent', st ? str_(st['MaHV']) : '', (num_(p.id) ? 'Cập nhật HV ' : 'Thêm HV ') + (p.name || ''));
  return { ok: true, id, code: st ? str_(st['MaHV']) : '', message: 'Đã lưu học viên' };
}
function apiSetStudentStatus(p) {
  const st = findRow_('HocVien', num_(p.id));
  if (!st) return { ok: false, error: 'Không tìm thấy học viên' };
  writeRow_('HocVien', num_(p.id), { TrangThai: p.status || 'Đang học' });
  if (p.status === 'Hoàn phí' && num_(p.refundAmount) > 0) {
    appendObj_('ThanhToan', { ID: nextId_('ThanhToan'), IDHocVien: num_(p.id), MaHV: str_(st['MaHV']),
      HoTen: str_(st['HoTen']), MaKhoa: str_(st['MaKhoa']), Ngay: today_(), SoTien: -Math.abs(num_(p.refundAmount)),
      Loai: 'Hoàn phí', PhuongThuc: 'Chuyển khoản', GhiChu: p.note || 'Hoàn phí học viên', KmTru: 0 });
  }
  logHistory_('setStudentStatus', str_(st['MaHV']), 'Đổi trạng thái: ' + (p.status || '') + (p.note ? ' — ' + p.note : ''));
  return { ok: true };
}
function apiTransferCourse(p) {
  const st = findRow_('HocVien', num_(p.id));
  if (!st) return { ok: false, error: 'Không tìm thấy học viên' };
  const oldCode = str_(st['MaKhoa']);
  const c = courseMap_()[p.toCourse];
  writeRow_('HocVien', num_(p.id), { MaKhoa: p.toCourse, Hang: c ? c.licenseClass : str_(st['Hang']) });
  logHistory_('transferCourse', str_(st['MaHV']), 'Chuyển khóa ' + oldCode + ' → ' + p.toCourse);
  return { ok: true, fromCourse: oldCode, toCourse: p.toCourse };
}
function slotRangeMin_(slot) {
  const m = String(slot || '').trim().match(/^(\d{1,2})(?::(\d{2}))?\s*-\s*(\d{1,2})(?::(\d{2}))?/);
  if (!m) return null;
  const f = (h, mi) => Number(h) * 60 + Number(mi || 0);
  return [f(m[1], m[2]), f(m[3], m[4])];
}
function slotStartMinutes_(slot) {
  const r = slotRangeMin_(slot);
  return r ? r[0] : 480;
}
function apiCreateSchedule(p) {
  const st = p.studentId ? bystId_(num_(p.studentId)) : null;
  const teacher = p.teacherId ? tab_('GiaoVien').objs.find(o => num_(o['ID']) === num_(p.teacherId)) : null;
  const nid = nextId_('LichHoc');
  appendObj_('LichHoc', {
    ID: nid, Ngay: p.date || today_(), KhungGio: p.slot || '08-10h', PhutBatDau: slotStartMinutes_(p.slot),
    LoaiBuoi: p.type || 'DAT', MaKhoa: p.courseCode || (st ? st.courseCode : ''),
    IDHocVien: p.studentId ? num_(p.studentId) : '', HoTenHV: st ? st.name : '',
    IDGiaoVien: p.teacherId ? num_(p.teacherId) : '', TenGiaoVien: teacher ? str_(teacher['HoTen']) : '',
    IDXe: p.vehicleId ? num_(p.vehicleId) : '', BienSo: p.plate || '',
    TrungTamLienKet: st ? st.centerName : '', KmKeHoach: num_(p.km),
    KmThucTe: p.actualKm === undefined || p.actualKm === null || p.actualKm === '' ? '' : num_(p.actualKm),
    DiemDanh: p.attendance || 'Đã học', NhanNhom: p.groupLabel || '', GhiChu: p.note || ''
  });
  logHistory_('createSchedule', st ? st.code : '', 'Thêm buổi ' + (p.type || '') + ' ngày ' + (p.date || ''));
  return { ok: true, id: nid };
}
function apiUpdateSchedule(p) {
  if (!findRow_('LichHoc', num_(p.id))) return { ok: false, error: 'Không tìm thấy buổi học' };
  const patch = {};
  if (p.date !== undefined) patch['Ngay'] = p.date;
  if (p.slot !== undefined) { patch['KhungGio'] = p.slot; patch['PhutBatDau'] = slotStartMinutes_(p.slot); }
  if (p.type !== undefined) patch['LoaiBuoi'] = p.type;
  if (p.teacherId !== undefined) {
    patch['IDGiaoVien'] = num_(p.teacherId);
    const t = tab_('GiaoVien').objs.find(o => num_(o['ID']) === num_(p.teacherId));
    if (t) patch['TenGiaoVien'] = str_(t['HoTen']);
  }
  if (p.vehicleId !== undefined) patch['IDXe'] = num_(p.vehicleId);
  if (p.plate !== undefined) patch['BienSo'] = p.plate;
  if (p.km !== undefined) patch['KmKeHoach'] = num_(p.km);
  if (p.actualKm !== undefined) patch['KmThucTe'] = p.actualKm === null || p.actualKm === '' ? '' : num_(p.actualKm);
  if (p.attendance !== undefined) patch['DiemDanh'] = p.attendance;
  if (p.note !== undefined) patch['GhiChu'] = p.note;
  writeRow_('LichHoc', num_(p.id), patch);
  logHistory_('updateSchedule', 'LH' + p.id, 'Sửa buổi học');
  return { ok: true };
}
function apiDeleteSchedule(p) {
  const ok = deleteRow_('LichHoc', num_(p.id));
  logHistory_('deleteSchedule', 'LH' + p.id, 'Xóa buổi học');
  return { ok: !!ok };
}
function apiAddPayment(p) {
  const st = bystId_(num_(p.studentId));
  const nid = nextId_('ThanhToan');
  appendObj_('ThanhToan', { ID: nid, IDHocVien: num_(p.studentId), MaHV: st ? st.code : '', HoTen: st ? st.name : '',
    MaKhoa: st ? st.courseCode : '', Ngay: p.date || today_(), SoTien: num_(p.amount),
    Loai: p.kind || 'Học phí', PhuongThuc: p.method || 'Tiền mặt', GhiChu: p.note || '', KmTru: 0 });
  logHistory_('addPayment', st ? st.code : '', 'Thu ' + (p.kind || 'Học phí') + ' ' + num_(p.amount) + 'đ');
  return { ok: true, id: nid };
}
function apiSaveTeacher(p) {
  let id = num_(p.id);
  if (!id) {
    id = nextId_('GiaoVien');
    appendObj_('GiaoVien', { ID: id, HoTen: '', TenNgan: '', Mau: '#90A4AE', Loai: 'chinh', SDT: '', DonGiaGio: 0, GioThang: 0, LuongThang: 0 });
  }
  const patch = {};
  if (p.name !== undefined) {
    patch['HoTen'] = p.name;
    patch['TenNgan'] = p.shortName !== undefined ? p.shortName : String(p.name).trim().split(' ').slice(-2).join(' ');
  }
  if (p.shortName !== undefined) patch['TenNgan'] = p.shortName;
  if (p.kind !== undefined) patch['Loai'] = p.kind;
  if (p.phone !== undefined) patch['SDT'] = p.phone;
  if (p.color !== undefined) patch['Mau'] = p.color;
  writeRow_('GiaoVien', id, patch);
  logHistory_('saveTeacher', 'GV' + id, 'Lưu GV ' + (p.name || ''));
  return { ok: true, id };
}
function apiSaveVehicle(p) {
  let id = num_(p.id);
  if (!id) {
    id = nextId_('Xe');
    appendObj_('Xe', { ID: id, BienSo: p.plate || '', HangXe: '', GVChuXe: '', KhoaPhuTrach: '', TrangThai: 'Đang sử dụng',
      KmDongHo: 0, KmBaoDuongCuoi: 0, ChuKyBaoDuongKm: 10000, NgayBaoDuong: '', HanDangKiem: '', HanBaoHiemTNDS: '', HanPhuHieu: '', XangThang: 0 });
  }
  const patch = {};
  if (p.plate !== undefined) patch['BienSo'] = p.plate;
  if (p.ownerTeacherId !== undefined || p.ownerTeacher !== undefined) {
    let ownerName = p.ownerTeacher || '';
    if (p.ownerTeacherId) {
      const t = tab_('GiaoVien').objs.find(o => num_(o['ID']) === num_(p.ownerTeacherId));
      if (t) ownerName = str_(t['HoTen']);
    }
    patch['GVChuXe'] = ownerName;
  }
  if (p.status !== undefined) patch['TrangThai'] = p.status;
  writeRow_('Xe', id, patch);
  logHistory_('saveVehicle', 'XE' + id, 'Lưu xe ' + (p.plate || ''));
  return { ok: true, id };
}
function apiSaveCenter(p) {
  let id = num_(p.id);
  if (!id) {
    id = nextId_('TrungTamLienKet');
    appendObj_('TrungTamLienKet', { ID: id, Ten: '', TyLeHoaHong: 0, SoHV: 0, DoanhThu: 0, HoaHong: 0 });
  }
  const patch = {};
  if (p.name !== undefined) patch['Ten'] = p.name;
  if (p.commissionRate !== undefined) patch['TyLeHoaHong'] = num_(p.commissionRate);
  if (p.studentsCount !== undefined) patch['SoHV'] = num_(p.studentsCount);
  if (p.revenue !== undefined) patch['DoanhThu'] = num_(p.revenue);
  if (p.revenue !== undefined && p.commissionRate !== undefined) patch['HoaHong'] = Math.round(num_(p.revenue) * num_(p.commissionRate) / 100);
  else if (p.commission !== undefined) patch['HoaHong'] = num_(p.commission);
  writeRow_('TrungTamLienKet', id, patch);
  logHistory_('saveCenter', 'TT' + id, 'Lưu TT ' + (p.name || ''));
  return { ok: true, id };
}
function apiDeleteCenter(p) {
  const ok = deleteRow_('TrungTamLienKet', num_(p.id));
  logHistory_('deleteCenter', 'TT' + p.id, 'Xóa trung tâm liên kết');
  return { ok: !!ok };
}
function apiSaveBotuc(p) {
  let id = num_(p.id);
  if (!id) {
    id = nextId_('BoTucThueXe');
    appendObj_('BoTucThueXe', { ID: id, Ngay: today_(), Khach: '', SDT: '', DichVu: 'Bổ túc tay lái', GiaoVien: '', BienSo: '', KhungGio: '', SoGio: 0, DonGia: 0, ThanhTien: 0, GhiChu: '' });
  }
  const patch = {};
  if (p.date !== undefined) patch['Ngay'] = p.date;
  if (p.customerName !== undefined) patch['Khach'] = p.customerName;
  if (p.phone !== undefined) patch['SDT'] = p.phone;
  if (p.service !== undefined) patch['DichVu'] = p.service;
  if (p.teacherName !== undefined) patch['GiaoVien'] = p.teacherName;
  if (p.plate !== undefined) patch['BienSo'] = p.plate;
  if (p.slot !== undefined) patch['KhungGio'] = p.slot;
  if (p.hours !== undefined) patch['SoGio'] = num_(p.hours);
  if (p.rate !== undefined) patch['DonGia'] = num_(p.rate);
  if (p.hours !== undefined && p.rate !== undefined) patch['ThanhTien'] = Math.round(num_(p.hours) * num_(p.rate));
  if (p.note !== undefined) patch['GhiChu'] = p.note;
  writeRow_('BoTucThueXe', id, patch);
  logHistory_('saveBotuc', 'BT' + id, 'Lưu bổ túc/thuê xe');
  return { ok: true, id, total: patch['ThanhTien'] };
}
function apiDeleteBotuc(p) {
  const ok = deleteRow_('BoTucThueXe', num_(p.id));
  logHistory_('deleteBotuc', 'BT' + p.id, 'Xóa dòng bổ túc/thuê xe');
  return { ok: !!ok };
}
function apiAddFuelLog(p) {
  const veh = tab_('Xe').objs.find(o => num_(o['ID']) === num_(p.vehicleId));
  const nid = nextId_('NhatKyNhienLieu');
  appendObj_('NhatKyNhienLieu', { ID: nid, IDXe: num_(p.vehicleId), BienSo: veh ? str_(veh['BienSo']) : '',
    Ngay: p.date || today_(), SoLit: num_(p.liters), KmDongHo: num_(p.odometerKm), SoTien: num_(p.amount) });
  logHistory_('addFuelLog', veh ? str_(veh['BienSo']) : '', 'Đổ xăng ' + num_(p.liters) + ' lít');
  return { ok: true, id: nid };
}
function apiDeleteFuelLog(p) {
  const ok = deleteRow_('NhatKyNhienLieu', num_(p.id));
  logHistory_('deleteFuelLog', 'NL' + p.id, 'Xóa nhật ký nhiên liệu');
  return { ok: !!ok };
}
function apiSaveAdmission(p) {
  let id = num_(p.id);
  if (!id) {
    id = nextId_('TuyenSinh');
    const maxN = tab_('TuyenSinh').objs.reduce((a, o) => Math.max(a, Number(String(o['MaDuKien']).replace(/\D+/g, '')) || 0), 0);
    appendObj_('TuyenSinh', { ID: id, MaDuKien: 'TS' + String(maxN + 1).padStart(3, '0'), HoTen: '', SDT: '',
      HangGPLX: 'B2', Nguon: '', NgayNopDau: today_(), HanBoSung: '', CCCD: 'Chưa nộp', KhamSK: 'Chưa nộp',
      Anh34: 'Chưa nộp', DonDeNghi: 'Chưa nộp', LyLich: 'Chưa nộp', NgayKhamSK: '', GhiChu: '' });
  }
  const patch = {};
  if (p.name !== undefined) patch['HoTen'] = p.name;
  if (p.phone !== undefined) patch['SDT'] = p.phone;
  if (p.licenseClass !== undefined) patch['HangGPLX'] = p.licenseClass;
  if (p.source !== undefined) patch['Nguon'] = p.source;
  if (p.firstSubmitDate !== undefined) patch['NgayNopDau'] = p.firstSubmitDate;
  if (p.healthExamDate !== undefined) patch['NgayKhamSK'] = p.healthExamDate;
  if (p.deadline !== undefined) patch['HanBoSung'] = p.deadline;
  if (p.note !== undefined) patch['GhiChu'] = p.note;
  writeRow_('TuyenSinh', id, patch);
  logHistory_('saveAdmission', 'TS' + id, 'Lưu hồ sơ tuyển sinh ' + (p.name || ''));
  return { ok: true, id };
}
function apiDeleteAdmission(p) {
  const ok = deleteRow_('TuyenSinh', num_(p.id));
  logHistory_('deleteAdmission', 'TS' + p.id, 'Xóa hồ sơ tuyển sinh');
  return { ok: !!ok };
}
function apiSetAdmissionDoc(p) {
  const colMap = { cccd: 'CCCD', health: 'KhamSK', photo: 'Anh34', form: 'DonDeNghi', resume: 'LyLich' };
  const col = colMap[p.field];
  if (!col) return { ok: false, error: 'Mục giấy tờ không hợp lệ' };
  const ok = writeRow_('TuyenSinh', num_(p.id), { [col]: p.submitted ? 'Đã nộp' : 'Chưa nộp' });
  if (!ok) return { ok: false, error: 'Không tìm thấy hồ sơ' };
  logHistory_('setAdmissionDoc', 'TS' + p.id, (p.submitted ? 'Đánh dấu đã nộp ' : 'Bỏ nộp ') + col);
  return { ok: true };
}
function apiSaveExamRound(p) {
  let id = num_(p.id);
  if (!id) {
    id = nextId_('DotSathach');
    appendObj_('DotSathach', { ID: id, TenDot: '', NgayThi: '', NoiThi: '', TrangThaiHoSo: 'Đang chuẩn bị' });
  }
  const patch = {};
  if (p.name !== undefined) patch['TenDot'] = p.name;
  if (p.examDate !== undefined) patch['NgayThi'] = p.examDate;
  if (p.location !== undefined) patch['NoiThi'] = p.location;
  if (p.dossierStatus !== undefined) patch['TrangThaiHoSo'] = p.dossierStatus;
  writeRow_('DotSathach', id, patch);
  logHistory_('saveExamRound', 'DOT' + id, 'Lưu đợt thi ' + (p.name || ''));
  return { ok: true, id };
}
function apiDeleteExamRound(p) {
  tab_('DangKyThi').objs.filter(o => num_(o['IDDot']) === num_(p.id))
    .forEach(o => deleteRow_('DangKyThi', num_(o['ID'])));
  const ok = deleteRow_('DotSathach', num_(p.id));
  logHistory_('deleteExamRound', 'DOT' + p.id, 'Xóa đợt thi + đăng ký của đợt');
  return { ok: !!ok };
}
function apiAddRoundRegistration(p) {
  const st = bystId_(num_(p.studentId));
  const nid = nextId_('DangKyThi');
  appendObj_('DangKyThi', { ID: nid, IDDot: num_(p.roundId), IDHocVien: num_(p.studentId),
    MaHV: st ? st.code : '', HoTen: st ? st.name : '', SBD: p.sbd || 'HP' + String(num_(p.studentId)).padStart(3, '0') });
  logHistory_('addRoundRegistration', st ? st.code : '', 'Đăng ký thi đợt ' + p.roundId);
  return { ok: true, id: nid };
}
function apiRemoveRoundRegistration(p) {
  const ok = deleteRow_('DangKyThi', num_(p.id));
  logHistory_('removeRoundRegistration', 'DK' + p.id, 'Hủy đăng ký thi');
  return { ok: !!ok };
}
function apiUpdateExamResult(p) {
  let id = num_(p.examId);
  if (!id) {
    const ex = tab_('KetQuaThi').objs.find(o => num_(o['IDHocVien']) === num_(p.studentId));
    if (ex) id = num_(ex['ID']);
    else {
      const st = bystId_(num_(p.studentId));
      id = nextId_('KetQuaThi');
      appendObj_('KetQuaThi', { ID: id, IDHocVien: num_(p.studentId), MaHV: st ? st.code : '', HoTen: st ? st.name : '',
        MaKhoa: st ? st.courseCode : '', NgayThi: p.examDate || today_(), SBD: p.sbd || '', LyThuyet: 'Chưa thi', DiemLT: '',
        SaHinh: 'Chưa thi', DiemSH: '', DuongTruong: 'Chưa thi', DiemDT: '', KetQua: 'Chưa thi',
        PhiThiLai: 0, GhiChuPhi: '', GhiChu: '', NgayCoBang: '', CoBang: 'Chưa' });
    }
  }
  const patch = {};
  if (p.examDate !== undefined) patch['NgayThi'] = p.examDate;
  if (p.sbd !== undefined) patch['SBD'] = p.sbd;
  if (p.theory !== undefined) patch['LyThuyet'] = p.theory;
  if (p.theoryScore !== undefined) patch['DiemLT'] = p.theoryScore === null ? '' : num_(p.theoryScore);
  if (p.sahinh !== undefined) patch['SaHinh'] = p.sahinh;
  if (p.sahinhScore !== undefined) patch['DiemSH'] = p.sahinhScore === null ? '' : num_(p.sahinhScore);
  if (p.road !== undefined) patch['DuongTruong'] = p.road;
  if (p.roadScore !== undefined) patch['DiemDT'] = p.roadScore === null ? '' : num_(p.roadScore);
  if (p.passed !== undefined) patch['KetQua'] = p.passed ? 'Đậu' : 'Rớt';
  if (p.retakeFee !== undefined) patch['PhiThiLai'] = num_(p.retakeFee);
  if (p.retakeFeeNote !== undefined) patch['GhiChuPhi'] = p.retakeFeeNote;
  if (p.note !== undefined) patch['GhiChu'] = p.note;
  writeRow_('KetQuaThi', id, patch);
  logHistory_('updateExamResult', 'KQ' + id, 'Cập nhật kết quả thi');
  return { ok: true, id };
}
function apiUpdateGraduation(p) {
  const st = bystId_(num_(p.studentId));
  const t = tab_('TotNghiep');
  const o = t.objs.find(x => num_(x['IDHocVien']) === num_(p.studentId));
  const vals = { NgayThiTN: p.examDate !== undefined ? p.examDate : (o ? dstr_(o['NgayThiTN']) : ''),
    KetQua: p.result !== undefined ? p.result : (o ? str_(o['KetQua']) : ''),
    SoChungChi: p.certNo !== undefined ? p.certNo : (o ? str_(o['SoChungChi']) : ''),
    NgayCap: p.certDate !== undefined ? p.certDate : (o ? dstr_(o['NgayCap']) : '') };
  if (o) {
    const row = t.headers.map(h => vals[h] !== undefined ? vals[h] : o[h]);
    t.sheet.getRange(o._row, 1, 1, t.headers.length).setValues([row]);
    dropTab_('TotNghiep');
  } else {
    t.sheet.appendRow([num_(p.studentId), st ? st.code : '', st ? st.name : '', st ? st.courseCode : '',
      vals.NgayThiTN, vals.KetQua, vals.SoChungChi, vals.NgayCap]);
    dropTab_('TotNghiep');
  }
  writeRow_('HocVien', num_(p.studentId), { NgayThiTotNghiep: vals.NgayThiTN, KetQuaTotNghiep: vals.KetQua,
    SoChungChi: vals.SoChungChi, NgayCapChungChi: vals.NgayCap });
  logHistory_('updateGraduation', st ? st.code : '', 'Cập nhật thi tốt nghiệp: ' + (p.result || ''));
  return { ok: true };
}
function apiUpdateLicenseStatus(p) {
  const st = findRow_('HocVien', num_(p.studentId));
  if (!st) return { ok: false, error: 'Không tìm thấy học viên' };
  const patch = { CoBang: p.hasLicense ? 'Có' : 'Chưa', NgayCoBang: p.licenseDate || '' };
  if (p.licenseScore !== undefined) patch['DiemBangLai'] = p.licenseScore === null ? '' : num_(p.licenseScore);
  writeRow_('HocVien', num_(p.studentId), patch);
  const ex = tab_('KetQuaThi').objs.find(o => num_(o['IDHocVien']) === num_(p.studentId));
  if (ex) writeRow_('KetQuaThi', num_(ex['ID']), { CoBang: p.hasLicense ? 'Có' : 'Chưa', NgayCoBang: p.licenseDate || '' });
  logHistory_('updateLicenseStatus', str_(st['MaHV']), 'Cập nhật cấp bằng: ' + (p.hasLicense ? 'Đã có bằng ' + (p.licenseDate || '') : 'Chưa có bằng'));
  return { ok: true };
}
function apiSendNotification(p) {
  const list = JSON.parse(cfg_('ThongBaoMau', '[]'));
  list.unshift({ date: today_(), title: p.title || 'Thông báo', message: p.message || '', sent: true });
  const t = tab_('CauHinh');
  const o = t.objs.find(x => str_(x['Khoa']) === 'ThongBaoMau');
  if (o) {
    const row = t.headers.map(h => h === 'GiaTri' ? JSON.stringify(list) : o[h]);
    t.sheet.getRange(o._row, 1, 1, t.headers.length).setValues([row]);
    dropTab_('CauHinh');
  }
  logHistory_('sendNotification', '', 'Gửi thông báo: ' + (p.title || ''));
  return { ok: true };
}
function apiEnrollAdmission(p) {
  const ao = tab_('TuyenSinh').objs.find(o => num_(o['ID']) === num_(p.admissionId));
  if (!ao) return { ok: false, error: 'Không tìm thấy hồ sơ tuyển sinh' };
  const adm = admissionObj_(ao);
  if (adm.docsStatus !== 'Đủ giấy tờ') return { ok: false, error: 'Hồ sơ chưa đủ giấy tờ' };
  const c = courseMap_()[p.courseCode];
  if (!c) return { ok: false, error: 'Không tìm thấy khóa học' };
  const hvObjsE = tab_('HocVien').objs;
  const maxSachE = hvObjsE.reduce((a, o) => Math.max(a, Number(String(o['MaHVSach']).replace(/\D+/g, '')) || 0), 0);
  const maxCodeE = hvObjsE.reduce((a, o) => Math.max(a, Number(String(o['MaHV']).replace(/\D+/g, '')) || 0), 0);
  const nid = nextId_('HocVien');
  const code = 'HV' + String(maxCodeE + 1).padStart(3, '0');
  appendObj_('HocVien', {
    ID: nid, MaHV: code, MaHVSach: String(maxSachE + 1), HoTen: adm.name, NgaySinh: '',
    CCCD: '', SDT: adm.phone, DiaChi: '', MaKhoa: c.code, Hang: c.licenseClass,
    TrungTamLienKet: adm.source && adm.source !== 'Quầy tuyển sinh' ? adm.source : '',
    GVChinh: '', XeQuen: '', BienSoKG: '', NgayDangKy: today_(), NgayKhamSK: adm.healthExamDate || '',
    HanKhamSK: adm.healthExpiry || '', GioCabin: 0, BuoiCabin: 0, DiemLyThuyet: '', DatLyThuyet: 'Chưa đạt',
    DaHocSaHinh: 'Chưa học', KmDATDaChay: 0, KmDATMucTieu: TARGET_KM[c.licenseClass] || 810, MocKmDAT: 0,
    GiamTien: 0, MuaLL: '', HocPhi: priceFor_(c.licenseClass), LyDoGiam: '', TrangThai: 'Đang học', NgayThiTotNghiep: '', KetQuaTotNghiep: '',
    SoChungChi: '', NgayCapChungChi: '', NgayCoBang: '', CoBang: 'Chưa', DiemBangLai: ''
  });
  deleteRow_('TuyenSinh', num_(p.admissionId));
  logHistory_('enrollAdmission', code, 'Nhập học từ hồ sơ ' + adm.expectedCode + ' → ' + code + ' vào khóa ' + c.code);
  return { ok: true, id: nid, code, name: adm.name, courseCode: c.code };
}
// ---------- Bang gia hoc phi theo hang (tab CauHinh, khoa HocPhi_*) ----------
const PRICE_KEYS = { 'B1': 'HocPhi_B1', 'B2': 'HocPhi_B2', 'C1': 'HocPhi_C1', 'C': 'HocPhi_C', 'Nâng dấu': 'HocPhi_NangDau' };
function priceFor_(cls) {
  const k = PRICE_KEYS[cls];
  const v = k ? num_(cfg_(k, 0)) : 0;
  return v > 0 ? v : (TUITION[cls] || 0);
}
function setCfg_(key, value) {
  const t = tab_('CauHinh');
  const row = t.objs.find(o => String(o['Khoa']) === key);
  if (row) t.sheet.getRange(row._row, 2).setValue(value);
  else t.sheet.appendRow([key, value]);
  invalidateDerived_();
}
function apiGetCenterInfo() {
  return {
    name: String(cfg_('TenTrungTam', '') || ''),
    address: String(cfg_('DiaChi', '') || ''),
    phone: String(cfg_('SDT', '') || ''),
    logo: String(cfg_('Logo', '') || ''),
    logoMark: String(cfg_('LogoMark', '') || ''),
  };
}
function apiSaveCenterInfo(p) {
  const name = String(p.name || '').trim();
  if (!name) throw new Error('Thiếu tên trung tâm');
  setCfg_('TenTrungTam', name);
  setCfg_('DiaChi', String(p.address || '').trim());
  setCfg_('SDT', String(p.phone || '').trim());
  if (p.logo !== undefined && p.logo !== null) {
    const lg = String(p.logo || '');
    if (lg.length > 49000) throw new Error('Ảnh logo quá lớn — chọn ảnh khác nhẹ hơn');
    setCfg_('Logo', lg);
  }
  if (p.logoMark !== undefined && p.logoMark !== null) {
    const lm = String(p.logoMark || '');
    if (lm.length > 49000) throw new Error('Ảnh logo quá lớn — chọn ảnh khác nhẹ hơn');
    setCfg_('LogoMark', lm);
  }
  logHistory_('saveCenterInfo', '', 'Cập nhật thông tin trung tâm: ' + name);
  return { ok: true };
}
function apiGetPrices() {
  const prices = {};
  Object.keys(PRICE_KEYS).forEach(c => { prices[c] = priceFor_(c); });
  return { prices };
}
function apiSavePrices(p) {
  const prices = p.prices || {};
  const t = tab_('CauHinh');
  Object.keys(PRICE_KEYS).forEach(cls => {
    if (prices[cls] === undefined) return;
    const key = PRICE_KEYS[cls];
    const row = t.objs.find(o => String(o['Khoa']) === key);
    if (row) t.sheet.getRange(row._row, 2).setValue(num_(prices[cls]));
    else t.sheet.appendRow([key, num_(prices[cls])]);
  });
  logHistory_('savePrices', '', 'Cập nhật bảng giá học phí theo hạng');
  return { ok: true };
}
function apiCheckCccd(p) {
  const cccd = String(p.cccd || '');
  const found = tab_('HocVien').objs.find(o => str_(o['CCCD']) === cccd);
  if (found) return { found: true, student: { id: num_(found['ID']), code: str_(found['MaHV']), name: str_(found['HoTen']), courseCode: str_(found['MaKhoa']), status: str_(found['TrangThai']) } };
  return { found: false };
}

// ---------- Router ----------
const READ_ACTIONS = {
  getdashboard: apiGetDashboard, getdatsummary: apiGetDatSummary,
  liststudents: apiListStudents, getstudent: apiGetStudent,
  listschedules: apiListSchedules, listschedulesrange: apiListSchedulesRange,
  listconflicts: apiListConflicts, listpayments: apiListPayments,
  listcenters: apiListCenters, getfleet: apiGetFleet, listfuellogs: apiListFuelLogs,
  listadmissions: apiListAdmissions, listexamrounds: apiListExamRounds,
  listexams: apiListExams, getreports: apiGetReports, getgovreports: apiGetGovReports,
  listnotifications: apiNotifications, getlookups: apiGetLookups, getprices: apiGetPrices,
  getcenterinfo: apiGetCenterInfo
};
const WRITE_ACTIONS = {
  savestudent: apiSaveStudent, setstudentstatus: apiSetStudentStatus,
  transferstudentcourse: apiTransferCourse, createschedule: apiCreateSchedule,
  updateschedule: apiUpdateSchedule, deleteschedule: apiDeleteSchedule,
  addpayment: apiAddPayment, saveteacher: apiSaveTeacher, savevehicle: apiSaveVehicle,
  savecenter: apiSaveCenter, deletecenter: apiDeleteCenter, saveprices: apiSavePrices,
  savecenterinfo: apiSaveCenterInfo,
  savebotuc: apiSaveBotuc, deletebotuc: apiDeleteBotuc,
  addfuellog: apiAddFuelLog, deletefuellog: apiDeleteFuelLog,
  saveadmission: apiSaveAdmission, deleteadmission: apiDeleteAdmission,
  setadmissiondoc: apiSetAdmissionDoc, saveexamround: apiSaveExamRound,
  deleteexamround: apiDeleteExamRound, addroundregistration: apiAddRoundRegistration,
  removeroundregistration: apiRemoveRoundRegistration, updateexamresult: apiUpdateExamResult,
  updategraduation: apiUpdateGraduation, updatelicensestatus: apiUpdateLicenseStatus,
  sendnotification: apiSendNotification, enrolladmission: apiEnrollAdmission,
  checkcccd: apiCheckCccd
};
// ---------- CacheService cho MỌI action đọc (v13: thế hệ THEO NHÓM + chia khối) ----------
// Khóa cache: hp3:<gen của nhóm>:<action>:<tham số đã sắp xếp theo tên khóa>.
// Gen theo nhóm xem READ_GROUP/WRITE_GROUPS bên dưới. Điều kiện tăng gen giữ
// nguyên: chỉ khi ghi thành công (checkcccd không trả ok:true nên không tăng
// thừa). Đây là ĐIỂM TĂNG GEN DUY NHẤT: mọi action ghi đều đi qua route_().
// TTL 1800s; trigger hâm nóng mỗi 10 phút (warmCacheHp) nạp lại cache trước
// khi hết hạn nên người dùng thật gần như không bao giờ gặp cache nguội.
// Payload < 95.000 ký tự cache nguyên khối; lớn hơn thì CHIA KHỐI 60K
// (cachePut_/cacheGet_) nên lịch tháng đầy buổi vẫn cache được. Không cache
// kết quả lỗi (res.ok === false). Mọi thao tác CacheService đều bọc
// try/catch: lỗi cache không bao giờ làm hỏng response.
const READ_CACHE_TTL = 1800; // giây
const READ_CACHE_MAXLEN = 95000; // ký tự JSON tối đa được phép cache nguyên khối
const GEN_TTL = 21600; // 6 giờ = tối đa CacheService, luôn >> TTL dữ liệu
// (1800s): nếu gen tự hết hạn và quay về '1' thì mọi khóa dữ liệu của thế hệ
// '1' cũ cũng đã hết hạn từ lâu -> không có chuyện dữ liệu cũ "sống lại".
// ---------- v13: thế hệ cache THEO NHÓM dữ liệu ----------
// Trước đây MỌI lượt ghi thành công đều tăng 1 gen chung -> toàn bộ cache đọc
// bị bỏ phí, lần mở mục nào sau khi lưu cũng phải tính lại từ đầu (lạnh 2–6s).
// Nay mỗi action đọc thuộc 1 nhóm; khóa cache chứa gen CỦA NHÓM ĐÓ. Action ghi
// chỉ tăng gen các nhóm nó thật sự làm thay đổi (WRITE_GROUPS) — ghi không có
// trong bảng thì tăng TẤT CẢ cho an toàn. Nhóm khác giữ nguyên cache nóng.
const READ_GROUP = {
  getdashboard: 'dash', getreports: 'reports', getgovreports: 'reports',
  liststudents: 'students', getstudent: 'students',
  listschedules: 'sched', listschedulesrange: 'sched', listconflicts: 'sched',
  getdatsummary: 'dat', getfleet: 'fleet', listfuellogs: 'fleet',
  listexams: 'exams', listexamrounds: 'exams', listadmissions: 'admissions',
  listpayments: 'tuition', getprices: 'tuition', listcenters: 'centers',
  listnotifications: 'alerts', getcenterinfo: 'config', getlookups: 'lookups'
};
const ALL_GROUPS = ['dash', 'reports', 'students', 'sched', 'dat', 'fleet', 'exams', 'admissions', 'tuition', 'centers', 'alerts', 'config', 'lookups'];
const WRITE_GROUPS = {
  savestudent: ['students', 'lookups', 'dash', 'exams', 'dat', 'reports'],
  setstudentstatus: ['students', 'lookups', 'dash', 'exams', 'dat', 'reports'],
  transferstudentcourse: ['students', 'lookups', 'dash', 'exams', 'dat'],
  addpayment: ['tuition', 'students', 'dash', 'reports', 'exams'],
  createschedule: ['sched', 'dash', 'dat', 'alerts'],
  updateschedule: ['sched', 'dash', 'dat', 'alerts'],
  deleteschedule: ['sched', 'dash', 'dat', 'alerts'],
  saveexamround: ['exams'], deleteexamround: ['exams'],
  addroundregistration: ['exams'], removeroundregistration: ['exams'],
  updateexamresult: ['exams', 'dash', 'reports'],
  updategraduation: ['exams', 'students', 'lookups', 'dash'],
  updatelicensestatus: ['exams', 'students', 'lookups'],
  saveadmission: ['admissions', 'alerts'], deleteadmission: ['admissions', 'alerts'],
  setadmissiondoc: ['admissions'],
  enrolladmission: ['admissions', 'students', 'lookups', 'dash', 'exams', 'dat', 'alerts'],
  sendnotification: ['alerts'],
  saveteacher: ['fleet', 'lookups', 'sched'], savevehicle: ['fleet', 'sched'],
  savecenter: ['centers', 'students', 'lookups'], deletecenter: ['centers', 'students', 'lookups'],
  savebotuc: ['fleet', 'reports'], deletebotuc: ['fleet', 'reports'],
  addfuellog: ['fleet', 'reports'], deletefuellog: ['fleet', 'reports'],
  saveprices: ['tuition', 'students', 'dash', 'reports'],
  savecenterinfo: ['config']
};
function groupGenKey_(g) { return 'hp3:gen:' + g; }
function cacheGen_(g) {
  try {
    const n = parseInt(CacheService.getScriptCache().get(groupGenKey_(g || 'misc')), 10);
    return isNaN(n) || n < 1 ? '1' : String(n);
  } catch (e) { return '1'; }
}
function bumpGroups_(groups) {
  try {
    const c = CacheService.getScriptCache();
    (groups && groups.length ? groups : ALL_GROUPS).forEach(g => {
      c.put(groupGenKey_(g), String(parseInt(cacheGen_(g), 10) + 1), GEN_TTL);
    });
  } catch (e) { /* cache lỗi không được chặn luồng chính */ }
}
function bumpCacheGen_() { bumpGroups_(ALL_GROUPS); } // giữ tương thích nếu còn nơi gọi cũ
function readCacheKey_(a, params) {
  const parts = Object.keys(params || {}).sort().map(k => {
    const v = params[k];
    return k + '=' + (v !== null && typeof v === 'object' ? JSON.stringify(v) : String(v));
  });
  return 'hp3:' + cacheGen_(READ_GROUP[a] || 'misc') + ':' + a + ':' + parts.join('&');
}
// ---------- v13: cache CHIA KHỐI cho payload lớn ----------
// Payload >= 95K ký tự trước đây KHÔNG được cache lần nào (vd lịch tháng khi
// đầy buổi ~174K) -> luôn tính lại. Nay cắt khối 60K ký tự (an toàn giới hạn
// 100KB/khóa kể cả khi tiếng Việt tốn nhiều byte), tối đa 8 khối (~480K).
const READ_CACHE_CHUNK = 60000;
function cacheGet_(key) {
  try {
    const c = CacheService.getScriptCache();
    const direct = c.get(key);
    if (direct) return direct;
    const n = parseInt(c.get(key + '#n'), 10);
    if (isNaN(n) || n < 2 || n > 8) return null;
    let out = '';
    for (let i = 0; i < n; i++) {
      const part = c.get(key + '#' + i);
      if (part === null || part === undefined) return null;
      out += part;
    }
    return out;
  } catch (e) { return null; }
}
function cachePut_(key, str) {
  try {
    const c = CacheService.getScriptCache();
    if (str.length < READ_CACHE_MAXLEN) { c.put(key, str, READ_CACHE_TTL); return; }
    const n = Math.ceil(str.length / READ_CACHE_CHUNK);
    if (n > 8) return;
    for (let i = 0; i < n; i++) c.put(key + '#' + i, str.slice(i * READ_CACHE_CHUNK, (i + 1) * READ_CACHE_CHUNK), READ_CACHE_TTL);
    c.put(key + '#n', String(n), READ_CACHE_TTL);
  } catch (e) { /* quá giới hạn thì thôi */ }
}
function route_(action, params) {
  const a = String(action || '').toLowerCase();
  if (READ_ACTIONS[a]) {
    const p = params || {};
    const key = readCacheKey_(a, p);
    const keyOk = key.length <= 240; // giới hạn khóa CacheService: 250 ký tự
    if (keyOk) {
      const hit = cacheGet_(key);
      if (hit) { try { return JSON.parse(hit); } catch (e) { /* JSON hỏng thì tính mới */ } }
    }
    const res = READ_ACTIONS[a](p);
    if (keyOk && !(res && res.ok === false)) {
      try {
        const str = JSON.stringify(res);
        if (str) cachePut_(key, str);
      } catch (e) { /* quá giới hạn thì thôi */ }
    }
    return res;
  }
  if (WRITE_ACTIONS[a]) {
    // Khóa ghi toàn cục: 2 lượt ghi đồng thời có thể sinh trùng ID (nextId_ tính
    // theo bản đọc trong bộ nhớ) hoặc ghi đè lẫn nhau — khóa lại, chờ tối đa 30s.
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) return { ok: false, error: 'Hệ thống đang xử lý một thao tác lưu khác — vui lòng thử lại sau vài giây' };
    try {
      const res = WRITE_ACTIONS[a](params || {});
      if (res && res.ok === true) bumpGroups_(WRITE_GROUPS[a]); // v13: chỉ tăng gen các nhóm liên quan (không có trong bảng -> tăng tất cả)
      return res;
    } finally { lock.releaseLock(); }
  }
  return { ok: false, error: 'Action không hợp lệ: ' + action };
}
function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const params = {};
    Object.keys(p).forEach(k => {
      if (k === 'action' || k === 'callback') return;
      let v = p[k];
      try { v = JSON.parse(v); } catch (err) {}
      params[k] = v;
    });
    const res = route_(p.action || 'getdashboard', params);
    // Chỉ nhận callback dạng tên hàm JS hợp lệ — chặn chèn mã qua tham số callback.
    const cbOk = typeof p.callback === 'string' && /^[A-Za-z_$][A-Za-z0-9_$.]*$/.test(p.callback);
    const body = cbOk ? p.callback + '(' + JSON.stringify(res) + ')' : JSON.stringify(res);
    return ContentService.createTextOutput(body).setMimeType(cbOk ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}
function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    return json_(route_(body.action, body.params || body));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}
// ---------- Hâm nóng cache (warmCacheHp) ----------
// Nạp sẵn cache cho các action đọc hay dùng bằng cách đi qua chính route_()
// -> khóa cache sinh ra khớp 100% với request thật của người dùng. Mỗi
// action bọc try/catch riêng: 1 action lỗi không chặn các action sau.
// listschedules chỉ hâm ngày hôm nay theo today_() (ngày hệ thống backend),
// listschedulesrange hâm TUẦN hiện tại (T2-CN) + dải THÁNG của chế độ Tháng;
// listpayments hâm tháng hiện tại.
function isoUtc_(ms) {
  const d = new Date(ms);
  const p2 = n => String(n).padStart(2, '0');
  return d.getUTCFullYear() + '-' + p2(d.getUTCMonth() + 1) + '-' + p2(d.getUTCDate());
}
function weekRangeOf_(iso) {
  const p = String(iso).slice(0, 10).split('-').map(Number);
  const ms = Date.UTC(p[0], (p[1] || 1) - 1, p[2] || 1);
  const dow = (new Date(ms).getUTCDay() + 6) % 7; // 0 = Thứ 2
  const mon = ms - dow * 86400000;
  return { from: isoUtc_(mon), to: isoUtc_(mon + 6 * 86400000) };
}
function warmOne_(action, params) {
  try { route_(action, params || {}); } catch (e) { /* bỏ qua, hâm tiếp action sau */ }
}
function warmCacheHp() {
  const T = today_();
  const wk = weekRangeOf_(T);
  warmOne_('getdashboard');
  warmOne_('getreports');
  warmOne_('getgovreports');
  warmOne_('getlookups');
  warmOne_('listcenters');
  warmOne_('listexamrounds');
  warmOne_('getfleet');
  warmOne_('listnotifications');
  warmOne_('getdatsummary');
  warmOne_('liststudents', { onlyEligible: false, page: 1, pageSize: 10 }); // đúng bộ tham số frontend gửi lần đầu (hâm {} là lệch khóa, không ai hưởng)
  warmOne_('liststudents', { page: 1, pageSize: 20 }); // combo của mục Học phí (v13 thêm: trước đây combo này nguội 4s)
  warmOne_('getprices');
  warmOne_('listexams');
  warmOne_('listadmissions');
  warmOne_('listconflicts');
  warmOne_('listfuellogs');
  warmOne_('listschedules', { date: T });
  warmOne_('listschedulesrange', { from: wk.from, to: wk.to });
  // Hâm dải THÁNG đúng khóa frontend dùng ở chế độ Tháng (Thứ 2 của tuần chứa
  // ngày 1 → +41 ngày). Từ v13 payload lớn được cache chia khối nên tháng cao
  // điểm (payload > 95K) vẫn hâm được bình thường.
  const mMon = weekRangeOf_(T.slice(0, 7) + '-01');
  const mp = mMon.from.split('-').map(Number);
  const mTo = isoUtc_(Date.UTC(mp[0], mp[1] - 1, mp[2]) + 41 * 86400000);
  warmOne_('listschedulesrange', { from: mMon.from, to: mTo });
  warmOne_('listpayments', { month: T.slice(0, 7) });
}
function caiTriggerWarm() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'warmCacheHp')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('warmCacheHp').timeBased().everyMinutes(10).create();
}
// ---------- Backup tự động 2h sáng ----------
function autoBackup() {
  const src = ss_();
  const copy = src.copy('Backup_' + src.getName() + '_' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd_HHmm'));
  const folders = DriveApp.getFoldersByName('Backup QuanLy HongPhuc');
  const f = folders.hasNext() ? folders.next() : DriveApp.createFolder('Backup QuanLy HongPhuc');
  DriveApp.getFileById(copy.getId()).moveTo(f);
  const files = [];
  const it = f.getFiles();
  while (it.hasNext()) files.push(it.next());
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  files.slice(14).forEach(x => x.setTrashed(true));
  logHistory_('autoBackup', '', 'Sao lưu Sheet: ' + copy.getId());
  return copy.getId();
}
function caiTriggerBackup() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'autoBackup')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('autoBackup').timeBased().everyDays(1).atHour(2).create();
}
