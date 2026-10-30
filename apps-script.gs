/**
 * GIAO THƯƠNG – Google Apps Script (miễn phí)
 * Nhận dữ liệu từ app: lưu vào Google Sheet + gửi email về HAMEE.
 *
 * CÀI ĐẶT:
 * 1. Tạo Google Sheet mới, 2 tab: "TinDang" và "KetNoi".
 *    Tab TinDang dòng 1:  loai | ho_ten | cong_ty | tieu_de | mo_ta | nganh | khu_vuc | sdt | ngay_dang | logo | trang_thai | email | website | linh_vuc | san_pham
 *    Tab KetNoi  dòng 1:  thoi_gian | ho_ten | cong_ty | sdt | email | website | linh_vuc | san_pham | loi_nhan | tin_quan_tam | loai | cong_ty_dang | nguoi_dang | sdt_nguoi_dang
 * 2. Tiện ích mở rộng → Apps Script → dán toàn bộ file này → sửa EMAIL_HAMEE.
 * 3. Triển khai → Triển khai mới → Loại: Ứng dụng web → Người truy cập: "Bất kỳ ai" → Triển khai.
 *    Copy link Web App dán vào SCRIPT_URL trong index.html.
 * 4. Tab TinDang: Tệp → Chia sẻ → Công bố lên web → chọn tab TinDang, định dạng CSV → copy link vào SHEET_CSV.
 *
 * DUYỆT TIN: tin hội viên gửi có trang_thai = "cho_duyet". Admin đổi thành "duyet" là tin hiện trên app.
 * ADMIN TỰ ĐĂNG: nhập thẳng 1 dòng vào tab TinDang, trang_thai = "duyet".
 * HỘI VIÊN: tab "HoiVien" (tự tạo khi nhập dữ liệu lần đầu). App đọc qua link Web App ?action=hv.
 * TRANG QUẢN TRỊ: admin.html — đăng nhập bằng ADMIN_KEY bên dưới (đổi thành mật khẩu riêng, không chia sẻ).
 */
const ADMIN_KEY = "DOI_MAT_KHAU_O_DAY"; // ← mật khẩu trang quản trị
const EMAIL_HAMEE = "hoicokhidien@gmail.com"; // ← email nhận thông báo (nhiều email cách nhau dấu phẩy)

const HV_HEAD = ["ten", "nhom", "dia_chi", "sdt", "email", "website", "logo", "url_hamee", "gioi_thieu", "san_pham", "trang_thai"];

function doGet(e) {
  if (e.parameter.action === "hv") {
    const sh = SpreadsheetApp.getActive().getSheetByName("HoiVien");
    const L = sh ? rows(sh).filter(r => r.trang_thai !== "an").map(r => { delete r._row; return r; }) : [];
    return json({ ok: true, hv: L });
  }
  return ContentService.createTextOutput("ok");
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const ss = SpreadsheetApp.getActive();
  const now = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm");
  const ngay = now.slice(0, 10);

  if (d.action === "dang_tin") {
    ss.getSheetByName("TinDang").appendRow([d.loai, d.ten, d.cty, d.tieu_de, d.mo_ta, d.nganh, d.khu_vuc, "'" + d.sdt, ngay, "", "cho_duyet", d.email, d.web, d.lv, d.sp]);
    MailApp.sendEmail({
      to: EMAIL_HAMEE,
      subject: `[GIAO THƯƠNG] Tin chờ duyệt – ${d.loai === "mua" ? "CẦN MUA" : "CẦN BÁN"}: ${d.tieu_de}`,
      body: `Có tin mới cần duyệt (${now})\n\nLoại: ${d.loai === "mua" ? "Nhu cầu mua" : "Nhu cầu bán"}\nTiêu đề: ${d.tieu_de}\nMô tả: ${d.mo_ta}\nNgành: ${d.nganh}\nKhu vực: ${d.khu_vuc}\n\nNgười gửi: ${d.ten}\nDoanh nghiệp: ${d.cty}\nSĐT/Zalo: ${d.sdt}\nEmail: ${d.email}\nWebsite: ${d.web}\nNgành nghề: ${d.lv}\nSản phẩm: ${d.sp}\n\n→ Duyệt: mở Google Sheet, tab TinDang, đổi cột trang_thai thành "duyet".\n${ss.getUrl()}`
    });
  }

  if (d.action === "ket_noi") {
    ss.getSheetByName("KetNoi").appendRow([now, d.ten, d.cty, "'" + d.sdt, d.email, d.web, d.lv, d.sp, d.loi_nhan, d.tin, d.loai, d.cong_ty_dang, d.nguoi_dang, "'" + d.sdt_nguoi_dang]);
    MailApp.sendEmail({
      to: EMAIL_HAMEE,
      subject: `[GIAO THƯƠNG] Yêu cầu kết nối: ${d.tin}`,
      body: `Có hội viên muốn kết nối (${now})\n\nNgười liên hệ: ${d.ten}\nDoanh nghiệp: ${d.cty}\nSĐT/Zalo: ${d.sdt}\nEmail: ${d.email}\nWebsite: ${d.web}\nNgành nghề: ${d.lv}\nSản phẩm: ${d.sp}\nLời nhắn: ${d.loi_nhan}\n\nQuan tâm tin: ${d.tin} (${d.loai === "mua" ? "Cần mua" : "Cần bán"})\nBên đăng: ${d.nguoi_dang} – ${d.cong_ty_dang} – ${d.sdt_nguoi_dang}`
    });
  }
  if (String(d.action || "").indexOf("admin_") === 0) return json(admin(d, ss, ngay));
  return ContentService.createTextOutput("ok");
}

/* ===================== QUẢN TRỊ ===================== */
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function rows(sh) {
  const v = sh.getDataRange().getDisplayValues(), h = v.shift();
  return v.map((r, i) => { const o = { _row: i + 2 }; h.forEach((k, j) => o[k] = r[j]); return o; });
}
function admin(d, ss, ngay) {
  if (!ADMIN_KEY || ADMIN_KEY === "DOI_MAT_KHAU_O_DAY" || d.key !== ADMIN_KEY) return { ok: false, error: "Sai mật khẩu" };
  if (d.action === "admin_list") {
    const hv = ss.getSheetByName("HoiVien");
    return { ok: true, tin: rows(ss.getSheetByName("TinDang")), ketnoi: rows(ss.getSheetByName("KetNoi")), hv: hv ? rows(hv) : [] };
  }
  if (d.action === "admin_hv_import") {
    let sh = ss.getSheetByName("HoiVien");
    if (!sh) sh = ss.insertSheet("HoiVien");
    if (d.replace) sh.clear();
    if (sh.getLastRow() === 0) { sh.appendRow(HV_HEAD); sh.getRange(1, 1, 1, HV_HEAD.length).setFontWeight("bold").setBackground("#1565d8").setFontColor("#ffffff"); sh.setFrozenRows(1); }
    const data = (d.rows || []).map(r => HV_HEAD.map(k => k === "sdt" && r[k] ? "'" + r[k] : (r[k] || "")));
    if (data.length) sh.getRange(sh.getLastRow() + 1, 1, data.length, HV_HEAD.length).setValues(data);
    return { ok: true, added: data.length };
  }
  const name = d.sheet === "HoiVien" ? "HoiVien" : "TinDang", key = name === "HoiVien" ? "ten" : "tieu_de";
  const sh = ss.getSheetByName(name);
  if (!sh) return { ok: false, error: "Chưa có tab " + name };
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  // kiểm tra đúng dòng trước khi sửa/xoá (tránh lệch dòng khi sheet vừa thay đổi)
  const check = r => r >= 2 && r <= sh.getLastRow() && sh.getRange(r, head.indexOf(key) + 1).getDisplayValue() === d.tieu_de_cu;
  if (d.action === "admin_status") {
    if (!check(d.row)) return { ok: false, error: "Dữ liệu đã thay đổi, tải lại trang" };
    sh.getRange(d.row, head.indexOf("trang_thai") + 1).setValue(d.trang_thai);
    return { ok: true };
  }
  if (d.action === "admin_save") {
    const f = d.data || {};
    if (f.sdt) f.sdt = "'" + f.sdt;
    if (d.row) {
      if (!check(d.row)) return { ok: false, error: "Dữ liệu đã thay đổi, tải lại trang" };
      Object.keys(f).forEach(k => { const c = head.indexOf(k); if (c >= 0) sh.getRange(d.row, c + 1).setValue(f[k]); });
    } else {
      if (name === "TinDang") { f.ngay_dang = f.ngay_dang || ngay; f.trang_thai = f.trang_thai || "duyet"; }
      sh.appendRow(head.map(k => f[k] || ""));
    }
    return { ok: true };
  }
  if (d.action === "admin_delete") {
    if (!check(d.row)) return { ok: false, error: "Dữ liệu đã thay đổi, tải lại trang" };
    sh.deleteRow(d.row);
    return { ok: true };
  }
  return { ok: false, error: "Không rõ thao tác" };
}
