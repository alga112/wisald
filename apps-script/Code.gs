/**
 * وصال — قاعدة بيانات Google Sheets
 * الصق هذا الملف في: Google Sheet ← الإضافات (Extensions) ← Apps Script
 * ثم: نشر (Deploy) ← نشر جديد ← تطبيق ويب، "التنفيذ بصفتي" و"من لديه حق الوصول: أي شخص".
 *
 * ينشئ أربع أوراق تلقائيًا: الجهات، الطلبات، سجل الوصول، الزوار.
 * عمود json في كل ورقة هو ما يقرأه الموقع — لا تعدّله يدويًا. بقية الأعمدة للقراءة فقط.
 */
const SH = {
  orgs: ["id", "النوع", "الاسم", "المدينة", "الجوال", "رقم الاتفاقية", "تاريخ التوقيع", "الموقّع", "كراسي متحركة", "أجهزة تنفس", "أسرّة طبية", "أجهزة قياس ضغط", "آخر تحديث", "json"],
  requests: ["id", "تاريخ الطلب", "المستشفى", "المدينة", "الصنف", "الكمية", "الأولوية", "الحالة", "الجمعية", "المسافة (كم)", "رقم التتبع", "المندوب", "آخر تحديث", "json"],
  audit: ["الوقت", "الجهة", "الإجراء", "json"],
  visitors: ["الوقت", "الاسم", "الجوال"]
};
const NAMES = { orgs: "الجهات", requests: "الطلبات", audit: "سجل الوصول", visitors: "الزوار" };
const ITEMS = ["كراسي متحركة", "أجهزة تنفس", "أسرّة طبية", "أجهزة قياس ضغط"];
const ST = ["بانتظار جهة مناسبة", "بانتظار موافقة الجمعية", "تجهيز الشحنة", "الشحنة في الطريق", "تم التسليم، بانتظار تأكيد المستشفى", "تم التحقق من الوصول"];
const ROLE = { h: "مستشفى", c: "جمعية", a: "جهة مشرفة" };

/** شغّل هذه الدالة مرة واحدة من المحرر لإنشاء الأوراق ومنح الصلاحيات */
function setup() { Object.keys(SH).forEach(sheet_); }

function sheet_(k) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(NAMES[k]);
  if (!s) {
    s = ss.insertSheet(NAMES[k]);
    s.appendRow(SH[k]);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, SH[k].length).setFontWeight("bold").setBackground("#0b5d4b").setFontColor("#ffffff");
    s.setRightToLeft(true);
  }
  return s;
}
const v_ = x => (x === null || x === undefined) ? "" : x;
const fdt_ = t => t ? Utilities.formatDate(new Date(t), "Asia/Riyadh", "yyyy-MM-dd HH:mm") : "";

function row_(k, id, d) {
  if (k === "orgs") {
    const s = d.stock || {}, c = d.contract || {};
    return [id, ROLE[d.role] || d.role, v_(d.name), v_(d.city), "'" + v_(d.phone), v_(c.no), fdt_(c.ts), v_(c.by),
      v_(s[0]), v_(s[1]), v_(s[2]), v_(s[3]), fdt_(Date.now()), JSON.stringify(d)];
  }
  return [id, fdt_(d.ts), v_(d.hname), v_(d.city), v_(ITEMS[d.item]), v_(d.qty), v_(d.pr), v_(ST[d.s]), v_(d.cname),
    v_(d.dist), v_(d.trk), v_(d.drv), fdt_(Date.now()), JSON.stringify(d)];
}
const auditRow_ = e => [fdt_(e.ts), v_(e.an || e.a), v_(e.t), JSON.stringify(e)];

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function readMap_(k) {
  const s = sheet_(k), n = s.getLastRow(), W = SH[k].length, m = {};
  if (n < 2) return m;
  s.getRange(2, 1, n - 1, W).getValues().forEach(r => {
    if (r[0] && r[W - 1]) { try { m[r[0]] = JSON.parse(r[W - 1]); } catch (x) {} }
  });
  return m;
}
function readAudit_() {
  const s = sheet_("audit"), n = s.getLastRow();
  if (n < 2) return [];
  const st = Math.max(2, n - 299);
  return s.getRange(st, 4, n - st + 1, 1).getValues()
    .map(r => { try { return JSON.parse(r[0]); } catch (x) { return null; } }).filter(Boolean);
}

function doGet(e) {
  if ((e.parameter.action || "") === "all") {
    return out_({ ok: true, orgs: readMap_("orgs"), requests: readMap_("requests"), audit: readAudit_() });
  }
  return out_({ ok: true, app: "wisal" });
}

function put_(k, id, d) {
  if (k !== "orgs" && k !== "requests") return;
  const s = sheet_(k), n = s.getLastRow(), row = row_(k, id, d);
  if (n > 1) {
    const ids = s.getRange(2, 1, n - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === id) { s.getRange(i + 2, 1, 1, row.length).setValues([row]); return; }
    }
  }
  s.appendRow(row);
}

function replaceAll_(k, rows) {
  const s = sheet_(k), n = s.getLastRow(), W = SH[k].length;
  if (n > 1) s.getRange(2, 1, n - 1, W).clearContent();
  if (rows.length) s.getRange(2, 1, rows.length, W).setValues(rows);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const type = (e.postData && e.postData.type) || "";
    // تسجيل زائر من صفحة الدخول (name, phone, ts)
    if (type.indexOf("application/x-www-form-urlencoded") === 0) {
      const p = e.parameter;
      sheet_("visitors").appendRow([fdt_(+p.ts || Date.now()), v_(p.name), "'" + v_(p.phone)]);
      return out_({ ok: true });
    }
    const b = JSON.parse(e.postData.contents);
    if (b.action === "put") put_(b.col, b.id, b.data);
    else if (b.action === "log") sheet_("audit").appendRow(auditRow_(b.e));
    else if (b.action === "bulk") {
      // إعادة تعبئة البيانات التجريبية (أول تشغيل، أو ?reset=1)
      replaceAll_("orgs", Object.keys(b.orgs || {}).map(id => row_("orgs", id, b.orgs[id])));
      replaceAll_("requests", Object.keys(b.requests || {}).map(id => row_("requests", id, b.requests[id])));
      replaceAll_("audit", (b.audit || []).map(auditRow_));
    }
    return out_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
