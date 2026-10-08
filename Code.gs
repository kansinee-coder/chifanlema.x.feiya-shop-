/**
 * Chifanlema — เชื่อมหน้าร้านกับ Google Sheet (Stock / Orders) + รับสลิป
 *  - GET  : ส่งยอดคงเหลือให้หน้าเว็บ
 *  - POST action "order" : ตรวจสต็อก แล้วบันทึกออเดอร์ลงแท็บ Orders (1 แถวต่อ 1 สินค้า) สถานะ "รอโอน"
 *  - POST action "slip"  : เก็บรูปสลิปในโฟลเดอร์ Drive "Chifanlema สลิป" ใส่ลิงก์ในแท็บ Orders และเปลี่ยนสถานะเป็น "รอตรวจสลิป"
 * วิธีติดตั้ง/อัปเดต: ดู SETUP.md
 */

const COL = {
  date: 1, id: 2, channel: 3, customer: 4, code: 5, /* F ชื่อสินค้า = สูตร */
  qty: 7, /* H ราคา, I ยอด = สูตร */ discount: 10, shipMethod: 11, shipFee: 12,
  status: 13, tracking: 14, note: 15,
  phone: 16, ig: 17, address: 18, total: 19, slip: 20   // P–T เพิ่มให้อัตโนมัติ
};
const EXTRA_HEADERS = ["เบอร์โทร", "IG", "ที่อยู่", "ยอดโอนรวม", "สลิป"];
const STATUSES = ["รอโอน", "รอตรวจสลิป", "ชำระแล้ว", "จัดส่งแล้ว", "ยกเลิก"];
const SLIP_FOLDER = "Chifanlema สลิป";
const MAX_SLIP_BYTES = 5 * 1024 * 1024;

function doGet() {
  return json({ ok: true, stock: readStock_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === "slip") return json(saveSlip_(d));
    return json(placeOrder_(d));
  } catch (err) {
    return json({ ok: false, reason: "error", message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// ---------- ออเดอร์ ----------
function placeOrder_(d) {
  const stock = readStock_();
  const short = [];
  d.items.forEach(it => {
    const s = stock[it.code];
    const left = s ? s.left : 0;
    if (it.qty > left) short.push({ code: it.code, left: Math.max(0, left) });
  });
  if (short.length) return { ok: false, reason: "stock", short: short, stock: stock };

  const sh = ordersSheet_();
  ensureSetup_(sh);
  let row = firstEmptyRow_(sh);
  d.items.forEach((it, i) => {
    set_(sh, row, COL.date, new Date());
    set_(sh, row, COL.id, d.id);
    set_(sh, row, COL.channel, "ลิงก์");
    set_(sh, row, COL.customer, d.name);
    set_(sh, row, COL.code, it.code);
    set_(sh, row, COL.qty, it.qty);
    set_(sh, row, COL.status, "รอโอน");
    if (i === 0) {
      set_(sh, row, COL.discount, d.discount || 0);
      set_(sh, row, COL.shipMethod, d.shipMethod);
      set_(sh, row, COL.shipFee, d.shipFee);
      set_(sh, row, COL.note, d.note || "");
      set_(sh, row, COL.phone, "'" + d.phone);
      set_(sh, row, COL.ig, d.ig || "");
      set_(sh, row, COL.address, d.address);
      set_(sh, row, COL.total, d.total);
    }
    row++;
  });
  SpreadsheetApp.flush();
  return { ok: true, id: d.id };
}

// ---------- สลิป ----------
function saveSlip_(d) {
  const sh = ordersSheet_();
  ensureSetup_(sh);
  const rows = rowsOfOrder_(sh, d.id);
  if (!rows.length) return { ok: false, reason: "notfound" };
  if (String(sh.getRange(rows[0], COL.slip).getValue())) return { ok: false, reason: "already" };
  if (!/^image\//.test(d.mime || "")) return { ok: false, reason: "type" };

  const bytes = Utilities.base64Decode(d.data);
  if (bytes.length > MAX_SLIP_BYTES) return { ok: false, reason: "size" };

  const file = slipFolder_().createFile(Utilities.newBlob(bytes, d.mime, d.id + ".jpg"));
  sh.getRange(rows[0], COL.slip).setFormula('=HYPERLINK("' + file.getUrl() + '","ดูสลิป")');
  rows.forEach(r => {
    if (sh.getRange(r, COL.status).getValue() === "รอโอน") sh.getRange(r, COL.status).setValue("รอตรวจสลิป");
  });
  SpreadsheetApp.flush();
  return { ok: true };
}

function slipFolder_() {
  const it = DriveApp.getFoldersByName(SLIP_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(SLIP_FOLDER);
}

// ---------- helpers ----------
function ordersSheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders"); }

function readStock_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Stock");
  const v = sh.getDataRange().getValues();
  const h = v[0].map(String);
  const iLeft = h.indexOf("คงเหลือ");
  const iStatus = h.indexOf("สถานะ");
  const out = {};
  for (let r = 1; r < v.length; r++) {
    const code = String(v[r][0]).trim();
    const status = String(v[r][iStatus] || "");
    if (!code || !status) continue;           // ข้ามแถวว่าง/บรรทัดหมายเหตุ
    out[code] = { left: Number(v[r][iLeft]) || 0, status: status };
  }
  return out;
}

function rowsOfOrder_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const ids = sh.getRange(2, COL.id, last - 1, 1).getValues();
  const out = [];
  ids.forEach((r, i) => { if (String(r[0]) === String(id)) out.push(i + 2); });
  return out;
}

function firstEmptyRow_(sh) {
  const last = Math.max(sh.getLastRow(), 2);
  const vals = sh.getRange(2, 1, last - 1, COL.code).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (!vals[i][COL.id - 1] && !vals[i][COL.code - 1]) return i + 2;
  }
  return last + 1;
}

function ensureSetup_(sh) {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("setup_v4") === "done") return;
  sh.getRange(1, COL.phone, 1, EXTRA_HEADERS.length).setValues([EXTRA_HEADERS])
    .setFontWeight("bold").setBackground("#C8102E").setFontColor("#FFFFFF");
  // เพิ่ม "รอโอน" ใน dropdown สถานะ
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(true).build();
  sh.getRange(2, COL.status, Math.max(sh.getMaxRows() - 1, 1), 1).setDataValidation(rule);
  props.setProperty("setup_v4", "done");
}

function set_(sh, r, c, v) { sh.getRange(r, c).setValue(v); }

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// รันฟังก์ชันนี้ 1 ครั้งจากหน้า Apps Script เพื่อกดอนุญาตสิทธิ์ Google Drive
function authorize() {
  slipFolder_();
  Logger.log("พร้อมแล้ว: โฟลเดอร์ " + SLIP_FOLDER);
}
