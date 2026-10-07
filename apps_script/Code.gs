/**
 * Self Service Shore - ตัวส่งอีเมลกลาง (Google Apps Script Web App)
 *
 * หน้า GitHub Pages ส่งไฟล์ Excel (base64) มาที่นี่ แล้วสคริปต์
 *   1) ส่งอีเมลแนบไฟล์จากบัญชี Google ที่ Deploy สคริปต์นี้ (shoreheungaline@gmail.com) ไปยัง RECIPIENTS
 *   2) ส่งอีเมลยืนยัน (auto-reply) พร้อมสำเนาไฟล์กลับไปหาผู้กรอก ตามอีเมล Contact ที่ใส่ในฟอร์ม
 * ไม่ต้องใช้ App Password และไม่ต้องมีเซิร์ฟเวอร์
 *
 * วิธีติดตั้งดู README.md หัวข้อ "ส่งอีเมลจากหน้า GitHub Pages ด้วย Google Apps Script"
 */

// ผู้รับอีเมลทุกครั้งที่มีการส่งข้อมูล (แก้ได้ตรงนี้ แล้ว Deploy เวอร์ชันใหม่)
var RECIPIENTS = [
  "narisaram@heungaline.co.th",
  "sirichai@heungaline.co.th",
  "logistics@heungaline.co.th"
];

// อีเมลยืนยันที่ส่งกลับผู้กรอก: ถ้าเขากดตอบกลับ จะไปที่อีเมลนี้
var CONFIRM_REPLY_TO = "logistics@heungaline.co.th";
var SEND_CONFIRMATION = true;                // ตั้งเป็น false เพื่อปิดอีเมลยืนยัน

var SENDER_NAME = "SHORE Self Service";
var MAX_PER_HOUR = 40;                       // จำกัดจำนวนการส่งต่อชั่วโมงทั้งระบบ (กันสแปม)
var MAX_CONFIRM_PER_ADDRESS_PER_DAY = 3;     // อีเมลยืนยันต่อที่อยู่ต่อวัน (กันใช้ระบบยิงอีเมลใส่คนอื่น)
var MAX_FILE_B64 = 3 * 1024 * 1024;          // ขนาดไฟล์ base64 สูงสุด ~2.2MB
var XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
var EMAIL_RE = /^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/;

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    var d = JSON.parse((e && e.postData && e.postData.contents) || "{}");

    // ช่องล่อบอท (คนจริงไม่เห็น/ไม่กรอก) : ถ้ามีค่า ตอบสำเร็จหลอกแล้วไม่ส่งจริง
    if (d.hp) return json_({ ok: true });

    var filename = String(d.filename || "");
    var fileB64 = String(d.fileB64 || "");
    if (!/^[A-Za-z0-9._-]{1,80}\.xlsx$/.test(filename)) return json_({ ok: false, error: "invalid filename" });
    if (!fileB64 || fileB64.length > MAX_FILE_B64) return json_({ ok: false, error: "invalid file size" });

    var bytes = Utilities.base64Decode(fileB64);
    // ไฟล์ .xlsx จริงต้องเป็น zip (ขึ้นต้นด้วย PK)
    if (bytes.length < 4 || bytes[0] !== 80 || bytes[1] !== 75) return json_({ ok: false, error: "not an xlsx file" });

    if (!underHourlyLimit_()) return json_({ ok: false, error: "too many submissions, please try again later" });

    var blob = Utilities.newBlob(bytes, XLSX_MIME, filename);
    var subject = oneLine_(d.subject || "[SHORE] New submission", 200);
    var body = String(d.body || "").substring(0, 8000) +
      "\n\n-- Sent automatically by Self Service Shore --";
    var options = {
      to: RECIPIENTS.join(","),
      subject: subject,
      body: body,
      name: SENDER_NAME,
      attachments: [blob]
    };
    var contactEmail = String(d.replyTo || "").trim();
    if (EMAIL_RE.test(contactEmail)) options.replyTo = contactEmail;

    MailApp.sendEmail(options);

    // อีเมลยืนยันกลับไปหาผู้กรอก (ล้มเหลวก็ไม่กระทบการส่งหลัก)
    var confirmedTo = "";
    if (SEND_CONFIRMATION && EMAIL_RE.test(contactEmail)) {
      try {
        confirmedTo = sendConfirmation_(contactEmail, d.info || {}, blob);
      } catch (cerr) {
        confirmedTo = "";
      }
    }
    return json_({ ok: true, confirmedTo: confirmedTo });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err).substring(0, 200) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// เปิดลิงก์สคริปต์ในเบราว์เซอร์ (GET) เพื่อเช็กว่า Deploy แล้ว
function doGet() {
  return json_({ ok: true, service: "Self Service Shore mailer" });
}

/** ส่งอีเมลยืนยันถึงผู้กรอก คืนที่อยู่อีเมลถ้าส่งแล้ว หรือ "" ถ้าไม่ได้ส่ง */
function sendConfirmation_(to, info, blob) {
  if (MailApp.getRemainingDailyQuota() < 1) return "";
  if (!underAddressLimit_(to)) return "";

  var name = oneLine_(info.contactName || "", 80);
  var vessel = oneLine_(info.vessel, 80), voy = oneLine_(info.voy, 40);
  var booking = oneLine_(info.booking, 60);
  var containers = (Array.isArray(info.containers) ? info.containers : []).slice(0, 50)
    .map(function (c, i) { return "  " + (i + 1) + ". " + oneLine_(c, 80); });

  var lines = [
    "Dear " + (name || "Sir/Madam") + ",",
    "",
    "Thank you. Heung-A Line has received your booking amendment request submitted via Self Service Shore.",
    "A copy of the submitted file is attached for your records.",
    "",
    "TERMINAL      : " + oneLine_(info.terminal, 120),
    "Vessel / Voy. : " + vessel + " V." + voy,
    "Shipper name  : " + oneLine_(info.shipper, 120),
    "Booking No.   : " + booking,
    "",
    "Containers (" + containers.length + "):"
  ].concat(containers, [
    "",
    "Our team will review your request and contact you if anything else is needed.",
    "If any information above is incorrect, please reply to this email.",
    "",
    "Best regards,",
    "Heung-A Line",
    "",
    "-- This is an automated message from Self Service Shore --"
  ]);

  MailApp.sendEmail({
    to: to,
    subject: oneLine_("[SHORE] We received your submission - " + vessel + " V." + voy + " / Booking " + booking, 200),
    body: lines.join("\n"),
    name: SENDER_NAME,
    replyTo: CONFIRM_REPLY_TO,
    attachments: [blob]
  });
  return to;
}

function oneLine_(v, max) {
  return String(v == null ? "" : v).replace(/[\r\n]+/g, " ").substring(0, max);
}

function underHourlyLimit_() {
  var props = PropertiesService.getScriptProperties();
  var key = "h_" + Utilities.formatDate(new Date(), "UTC", "yyyyMMddHH");
  var n = Number(props.getProperty(key) || 0);
  if (n >= MAX_PER_HOUR) return false;
  props.setProperty(key, String(n + 1));
  return true;
}

/** จำกัดจำนวนอีเมลยืนยันต่อที่อยู่ต่อวัน และลบตัวนับของวันก่อน ๆ ทิ้ง */
function underAddressLimit_(address) {
  var props = PropertiesService.getScriptProperties();
  var day = Utilities.formatDate(new Date(), "UTC", "yyyyMMdd");
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(address).toLowerCase());
  var hex = digest.map(function (b) { return ((b + 256) % 256).toString(16).padStart(2, "0"); }).join("");
  var key = "c_" + day + "_" + hex;

  props.getKeys().forEach(function (k) {
    if (k.indexOf("c_") === 0 && k.indexOf("c_" + day + "_") !== 0) props.deleteProperty(k);
  });

  var n = Number(props.getProperty(key) || 0);
  if (n >= MAX_CONFIRM_PER_ADDRESS_PER_DAY) return false;
  props.setProperty(key, String(n + 1));
  return true;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** กด Run ฟังก์ชันนี้ 1 ครั้งตอนติดตั้ง เพื่อให้ Google ขออนุญาตส่งอีเมล (ส่งอีเมลทดสอบจริง 1 ฉบับ) */
function authorizeAndTest() {
  MailApp.sendEmail({
    to: RECIPIENTS.join(","),
    subject: "[SHORE] Test - Apps Script mailer is working",
    body: "This is a test email from the Self Service Shore Apps Script mailer.",
    name: SENDER_NAME
  });
}
