/**
 * Self Service Shore - ตัวส่งอีเมลกลาง (Google Apps Script Web App)
 *
 * หน้า GitHub Pages ส่งไฟล์ Excel (base64) มาที่นี่ แล้วสคริปต์ส่งอีเมลแนบไฟล์
 * จากบัญชี Google ที่ Deploy สคริปต์นี้ (shoreheungaline@gmail.com) ไปยัง RECIPIENTS
 * ไม่ต้องใช้ App Password และไม่ต้องมีเซิร์ฟเวอร์
 *
 * วิธีติดตั้งดู README.md หัวข้อ "ส่งอีเมลจากหน้า GitHub Pages ด้วย Google Apps Script"
 */

// ผู้รับอีเมลทุกครั้งที่มีการส่งข้อมูล (แก้ได้ตรงนี้ แล้ว Deploy เวอร์ชันใหม่)
var RECIPIENTS = [
  "narisaram@heungaline.co.th",
  "sirichai@heungaline.co.th"
];

var SENDER_NAME = "SHORE Self Service";
var MAX_PER_HOUR = 40;                       // จำกัดจำนวนอีเมลต่อชั่วโมงทั้งระบบ (กันสแปม)
var MAX_FILE_B64 = 3 * 1024 * 1024;          // ขนาดไฟล์ base64 สูงสุด ~2.2MB
var XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

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

    var subject = String(d.subject || "[SHORE] New submission").replace(/[\r\n]+/g, " ").substring(0, 200);
    var body = String(d.body || "").substring(0, 8000) +
      "\n\n-- Sent automatically by Self Service Shore --";
    var options = {
      to: RECIPIENTS.join(","),
      subject: subject,
      body: body,
      name: SENDER_NAME,
      attachments: [Utilities.newBlob(bytes, XLSX_MIME, filename)]
    };
    var replyTo = String(d.replyTo || "").trim();
    if (/^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/.test(replyTo)) options.replyTo = replyTo;

    MailApp.sendEmail(options);
    return json_({ ok: true });
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

function underHourlyLimit_() {
  var props = PropertiesService.getScriptProperties();
  var key = "h_" + Utilities.formatDate(new Date(), "UTC", "yyyyMMddHH");
  var n = Number(props.getProperty(key) || 0);
  if (n >= MAX_PER_HOUR) return false;
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
