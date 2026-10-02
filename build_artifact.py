# -*- coding: utf-8 -*-
"""
build_artifact.py
สร้างเวอร์ชัน Claude Artifact ของฟอร์ม (ไฟล์ HTML ไฟล์เดียว จบในตัว)

ต่างจากเวอร์ชัน docs/ (GitHub Pages) ตรงที่:
  - ฝังไฟล์แม่แบบ Excel + Check.xlsx + โลโก้ เป็น base64 ไว้ในไฟล์เดียว
    (Artifact ไม่มีไฟล์พี่น้องให้ fetch())
  - บันทึกไฟล์ผลลัพธ์ผ่าน capability "downloads" ของ Artifact แทนการคลิกลิงก์ดาวน์โหลด
    (ลิงก์ดาวน์โหลดแบบเดิมถูกบล็อกใน sandbox ของ Artifact)
  - ตัด <!DOCTYPE>/<html>/<head>/<body> ออก เหลือแค่ <title> + <style> + เนื้อหา
    ตามข้อกำหนดของ Artifact tool

รันเพื่อสร้าง/อัปเดต artifact_src/self-service-shore.html:
    python build_artifact.py
"""

from __future__ import annotations

import base64
import os
import re

BASE = os.path.dirname(os.path.abspath(__file__))
SRC_HTML = os.path.join(BASE, "docs", "index.html")
OUT_DIR = os.path.join(BASE, "artifact_src")
OUT_HTML = os.path.join(OUT_DIR, "self-service-shore.html")

TEMPLATES = {
    "A0": "docs/templates/A0-SHORE.xlsx",
    "B3": "docs/templates/B3-SHORE.xlsx",
    "B5C3": "docs/templates/B5C3-SHORE.xlsx",
    "A3C1C2": "docs/templates/A3C1C2-HUTCHISON.xlsx",
}
CHECK_XLSX = "docs/Check.xlsx"
VSLNAME_XLS = "docs/VSLNAME.xls"
LOGO_PNG = "docs/logo.png"
BG_JPG = "docs/bg-containers.jpg"


def b64_of(rel_path: str) -> str:
    with open(os.path.join(BASE, rel_path), "rb") as f:
        return base64.b64encode(f.read()).decode("ascii")


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    with open(SRC_HTML, encoding="utf-8") as f:
        html = f.read()

    # ---- 1) ตัด wrapper <!DOCTYPE>/<html>/<head>/<body> ----
    m_title = re.search(r"<title>.*?</title>", html, re.S)
    m_style = re.search(r"<style>.*?</style>", html, re.S)
    m_body = re.search(r"<body>(.*)</body>\s*</html>\s*$", html, re.S)
    if not (m_title and m_style and m_body):
        raise SystemExit("โครงสร้าง docs/index.html เปลี่ยนไป - แก้สคริปต์นี้ตามด้วย")
    body_inner = m_body.group(1)

    # ---- 2) โลโก้ + รูปพื้นหลัง (watermark) -> data URI ----
    logo_b64 = b64_of(LOGO_PNG)
    old_logo_src = 'src="logo.png?v=2"'
    assert old_logo_src in body_inner, "หา src=\"logo.png?v=2\" ใน docs/index.html ไม่เจอ (เลขเวอร์ชัน cache-bust เปลี่ยนไปหรือเปล่า?)"
    body_inner = body_inner.replace(
        old_logo_src, f'src="data:image/png;base64,{logo_b64}"'
    )
    bg_b64 = b64_of(BG_JPG)
    style_text = m_style.group(0).replace(
        'url("bg-containers.jpg")', f'url("data:image/jpeg;base64,{bg_b64}")'
    )
    assert 'data:image/jpeg;base64,' in style_text, "หา url(\"bg-containers.jpg\") ใน <style> ไม่เจอ"

    # ---- 3) ฝังไฟล์แม่แบบ + Check.xlsx เป็น base64 (แทรกก่อนสคริปต์หลัก) ----
    templates_js = ",\n  ".join(f'"{k}": "{b64_of(p)}"' for k, p in TEMPLATES.items())
    check_b64 = b64_of(CHECK_XLSX)
    vslname_b64 = b64_of(VSLNAME_XLS)
    embedded_block = (
        "<script>\n"
        "/* ไฟล์แม่แบบ Excel + Check.xlsx + VSLNAME.xls ฝังไว้เป็น base64 (Artifact ไม่มีไฟล์พี่น้องให้ fetch) */\n"
        "const TEMPLATES_B64 = {\n  " + templates_js + "\n};\n"
        'const CHECK_B64 = "' + check_b64 + '";\n'
        'const VSLNAME_B64 = "' + vslname_b64 + '";\n'
        "function b64ToArrayBuffer(b64){\n"
        "  const bin = atob(b64);\n"
        "  const bytes = new Uint8Array(bin.length);\n"
        "  for(let i=0;i<bin.length;i++) bytes[i] = bin.charCodeAt(i);\n"
        "  return bytes.buffer;\n"
        "}\n"
        "\n"
        "/* ผู้รับอีเมล + เนื้อหาอีเมล สำหรับการส่งผ่าน Gmail (เวอร์ชัน docs/ แบบ static ตัดออกแล้ว\n"
        "   เพราะ GitHub Pages ส่งอีเมลตรงไม่ได้ แต่ Artifact เชื่อม Gmail ของผู้ใช้ได้จริง จึงคงไว้ที่นี่) */\n"
        'const EMAIL_RECIPIENTS = ["dongykong.naris@gmail.com", "sirichai@heungaline.co.th"];\n'
        "function buildMailBody(payload, rows, outFile, warnings){\n"
        "  const lines = [\n"
        '    "A new SHORE submission has been sent via the Self Service Shore form",\n'
        '    "",\n'
        '    "TERMINAL      : "+payload.terminal,\n'
        '    "Vessel / Voy. : "+payload.vessel+" V."+payload.voy,\n'
        '    "Shipper name  : "+payload.shipper,\n'
        '    "POD           : "+payload.pod,\n'
        '    "Booking No.   : "+payload.booking,\n'
        '    "",\n'
        '    "Containers ("+payload.rows.length+"):"\n'
        "  ];\n"
        "  payload.rows.forEach((r,i)=>{\n"
        "    const extra = [];\n"
        '    if(r.commodity) extra.push("commodity="+r.commodity);\n'
        '    if(r.temp) extra.push("temp="+r.temp);\n'
        '    if(r.humidity) extra.push("humidity="+r.humidity);\n'
        '    if(r.vent) extra.push("vent="+r.vent);\n'
        '    if(r.dgUn) extra.push("DG/UN="+r.dgUn);\n'
        '    if(r.overHeight) extra.push("overHeight="+r.overHeight);\n'
        '    if(r.overWidth) extra.push("overWidth="+r.overWidth);\n'
        '    const extraStr = extra.length ? "  "+extra.join("  ") : "";\n'
        '    lines.push("  "+(i+1)+". "+r.container+"  size="+r.size+"  status="+r.status+extraStr);\n'
        "  });\n"
        '  const contactLine = "Contact: "+(payload.contactName||"")+" "+(payload.contactPhone||"")+\n'
        '    ((payload.contactEmail||"") ? " "+payload.contactEmail : "");\n'
        '  lines.push("", contactLine);\n'
        '  if(payload.remark) lines.push("Special Condition Request: "+payload.remark);\n'
        '  if(warnings && warnings.length) lines.push("", "Notes:", ...warnings.map(w=>"  - "+w));\n'
        '  lines.push("", "Attachment: "+outFile);\n'
        "  return lines.join(\"\\n\");\n"
        "}\n"
        "</script>\n"
    )

    # ---- 4) แก้ loadConfig ให้อ่านจาก CHECK_B64/VSLNAME_B64 แทน fetch ----
    old_load = '''  try{
    const res = await fetch("Check.xlsx", {cache:"no-store"});
    if(res.ok){
      const wb = XLSX.read(await res.arrayBuffer(), {type:"array"});
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, {header:1, blankrows:false, defval:""});
      const T=[], S=[], ST=[]; let cur=null;
      for(const row of rows){
        const a = String(row[0] ?? "").trim().toLowerCase();
        const b = String(row[1] ?? "").trim();
        if(a){
          cur = a.includes("terminal") ? T : a.includes("size") ? S : a.includes("status") ? ST : null;
        }
        if(b && cur) cur.push(b);
      }
      cfg = {
        terminals: T.length ? T : CFG_FALLBACK.terminals,
        sizes:     S.length ? S : CFG_FALLBACK.sizes,
        statuses:  ST.length ? ST : CFG_FALLBACK.statuses,
        vessels:   CFG_FALLBACK.vessels
      };
    }
  }catch(e){ /* ใช้ fallback */ }

  try{
    const res = await fetch("VSLNAME.xls", {cache:"no-store"});
    if(res.ok){
      const wb = XLSX.read(await res.arrayBuffer(), {type:"array"});
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, {header:1, blankrows:false, defval:""});
      const V = rows.slice(1).map(r=>String(r[0] ?? "").trim()).filter(Boolean);
      if(V.length) cfg.vessels = V;
    }
  }catch(e){ /* ใช้ fallback */ }'''
    new_load = '''  try{
    const wb = XLSX.read(b64ToArrayBuffer(CHECK_B64), {type:"array"});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, blankrows:false, defval:""});
    const T=[], S=[], ST=[]; let cur=null;
    for(const row of rows){
      const a = String(row[0] ?? "").trim().toLowerCase();
      const b = String(row[1] ?? "").trim();
      if(a){
        cur = a.includes("terminal") ? T : a.includes("size") ? S : a.includes("status") ? ST : null;
      }
      if(b && cur) cur.push(b);
    }
    cfg = {
      terminals: T.length ? T : CFG_FALLBACK.terminals,
      sizes:     S.length ? S : CFG_FALLBACK.sizes,
      statuses:  ST.length ? ST : CFG_FALLBACK.statuses,
      vessels:   CFG_FALLBACK.vessels
    };
  }catch(e){ /* ใช้ fallback */ }

  try{
    const wb = XLSX.read(b64ToArrayBuffer(VSLNAME_B64), {type:"array"});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, blankrows:false, defval:""});
    const V = rows.slice(1).map(r=>String(r[0] ?? "").trim()).filter(Boolean);
    if(V.length) cfg.vessels = V;
  }catch(e){ /* ใช้ fallback */ }'''
    assert old_load in body_inner, "หา loadConfig เดิมไม่เจอ"
    body_inner = body_inner.replace(old_load, new_load)

    # ---- 5) แก้ loadTemplateBytes ให้อ่านจาก TEMPLATES_B64 (ฝังในไฟล์) แทน fetch ----
    old_fill = '''async function loadTemplateBytes(conf){
  const res = await fetch(conf.tpl, {cache:"no-store"});
  if(!res.ok) throw new Error("Could not load template file: " + conf.tpl + " (HTTP " + res.status + ")");
  return res.arrayBuffer();
}'''
    new_fill = '''async function loadTemplateBytes(conf){
  const b64 = TEMPLATES_B64[conf.key];
  if(!b64) throw new Error("Embedded template file not found: " + conf.key);
  return b64ToArrayBuffer(b64);
}'''
    assert old_fill in body_inner, "หา loadTemplateBytes เดิมไม่เจอ"
    body_inner = body_inner.replace(old_fill, new_fill)

    # ---- 6) เปลี่ยนการ "ดาวน์โหลด" ให้ใช้ capability "downloads" ของ Artifact ----
    old_trigger = '''function triggerDownload(buf, name){
  const blob = new Blob([buf], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url), 4000);
}'''
    new_trigger = '''async function saveFile(buf, name){
  const blob = new Blob([buf], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  if(window.claude && window.claude.use){
    const downloads = await window.claude.use("downloads");
    if(downloads){
      const res = await downloads.save({filename: name, data: blob}); // อาจ throw {code:"declined",...}
      return res.status; // "saved" | "delivered"
    }
  }
  // นอก Artifact (เช่น เปิดไฟล์นี้ตรง ๆ ตอนพัฒนา) - ใช้ลิงก์ดาวน์โหลดแบบเดิม
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url), 4000);
  return "saved";
}

/* ================= ส่งอีเมลอัตโนมัติผ่าน Gmail ที่เชื่อมกับบัญชี Claude ของผู้ใช้ =================
   ใช้ capability "mcp" ของ Artifact เรียก connector "Gmail" โดยตรง (ไม่ผ่านเซิร์ฟเวอร์ใด ๆ)
   ทดสอบรูปแบบคำขอ/ไฟล์แนบจริงแล้วก่อนขึ้นระบบ (ผ่าน create_draft แบบไม่ส่งจริง) */
const GMAIL_TOOL = "send_message";
const GMAIL_ERROR_COPY = {
  needs_reauth:        "Your Gmail account needs to be reconnected (go to Settings → Connectors on claude.ai)",
  server_not_connected:"Gmail is not connected to your Claude account (go to Settings → Connectors)",
  selection_required:  "Multiple Gmail accounts are connected — please choose one in the popup and submit again",
  not_in_manifest:      "This page has not been authorized to use Gmail",
  consent_required:     "You have not authorized this page to use Gmail — submit again to request permission",
  blocked_by_policy:    "Your organization has disabled Gmail access from this page",
  approval_required:    "This email requires approval before sending (organization policy)",
  tool_error:           "Gmail reported that the send failed",
  cancelled:            "Email sending was cancelled midway (it's unclear whether it sent — please check Gmail before retrying)",
  not_granted:          "This page does not have permission to use the connector",
  capability_disabled:  "The connector is not available on this page right now",
};
function arrayBufferToB64(buf){
  const bytes = new Uint8Array(buf);
  let bin = "";
  const chunk = 0x8000;
  for(let i=0;i<bytes.length;i+=chunk) bin += String.fromCharCode.apply(null, bytes.subarray(i, i+chunk));
  return btoa(bin);
}
async function tryGmailSend(buf, filename, toList, subject, bodyText){
  if(!(window.claude && window.claude.use)) return {ok:false, reason:null}; // ไม่ได้รันใน Artifact
  let mcp;
  try{ mcp = await window.claude.use("mcp"); }catch(e){ mcp = null; }
  if(!mcp) return {ok:false, reason:null};
  try{
    const res = await mcp.callTool("Gmail", GMAIL_TOOL, {
      to: toList,
      subject: subject,
      body: bodyText,
      attachments: [{
        content: arrayBufferToB64(buf),
        filename: filename,
        mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      }]
    });
    return {ok:true, payload: res.payload};
  }catch(e){
    const code = e && e.code;
    return {ok:false, reason: GMAIL_ERROR_COPY[code] || ("Gmail send failed (" + (code || "unknown") + ")")};
  }
}'''
    assert old_trigger in body_inner, "หา triggerDownload เดิมไม่เจอ"
    body_inner = body_inner.replace(old_trigger, new_trigger)

    # ---- 7) submitForm: ลองส่งผ่าน Gmail ก่อน ถ้าไม่สำเร็จค่อย saveFile สำรอง ----
    old_submit = '''  const btn = $("#btnSubmit");
  btn.disabled = true; btn.textContent = "Generating file…";
  try{
    const rows = buildRows(payload);
    const { buf, written, warnings } = await fillTemplate(conf, rows);
    triggerDownload(buf, conf.out);
    showResult(conf.out, written, warnings);
    window.scrollTo({top:document.body.scrollHeight, behavior:"smooth"});
  }catch(e){
    console.error(e);
    let msg = e.message || String(e);
    if(/Failed to fetch|NetworkError|HTTP 0/i.test(msg))
      msg = "Could not load the template file — this page must be opened via http/https (GitHub Pages or `python -m http.server`), not by double-clicking the file";
    showResultError(msg);
  }finally{
    btn.disabled = false; btn.textContent = "Submit";
  }
}'''
    new_submit = '''  const btn = $("#btnSubmit");
  btn.disabled = true; btn.textContent = "Generating file…";
  try{
    const rows = buildRows(payload);
    const { buf, written, warnings } = await fillTemplate(conf, rows);

    const subject = "[SHORE] "+conf.key+" - "+payload.vessel+" V."+payload.voy+" - Booking "+payload.booking;
    const bodyText = buildMailBody(payload, rows, conf.out, warnings);

    btn.textContent = "Sending email via Gmail…";
    const gmail = await tryGmailSend(buf, conf.out, EMAIL_RECIPIENTS, subject, bodyText);
    if(gmail.ok){
      showResultGmailSent(conf.out, written, warnings);
      window.scrollTo({top:document.body.scrollHeight, behavior:"smooth"});
      return;
    }

    // Gmail send failed (or not running in an Artifact with this capability) -> fallback: save file
    btn.textContent = "Saving file…";
    let saveStatus;
    try{
      saveStatus = await saveFile(buf, conf.out);
    }catch(e){
      if(e && e.code === "declined"){
        showResultError("File save cancelled — your data is still in the form; submit again when ready");
        return;
      }
      throw e;
    }
    showResult(conf.out, written, warnings, saveStatus, gmail.reason);
    window.scrollTo({top:document.body.scrollHeight, behavior:"smooth"});
  }catch(e){
    console.error(e);
    let msg = e.message || String(e);
    if(/Failed to fetch|NetworkError|HTTP 0/i.test(msg))
      msg = "Could not load the template file — this page must be opened via http/https (GitHub Pages or `python -m http.server`), not by double-clicking the file";
    showResultError(msg);
  }finally{
    btn.disabled = false; btn.textContent = "Submit";
  }
}'''
    assert old_submit in body_inner, "old submitForm not found"
    body_inner = body_inner.replace(old_submit, new_submit)

    # ---- 8) showResult: เพิ่ม showResultGmailSent + ปรับ showResult ให้ตรงกับ "บันทึกไฟล์" สำรอง ----
    old_result = '''function showResult(outFile, rows, warnings){
  const warn = (warnings && warnings.length)
    ? '<ul>'+warnings.map(w=>'<li>'+escapeHtml(w)+'</li>').join("")+'</ul>' : "";
  $("#resultBox").innerHTML =
    '<div class="result">'+
      '<h2>✓ File saved successfully</h2>'+
      '<div>File <span class="file">'+escapeHtml(outFile)+'</span> has been downloaded to your device</div>'+
      '<div>Containers: '+rows+'</div>'+ warn +
      '<div style="margin-top:10px;color:var(--muted);font-size:13px">'+
      'You can now forward or email this file to the terminal or agent as needed'+
      '</div>'+
    '</div>';
}'''
    new_result = '''function showResultGmailSent(outFile, rows, warnings){
  const warn = (warnings && warnings.length)
    ? '<ul>'+warnings.map(w=>'<li>'+escapeHtml(w)+'</li>').join("")+'</ul>' : "";
  const toList = EMAIL_RECIPIENTS.map(escapeHtml).join(", ");
  $("#resultBox").innerHTML =
    '<div class="result">'+
      '<h2>✓ Email sent successfully via your Gmail</h2>'+
      '<div>File <span class="file">'+escapeHtml(outFile)+'</span> has been sent to '+toList+'</div>'+
      '<div>Containers: '+rows+'</div>'+ warn +
    '</div>';
}
function showResult(outFile, rows, warnings, saveStatus, gmailReason){
  const warn = (warnings && warnings.length)
    ? '<ul>'+warnings.map(w=>'<li>'+escapeHtml(w)+'</li>').join("")+'</ul>' : "";
  const toList = EMAIL_RECIPIENTS.map(escapeHtml).join(", ");
  const savedLine = saveStatus === "delivered"
    ? 'File <span class="file">'+escapeHtml(outFile)+'</span> has been sent to your selected destination'
    : 'File <span class="file">'+escapeHtml(outFile)+'</span> has been saved';
  const gmailLine = gmailReason
    ? '<li>Automatic Gmail sending failed: '+escapeHtml(gmailReason)+' — intended recipients: '+toList+'</li>'
    : '';
  $("#resultBox").innerHTML =
    '<div class="result">'+
      '<h2>✓ File saved successfully</h2>'+
      '<div>'+savedLine+'</div>'+
      '<div>Containers: '+rows+'</div>'+ warn +
      '<ul>'+gmailLine+
      '<li>You can now forward or email this file to the terminal or agent as needed</li>'+
      '</ul>'+
    '</div>';
}'''
    assert old_result in body_inner, "old showResult not found"
    body_inner = body_inner.replace(old_result, new_result)

    # ---- 9) ประกอบไฟล์สุดท้าย ----
    out = m_title.group(0) + "\n" + style_text + "\n" + embedded_block + body_inner
    with open(OUT_HTML, "w", encoding="utf-8", newline="\n") as f:
        f.write(out)

    print(f"เขียนแล้ว: {OUT_HTML} ({os.path.getsize(OUT_HTML)/1024:.0f} KB)")


if __name__ == "__main__":
    main()
