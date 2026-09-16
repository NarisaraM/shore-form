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
    "A2": "docs/templates/A2-FORM-A.xlsx",
    "A3C1C2": "docs/templates/A3C1C2-HUTCHISON.xlsx",
}
CHECK_XLSX = "docs/Check.xlsx"
VSLNAME_XLS = "docs/VSLNAME.xls"
LOGO_PNG = "docs/logo.png"


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

    # ---- 2) โลโก้ -> data URI ----
    logo_b64 = b64_of(LOGO_PNG)
    body_inner = body_inner.replace(
        'src="logo.png"', f'src="data:image/png;base64,{logo_b64}"'
    )

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
        '    "มีการส่งข้อมูล SHORE ใหม่ผ่านระบบแบบฟอร์มสำรวจข้อมูล (Self Service Shore)",\n'
        '    "",\n'
        '    "TERMINAL      : "+payload.terminal,\n'
        '    "Vessel / Voy. : "+payload.vessel+" V."+payload.voy,\n'
        '    "Shipper name  : "+payload.shipper,\n'
        '    "POD           : "+payload.pod,\n'
        '    "Booking No.   : "+payload.booking,\n'
        '    "",\n'
        '    "จำนวนตู้ ("+payload.rows.length+" ใบ):"\n'
        "  ];\n"
        "  payload.rows.forEach((r,i)=>{\n"
        "    const extra = [];\n"
        '    if(r.commodity) extra.push("commodity="+r.commodity);\n'
        '    if(r.temp) extra.push("temp="+r.temp);\n'
        '    if(r.vent) extra.push("vent="+r.vent);\n'
        '    if(r.dgUn) extra.push("DG/UN="+r.dgUn);\n'
        '    const extraStr = extra.length ? "  "+extra.join("  ") : "";\n'
        '    lines.push("  "+(i+1)+". "+r.container+"  size="+r.size+"  status="+r.status+extraStr);\n'
        "  });\n"
        '  const contactLine = "Contact: "+(payload.contactName||"")+" "+(payload.contactPhone||"")+\n'
        '    ((payload.contactEmail||"") ? " "+payload.contactEmail : "");\n'
        '  lines.push("", "ผู้ติดต่อ: "+contactLine);\n'
        '  if(payload.remark) lines.push("Special Condition Request: "+payload.remark);\n'
        '  if(warnings && warnings.length) lines.push("", "หมายเหตุ:", ...warnings.map(w=>"  - "+w));\n'
        '  lines.push("", "ไฟล์แนบ: "+outFile);\n'
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
  if(!res.ok) throw new Error("โหลดไฟล์แม่แบบไม่ได้: " + conf.tpl + " (HTTP " + res.status + ")");
  return res.arrayBuffer();
}'''
    new_fill = '''async function loadTemplateBytes(conf){
  const b64 = TEMPLATES_B64[conf.key];
  if(!b64) throw new Error("ไม่พบไฟล์แม่แบบที่ฝังไว้: " + conf.key);
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
  needs_reauth:        "บัญชี Gmail ของคุณต้องเชื่อมต่อใหม่ (ไปที่ Settings → Connectors ใน claude.ai)",
  server_not_connected:"ยังไม่ได้เชื่อมต่อ Gmail กับบัญชี Claude ของคุณ (ไปที่ Settings → Connectors)",
  selection_required:  "มีหลายบัญชี Gmail เชื่อมต่ออยู่ กรุณาเลือกบัญชีในกล่องที่เด้งขึ้น แล้วกดส่งข้อมูลอีกครั้ง",
  not_in_manifest:      "หน้านี้ยังไม่ได้รับอนุญาตให้ใช้ Gmail",
  consent_required:     "คุณยังไม่อนุญาตให้หน้านี้ใช้ Gmail กดส่งข้อมูลอีกครั้งเพื่อขออนุญาต",
  blocked_by_policy:    "องค์กรของคุณปิดการใช้ Gmail จากหน้านี้ไว้",
  approval_required:    "อีเมลนี้ต้องขออนุมัติก่อนส่ง (นโยบายองค์กร)",
  tool_error:           "Gmail รายงานว่าส่งไม่สำเร็จ",
  cancelled:            "การส่งอีเมลถูกยกเลิกกลางคัน (ไม่แน่ใจว่าส่งไปแล้วหรือไม่ กรุณาตรวจ Gmail ก่อนส่งซ้ำ)",
  not_granted:          "หน้านี้ไม่ได้รับสิทธิ์ใช้ตัวเชื่อมต่อ",
  capability_disabled:  "ระบบเชื่อมต่อใช้ไม่ได้ในหน้านี้ขณะนี้",
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
    return {ok:false, reason: GMAIL_ERROR_COPY[code] || ("Gmail ส่งไม่สำเร็จ (" + (code || "unknown") + ")")};
  }
}'''
    assert old_trigger in body_inner, "หา triggerDownload เดิมไม่เจอ"
    body_inner = body_inner.replace(old_trigger, new_trigger)

    # ---- 7) submitForm: ลองส่งผ่าน Gmail ก่อน ถ้าไม่สำเร็จค่อย saveFile สำรอง ----
    old_submit = '''  const btn = $("#btnSubmit");
  btn.disabled = true; btn.textContent = "กำลังสร้างไฟล์…";
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
      msg = "โหลดไฟล์แม่แบบไม่ได้ — ต้องเปิดหน้านี้ผ่าน http/https (GitHub Pages หรือ `python -m http.server`) ไม่ใช่ดับเบิลคลิกไฟล์";
    showResultError(msg);
  }finally{
    btn.disabled = false; btn.textContent = "ส่งข้อมูล";
  }
}'''
    new_submit = '''  const btn = $("#btnSubmit");
  btn.disabled = true; btn.textContent = "กำลังสร้างไฟล์…";
  try{
    const rows = buildRows(payload);
    const { buf, written, warnings } = await fillTemplate(conf, rows);

    const subject = "[SHORE] "+conf.key+" - "+payload.vessel+" V."+payload.voy+" - Booking "+payload.booking;
    const bodyText = buildMailBody(payload, rows, conf.out, warnings);

    btn.textContent = "กำลังส่งอีเมลผ่าน Gmail…";
    const gmail = await tryGmailSend(buf, conf.out, EMAIL_RECIPIENTS, subject, bodyText);
    if(gmail.ok){
      showResultGmailSent(conf.out, written, warnings);
      window.scrollTo({top:document.body.scrollHeight, behavior:"smooth"});
      return;
    }

    // Gmail ส่งไม่สำเร็จ (หรือไม่ได้รันใน Artifact ที่มี capability นี้) -> สำรอง: บันทึกไฟล์
    btn.textContent = "กำลังบันทึกไฟล์…";
    let saveStatus;
    try{
      saveStatus = await saveFile(buf, conf.out);
    }catch(e){
      if(e && e.code === "declined"){
        showResultError("ยกเลิกการบันทึกไฟล์แล้ว — ข้อมูลยังอยู่ในฟอร์ม กดส่งข้อมูลอีกครั้งเมื่อพร้อม");
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
      msg = "โหลดไฟล์แม่แบบไม่ได้ — ต้องเปิดหน้านี้ผ่าน http/https (GitHub Pages หรือ `python -m http.server`) ไม่ใช่ดับเบิลคลิกไฟล์";
    showResultError(msg);
  }finally{
    btn.disabled = false; btn.textContent = "ส่งข้อมูล";
  }
}'''
    assert old_submit in body_inner, "หา submitForm เดิมไม่เจอ"
    body_inner = body_inner.replace(old_submit, new_submit)

    # ---- 8) showResult: เพิ่ม showResultGmailSent + ปรับ showResult ให้ตรงกับ "บันทึกไฟล์" สำรอง ----
    old_result = '''function showResult(outFile, rows, warnings){
  const warn = (warnings && warnings.length)
    ? '<ul>'+warnings.map(w=>'<li>'+escapeHtml(w)+'</li>').join("")+'</ul>' : "";
  $("#resultBox").innerHTML =
    '<div class="result">'+
      '<h2>✓ บันทึกไฟล์สำเร็จ</h2>'+
      '<div>ไฟล์ <span class="file">'+escapeHtml(outFile)+'</span> ถูกดาวน์โหลดไว้ในเครื่องแล้ว</div>'+
      '<div>จำนวนตู้: '+rows+' ใบ</div>'+ warn +
      '<div style="margin-top:10px;color:var(--muted);font-size:13px">'+
      'นำไฟล์นี้ไปส่งต่อ/แนบอีเมลให้ท่าเรือหรือเอเย่นต์ตามขั้นตอนของท่านได้เลย'+
      '</div>'+
    '</div>';
}'''
    new_result = '''function showResultGmailSent(outFile, rows, warnings){
  const warn = (warnings && warnings.length)
    ? '<ul>'+warnings.map(w=>'<li>'+escapeHtml(w)+'</li>').join("")+'</ul>' : "";
  const toList = EMAIL_RECIPIENTS.map(escapeHtml).join(", ");
  $("#resultBox").innerHTML =
    '<div class="result">'+
      '<h2>✓ ส่งอีเมลสำเร็จผ่าน Gmail ของคุณ</h2>'+
      '<div>ไฟล์ <span class="file">'+escapeHtml(outFile)+'</span> ถูกส่งไปที่ '+toList+' เรียบร้อยแล้ว</div>'+
      '<div>จำนวนตู้: '+rows+' ใบ</div>'+ warn +
    '</div>';
}
function showResult(outFile, rows, warnings, saveStatus, gmailReason){
  const warn = (warnings && warnings.length)
    ? '<ul>'+warnings.map(w=>'<li>'+escapeHtml(w)+'</li>').join("")+'</ul>' : "";
  const toList = EMAIL_RECIPIENTS.map(escapeHtml).join(", ");
  const savedLine = saveStatus === "delivered"
    ? 'ไฟล์ <span class="file">'+escapeHtml(outFile)+'</span> ถูกส่งไปยังปลายทางที่คุณเลือกแล้ว'
    : 'ไฟล์ <span class="file">'+escapeHtml(outFile)+'</span> ถูกบันทึกไว้แล้ว';
  const gmailLine = gmailReason
    ? '<li>ส่งผ่าน Gmail อัตโนมัติไม่สำเร็จ: '+escapeHtml(gmailReason)+' — ผู้รับที่ตั้งใจส่ง: '+toList+'</li>'
    : '';
  $("#resultBox").innerHTML =
    '<div class="result">'+
      '<h2>✓ บันทึกไฟล์สำเร็จ</h2>'+
      '<div>'+savedLine+'</div>'+
      '<div>จำนวนตู้: '+rows+' ใบ</div>'+ warn +
      '<ul>'+gmailLine+
      '<li>นำไฟล์นี้ไปส่งต่อ/แนบอีเมลให้ท่าเรือหรือเอเย่นต์ตามขั้นตอนของท่านได้เลย</li>'+
      '</ul>'+
    '</div>';
}'''
    assert old_result in body_inner, "หา showResult เดิมไม่เจอ"
    body_inner = body_inner.replace(old_result, new_result)

    # ---- 9) ประกอบไฟล์สุดท้าย ----
    out = m_title.group(0) + "\n" + m_style.group(0) + "\n" + embedded_block + body_inner
    with open(OUT_HTML, "w", encoding="utf-8", newline="\n") as f:
        f.write(out)

    print(f"เขียนแล้ว: {OUT_HTML} ({os.path.getsize(OUT_HTML)/1024:.0f} KB)")


if __name__ == "__main__":
    main()
