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
    embedded_block = (
        "<script>\n"
        "/* ไฟล์แม่แบบ Excel + Check.xlsx ฝังไว้เป็น base64 (Artifact ไม่มีไฟล์พี่น้องให้ fetch) */\n"
        "const TEMPLATES_B64 = {\n  " + templates_js + "\n};\n"
        'const CHECK_B64 = "' + check_b64 + '";\n'
        "function b64ToArrayBuffer(b64){\n"
        "  const bin = atob(b64);\n"
        "  const bytes = new Uint8Array(bin.length);\n"
        "  for(let i=0;i<bin.length;i++) bytes[i] = bin.charCodeAt(i);\n"
        "  return bytes.buffer;\n"
        "}\n"
        "</script>\n"
    )

    # ---- 4) แก้ loadConfig ให้อ่านจาก CHECK_B64 แทน fetch ----
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
        statuses:  ST.length ? ST : CFG_FALLBACK.statuses
      };
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
      statuses:  ST.length ? ST : CFG_FALLBACK.statuses
    };
  }catch(e){ /* ใช้ fallback */ }'''
    assert old_load in body_inner, "หา loadConfig เดิมไม่เจอ"
    body_inner = body_inner.replace(old_load, new_load)

    # ---- 5) แก้ fillTemplate ให้อ่านจาก TEMPLATES_B64 แทน fetch ----
    old_fill = '''  const res = await fetch(conf.tpl, {cache:"no-store"});
  if(!res.ok) throw new Error("โหลดไฟล์แม่แบบไม่ได้: " + conf.tpl + " (HTTP " + res.status + ")");
  const wb = XLSX.read(await res.arrayBuffer(), {type:"array", cellStyles:true});'''
    new_fill = '''  const b64 = TEMPLATES_B64[conf.key];
  if(!b64) throw new Error("ไม่พบไฟล์แม่แบบที่ฝังไว้: " + conf.key);
  const wb = XLSX.read(b64ToArrayBuffer(b64), {type:"array", cellStyles:true});'''
    assert old_fill in body_inner, "หา fillTemplate เดิมไม่เจอ"
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
}'''
    assert old_trigger in body_inner, "หา triggerDownload เดิมไม่เจอ"
    body_inner = body_inner.replace(old_trigger, new_trigger)

    # ---- 7) submitForm: เรียก saveFile แทน triggerDownload, จัดการ "declined" ----
    old_submit = '''  const btn = $("#btnSubmit");
  btn.disabled = true; btn.textContent = "กำลังสร้างไฟล์…";
  try{
    const rows = buildRows(payload);
    const { buf, written, warnings } = await fillTemplate(conf, rows);
    triggerDownload(buf, conf.out);
    showResult(conf.out, written, warnings);
    window.scrollTo({top:document.body.scrollHeight, behavior:"smooth"});
    setTimeout(()=>openMailDraft(payload, rows, conf.out, warnings), 600);
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
    showResult(conf.out, written, warnings, saveStatus);
    window.scrollTo({top:document.body.scrollHeight, behavior:"smooth"});
    setTimeout(()=>openMailDraft(payload, rows, conf.out, warnings), 600);
  }catch(e){
    console.error(e);
    showResultError(e.message || String(e));
  }finally{
    btn.disabled = false; btn.textContent = "ส่งข้อมูล";
  }
}'''
    assert old_submit in body_inner, "หา submitForm เดิมไม่เจอ"
    body_inner = body_inner.replace(old_submit, new_submit)

    # ---- 8) showResult: ข้อความให้ตรงกับ "บันทึกไฟล์" แทน "ดาวน์โหลดอัตโนมัติ" ----
    old_result = '''function showResult(outFile, rows, warnings){
  const warn = (warnings && warnings.length)
    ? '<ul>'+warnings.map(w=>'<li>'+escapeHtml(w)+'</li>').join("")+'</ul>' : "";
  const toList = EMAIL_RECIPIENTS.map(escapeHtml).join(", ");
  $("#resultBox").innerHTML =
    '<div class="result">'+
      '<h2>✓ สร้างไฟล์สำเร็จ — กำลังเปิดโปรแกรมอีเมลให้</h2>'+
      '<div>ไฟล์ <span class="file">'+escapeHtml(outFile)+'</span> ถูกดาวน์โหลดไว้แล้ว</div>'+
      '<div>จำนวนตู้: '+rows+' ใบ</div>'+ warn +
      '<ul><li>โปรแกรมอีเมลของคุณจะเปิดขึ้น พร้อมผู้รับ/หัวเรื่อง/เนื้อหาให้แล้ว (ถึง '+toList+')</li>'+
      '<li><b>กรุณาแนบไฟล์ '+escapeHtml(outFile)+' ที่เพิ่งดาวน์โหลด</b> แล้วกด Send ในโปรแกรมอีเมลของคุณ</li>'+
      '<li>เว็บเวอร์ชันนี้ (GitHub Pages) ไม่มีเซิร์ฟเวอร์ จึงแนบไฟล์และส่งให้อัตโนมัติ 100% ไม่ได้ — '+
      'ถ้าต้องการให้ส่งอัตโนมัติจริงไม่ต้องแนบเอง ใช้เวอร์ชัน Python (<code>python run.py</code>)</li></ul>'+
    '</div>';
}'''
    new_result = '''function showResult(outFile, rows, warnings, saveStatus){
  const warn = (warnings && warnings.length)
    ? '<ul>'+warnings.map(w=>'<li>'+escapeHtml(w)+'</li>').join("")+'</ul>' : "";
  const toList = EMAIL_RECIPIENTS.map(escapeHtml).join(", ");
  const savedLine = saveStatus === "delivered"
    ? 'ไฟล์ <span class="file">'+escapeHtml(outFile)+'</span> ถูกส่งไปยังปลายทางที่คุณเลือกแล้ว'
    : 'ไฟล์ <span class="file">'+escapeHtml(outFile)+'</span> ถูกบันทึกไว้แล้ว';
  $("#resultBox").innerHTML =
    '<div class="result">'+
      '<h2>✓ สร้างไฟล์สำเร็จ — กำลังเปิดโปรแกรมอีเมลให้</h2>'+
      '<div>'+savedLine+'</div>'+
      '<div>จำนวนตู้: '+rows+' ใบ</div>'+ warn +
      '<ul><li>โปรแกรมอีเมลของคุณจะเปิดขึ้น พร้อมผู้รับ/หัวเรื่อง/เนื้อหาให้แล้ว (ถึง '+toList+')</li>'+
      '<li><b>กรุณาแนบไฟล์ '+escapeHtml(outFile)+' ที่เพิ่งบันทึกไว้</b> แล้วกด Send ในโปรแกรมอีเมลของคุณ</li>'+
      '<li>หน้านี้ (Artifact) ไม่มีเซิร์ฟเวอร์ จึงแนบไฟล์และส่งให้อัตโนมัติ 100% ไม่ได้ — '+
      'ถ้าต้องการให้ส่งอัตโนมัติจริงไม่ต้องแนบเอง ใช้เวอร์ชัน Python (<code>python run.py</code>)</li></ul>'+
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
