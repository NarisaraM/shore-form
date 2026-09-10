# -*- coding: utf-8 -*-
"""
pdf_extract.py
อ่านไฟล์ PDF ทั้งหมดในโฟลเดอร์ pdf_in/ ด้วย pdfplumber (ดึง "ข้อความ" ออกมาล้วน ๆ)
แล้ววิเคราะห์ด้วยกฎ (regex/keyword) *โดยไม่ใช้ AI* เพื่อเดาค่าเบื้องต้นสำหรับกรอกฟอร์ม

ขั้นตอน:
  1) pdfplumber เปิดไฟล์ -> page.extract_text() ต่อกันทุกหน้า  (ส่วนนี้ไม่มี AI)
  2) ฟังก์ชัน parse_text() ใช้กฎล้วน ๆ หาค่า:
       - เลขตู้ (Container No.)  รูปแบบ 4 ตัวอักษร + 7 ตัวเลข
       - Booking No.            คำว่า BOOKING / BKG ตามด้วยรหัส
       - Vessel / Voyage
       - POD (Port of Discharge)
       - Shipper
       - ขนาดตู้ 20/40 + GP/HQ/RF/UT/TK
"""

from __future__ import annotations

import glob
import os
import re
from typing import Dict, List

try:
    import pdfplumber  # type: ignore
    HAS_PDFPLUMBER = True
except Exception:  # pragma: no cover
    pdfplumber = None
    HAS_PDFPLUMBER = False


CONTAINER_RE = re.compile(r"\b([A-Z]{4})[ \-]?(\d{7})\b")
BOOKING_RE = re.compile(
    r"(?:BOOKING|BKG|B/?KG|BOOKING\s*NO|BKG\s*NO)[\s.:#]*([A-Z0-9\-]{6,})",
    re.IGNORECASE,
)
VOYAGE_RE = re.compile(r"(?:VOY|VOYAGE|V\.)[\s.:#]*([0-9]{2,4}[A-Z]?)", re.IGNORECASE)
VESSEL_RE = re.compile(
    r"(?:VESSEL|VSL|M\.?V\.?|BY\s+VESSEL|ชื่อเรือ)[\s.:#]*"
    r"([A-Z][A-Z0-9\-'. ]{2,40}?)(?:\s+(?:VOY|V\.|VOYAGE)\b|$)",
    re.IGNORECASE,
)
POD_RE = re.compile(
    r"(?:POD|PORT\s*OF\s*DISCHARGE|DISCHARGE\s*PORT|DISCH\s*PORT)[\s.:#]*"
    r"([A-Za-z][A-Za-z0-9\-,'/() ]{2,40})",
    re.IGNORECASE,
)
SHIPPER_RE = re.compile(
    r"(?:SHIPPER|SHIPPER\s*NAME|ผู้ส่งออก)[\s.:#]*"
    r"([A-Za-z0-9][A-Za-z0-9\-,.'&/() ]{3,60})",
    re.IGNORECASE,
)
SIZE_RE = re.compile(r"\b(20|40|45)\s?'?\s?(GP|DC|HQ|HC|RF|RH|UT|OT|TK|FR|PL)\b", re.IGNORECASE)

STATUS_KEYWORDS = {
    "FULL": ("FULL", "F/L", "LADEN", "ตู้หนัก", "บรรจุ"),
    "EMPTY": ("EMPTY", "MT", "ตู้เปล่า"),
}


def extract_text_from_pdf(path: str) -> str:
    """ดึงข้อความจาก PDF หนึ่งไฟล์ (ทุกหน้า) - ไม่มี AI"""
    if not HAS_PDFPLUMBER:
        raise RuntimeError("ยังไม่ได้ติดตั้ง pdfplumber (pip install pdfplumber)")
    parts: List[str] = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            if txt:
                parts.append(txt)
    return "\n".join(parts)


def _clean(s: str) -> str:
    return re.sub(r"\s{2,}", " ", s.strip(" .:-\t")).strip()


def parse_text(text: str) -> Dict:
    """วิเคราะห์ข้อความด้วยกฎล้วน ๆ -> dict ค่าที่เดาได้"""
    upper = text.upper()

    containers: List[Dict[str, str]] = []
    seen = set()
    for m in CONTAINER_RE.finditer(upper):
        cn = f"{m.group(1)}{m.group(2)}"
        if cn in seen:
            continue
        seen.add(cn)
        # หา size/status ที่อยู่บรรทัดใกล้ ๆ
        window = upper[max(0, m.start() - 60): m.end() + 60]
        size_m = SIZE_RE.search(window)
        size = ""
        if size_m:
            code = size_m.group(2).upper()
            code = {"DC": "GP", "HC": "HQ", "RH": "RF"}.get(code, code)
            size = f"{size_m.group(1)} {code}"
        status = ""
        for key, kws in STATUS_KEYWORDS.items():
            if any(k in window for k in kws):
                status = key
                break
        containers.append({"container": cn, "size": size, "status": status})

    def first(regex):
        m = regex.search(text)
        return _clean(m.group(1)) if m else ""

    booking = first(BOOKING_RE)
    voyage = first(VOYAGE_RE)
    vessel = first(VESSEL_RE)
    pod = first(POD_RE)
    shipper = first(SHIPPER_RE)

    # size รวมของทั้งเอกสาร (ถ้าตู้ไม่ได้ระบุราย ๆ)
    doc_size = ""
    sm = SIZE_RE.search(upper)
    if sm:
        code = sm.group(2).upper()
        code = {"DC": "GP", "HC": "HQ", "RH": "RF"}.get(code, code)
        doc_size = f"{sm.group(1)} {code}"

    for c in containers:
        if not c["size"]:
            c["size"] = doc_size

    return {
        "vessel": vessel,
        "voy": voyage,
        "shipper": shipper,
        "pod": pod,
        "booking": booking,
        "containers": containers,
    }


def _merge(dst: Dict, src: Dict) -> None:
    for k in ("vessel", "voy", "shipper", "pod", "booking"):
        if not dst.get(k) and src.get(k):
            dst[k] = src[k]
    have = {c["container"] for c in dst["containers"]}
    for c in src.get("containers", []):
        if c["container"] not in have:
            dst["containers"].append(c)
            have.add(c["container"])


def parse_folder(pdf_dir: str) -> Dict:
    """
    อ่านทุกไฟล์ .pdf ใน pdf_dir แล้วรวมผลการวิเคราะห์เป็นชุดเดียว
    คืน: {"ok":bool, "files":[...], "suggestions":{...}, "error":str}
    """
    if not HAS_PDFPLUMBER:
        return {
            "ok": False,
            "files": [],
            "suggestions": {},
            "error": "ยังไม่ได้ติดตั้ง pdfplumber - รัน: pip install pdfplumber",
        }

    files = sorted(glob.glob(os.path.join(pdf_dir, "*.pdf")))
    if not files:
        return {
            "ok": False,
            "files": [],
            "suggestions": {},
            "error": f"ไม่พบไฟล์ .pdf ในโฟลเดอร์ {pdf_dir}",
        }

    merged: Dict = {
        "vessel": "", "voy": "", "shipper": "", "pod": "", "booking": "",
        "containers": [],
    }
    done: List[str] = []
    for f in files:
        try:
            text = extract_text_from_pdf(f)
            _merge(merged, parse_text(text))
            done.append(os.path.basename(f))
        except Exception as exc:
            print(f"  [pdf_extract] อ่าน {f} ไม่ได้: {exc}")

    return {"ok": True, "files": done, "suggestions": merged, "error": ""}
