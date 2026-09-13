# -*- coding: utf-8 -*-
"""
build_docs.py
สร้างไฟล์สำหรับเวอร์ชัน static (GitHub Pages) ไว้ในโฟลเดอร์ docs/

ทำสิ่งต่อไปนี้:
  1) แปลงไฟล์แม่แบบ .xls -> .xlsx สด ๆ ผ่าน Microsoft Excel (lib/xls_convert)
     - ไฟล์ .xls ต้นฉบับผ่าน scrub_templates.py มาแล้ว (ลบข้อมูลตัวอย่าง/เบอร์โทรเจ้าหน้าที่)
       จึงไม่ต้องล้างซ้ำตรงนี้อีก และไม่ต้องตัดรูปภาพ/ฟอร์แมตใด ๆ ออก
       -> ไฟล์ที่ได้จึงคงฟอนต์/สี/เส้นขอบ/ความกว้างคอลัมน์/รูปภาพ เหมือนต้นฉบับ 100%
       (docs/index.html เขียนค่าลงไฟล์ด้วยการ "ผ่าตัด XML" ผ่าน JSZip ไม่ใช้ SheetJS/ExcelJS
        อ่าน-สร้างสมุดงานใหม่ทั้งเล่ม จึงไม่ทำลายฟอร์แมตเดิม)
  2) คัดลอกเข้า docs/templates/ โดยตั้งชื่อไม่มีช่องว่าง
  3) คัดลอก Check.xlsx เข้า docs/
  4) วางไฟล์ .nojekyll

รันครั้งเดียว (หรือทุกครั้งที่แก้ไฟล์แม่แบบ .xls/.xlsx):  python build_docs.py
"""

from __future__ import annotations

import os
import shutil

from lib.xls_convert import ensure_xlsx

BASE = os.path.dirname(os.path.abspath(__file__))
INPUT = os.path.join(BASE, "input")
DOCS = os.path.join(BASE, "docs")
DOCS_TPL = os.path.join(DOCS, "templates")
CACHE = os.path.join(BASE, "build", "templates")

# ไฟล์แม่แบบต้นฉบับ -> ชื่อไฟล์ใน docs/templates/ (ไม่มีช่องว่าง)
TEMPLATES = {
    "A0-SHORE.xls": "A0-SHORE.xlsx",
    "B3-SHORE.xls": "B3-SHORE.xlsx",
    "B5C3-SHORE.xls": "B5C3-SHORE.xlsx",
    "A2-FORM  A.xlsx": "A2-FORM-A.xlsx",
    "A3C1C2-HUTCHISON.xls": "A3C1C2-HUTCHISON.xlsx",
}


def main():
    os.makedirs(DOCS_TPL, exist_ok=True)

    for src_name, out_name in TEMPLATES.items():
        src = os.path.join(INPUT, src_name)
        if not os.path.isfile(src):
            print(f"  [!] ไม่พบ input/{src_name} - ข้าม")
            continue

        xlsx = ensure_xlsx(src, CACHE, force=True)  # แปลงสดทุกครั้งเพื่อความชัวร์
        dst = os.path.join(DOCS_TPL, out_name)
        shutil.copyfile(xlsx, dst)
        print(f"  {out_name}: คัดลอกแล้ว ({os.path.getsize(dst)/1024:.0f} KB, เก็บฟอร์แมตต้นฉบับครบ)")

    # Check.xlsx
    shutil.copyfile(os.path.join(INPUT, "Check.xlsx"), os.path.join(DOCS, "Check.xlsx"))
    print("  คัดลอก Check.xlsx")

    # .nojekyll (กัน GitHub Pages ประมวลผลแบบ Jekyll)
    open(os.path.join(DOCS, ".nojekyll"), "w").close()
    print("  สร้าง docs/.nojekyll")

    print("\nเสร็จแล้ว - เนื้อหา static อยู่ใน docs/")


if __name__ == "__main__":
    main()
