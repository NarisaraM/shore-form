# -*- coding: utf-8 -*-
"""
scrub_templates.py
ลบ "ข้อมูลตัวอย่าง / ข้อมูลส่วนบุคคล" (เบอร์โทรเจ้าหน้าที่, ตารางตัวอย่าง, โน้ตเก่า)
ออกจากไฟล์แม่แบบต้นฉบับที่รากโปรเจกต์ *ในที่เดิม* โดยคงรูปแบบไฟล์ (.xls / .xlsx) และเลย์เอาต์ไว้

ล้างเฉพาะ "ค่าในเซลล์" ตามช่วงที่ระบุ ไม่ลบแถว ไม่แตะหัวตาราง/โครงฟอร์ม
ทำให้เวอร์ชัน Python (run.py) ยังทำงานเหมือนเดิมทุกประการ

ใช้ Microsoft Excel ผ่าน COM (ต้องมี Excel ติดตั้ง)
รันครั้งเดียว:  python scrub_templates.py
"""

from __future__ import annotations

import os
import sys

import win32com.client as win32

BASE = os.path.dirname(os.path.abspath(__file__))

# xlExcel8 = .xls (97-2003), xlOpenXMLWorkbook = .xlsx
FMT = {".xls": 56, ".xlsx": 51}

# ไฟล์ -> { ชื่อชีต : [ ช่วงเซลล์ที่ต้องล้างค่า ] }
PLAN = {
    "A0-SHORE.xls": {
        # เบอร์โทรเจ้าหน้าที่ + แถวตัวอย่างเที่ยวเรือเก่า + โน้ตที่มีเบอร์มือถือเอเย่น
        "Sheet1": ["C29:C33", "D30:K43", "G27"],
    },
    "B3-SHORE.xls": {
        "CHORE CY": ["A37"],                # OOY 02-0596275
    },
    "B5C3-SHORE.xls": {
        "Sheet1": ["A2:I9"],                     # ตารางตัวอย่าง
        "DataImport": ["A39:A43", "A60:A70"],    # เบอร์โทรเจ้าหน้าที่ + รายชื่อเรือตัวอย่าง
    },
    "A3C1C2-HUTCHISON.xls": {
        "Sheet1": ["A2:I9"],              # ตารางตัวอย่าง
        "HPT": ["A29"],                   # โน้ต SHORE CHANGE VESSEL ...
    },
}


def main():
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        for fname, sheets in PLAN.items():
            path = os.path.join(BASE, fname)
            if not os.path.isfile(path):
                print(f"  [!] ไม่พบ {fname} - ข้าม")
                continue
            ext = os.path.splitext(fname)[1].lower()
            wb = excel.Workbooks.Open(os.path.abspath(path))
            cleared = 0
            for sheet_name, ranges in sheets.items():
                try:
                    ws = wb.Worksheets(sheet_name)
                except Exception:
                    print(f"  [!] {fname}: ไม่พบชีต {sheet_name!r}")
                    continue
                for rng in ranges:
                    for c in ws.Range(rng):
                        if c.Value is not None:
                            cleared += 1
                        try:
                            if c.MergeCells:
                                c.MergeArea.ClearContents()
                            else:
                                c.ClearContents()
                        except Exception:
                            try:
                                c.Value = None
                            except Exception:
                                pass
            wb.Save()  # บันทึกทับในรูปแบบเดิม
            wb.Close(SaveChanges=False)
            print(f"  {fname}: ล้าง {cleared} เซลล์")
    finally:
        excel.Quit()
    print("\nเสร็จแล้ว - ไฟล์แม่แบบต้นฉบับถูก scrub ในที่เดิม")


if __name__ == "__main__":
    if not sys.platform.startswith("win"):
        sys.exit("ต้องรันบน Windows ที่มี Microsoft Excel")
    main()
