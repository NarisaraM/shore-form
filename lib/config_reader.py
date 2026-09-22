# -*- coding: utf-8 -*-
"""
config_reader.py
อ่านไฟล์ Check.xlsx เพื่อดึง "ตัวเลือก" (dropdown) ออกมาเป็นรายการข้อ ๆ
- คอลัมน์ A เป็นหัวข้อ (Terminal / SIZE / STATUS F/E)
- คอลัมน์ B เป็นค่าตัวเลือกของหัวข้อนั้น ๆ ไล่ลงมาจนกว่าจะเจอหัวข้อใหม่

และอ่านไฟล์ VSLNAME.xls (คอลัมน์ A แถวแรกเป็นหัวข้อ "Vessel Name" ที่เหลือเป็นรายชื่อเรือ)
เพื่อดึงตัวเลือกชื่อเรือ

ไม่มีการใช้ AI ใด ๆ ทั้งสิ้น เป็นการอ่านไฟล์ Excel ตรง ๆ ด้วย openpyxl/xlrd
"""

from __future__ import annotations

import os
from typing import Dict, List

import openpyxl
import xlrd

# ชื่อไฟล์คำสั่งตัวเลือก
CHECK_FILE = os.path.join("input", "Check.xlsx")
VSLNAME_FILE = os.path.join("input", "VSLNAME.xls")

# ค่าที่ใช้สำรอง เผื่อเปิดไฟล์ Check.xlsx ไม่ได้ (คัดลอกจากไฟล์จริง ณ วันที่สร้างสคริปต์)
FALLBACK = {
    "terminals": [
        "LCMT Company LTD, ( under LCB1 Group)  A0",
        "ESCO (EASTERN SEA LCH CNTR TML/B3)",
        "B5/C3 LCIT (LAEM CHABANG INTERNATIONAL TERMINAL CO., LTD)",
        "A3 (C1,C2) (Hutchison Laemchabang Terminal Limited, HLT)",
    ],
    "sizes": [
        "20 GP", "40 GP", "40 HQ", "20 RH", "40 RH",
        "20 UT", "40 UT", "20 TK", "40 TK", "20 FR", "40 FR",
    ],
    "statuses": ["FULL", "EMPTY"],
    "vessels": [
        "CA MANILA", "DONGJIN CONFIDENT", "HEUNG-A BANGKOK", "HEUNG-A HOCHIMINH",
        "INCHEON VOYAGER", "KMTC BANGKOK", "KMTC GWANGYANG", "KMTC JAKARTA",
        "KMTC SURABAYA", "KMTC TAIPEIS", "KMTC ULSAN", "KMTC XIAMEN",
        "LAEM CHABANG VOYAGER", "PANCON CHAMPION", "PEGASUS PROTO",
        "SAWASDEE ALTAIR", "SAWASDEE ATLANTIC", "SAWASDEE BALTIC",
        "SAWASDEE CAPELLA", "SAWASDEE DENEB", "SAWASDEE INCHEON",
        "SAWASDEE MIMOSA", "SAWASDEE RIGEL", "SAWASDEE SPICA",
        "SAWASDEE SUNRISE", "SAWASDEE VEGA", "SKY ORION", "STARSHIP JUPITER",
        "TIANJIN BRIDGE", "TS TIANJIN", "TS XIAMEN", "YEOSU VOYAGER",
    ],
}

# คำที่ถือว่าเป็น "หัวข้อ" ในคอลัมน์ A
_HEADER_TERMINAL = ("terminal",)
_HEADER_SIZE = ("size",)
_HEADER_STATUS = ("status",)


def _norm(value) -> str:
    return str(value).strip() if value is not None else ""


def _read_vessels(base_dir: str) -> List[str]:
    path = os.path.join(base_dir, VSLNAME_FILE)
    if not os.path.isfile(path):
        return list(FALLBACK["vessels"])

    try:
        book = xlrd.open_workbook(path)
        sheet = book.sheets()[0]
    except Exception:
        return list(FALLBACK["vessels"])

    vessels: List[str] = []
    for r in range(1, sheet.nrows):  # แถว 0 เป็นหัวข้อ "Vessel Name"
        name = _norm(sheet.cell_value(r, 0)) if sheet.ncols > 0 else ""
        if name:
            vessels.append(name)

    return vessels or list(FALLBACK["vessels"])


def read_config(base_dir: str) -> Dict[str, List[str]]:
    """
    คืนค่า dict:
        {
          "terminals": [...],   # ตัวเลือก TERMINAL
          "sizes":     [...],   # ตัวเลือก SIZE
          "statuses":  [...],   # ตัวเลือก STATUS F/E
        }
    """
    path = os.path.join(base_dir, CHECK_FILE)
    if not os.path.isfile(path):
        return {k: list(v) for k, v in FALLBACK.items()}

    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception:
        return {k: list(v) for k, v in FALLBACK.items()}

    ws = wb.worksheets[0]

    terminals: List[str] = []
    sizes: List[str] = []
    statuses: List[str] = []
    current: List[str] | None = None

    for row in ws.iter_rows(values_only=True):
        col_a = _norm(row[0]) if len(row) > 0 else ""
        col_b = _norm(row[1]) if len(row) > 1 else ""

        if col_a:
            low = col_a.lower()
            if any(h in low for h in _HEADER_TERMINAL):
                current = terminals
            elif any(h in low for h in _HEADER_SIZE):
                current = sizes
            elif any(h in low for h in _HEADER_STATUS):
                current = statuses
            else:
                current = None

        if col_b and current is not None:
            current.append(col_b)

    try:
        wb.close()
    except Exception:
        pass

    result = {
        "terminals": terminals or list(FALLBACK["terminals"]),
        "sizes": sizes or list(FALLBACK["sizes"]),
        "statuses": statuses or list(FALLBACK["statuses"]),
        "vessels": _read_vessels(base_dir),
    }
    return result


if __name__ == "__main__":
    import json

    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    print(json.dumps(read_config(here), ensure_ascii=False, indent=2))
