# -*- coding: utf-8 -*-
"""
run.py  -  จุดเริ่มต้นใช้งาน (ดับเบิลคลิก หรือ `python run.py`)

สิ่งที่ทำ:
  1) ตรวจ/ติดตั้งไลบรารีที่จำเป็น (openpyxl, xlrd, pdfplumber, pywin32) ถ้ายังไม่มี
  2) แปลงไฟล์แม่แบบ .xls -> .xlsx เก็บไว้ใน build/templates/ (ทำครั้งเดียว)
  3) เปิดเว็บเซิร์ฟเวอร์ที่ http://127.0.0.1:8000 แล้วเปิดเบราว์เซอร์ให้อัตโนมัติ
"""

from __future__ import annotations

import os
import subprocess
import sys
import threading
import time
import webbrowser

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PORT = 8000

REQUIRED = [
    ("openpyxl", "openpyxl"),
    ("xlrd", "xlrd"),
    ("pdfplumber", "pdfplumber"),
]
if sys.platform.startswith("win"):
    REQUIRED.append(("win32com.client", "pywin32"))


def ensure_packages():
    missing = []
    for module_name, pip_name in REQUIRED:
        try:
            __import__(module_name)
        except Exception:
            missing.append(pip_name)
    if not missing:
        return
    print(f"  กำลังติดตั้งไลบรารีที่ขาด: {', '.join(missing)}")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", *missing])
    print("  ติดตั้งเสร็จแล้ว")


def prebuild_templates():
    """แปลง .xls ทุกไฟล์ล่วงหน้า เพื่อไม่ให้ผู้ใช้รอตอนกดส่งข้อมูลครั้งแรก"""
    from lib.excel_fill import TERMINALS
    from lib.xls_convert import ensure_xlsx

    cache_dir = os.path.join(BASE_DIR, "build", "templates")
    for conf in TERMINALS.values():
        src = os.path.join(BASE_DIR, conf["src"])
        if not os.path.isfile(src):
            print(f"  [!] ไม่พบไฟล์แม่แบบ {conf['src']} - ข้ามไปก่อน")
            continue
        try:
            ensure_xlsx(src, cache_dir)
        except Exception as exc:
            print(f"  [!] เตรียมแม่แบบ {conf['src']} ไม่สำเร็จ: {exc}")


def open_browser_later(url: str):
    time.sleep(1.2)
    try:
        webbrowser.open(url)
    except Exception:
        pass


def main():
    os.chdir(BASE_DIR)
    sys.path.insert(0, BASE_DIR)

    ensure_packages()

    print("  เตรียมไฟล์แม่แบบ Excel ...")
    prebuild_templates()

    from server import serve

    url = f"http://127.0.0.1:{PORT}"
    threading.Thread(target=open_browser_later, args=(url,), daemon=True).start()
    serve(port=PORT)


if __name__ == "__main__":
    main()
