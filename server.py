# -*- coding: utf-8 -*-
"""
server.py
เว็บเซิร์ฟเวอร์เล็ก ๆ ด้วยไลบรารีมาตรฐานของ Python (ไม่ต้องลง Flask)
ให้บริการหน้าฟอร์ม + API สำหรับอ่าน Check.xlsx / อ่าน PDF / เขียนผลลัพธ์ลง Excel

เส้นทาง (routes):
  GET  /                -> web/index.html
  GET  /api/config      -> ตัวเลือก TERMINAL / SIZE / STATUS จาก Check.xlsx
  GET  /api/mail-status -> ตรวจว่าตั้งค่า SMTP ไว้แล้วหรือยัง
  POST /api/parse-pdf   -> อ่านไฟล์ PDF ในโฟลเดอร์ pdf_in/ แล้วเดาค่ากรอกฟอร์ม
  POST /api/submit      -> เขียนข้อมูลลงแม่แบบ Excel ของ TERMINAL ที่เลือก แล้วส่งอีเมลแนบไฟล์
  GET  /download/<file> -> (สำรอง/สำหรับผู้ดูแลระบบ) ดาวน์โหลดไฟล์ผลลัพธ์จากโฟลเดอร์ output/
"""

from __future__ import annotations

import json
import os
import sys
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WEB_DIR = os.path.join(BASE_DIR, "web")
PDF_DIR = os.path.join(BASE_DIR, "pdf_in")
OUT_DIR = os.path.join(BASE_DIR, "output")
CACHE_DIR = os.path.join(BASE_DIR, "build", "templates")

sys.path.insert(0, BASE_DIR)
from lib.config_reader import read_config          # noqa: E402
from lib.pdf_extract import parse_folder           # noqa: E402
from lib.excel_fill import fill, TERMINALS         # noqa: E402
from lib import mailer                             # noqa: E402

# ผู้รับอีเมลทุกครั้งที่มีการส่งข้อมูล (แก้ได้ตรงนี้)
EMAIL_RECIPIENTS = [
    "dongykong.naris@gmail.com",
    "sirichai@heungaline.co.th",
]


def _json_bytes(obj) -> bytes:
    return json.dumps(obj, ensure_ascii=False).encode("utf-8")


def _build_mail(payload: dict, result: dict):
    """สร้างหัวเรื่อง + เนื้อหาอีเมลจากข้อมูลที่กรอกในฟอร์ม"""
    terminal = payload.get("terminal", "")
    vessel = payload.get("vessel", "")
    voy = payload.get("voy", "")
    shipper = payload.get("shipper", "")
    pod = payload.get("pod", "")
    booking = payload.get("booking", "")
    special_cond = payload.get("remark", "")
    contact_name = payload.get("contactName", "")
    contact_phone = payload.get("contactPhone", "")
    contact_email = payload.get("contactEmail", "")
    contact_line = "Contact: " + " ".join(p for p in (contact_name, contact_phone) if p)
    if contact_email:
        contact_line += " " + contact_email
    rows = payload.get("rows", [])

    subject = f"[SHORE] {result.get('terminal_key','')} - {vessel} V.{voy} - Booking {booking}"

    lines = [
        "มีการส่งข้อมูล SHORE ใหม่ผ่านระบบแบบฟอร์มสำรวจข้อมูล (Self Service Shore)",
        "",
        f"TERMINAL      : {terminal}",
        f"Vessel / Voy. : {vessel} V.{voy}",
        f"Shipper name  : {shipper}",
        f"POD           : {pod}",
        f"Booking No.   : {booking}",
        "",
        f"จำนวนตู้ ({len(rows)} ใบ):",
    ]
    for i, r in enumerate(rows, 1):
        extra = []
        if r.get("commodity"):
            extra.append(f"commodity={r['commodity']}")
        if r.get("temp"):
            extra.append(f"temp={r['temp']}")
        if r.get("vent"):
            extra.append(f"vent={r['vent']}")
        if r.get("dgUn"):
            extra.append(f"DG/UN={r['dgUn']}")
        extra_str = ("  " + "  ".join(extra)) if extra else ""
        lines.append(
            f"  {i}. {r.get('container','')}  size={r.get('size','')}  status={r.get('status','')}{extra_str}"
        )
    lines += ["", f"ผู้ติดต่อ: {contact_line}"]
    if special_cond:
        lines += [f"Special Condition Request: {special_cond}"]
    if result.get("warnings"):
        lines += ["", "หมายเหตุ:"] + [f"  - {w}" for w in result["warnings"]]
    lines += ["", f"ไฟล์แนบ: {result.get('out_file','')}", "", "-- ส่งอัตโนมัติจากระบบ Self Service Shore --"]

    return subject, "\n".join(lines)


class Handler(BaseHTTPRequestHandler):
    server_version = "ShoreForm/1.0"

    # ------------------------------------------------------------------ utils
    def _send(self, code: int, body: bytes, ctype: str = "application/json; charset=utf-8",
              extra_headers: dict | None = None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra_headers or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _send_json(self, obj, code: int = 200):
        self._send(code, _json_bytes(obj))

    def _read_body(self) -> dict:
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0:
            return {}
        raw = self.rfile.read(length)
        try:
            return json.loads(raw.decode("utf-8"))
        except Exception:
            return {}

    def log_message(self, fmt, *args):  # noqa: A003 - ปิด log รก ๆ
        sys.stderr.write("  %s - %s\n" % (self.address_string(), fmt % args))

    # -------------------------------------------------------------------- GET
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path in ("/", "/index.html"):
            return self._serve_file(os.path.join(WEB_DIR, "index.html"),
                                    "text/html; charset=utf-8")

        if path == "/api/config":
            cfg = read_config(BASE_DIR)
            cfg["terminal_keys"] = {name: TERMINALS[name]["key"]
                                    for name in cfg["terminals"] if name in TERMINALS}
            return self._send_json(cfg)

        if path == "/api/mail-status":
            return self._send_json({
                "configured": mailer.is_configured(BASE_DIR),
                "recipients": EMAIL_RECIPIENTS,
            })

        if path.startswith("/download/"):
            name = urllib.parse.unquote(path[len("/download/"):])
            return self._serve_download(name)

        if path == "/logo.png":
            return self._serve_file(os.path.join(BASE_DIR, "logo.png"), "image/png")

        if path.startswith("/web/"):
            return self._serve_file(os.path.join(BASE_DIR, path.lstrip("/")),
                                    self._guess_ctype(path))

        return self._send_json({"error": "not found"}, 404)

    do_HEAD = do_GET

    # ------------------------------------------------------------------- POST
    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == "/api/parse-pdf":
            result = parse_folder(PDF_DIR)
            return self._send_json(result)

        if path == "/api/submit":
            payload = self._read_body()
            try:
                result = fill(payload, BASE_DIR, CACHE_DIR, OUT_DIR)
            except Exception as exc:
                import traceback
                traceback.print_exc()
                return self._send_json({"ok": False, "error": str(exc), "warnings": []}, 500)

            if not result.get("ok"):
                return self._send_json(result, 400)

            # เขียนไฟล์ Excel สำเร็จแล้ว -> ส่งอีเมลแนบไฟล์ (แทนการให้ดาวน์โหลด)
            try:
                subject, body = _build_mail(payload, result)
                mailer.send_excel_email(
                    BASE_DIR,
                    to_list=EMAIL_RECIPIENTS,
                    subject=subject,
                    body_text=body,
                    attachment_path=result["out_path"],
                    attachment_name=result["out_file"],
                )
                result["emailed"] = True
                result["to"] = EMAIL_RECIPIENTS
                return self._send_json(result, 200)
            except Exception as exc:
                # เขียนไฟล์ไว้ที่ output/ แล้ว (เป็นสำรอง) แต่ส่งเมลไม่สำเร็จ -> แจ้งผู้ใช้ตรง ๆ
                result["ok"] = False
                result["emailed"] = False
                result["error"] = str(exc)
                return self._send_json(result, 502)

        return self._send_json({"error": "not found"}, 404)

    # ---------------------------------------------------------------- helpers
    @staticmethod
    def _guess_ctype(path: str) -> str:
        if path.endswith(".css"):
            return "text/css; charset=utf-8"
        if path.endswith(".js"):
            return "application/javascript; charset=utf-8"
        if path.endswith(".html"):
            return "text/html; charset=utf-8"
        return "application/octet-stream"

    def _serve_file(self, fs_path: str, ctype: str):
        if not os.path.isfile(fs_path):
            return self._send_json({"error": f"missing {os.path.basename(fs_path)}"}, 404)
        with open(fs_path, "rb") as fh:
            self._send(200, fh.read(), ctype)

    def _serve_download(self, name: str):
        # กันการทะลุโฟลเดอร์
        safe = os.path.basename(name)
        fs_path = os.path.join(OUT_DIR, safe)
        if not os.path.isfile(fs_path):
            return self._send_json({"error": "ไม่พบไฟล์ผลลัพธ์"}, 404)
        with open(fs_path, "rb") as fh:
            data = fh.read()
        self._send(
            200, data,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            {"Content-Disposition": f'attachment; filename="{safe}"'},
        )


def serve(host: str = "127.0.0.1", port: int = 8000):
    os.makedirs(PDF_DIR, exist_ok=True)
    os.makedirs(OUT_DIR, exist_ok=True)
    os.makedirs(CACHE_DIR, exist_ok=True)
    httpd = ThreadingHTTPServer((host, port), Handler)
    print(f"  ระบบแบบฟอร์มสำรวจข้อมูล SHORE พร้อมใช้งานที่  http://{host}:{port}")
    print("  กด Ctrl+C เพื่อหยุด")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n  หยุดเซิร์ฟเวอร์แล้ว")
        httpd.shutdown()


if __name__ == "__main__":
    p = int(sys.argv[1]) if len(sys.argv) > 1 else 8000
    serve(port=p)
