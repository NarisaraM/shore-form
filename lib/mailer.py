# -*- coding: utf-8 -*-
"""
mailer.py
ส่งไฟล์ Excel ที่กรอกเสร็จแล้วออกเป็นอีเมล (แนบไฟล์) ผ่าน SMTP

ตั้งค่าบัญชีที่ใช้ส่งได้ 2 ทาง (อย่างใดอย่างหนึ่ง หรือผสมกันได้ - ตัวแปรแวดล้อมมาก่อน):

  1) ตัวแปรแวดล้อม (environment variables):
       SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD,
       SMTP_FROM, SMTP_FROM_NAME, SMTP_USE_TLS (1/0)

  2) ไฟล์ smtp_config.json ที่รากโปรเจกต์ (ไม่ถูก commit ขึ้น git)
     คัดลอกจาก smtp_config.example.json แล้วกรอกข้อมูลจริง เช่น:
       {
         "host": "smtp.gmail.com",
         "port": 587,
         "use_tls": true,
         "username": "youraccount@gmail.com",
         "password": "รหัสผ่านแอป (App Password) 16 หลัก",
         "from_addr": "youraccount@gmail.com",
         "from_name": "SHORE Self Service"
       }

ไม่มีการเก็บรหัสผ่านไว้ในโค้ดหรือใน git แต่อย่างใด
"""

from __future__ import annotations

import json
import os
import smtplib
import ssl
from email.message import EmailMessage
from typing import Dict, List, Optional

CONFIG_FILE = "smtp_config.json"

XLSX_MIME = (
    "application",
    "vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)


def _truthy(value, default: bool) -> bool:
    if value is None or value == "":
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() not in ("0", "false", "no")


def _load_config(base_dir: str) -> Dict:
    # None = "ไม่ได้ตั้งค่าตัวนี้" (ต่างจาก "" ที่แปลว่าตั้งเป็นค่าว่างชัดเจน)
    cfg = {
        "host": os.environ.get("SMTP_HOST") or None,
        "port": os.environ.get("SMTP_PORT") or None,
        "use_tls": os.environ.get("SMTP_USE_TLS"),
        "username": os.environ.get("SMTP_USER") or None,
        "password": os.environ.get("SMTP_PASSWORD") or None,
        "from_addr": os.environ.get("SMTP_FROM") or None,
        "from_name": os.environ.get("SMTP_FROM_NAME") or None,
    }

    path = os.path.join(base_dir, CONFIG_FILE)
    if os.path.isfile(path):
        try:
            with open(path, encoding="utf-8") as f:
                file_cfg = json.load(f)
        except Exception:
            file_cfg = {}
        for k, v in file_cfg.items():
            if cfg.get(k) is None:
                cfg[k] = v

    cfg["port"] = int(cfg["port"]) if cfg.get("port") else 587
    cfg["use_tls"] = _truthy(cfg.get("use_tls"), default=True)
    cfg["from_name"] = cfg.get("from_name") or "SHORE Self Service"
    for k in ("host", "username", "password", "from_addr"):
        cfg[k] = cfg.get(k) or ""
    return cfg


def is_configured(base_dir: str) -> bool:
    cfg = _load_config(base_dir)
    return bool(cfg.get("host") and cfg.get("username") and cfg.get("password") and cfg.get("from_addr"))


def send_excel_email(
    base_dir: str,
    to_list: List[str],
    subject: str,
    body_text: str,
    attachment_path: str,
    attachment_name: str,
    cc_list: Optional[List[str]] = None,
) -> None:
    """
    ส่งอีเมลแนบไฟล์ Excel หนึ่งไฟล์ ให้ to_list (และ cc_list ถ้ามี)
    โยน RuntimeError ที่มีข้อความอ่านง่ายถ้ายังไม่ได้ตั้งค่า SMTP หรือส่งไม่สำเร็จ
    """
    cfg = _load_config(base_dir)
    if not (cfg.get("host") and cfg.get("username") and cfg.get("password") and cfg.get("from_addr")):
        raise RuntimeError(
            "ยังไม่ได้ตั้งค่าการส่งอีเมล (SMTP) — คัดลอก smtp_config.example.json "
            "เป็น smtp_config.json แล้วกรอกข้อมูลบัญชีที่ใช้ส่งให้ครบ (ดู README.md)"
        )
    if not to_list:
        raise RuntimeError("ไม่มีผู้รับอีเมล (to_list ว่าง)")
    if not os.path.isfile(attachment_path):
        raise RuntimeError(f"ไม่พบไฟล์ที่จะแนบ: {attachment_path}")

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = f'{cfg["from_name"]} <{cfg["from_addr"]}>'
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg.set_content(body_text)

    with open(attachment_path, "rb") as f:
        data = f.read()
    msg.add_attachment(data, maintype=XLSX_MIME[0], subtype=XLSX_MIME[1], filename=attachment_name)

    all_rcpt = list(to_list) + list(cc_list or [])
    host = cfg["host"]
    port = int(cfg["port"])

    try:
        if port == 465:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(host, port, context=context, timeout=20) as server:
                server.login(cfg["username"], cfg["password"])
                server.send_message(msg, to_addrs=all_rcpt)
        else:
            with smtplib.SMTP(host, port, timeout=20) as server:
                server.ehlo()
                if cfg.get("use_tls", True):
                    server.starttls(context=ssl.create_default_context())
                    server.ehlo()
                server.login(cfg["username"], cfg["password"])
                server.send_message(msg, to_addrs=all_rcpt)
    except Exception as exc:
        raise RuntimeError(f"ส่งอีเมลไม่สำเร็จ: {exc}") from exc
