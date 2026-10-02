FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
# มี PORT = โหมดโฮสต์สาธารณะ (จำกัดการส่ง, ไม่เปิด /download, ใช้โฟลเดอร์ชั่วคราวต่อคำขอ)
ENV PORT=8000
CMD ["python", "server.py"]
