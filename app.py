import os
import json
import requests
from flask import Flask, request, jsonify
from io import BytesIO
from openpyxl import load_workbook

app = Flask(__name__)

# ---------------------------------------------------------
# ตั้งค่า App ID และ App Secret ของ Feishu
# ---------------------------------------------------------
APP_ID = "cli_aabfaea8b0619bfc"
APP_SECRET = "3emUt5KWwH01BlhIKADP2bCb5C062oxt"


# ---------------------------------------------------------
# ฟังก์ชันขอ Token จาก Feishu
# ---------------------------------------------------------
def get_tenant_access_token():
    """ขอ Tenant Access Token จาก Feishu Open API"""
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    payload = {"app_id": APP_ID, "app_secret": APP_SECRET}
    res = requests.post(url, json=payload).json()
    return res.get("tenant_access_token")


# ---------------------------------------------------------
# ฟังก์ชันดาวน์โหลดไฟล์จากแชท Feishu
# ---------------------------------------------------------
def download_feishu_file(message_id, file_key, token):
    """ดาวน์โหลดไฟล์ที่ผู้ใช้ส่งมาในแชท โดยใช้ message_id และ file_key"""
    url = f"https://open.feishu.cn/open-apis/im/v1/messages/{message_id}/resources/{file_key}?type=file"
    headers = {"Authorization": f"Bearer {token}"}
    res = requests.get(url, headers=headers)
    if res.status_code == 200:
        return res.content
    return None


# ---------------------------------------------------------
# ฟังก์ชันวิเคราะห์ไฟล์ Excel
# ---------------------------------------------------------
def analyze_excel(file_bytes):
    """
    อ่านไฟล์ Excel (.xlsx) และสรุปข้อมูล:
    - จำนวนแถวทั้งหมดที่มีข้อมูล
    - จำนวน Sheet ทั้งหมด
    - ชื่อ Sheet แต่ละอัน
    """
    try:
        wb = load_workbook(filename=BytesIO(file_bytes), data_only=True, read_only=True)

        sheet_names = wb.sheetnames
        total_sheets = len(sheet_names)

        # วิเคราะห์ Sheet แรก (Active Sheet)
        ws = wb.active
        total_rows = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if any(cell is not None for cell in row):
                total_rows += 1

        wb.close()

        # สร้างข้อความสรุป
        summary = (
            f"📊 ผลการวิเคราะห์ไฟล์ Excel\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"📄 จำนวน Sheet: {total_sheets} sheet(s)\n"
            f"📋 ชื่อ Sheet: {', '.join(sheet_names)}\n"
            f"📝 จำนวนแถวข้อมูล (Sheet แรก): {total_rows} แถว\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"✅ วิเคราะห์เสร็จสิ้น!"
        )
        return summary

    except Exception as e:
        return f"❌ เกิดข้อผิดพลาดในการอ่าน Excel: {str(e)}"


# ---------------------------------------------------------
# ฟังก์ชันตอบกลับข้อความในแชท
# ---------------------------------------------------------
def reply_message(message_id, text, token):
    """ส่งข้อความตอบกลับในแชท Feishu"""
    url = f"https://open.feishu.cn/open-apis/im/v1/messages/{message_id}/reply"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json; charset=utf-8"
    }
    payload = {
        "msg_type": "text",
        "content": json.dumps({"text": text})
    }
    requests.post(url, headers=headers, json=payload)


# ---------------------------------------------------------
# Webhook Endpoint (รับ Event จาก Feishu)
# ---------------------------------------------------------
@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json

    # ขั้นตอนที่ 1: ยืนยัน URL Challenge (Feishu จะส่งมาครั้งแรกเมื่อผูก Webhook)
    if "challenge" in data:
        return jsonify({"challenge": data["challenge"]})

    # ขั้นตอนที่ 2: จัดการ Event ข้อความ
    if "header" in data and "event" in data:
        event_type = data["header"].get("event_type")

        # เช็คว่าเป็น Event "มีคนส่งข้อความเข้ามา" หรือไม่
        if event_type == "im.message.receive_v1":
            message = data["event"].get("message", {})
            message_type = message.get("message_type")
            message_id = message.get("message_id")

            # ถ้าเป็นไฟล์
            if message_type == "file":
                content = json.loads(message.get("content", "{}"))
                file_key = content.get("file_key")
                file_name = content.get("file_name", "")

                # เช็คว่าเป็นไฟล์ .xlsx หรือไม่
                if file_name.endswith(".xlsx"):
                    token = get_tenant_access_token()

                    # ตอบกลับก่อนว่ากำลังทำงาน
                    reply_message(message_id, f"⏳ กำลังดาวน์โหลดและวิเคราะห์ไฟล์ '{file_name}'...", token)

                    # ดาวน์โหลดไฟล์
                    file_bytes = download_feishu_file(message_id, file_key, token)

                    if file_bytes:
                        # วิเคราะห์ข้อมูลและตอบกลับ
                        result_text = analyze_excel(file_bytes)
                        reply_message(message_id, result_text, token)
                    else:
                        reply_message(message_id, "❌ ดาวน์โหลดไฟล์ไม่สำเร็จ กรุณาลองใหม่อีกครั้ง", token)
                else:
                    token = get_tenant_access_token()
                    reply_message(message_id, f"⚠️ รองรับเฉพาะไฟล์นามสกุล .xlsx เท่านั้นครับ\nไฟล์ที่ส่งมา: {file_name}", token)

            # ถ้าเป็นข้อความปกติ (ไม่ใช่ไฟล์)
            elif message_type == "text":
                content = json.loads(message.get("content", "{}"))
                text = content.get("text", "")

                token = get_tenant_access_token()
                reply_message(message_id, "👋 สวัสดีครับ! ส่งไฟล์ Excel (.xlsx) มาได้เลย\nผมจะวิเคราะห์และสรุปข้อมูลให้ทันทีครับ!", token)

    return jsonify({"status": "ok"})


# ---------------------------------------------------------
# Health Check (สำหรับ Render ตรวจสอบว่าเซิร์ฟเวอร์ยังทำงานอยู่)
# ---------------------------------------------------------
@app.route("/", methods=["GET"])
def health():
    return "JIRAYUTBOT is running! 🤖"


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
