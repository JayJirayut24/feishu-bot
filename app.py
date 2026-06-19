import os
import json
import re
import requests
import threading
from flask import Flask, request, jsonify
from io import BytesIO
from datetime import datetime, timezone, timedelta
from openpyxl import load_workbook

app = Flask(__name__)

# ---------------------------------------------------------
# ตั้งค่า App ID และ App Secret ของ Feishu
# ---------------------------------------------------------
APP_ID = "cli_aabfaea8b0619bfc"
APP_SECRET = "3emUt5KWwH01BlhIKADP2bCb5C062oxt"

# ---------------------------------------------------------
# ตั้งค่า Feishu Wiki Spreadsheet ปลายทาง
# ---------------------------------------------------------
WIKI_TOKEN = "UbCZwapNyiN15YkEKADcFyUHnWf"
TARGET_SHEET_NAMES = ["ยิงส่ง - ITCBI", "ยิงถึง - ITCBI"]
BRANCH_CODE_SHEET = "ยิงส่ง - ITCBI"

# ---------------------------------------------------------
# Bangkok Timezone (UTC+7)
# ---------------------------------------------------------
BKK_TZ = timezone(timedelta(hours=7))

# ---------------------------------------------------------
# ตัวนับรายวัน (รีเซ็ตทุกเที่ยงคืน 00:00)
# ---------------------------------------------------------
daily_counter = {"date": None, "count": 0}


def get_daily_count():
    """เช็คและรีเซ็ตตัวนับถ้าข้ามวัน"""
    today = datetime.now(BKK_TZ).strftime("%d/%m/%Y")
    if daily_counter["date"] != today:
        daily_counter["date"] = today
        daily_counter["count"] = 0
    return daily_counter


def add_to_daily_count(count):
    """เพิ่มจำนวนเข้าตัวนับรายวัน"""
    dc = get_daily_count()
    dc["count"] += count
    return dc["count"], dc["date"]


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
# ฟังก์ชันดึงข้อมูลจากไฟล์ Excel
# ---------------------------------------------------------
def extract_data_from_excel(file_bytes):
    """
    อ่านไฟล์ Excel (.xlsx) และดึงข้อมูล 3 ประเภท:
    1. เลข AWB 13 หลัก (เช่น 7989935047501)
    2. รหัส B ตามด้วยตัวเลข (เช่น B28999230)
    3. รหัสสาขา 6 หลัก (เช่น 811146)
    """
    try:
        wb = load_workbook(filename=BytesIO(file_bytes), data_only=True, read_only=True)
        awb_list = []
        branch_codes = []

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row in ws.iter_rows(values_only=True):
                for cell in row:
                    if cell is None:
                        continue

                    # แปลงค่าเซลล์เป็น string
                    cell_str = str(cell).strip()

                    # ถ้าเป็นตัวเลขทศนิยม (เช่น 7989935047501.0) ให้ตัด .0 ออก
                    if isinstance(cell, float) and cell == int(cell):
                        cell_str = str(int(cell))

                    # เงื่อนไข 1: ตัวเลข 12 หลักพอดี → AWB
                    if re.match(r'^\d{12}$', cell_str):
                        awb_list.append(cell_str)
                    # เงื่อนไข 2: ขึ้นต้นด้วย B ตามด้วยตัวเลข → AWB
                    elif re.match(r'^B\d+$', cell_str):
                        awb_list.append(cell_str)
                    # เงื่อนไข 3: ตัวเลข 6 หลักพอดี → รหัสสาขา
                    elif re.match(r'^\d{6}$', cell_str):
                        branch_codes.append(cell_str)

        wb.close()

        # ลบ AWB ซ้ำ (เก็บลำดับเดิม)
        seen_awb = set()
        unique_awb = []
        for awb in awb_list:
            if awb not in seen_awb:
                seen_awb.add(awb)
                unique_awb.append(awb)

        # รหัสสาขาไม่ลบซ้ำ (เพราะ 1 รหัสสาขาอาจมีหลายพัสดุ)
        return unique_awb, branch_codes

    except Exception:
        return [], []


# ---------------------------------------------------------
# ฟังก์ชันดึง Spreadsheet Token จาก Wiki Token
# ---------------------------------------------------------
def get_spreadsheet_token(token):
    """ดึง obj_token (Spreadsheet Token จริง) จาก Wiki Node"""
    url = f"https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token={WIKI_TOKEN}"
    headers = {"Authorization": f"Bearer {token}"}
    try:
        res = requests.get(url, headers=headers).json()
        if res.get("code") != 0:
            return None, f"API Error: {res.get('msg')} (Code: {res.get('code')})"
        node = res.get("data", {}).get("node", {})
        return node.get("obj_token"), None
    except Exception as e:
        return None, f"Request Error: {str(e)}"


# ---------------------------------------------------------
# ฟังก์ชันดึง Sheet ID ของ Sheet ที่ต้องการ
# ---------------------------------------------------------
def get_sheet_ids(spreadsheet_token, token):
    """ค้นหา Sheet ID ของ Sheet ชื่อ 'ยิงส่ง - ITCBI' และ 'ยิงถึง - ITCBI'"""
    url = f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/query"
    headers = {"Authorization": f"Bearer {token}"}
    try:
        res = requests.get(url, headers=headers).json()
        sheets = res.get("data", {}).get("sheets", [])

        sheet_ids = {}
        for sheet in sheets:
            title = sheet.get("title", "")
            sheet_id = sheet.get("sheet_id", "")
            if title in TARGET_SHEET_NAMES:
                sheet_ids[title] = sheet_id

        return sheet_ids
    except Exception:
        return {}


# ---------------------------------------------------------
# ฟังก์ชันเพิ่มข้อมูลลงใน Feishu Spreadsheet
# ---------------------------------------------------------
def append_to_feishu_sheet(spreadsheet_token, sheet_id, data_list, column, token):
    """เพิ่มข้อมูลต่อท้ายคอลัมน์ที่ระบุของ Sheet"""
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values_append"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json; charset=utf-8"
    }

    # จัดข้อมูลเป็นแถวๆ สำหรับคอลัมน์ที่ระบุ
    values = [[item] for item in data_list]

    payload = {
        "valueRange": {
            "range": f"{sheet_id}!{column}:{column}",
            "values": values
        }
    }

    params = {"insertDataOption": "INSERT_ROWS"}

    try:
        res = requests.post(url, headers=headers, json=payload, params=params).json()
        return res
    except Exception as e:
        return {"code": -1, "msg": str(e)}


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
# ฟังก์ชันจัดการ Event แบบ Asynchronous
# ---------------------------------------------------------
def process_event(data):
    """ประมวลผล Event ใน Background Thread เพื่อป้องกัน Feishu Timeout"""
    if "header" not in data or "event" not in data:
        return

    event_type = data["header"].get("event_type")

    if event_type == "im.message.receive_v1":
        message = data["event"].get("message", {})
        message_type = message.get("message_type")
        message_id = message.get("message_id")

        # ถ้าเป็นไฟล์
        if message_type == "file":
            content = json.loads(message.get("content", "{}"))
            file_key = content.get("file_key")
            file_name = content.get("file_name", "")

            if file_name.endswith(".xlsx"):
                token = get_tenant_access_token()

                # ตอบกลับว่ากำลังทำงาน
                reply_message(message_id, f"⏳ กำลังประมวลผลไฟล์ '{file_name}'...", token)

                # ดาวน์โหลดไฟล์
                file_bytes = download_feishu_file(message_id, file_key, token)

                if file_bytes:
                    # ดึงข้อมูลจาก Excel
                    awb_list, branch_codes = extract_data_from_excel(file_bytes)

                    if not awb_list and not branch_codes:
                        reply_message(
                            message_id,
                            "⚠️ ไม่พบข้อมูล AWB (12 หลัก / รหัส B) หรือรหัสสาขา (6 หลัก) ในไฟล์นี้ครับ",
                            token
                        )
                        return

                    # ดึง Spreadsheet Token จาก Wiki
                    spreadsheet_token, error_msg = get_spreadsheet_token(token)

                    if not spreadsheet_token:
                        reply_message(
                            message_id,
                            f"❌ ไม่สามารถเข้าถึง Wiki Spreadsheet ได้\n"
                            f"สาเหตุ: {error_msg}\n"
                            "กรุณาตรวจสอบว่า JIRAYUTBOT มีสิทธิ์เข้าถึง Wiki และ Sheet แล้ว",
                            token
                        )
                        return

                    # ดึง Sheet ID
                    sheet_ids = get_sheet_ids(spreadsheet_token, token)

                    if not sheet_ids:
                        reply_message(
                            message_id,
                            "❌ ไม่พบ Sheet 'ยิงส่ง - ITCBI' หรือ 'ยิงถึง - ITCBI'\n"
                            "กรุณาตรวจสอบชื่อ Sheet ใน Feishu Spreadsheet",
                            token
                        )
                        return

                    results = []

                    # === AWB → คอลัมน์ A ของทั้ง 2 Sheet ===
                    if awb_list:
                        for sheet_name, sheet_id in sheet_ids.items():
                            res = append_to_feishu_sheet(
                                spreadsheet_token, sheet_id, awb_list, "A", token
                            )
                            code = res.get("code", -1)
                            if code == 0:
                                results.append(f"✅ {sheet_name}: เพิ่ม AWB {len(awb_list)} รายการ")
                            else:
                                msg = res.get("msg", "Unknown error")
                                results.append(f"❌ {sheet_name}: {msg}")

                    # === รหัสสาขา → คอลัมน์ F ของ "ยิงส่ง - ITCBI" เท่านั้น ===
                    if branch_codes and BRANCH_CODE_SHEET in sheet_ids:
                        sheet_id = sheet_ids[BRANCH_CODE_SHEET]
                        res = append_to_feishu_sheet(
                            spreadsheet_token, sheet_id, branch_codes, "F", token
                        )
                        code = res.get("code", -1)
                        if code == 0:
                            results.append(
                                f"✅ รหัสสาขา: เพิ่ม {len(branch_codes)} รายการ → ยิงส่ง - ITCBI (คอลัมน์ F)"
                            )
                        else:
                            msg = res.get("msg", "Unknown error")
                            results.append(f"❌ รหัสสาขา: {msg}")

                    # === อัปเดตตัวนับรายวัน ===
                    total_awb = len(awb_list)
                    daily_total, today_date = add_to_daily_count(total_awb)

                    # นับแยกประเภท
                    count_12 = sum(1 for a in awb_list if re.match(r'^\d{12}$', a))
                    count_b = sum(1 for a in awb_list if re.match(r'^B\d+$', a))

                    # === สร้างข้อความสรุป ===
                    summary = (
                        f"📊 ผลการประมวลผลไฟล์ '{file_name}'\n"
                        f"━━━━━━━━━━━━━━━━━━━━\n"
                        f"🔢 จำนวน AWB: {count_12} เลข 12 หลัก และ {count_b} รหัส B "
                        f"ที่กรอกลง feishu.cn\n"
                        f"📅 จำนวน AWB วันที่ {today_date} รวม {daily_total}\n"
                    )

                    if branch_codes:
                        summary += f"🏢 รหัสสาขา: {len(branch_codes)} รายการ\n"

                    summary += (
                        f"━━━━━━━━━━━━━━━━━━━━\n"
                        + "\n".join(results) + "\n"
                        f"━━━━━━━━━━━━━━━━━━━━\n"
                        f"✅ ประมวลผลเสร็จสิ้น!"
                    )

                    reply_message(message_id, summary, token)
                else:
                    reply_message(
                        message_id,
                        "❌ ดาวน์โหลดไฟล์ไม่สำเร็จ กรุณาลองใหม่อีกครั้ง",
                        token
                    )
            else:
                token = get_tenant_access_token()
                reply_message(
                    message_id,
                    f"⚠️ รองรับเฉพาะไฟล์นามสกุล .xlsx เท่านั้นครับ\n"
                    f"ไฟล์ที่ส่งมา: {file_name}",
                    token
                )

        # ถ้าเป็นข้อความปกติ
        elif message_type == "text":
            token = get_tenant_access_token()
            reply_message(
                message_id,
                "👋 สวัสดีครับ! ส่งไฟล์ Excel (.xlsx) มาได้เลย\n"
                "ผมจะดึงเลข AWB (12 หลัก), รหัส B และรหัสสาขา (6 หลัก)\n"
                "แล้วบันทึกลง Feishu Sheet ให้อัตโนมัติครับ!",
                token
            )

# ---------------------------------------------------------
# Webhook Endpoint (รับ Event จาก Feishu)
# ---------------------------------------------------------
@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json

    # ขั้นตอนที่ 1: ยืนยัน URL Challenge
    if "challenge" in data:
        return jsonify({"challenge": data["challenge"]})

    # ขั้นตอนที่ 2: รันการประมวลผลใน Background Thread
    thread = threading.Thread(target=process_event, args=(data,))
    thread.start()

    # ตอบกลับ Feishu ทันทีว่า "รับทราบแล้ว" เพื่อไม่ให้ Feishu ส่งซ้ำ
    return jsonify({"status": "ok"})


# ---------------------------------------------------------
# Health Check
# ---------------------------------------------------------
@app.route("/", methods=["GET"])
def health():
    return "JIRAYUTBOT is running! 🤖"


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
