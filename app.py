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
def append_to_feishu_sheet(spreadsheet_token, sheet_id, values, start_col, token):
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values_append"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    # แปลงข้อมูลเป็น array ของ array (Feishu ต้องการแบบนี้)
    if len(values) > 0 and isinstance(values[0], list):
        sheet_values = [[str(item) for item in row] for row in values]
        max_cols = max(len(row) for row in sheet_values)
        end_col = chr(ord(start_col.upper()) + max_cols - 1)
        range_str = f"{sheet_id}!{start_col}:{end_col}"
    else:
        sheet_values = [[str(v)] for v in values]
        range_str = f"{sheet_id}!{start_col}:{start_col}"

    payload = {
        "valueRange": {
            "range": range_str,
            "values": sheet_values
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

        token = get_tenant_access_token()
        awb_list = []
        branch_codes = []
        is_valid_input = False

        # ---------------------------------------------
        # 1. กรณีเป็นไฟล์
        # ---------------------------------------------
        if message_type == "file":
            content = json.loads(message.get("content", "{}"))
            file_key = content.get("file_key")
            file_name = content.get("file_name", "")

            if file_name.endswith(".xlsx"):
                is_valid_input = True
                reply_message(message_id, f"⏳ กำลังประมวลผลไฟล์ '{file_name}'...", token)

                file_bytes = download_feishu_file(message_id, file_key, token)
                if file_bytes:
                    awb_list, branch_codes = extract_data_from_excel(file_bytes)

                    # ดึงรหัสสาขาจากชื่อไฟล์ (ถ้าใน Excel ไม่มีรหัสสาขา)
                    if not branch_codes:
                        filename_branch_match = re.search(r'(?<!\d)(\d{6})(?!\d)', file_name)
                        if filename_branch_match:
                            branch_code_from_name = filename_branch_match.group(1)
                            if awb_list:
                                branch_codes = [branch_code_from_name] * len(awb_list)
                            else:
                                branch_codes = [branch_code_from_name]
                else:
                    reply_message(message_id, "❌ ดาวน์โหลดไฟล์ไม่สำเร็จ กรุณาลองใหม่อีกครั้ง", token)
                    return
            else:
                reply_message(message_id, f"⚠️ รองรับเฉพาะไฟล์นามสกุล .xlsx เท่านั้นครับ", token)
                return

        # ---------------------------------------------
        # 2. กรณีเป็นข้อความปกติ (พิมพ์เข้ามาในแชท)
        # ---------------------------------------------
        elif message_type == "text":
            content = json.loads(message.get("content", "{}"))
            text = content.get("text", "")

            # ค้นหาเลข 12 หลัก
            awb_list.extend(re.findall(r'(?<!\d)\d{12}(?!\d)', text))
            # ค้นหารหัส B
            awb_list.extend(re.findall(r'\bB\d+\b', text))
            # ค้นหารหัสสาขา 6 หลัก
            branch_codes.extend(re.findall(r'(?<!\d)\d{6}(?!\d)', text))
            
            # ลบ AWB ซ้ำ (แต่เก็บลำดับเดิม)
            seen = set()
            awb_list = [x for x in awb_list if not (x in seen or seen.add(x))]

            # ถ้าระบุรหัสสาขามา 1 ตัว แต่มีหลาย AWB ให้เบิ้ลรหัสสาขาให้เท่ากับ AWB
            if len(branch_codes) == 1 and len(awb_list) > 1:
                branch_codes = [branch_codes[0]] * len(awb_list)

            if awb_list or branch_codes:
                is_valid_input = True
                reply_message(message_id, f"⏳ กำลังบันทึกข้อมูลจากข้อความ...", token)
            else:
                reply_message(
                    message_id,
                    "👋 สวัสดีครับ! ส่งไฟล์ Excel (.xlsx) หรือพิมพ์เลข AWB (12 หลัก), รหัส B, รหัสสาขา (6 หลัก) มาให้ผมได้เลยครับ",
                    token
                )
                return

        # ==========================================
        # เริ่มกระบวนการเขียนลง Feishu Sheet
        # ==========================================
        if not is_valid_input:
            return

        if not awb_list and not branch_codes:
            reply_message(
                message_id,
                "⚠️ ไม่พบข้อมูล AWB (12 หลัก / รหัส B) หรือรหัสสาขา (6 หลัก) เลยครับ",
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

        for sheet_name, sheet_id in sheet_ids.items():
            if sheet_name == BRANCH_CODE_SHEET:
                # "ยิงส่ง - ITCBI": รวม AWB กับรหัสสาขาเป็นบรรทัดเดียวกัน (A ถึง F)
                rows = []
                max_len = max(len(awb_list), len(branch_codes)) if branch_codes else len(awb_list)
                for i in range(max_len):
                    awb = awb_list[i] if i < len(awb_list) else ""
                    br = branch_codes[i] if branch_codes and i < len(branch_codes) else ""
                    if br:
                        # [Col A, Col B, Col C, Col D, Col E, Col F]
                        rows.append([awb, "", "", "", "", br])
                    else:
                        rows.append([awb]) # เขียนแค่คอลัมน์ A
                
                if rows:
                    res = append_to_feishu_sheet(spreadsheet_token, sheet_id, rows, "A", token)
                    code = res.get("code", -1)
                    if code == 0:
                        results.append(f"✅ {sheet_name}: บันทึกข้อมูลสำเร็จ")
                    else:
                        msg = res.get("msg", "Unknown error")
                        results.append(f"❌ {sheet_name}: {msg}")
            else:
                # "ยิงถึง - ITCBI": มีแค่ AWB
                if awb_list:
                    res = append_to_feishu_sheet(spreadsheet_token, sheet_id, awb_list, "A", token)
                    code = res.get("code", -1)
                    if code == 0:
                        results.append(f"✅ {sheet_name}: บันทึกข้อมูลสำเร็จ")
                    else:
                        msg = res.get("msg", "Unknown error")
                        results.append(f"❌ {sheet_name}: {msg}")

        # === อัปเดตตัวนับรายวัน ===
        total_awb = len(awb_list)
        daily_total, today_date = add_to_daily_count(total_awb)

        # === สร้างข้อความสรุป ===
        summary = (
            f"📅 จำนวน AWB วันที่ {today_date} รวม {daily_total}\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
        )

        # ถ้ามีข้อผิดพลาด ให้แสดงด้วย
        errors = [msg for msg in results if "❌" in msg]
        if errors:
            summary += "\n".join(errors) + "\n━━━━━━━━━━━━━━━━━━━━\n"

        summary += "✅ ประมวลผลเสร็จสิ้น!"

        reply_message(message_id, summary, token)

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
