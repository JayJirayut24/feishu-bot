import os
import json
import re
import unicodedata
import requests
import threading
from flask import Flask, request, jsonify
from io import BytesIO
from datetime import datetime, timezone, timedelta
from openpyxl import load_workbook

app = Flask(__name__)

# ---------------------------------------------------------
# ป้องกัน Feishu ส่ง Event ซ้ำ (Deduplication)
# ---------------------------------------------------------
processed_event_ids = set()
processed_event_lock = threading.Lock()

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
# ระบบเก็บ State จำนวน AWB แบบรายวัน และ รหัสสาขา
# ---------------------------------------------------------
STATE_FILE = "state.json"
app_state = {
    "current_date": "",
    "daily_count": 0,
    "branch_summary": {}
}

# โหลดข้อมูลเก่าถ้ามี
if os.path.exists(STATE_FILE):
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            app_state = json.load(f)
    except:
        pass

def save_state():
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(app_state, f, ensure_ascii=False)
    except:
        pass

def add_to_daily_count(new_count, branch_list):
    tz = timezone(timedelta(hours=7))
    now = datetime.now(tz)
    today_str = now.strftime("%d/%m/%Y")
    
    # ถ้าระบบข้ามวันแล้ว ให้รีเซ็ตตัวนับและสรุปยอดเป็น 0
    if app_state.get("current_date") != today_str:
        app_state["current_date"] = today_str
        app_state["daily_count"] = 0
        app_state["branch_summary"] = {}
        
    app_state["daily_count"] += new_count
    
    # นับยอดแต่ละรหัสสาขา
    for br in branch_list:
        if br:
            app_state["branch_summary"][br] = app_state["branch_summary"].get(br, 0) + 1
            
    save_state()
    return app_state["daily_count"], today_str


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
def _cell_to_str(cell):
    if cell is None:
        return ""
    if isinstance(cell, float) and cell == int(cell):
        return str(int(cell))
    return str(cell).strip()

def _normalize(text):
    """Normalize Unicode + ลบ whitespace เพื่อเปรียบเทียบ header ภาษาไทยได้แม่นยำ"""
    return unicodedata.normalize("NFC", "".join(text.split()))

def extract_data_from_excel(file_bytes):
    try:
        wb = load_workbook(filename=BytesIO(file_bytes), data_only=True, read_only=True)
        awb_list = []
        branch_codes = []

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            all_rows = list(ws.iter_rows(values_only=True))
            if not all_rows:
                continue

            # === ค้นหา header "สาขาที่ถูกต้อง" ใน 5 แถวแรก ===
            TARGET = _normalize("สาขาที่ถูกต้อง")
            correct_branch_col = None
            awb_col = None
            data_start_row = 0

            for row_idx, row in enumerate(all_rows[:5]):
                if not row:
                    continue
                for i, cell in enumerate(row):
                    h = _normalize(_cell_to_str(cell))
                    if TARGET in h:
                        correct_branch_col = i
                    if h.upper() == "AWB":
                        awb_col = i
                if correct_branch_col is not None:
                    data_start_row = row_idx + 1
                    break

            if correct_branch_col is not None:
                # โหมดพิเศษ: จับคู่ AWB กับสาขาที่ถูกต้องตามคอลัมน์
                for row in all_rows[data_start_row:]:
                    if not row:
                        continue

                    awb_str = ""
                    if awb_col is not None and awb_col < len(row):
                        candidate = _cell_to_str(row[awb_col])
                        if re.match(r'^\d{12}$', candidate) or re.match(r'^B\d+$', candidate):
                            awb_str = candidate

                    if not awb_str:
                        for j, cell in enumerate(row):
                            if j == correct_branch_col:
                                continue
                            candidate = _cell_to_str(cell)
                            if re.match(r'^\d{12}$', candidate) or re.match(r'^B\d+$', candidate):
                                awb_str = candidate
                                break

                    if not awb_str:
                        continue

                    branch_str = ""
                    if correct_branch_col < len(row):
                        candidate = _cell_to_str(row[correct_branch_col])
                        if re.match(r'^\d{6}$', candidate):
                            branch_str = candidate

                    awb_list.append(awb_str)
                    branch_codes.append(branch_str)

                # ลบ AWB ซ้ำ (พร้อม branch ที่จับคู่)
                seen = set()
                deduped_awb, deduped_br = [], []
                for awb, br in zip(awb_list, branch_codes):
                    if awb not in seen:
                        seen.add(awb)
                        deduped_awb.append(awb)
                        deduped_br.append(br)
                awb_list, branch_codes = deduped_awb, deduped_br

            else:
                # โหมดปกติ: ข้ามแถวแรก (header) แล้ว pre-scan หาคอลัมน์ที่ค่า 6 หลักเหมือนกันทุกแถว
                col_six = {}
                for row in all_rows[1:]:
                    if not row:
                        continue
                    for j, cell in enumerate(row):
                        v = _cell_to_str(cell)
                        if v and re.match(r'^\d{6}$', v):
                            col_six.setdefault(j, []).append(v)

                # คอลัมน์ที่ทุกค่าเหมือนกัน (เช่น 812087 ทั้งหมด) = สาขาผิด → ข้าม
                skip_cols = {j for j, vals in col_six.items()
                             if len(vals) > 1 and len(set(vals)) == 1}

                for row in all_rows[1:]:
                    if not row:
                        continue
                    for j, cell in enumerate(row):
                        if j in skip_cols:
                            continue
                        cell_str = _cell_to_str(cell)
                        if not cell_str:
                            continue
                        if re.match(r'^\d{12}$', cell_str):
                            awb_list.append(cell_str)
                        elif re.match(r'^B\d+$', cell_str):
                            awb_list.append(cell_str)
                        elif re.match(r'^\d{6}$', cell_str):
                            branch_codes.append(cell_str)

                # ลบ AWB ซ้ำ
                seen = set()
                unique_awb = []
                for awb in awb_list:
                    if awb not in seen:
                        seen.add(awb)
                        unique_awb.append(awb)
                awb_list = unique_awb

        wb.close()
        return awb_list, branch_codes

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
# ฟังก์ชันล้างข้อมูลทุกแถวใน Sheet (ลบแถวจริง ไม่ใช่แค่ล้าง)
# ---------------------------------------------------------
def clear_all_sheet_columns(spreadsheet_token, sheet_ids_dict, token):
    auth_headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    errors = []

    for sheet_name, sheet_id in sheet_ids_dict.items():
        # 1. อ่านคอลัมน์ A เพื่อหาแถวสุดท้ายที่มีข้อมูล
        read_url = (
            f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/"
            f"{spreadsheet_token}/values/{sheet_id}!A2:A5000"
        )
        last_data_row = 0
        try:
            read_res = requests.get(
                read_url, headers={"Authorization": f"Bearer {token}"}
            ).json()
            raw = read_res.get("data", {}).get("valueRange", {}).get("values") or []
            for i, row in enumerate(raw):
                if row and row[0] not in (None, "", "None"):
                    last_data_row = i + 1  # relative index (1-based from row 2)
        except Exception:
            last_data_row = 0

        if last_data_row == 0:
            continue

        # 2. ลบแถวด้วย DELETE /dimension_range
        # startIndex=1 = row 2 (0-based), endIndex=last_data_row+1 (exclusive)
        delete_url = (
            f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/"
            f"{spreadsheet_token}/dimension_range"
        )
        payload = {
            "dimension": {
                "sheetId": sheet_id,
                "majorDimension": "ROWS",
                "startIndex": 2,               # 1-based: row 2 = แถวข้อมูลแรก (ข้ามหัวข้อ)
                "endIndex": last_data_row + 2  # exclusive, 1-based
            }
        }
        try:
            res = requests.delete(delete_url, headers=auth_headers, json=payload)
            try:
                data = res.json()
                if data.get("code", -1) != 0:
                    errors.append(
                        f"{sheet_name}: code={data.get('code')} msg={data.get('msg', '?')}"
                    )
            except ValueError:
                if res.status_code not in (200, 204):
                    errors.append(f"{sheet_name}: HTTP {res.status_code}")
        except Exception as e:
            errors.append(f"{sheet_name}: {str(e)}")

    if errors:
        return False, " | ".join(errors)
    return True, ""


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

            # --- ฟีเจอร์ "ลบข้อมูล" (Clear/Reset) ---
            clean_text = text.lower()
            if any(cmd in clean_text for cmd in ["clear", "ลบข้อมูล", "reset"]):
                try:
                    _spreadsheet_token, _error = get_spreadsheet_token(token)
                    if not _spreadsheet_token:
                        reply_message(message_id, f"❌ ไม่สามารถเข้าถึง Spreadsheet ได้: {_error}", token)
                        return

                    _sheet_ids = get_sheet_ids(_spreadsheet_token, token)
                    if not _sheet_ids:
                        reply_message(message_id, "❌ ไม่พบ Sheet ปลายทาง", token)
                        return

                    ok, err = clear_all_sheet_columns(_spreadsheet_token, _sheet_ids, token)

                    # รีเซ็ตตัวนับรายวัน
                    app_state["current_date"] = ""
                    app_state["daily_count"] = 0
                    app_state["branch_summary"] = {}
                    save_state()

                    if ok:
                        msg = "🗑️ ล้างข้อมูลทุกคอลัมน์เรียบร้อยแล้วครับ\n✅ ยิงส่ง - ITCBI, ยิงถึง - ITCBI\nรีเซ็ตยอดรายวันเป็น 0 แล้ว"
                    else:
                        msg = f"❌ ล้างข้อมูลไม่สำเร็จ: {err}\nรีเซ็ตยอดรายวันเป็น 0 แล้ว"

                    reply_message(message_id, msg, token)

                except Exception as e:
                    reply_message(message_id, f"❌ เกิดข้อผิดพลาดในระบบลบข้อมูล: {str(e)}", token)

                return

            # --- ฟีเจอร์ "สรุปยอด" ---
            if "สรุปยอด" in text:
                tz = timezone(timedelta(hours=7))
                today_str = datetime.now(tz).strftime("%d/%m/%Y")
                
                if app_state.get("current_date") != today_str:
                    reply_message(message_id, f"📊 สรุปยอดวันที่ {today_str}\nยังไม่มีข้อมูลในระบบครับ", token)
                    return
                
                summary_text = f"📊 สรุปยอดวันที่ {today_str}\n"
                summary_text += "━━━━━━━━━━━━━━━━━━━━\n"
                
                branch_summary = app_state.get("branch_summary", {})
                if branch_summary:
                    # เรียงตามรหัสสาขา
                    for br in sorted(branch_summary.keys()):
                        summary_text += f"🏢 สาขา {br} : {branch_summary[br]} รายการ\n"
                else:
                    summary_text += "ไม่มีข้อมูลรหัสสาขา\n"
                    
                summary_text += "━━━━━━━━━━━━━━━━━━━━\n"
                summary_text += f"รวมทั้งหมด: {app_state.get('daily_count', 0)} รายการ\n"
                
                reply_message(message_id, summary_text, token)
                return

            # ตรวจพบลูกน้ำ (,) ให้แจ้ง Error และหยุดการทำงานทันที
            if "," in text:
                reply_message(
                    message_id,
                    f"⚠️ ตรวจพบเครื่องหมาย ( , ) ในข้อความ\nERROR กรุณาลบ , ออกแล้วส่งใหม่ครับ",
                    token
                )
                return

            # ประมวลผลทีละบรรทัด ป้องกันบรรทัดเหลื่อมกัน
            lines = text.split("\n")
            current_branch = ""
            
            for line in lines:
                line_awbs = []
                # ค้นหาเลข 12 หลัก
                line_awbs.extend(re.findall(r'(?<!\d)\d{12}(?!\d)', line))
                # ค้นหารหัส B
                line_awbs.extend(re.findall(r'\bB\d+\b', line))
                # ค้นหารหัสสาขา 6 หลัก
                line_branches = re.findall(r'(?<!\d)\d{6}(?!\d)', line)

                if not line_awbs and not line_branches:
                    continue

                # ลบ AWB ซ้ำในบรรทัดเดียวกัน (แต่เก็บลำดับเดิม)
                seen = set()
                line_awbs = [x for x in line_awbs if not (x in seen or seen.add(x))]

                if line_branches:
                    # ถ้าเจอสาขา ให้จำสาขาล่าสุดเอาไว้เผื่อบรรทัดถัดไปไม่มี
                    current_branch = line_branches[-1]
                elif line_awbs:
                    # มี "ยิงไป/ยิงถึง" แต่ไม่มีรหัส 6 หลัก = ชื่อสาขา → เว้นว่าง
                    if "ยิงไป" in line or "ยิงถึง" in line:
                        line_branches = []
                        current_branch = ""
                    elif current_branch:
                        # ไม่มีคำระบุทิศทาง → สืบทอด current_branch จากบรรทัดก่อน
                        line_branches = [current_branch] * len(line_awbs)

                # ถ้าระบุรหัสสาขามา 1 ตัว แต่มีหลาย AWB ในบรรทัดนี้ ให้เบิ้ลรหัสสาขาให้เท่ากับ AWB
                if len(line_branches) == 1 and len(line_awbs) > 1:
                    line_branches = [line_branches[0]] * len(line_awbs)

                # ถ้าบรรทัดนี้มีแค่ "รหัสสาขา" (ไม่มี AWB) เราจะไม่สร้างบรรทัดว่างๆ ใน Sheet! (แค่จำไว้เฉยๆ)
                if not line_awbs:
                    continue

                # จับคู่ AWB กับ Branch เข้าด้วยกัน เพื่อป้องกันบรรทัดเหลื่อมข้ามบรรทัด
                max_len = max(len(line_awbs), len(line_branches))
                for i in range(max_len):
                    awb_list.append(line_awbs[i] if i < len(line_awbs) else "")
                    branch_codes.append(line_branches[i] if i < len(line_branches) else "")

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
        today_time_str = datetime.now(timezone(timedelta(hours=7))).strftime("%d/%m/%Y %H:%M:%S")

        for sheet_name, sheet_id in sheet_ids.items():
            if sheet_name == BRANCH_CODE_SHEET:
                # "ยิงส่ง - ITCBI": รวม AWB กับรหัสสาขาเป็นบรรทัดเดียวกัน (A ถึง H)
                rows = []
                max_len = max(len(awb_list), len(branch_codes)) if branch_codes else len(awb_list)
                for i in range(max_len):
                    awb = awb_list[i] if i < len(awb_list) else ""
                    br = branch_codes[i] if branch_codes and i < len(branch_codes) else ""
                    if br:
                        # [Col A, Col B, Col C, Col D, Col E, Col F, Col G, Col H]
                        rows.append([awb, "", "", "", "", br, "", today_time_str])
                    else:
                        rows.append([awb, "", "", "", "", "", "", today_time_str]) # เขียนแค่คอลัมน์ A และลงเวลาที่ H
                
                if rows:
                    res = append_to_feishu_sheet(spreadsheet_token, sheet_id, rows, "A", token)
                    code = res.get("code", -1)
                    if code == 0:
                        results.append(f"✅ {sheet_name}: บันทึกข้อมูลสำเร็จ")
                    else:
                        msg = res.get("msg", "Unknown error")
                        results.append(f"❌ {sheet_name}: {msg}")
            else:
                # "ยิงถึง - ITCBI": มีแค่ AWB และวันที่
                if awb_list:
                    rows = []
                    for awb in awb_list:
                        # [Col A, Col B, Col C, Col D, Col E, Col F]
                        rows.append([awb, "", "", "", "", today_time_str])
                    res = append_to_feishu_sheet(spreadsheet_token, sheet_id, rows, "A", token)
                    code = res.get("code", -1)
                    if code == 0:
                        results.append(f"✅ {sheet_name}: บันทึกข้อมูลสำเร็จ")
                    else:
                        msg = res.get("msg", "Unknown error")
                        results.append(f"❌ {sheet_name}: {msg}")

        # === อัปเดตตัวนับรายวัน ===
        total_awb = len(awb_list)
        daily_total, today_date = add_to_daily_count(total_awb, branch_codes)

        # === สร้างข้อความสรุป ===
        summary = (
            f"📅 จำนวน AWB ของไฟล์นี้ : {total_awb}\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"📅 AWB วันที่ {today_date} รวม {daily_total}\n"
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

    # ขั้นตอนที่ 2: ตรวจสอบ event_id ไม่ให้ประมวลผลซ้ำ
    event_id = data.get("header", {}).get("event_id")
    if event_id:
        with processed_event_lock:
            if event_id in processed_event_ids:
                return jsonify({"status": "ok"})
            processed_event_ids.add(event_id)
            # เก็บแค่ 1000 event ล่าสุด ป้องกัน memory leak
            if len(processed_event_ids) > 1000:
                oldest = next(iter(processed_event_ids))
                processed_event_ids.discard(oldest)

    # ขั้นตอนที่ 3: รันการประมวลผลใน Background Thread
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
