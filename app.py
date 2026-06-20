# -*- coding: utf-8 -*-
"""
input_data.py
-------------
ระบบดึงข้อมูลจาก Feishu Sheet แล้วคีย์เข้า PDA ผ่าน ADB
**ไม่มีระบบตรวจจับ popup ใดๆ ทั้งสิ้น** เพื่อให้คีย์ข้อมูลได้เร็วที่สุด

การตรวจ popup เป็นหน้าที่ของ recheck.py (รันแยก process)
input_data.py แค่ "ฟังสัญญาณ" จาก recheck.py ผ่านไฟล์สัญญาณใน common.py:
  - ถ้าได้สัญญาณ PAUSE  -> หยุดคีย์ทันที รอคำสั่งต่อ
  - RESUME:CLEAR_DELIVERY -> ลบช่อง delivery แล้วคีย์ A ตัวถัดไปในกลุ่มเดิม
  - RESUME:SKIP_GROUP     -> ข้ามทั้งกลุ่มงานนี้ ไปขึ้นกลุ่มใหม่

รันผ่าน command line:
    python input_data.py
หรือใส่ session cookie ใน env var FEISHU_SESSION
"""
import sys
import time
import datetime
import io
from pathlib import Path

# บังคับให้ stdout/stderr เป็น UTF-8 กันภาษาไทยเพี้ยน/error บน Windows cmd
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    # สำหรับ Python เก่าที่ไม่มี reconfigure
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

from common import (
    COORDS,
    STATE_FILE, CMD_FILE, STOP_FILE, SCAN_COUNTER_FILE,
    write_signal, read_signal, clear_signal,
    adb_tap, adb_clear_field, adb_broadcast_text,
    run_cmd, read_feishu_rows, FEISHU_SHEET_URL, ADB_PATH,
)
from upload_notify import upload_to_cloudinary, send_dingtalk_links


def log(msg):
    """พิมพ์ log ออกหน้าจอ command line พร้อม timestamp"""
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


# =========================================================
# ปรับความเร็วตรงนี้ได้ง่ายๆ (หน่วยเป็นวินาที)
# DELAY_AFTER_CONFIRM = เวลารอให้แอป PDA ประมวลผลหลังกดยืนยัน
#   - ค่าน้อย = เร็วขึ้น แต่ถ้าน้อยไปแอปอาจรับเลขถัดไปไม่ทัน (ช่องไม่ทันเคลียร์)
#   - ลองไล่จาก 0.15 -> 0.10 -> 0.08 ถ้า PDA ยังตามทัน
#   - ถ้าเริ่มมีเลขตกหล่น/คีย์ทับ ให้เพิ่มกลับขึ้น
# =========================================================
DELAY_AFTER_CONFIRM = 0.25


class InputDataRunner:
    def __init__(self):
        self.running = True
        # เคลียร์สัญญาณเก่าทิ้งก่อนเริ่ม
        clear_signal(CMD_FILE)
        clear_signal(STOP_FILE)
        write_signal(STATE_FILE, "IDLE")

    # -----------------------------------------------------
    # ตรวจสอบว่าควรหยุดทั้งระบบหรือยัง (กด stop)
    # -----------------------------------------------------
    def should_stop(self) -> bool:
        if STOP_FILE.exists():
            self.running = False
            return True
        return False

    # -----------------------------------------------------
    # คอยจังหวะที่ recheck.py สั่ง PAUSE
    # คืนค่า action ที่ recheck สั่งมา: "CLEAR_DELIVERY" / "SKIP_GROUP" / "" (ไม่มี popup)
    # -----------------------------------------------------
    def check_pause(self) -> str:
        """
        อ่านคำสั่งจาก recheck.py
        - ถ้าเป็น PAUSE: รอจนกว่า recheck จะปิด popup เสร็จและส่ง RESUME:<action> กลับมา
        - คืน action string ให้ลูปหลักจัดการต่อ
        - ถ้าไม่มีคำสั่ง: คืน "" (ทำงานปกติ)
        """
        cmd = read_signal(CMD_FILE)
        if not cmd:
            return ""

        if cmd == "PAUSE":
            log("   ⏸️  ได้รับสัญญาณ PAUSE จาก recheck.py — หยุดรอ recheck ปิด popup...")
            write_signal(STATE_FILE, "PAUSED")

            # รอ recheck ปิด popup เสร็จ แล้วเปลี่ยนคำสั่งเป็น RESUME:<action>
            while self.running:
                if self.should_stop():
                    return ""
                cur = read_signal(CMD_FILE)
                if cur.startswith("RESUME:"):
                    action = cur.split(":", 1)[1].strip()
                    clear_signal(CMD_FILE)  # เคลียร์คำสั่งทิ้ง พร้อมรับรอบใหม่
                    write_signal(STATE_FILE, "RUNNING")
                    log(f"   ▶️  recheck ปิด popup เสร็จ -> สั่งให้ทำต่อแบบ: {action}")
                    return action
                time.sleep(0.1)
            return ""

        # เผื่อกรณีมี RESUME ค้างอยู่โดยไม่ได้ PAUSE ก่อน -> เคลียร์ทิ้ง
        if cmd.startswith("RESUME:"):
            action = cmd.split(":", 1)[1].strip()
            clear_signal(CMD_FILE)
            return action

        return ""

    # -----------------------------------------------------
    # ลบข้อมูลเก่าในช่อง delivery (วิธีเดิมจาก src.py)
    # -----------------------------------------------------
    def clear_delivery_field(self):
        log("   🧹 ลบข้อมูลเก่าในช่อง delivery (วิธีเดิม src.py)")
        adb_tap(*COORDS["delivery_no"])
        time.sleep(0.05)
        adb_clear_field()
        time.sleep(0.05)

    # -----------------------------------------------------
    # คีย์ F + G (เปิดกลุ่มงานใหม่) = STEP 1
    # -----------------------------------------------------
    def key_group_header(self, f_data, g_data):
        # พิมพ์ F และ G แค่ครั้งเดียวต่อ 1 กลุ่ม  (เหมือน src.py)
        if f_data:
            adb_tap(*COORDS["open_input"])
            time.sleep(0.9)
            adb_tap(*COORDS["station_next"])
            time.sleep(0.4)
            adb_broadcast_text(f_data)
            time.sleep(0.4)
            adb_tap(*COORDS["db_suggestion"])
            time.sleep(0.2)
            adb_tap(*COORDS["select_branch"])
            time.sleep(0.2)

        if g_data:
            adb_tap(*COORDS["task_code"])
            time.sleep(0.01)
            adb_clear_field()
            time.sleep(0.01)
            adb_broadcast_text(g_data)
            time.sleep(0.01)

    # -----------------------------------------------------
    # คีย์เลขพัสดุ (คอลัมน์ A) + กดยืนยัน แบบ Fast Combo
    # tap -> ADB_CLEAR_TEXT -> ADB_INPUT_TEXT -> tap confirm  รวบเป็น shell เดียว
    # -----------------------------------------------------
    def key_delivery(self, a_data):
        if a_data:
            t0 = time.time()
            safe_text = str(a_data).strip()
            dx, dy = COORDS["delivery_no"]
            cx, cy = COORDS["confirm_btn"]
            cmd = (
                f"input tap {int(dx)} {int(dy)} && "
                f"am broadcast -a ADB_CLEAR_TEXT && "
                f"am broadcast -a ADB_INPUT_TEXT --es msg '{safe_text}' && "
                f"input tap {int(cx)} {int(cy)}"
            )
            run_cmd([ADB_PATH, "shell", cmd])
            time.sleep(DELAY_AFTER_CONFIRM)
            elapsed = time.time() - t0
            log(f"   ⚡ ป้อน+ยืนยัน {elapsed:.2f}s")

    # -----------------------------------------------------
    # ลูปการทำงานหลัก
    # -----------------------------------------------------
    def run(self):
        try:
            rows = read_feishu_rows()
            total = len(rows)
            log(f"พบข้อมูลทั้งหมด {total} แถว")

            # -------- จัดกลุ่มข้อมูลตาม F และ G (เหมือน src.py) --------
            grouped_data = {}
            skipped_rows = []
            for row in rows:
                g = row["G"]
                if not g or not g.upper().startswith("ZXZB"):
                    skipped_rows.append(row["A"])
                    continue
                key = (row["F"], row["G"])
                grouped_data.setdefault(key, []).append(row["A"])

            if skipped_rows:
                log(f"⚠️  ข้าม {len(skipped_rows)} แถวที่ไม่มีเลขงาน (G ไม่พบ/0): "
                    f"{', '.join(str(x) for x in skipped_rows[:5])}"
                    f"{'...' if len(skipped_rows) > 5 else ''}")
            log(f"จัดกลุ่มได้ทั้งหมด {len(grouped_data)} กลุ่มคำสั่งงาน\n")
            write_signal(STATE_FILE, "RUNNING")

            idx = 0
            # -------- วนทีละกลุ่ม --------
            for (f_data, g_data), a_list in grouped_data.items():
                if self.should_stop():
                    break

                log("=====================================")
                log(f"📌 [เริ่มกลุ่มใหม่] สาขา: {f_data} | เลขงาน: {g_data}")
                log(f"   (มีรายการที่ต้องคีย์ทั้งหมด {len(a_list)} รายการ)")
                log("=====================================")

                # ก่อนขึ้นกลุ่มใหม่ เช็คว่ามี popup ค้างจากแถวก่อนหน้าไหม (non-blocking)
                pre = self.peek_recheck()
                if pre == "SKIP_GROUP":
                    log("   Popup -> skip group before step1")
                    self.clear_delivery_field()
                    continue
                elif pre == "CLEAR_DELIVERY":
                    log("   🧹 เคลียร์ popup ค้างก่อนขึ้นกลุ่มใหม่เรียบร้อย")
                    self.clear_delivery_field()

                # STEP 1: เปิดกลุ่มงานใหม่
                self.key_group_header(f_data, g_data)

                # หลัง STEP 1 ตรวจ popup (non-blocking)
                step1_action = self.peek_recheck()
                if step1_action == "SKIP_GROUP":
                    log(f"   ⏭️ popup หลัง STEP 1 -> ข้ามกลุ่ม {g_data} ไปกลุ่มถัดไป")
                    continue

                skip_group = False  # ธงสำหรับ popup 2 (ปิดตู้ ห้ามทำซ้ำ)

                # -------- วนคีย์ A ในกลุ่มนี้ --------
                i = 0
                while i < len(a_list):
                    if self.should_stop():
                        break

                    a_data = a_list[i]
                    idx += 1
                    log(f"\n   -> แถวที่ {idx}/{total} | เลขพัสดุ: {a_data}")

                    # เช็ค PAUSE ก่อนเริ่มคีย์ กันกรณี recheck เพิ่งเจอ popup ของแถวก่อน
                    # (popup อาจยังค้างบนจอ ถ้าคีย์ทับจะพลาด)
                    pre_action = self.peek_recheck()
                    if pre_action == "SKIP_GROUP":
                        log(f"   ⏭️ popup 'ปิดตู้แล้ว' -> ข้ามทั้งกลุ่มงาน {g_data}")
                        skip_group = True
                        break
                    elif pre_action == "CLEAR_DELIVERY":
                        # recheck ลบ delivery ให้แล้ว คีย์ A ตัวนี้ต่อได้เลย (ไม่ +i ไม่ข้าม)
                        log("   🔁 popup จัดการแล้ว -> คีย์เลขพัสดุตัวนี้ต่อ")

                    # คีย์ A + กดยืนยัน (Fast Combo) — คีย์รัวไม่รอ
                    self.key_delivery(a_data)

                    # โหมดเร็วสุด: เช็ค PAUSE แบบแวบเดียว (non-blocking) ไม่หน่วงเลย
                    # ถ้า recheck เจอ popup มันจะ PAUSE ค้างไว้ -> เราจับได้ที่นี่
                    # หรือถ้าจับไม่ทันแถวนี้ ก็จะไปจับที่ pre_action ของแถวถัดไป
                    # (recheck วิ่งตรวจตลอดเวลาอยู่แล้ว ไม่พลาด popup แน่นอน)
                    action = self.peek_recheck()

                    if action == "SKIP_GROUP":
                        # popup 2: ข้ามทั้งกลุ่ม ไปขึ้นกลุ่มใหม่เลย
                        log(f"   ⏭️ popup 'ปิดตู้แล้ว' -> ข้ามทั้งกลุ่มงาน {g_data}")
                        skip_group = True
                        break

                    elif action == "CLEAR_DELIVERY":
                        # popup 1 หรือ 3: recheck ปิด popup + ลบ delivery ให้แล้ว
                        # คีย์ A ตัวถัดไปในกลุ่มเดิม
                        log(f"   ⏭️ ข้ามแถวที่ {idx} (popup) -> ไปเลขพัสดุตัวถัดไป")
                        i += 1
                        continue

                    else:
                        log("   ✓ เสร็จสมบูรณ์")
                        i += 1

                if skip_group:
                    continue  # ไปกลุ่มถัดไป

            log("\n====== อัปโหลดเสร็จสิ้น ======")
            write_signal(STATE_FILE, "DONE")

            if self.running:
                self.take_final_screenshot()

        except ValueError as ve:
            log(f"❌ เกิดข้อผิดพลาด: {ve}")
            write_signal(STATE_FILE, "ERROR")
        except Exception as e:
            log(f"❌ เกิดข้อผิดพลาดที่ไม่คาดคิด: {e}")
            write_signal(STATE_FILE, "ERROR")
        finally:
            self.running = False
            write_signal(STATE_FILE, "DONE")

    # -----------------------------------------------------
    # เช็คสัญญาณ PAUSE จาก recheck.py แบบ "แวบเดียว" (non-blocking)
    # ไม่หน่วงเวลาเลยถ้าไม่มี popup -> นี่คือกุญแจความเร็ว
    #
    # - ถ้า recheck สั่ง PAUSE -> เข้า check_pause() รอ recheck ปิด popup
    #   แล้วคืน action ("CLEAR_DELIVERY"/"SKIP_GROUP")
    # - ถ้าไม่มี PAUSE -> คืน "" ทันที ไปคีย์แถวถัดไปได้เลย
    # -----------------------------------------------------
    def peek_recheck(self) -> str:
        cmd = read_signal(CMD_FILE)
        if cmd == "PAUSE":
            return self.check_pause()
        return ""

    # -----------------------------------------------------
    # รอ recheck ตรวจ popup ให้เสร็จก่อนตัดสินว่าแถวนี้สำเร็จ
    #
    # หลักการ: อ่านเลข scan counter ของ recheck ก่อน แล้วรอจน recheck
    # ตรวจหน้าจอเพิ่มอย่างน้อย 2 รอบ (= มั่นใจว่ามันตรวจ "หลังกดยืนยัน" แล้ว)
    # - ถ้าระหว่างรอ recheck สั่ง PAUSE -> จัดการ popup แล้วคืน action
    # - ถ้าครบ 2 รอบแล้วไม่มี PAUSE -> ไม่มี popup จริง คืน "" ไปต่อได้
    #
    # ใช้เวลาประมาณ POLL_INTERVAL*2 (~0.6 วิ) ต่อแถว ซึ่งจำเป็น
    # เพื่อความถูกต้อง (กันการขึ้นกลุ่มใหม่ทับ popup)
    # -----------------------------------------------------
    def wait_recheck_verdict(self) -> str:
        start_count = self._read_scan_count()
        # ต้องการให้ recheck ตรวจเพิ่มอย่างน้อย 1 รอบหลังจุดนี้
        # (จุดนี้เรียกแค่ตอนขึ้นกลุ่มใหม่/หลัง STEP 1 ไม่ใช่ทุกแถว จึงไม่กระทบความเร็วรวม)
        need_count = start_count + 1
        # timeout กันค้าง (เผื่อ recheck ตาย) — สูงสุด ~3 วิ
        deadline = time.time() + 3.0

        while time.time() < deadline:
            if self.should_stop():
                return ""

            # เจอ PAUSE -> มี popup จัดการเลย
            cmd = read_signal(CMD_FILE)
            if cmd == "PAUSE":
                return self.check_pause()

            # recheck ตรวจครบรอบแล้วและไม่มี popup -> ปลอดภัย ไปต่อ
            if self._read_scan_count() >= need_count:
                return ""

            time.sleep(0.05)

        # timeout: เช็ค PAUSE ครั้งสุดท้ายก่อนยอมไปต่อ
        if read_signal(CMD_FILE) == "PAUSE":
            return self.check_pause()
        return ""

    def _read_scan_count(self) -> int:
        try:
            return int(read_signal(SCAN_COUNTER_FILE) or "0")
        except Exception:
            return 0

    def _old_peek_recheck(self) -> str:
        cmd = read_signal(CMD_FILE)
        if cmd == "PAUSE":
            # มี popup! หยุดรอ recheck ปิด popup + ส่ง action กลับ
            return self.check_pause()
        return ""

    # -----------------------------------------------------
    # ถ่ายภาพหน้าจอตอนจบ (คงจาก src.py)
    # -----------------------------------------------------
    def take_final_screenshot(self):
        save_path = None
        try:
            log("\n📸 กำลังบันทึกภาพหน้าจอ PDA...")
            now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            folder_name = f"photo_pda_{now_str}"

            base_dir = Path(__file__).parent
            save_dir = base_dir / folder_name
            save_dir.mkdir(parents=True, exist_ok=True)

            save_path = save_dir / f"screenshot_{now_str}.png"

            run_cmd([ADB_PATH, "shell", "screencap", "-p", "/sdcard/final_screen.png"])
            run_cmd([ADB_PATH, "pull", "/sdcard/final_screen.png", str(save_path)])
            run_cmd([ADB_PATH, "shell", "rm", "/sdcard/final_screen.png"])

            log(f"✅ บันทึกภาพเสร็จสิ้นที่โฟลเดอร์: {folder_name}")
        except Exception as e:
            log(f"❌ เกิดข้อผิดพลาดในการบันทึกภาพ: {e}")

        # ─── อัปโหลด Cloudinary + แจ้งเตือน DingTalk ───────────────
        if save_path is None or not save_path.exists():
            log("⚠️  ไม่พบไฟล์ภาพ — ข้ามขั้นตอนอัปโหลด")
            return

        try:
            log("\n☁️  กำลังอัปโหลดภาพขึ้น Cloudinary...")
            image_url = upload_to_cloudinary(save_path, resource_type="image")
            log(f"✅ อัปโหลดภาพสำเร็จ: {image_url}")
        except Exception as e:
            log(f"❌ อัปโหลดภาพล้มเหลว: {e}")
            return

        try:
            log("📨 กำลังส่งลิงก์ไปยัง DingTalk...")
            send_dingtalk_links(
                image_url=image_url,
                sheet_url=FEISHU_SHEET_URL,
            )
            log("✅ ส่งแจ้งเตือน DingTalk สำเร็จ")
        except Exception as e:
            log(f"❌ ส่ง DingTalk ล้มเหลว: {e}")


def main():
    log("กำลังเชื่อมต่อ Feishu Sheet...")
    runner = InputDataRunner()
    runner.run()


if __name__ == "__main__":
    main()
