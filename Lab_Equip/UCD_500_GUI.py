# -*- coding: utf-8 -*-
"""
Created on Mon Mar 24 13:43:18 2025

@author: CC.Cheng
"""
import time
from pywinauto import Application, Desktop
from datetime import datetime
import xlwings as xw

# -------- CONFIG --------
lane_counts = [ "2"] # ["1", "2", "4"]
bitrates = ["1.62", "2.70", "5.40", "8.10"] # ["1.62", "2.70", "5.40", "6.75", "8.10"]

# -------- CONNECT TO UCD GUI --------
def connect_to_ucd():
    windows = Desktop(backend="uia").windows(title_re="UCD.*")
    if not windows:
        raise RuntimeError("UCD GUI window not found. Is it open?")
    target_win = windows[0]
    app = Application(backend="uia").connect(handle=target_win.handle)
    return app.window(handle=target_win.handle)

# -------- VERIFY LT BY READING STATUS PANEL --------
def verify_lt_by_bitrate_safe(main_win, expected_rate: str) -> str:
    import time
    time.sleep(1.5)
    try:
        text_controls = main_win.descendants(control_type="Text")

        for i in range(len(text_controls) - 1):
            try:
                label = text_controls[i].window_text().strip().lower()
                if label.startswith("bit rate"):
                    try:
                        value = text_controls[i + 1].window_text().strip().lower()
                        try:
                            expected_float = float(expected_rate)
                            actual_float = float(value.replace("gbps", "").strip())

                            if abs(expected_float - actual_float) < 0.01:
                                return f"✅ PASS (Bitrate matched: {value})"
                            else:
                                return f"❌ FAIL (Expected: {expected_rate}, Got: {value})"
                        except Exception:
                            return f"❌ FAIL (Unable to parse Bitrate: '{value}')"

                    except Exception:
                        return "❌ FAIL (Bitrate label found, but couldn't read value)"
            except Exception:
                continue

        return "❌ FAIL (Bitrate label not found)"
    except Exception as e:
        return f"❌ ERROR verifying bitrate: {e}"

# -------- MAIN --------
main_win = connect_to_ucd()
radios = main_win.descendants(control_type="RadioButton")

# Manually mapped indexes from your radio buttons
lane_radio_buttons = {
    # "1": radios[8],
    "2": radios[9],
    # "4": radios[10]
}

bitrate_radio_buttons = {
    "1.62": radios[11],
    "2.70": radios[12],
    "5.40": radios[13],
    # "6.75": radios[14],
    "8.10": radios[15]
}

results = []

for lane in lane_counts:
    for rate in bitrates:
        print(f"\n⚙️ Setting Lane = {lane}, Bitrate = {rate} Gbps")
        try:
            lane_radio_buttons[lane].select()
            bitrate_radio_buttons[rate].select()
            time.sleep(0.5)

            # Click "Link Training"
            main_win.child_window(title="Link Training", control_type="Button").click_input()
            print("🔄 Link Training triggered...")
            time.sleep(1.5)

            result = verify_lt_by_bitrate_safe(main_win, expected_rate=rate)
            print(result)
            results.append((lane, rate, result))

        except Exception as e:
            print(f"❌ Failed: {e}")
            results.append((lane, rate, f"Error: {e}"))

# Save to EXCEL
book = xw.Book()
sheet = book.sheets[0]
sheet.range("A1").value=results
today_datetime = datetime.today().strftime('%m%d%Y_%H%M%S')
filename_o = r"C:\Users\cc.cheng\Downloads\testtest_%s.xlsx"%(today_datetime)
book.save(filename_o)
book.close()
