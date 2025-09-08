# -*- coding: utf-8 -*-
"""
Created on Mon Mar 24 13:43:18 2025

@author: CC.Cheng
"""
import time
from pywinauto import Application, Desktop
from datetime import datetime
import xlwings as xw
from aardvark_controller import AardvarkController
from Agilent_E3631A import Agilent_E3631A

# -------- CONFIG --------
lane_counts = ["2"]  # ["1", "2", "4"]
bitrates = ["5.40"]  # ["1.62", "2.70", "5.40", "6.75", "8.10"]
# Part 2
# dir_pwr = {1.254:'1p14',1.322:'1p2',1.384:'1p26'}
# dir_pwr_RD = {1.194:'1p14',1.254:'1p2',1.314:'1p26'}
# Part 25
dir_pwr = {1.269: '1p14', 1.337: '1p2', 1.4: '1p26'}
dir_pwr_RD = {1.199: '1p14', 1.263: '1p2', 1.325: '1p26'}


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


def CDR_cap_array_test(VCO_trim=0x7FC):
    VCO_trim_low = VCO_trim & 0xFF  # Bits [5:0]
    VCO_trim_mid = (VCO_trim >> 8) & 0xFF  # Bits [13:6]
    VCO_trim_high = (VCO_trim >> 16) & 0x0F  # Bits [17:14]
    # Write commands
    # For CDR0
    aardvark.basic_write(0x0C, [0x00, 0x03])
    aardvark.basic_write(0x48, [0x0B, 0x32])
    aardvark.basic_write(0x4D, [0x20, 0x05])
    aardvark.basic_write(0x52, [0x0B, 0x32])
    aardvark.basic_write(0x57, [0x20, 0x05])
    aardvark.basic_write(0x4D, [0x21])
    aardvark.basic_write(0x57, [0x21])
    aardvark.basic_write(0x20, [0x00, 0x0B])
    aardvark.basic_write(0x23, [0x1A, 0x08, 0x54, 0x06, 0x03, 0x00])

    aardvark.basic_write(0x4D, [0xA1])
    aardvark.basic_write(0x57, [0x21])
    aardvark.basic_write(0x20, [0x01])
    aardvark.basic_write(0x4A, [VCO_trim_low, VCO_trim_mid, VCO_trim_high])
    aardvark.basic_write(0x0C, [0x00, 0x38])
    # # For CDR1
    aardvark.basic_write(0x0C, [0x00, 0x03])
    aardvark.basic_write(0x20, [0x00, 0x0B])
    aardvark.basic_write(0x23, [0x1A, 0x08, 0x54, 0x06, 0x03, 0x00])

    aardvark.basic_write(0x4D, [0x21])
    aardvark.basic_write(0x57, [0xA1])
    aardvark.basic_write(0x20, [0x01])
    aardvark.basic_write(0x54, [VCO_trim_low, VCO_trim_mid, VCO_trim_high])
    aardvark.basic_write(0x0C, [0x00, 0x38])
    aardvark.basic_write(0x43, [0x1F])
    aardvark.basic_write(0x0C, [0x00, 0xB8])


def CDR_chrg_pmp_test(chrg_pmp=0xFF):
    # chrg_pmp default 0xCF
    aardvark.basic_write(0x4F, [chrg_pmp])
    aardvark.basic_write(0x59, [chrg_pmp])


def PowerSupply(v12):
    return dir_pwr.get(v12, "Unknown")


def PowerSupply_RD(v12):
    return dir_pwr_RD.get(v12, "Unknown")


# -------- MAIN --------
main_win = connect_to_ucd()
radios = main_win.descendants(control_type="RadioButton")

# Manually mapped indexes from your radio buttons
# MUST DISPLAY UCD500 TX Link TAB at TOP of screen
lane_radio_buttons = {
    "1": radios[8],
    "2": radios[9],
    "4": radios[10]
}
bitrate_radio_buttons = {
    "1.62": radios[11],
    "2.70": radios[12],
    "5.40": radios[13],
    "6.75": radios[14],
    "8.10": radios[15]
}

ADDR_PWR = "GPIB0::12::INSTR"
pwr = Agilent_E3631A(ADDR_PWR)

aardvark = AardvarkController(i2c_address=0x10, bitrate=400)
aardvark.open()

LOSDET_GAIN_BOOST = list(range(8))
LOS_TH = list(range(8))
LOSDET_DLY_X3 = [1]

V12RT = [1.269, 1.337, 1.4]
V12RD = [1.199, 1.263, 1.325]
# dir_pwr = {1.269:'1p14',1.337:'1p2',1.4:'1p26'}
# dir_pwr_RD = {1.199:'1p14',1.263:'1p2',1.325:'1p26'}
MODE = ['RT', 'RD']
CABLE = '0p2m'  # '0p2m', '1p8m'
# -------- AUTOMATION --------
part = 25
Iteration = 3
results = [["LOSDET_GAIN_BOOST", "LOS_TH", "LOSDET_DLY_X3", "MODE", "LC", "LR", "VDD", "UCD LT"]]

for mode in MODE:
    pwr.set_outputOFF()
    time.sleep(1)
    pwr.set_VoltageP6V(1.2)  # SET 1.2V 1A for INITIALIZATION
    pwr.get_CurrP6V()
    pwr.set_outputON()
    time.sleep(1)
    if mode == 'RT':
        aardvark.execute_batch_file(r"C:\WORK\Swift\A0_RT_0x32_I2C_SEQ_PIN_AUTO.xml")
        CDR_cap_array_test()
        CDR_chrg_pmp_test()
        V12 = V12RT
    elif mode == 'RD':
        aardvark.execute_batch_file("C:\WORK\Swift\A0_RD_I2C_SEQ_PIN_AUTO.xml")
        V12 = V12RD

    for v12 in V12:
        pwr.set_VoltageP6V(v12)
        current = pwr.get_CurrP6V()

        for lane in lane_counts:
            for rate in bitrates:
                print(f"\n⚙️ Setting Lane = {lane}, Bitrate = {rate} Gbps")
                try:
                    lane_radio_buttons[lane].select()
                    bitrate_radio_buttons[rate].select()
                    for v3 in LOSDET_DLY_X3:
                        for v1 in LOSDET_GAIN_BOOST:
                            aardvark.basic_write(0x63, [v1 << 4])
                            LOS_CONTROL0 = aardvark.read_register(0x63, 1)
                            for v2 in LOS_TH:
                                aardvark.basic_write(0x64, [(v3 << 3) + v2])
                                LOS_CONTROL1 = (aardvark.read_register(0x64, 1))[0]
                                bit1to2 = hex(int(LOS_CONTROL1, 16) & 0x07)
                                bit3 = hex((int(LOS_CONTROL1, 16) >> 3) & 0x1)

                                # Click "Link Training"
                                pass_count = 0
                                fail_count = 0
                                fail_messages = []

                                for trial in range(Iteration):
                                    try:
                                        main_win.child_window(title="Link Training",
                                                              control_type="Button").click_input()
                                        print(f"🔄 Link Training attempt {trial + 1}/%s..." % Iteration)

                                        result = verify_lt_by_bitrate_safe(main_win, expected_rate=rate)
                                        print(f"   ▶ {result}")

                                        if "✅" in result:
                                            pass_count += 1
                                        else:
                                            fail_count += 1
                                            fail_messages.append(result)
                                        time.sleep(1)

                                    except Exception as e:
                                        fail_count += 1
                                        msg = f"Error: {e}"
                                        fail_messages.append(msg)
                                        print(f"❌ Error on attempt {trial + 1}: {msg}")

                                summary = f"'{pass_count}/%s" % Iteration
                                if mode == 'RT':
                                    preview = [LOS_CONTROL0[0], bit1to2, bit3, mode, lane, rate, PowerSupply(v12),
                                               summary]
                                elif mode == 'RD':
                                    preview = [LOS_CONTROL0[0], bit1to2, bit3, mode, lane, rate, PowerSupply_RD(v12),
                                               summary]

                                print("📊 Summary:", preview)
                                print("=============================================\n")
                                results.append(preview)

                except Exception as e:
                    print(f"❌ Failed: {e}")
                    if mode == 'RT':
                        results.append(
                            [LOS_CONTROL0[0], bit1to2, bit3, mode, lane, rate, PowerSupply(v12), f"Error: {e}"])
                    elif mode == 'RD':
                        results.append(
                            [LOS_CONTROL0[0], bit1to2, bit3, mode, lane, rate, PowerSupply_RD(v12), f"Error: {e}"])

# Save to EXCEL
pwr.set_outputOFF()
book = xw.Book()
sheet = book.sheets[0]
sheet.range("A1").value = results
today_datetime = datetime.today().strftime('%m%d%Y_%H%M%S')
filename_o = r"C:\WORK\Swift\Result\part%s_2L_HBR2_%s_%s.xlsx" % (part, CABLE, today_datetime)
book.save(filename_o)
book.close()
aardvark.close()
pwr.close()
