# -*- coding: utf-8 -*-
"""
Created on Mon Apr 29 16:11:07 2025

@author: Caesal Cheng
"""

from DPR_100 import DPR_100
from DPT_200 import DPT_200
from temperature_chamber_caesal import TemperatureChamber
from aardvark_controller import AardvarkController
from Keysight_91304A import Keysight_91304A
from Keysight_M8070A import Keysight_M8070A
import time
import xlwings as xw

###############################################
dir_rate = {0x1E: 'HBR3', 0x14: 'HBR2', 0x0A: 'HBR', 0x06: 'RBR'}
dir_lane = {0x4: '4Lane', 0x2: '2Lane', 0x01: '1Lane'}
VoltageSwing = [0, 1, 2, 3]
PreEmphasis = [0, 1, 2, 3]

VoltageSwing_PreEmphasis = [[vs, pe] for vs in VoltageSwing for pe in PreEmphasis if vs + pe <= 3]
print(VoltageSwing_PreEmphasis)

TARGET_TEMPERATURE = list(range(-15, 106, 10))
# print(TARGET_TEMPERATURE)

LEVEL0_VSL = list(range(0, 16, 1))  # 4-bit values (0 to 15)
LEVEL0_PE = list(range(0, 16, 1))  # 4-bit values (0 to 15)
OUTPUT_CTRL1 = []  # 0x44
for pe in LEVEL0_PE:
    for vsl in LEVEL0_VSL:
        combined_value = (pe << 4) | vsl
        OUTPUT_CTRL1.append(hex(combined_value))


# print(OUTPUT_CTRL1)

###############################################
def LinkRate(LR):
    return dir_rate.get(LR, "Unknown")


def LaneCount(LC):
    return dir_lane.get(LC, "Unknown")


def LinkTraining_DPT200toDPR100(LR, LC, VS, PE):
    dpcd103 = (PE << 3) + VS
    source.DPCD_write(0x100, LR)
    source.DPCD_write(0x101, LC + 128)
    source.DPCD_Nwrite(0x102, [0x21, 0x0, 0x0, 0x0, 0x0])  # TPS 1
    sink.DPCD_write(0x202, 0x11)  # CR Done
    sink.SetL0L1VSPE(VS, PE)  # VS PE
    if LC == 4:
        sink.SetL2L3VSPE(VS, PE)  # VS PE
    source.DPCD_Nread(0x202, 5)
    if LC < 4:
        source.DPCD_Nwrite(0x103, [dpcd103, dpcd103])
    else:
        source.DPCD_Nwrite(0x103, [dpcd103, dpcd103, dpcd103, dpcd103])
    source.DPCD_Nwrite(0x102, [0x23, dpcd103, dpcd103, dpcd103, dpcd103, 0x0, 0x0])  # TPS 3
    sink.DPCD_write(0x202, 0x11)
    sink.SetL0L1VSPE(VS, PE)  # VS PE
    if LC == 4:
        sink.SetL2L3VSPE(VS, PE)  # VS PE
    source.DPCD_Nread(0x202, 5)
    if LC < 4:
        source.DPCD_Nwrite(0x103, [dpcd103, dpcd103])
    else:
        source.DPCD_Nwrite(0x103, [dpcd103, dpcd103, dpcd103, dpcd103])
    sink.DPCD_write(0x202, 0x77)
    if LC == 4:
        sink.DPCD_write(0x203, 0x77)
    source.DPCD_Nread(0x202, 2)


def NoCTLE_Rate(LR):
    scope.setDefault()  # Reset Scope configuration
    scope.autoScale()  # Autoscale signals
    scope.add_measVpp(2)  # Vpp including Pre-Emphasis
    scope.add_measVampl(2)  # Vpp excluding Pre-Emphasis
    scope.changeCHdiff(2)  # Enable Differential signal
    scope.dispOffCH(4)  # Disable common mode signal on CH4
    scope.writeVISA(":DISP:MAIN 0,CHAN2")  # Dusable Display for differential signal
    scope.setAutoThr("CHAN2")  # Calc measurement threshold. (Not sure what it is but it's important)

    if LR == 0x14:
        scope.writeVISA(
            ":MEAS:CLOC:METH SOPLL,5.4e9,6.28e6,1")  # CDR Configuration (Reference CDR for HBR2 based on Spec.)
    elif LR == 0xA:
        scope.writeVISA(
            ":MEAS:CLOC:METH SOPLL,2.7e9,6.28e6,1")  # CDR Configuration (Reference CDR for HBR based on Spec.)
    scope.writeVISA(":DISP:MAIN 0,CHAN2")  # Dusable Display for differential signal
    scope.writeVISA(":MTES:FOLD 1,CHAN2")  # Enable Eye Diagram (EQU--> CTLE, CHAN2--> Raw signal)
    scope.stop()

    scope.set_JitterNoiseSrc("CHAN2")  # Select CHAN2 signal for Jitter measurent
    scope.set_JiiterNoiseUnit()  # Change Unit from ps to UI.
    scope.set_JitterNoiseBER()  # Select BER for TJ measurement (BER-9 is default)
    scope.set_JitterNoiseEn(True)  # Enable Jitter Measurement
    scope.clearDis()  # Clear previous capture
    scope.setMemDepth(10e6)  # Num of measurement point to capture
    scope.single()  # Single shot
    scope.add_measEyeHeight()  #
    scope.setIsimStat(2, "PORT41")  # Enable S-parameter cable emulation (Just enable. No cable model in default)


def CTLE_Rate(LR):
    scope.setDefault()  # Reset Scope configuration
    scope.autoScale()  # Autoscale signals
    scope.add_measVpp(2)  # Vpp including Pre-Emphasis
    scope.add_measVampl(2)  # Vpp excluding Pre-Emphasis
    scope.changeCHdiff(2)  # Enable Differential signal
    scope.dispOffCH(4)  # Disable common mode signal on CH4
    scope.writeVISA(":DISP:MAIN 0,CHAN2")  # Disable Display for differential signal
    scope.setAutoThr("CHAN2")  # Calc measurement threshold. (Not sure what it is but it's important)

    scope.setCtleSrc(2)  # Select CH2 as an input for CTLE
    if LR == 0x14:
        scope.setCtleDataRate(5.4e9)  # CTLE Configurations for TP3_RQ(HBR2 Reference Equalizer) based on DP Spec
        scope.setCtleOfPoles(3)  # |
        scope.setCtleDCgain(1)  # |
        scope.setCtlePnFreq(0, 640e6)  # |
        scope.setCtlePnFreq(1, 2.7e9)  # |
        scope.setCtlePnFreq(2, 4.5e9)  # |
        scope.setCtlePnFreq(3, 13.5e9)  # |
    elif LR == 0xA:
        scope.setCtleDataRate(
            2.7e9)  # CTLE Configurations for TP3_RQ(Enhanced HBR Reference Equalizer) based on DP Spec
        scope.setCtleOfPoles(2)  # |
        scope.setCtleDCgain(1)  # |
        scope.setCtlePnFreq(0, 380e6)  # |
        scope.setCtlePnFreq(1, 1.15e9)  # |
        scope.setCtlePnFreq(2, 2e9)  # |

    scope.setCtleEn()  # |-> CTLE Configurations

    scope.setCtleScaling(True)  # Scaling for equalized signal
    vpp_src = scope.getVpp(2)  #
    scope.setCtleScaling(True, vpp_src)  #
    vpp_eq = scope.getVpp("EQU")  #
    scope.setCtleScaling(True, vpp_eq / 8)  # scope.setCtleScaling(True,vpp_eq/6)
    scope.setAutoThr("EQU")  #

    if LR == 0x14:
        scope.writeVISA(
            ":MEAS:CLOC:METH SOPLL,5.4e9,6.28e6,1")  # CDR Configuration (Reference CDR for HBR2 based on Spec.)
    elif LR == 0xA:
        scope.writeVISA(
            ":MEAS:CLOC:METH SOPLL,2.7e9,6.28e6,1")  # CDR Configuration (Reference CDR for HBR based on Spec.)
    scope.writeVISA(":DISP:MAIN 0,EQU")  # Disable Display for differential signal
    scope.writeVISA(":MTES:FOLD 1,EQU")  # Enable Eye Diagram (EQU--> CTLE, CHAN2--> Raw signal)
    scope.stop()

    scope.set_JitterNoiseSrc("EQU")  # Select equalized signal for Jitter measurent
    scope.set_JiiterNoiseUnit()  # Change Unit from ps to UI.
    scope.set_JitterNoiseBER()  # Select BER for TJ measurement (BER-9 is default)
    scope.set_JitterNoiseEn(True)  # Enable Jitter Measurement
    scope.clearDis()  # Clear previous capture
    scope.setMemDepth(10e6)  # Num of measurement point to capture
    scope.single()  # Single shot
    scope.add_measEyeHeight()  #
    scope.setIsimStat(2, "PORT41")  # Enable S-parameter cable emulation (Just enable. No cable model in default)


###############################################
DIR_Sim = r"C:\Users\Public\Documents\Infiniium\Filters"
ADDR_BERT = "TCPIP0::localhost::hislip0::INSTR"
ADDR_SCOPE = "TCPIP0::172.17.11.12::inst0::INSTR"

bert = Keysight_M8070A(ADDR_BERT)
scope = Keysight_91304A(ADDR_SCOPE)
sink = DPR_100(9)
source = DPT_200(4)
temp = TemperatureChamber(5)
aardvark = AardvarkController(i2c_address=0x10, bitrate=400)
aardvark.open()
aardvark.execute_batch_file(r"C:\Caesal\Bootup_Cali\A0_RD_I2C_SEQ_PIN_AUTO.xml")

###############################################
###############################################
############SCOPE Initialization###############
###############################################
###############################################
print("###############Initialization : START##################")
scope.setDefault()  # Reset Scope configuration
scope.autoScale()  # Autoscale signals
scope.add_measVpp(2)  # Vpp including Pre-Emphasis
scope.add_measVampl(2)  # Vpp excluding Pre-Emphasis
scope.changeCHdiff(2)  # Enable Differential signal
scope.dispOffCH(4)  # Disable common mode signal on CH4
scope.writeVISA(":DISP:MAIN 0,CHAN2")  # Dusable Display for differential signal
scope.setAutoThr("CHAN2")  # Calc measurement threshold. (Not sure what it is but it's important)
scope.setCtleSrc(2)  # Select CH2 as an input for CTLE
scope.setCtleDataRate(5.4e9)  # CTLE Configurations for TP3_RQ(HBR2) based on DP Spec
scope.setCtleOfPoles(3)  # |
scope.setCtleDCgain(1)  # |
scope.setCtlePnFreq(0, 640e6)  # |
scope.setCtlePnFreq(1, 2.7e9)  # |
scope.setCtlePnFreq(2, 4.5e9)  # |
scope.setCtlePnFreq(3, 13.5e9)  # |
scope.setCtleEn()  # |-> CTLE Configurations

scope.setCtleScaling(True)  # Scaling for equalized signal
vpp_src = scope.getVpp(2)  #
scope.setCtleScaling(True, vpp_src)  #
vpp_eq = scope.getVpp("EQU")  #
scope.setCtleScaling(True, vpp_eq / 8)  # scope.setCtleScaling(True,vpp_eq/6)
scope.setAutoThr("EQU")  #

scope.writeVISA(":MEAS:CLOC:METH SOPLL,5.4e9,6.28e6,1")  # CDR Configuration (Reference CDR for HBR2 based on Spec.)
scope.writeVISA(":DISP:MAIN 0,EQU")  # Dusable Display for differential signal
scope.writeVISA(":MTES:FOLD 1,EQU")  # Enable Eye Diagram (EQU--> CTLE, CHAN2--> Raw signal)
scope.stop()

scope.set_JitterNoiseSrc("EQU")  # Select equalized signal for Jitter measurent
scope.set_JiiterNoiseUnit()  # Change Unit from ps to UI.
scope.set_JitterNoiseBER()  # Select BER for TJ measurement (BER-9 is default)
scope.set_JitterNoiseEn(True)  # Enable Jitter Measurement
scope.clearDis()  # Clear previous capture
scope.setMemDepth(10e6)  # Num of measurement point to capture
scope.single()  # Single shot
# time.sleep(60)
scope.add_measEyeHeight()  #
scope.setIsimStat(2, "PORT41")  # Enable S-parameter cable emulation (Just enable. No cable model in default)
print("###############Initialization : Done##################")

###############################################
###############################################
###############################################
###############################################
# aardvark.execute_batch_file(r"C:\WORK\Swift\Bootup_Cali\redriver_I2c_program_79_part4.xml")
DIR_SAVE = r"C:\\Share\Swift\TestChip_A1\Caesal"
result_all = [
    ["Part", "LR", "LC", "Temperature", "Cable", "JBERT", "JBERT_ISI", "V12", "SWIFT_RXEQ", "VSL_PE", "tj", "rj", "dj",
     "eye H", "eye W", "Current", "vpp", "vamp"]]
###############################################

TEMPERATURE = [20]  # [-40,-20,0,20,40,60,80,105]
SIM_FILE_LIST = ["DoNothing", "1m_Cable", "4p5m_Cable"]  # ["DoNothing","1m_Cable","4p5m_Cable","CIC_rev0p6"]
JBERT_VS = [800]  # [400,600,800]
V12 = [1.2, 1.26, 1.14]
SWIFT_RXEQ = ["ENABLED"]  # ["ENABLED","DISABLED"]
JBERT_ISI = [0, -3, -6]  # [0,-3,-6]
LinkRa = [0x14, 0xA]  # [0x14,0xA,0x6]
VoltageSwing_PreEmphasis = [[1, 1]]

LC = 0x1
# LC = 0x2
# VS = 1
# PE = 1
v12 = 1.2
VSL_PE = '0x0'
current = 0
part = 1
mode = "ReDriverMode"  # ["ReTimerMode","ReDriverMode"]

###############################################
###############################################
###############################################
###############################################
time_start = time.time()
aardvark.execute_batch_file(r"C:\Caesal\Bootup_Cali\A0_RD_I2C_SEQ_PIN_AUTO.xml")
# aardvark.execute_batch_file(r"C:\Caesal\Bootup_Cali\A0_RT_I2C_SEQ_PIN_AUTO.xml")

LinkTraining_DPT200toDPR100(0x14, 0x2, 0, 2)  # Establish Link to config CTLE on scope
for target_temp in TEMPERATURE:
    for LR in LinkRa:
        for vs_in in JBERT_VS:
            bert.set_DoutAmpl(1, 1,
                              vs_in / 2000)  # JBERT V amplitude (It's single-end configuration, so 200mV means 400mV in differential)
            if LR == 0xA and LC == 0x1:
                bert.set_BitRate(2.7e9)
                bert.set_ISIFreq(1, 1.35e9)
                bert.restart_All()
                bert.break_Sequence(1)
                bert.break_Sequence(1)
                bert.break_Sequence(1)
                CTLE_Rate(LR)  # CTLE Configurations for TP3_RQ(HBR) based on DP Spec
            elif LR == 0x14 and LC == 0x1:
                bert.set_BitRate(5.4e9)
                bert.set_ISIFreq(1, 2.7e9)
                bert.restart_All()
                bert.break_Sequence(1)
                bert.break_Sequence(1)
                CTLE_Rate(LR)  # CTLE Configurations for TP3_RQ(HBR2) based on DP Spec
            if LR == 0xA and LC == 0x2:
                bert.set_BitRate(2.7e9)
                bert.set_ISIFreq(1, 1.35e9)
                bert.set_ISIFreq(2, 1.35e9)
                bert.restart_All()
                bert.break_Sequence(1)
                bert.break_Sequence(1)
                bert.break_Sequence(1)
                bert.break_Sequence(2)
                bert.break_Sequence(2)
                bert.break_Sequence(2)
                CTLE_Rate(LR)  # CTLE Configurations for TP3_RQ(HBR) based on DP Spec
            elif LR == 0x14 and LC == 0x2:
                bert.set_BitRate(5.4e9)
                bert.set_ISIFreq(1, 2.7e9)
                bert.set_ISIFreq(2, 2.7e9)
                bert.restart_All()
                bert.break_Sequence(1)
                bert.break_Sequence(1)
                bert.break_Sequence(2)
                bert.break_Sequence(2)
                CTLE_Rate(LR)  # CTLE Configurations for TP3_RQ(HBR2) based on DP Spec
            for isi in JBERT_ISI:
                isi_i = -0.03 if isi == 0 else isi
                bert.set_Isi(1, 1, isi_i)  # Insertion loss from JBERT
                if LC == 0x2:
                    bert.set_Isi(1, 2, isi_i)  # Insertion loss from JBERT

                for sim_file in SIM_FILE_LIST:

                    scope.recallIsimFile(2, DIR_Sim + "\\" + sim_file + ".tf4")  # Select Cable model
                    scope.clearDis()
                    for EQ in SWIFT_RXEQ:
                        for VSPE in VoltageSwing_PreEmphasis:
                            VS = VSPE[0]
                            PE = VSPE[1]
                            LinkTraining_DPT200toDPR100(LR, 0x2, VS, PE)
                            # for VSL_PE in OUTPUT_CTRL1:
                            test_s = time.time()
                            #    aardvark.write_register(0x44, VSL_PE)

                            scope.clearDis()  # Clear previous capture
                            scope.single()  # Single shot
                            time.sleep(2)

                            eyeW = scope.getEyeWidth()
                            eyeH = scope.getEyeHeight()
                            jitter = scope.get_JitterNoiseTjRjDj()
                            tj = jitter[0][1]
                            rj = jitter[1][1]
                            dj = jitter[2][1]
                            vpp = scope.getVpp("EQU")
                            vamp = scope.getVamp("EQU")

                            count = 0
                            while (eyeH == 0 or tj > 500) and count < 4:
                                LinkTraining_DPT200toDPR100(LR, 0x2, VS, PE)
                                scope.clearDis()  # Clear previous capture
                                scope.single()  # Single shot
                                time.sleep(2)
                                eyeW = scope.getEyeWidth()
                                eyeH = scope.getEyeHeight()
                                jitter = scope.get_JitterNoiseTjRjDj()
                                tj = jitter[0][1]
                                rj = jitter[1][1]
                                dj = jitter[2][1]
                                vpp = scope.getVpp("EQU")
                                vamp = scope.getVamp("EQU")
                                print("Eye diagram is too narrow to measure jitter, retrying again, count = %s", count)
                                count += 1
                                if count == 4:
                                    print(
                                        "Eye diagram is too narrow to measure jitter, leave whatever it is, count = %s",
                                        count)

                            # Save the data into "result_all"
                            result = [part, LR, LC, target_temp, sim_file, vs_in, isi, v12, EQ, VS * 10 + PE, tj, rj,
                                      dj, eyeH, eyeW, current, vpp, vamp]
                            result_all.append(result)

                            # Save eye diagram image
                            if LR == 0xA:
                                img_name = DIR_SAVE + r"\SWIFT_TP3_EQ part%s_RX1_PRBS7_%s_%s_%smV_Cable%s_ISI%sdB_EQ%s_VS%d_PE%d_%dC_%s" % (
                                    part, LinkRate(LR), LaneCount(LC), vs_in, sim_file, isi, EQ, VS, PE, target_temp,
                                    mode)
                            if LR == 0x14:
                                img_name = DIR_SAVE + r"\SWIFT_TP3_EQ_part%s_RX1_CP2520_%s_%s_%smV_Cable%s_ISI%sdB_EQ%s_VS%d_PE%d_%dC_%s" % (
                                    part, LinkRate(LR), LaneCount(LC), vs_in, sim_file, isi, EQ, VS, PE, target_temp,
                                    mode)

                            #    img_name = DIR_SAVE + r"\OUTPUT_CTRL1=%s_%s_TP2_PRBS7_%s_800mV_ISI=0dB_%dmV_%ddB_%dC"%(VSL_PE,LinkRate(LR),LaneCount(LC),VS,PE,target_temp)
                            #    print (img_name,time.time()-test_s)
                            #    print("OUTPUT_CTRL1=%s %s TP2 PRBS7 %s 800mV ISI=0dB %dmV %ddB %d°C"%(VSL_PE,LinkRate(LR),LaneCount(LC),VS,PE,target_temp))
                            print(
                                "LR = %s VSPE = %s Cable = %s ISI = %s Vpp = %.3f Vamp = %.3f EyeH = %.3f EyeW = %e Tj = %.3f Rj = %.3f Dj = %.3f" % (
                                    LR, VSPE, sim_file, isi, vpp, vamp, eyeH, eyeW, tj, rj, dj))

                            scope.saveImg(img_name)

                        ###############################################
###############################################
###############################################
time_end = time.time()
book = xw.Book()
sheet = book.sheets[0]
sheet.range("A1").value = result_all
filename_o = r"C:\Caesal\test\SWIFT_part#1#3#16#20#29#30#32_RX1_L0_TP3EQ_RD.xlsx"
# filename_o = r"C:\Caesal\test\SWIFT_HBR_Part5-14_RX1_L0_EQOFF.xlsx"
# filename_o = r"\\%s\%s\EyeDiagram_OUTPUT_CTRL1_TP2_TP2_PRBS7_800mV_ISI=0dB_#04_FIB.xlsx"%(ADDR_SCOPE.split("::")[1],DIR_SAVE.split(r"\\")[1])
book.save(filename_o)
book.close()
print(time_end - time_start)

aardvark.close()
sink.close()
source.close()
temp.Set_Mon_temperature(20)
time.sleep(60)
temp.close()
###############################################


# Old Documents(Dont delete)
# sink.Set_DPVersion = (1.1)
# sink.MaxCapability(0x14,0x1,0)
# for LR in LinkRate:
#     for LC in LaneCount:
#         for VSPE in VoltageSwing_PreEmphasis:
#             VS = VSPE[0]
#             PE = VSPE[1]
#             print("Link Rate: ",LR, "Lane Count: ",LC, "VS: ",VS, "PE: ",PE)
#             LinkTraining_DPT200toDPR100(LR,LC,VS,PE)
#             time.sleep(2)

################################################

# AUTOMATED_TEST_REQUEST
# DPCD_TEST_REQ
# sink.DPCD_RWR(0x218,0x8) #Request Pattern Change
# DPCD_MAX_DOWNSPREAD
# s.DPCD_RWR(0x03,0x0)
###############################################
# sink.Set_AutoTest(LR,LC,PT) # LinkRate LaneCount Pattern
###############################################
##DPCD_L1L0_ADJ_REQ
##Fake Link
# print (sink.MaxCapability(LR,LC,0))
##sink.FakeLink()
# source.DPCD_read(0x202)
# sink.DPCD_RWR(0x201,0x2)
##sink.HPD_PULSE_L()
# time.sleep(5)
# source.DPCD_Nwrite(0x102,0x21)
# sink.DPCD_RWR(0x202,0x11)
# source.DPCD_read(0x202)
# sink.DPCD_RWR(0x202,0x77)
# sink.DPCD_RWR(0x203,0x00)
##sink.HPD_PULSE_H()
##s.DPCD_RWR(0x218,1)
################################################
# sink.SetL0L1VSPE(VS,PE) # VS PE
# sink.SetL2L3VSPE(VS,PE) # VS PE
##time.sleep(1)
# print (sink.HPD_ShortPULSE(1))
##time.sleep(1)
# source.DPCD_Nread(0x200,7)

