# -*- coding: utf-8 -*-
"""
Created on Mon Nov 21 16:11:07 2024

@author: Caesal Cheng
"""

from DPR_100 import DPR_100
from DPT_200 import DPT_200
# from temperature_chamber_caesal import TemperatureChamber
from aardvark_controller import AardvarkController
from Keysight_91304A import Keysight_91304A
from Keysight_M8070A import Keysight_M8070A
from Agilent_E3631A import Agilent_E3631A
from Agilent_34401A import Agilent_34401A
import time
from datetime import datetime
import xlwings as xw

###############################################
dir_rate = {0x1E: 'HBR3', 0x14: 'HBR2', 0x0A: 'HBR', 0x06: 'RBR'}
dir_lane = {0x4: '4Lane', 0x2: '2Lane', 0x01: '1Lane'}
# dir_lane = {0x2:'Lane1',0x01:'Lane0'}
V12 = [1.168, 1.228, 1.108]  # [1.2,1.26,1.14] RT -32mV
V12_RD = [1.186, 1.246, 1.126]  # [1.2,1.26,1.14] RD -14mV

dir_pwr = {1.168: '1p2', 1.228: '1p26', 1.108: '1p14', 1.186: '1p2', 1.246: '1p26', 1.126: '1p14'}
VoltageSwing = [0, 1, 2, 3]
PreEmphasis = [0, 1, 2, 3]

VoltageSwing_PreEmphasis = [[vs, pe] for vs in VoltageSwing for pe in PreEmphasis if vs + pe <= 3]
# print(VoltageSwing_PreEmphasis)

TARGET_TEMPERATURE = list(range(-15, 106, 10))


# print(TARGET_TEMPERATURE)

# LEVEL0_VSL = list(range(0, 16, 1))  # 4-bit values (0 to 15)
# LEVEL0_PE  = list(range(0, 16, 1))  # 4-bit values (0 to 15)
# OUTPUT_CTRL1 = [] # 0x44
# for pe in LEVEL0_PE:
#    for vsl in LEVEL0_VSL:
#        combined_value = (pe << 4) | vsl
#        OUTPUT_CTRL1.append(hex(combined_value))
# print(OUTPUT_CTRL1)

###############################################
###############################################
def LinkRate(LR):
    return dir_rate.get(LR, "Unknown")


def LaneCount(LC):
    return dir_lane.get(LC, "Unknown")


def PowerSupply(v12):
    return dir_pwr.get(v12, "Unknown")


def float_to_p_string(value):
    return str(value).replace('.', 'p')


def ModeSelect(CHIP, mode):
    if CHIP == 'PARADE':
        if mode == "ReTimerMode_noSSC":
            pwr.set_outputOFF()
            pwr.set_VoltageP6V(0)  # Force CONFIG1 GPIO to HIGH
            time.sleep(1)
            pwr.set_outputON()  # Power Cycle
            time.sleep(1)
            pwr.get_CurrP25V()
            bert.set_SSC('OFF')
        elif mode == "ReTimerMode_SSC":
            pwr.set_outputOFF()
            pwr.set_VoltageP6V(0)  # Force CONFIG1 GPIO to HIGH
            time.sleep(1)
            pwr.set_outputON()  # Power Cycle
            time.sleep(1)
            pwr.get_CurrP25V()
            bert.set_SSC()
        elif mode == "ReDriverMode":
            pwr.set_outputOFF()
            pwr.set_VoltageP6V(1.6)  # Force CONFIG1 GPIO to MID
            time.sleep(1)
            pwr.set_outputON()  # Power Cycle
            time.sleep(1)
            pwr.get_CurrP25V()
            bert.set_SSC()
        else:
            print("#################  WRONG MODE  #################")
    elif CHIP == 'SWIFT':
        bert.set_SSC()
        if mode == "ReDriverMode":
            pwr.set_outputOFF()
            pwr.set_VoltageP6V(1.6)  # Force CONFIG1/4 GPIO to MID
            time.sleep(1)
            pwr.set_outputON()  # Power Cycle
            time.sleep(1)
            pwr.get_CurrP25V()
            bert.set_SSC()
            aardvark.execute_batch_file(r"C:\WORK\Swift\Bootup_Cali\A0_RD_I2C_SEQ_PIN_AUTO.xml")
            time.sleep(1)
        elif mode == "ReTimerMode":
            pwr.set_outputOFF()
            pwr.set_VoltageP6V(0)  # Force CONFIG1/4 GPIO to HIGH
            time.sleep(1)
            pwr.set_outputON()
            time.sleep(1)
            pwr.get_CurrP25V()
            bert.set_SSC()
            aardvark.execute_batch_file(r"C:\WORK\Swift\Bootup_Cali\A0_RT_0x32_I2C_SEQ_PIN_AUTO.xml")
            #            aardvark.execute_batch_file(r"C:\Caesal\Bootup_Cali\A0_RT_0x32_I2C_SEQ_PIN_AUTO_SWIF225.xml")
            time.sleep(1)
        else:
            print("#################  WRONG MODE  #################")
    else:
        print("#################  WRONG CHIP  #################")
    print("#################  %s %s  #################" % (CHIP, mode))


def LinkTraining_DPT200toDPR100(LR, LC, VS, PE):
    dpcd103 = (PE << 3) + VS
    sink.DPCD_write(0x3, 0x1)
    source.DPCD_Nread(0x0, 16)
    source.DPCD_write(0x100, LR)
    source.DPCD_write(0x101, LC + 128)
    if LR == 0x14:
        source.DPCD_Nwrite(0x10B, [0x7, 0x7])  # SSC
        source.DPCD_write(0x107, 0x10)  # SSC
    else:
        source.DPCD_Nwrite(0x10B, [0x3, 0x3])  # SSC
        source.DPCD_write(0x107, 0x10)  # SSC

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
    source.DPCD_Nread(0x202, 6)
    if LC < 4:
        source.DPCD_Nwrite(0x103, [dpcd103, dpcd103])
    else:
        source.DPCD_Nwrite(0x103, [dpcd103, dpcd103, dpcd103, dpcd103])
    sink.DPCD_write(0x202, int(hex([0, 0x7, 0x77][LC]), 16))
    if LC == 4:
        sink.DPCD_write(0x203, 0x77)
    source.DPCD_Nwrite(0x102, [0x00])
    source.DPCD_Nread(0x202, LC)


def JBERT_Rate_Pattern(LR, PT, LC):
    if LR == 0x14:
        bert.set_BitRate(5.4e9)
        bert.set_ISIFreq(1, 2.7e9)
        bert.set_ISIFreq(2, 2.7e9)
        bert.restart_All()
        bert.break_Sequence(1)
        bert.break_Sequence(1)
        bert.break_Sequence(2)
        bert.break_Sequence(2)  # CP2520 Pattern
    elif LR == 0xA:
        bert.set_BitRate(2.7e9)
        bert.set_ISIFreq(1, 1.35e9)
        bert.set_ISIFreq(2, 1.35e9)
    if PT == 'PLTPAT':
        bert.restart_All()
        for _ in range(4):
            bert.break_Sequence(1)
            bert.break_Sequence(2)  # PLTPAT
    elif PT == 'PRBS7':
        bert.restart_All()
        for _ in range(3):
            bert.break_Sequence(1)
            bert.break_Sequence(2)  # PRBS7 Pattern
    elif PT == 'CP2520':
        bert.restart_All()
        for _ in range(1):
            bert.break_Sequence(1)
            bert.break_Sequence(2)  # CP2520 Pattern
    elif PT == 'D10d2':
        bert.restart_All()
        for _ in range(5):
            bert.break_Sequence(1)
            bert.break_Sequence(2)  # D10d2 Pattern
    if LC == 0x1:
        bert.setLocalOutput(1, True)
        bert.setLocalOutput(2, False)
    elif LC == 0x2:
        bert.setLocalOutput(1, True)
        bert.setLocalOutput(2, True)


def NoCTLE_Rate(LR, stage='none'):
    if stage == 'none':
        scope.setDefault()  # Reset Scope configuration
        scope.autoScale()  # Autoscale signals
        scope.add_measVpp(2)  # Vpp including Pre-Emphasis
        scope.add_measVampl(2)  # Vpp excluding Pre-Emphasis
        scope.changeCHdiff(2)  # Enable Differential signal
        scope.dispOffCH(1)  # Disable common mode signal on CH4
        scope.dispOffCH(3)  # Disable common mode signal on CH1
        scope.dispOffCH(4)  # Disable common mode signal on CH3
        scope.writeVISA(":DISP:MAIN 0,CHAN2")  # Disable Display for differential signal
        scope.writeVISA(":DISP:MAIN 0,CHAN1")  # Disable Display for differential signal
        scope.writeVISA(":DISP:MAIN 0,CHAN3")  # Disable Display for differential signal
        scope.writeVISA(":DISP:MAIN 0,CHAN4")  # Disable Display for differential signal
        scope.setAutoThr("CHAN2")  # Calc measurement threshold.

        if LR == 0x14:
            scope.writeVISA(
                ":MEAS:CLOC:METH SOPLL,5.4e9,6.28e6,1")  # CDR Configuration (Reference CDR for HBR2 based on Spec.)
        elif LR == 0xA:
            scope.writeVISA(
                ":MEAS:CLOC:METH SOPLL,2.7e9,6.28e6,1")  # CDR Configuration (Reference CDR for HBR based on Spec.)
        scope.writeVISA(":DISP:MAIN 0,CHAN2")  # Disable Display for differential signal
        scope.writeVISA(":DISP:MAIN 0,CHAN1")  # Disable Display for differential signal
        scope.writeVISA(":DISP:MAIN 0,CHAN3")  # Disable Display for differential signal
        scope.writeVISA(":DISP:MAIN 0,CHAN4")  # Disable Display for differential signal
        scope.writeVISA(":MTES:FOLD 1,CHAN2")  # Enable Eye Diagram (EQU--> CTLE, CHAN2--> Raw signal)
        scope.stop()

        scope.set_JitterNoiseSrc("CHAN2")  # Select CHAN2 signal for Jitter measurent
        scope.set_JiiterNoiseUnit()  # Change Unit from ps to UI.
        scope.set_JitterNoiseBER()  # Select BER for TJ measurement (BER-9 is default)
        scope.set_JitterNoiseEn(True)  # Enable Jitter Measurement
        scope.clearDis()  # Clear previous capture
        scope.setMemDepth(10e6)  # Num of measurement point to capture
        #    scope.single()                      # Single shot
        scope.add_measEyeHeight()  #
        scope.setIsimStat(2, "PORT41")  # Enable S-parameter cable emulation (Just enable. No cable model in default)
    elif stage == 'TP0':
        scope.setDefault()  # Reset Scope configuration
        scope.autoScale()  # Autoscale signals
        scope.add_measVpp(1)  # Vpp including Pre-Emphasis
        scope.add_measVampl(1)  # Vpp excluding Pre-Emphasis
        scope.changeCHdiff(1)  # Enable Differential signal
        scope.dispOffCH(2)  # Disable common mode signal on CH4
        scope.dispOffCH(3)  # Disable common mode signal on CH1
        scope.dispOffCH(4)  # Disable common mode signal on CH3
        scope.writeVISA(":DISP:MAIN 0,CHAN2")  # Disable Display for differential signal
        scope.writeVISA(":DISP:MAIN 0,CHAN1")  # Disable Display for differential signal
        scope.writeVISA(":DISP:MAIN 0,CHAN3")  # Disable Display for differential signal
        scope.writeVISA(":DISP:MAIN 0,CHAN4")  # Disable Display for differential signal
        scope.setAutoThr("CHAN1")  # Calc measurement threshold.

        if LR == 0x14:
            scope.writeVISA(
                ":MEAS:CLOC:METH SOPLL,5.4e9,6.28e6,1")  # CDR Configuration (Reference CDR for HBR2 based on Spec.)
        elif LR == 0xA:
            scope.writeVISA(
                ":MEAS:CLOC:METH SOPLL,2.7e9,6.28e6,1")  # CDR Configuration (Reference CDR for HBR based on Spec.)
        scope.writeVISA(":DISP:MAIN 0,CHAN2")  # Disable Display for differential signal
        scope.writeVISA(":DISP:MAIN 0,CHAN1")  # Disable Display for differential signal
        scope.writeVISA(":DISP:MAIN 0,CHAN3")  # Disable Display for differential signal
        scope.writeVISA(":DISP:MAIN 0,CHAN4")  # Disable Display for differential signal
        scope.writeVISA(":MTES:FOLD 1,CHAN1")  # Enable Eye Diagram (EQU--> CTLE, CHAN2--> Raw signal)
        scope.stop()

        scope.set_JitterNoiseSrc("CHAN1")  # Select CHAN2 signal for Jitter measurent
        scope.set_JiiterNoiseUnit()  # Change Unit from ps to UI.
        scope.set_JitterNoiseBER()  # Select BER for TJ measurement (BER-9 is default)
        scope.set_JitterNoiseEn(True)  # Enable Jitter Measurement
        scope.clearDis()  # Clear previous capture
        scope.setMemDepth(10e6)  # Num of measurement point to capture
        #    scope.single()                      # Single shot
        scope.add_measEyeHeight()  #
        scope.setIsimStat(1, "PORT41")  # Enable S-parameter cable emulation (Just enable. No cable model in default)


def CTLE_Rate(LR):
    scope.setDefault()  # Reset Scope configuration
    scope.autoScale()  # Autoscale signals
    scope.add_measVpp(2)  # Vpp including Pre-Emphasis
    scope.add_measVampl(2)  # Vpp excluding Pre-Emphasis
    scope.changeCHdiff(2)  # Enable Differential signal
    scope.dispOffCH(4)  # Disable common mode signal on CH4
    scope.writeVISA(":DISP:MAIN 0,CHAN2")  # Disable Display for differential signal
    scope.writeVISA(":DISP:MAIN 0,CHAN1")  # Disable Display for differential signal
    scope.writeVISA(":DISP:MAIN 0,CHAN3")  # Disable Display for differential signal
    scope.writeVISA(":DISP:MAIN 0,CHAN4")  # Disable Display for differential signal
    scope.setAutoThr("CHAN2")  # Calc measurement threshold.

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
    #    scope.single()                      # Single shot
    scope.add_measEyeHeight()  #
    scope.setIsimStat(2, "PORT41")  # Enable S-parameter cable emulation (Just enable. No cable model in default)


def AUTOSCALE_CTLE():
    scope.setCtleScaling(True)  # Scaling for equalized signal
    vpp_src = scope.getVpp(2)  #
    scope.setCtleScaling(True, vpp_src)  #
    vpp_eq = scope.getVpp("EQU")  #
    scope.setCtleScaling(True, vpp_eq / 8)  # scope.setCtleScaling(True,vpp_eq/6)
    scope.setAutoThr("EQU")  #


def CDR_cap_array_test(VCO_trim=0x7FC):
    VCO_trim_low = VCO_trim & 0xFF  # Bits [5:0]
    VCO_trim_mid = (VCO_trim >> 8) & 0xFF  # Bits [13:6]
    VCO_trim_high = (VCO_trim >> 16) & 0x0F  # Bits [17:14]
    aardvark.basic_write(0x4A, [VCO_trim_low, VCO_trim_mid, VCO_trim_high])
    aardvark.basic_write(0x54, [VCO_trim_low, VCO_trim_mid, VCO_trim_high])
    aardvark.basic_write(0x0C, [0x00, 0xB8])


def CDR_chrg_pmp_test(chrg_pmp=0xFF):
    # chrg_pmp default 0xCF, need to try 0xAF, 0xBF, 0xCF, 0xDF, 0xEF, 0xFF
    aardvark.basic_write(0x4F, [chrg_pmp])
    aardvark.basic_write(0x59, [chrg_pmp])


def CDR_rz_sel(rz_sel):
    # rz_sel default 5k, 0x24
    aardvark.basic_write(0x4D, [rz_sel])
    aardvark.basic_write(0x57, [rz_sel])


def CDR_duty_cycle_correction(switch):
    if switch == 1:
        aardvark.basic_write(0x48, [0x2A])  # write value only
        aardvark.basic_write(0x52, [0x2A])  # write value only
        result = "Yuichi Programming"
        print(result)
        return result
    else:
        aardvark.basic_write(0x48, [0x0A])  # write value only
        aardvark.basic_write(0x52, [0x0A])  # write value only
        result = "Original Programming"
        print(result)
        return result


def Serial_Num():
    aardvark.basic_write(0xB0, [0x01])
    aardvark.basic_write(0xB6, [0x00])
    aardvark.basic_write(0xB1, [0x01])
    time.sleep(0.05)
    serial_Num = aardvark.read_register(0xB9, 1)
    return serial_Num


###############################################
DIR_SAVE = r"C:\\Share\Swift\TestChip_A1\Caesal"
DIR_Sim = r"C:\Users\Public\Documents\Infiniium\Filters"
ADDR_BERT = "TCPIP0::localhost::hislip0::INSTR"
ADDR_SCOPE = "TCPIP0::172.17.11.12::inst0::INSTR"
ADDR_PWR = "GPIB1::5::INSTR"
ADDR_MITT = "GPIB1::22::INSTR"
#
sink = DPR_100(9)
source = DPT_200(4)
bert = Keysight_M8070A(ADDR_BERT)
scope = Keysight_91304A(ADDR_SCOPE)
pwr = Agilent_E3631A(ADDR_PWR)
mitt = Agilent_34401A(ADDR_MITT)
# temp = TemperatureChamber(5)
aardvark = AardvarkController(i2c_address=0x10, bitrate=400)
aardvark.open()

time.sleep(5)
###############################################
###############################################
###############    VARIABLES   ################
###############################################
###############################################
CHIP = 'SWIFT'  # 'PARADE' 'SWIFT'
TEMPERATURE = [20]  # [-40,-20,0,20,40,60,80,105]
SIM_FILE_LIST_TP3 = ["0p5m_2_S", "1p3m_1_S", "2m_5_S",
                     "3m_3_S"]  # ["DoNothing","1m_Cable","4p5m_Cable","CIC_rev0p6"] ["0p5m_2_S","1p3m_1_S","2m_5_S","3m_3_S","CIC_rev0p6"]
JBERT_VS = [800]
LinkRa = [0x14]  # [0x14,0xA,0x6]
LCC = [0x2]  # lane0 : 0x1  lane1: 0x2
V12 = [1.2, 1.26, 1.14]  # [1.2,1.26,1.14]
JBERT_ISI = [0, -1, -2, -3, -6, -9]  # [0,-1,-2,-3,-6,-9,-16]
PATTERN = ['CP2520']  # ['D10d2','CP2520']
VoltageSwing_PreEmphasis = [[0, 0], [0, 1], [0, 2], [1, 1], [1, 2], [2, 1]]
# VoltageSwing_PreEmphasis = [[vs, pe] for vs in VoltageSwing for pe in PreEmphasis if vs + pe <= 3]
STAGE = ["TP2", "TP3", "TP3_EQ"]  # ["TP0","TP2","TP3","TP3_EQ"]
SWIFT_RXEQ = ["ENABLED"]
################# MODE #######################
if CHIP == 'SWIFT':
    MODE = ["ReTimerMode", "ReDriverMode"]  # ["ReTimerMode","ReDriverMode"] SWIFT
elif CHIP == 'PARADE':
    MODE = ["ReDriverMode"]  # ["ReTimerMode_noSSC","ReTimerMode_SSC","ReDriverMode"] # PARADE

################## TP0 #######################
VoltageSwing_PreEmphasis_TP0 = [[0, 0]]
V12_TP0 = [1.2]
################## HBR #######################
JBERT_ISI_HBR = [0, -3, -6, -9]
PATTERN_HBR = ['PRBS7']
################# Programming #######################
DUTYCYC = [0]  # CDR_duty_cycle_correction 1 = ENABLED 0 = DISABLED

# Manual Setup required
RX_PORT = 1  # RX1 : 1 RX2: 2
# LC = 0x1 # lane0 : 0x1  lane1: 0x2
current = 0

###############################################
###################AUTOMATION##################
###############################################
result_all = [
    ["Part", "LR", "LC", "Temperature", "Cable", "JBERT", "JBERT_ISI", "V12", "SWIFT_RXEQ", "VSL_PE", "tj", "rj", "dj",
     "eye H", "eye W", "eye W UI", "Current", "vpp", "vamp", "Mode", "Stage", "Special Programming"]]
METER_CORRECTION = True  # USING Multimeter to ensure the Power Supply is correct on the pin
retry = 0  # Capture x more times when EYE and Jitter is invalid
part = 8  # 'p01' for PARADE
print("#################  NOW I'M TESTING %s PART # %s RX %s #################" % (CHIP, part, RX_PORT))
###############################################
time_start = time.time()

pwr.set_outputOFF()
pwr.set_VoltageP25V(1.2)  # SET 1.2V 1A for INITIALIZATION
pwr.get_CurrP25V()
pwr.set_outputON()
time.sleep(1)
serial = Serial_Num()[0]
print("################# Serial Number = %s #################" % (serial))

for target_temp in TEMPERATURE:  # SET TEMPERATURE FROM CHAMBER
    #    temp.Set_Mon_temperature(target_temp)
    #    time.sleep(2)

    for LR in LinkRa:
        if LR == 0xA:
            JBERT_ISI = JBERT_ISI_HBR
            PT = PATTERN_HBR
        elif LR == 0x14:
            JBERT_ISI = JBERT_ISI
            PT = PATTERN
        for mode in MODE:
            ModeSelect(CHIP, mode)
            CDR_cap_array_test()
            CDR_chrg_pmp_test()
            #            aardvark.basic_write(0x44, [0x11,0x11,0x11,0x11]) # write value only
            for duty_cycle in DUTYCYC:
                Yuichi = CDR_duty_cycle_correction(duty_cycle)
                for vs_in in JBERT_VS:
                    bert.set_DoutAmpl(1, 1,
                                      vs_in / 2000)  # JBERT V amplitude (It's single-end configuration, so 200mV means 400mV in differential)
                    for LC in LCC:
                        for pt in PT:
                            bert.OutputOff(1)
                            time.sleep(1)
                            bert.OutputOn(1)
                            JBERT_Rate_Pattern(LR, pt, LC)  # JBERT Set LinkRate on L0L1 and Pattern
                            JBERT_Rate_Pattern(LR, pt, LC)  # ALWAYS config L0&L1 together when we do 2L LT
                            time.sleep(2)
                            LinkTraining_DPT200toDPR100(LR, LC, 3, 0)  # Establish LT to config CTLE on scope
                            for stage in STAGE:
                                if stage == 'TP0':
                                    SIM_FILE_LIST = ["DoNothing"]  # TP0 has no TX cable applied
                                    print("#################  TP0  #################")
                                    VoltageSwing_PreEmphasis = VoltageSwing_PreEmphasis_TP0
                                    V12 = V12_TP0
                                    NoCTLE_Rate(LR, stage)  # No EQ on Oscilloscope
                                if stage == 'TP2':
                                    SIM_FILE_LIST = ["DoNothing"]  # TP2 has no TX cable applied
                                    print("#################  TP2  #################")
                                    VoltageSwing_PreEmphasis = VoltageSwing_PreEmphasis
                                    V12 = V12
                                    NoCTLE_Rate(LR)  # No EQ on Oscilloscope
                                elif stage == 'TP3':
                                    SIM_FILE_LIST = SIM_FILE_LIST_TP3
                                    print("#################  TP3  #################")
                                    VoltageSwing_PreEmphasis = VoltageSwing_PreEmphasis
                                    V12 = V12
                                    NoCTLE_Rate(LR)  # No EQ on Oscilloscope
                                elif stage == 'TP3_EQ':
                                    SIM_FILE_LIST = SIM_FILE_LIST_TP3
                                    print("#################  TP3 EQ  #################")
                                    VoltageSwing_PreEmphasis = VoltageSwing_PreEmphasis
                                    V12 = V12
                                    CTLE_Rate(LR)  # CTLE Configurations for TP3_EQ(HBR2/HBR) based on DP Spec
                                for sim_file in SIM_FILE_LIST:  # Select Cable model
                                    if stage == 'TP0':
                                        scope.recallIsimFile(1, DIR_Sim + "\\" + sim_file + ".tf4")
                                    else:
                                        scope.recallIsimFile(2, DIR_Sim + "\\" + sim_file + ".tf4")

                                    scope.clearDis()
                                    for isi in JBERT_ISI:
                                        if isi == 0:
                                            isi_i = -0.03
                                        elif (isi == -9 and LR == 0xA):
                                            isi_i = -8.1  # HBR(freq = 1.35GHz) can only reach -8.1dB ISL
                                        elif (isi == -16 and LR == 0x14):
                                            isi_i = -16.2
                                        else:
                                            isi_i = isi
                                        bert.set_Isi(1, 1, isi_i)  # Insertion loss from JBERT
                                        bert.set_Isi(1, 2, isi_i)  # Insertion loss from JBERT
                                        for EQ in SWIFT_RXEQ:
                                            for VSPE in VoltageSwing_PreEmphasis:
                                                VS = VSPE[0]
                                                PE = VSPE[1]
                                                LinkTraining_DPT200toDPR100(LR, LC, VS,
                                                                            PE)  # Always 2L LT when testing L0 or L1
                                                for v12 in V12:
                                                    pwr.set_VoltageP25V(v12)
                                                    pwr.get_CurrP25V()
                                                    if METER_CORRECTION == True:
                                                        #                                                print("Correcting VDD ...")
                                                        step = 0.003  # Voltage step adjustment
                                                        tolerance = 0.0016  # Acceptable difference
                                                        real_vol = mitt.meas_vol()
                                                        v12_adj = v12
                                                        while abs(real_vol - v12) > tolerance:
                                                            if real_vol < v12:
                                                                v12_adj += step
                                                            else:
                                                                v12_adj -= step
                                                            pwr.set_VoltageP25V(v12_adj)

                                                            pwr.get_CurrP25V()
                                                            real_vol = mitt.meas_vol()
                                                        print(
                                                            "Corrected VDD to V%.3f supply to acheive V%s by observing as V%.3f (%.1fmV diff)" % (
                                                                v12_adj, v12, real_vol, (v12_adj - v12) * 1000))

                                                    test_s = time.time()

                                                    scope.clearDis()  # Clear previous capture
                                                    scope.single()  # Single shot
                                                    time.sleep(2)
                                                    eyeW = scope.getEyeWidth()
                                                    eyeH = scope.getEyeHeight()
                                                    jitter = scope.get_JitterNoiseTjRjDj()
                                                    tj = jitter[0][1]
                                                    rj = jitter[1][1]
                                                    dj = jitter[2][1]
                                                    if stage == 'TP3_EQ':
                                                        vpp = scope.getVpp("EQU")
                                                        vamp = scope.getVamp("EQU")
                                                    elif stage == 'TP0':
                                                        vpp = scope.getVpp(1)
                                                        vamp = scope.getVamp(1)
                                                    else:
                                                        vpp = scope.getVpp(2)
                                                        vamp = scope.getVamp(2)

                                                    # Capture x more times when EYE and Jitter is invalid
                                                    count = 0
                                                    while (eyeH == 0 or tj > 500) and count < retry + 1:
                                                        if count == retry:
                                                            print(
                                                                "Eye diagram is still too narrow to measure jitter, leave whatever it is")
                                                            break
                                                        print(
                                                            "Eye diagram is too narrow to measure jitter, retrying again, count = %s" % (
                                                                        count + 1))
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
                                                        if stage == 'TP3_EQ':
                                                            vpp = scope.getVpp("EQU")
                                                            vamp = scope.getVamp("EQU")
                                                        elif stage == 'TP0':
                                                            vpp = scope.getVpp(1)
                                                            vamp = scope.getVamp(1)
                                                        else:
                                                            vpp = scope.getVpp(2)
                                                            vamp = scope.getVamp(2)
                                                        count += 1
                                                    if LR == 0xA:
                                                        eyeW_UI = eyeW / (1 / 2.7e9)
                                                    if LR == 0x14:
                                                        eyeW_UI = eyeW / (1 / 5.4e9)
                                                    current = pwr.get_CurrP25V()
                                                    # Save the data into "result_all"
                                                    result = [part, LinkRate(LR), LaneCount(LC), target_temp, sim_file,
                                                              vs_in, isi, v12, EQ, VS * 10 + PE, tj, rj, dj, eyeH, eyeW,
                                                              eyeW_UI, current, vpp, vamp, mode, stage, Yuichi]
                                                    result_all.append(result)

                                                    today_datetime = datetime.today().strftime('%m%d%Y_%H%M%S')

                                                    # Save eye diagram image
                                                    if LR == 0xA:
                                                        img_name = DIR_SAVE + r"\part%s\%s_%s_part%s_V%s_RX%s_PRBS7_%s_%s_%smV_Cable%s_ISI%sdB_EQ%s_VS%d_PE%d_%dC_%s_%s_%s" % (
                                                            part, CHIP, stage, part, float_to_p_string(v12), RX_PORT,
                                                            LinkRate(LR), LaneCount(LC), vs_in, sim_file, isi, EQ, VS,
                                                            PE, target_temp, mode, Yuichi, today_datetime)
                                                    if LR == 0x14:
                                                        img_name = DIR_SAVE + r"\part%s\%s_%s_part%s_V%s_RX%s_CP2520_%s_%s_%smV_Cable%s_ISI%sdB_EQ%s_VS%d_PE%d_%dC_%s_%s_%s" % (
                                                            part, CHIP, stage, part, float_to_p_string(v12), RX_PORT,
                                                            LinkRate(LR), LaneCount(LC), vs_in, sim_file, isi, EQ, VS,
                                                            PE, target_temp, mode, Yuichi, today_datetime)

                                                    print(
                                                        "TEMP = %s LR = %s VSPE = %s Cable = %s ISI = %s Vpp = %.3f Vamp = %.3f EyeH = %.3f EyeW_UI = %.3f Tj = %.3f Current = %.3f" % (
                                                            target_temp, LinkRate(LR), VSPE, sim_file, isi, vpp, vamp,
                                                            eyeH, eyeW_UI, tj, current))

                                                    scope.saveImg(img_name)

pwr.set_outputOFF()
# temp.Set_Mon_temperature(20)
# time.sleep(40)
# temp.Set_Mon_temperature(105)
time.sleep(5)

###############################################
##############EXCEL File Save##################
###############################################
time_end = time.time()
pwr.set_outputOFF()
book = xw.Book()
sheet = book.sheets[0]
sheet.range("A1").value = result_all
today_datetime = datetime.today().strftime('%m%d%Y_%H%M%S')
filename_o = r"C:\Caesal\test\%s_part%s_serial%s_RX%s_%s_TP2TP3TP3EQ_HBR2_RDRT_VT_%s.xlsx" % (CHIP, part, serial,
                                                                                              RX_PORT, LaneCount(LC),
                                                                                              today_datetime)
book.save(filename_o)
book.close()
print(time_end - time_start)

###############################################
#################Device Close##################
###############################################

aardvark.close()
sink.close()
source.close()
# temp.Set_Mon_temperature(20)
# time.sleep(60)
# temp.close()
pwr.close()
mitt.close()
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

