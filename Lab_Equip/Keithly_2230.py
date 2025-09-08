"""
#---------------------------------------------------------------------------------------------------
# Project   : Aston RTSSM validation framework
# Version   : v1.30($Rev: 6768 $), $Date: 2016-11-14 19:26:11 -0800 (Mon, 14 Nov 2016) $
#       SVN : $Id: Keithly_2280.py 6768 2016-11-15 03:26:11Z kasakawa $
# Author    : Koji Asakawa
# Contents  : This contains interface and command
# Copyright : MegaChips Technology America Co.
#             *** MegaChips Technology America STRICTLY CONFIDENTIAL ***
#---------------------------------------------------------------------------------------------------
# History
# v1,00 --- Created : Wed Nov 09 22:15:31 2016 v1.00
# v1.20 --- Add: Selecting the range of current measurement
# v1.21 --- Add: SVN version stamp
"""

# -- Includes ---------------------------------------------------------------------------------------

import time
import pyvisa as visa


# -- Class ------------------------------------------------------------------------------------------

class Keithly_2230(object):

    # constructer and destructer -------------------------------------------------------------------
    def __init__(self, Address="GPIB0::30::INSTR", timeout=180000):  # constructer
        #    def __init__(self, _RST=False, Address="GPIB0::30::INSTR"): # old version constructer
        ## Initialize connection with this machine
        self.h = None
        self.rm = visa.ResourceManager()
        self.h = self.rm.open_resource(Address)
        self.visa_TermChar = '\n'  # Termination characters for VISA Read/Write (default=\r\n)
        self.h.timeout = timeout  # 3 minute

    #        ##write codes for initialization
    #        if _RST==True:
    #            self.setInitialyze()
    #            time.sleep(0.8)
    #        self.setSystem_Remote()

    def __del__(self):  # destructer
        # self.setSystem_Local()
        # self.rm.close()
        pass

    # Wrapping function of VISA commands -----------------------------------------------------------
    def VISA_write(self, cmd):
        # self.h.write(cmd,termination=self.visa_TermChar)
        self.h.write(cmd)

    def VISA_read(self):
        buff = self.h.read()
        buff = buff.replace("\n", "")
        buff = buff.replace("\r", "")
        return buff

    def VISA_query(self, cmd):
        self.VISA_write(cmd)
        time.sleep(0.010)
        return self.VISA_read()

    # Instances ------------------------------------------------------------------------------------
    def get_IDN(self):
        return self.VISA_query("*IDN?")

    def set_system_local(self):
        self.VISA_write("SYST:LOC")

    def set_system_remote(self):
        self.VISA_write("SYST:RWL")

    def set_system_default(self):
        self.VISA_write("SYST:POS RST")

    def set_voltage_ch1(self, volt=0.0):
        self.VISA_write("INST:SEL CH1")
        self.VISA_write("VOLT " + str(volt))

    def set_voltage_ch2(self, volt=0.0):
        self.VISA_write("INST:SEL CH2")
        self.VISA_write("VOLT " + str(volt))

    def set_voltage_ch3(self, volt=0.0):
        self.VISA_write("INST:SEL CH3")
        self.VISA_write("VOLT " + str(volt))

    def set_voltage_ch(self, ch=1, volt=0.0):
        self.VISA_write("INST:SEL CH%d" % ch)
        self.VISA_write("VOLT " + str(volt))

    def set_current_ch1(self, amp=0.0):
        self.VISA_write("INST:SEL CH1")
        self.VISA_write("CURR " + str(amp) + "A")

    def set_current_ch2(self, amp=0.0):
        self.VISA_write("INST:SEL CH2")
        self.VISA_write("CURR " + str(amp) + "A")

    def set_current_ch3(self, amp=0.0):
        self.VISA_write("INST:SEL CH3")
        self.VISA_write("CURR " + str(amp) + "A")

    def get_voltage_ch1(self):
        self.VISA_write("INST:SEL CH1")
        return float(self.VISA_query("VOLT?"))

    def get_voltage_ch2(self):
        self.VISA_write("INST:SEL CH2")
        return float(self.VISA_query("VOLT?"))

    def get_voltage_ch3(self):
        self.VISA_write("INST:SEL CH3")
        return float(self.VISA_query("VOLT?"))

    def get_meas_current_ch1(self):
        self.VISA_write("INST:SEL CH1")
        try:
            splitbuf = self.VISA_query(":MEAS:CURR?")
        except:
            splitbuf = 0.0

        try:
            ret = float(splitbuf)
        except:
            ret = float(splitbuf.split(",")[0][:-1])
        return ret

    def get_meas_current_ch2(self):
        self.VISA_write("INST:SEL CH2")
        try:
            splitbuf = self.VISA_query(":MEAS:CURR?")
        except:
            splitbuf = 0.0
        return float(splitbuf)

    def get_meas_current_ch3(self):
        self.VISA_write("INST:SEL CH3")
        try:
            splitbuf = self.VISA_query(":MEAS:CURR?")
        except:
            splitbuf = 0.0
        return float(splitbuf)

    def set_output_on(self):
        self.VISA_write("OUTP 1")

    def set_output_off(self):
        self.VISA_write("OUTP 0")


if __name__ == "__main__":
    K2230 = Keithly_2230(Address="GPIB1::30::INSTR")

    #    try:

    #        PSU12V_K2280.setVoltageLimit(1.25)
    #        PSU12V_K2280.setCurrentLimit(0.8)
    #        PSU18V_K2280.setVoltageLimit(1.8)
    #        PSU18V_K2280.setCurrentLimit(0.8)
    #
    #        time.sleep(1)
    #
    #        PSU18V_K2280.setOutput_ON()
    #        PSU12V_K2280.setOutput_ON()

    print(K2230.getIDN())
    K2230.setCurrentCh1(0.5)
    K2230.setVoltageCh1(1.0)
    K2230.setVoltageCh2(1.8)
    print(K2230.getMeasCurrentCh1())
    print(K2230.getMeasCurrentCh2())
    K2230.setOutput_ON()
    K2230.setOutput_OFF()
#    K2230.setSystem_Local()
#    del K2230
