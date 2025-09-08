# ---------------------------------------------------------------------------------------------------
# Project   : Aston
# Filename  : Agilent_E3631A.py
# Version   : v1.00, 08-25-2024
# Contents  : This contains interface and command
# Copyright : MegaChips Technology America Co.
#             *** MegaChips Technology America STRICTLY CONFIDENTIAL ***
# ---------------------------------------------------------------------------------------------------

# -- Includes ---------------------------------------------------------------------------------------

import time
import pyvisa as visa


# -- Class ------------------------------------------------------------------------------------------

class Agilent_E3631A(object):

    # === Note for this instrument =================================================================
    # - *RST takes over 700msec practically. Wait over 800msec after asserting this command.
    # - first, put the divice "remote mode" before sending command
    # - set a termination code of VISA with '\n' as only 'LF' code. It's necessary to read and write with VISA.
    # ==============================================================================================

    # constructer and destructer -------------------------------------------------------------------
    def __init__(self, Address="GPIB0::5::INSTR"):  # constructer
        ## Initialize connection with this machine
        self.h = None
        rm = visa.ResourceManager()
        self.h = rm.open_resource(Address, read_termination='\r')
        self.h.timeout = 20000
        self.visa_TermChar = '\n'  # Termination characters for VISA Read/Write (default=\r\n)

    # Wrapping function of VISA commands -----------------------------------------------------------
    def VISA_write(self, cmd):
        self.h.write(cmd, termination=self.visa_TermChar)

    def VISA_read(self):
        return self.h.read(termination=self.visa_TermChar)

    def VISA_query(self, cmd):
        self.h.write(cmd, termination=self.visa_TermChar)
        time.sleep(0.010)
        return self.h.read(termination=self.visa_TermChar)

    # Instances ------------------------------------------------------------------------------------
    def get_idn(self):
        return self.VISA_query("*IDN?")

    def set_outputON(self):
        self.VISA_write("OUTPUT 1")

    def set_outputOFF(self):
        self.VISA_write("OUTPUT 0")

    def getISI_stat(self):
        return self.VISA_query("OUTPUT?")

    def sel_Output(self, num):  # 1:P6V 2:P25V 3:N25V
        self.VISA_write("INST:NSEL %d" % num)

    def set_Voltage(self, v):
        self.VISA_write("VOLT %.3e" % v)

    def set_VoltageP6V(self, v):
        self.sel_Output(1)
        self.set_Voltage(v)

    def set_VoltageP25V(self, v):
        self.sel_Output(2)
        self.set_Voltage(v)

    def set_AllVoltage(self, p6v, p25v, n25v):
        self.set_VoltageP6V(p6v)
        self.set_VoltageP25V(p25v)
        self.sel_Output(3)
        self.set_Voltage(n25v)

    def get_Current(self):
        return float(self.VISA_query("MEAS:CURR?"))

    def get_CurrP6V(self):
        try:
            curr = float(self.VISA_query("MEAS:CURR? P6V"))
        except:
            print('Cannot Read value of current')
            curr = 0
        return curr

    def get_CurrP25V(self):
        try:
            curr = float(self.VISA_query("MEAS:CURR? P25V"))
        except:
            print('Cannot Read value of current')
            curr = 0
        return curr

    def get_CurrN25V(self):
        try:
            curr = float(self.VISA_query("MEAS:CURR? N25V"))
        except:
            print('VISA Cannot Read value of current')
            curr = 0
        return curr

    def get_AllCurrent(self):
        dat = []
        dat.append(self.VISA_query("MEAS:CURR? P6V"))
        dat.append(self.VISA_query("MEAS:CURR? P25V"))
        dat.append(self.VISA_query("MEAS:CURR? N25V"))
        return dat

    def get_VoltageP6V(self):
        self.VISA_write('INST:SEL P6V')
        vol = self.VISA_query('VOLT?')
        return float(vol)

    def get_VoltageP25V(self):
        self.VISA_write('INST:SEL P25V')
        vol = self.VISA_query('VOLT?')
        return float(vol)

    def close(self):
        self.h.close()

# if __name__ == '__main__' :
#     ps =   Agilent_E3631A(Address="GPIB1::5::INSTR")
#     ps.set_VoltageP25V(12)
#     ps.set_VoltageP6V(0)
#     print(ps.get_VoltageP25V())
#     print(ps.get_VoltageP6V())

#    print(dm.meas_vol())
