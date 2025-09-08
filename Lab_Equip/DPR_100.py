# -*- coding: utf-8 -*-
"""
Created on Wed Apr 24 14:36:08 2024

@author: CC.Cheng
"""

# -- Includes ---------------------------------------------------------------------------------------
import serial
from serial import Serial
import time

dir_rate = {'HBR3': 0x1E, 'HBR2': 0x14, 'HBR': 0x0A, 'RBR': 0x06}


# -- Class ------------------------------------------------------------------------------------------

class DPR_100(Serial):
    def __init__(self, com=9):
        super(DPR_100, self).__init__(
            port='COM%d' % com,
            baudrate=115200,
            parity=serial.PARITY_NONE,
            timeout=0.1,
            bytesize=serial.EIGHTBITS)

        self.rst_buf()

    def rst_buf(self):
        self.reset_input_buffer()
        self.reset_output_buffer()

    def chk_FW_Version(self):
        self.write(bytearray([0x04, 0x72, 0x1C, 0x6E]))
        rdat = list(bytearray(self.read(10)))
        print(map(hex, rdat))
        print('Major : %x' % rdat[3])
        print('Minor : %x' % rdat[4])
        print('Rev   : %x' % rdat[5])

    def _add_CheckSum(self, data):
        a = sum(data)
        chksum = (((a + 1) ^ 0xFF) + 2) & 0xFF
        return data + [chksum]

    def wait_Ack(self):
        rdat = list(bytearray(self.read(4)))
        cnt = 0
        while 1:
            if len(rdat) < 4:
                time.sleep(0.01)
                rdat += list(bytearray(self.read(1)))
            elif rdat[-4:] == [0x04, 0x72, 0x0C, 0x7E]:
                return 1
            elif rdat[-4:] == [0x04, 0x72, 0x0B, 0x7F]:
                return -1
            else:
                cnt += 1
                if cnt == 100:
                    self.close()
                    raise Exception("No response from DPR-100")

    def DPCD_read(self, addr):
        addr_l = addr & 0xFF
        addr_m = (addr >> 8) & 0xFF
        data = [0x06, 0x72, 0x1A, addr_m, addr_l]
        data = self._add_CheckSum(data)
        # print ("R",list(map(hex,data)))
        self.write(bytearray(data))
        rdat = list(bytearray(self.read(8)))
        if rdat[0:-2] == [0x05, 0x72, 0x1A]:
            return rdat[3]
        else:
            return -1

    def MaxCapability(self, LR, LC, EnFrame=0):
        data = [0x7, 0x72, 0xA0, LC, LR, EnFrame & 1]
        data = self._add_CheckSum(data)
        self.write(bytearray(data))

        return self.wait_Ack()

    def DPCD_RWR(self, addr, wdata):
        ori = hex(self.DPCD_read(addr))
        self.DPCD_write(addr, wdata)
        self.DPCD_read(addr)
        new = hex(self.DPCD_read(addr))
        print("HDCP", hex(addr), "=", new, "from", ori)

    def DPCD_write(self, addr, wdata):
        addr_l = addr & 0xFF
        addr_m = (addr >> 8) & 0xFF
        wdata &= 0xFF
        data = [0x07, 0x72, 0x1B, addr_m, addr_l, wdata]
        data = self._add_CheckSum(data)
        # print ("W",list(map(hex,data)))
        self.write(bytearray(data))

        return self.wait_Ack()

    def HPD_ShortPULSE(self, ms=1):
        # duration &= 0xFF
        data = [0x06, 0x72, 0xA5, ms & 0xFF, ms >> 8 & 0xFF]  # 1 ms interval
        data = self._add_CheckSum(data)
        self.write(bytearray(data))
        return self.wait_Ack()
        print("Short Pulse Issued")

    def HPD_PULSE_H(self):
        data = [0x06, 0x72, 0xA5, 0xFF, 0xFF]
        data = self._add_CheckSum(data)
        self.write(bytearray(data))

    def HPD_PULSE_L(self):
        data = [0x06, 0x72, 0xA5, 0x00, 0x00]
        data = self._add_CheckSum(data)
        self.write(bytearray(data))

    def Set_DPVersion(self, ver):  # 1.1/1.2/1.4
        self.HPD_PULSE_H()
        ver = hex(0x10 + int(ver * 10 - 10))
        self.DPCD_RWR(0x0000, ver)
        self.DPCD_RWR(0x0001, 0x14)
        self.DPCD_RWR(0x0002, 0x04)
        self.DPCD_RWR(0x0003, 0x00)
        self.DPCD_RWR(0x0006, 0x01)
        self.HPD_ShortPULSE()

    def Set_AutoTest(self, LR, LC, PT):
        # DPCD_TEST_LINK_RATE
        self.DPCD_RWR(0x219, LR)
        # DPCD_TEST_LANE_COUNT
        self.DPCD_RWR(0x220, LC)  # Lane
        # DPCD_PHY_TEST_PATTERN
        # 01h = D10.2 test pattern (unscrambled; same as PHY TPS1).
        # 02h = Symbol Error Measurement pattern.
        # 03h = PRBS7.
        # 04h = 80-bit custom pattern.
        # 05h = CP2520 Pattern 1.
        # 06h = CP2520 Pattern 2.
        # 07h = CP2520 Pattern 3 (which is TPS4).
        self.DPCD_RWR(0x248, PT)
        if self.DPCD_read(0x248) == 0x4:
            print('Writing PLTPAT 80 bits')
            self.DPCD_RWR(0x250, 0b00011111)
            self.DPCD_RWR(0x251, 0b01111100)
            self.DPCD_RWR(0x252, 0b11110000)
            self.DPCD_RWR(0x253, 0b11000001)
            self.DPCD_RWR(0x254, 0b00000111)
            self.DPCD_RWR(0x255, 0b00011111)
            self.DPCD_RWR(0x256, 0b01111100)
            self.DPCD_RWR(0x257, 0b11110000)
            self.DPCD_RWR(0x258, 0b11000001)
            self.DPCD_RWR(0x259, 0b00000111)

    def SetL0L1VSPE(self, VS, PE):
        L0vspe = ((PE << 2) + VS)
        L1vspe = ((PE << 2) + VS)
        L1L0vspe = (L1vspe << 4) + L0vspe
        # DPCD_L1L0_ADJ_REQ
        self.DPCD_write(0x206, L1L0vspe)

    def SetL2L3VSPE(self, VS, PE):
        L2vspe = ((PE << 2) + VS)
        L3vspe = ((PE << 2) + VS)
        L3L2vspe = (L3vspe << 4) + L2vspe
        # DPCD_L3L2_ADJ_REQ
        self.DPCD_write(0x207, L3L2vspe)
        # BIT0~1
        # 00b = Voltage Swing Level 0.
        # 01b = Voltage Swing Level 1.
        # 10b = Voltage Swing Level 2.
        # 11b = Voltage Swing Level 3.
        # BIT2~3
        # 00b = Pre-emphasis Level 0.
        # 01b = Pre-emphasis Level 1.
        # 10b = Pre-emphasis Level 2.
        # 11b = Pre-emphasis Level 3.

    def FakeLink(self):
        self.DPCD_RWR(0x200, 0x01)
        self.DPCD_RWR(0x202, 0x77)
        self.DPCD_RWR(0x203, 0x77)
        self.DPCD_RWR(0x204, 0x81)

    def FullAuto(self, ver, LR, LC, PT, VS, PE):
        self.HPD_ShortPULSE()
        self.Set_DPVersion = (ver)
        ###############################################

        # AUTOMATED_TEST_REQUEST
        self.DPCD_RWR(0x201, 0x2)
        # DPCD_TEST_REQ
        self.DPCD_RWR(0x218, 0x8)  # Request Pattern Change
        # DPCD_MAX_DOWNSPREAD
        self.DPCD_RWR(0x03, 0x0)
        ###############################################
        self.Set_AutoTest(LR, LC, PT)  # LinkRate LaneCount Pattern
        ###############################################
        # DPCD_L1L0_ADJ_REQ
        self.SetL0L1VSPE(VS, PE)  # VS PE
        # DPCD_L3L2_ADJ_REQ
        self.SetL2L3VSPE(VS, PE)  # VS PE
        # Fake Link
        self.FakeLink()
        ###############################################
        time.sleep(1)
        self.HPD_ShortPULSE()
        time.sleep(1)
        self.HPD_ShortPULSE()


"""    
if __name__ == '__main__' :
    s = DPR_100(16)
    print(hex(s.DPCD_read(0x202)))
    s.close()
"""