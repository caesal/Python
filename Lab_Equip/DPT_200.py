# ---------------------------------------------------------------------------------------------------
# Project   : Aston
# Filename  : Agilent_E3631A.py
# Version   : v1.00, 08-25-2016
# Author    : Mizuki Yurimoto
# Contents  : This contains interface and command
# Copyright : MegaChips Technology America Co.
#             *** MegaChips Technology America STRICTLY CONFIDENTIAL ***
# ---------------------------------------------------------------------------------------------------

# -- Includes ---------------------------------------------------------------------------------------
import serial
from serial import Serial
import time

dir_rate = {'HBR3': 0x1E, 'HBR2': 0x14, 'HBR': 0x0A, 'RBR': 0x06}


# -- Class ------------------------------------------------------------------------------------------

class DPT_200(Serial):
    def __init__(self, com=9):
        super(DPT_200, self).__init__(
            port='COM%d' % com,
            baudrate=115200,
            parity=serial.PARITY_NONE,
            timeout=0.1,
            bytesize=serial.EIGHTBITS)

        self.rst_buf()

    def rst_buf(self):
        self.reset_input_buffer()
        self.reset_output_buffer()

    def DPT_read(self, addr):
        rdat = []
        rdat = list(bytearray(self.read(1)))
        cnt = 0
        while 1:
            dsum = sum(rdat) & 0xFF
            chksum = ((dsum + 1) ^ 0xFF) + 2
            rdat = rdat + list(bytearray(self.read(1)))
            chksum = ((dsum + 1) ^ 0xFF) + 2

            if chksum == rdat[-1]:
                return rdat
            elif cnt == 0x100:
                return -1
            else:
                cnt += 1

    def chk_FW_Version(self):
        self.write(bytearray([0x04, 0x72, 0x1C, 0x6E]))
        rdat = list(bytearray(self.read(10)))
        print(map(hex, rdat))
        print('Major : %x' % rdat[3])
        print('Minor : %x' % rdat[4])
        print('Rev   : %x' % rdat[5])

    def _add_CheckSum(self, data):
        #        print(data)
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
                    raise Exception("No response from DPT200")

    def DPCD_read(self, addr):
        addr_l = addr & 0xFF
        addr_m = (addr >> 8) & 0xFF
        data = [0x06, 0x72, 0x1A, addr_m, addr_l]
        data = self._add_CheckSum(data)
        self.write(bytearray(data))
        rdat = list(bytearray(self.read(8)))
        if rdat[0:-2] == [0x05, 0x72, 0x1A]:
            return rdat[3]
        else:
            return -1

    def DPCD_Nread(self, addr, n):
        addr_l = addr & 0xFF
        addr_m = (addr >> 8) & 0xFF
        data = [0x08, 0x72, 0x5B, addr_l, addr_m, 0x00, n]
        data = self._add_CheckSum(data)
        self.write(bytearray(data))
        rdat = list(bytearray(self.read(4 + n)))
        if rdat[:3] == [4 + n, 0x72, 0x5B]:
            return rdat[3:-1]
        else:
            return [-1]

    def DPCD_write(self, addr, wdata):
        addr_l = addr & 0xFF
        addr_m = (addr >> 8) & 0xFF
        wdata &= 0xFF
        data = [0x07, 0x72, 0x1B, addr_m, addr_l, wdata]
        data = self._add_CheckSum(data)
        self.write(bytearray(data))

        return self.wait_Ack()

    def I2CovrAUX_write(self, slave_addr, *wdata):
        data = [0x07 + len(wdata), 0x72, 0x66, slave_addr, 1, len(wdata)] + list(wdata)
        data = self._add_CheckSum(data)
        # print (list(map(hex,data)))
        self.write(bytearray(data))

    def DPCD_Nwrite(self, saddr, *wdata):
        addr_l = saddr & 0xFF
        addr_m = (saddr >> 8) & 0xFF
        n = len(*wdata)
        #        print(n)
        #        print(*wdata)
        if n == 0:
            print('No Write Date : Return -1')
            return -1
        data = [0x08 + n, 0x72, 0x5C, addr_l, addr_m, 0x00, n] + list(*wdata)
        data = self._add_CheckSum(data)
        self.write(bytearray(data))
        return self.wait_Ack()

    def cr_done_phase(self, data_rate='HBR3', lanecnt=4):
        bw = {'HBR3': 0x1E, 'HBR2': 0x14, 'HBR': 0x0A, 'RBR': 0x06}
        self.DPCD_write(0x600, 0x02)
        self.DPCD_write(0x600, 0x01)  # Power Down
        self.DPCD_write(0x102, 0x0)  # 0= Link Trainning disabled
        self.DPCD_write(0x107, 0x10)  # 0x10 Enable frequency down spread :0x00 is disabled
        self.DPCD_Nwrite(0x100, [bw[data_rate], lanecnt])  # nLane = 4 : Set BW and Lane Count in 0x100 and 0x101

        # ====CLOCK RECOVERY PHASE ========
        # Set DUT Pattern (JBERT patter is set to Clock above)
        self.DPCD_write(0x102, 0x1)  # Set Clock Pattern(TPS1)  - D10.2 symbols  Scrambler enabled[5]?
        self.DPCD_Nwrite(0x103, [0x00, 0x00, 0x00,
                                 0x00])  # All 4 lanes Voltage Level 0 and Pre-Emphasis Level0 for all lanes 0,1,2,3
        time.sleep(1)
        # Check Frequency Lock Status for Lane= ML
        cr_done_status_pass = 0
        for i in range(lanecnt):
            cr_done_status_pass |= 0x1 << (4 * i)
        cr_done_status = self.DPCD_Nread(0x202, 2)
        cr_done_status = cr_done_status[1] << 8 | cr_done_status[0]
        print("CR_DONE :: " + hex(cr_done_status))
        return cr_done_status_pass == cr_done_status

    def cr_done_phycts(self, data_rate='HBR3', lut=0):
        bw = {'HBR3': 0x1E, 'HBR2': 0x14, 'HBR': 0x0A, 'RBR': 0x06}
        self.DPCD_write(0x600, 0x02)
        self.DPCD_write(0x600, 0x01)  # Power Down
        self.DPCD_write(0x102, 0x0)  # 0= Link Trainning disabled
        self.DPCD_write(0x107, 0x10)  # 0x10 Enable frequency down spread :0x00 is disabled
        self.DPCD_write(0x270,
                        0x80 | lut << 4)  # Enable PHY CTS test for Lanex 0x80 for L0, 0x90 for L1, 0xA0 for L2, 0xB0 for L3
        self.DPCD_Nwrite(0x100, [bw[data_rate], 4])  # nLane = 4 : Set BW and Lane Count in 0x100 and 0x101

        # ====CLOCK RECOVERY PHASE ========
        # Set DUT Pattern (JBERT patter is set to Clock above)
        self.DPCD_write(0x102, 0x1)  # Set Clock Pattern(TPS1)  - D10.2 symbols  Scrambler enabled[5]?
        self.DPCD_Nwrite(0x103, [0x00, 0x00, 0x00,
                                 0x00])  # All 4 lanes Voltage Level 0 and Pre-Emphasis Level0 for all lanes 0,1,2,3
        time.sleep(1)
        # Check Frequency Lock Status for Lane= ML
        cr_done_status_pass = (0x01 if lut in [0, 2] else 0x10)
        cr_done_status = self.DPCD_Nread(0x202, 2)[int(lut / 2)]
        print("CR_DONE :: " + hex(cr_done_status))
        return cr_done_status_pass == cr_done_status

    def cr_done_phycts_no600(self, data_rate='HBR3', lut=0):
        bw = {'HBR3': 0x1E, 'HBR2': 0x14, 'HBR': 0x0A, 'RBR': 0x06}
        #        self.DPCD_write(0x600,0x02)
        self.DPCD_write(0x600, 0x01)  # Power Down
        self.DPCD_write(0x102, 0x0)  # 0= Link Trainning disabled
        self.DPCD_write(0x107, 0x10)  # 0x10 Enable frequency down spread :0x00 is disabled
        self.DPCD_write(0x270,
                        0x80 | lut << 4)  # Enable PHY CTS test for Lanex 0x80 for L0, 0x90 for L1, 0xA0 for L2, 0xB0 for L3
        self.DPCD_Nwrite(0x100, [bw[data_rate], 4])  # nLane = 4 : Set BW and Lane Count in 0x100 and 0x101

        # ====CLOCK RECOVERY PHASE ========
        # Set DUT Pattern (JBERT patter is set to Clock above)
        self.DPCD_write(0x102, 0x1)  # Set Clock Pattern(TPS1)  - D10.2 symbols  Scrambler enabled[5]?
        self.DPCD_Nwrite(0x103, [0x00, 0x00, 0x00,
                                 0x00])  # All 4 lanes Voltage Level 0 and Pre-Emphasis Level0 for all lanes 0,1,2,3
        time.sleep(1)
        # Check Frequency Lock Status for Lane= ML
        cr_done_status_pass = (0x01 if lut in [0, 2] else 0x10)
        cr_done_status = self.DPCD_Nread(0x202, 2)[int(lut / 2)]
        print("CR_DONE :: " + hex(cr_done_status))
        return cr_done_status_pass == cr_done_status

    def symbol_lock_phycts(self, tps=0x7, lut=0):
        # Change DUT Pattern
        self.DPCD_write(0x102, tps)  # Set Symbol Pattern
        # Important to Write VSL and PEL registers again after pattern setting
        self.DPCD_Nwrite(0x103, [0x00, 0x00, 0x00,
                                 0x00])  # All 4 lanes Voltage Level 0 and Pre-Emphasis Level0 for all lanes 0,

        # Check Symbol Lock Status for Lane= ML
        symbol_lock_status_pass = 0x07 if lut in [0, 2] else 0x70
        #        symbol_lock = self.DPCD_Nread(0x202,6)[int(lut/2)]
        symbol_lock = self.DPCD_Nread(0x202, 2)[int(lut / 2)]
        print("SYMBOL_LOCK :: " + hex(symbol_lock))

        #        self.DPCD_write(0x102,0x0)  # 0= Trainning disabled
        return symbol_lock_status_pass == symbol_lock

    def symbol_lock_phase(self, tps=0x7, lanecnt=4):
        # Change DUT Pattern
        self.DPCD_write(0x102, tps)  # Set Symbol Pattern
        # Important to Write VSL and PEL registers again after pattern setting
        self.DPCD_Nwrite(0x103, [0x00, 0x00, 0x00,
                                 0x00])  # All 4 lanes Voltage Level 0 and Pre-Emphasis Level0 for all lanes 0,

        # Check Symbol Lock Status for Lane= ML
        symbol_lock_status_pass = 0
        for i in range(lanecnt):
            symbol_lock_status_pass |= 0x7 << (4 * i)
        symbol_lock = self.DPCD_Nread(0x202, 2)
        symbol_lock = symbol_lock[1] << 8 | symbol_lock[0]
        print("SYMBOL_LOCK :: " + hex(symbol_lock))
        return symbol_lock_status_pass == symbol_lock

    def get_errcnt(self, ml):
        mlcnt = self.DPCD_Nread(0x210 + ml * 2,
                                2)  # 210h/211h for Lane0, 212h/213h for Lane1,214h/215h for Lane2,216h/217h for Lane3
        mlcnt = mlcnt[1] << 8 | mlcnt[0]
        return mlcnt

    def test_TPS4_seq(self):
        self.DPCD_write(0x102, 0x00)  # Pattern
        self.DPCD_write(0x10b, 0x05)  #

    #    print hex(dpt.DPCD_read(0x202))

    def set_LTparam(self, drate='HBR3', lane=4):
        data = [0x0A, 0x72, 0x52, 0x01, 0x01, 0x00, 0x00, 0x02, 0x01]
        data = self._add_CheckSum(data)
        self.write(bytearray(data))
        rdat = list(bytearray(self.read(8)))
        time.sleep(0.1)

        if rdat == [0x04, 0x72, 0x0C, 0x7E]:
            return 1
        else:
            return -1

    def set_TXEQ_by_request(self):
        time.sleep(0.1)
        adj_req = self.DPCD_Nread(0x206, 2)
        vol0 = (adj_req[0] & (0x3 << 0)) >> 0
        pre0 = (adj_req[0] & (0x3 << 2)) >> 2
        vol1 = (adj_req[0] & (0x3 << 4)) >> 4
        pre1 = (adj_req[0] & (0x3 << 6)) >> 6
        vol2 = (adj_req[1] & (0x3 << 0)) >> 0
        pre2 = (adj_req[1] & (0x3 << 2)) >> 2
        vol3 = (adj_req[1] & (0x3 << 4)) >> 4
        pre3 = (adj_req[1] & (0x3 << 6)) >> 6
        req0 = (pre0 << 3) + vol0
        req1 = (pre1 << 3) + vol1
        req2 = (pre2 << 3) + vol2
        req3 = (pre3 << 3) + vol3
        adj_rpl = [req0, req1, req2, req3]
        time.sleep(0.1)
        self.DPCD_Nwrite(0x103, adj_rpl)
        return adj_rpl

# if __name__ == '__main__' :
#    s = DPT_200(4)
#    rply = s.DPCD_read(0x202)
#    s.I2CovrAUX_write(0x37<<1,0x51,0x82,0x01,0x10,0xAC)
#    # s.I2CovrAUX_write(0xA0,0x80)
#    import sys
#    s.close()
## Read DPCD capability
#    sys.exit()
#
#
## Set data rate
#    time.sleep(0.1)
#
## Set pattern and TXEQ
#    time.sleep(0.1)
#    s.DPCD_Nwrite(0x102, [0x21, 0, 0, 0, 0])
#
## Read sink status
#    time.sleep(0.1)
#    rply = s.DPCD_Nread(0x202,2)
#
#    adj_rpl = s.set_TXEQ_by_request()
#    time.sleep(0.1)
#    rply = s.DPCD_Nread(0x202,2)
#
## Set pattern and TXEQ
#    time.sleep(0.1)
#    s.DPCD_Nwrite(0x102, [0x23] + adj_rpl)
#    rply = s.DPCD_Nread(0x202,2)
#    s.set_TXEQ_by_request()
#    rply = s.DPCD_Nread(0x202,2)
#
#
#    s.DPCD_write(0x102, 0)
##    s.DPCD_write(0x600, 5)


#    print map(hex,s.DPCD_Nread(0x210,4))
#    time.sleep(0.1)
#    print map(hex,s.DPCD_Nread(0x210,4))
#    print hex(s.DPCD_read(0x00))
#    print hex(s.DPCD_read(0x01))
#    print hex(s.DPCD_read(0x02))
##
##    s.chk_FW_Version()
##    dats = s.DPCD_Nread(0x00,4)
##    print map(hex,dats)
##    print hex(s.DPCD_read(0x0000))
##    print s.DPCD_write(0x0100,0x1E)
##    print s.set_LTparam()
##    print '-------------'
#    s.close()

