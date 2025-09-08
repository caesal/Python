# -*- coding: utf-8 -*-
"""
Created on Mon Apr 29 16:11:07 2024

@author: Caesal Cheng
"""

from DPR_100 import DPR_100
from DPT_200 import DPT_200

###############################################

LinkRate = [0x6, 0xA, 0x14]
LaneCount = [0x1, 0x2]
VoltageSwing = [0, 1, 2, 3]
PreEmphasis = [0, 1, 2, 3]

VoltageSwing_PreEmphasis = [[vs, pe] for vs in VoltageSwing for pe in PreEmphasis if vs + pe <= 3]


# print(VoltageSwing_PreEmphasis)

###############################################
class DPT200toDPR100():
    def LinkTraining_DPT200toDPR100(self, LR, LC, VS, PE):
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


###############################################
sink = DPR_100(9)
source = DPT_200(4)
LT = DPT200toDPR100()
###############################################
LR = 0xA
LC = 0x1

LT.LinkTraining_DPT200toDPR100(LR, LC, 0, 0)

sink.close()
source.close()
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

