# -*- coding: utf-8 -*-
"""
Created on Wed Oct  5 09:34:25 2022

@author: CaesalCheng
"""

import time
from PyGprobe import PyGprobe
import sys
import time

i = 0
import csv

AUX_REPLY = 1
aux = 1

COM = 12
filename = "60000_61CFF_Result.csv"
f = open(filename, "w")
f.write("%s,%s,%s\n" % ('Register_Num', 'AUX_Reply', 'AUX_Result'))

gp = PyGprobe(COM, "JaguarB0")

# gp.write_reg(0x4003c228,0xFF)
for i in range(393216, 400640):
    gp.write_reg(0x4003c21C, i)
    gp.write_reg(0x4003c218, 0x000000a9)  # Read command
    # gp.write_reg(0x4003c218,0x000000a8)
    AUX_DATA_REPLY = hex(gp.read_reg(0x4003c23c))
    REG_NUM = hex(i)

    if AUX_DATA_REPLY == '0x00' or AUX_DATA_REPLY == '0x0':
        aux = ('RD_00_ACK')
    else:
        aux = ('RD_ACK')

    f.write("%s,%s,%s\n" % (REG_NUM, AUX_DATA_REPLY, aux))

f.close()
gp.close_gp()

