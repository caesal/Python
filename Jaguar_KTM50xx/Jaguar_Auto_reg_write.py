# -*- coding: utf-8 -*-
"""
Created on Tue Oct  4 13:55:45 2022

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
# filename = "100_10F_Result.csv" #File Name
# f = open(filename,"w")
# f.write("%s,%s,%s\n"%('Register_Num','AUX_Reply','AUX_Result'))


gp = PyGprobe(COM, "JaguarB0")

gp.write_reg(0x4003c228, 0xFF)  # Required 0xFF for certain Register
for i in range(259, 272):
    gp.write_reg(0x4003c21C, i)  # identify the Register
    gp.write_reg(0x4003c218, 0x000000a9)  # Read command
    gp.write_reg(0x4003c218, 0x000000a8)  # Write command
    gp.write_reg(0x4003c218, 0x000000a9)  # Read command
    AUX_REPLY = hex(gp.read_reg(0x4003c238))  # Replied value from Register
    REG_NUM = hex(i)

    if AUX_REPLY == '0x100' or AUX_REPLY == '0x0':
        aux = ('AUX_ACK')
    elif AUX_REPLY == '0x101' or AUX_REPLY == '0x101':
        aux = ('AUX_NACK')
    elif AUX_REPLY == '0x2':
        aux = ('AUX_DEFER')
    # f.write("%s,%s,%s\n"%(REG_NUM,AUX_REPLY,aux))

# f.close()
gp.close_gp()




