# -*- coding: utf-8 -*-
"""
Created on Thu Feb  2 17:18:26 2023

    FINAL BATCHING FW SEQUENCE VIA I2C


Python Script for ISP over I2C Flashing by using aardvark I2C master
Need to set bit0 of 0x40010400(I2C_SLAVE_EN) in MCDP5200 beforehand.
"""

from aardvark_py import *
import time, os, tqdm
import sys
import tkinter
import tkinter.filedialog

"""
General Functions

"""


def GUI_OpenFileName():
    root = tkinter.Tk()
    root.withdraw()
    fTyp = [("BIN File", "*.bin")]
    filename = tkinter.filedialog.askopenfilename(filetypes=fTyp)
    return filename


def ReadBinDataFrmFile_i2c(filename, dlength):
    '''
    Open .BIN file to program the ISP and SPI

    '''
    f = open(filename, "rb")
    while 1:
        chunk = f.read(dlength)
        yield chunk


def Addr2Byte(addr):
    '''
    Change the format from Address to Byte

    '''
    wadr_list = []
    for shift in range(24, -1, -8):
        wadr_list.append((addr >> shift) & 0xFF)
    return wadr_list


def GetCRC(buf):
    '''
    Generate CRC at the end of each chunk

    '''
    crc = 0x1021
    for byte_buf in bytearray(buf):
        data = byte_buf
        for i in range(8):
            flag = data ^ (crc >> 8)
            crc = (crc << 1) & 0xFFFF
            if flag & 0x80:
                crc = (crc ^ 0x1021) & 0xFFFF
            data <<= 1
    return crc


"""
ISP over I2C master class

"""


class ISP_I2CM(object):
    bitrate = 100  # kHz

    def __init__(self, slave_id=0x14, unique_id=-1, re_open=True):

        '''
        Initialized All of Constants that we will use in next step

        '''
        # Initialized All of Constants that we will use in next step
        self.SlaveID = slave_id
        self.re_open = re_open
        self.unique_id = unique_id
        self.handle = 0
        self.open_i2c()

    def open_i2c(self):

        '''
        Established the Connection with I2C AARDVARK

        '''
        (num, ports, unique_ids) = aa_find_devices_ext(16, 16)  # Find all the attached devices
        if num == 0:
            raise Exception("Aardvark I2C is not connected!!")
        s = ""
        for port, u_id in zip(ports, unique_ids):  # Open port
            if self.unique_id == -1 or self.unique_id == u_id:
                if self.re_open == True and port >= 0x8000:
                    for h in range(1, 100):
                        if port & 0x7FFF == aa_port(h):
                            s = "(Re-connect)"
                            self.handle = h
                            break
                    else:
                        raise Exception("Failed to re-open aardvark I2c")
                    break
                elif port < 0x8000:
                    self.handle = aa_open(port)
                    break
        else:
            raise Exception("No available aardvark I2C, Plase check <re_open> or <unique_id>")

        if self.handle < 0:
            raise Exception("Problem with I2C open")
        else:
            print("Aardvark : Connected to %d%s" % (aa_unique_id(self.handle), s))
        aa_configure(self.handle, AA_CONFIG_SPI_I2C)  # Ensure that the I2C subsystem is enabled
        bitrate = aa_i2c_bitrate(self.handle, self.bitrate)  # Set the bitrate

    def close_i2c(self):

        '''
        Break the Connection with I2C AARDVARK

        '''
        aa_close(self.handle)

    def writeDDC2Bi3_I2C(self, *wdata, ReadAfrWrite=True):

        '''
        Write & Read Complex

        '''
        data_out = array('B')
        for w in wdata:
            data_out.append(w & 0xFF)
        if ReadAfrWrite == True:  # Normal Case: Read I2C after Write
            c1, c2, rval, c3 = (0, 0, 0, 0)
            l = aa_i2c_write(self.handle, self.SlaveID, AA_I2C_NO_FLAGS, data_out)  # WRITE COMMAND
            for i in range(10000):  # READ LOOP
                c1, c2, rval, c3 = (0, 0, 0, 0)
                c1, rval = aa_i2c_read(self.handle, self.SlaveID, AA_I2C_NO_FLAGS, 8)  # READ COMMAND
                eee = [varx[1] for varx in enumerate(list(map(hex, rval))) if varx[0] in [5, 6]]  # PULL OUT Bit 5&6
                eeee = [varx[1] for varx in enumerate(list(map(hex, rval))) if varx[0] in [1]]  # PULL OUT Bit 1

                if eee == ['0x3', '0xb']:
                    print('         \n<NACK>:  ', eee)  # If NACK occur, Stop the Running
                    break
                elif eee == ['0x3', '0xc']:
                    # print('          \n<ACK>:  ', eee)                                          # Always ACK so disable this print msg for clean up
                    break
                elif eeee == ['0x80']:
                    # print('\n------------ Waiting, Device is Busy. -------------------')        #Waiting always happen, so disable this print msg for clean up
                    c1, c2, rval, c3 = (0, 0, 0, 0)
                    c1, rval = aa_i2c_read(self.handle, self.SlaveID, AA_I2C_NO_FLAGS,
                                           8)  # Send READ COMMAND again until receive ACK
                    time.sleep(0.5)
                else:
                    print('\nERROR ERROR ERROR ERROR ERROR ERROR ERROR ERROR ERROR ERROR ERROR ERROR ')
                    sys.exit()

        else:
            c1, c2, rval, c3 = (0, 0, 0, 0)  # Special Case: Write I2C without READ back
            l = aa_i2c_write(self.handle, self.SlaveID, AA_I2C_NO_FLAGS, data_out)  # WRITE COMMAND

    def _send_command(self, *data):

        '''
        COMBINING HEADER & RAW DATA & CHECKSUM

        '''
        data = list(data)
        length = 3 + len(data)
        prefix = [0xC2, 0x00, 0x00]
        header = [0x51, 0x80 | length] + prefix
        checksum = 0
        wdata = header + data
        for w in wdata:
            checksum ^= w
        checksum ^= 0xE6
        checksum = [checksum]
        ddd = header + data + checksum
        self.writeDDC2Bi3_I2C(*ddd)

    def FastFlashWrite(self, filename, buffer_size=0x1000):

        '''
        FASTFLASHWRITE COMMAND COMPLEX [BUFFER = 0x1000 is MAX]

        '''
        print('\nFast Flash Write Start')
        print('\nWriting File Name : ', filename)
        print('\n ')
        bytesize = os.path.getsize(filename)
        cnt_full = -(-bytesize // buffer_size)  # Division & Round Up each CHUNK
        getChunk = ReadBinDataFrmFile_i2c(filename, buffer_size)
        reg_addr = [0x00, 0x00, 0x00, 0x00]
        data_len = [0x00, 0x00, 0x00, 0x00]
        pbar = tqdm.tqdm(range(cnt_full))
        pbar.set_description("Fast Flash Write")
        for i in pbar:
            chunk = next(getChunk)
            if chunk == '' or len(chunk) == 0:
                print('File is Invalid')
                break
            else:
                FFW = [0x51, 0x85, 0xc2, 0x00, 0x00, 0x03, 0x10, 0xe3]  # WRITE COMMAND in SPECIAL CASE
                self.writeDDC2Bi3_I2C(*FFW, ReadAfrWrite=False)
                data_len = Addr2Byte(len(chunk))
                bigdata = bytearray(reg_addr) + bytearray(data_len) + bytearray(chunk) + bytearray(
                    Addr2Byte(GetCRC(chunk)))
                if reg_addr[2] == 240:  # REGISTER ADDRESS ADDING UP
                    reg_addr[1] += 1
                    reg_addr[2] = 0
                elif reg_addr[1] == 255:
                    reg_addr[0] += 1
                    reg_addr[1] = 0
                    reg_addr[2] = 0
                else:
                    reg_addr[2] += 16
                time.sleep(0.5)
                self.writeDDC2Bi3_I2C(*bigdata, ReadAfrWrite=True)  # WRITE COMMAND with each CHUNK
        pbar.close()
        print('\nFast Flash Write Done')

    def Debug_On(self):
        print('\nDebug ON')
        self._send_command(0x02, 0x09)

    def Enter_IromMode(self):
        print('\nRequest to enter force-IROM mode')
        self._send_command(0x07, 0x07, 0x50, 0x00, 0x00, 0x10)

    def Enter_SpiDriverWriteState(self):
        print('\nRequest to enter ISP driver write state')
        self._send_command(0x07, 0x07, 0x50, 0x00, 0x00, 0x11)

    def Run(self):
        print('\nRequest to run ISP driver')
        self._send_command(0x07, 0x07, 0x50, 0x00, 0x00, 0x12)  # Request to run ISP driver

    def FlashErase(self):
        print('\nErase inactive bank')
        self._send_command(0x07, 0x07, 0x50, 0x00, 0x00, 0x20)  # Erase inactive bank
        print('\nErase Done')

    def HardReset(self):
        print('\nHard Reset')
        i2cm._send_command(0x07, 0x07, 0x50, 0x00, 0x01, 0x10)  # Hard Reset


# =========== #
#     Main    #
# =========== #

if __name__ == '__main__':
    """
    I2CM instance
    Slave Address is 0x73 (0xE6/7 in 8bit address)

    """
    Slave_address = 0x73
    i2cm = ISP_I2CM(Slave_address)

    print('\nI2C Transaction Excel created')

    """
    Below sequence is same as Kinetic_Smart_Connectivity_ISP_over_I2C_Interface_KTM5xxx.pptx
    Please refer p.11 on the doc
    """
    # Pre: Debug Pn
    # i2cm.Debug_On()
    # time.sleep(0.2)

    # 1 : Request to Enter Force-IROM mode
    i2cm.Enter_IromMode()  # Request to enter force-IROM mode
    time.sleep(0.2)

    # 2 : Reqest to enter ISP dirver write state
    i2cm.Enter_SpiDriverWriteState()  # Request to enter ISP driver write state
    time.sleep(0.1)

    # 3 : Write isp over i2c driver
    i2c_isp_drv_file = GUI_OpenFileName()
    i2cm.FastFlashWrite(i2c_isp_drv_file)  # Set driver file

    # 4 : Request to run ISP driver
    i2cm.Run()
    time.sleep(0.1)

    # 5 : Erase inactive bank
    i2cm.FlashErase()
    time.sleep(0.5)

    # 6 : Flash fw bin file
    fw_bin_file = GUI_OpenFileName()
    i2cm.FastFlashWrite(fw_bin_file)  # Set driver file
    time.sleep(3)

    # 7 : Hard Reset
    i2cm.HardReset()

    i2cm.close_i2c()
