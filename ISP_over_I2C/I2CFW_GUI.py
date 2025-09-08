# -*- coding: utf-8 -*-
"""
Created on Thu Jan  9 11:39:55 2025

@author: CC.Cheng
"""

from aardvark_py import *
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import time, os, sys, tqdm
import ctypes

"""
General Functions

"""
GUI_VERSION = 'ISP over I2C Tool Ver0.0.4'
GUI_SIZE = '440x630'
GUI_THEME = 'classic'  # e.g., 'winnative', 'clam', 'alt', 'default', 'classic', 'vista', 'xpnative'


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
        # else :
        #     print ("Aardvark : Connected to %d%s"%(aa_unique_id(self.handle),s))
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
                    # print('         \n<NACK>:  ', eee)                                           # If NACK occur, Stop the Running
                    raise Exception("I2C NACK.\n")
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
                    # print('\nERROR ERROR ERROR ERROR ERROR ERROR ERROR ERROR ERROR ERROR ERROR ERROR ')
                    raise Exception("I2C ERROR.\n")
                    # sys.exit()

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

    def FastFlashWrite(self, filename, buffer_size=0x1000, progress_callback=None, log_callback=None, stop_flag=None):
        """
        Fast Flash Write Command Complex [BUFFER = 0x1000 is MAX]

        Args:
            filename (str): Path to the binary file to be written.
            buffer_size (int): Size of each chunk to be written.
            progress_callback (function): Function to update progress bar and label.
            log_callback (function): Function to log messages to GUI.
            stop_flag (function): A callable to check if the process should stop.
        """
        import time

        # Start timing
        start_time = time.time()

        # Get file size and calculate total chunks
        bytesize = os.path.getsize(filename)
        total_chunks = -(-bytesize // buffer_size)  # Total chunks (ceil division)
        processed_bytes = 0

        # Logging initialization
        if log_callback:
            log_callback(f"File size: {bytesize} bytes | Buffer size: {buffer_size} bytes")
            log_callback(f"Total chunks: {total_chunks}")

        # Chunk generator
        getChunk = ReadBinDataFrmFile_i2c(filename, buffer_size)
        reg_addr = [0x00, 0x00, 0x00, 0x00]  # Initialize register address

        for i in range(total_chunks):
            if stop_flag and stop_flag():  # Check if the stop flag is set
                if log_callback:
                    log_callback("Stop requested during Fast Flash Write. Aborting...\n")
                return  # Exit the method immediately

            chunk = next(getChunk)
            if not chunk:
                if log_callback:
                    log_callback("File is invalid or end of file reached.")
                break

            # Send special Fast Flash Write command
            FFW = [0x51, 0x85, 0xc2, 0x00, 0x00, 0x03, 0x10, 0xe3]
            self.writeDDC2Bi3_I2C(*FFW, ReadAfrWrite=False)

            # Prepare the chunk data with header and checksum
            data_len = Addr2Byte(len(chunk))
            bigdata = (
                    bytearray(reg_addr)
                    + bytearray(data_len)
                    + bytearray(chunk)
                    + bytearray(Addr2Byte(GetCRC(chunk)))
            )

            # Increment register address
            if reg_addr[2] == 240:
                reg_addr[1] += 1
                reg_addr[2] = 0
            elif reg_addr[1] == 255:
                reg_addr[0] += 1
                reg_addr[1] = 0
                reg_addr[2] = 0
            else:
                reg_addr[2] += 16

            # Simulate write delay
            time.sleep(0.1)
            self.writeDDC2Bi3_I2C(*bigdata, ReadAfrWrite=True)

            # Update progress
            processed_bytes += len(chunk)
            elapsed_time = time.time() - start_time
            percentage = (i + 1) / total_chunks * 100
            speed = (processed_bytes / elapsed_time) / 1000 if elapsed_time > 0 else 0  # Bytes per second
            eta = (total_chunks - (i + 1)) / (i + 1) * elapsed_time if i + 1 > 0 else 0  # Estimated time

            # Update the progress bar and label in GUI
            if progress_callback:
                progress_callback(percentage, speed, eta)

        # Log completion
        if log_callback:
            log_callback("Fast Flash Write Completed\n")

    def Enter_IromMode(self):
        # print('\nRequest to enter force-IROM mode')
        self._send_command(0x07, 0x07, 0x50, 0x00, 0x00, 0x10)

    def Enter_SpiDriverWriteState(self):
        # print('\nRequest to enter ISP driver write state')
        self._send_command(0x07, 0x07, 0x50, 0x00, 0x00, 0x11)

    def Run(self):
        # print('\nRequest to run ISP driver')
        self._send_command(0x07, 0x07, 0x50, 0x00, 0x00, 0x12)  # Request to run ISP driver

    def FlashErase(self):
        # print('\nErase inactive bank')
        self._send_command(0x07, 0x07, 0x50, 0x00, 0x00, 0x20)  # Erase inactive bank
        # print('\nErase Done')

    def HardReset(self):
        # print('\nHard Reset')
        self._send_command(0x07, 0x07, 0x50, 0x00, 0x01, 0x10)  # Hard Reset


class I2C_GUI:
    def __init__(self, root):
        self.root = root
        self.root.title(GUI_VERSION)
        self.root.geometry(GUI_SIZE)

        # Set an icon if available
        self.root.iconbitmap('I2C_application_icon.ico')

        # Slave address
        self.slave_address = tk.StringVar(value="0x73")
        self.stop_requested = False  # Flag to stop the process

        # GUI Elements
        self.create_widgets()

        # I2CM instance
        self.i2cm = None

    def create_widgets(self):
        tk.Label(self.root, text="Slave Address (hex):         \n Jaguar & Mustang = 0x73"
                 ).grid(row=0, column=0, pady=5, padx=5, sticky="nw")
        tk.Entry(self.root, textvariable=self.slave_address).grid(row=0, column=1, pady=5, padx=0)

        # Log Box with Scrollbar
        log_frame = tk.Frame(self.root)
        log_frame.grid(row=1, column=0, columnspan=2, pady=10, padx=10, sticky="nsew")
        self.log_text = tk.Text(log_frame, height=25, width=50, state="disabled", wrap="word")
        scrollbar = tk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.config(yscrollcommand=scrollbar.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Progress Bar with Overlay Label
        self.progress = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress, maximum=100)
        self.progress_bar.grid(row=2, column=0, columnspan=2, pady=5, padx=10, sticky="we")

        # Get the background color of the progress bar
        style = ttk.Style()
        Win_BG = '#6bfaa4'

        style.configure('TProgressbar', background=Win_BG)

        # Create a label with a matching background
        self.progress_label = tk.Label(
            self.root,
            text="Progress: 0% | Speed: 0 B/s | ETA: 0s",
            anchor="center"
        )
        self.progress_label.grid(row=3, column=0, columnspan=2)

        # Buttons
        tk.Button(self.root, text="Flash FW", command=self.start_process).grid(row=4, column=0, pady=10)
        tk.Button(self.root, text="Clean Log", command=self.clean_log).grid(row=4, column=1, pady=10)
        tk.Button(self.root, text="Stop", command=self.stop_process).grid(row=5, column=0, pady=10)  # Stop Button
        tk.Button(self.root, text="Exit", command=self.exit_application).grid(row=5, column=1, pady=10)

    def update_progress(self, percentage, speed, eta):
        """
        Update the progress bar and overlay label with percentage, speed, and ETA.
        """
        self.progress.set(percentage)
        self.progress_bar.update()
        self.progress_label.config(text=f"Progress: {percentage:.2f}% | Speed: {speed:.2f} KB/s | ETA: {eta:.0f}s")
        self.root.update()  # Ensure the GUI updates immediately

    def log_message(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.config(state="disabled")
        self.log_text.see(tk.END)
        self.root.update()  # Ensure the log updates immediately

    def clean_log(self):
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state="disabled")

    def open_file_dialog(self):
        return filedialog.askopenfilename(filetypes=[("BIN File", "*.bin")])

    def initialize_i2cm(self):
        try:
            slave_id = int(self.slave_address.get(), 16)
            self.i2cm = ISP_I2CM(slave_id)
            self.log_message("Initialized I2C with slave address: " + hex(slave_id) + "\n")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to initialize I2C: {e}")

    def stop_process(self):
        """Stop the current process and reset the progress bar."""
        self.log_message("Stop button pressed. Stopping the process...")
        self.stop_requested = True  # Set the flag to stop the process
        self.progress.set(0)  # Reset the progress bar
        self.progress_label.config(text="Progress: 0% | Speed: 0 B/s | ETA: 0s")  # Reset the label
        self.root.update()  # Ensure the GUI updates immediately

    def start_process(self):
        self.stop_requested = False  # Reset the flag at the start of the process
        self.initialize_i2cm()
        if not self.i2cm:
            return

        try:
            # Step 1: Enter Force-IROM Mode
            self.log_message("Step 1: Entering Force-IROM Mode\n")
            self.i2cm.Enter_IromMode()
            time.sleep(1)
            self.root.update()

            # Step 2: Enter ISP Driver Write State
            self.log_message("Step 2: Entering ISP Driver Write State\n")
            self.i2cm.Enter_SpiDriverWriteState()
            time.sleep(1)
            self.root.update()

            # Step 3: Write ISP Driver
            self.log_message("Step 3: Writing ISP Driver\n")
            i2c_isp_drv_file = "jaguar-b0_mca_i2c_isp_driver_payload_Disable_WDT.bin"
            if not i2c_isp_drv_file:
                raise Exception("No ISP driver file selected.\n")
            self.log_message(f"Selected ISP Driver File:\n {i2c_isp_drv_file}\n")
            self.i2cm.FastFlashWrite(
                i2c_isp_drv_file,
                progress_callback=self.update_progress,
                log_callback=self.log_message,
                stop_flag=lambda: self.stop_requested  # Pass the stop flag
            )
            if self.stop_requested:
                raise Exception("Process stopped by user.")
            time.sleep(1)

            # 4: Run ISP Driver
            self.log_message("Step 4: Running ISP Driver\n")
            self.i2cm.Run()
            time.sleep(1)

            # 5: Erase Inactive Bank
            self.log_message("Step 5: Erasing Inactive Bank\n")
            self.i2cm.FlashErase()
            self.log_message("Erase Done\n")
            time.sleep(1)

            # 6: Flash Firmware
            self.log_message("Step 6: Flashing Firmware\n")
            fw_bin_file = self.open_file_dialog()
            if not fw_bin_file:
                raise Exception("No firmware file selected.\n")
            self.log_message(f"Selected Firmware File:\n {fw_bin_file}\n")
            self.i2cm.FastFlashWrite(
                fw_bin_file,
                progress_callback=self.update_progress,
                log_callback=self.log_message,
                stop_flag=lambda: self.stop_requested  # Pass the stop flag
            )
            if self.stop_requested:
                raise Exception("Process stopped by user.")
            time.sleep(3)

            # 7: Hard Reset
            self.log_message("Step 7: Performing Hard Reset\n")
            self.i2cm.HardReset()

            messagebox.showinfo("Success", "I2C Flashing process completed successfully!\n")
        except Exception as e:
            self.log_message(f"Error: {e}")
            if self.stop_requested:
                self.log_message("Performing Hard Reset after stop request...")
                self.i2cm.HardReset()
                self.stop_requested = False  # Reset the flag
        finally:
            if self.i2cm:
                self.i2cm.close_i2c()

    def exit_application(self):
        if self.i2cm:
            self.i2cm.close_i2c()
        self.root.destroy()
        sys.exit()


if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style(root)
    style.theme_use(GUI_THEME)
    app = I2C_GUI(root)
    root.mainloop()
