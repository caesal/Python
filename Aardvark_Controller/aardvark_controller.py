# -*- coding: utf-8 -*-
"""
Created on Thu Feb 13 11:32:00 2025

@author: CC.Cheng
"""

import time
from array import array
import xml.etree.ElementTree as ET
import logging
from aardvark_py import *
import os

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)


class AardvarkController:
    def __init__(self, i2c_address, bitrate=400):
        self.i2c_address = i2c_address
        self.bitrate = bitrate
        self.handle = None

    def open(self):
        """Open the Aardvark device."""
        self.handle = aa_open(0)
        if self.handle <= 0:
            raise RuntimeError("Error: Unable to open Aardvark device.")
        print("Aardvark opened successfully!")
        self.configure()

    def configure(self):
        """Configure the Aardvark for I2C communication."""
        aa_configure(self.handle, AA_CONFIG_SPI_I2C)
        aa_i2c_bitrate(self.handle, self.bitrate)
        aa_i2c_pullup(self.handle, AA_I2C_PULLUP_BOTH)
        aa_target_power(self.handle, AA_TARGET_POWER_BOTH)
        print(f"Aardvark configured with I2C Bitrate: {self.bitrate} kHz")

    def close(self):
        """Close the Aardvark device."""
        if self.handle is not None:
            aa_close(self.handle)
            self.handle = None
            print("Aardvark closed.")

    def basic_write(self, register, data):
        """Write data to a specific register."""
        if self.handle is None:
            raise RuntimeError("Aardvark device is not open.")

        # Ensure data is formatted correctly as a list of unsigned bytes
        data_out = array('B', [register & 0xFF] + data)  # Convert to an array of bytes

        # Send data to the I2C slave
        num_written = aa_i2c_write(self.handle, self.i2c_address, AA_I2C_NO_FLAGS, data_out)

        if num_written < 0:
            logger.error(f"Error writing to register {hex(register)}: {num_written}")
        elif num_written != len(data_out):
            logger.warning(f"Warning: Expected to write {len(data_out)} bytes, but only wrote {num_written}.")
        else:
            logger.info(f"Successfully wrote {[hex(b) for b in data]} to register {hex(register)}")

    def read_register(self, register, length):
        """Read data from a specific register."""
        if self.handle is None:
            raise RuntimeError("Aardvark device is not open.")

        # Convert register to a properly formatted list of bytes
        register_address = array('B', [register & 0xFF])  # Ensure it's a list of unsigned bytes

        # Send register address to the I2C slave
        num_written = aa_i2c_write(self.handle, self.i2c_address, AA_I2C_NO_FLAGS, register_address)

        if num_written < 0:
            logger.error(f"Error sending register address {hex(register)}: {num_written}")
            return None
        elif num_written != len(register_address):
            logger.warning(f"Warning: Expected to write {len(register_address)} bytes, but only wrote {num_written}.")

        # Read requested number of bytes
        num_read, data_in = aa_i2c_read(self.handle, self.i2c_address, AA_I2C_NO_FLAGS, length)

        if num_read < 0:
            logger.error(f"Error reading from register {hex(register)}: {num_read}")
            return None
        elif num_read != length:
            logger.warning(f"Warning: Expected to read {length} bytes, but only read {num_read}.")

        # Convert data to hex format
        data_list = [hex(byte) for byte in data_in]

        # Format output to show hex values
        logger.info(f"Read {data_list} from register {hex(register)}")

        return data_list  # Return as a list of hex values

    def write_register(self, register, data):
        """Special write function: read before write, write, then verify with a read-back."""
        if self.handle is None:
            raise RuntimeError("Aardvark device is not open.")

        length = len(data)  # Determine the length of data to write

        # Step 1: Read current value from the register
        before_write = self.read_register(register, length)
        if before_write is None:
            logger.error(f"Failed to read from register {hex(register)} before writing.")
            return

        # Step 2: Perform write operation
        self.basic_write(register, data)

        # Step 3: Read back the value after write
        after_write = self.read_register(register, length)
        if after_write is None:
            logger.error(f"Failed to read from register {hex(register)} after writing.")
            return

        # Step 4: Compare values
        if before_write == after_write:
            logger.warning(f"[WARNING] Write did not change the register {hex(register)} value!")
        else:
            logger.info(f"Write successful! Register {hex(register)} changed from {before_write} to {after_write}")

    def execute_batch_file(self, filename):
        """Execute I2C batch commands from an XML file."""
        if self.handle is None:
            logger.error("Aardvark device is not open.")
            raise RuntimeError("Aardvark device is not open.")

        if not os.path.exists(filename):
            logger.error(f"Batch file '{filename}' not found.")
            return

        try:
            # Parse XML file
            tree = ET.parse(filename)
            root = tree.getroot()

            for command in root.findall("i2c_write"):
                addr = int(command.get("addr"), 16)  # Convert device address to int
                count = int(command.get("count"))  # Number of bytes in the command
                radix = int(command.get("radix"))  # Radix (base of the numbers, e.g., base 16 for hex)

                # Extract data as a list of integers
                data_list = [int(byte, radix) for byte in command.text.strip().split()]

                # Check if we have enough bytes (first byte is register address, rest are values)
                if len(data_list) < 2:
                    logger.error(f"Invalid data format in batch file: {data_list}")
                    continue

                # First byte is the register start address
                start_register = data_list[0]
                values = data_list[1:]  # Remaining bytes are data to write

                # Write each value to consecutive registers
                for i, value in enumerate(values):
                    current_register = start_register + i  # Increment register address
                    self.basic_write(current_register, [value])
                    # logger.warning(f"Wrote {hex(value)} to register {hex(current_register)}")

        except ET.ParseError:
            logger.error(f"Error parsing XML file '{filename}'.")
        except Exception as e:
            logger.error(f"Unexpected error processing batch file: {str(e)}")

    def execute_MTP_batch_file(self, filename):
        """Execute I2C batch commands from an XML file."""
        if self.handle is None:
            logger.error("Aardvark device is not open.")
            raise RuntimeError("Aardvark device is not open.")

        if not os.path.exists(filename):
            logger.error(f"Batch file '{filename}' not found.")
            return

        try:
            # Parse XML file
            tree = ET.parse(filename)
            root = tree.getroot()

            for command in root.findall("i2c_write"):
                addr = int(command.get("addr"), 16)  # Convert device address to int
                count = int(command.get("count"))  # Number of bytes in the command
                radix = int(command.get("radix"))  # Radix (base of the numbers, e.g., base 16 for hex)

                # Extract data as a list of integers
                data_list = [int(byte, radix) for byte in command.text.strip().split()]

                # Check if we have enough bytes (first byte is register address, rest are values)
                if len(data_list) < 2:
                    logger.error(f"Invalid data format in batch file: {data_list}")
                    continue

                # First byte is the register start address
                start_register = 0x37
                values = data_list[0:]  # Remaining bytes are data to write

                # Write each value to consecutive registers
                for i, value in enumerate(values):
                    current_register = start_register + i  # Increment register address
                    self.basic_write(current_register, [value])
                    # logger.warning(f"Wrote {hex(value)} to register {hex(current_register)}")

        except ET.ParseError:
            logger.error(f"Error parsing XML file '{filename}'.")
        except Exception as e:
            logger.error(f"Unexpected error processing batch file: {str(e)}")

    def __del__(self):
        """Destructor to ensure the Aardvark device is closed properly."""
        self.close()

###############################################
# # Initialize Aardvark with I2C slave address 0x10
# aardvark = AardvarkController(i2c_address=0x10, bitrate=400)

# # Open and configure Aardvark
# aardvark.open()

# # Write to register 0x44 with values 0x33 and 0x11
# aardvark.basic_write(0x44, [0x33, 0x11]) # write value only
# aardvark.write_register(0x44, [0x11, 0x33]) # read write read sequence

# # Read 2 bytes from register 0x44
# data = aardvark.read_register(0x44, 2)

# # Execute a batch command file
# aardvark.execute_batch_file(r"C:\WORK\Swift\Bootup_Cali\redriver_I2c_program_79_part4.xml")

# # Close Aardvark connection
# aardvark.close()
