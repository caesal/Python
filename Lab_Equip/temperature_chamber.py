# -*- coding: utf-8 -*-
"""
Created on Thu Oct 17 14:15:51 2024

@author: Caesal Cheng
"""

import time
import serial
from serial import Serial
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


class TemperatureChamber(Serial):
    def __init__(self, com=3):
        super(TemperatureChamber, self).__init__(
            port=f'COM{com}',
            baudrate=9600,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=1,
            bytesize=serial.EIGHTBITS
        )
        self.tolerance = 0.6  # Change tolerance here
        self.duration = 15  # Duration to hold temperature within tolerance

    def calculate_checksum(self, command_body):
        """Calculate the checksum for a given command body."""
        ascii_values = [ord(char) for char in command_body]
        total_sum = sum(ascii_values)
        hex_sum = format(total_sum, 'X')
        return hex_sum[-2:].zfill(2)

    def set_temperature_set_point(self, temperature, address='01'):
        """Send command to set temperature set point."""
        try:
            sign = '00' if temperature >= 0 else 'FF'
            temperature_decimal = f"{int(abs(temperature) * 10):04}"
            command_body = f'{address}0200{temperature_decimal}{sign}'
            checksum = self.calculate_checksum(command_body)
            command = f'\x02L{command_body}{checksum}\x03'

            logging.info(f"Sending command to set temperature: {repr(command)}")
            self.write(command.encode())
            response = self.read_until(b'\x06').decode()

            if response == '\x02L01000D\x06':
                logging.info("Set point updated successfully")
            else:
                logging.warning("Failed to update set point")
        except Exception as e:
            logging.error(f"Error setting temperature: {e}")

    def read_temperature(self, address='01'):
        """Read the current temperature from the device."""
        try:
            command_body = f'{address}00'
            checksum = self.calculate_checksum(command_body)
            command = f'\x02L{command_body}{checksum}\x03'

            self.write(command.encode())
            response = self.read_until(b'\x06').decode()

            if response.startswith('\x02L') and response.endswith('\x06'):
                sign = response[7:8]
                temperature = response[8:12]

                try:
                    temperature = int(temperature) / 10
                    if sign == '5':
                        temperature = -temperature
                    return temperature
                except ValueError:
                    logging.error("Error parsing temperature value")
                    return None
            else:
                logging.warning("Invalid response received")
                return None
        except Exception as e:
            logging.error(f"Error reading temperature: {e}")
            return None

    def monitor_temperature(self, target_temperature, address='01'):
        """Monitor the temperature and ensure it remains within tolerance for the specified duration."""
        consecutive_count = 0
        while consecutive_count < self.duration:
            current_temperature = self.read_temperature(address)
            if current_temperature is not None:
                logging.info(f"Current temperature: {current_temperature}°C")
                if abs(current_temperature - target_temperature) <= self.tolerance:
                    consecutive_count += 1
                else:
                    consecutive_count = 0  # Reset count if out of tolerance
            else:
                logging.warning("Skipping iteration due to read error")
            time.sleep(1)

        logging.info("Done: Temperature has stabilized within tolerance")

    def Set_Mon_temperature(self, target_temperature):
        try:
            self.set_temperature_set_point(target_temperature)
            self.monitor_temperature(target_temperature)
        except serial.SerialException as e:
            logging.error(f"Serial connection error: {e}")

# target_temperature = -40  # Change target temperature here
# chamber = TemperatureChamber(3)
# chamber.Set_Mon_temperature(target_temperature)
