# -*- coding: utf-8 -*-
"""
Created on Wed Dec 11 13:51:41 2024

@author: CC.Cheng
"""

def char_to_hex(char):
    """Convert a hexadecimal character to its integer value."""
    if '0' <= char <= '9':
        return ord(char) - ord('0')
    elif 'A' <= char <= 'F':
        return ord(char) - ord('A') + 10
    elif 'a' <= char <= 'f':
        return ord(char) - ord('a') + 10
    else:
        raise ValueError(f"Illegal character: {char}")


def calculate_crc16(data: str, poly: int = 0x1021, init: int = 0x1021):
    """Calculate CRC16 based on the logic in the provided C code."""
    crc = init
    is_high_nibble = True
    current_byte = 0

    for char in data:
        if char.isspace():
            continue  # Skip spaces or newline characters

        nibble = char_to_hex(char)

        if is_high_nibble:
            current_byte = nibble << 4
            is_high_nibble = False
        else:
            current_byte |= nibble
            is_high_nibble = True

            # Process each bit of the byte
            for _ in range(8):
                msb = (crc >> 15) & 1
                input_bit = (current_byte >> 7) & 1
                current_byte = (current_byte << 1) & 0xFF
                if msb ^ input_bit:
                    crc = ((crc << 1) ^ poly) & 0xFFFF
                else:
                    crc = (crc << 1) & 0xFFFF

    return crc


# Example usage
# input_data = "70000000F003000000000000000000000000"
input_data = "AA5500300000000000020B00000000000000CD059808940080072C00650424003804050085B60207FF0C080C006E"
# input_data = "aa5500880000000000020b00000000000000cd059808940080072c00650424003804050085b60207ff0c080c006ed205c805a2088e088a0776076f045b0442042e040000d49b740000000000001201010d0059c66000000085b654821c46419464865d5625c7708a9ea6bdaa9bb628000000000000000000000006c600008900c5818e670100"
crc16_result = calculate_crc16(input_data)
print(crc16_result)
print(f"CRC16: {crc16_result:04X}")
