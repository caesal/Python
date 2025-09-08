import struct
import argparse
import os

def calculate_crc16(Accum: int, Buf: bytes, W_len: int) -> int:
    PolDiv = 0x1021
    for j in range(W_len):
        DataIn = Buf[j]
        for _ in range(8):
            flag = DataIn ^ (Accum >> 8)
            Accum <<= 1

            if flag & 0x80:
                Accum ^= PolDiv

            DataIn <<= 1
            Accum &= 0xFFFF
    return Accum

def get_dynamic_size(data: bytes) -> int:
    for i in range(len(data) - 1, -1, -1):
        if data[i] != 0:
            return i + 1
    return 0

def modify_binary_file(modetype: str, input_file: str, output_file: str, xml_file: str, slave_addr: int = 0x54):
    fixed_start = bytes([0x51, 0x85, 0xC2, 0x00, 0x00, 0x03, 0x10, 0x6B])
    physical_addr = 0
    chunk_size = 1024

    with open(input_file, 'rb') as f:
        original_data = f.read()

    if modetype == "codepatch":
        physical_addr = 0x300000
        total_size = get_dynamic_size(original_data)
        print(f"Detected dynamic size for codepatch: {total_size} bytes")
    elif modetype == "mtpfull":
        physical_addr = 0x13000
        total_size = len(original_data)
        print(f"Total size for mtpfull: {total_size} bytes")
    elif modetype == "mtpproduct":
        physical_addr = 0x13000
        chunk_size = 0x164
        total_size = len(original_data)
        print(f"Total size for mtpproduct: {total_size} bytes")
    elif modetype == "mtpconfig":
        physical_addr = 0x13164
        chunk_size = 0x7E
        total_size = len(original_data)
        print(f"Total size for mtpconfig: {total_size} bytes")
    elif modetype == "eeprom":
        physical_addr = 0x164
        chunk_size = 8
        total_size = 0x80
        print(f"Total size for eeprom: {total_size} bytes")
    else:
        raise ValueError("Invalid type. Please choose one of: codepatch, mtpfull, mtpproduct, mtpconfig, eeprom, eepapp")

    modified_data = bytes()
    non_modified_data = bytes()
    non_modified_file_name = modetype + ".bin"
    batch_file_name = "APPSTEST_EEPROM_Updata_script.batch"
    with open(xml_file, 'w') as xml_f:   
        xml_f.write(f'<aardvark>\n')

        if modetype in ["codepatch", "mtpfull"]:
            num_chunks = total_size // chunk_size
            remaining_size = total_size % chunk_size

            for i in range(num_chunks):
                current_chunk_size = chunk_size
                
                chunk_data = original_data[i * chunk_size:(i + 1) * chunk_size]
                    
                payload_size = struct.pack('>I', current_chunk_size)

                byte_array = fixed_start + struct.pack('>I', physical_addr) + payload_size

                Accum = 0x1021
                crc16 = calculate_crc16(Accum, chunk_data, len(chunk_data))
                crc16_bytes = struct.pack('>H', crc16)
                full_crc16_bytes = bytes([0x00, 0x00]) + crc16_bytes

                modified_data += byte_array + chunk_data + full_crc16_bytes

                count = len(byte_array) + len(chunk_data) + len(full_crc16_bytes)

                data_with_crc = byte_array + chunk_data + full_crc16_bytes
                hex_data = " ".join(f"{b:02X}" for b in data_with_crc)
                
                xml_f.write(f'<i2c_write addr="0x37" count="{count}" nostop="0" radix="16">\n')
                xml_f.write(f'    {hex_data}\n')
                xml_f.write(f'</i2c_write>\n')
                xml_f.write(f'<sleep ms="1"/>\n')
                xml_f.write(f'<i2c_read addr="0x37" count="8"></i2c_read>\n')
                xml_f.write(f'<sleep ms="1"/>\n')
                
                physical_addr += chunk_size

            if remaining_size > 0:
                current_chunk_size = remaining_size
                
                chunk_data = original_data[num_chunks * chunk_size:num_chunks * chunk_size + current_chunk_size]
                
                if len(chunk_data) % 2 != 0:
                    current_chunk_size = current_chunk_size + 1
                    chunk_data += bytes([0x00])
                    
                payload_size = struct.pack('>I', current_chunk_size)

                byte_array = fixed_start + struct.pack('>I', physical_addr) + payload_size

                Accum = 0x1021
                crc16 = calculate_crc16(Accum, chunk_data, len(chunk_data))
                crc16_bytes = struct.pack('>H', crc16)
                full_crc16_bytes = bytes([0x00, 0x00]) + crc16_bytes

                modified_data += byte_array + chunk_data + full_crc16_bytes

                count = len(byte_array) + len(chunk_data) + len(full_crc16_bytes)

                data_with_crc = byte_array + chunk_data + full_crc16_bytes
                hex_data = " ".join(f"{b:02X}" for b in data_with_crc)

                xml_f.write(f'<i2c_write addr="0x37" count="{count}" nostop="0" radix="16">\n')
                xml_f.write(f'    {hex_data}\n')
                xml_f.write(f'</i2c_write>\n')
                xml_f.write(f'<sleep ms="1"/>\n')
                xml_f.write(f'<i2c_read addr="0x37" count="8"></i2c_read>\n')
                xml_f.write(f'<sleep ms="1"/>\n')
                
            xml_f.write(f'</aardvark>\n')
            print(f"XML has been created as {xml_file}")

        elif modetype in ["mtpproduct", "mtpconfig"]:
            offset = physical_addr - 0x13000

            current_chunk_size = chunk_size
            payload_size = struct.pack('>I', current_chunk_size)

            byte_array = fixed_start + struct.pack('>I', physical_addr) + payload_size
            chunk_data = original_data[offset : offset + chunk_size]

            Accum = 0x1021
            crc16 = calculate_crc16(Accum, chunk_data, len(chunk_data))
            crc16_bytes = struct.pack('>H', crc16)
            full_crc16_bytes = bytes([0x00, 0x00]) + crc16_bytes

            modified_data += byte_array + chunk_data + full_crc16_bytes

            count = len(byte_array) + len(chunk_data) + len(full_crc16_bytes)

            data_with_crc = byte_array + chunk_data + full_crc16_bytes
            hex_data = " ".join(f"{b:02X}" for b in data_with_crc)
            
            xml_f.write(f'<i2c_write addr="0x37" count="{count}" nostop="0" radix="16">\n')
            xml_f.write(f'    {hex_data}\n')
            xml_f.write(f'</i2c_write>\n')
            xml_f.write(f'<sleep ms="1"/>\n')
            xml_f.write(f'<i2c_read addr="0x37" count="8"></i2c_read>\n')
            xml_f.write(f'<sleep ms="1"/>\n')
            
            physical_addr += chunk_size

            xml_f.write(f'</aardvark>\n')
            print(f"XML has been created as {xml_file}")
            
            non_modified_data = chunk_data
            with open(non_modified_file_name, 'wb') as non_m_f:
                non_m_f.write(non_modified_data)
                
        elif modetype in ["eeprom"]:
            config_data = original_data[physical_addr : physical_addr + total_size]
            regaddr = 0x0
            
            non_modified_data = config_data
            with open("eepconfig.bin", 'wb') as non_m_f:
                non_m_f.write(non_modified_data)
                
            with open(batch_file_name, 'w') as batch_f: 
                for i in range(total_size // chunk_size):
                    data_chunk = config_data[i * chunk_size:(i + 1) * chunk_size]
                    data_with_addr = struct.pack('>B', regaddr) + config_data[i * chunk_size:(i + 1) * chunk_size]
                    hex_data = " ".join(f"{b:02X}" for b in data_with_addr)
                    
                    xml_f.write(f'<i2c_write addr="{hex(slave_addr)}" count="{chunk_size+1}" nostop="0" radix="16">\n')
                    xml_f.write(f'    {hex_data}\n')
                    xml_f.write(f'</i2c_write>\n')
                    xml_f.write(f'<sleep ms="1"/>\n')
                    
                    batch_data = " ".join(f"0x{byte:02X}" for byte in data_chunk)
                    batch_f.write(f"AppsTest 31 0x{regaddr:02X} 8 {batch_data}\n")
                    
                    modified_data += struct.pack('>B', regaddr) + config_data[i * chunk_size:(i + 1) * chunk_size]                
                    regaddr = regaddr + chunk_size
            
            xml_f.write(f'</aardvark>\n')
            print(f"XML has been created as {xml_file}")
            print(f"Batch file has been created as {batch_file_name}")

    with open(output_file, 'wb') as f:
        f.write(modified_data)

    print(f"File has been modified and saved to {output_file}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Modify a binary file by adding dynamic byte array and CRC16 checksum based on type.')
    
    parser.add_argument('modetype', type=str, choices=['codepatch', 'mtpfull', 'mtpproduct', 'mtpconfig', 'eeprom'],
                        help='Choose the type: codepatch, mtpfull, mtpproduct, mtpconfig, eeprom')
    
    parser.add_argument('input_file', type=str, help='Relative path of the input binary file')
    
    parser.add_argument('output_file', type=str, help='Relative path of the output binary file')
    
    parser.add_argument('xml_file', type=str, help='Relative path of the XML file to be generated')
    
    parser.add_argument('slave_addr', nargs='?', type=lambda x: int(x, 16), default=0x54, 
                        help='Slave address for EEPROM mode (default is 0x54)')
    
    args = parser.parse_args()

    modify_binary_file(args.modetype, args.input_file, args.output_file, args.xml_file, args.slave_addr)
