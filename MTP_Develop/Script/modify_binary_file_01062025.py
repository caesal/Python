import struct
import argparse
import os

def calculate_crc16(Buf: bytes, W_len: int) -> int:
    Accum = 0x1021
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

def modify_binary_file(modetype: str, input_file: str, output_file: str, xml_file: str, slave_addr: int = 0x54, chunk_size: int = 1024, engineer: str = "customer"):
    Product_Data_size = 356
    Config_Data_size = 136
    if engineer == "indie":
        DDC2BI_header = bytes([0x51, 0x85, 0xC2, 0x00, 0x00, 0x03, 0x11, 0x6A])
    else:
        DDC2BI_header = bytes([0x51, 0x85, 0xC2, 0x00, 0x00, 0x03, 0x10, 0x6B])
        
    physical_addr = 0

   # CalCrcAndSizeCodePatch process(merge)
    if modetype == "codepatch":
        
        with open(input_file, 'rb') as f:
            original_data = f.read()
        
        total_size = get_dynamic_size(original_data)
        
        intermediate_file = "codePatchCRC.bin"  

        if len(original_data) < 32:
            raise ValueError("Input data is too short for codepatch operation")

        ConfigHeader = original_data[0:32]
        ConfigLen = original_data[4:6]

        print(" ".join(f"{b:02x}" for b in ConfigHeader))
        
        print(" ".join(f"{b:02x}" for b in ConfigLen))
        
        CalculatedData = original_data[32:total_size]
        CalculatedDataSize = struct.pack('<I', len(CalculatedData))

        print(f"codepatch size w/o ConfigHeader: {CalculatedDataSize} bytes")
        
        crc16 = calculate_crc16(CalculatedData, len(CalculatedData))
        crc16_bytes = original_data[0:2] + struct.pack('>H', crc16)

        modified_data = b''.join([
            crc16_bytes,
            ConfigLen,
            CalculatedDataSize,
            ConfigHeader[10:32],
            CalculatedData
        ])

        print(f"Crc16Vale: {hex(crc16)}")
        
        with open(intermediate_file, 'wb') as f:
            f.write(modified_data)
        input_file = intermediate_file    #Update input_file


    with open(input_file, 'rb') as f:
        original_data = f.read()
        
    total_size = get_dynamic_size(original_data)
    
    if total_size < Config_Data_size:
        raise ValueError("Incooret input file")

    if modetype == "codepatch":
        physical_addr = 0x300000
        print(f"Detected dynamic size for codepatch: {total_size} bytes")
    elif modetype == "mtpfull":
        physical_addr = 0x13000
        print(f"Total size for mtpfull: {total_size} bytes")
    elif modetype == "mtpproduct":
        physical_addr = 0x13000
        total_size = Product_Data_size
        print(f"Size of mtpproduct: {Product_Data_size} bytes")
    elif modetype == "mtpconfig":
        physical_addr = 0x13164
        total_size = Config_Data_size
        print(f"Size of mtpconfig: {Config_Data_size} bytes")
    elif modetype == "mtpconfigBin":
        physical_addr = 0x13164
        total_size = Config_Data_size
        print(f"Size of mtpconfigBin: {Config_Data_size} bytes")
    elif modetype == "eeprom":
        physical_addr = 0x164
        chunk_size = 8
        total_size = 136
        print(f"Size of configuration data: {Config_Data_size} bytes")
    else:
        raise ValueError("Invalid type. Please choose one of: codepatch, mtpfull, mtpproduct, mtpconfig, eeprom")

    record_data = bytes()
    non_modified_data = bytes()
    non_modified_file_name = modetype + ".bin"
    batch_file_name = "APPSTEST_EEPROM_Updata_script.batch"
    with open(xml_file, 'w') as xml_f:   
        xml_f.write(f'<aardvark>\n')

        if modetype in ["codepatch", "mtpfull"]:
            num_chunks = total_size // chunk_size
            remaining_size = total_size % chunk_size

            for i in range(num_chunks):          
                chunk_data = original_data[i * chunk_size:(i + 1) * chunk_size]
                
                DDC2BI_header_with_addr_size = DDC2BI_header + struct.pack('>I', physical_addr) + struct.pack('>I', chunk_size)

                crc16 = calculate_crc16(chunk_data, len(chunk_data))
                crc16_bytes = bytes([0x00, 0x00]) + struct.pack('>H', crc16)
                
                count = len(DDC2BI_header_with_addr_size) + len(chunk_data) + len(crc16_bytes)

                data_with_crc = DDC2BI_header_with_addr_size + chunk_data + crc16_bytes
                hex_data = " ".join(f"{b:02X}" for b in data_with_crc)
                xml_f.write(f'<i2c_write addr="0x37" count="{count}" nostop="0" radix="16">\n')
                xml_f.write(f'    {hex_data}\n')
                xml_f.write(f'</i2c_write>\n')
                xml_f.write(f'<sleep ms="2"/>\n')
                xml_f.write(f'<i2c_read addr="0x37" count="8"></i2c_read>\n')
                xml_f.write(f'<sleep ms="1"/>\n')
                
                record_data += DDC2BI_header_with_addr_size + chunk_data + crc16_bytes
                
                physical_addr += chunk_size

            if remaining_size > 0:
                chunk_data = original_data[num_chunks * chunk_size:num_chunks * chunk_size + remaining_size]
                
                if len(chunk_data) % 2 != 0:
                    remaining_size = remaining_size + 1
                    chunk_data += bytes([0x00])
                    
                DDC2BI_header_with_addr_size = DDC2BI_header + struct.pack('>I', physical_addr) + struct.pack('>I', remaining_size)

                crc16 = calculate_crc16(chunk_data, len(chunk_data))
                crc16_bytes = bytes([0x00, 0x00]) + struct.pack('>H', crc16)

                count = len(DDC2BI_header_with_addr_size) + len(chunk_data) + len(crc16_bytes)

                data_with_crc = DDC2BI_header_with_addr_size + chunk_data + crc16_bytes
                hex_data = " ".join(f"{b:02X}" for b in data_with_crc)
                xml_f.write(f'<i2c_write addr="0x37" count="{count}" nostop="0" radix="16">\n')
                xml_f.write(f'    {hex_data}\n')
                xml_f.write(f'</i2c_write>\n')
                xml_f.write(f'<sleep ms="2"/>\n')
                xml_f.write(f'<i2c_read addr="0x37" count="8"></i2c_read>\n')
                xml_f.write(f'<sleep ms="1"/>\n')
                
                record_data += DDC2BI_header_with_addr_size + chunk_data + crc16_bytes
                
            xml_f.write(f'</aardvark>\n')
            print(f"XML has been created as {xml_file}")

        elif modetype in ["mtpproduct", "mtpconfig"]:
            offset = physical_addr - 0x13000
            num_chunks = total_size // chunk_size
            remaining_size = total_size % chunk_size

            for i in range(num_chunks):
                chunk_data = original_data[offset + (i * chunk_size) : offset + ((i + 1) * chunk_size)]

                DDC2BI_header_with_addr_size = DDC2BI_header + struct.pack('>I', physical_addr) + struct.pack('>I', chunk_size)

                crc16 = calculate_crc16(chunk_data, len(chunk_data))
                crc16_bytes = bytes([0x00, 0x00]) + struct.pack('>H', crc16)

                count = len(DDC2BI_header_with_addr_size) + len(chunk_data) + len(crc16_bytes)

                data_with_crc = DDC2BI_header_with_addr_size + chunk_data + crc16_bytes
                hex_data = " ".join(f"{b:02X}" for b in data_with_crc)
                xml_f.write(f'<i2c_write addr="0x37" count="{count}" nostop="0" radix="16">\n')
                xml_f.write(f'    {hex_data}\n')
                xml_f.write(f'</i2c_write>\n')
                xml_f.write(f'<sleep ms="2"/>\n')
                xml_f.write(f'<i2c_read addr="0x37" count="8"></i2c_read>\n')
                xml_f.write(f'<sleep ms="1"/>\n')
                
                record_data += DDC2BI_header_with_addr_size + chunk_data + crc16_bytes
                non_modified_data += chunk_data
            
                physical_addr += chunk_size
            
            if remaining_size > 0:
                chunk_data = original_data[offset + (num_chunks * chunk_size):offset + (num_chunks * chunk_size) + remaining_size]
                
                if len(chunk_data) % 2 != 0:
                    remaining_size = remaining_size + 1
                    chunk_data += bytes([0x00])
                    
                DDC2BI_header_with_addr_size = DDC2BI_header + struct.pack('>I', physical_addr) + struct.pack('>I', remaining_size)
                
                crc16 = calculate_crc16(chunk_data, len(chunk_data))
                crc16_bytes = bytes([0x00, 0x00]) + struct.pack('>H', crc16)
                
                count = len(DDC2BI_header_with_addr_size) + len(chunk_data) + len(crc16_bytes)
                
                data_with_crc = DDC2BI_header_with_addr_size + chunk_data + crc16_bytes
                hex_data = " ".join(f"{b:02X}" for b in data_with_crc)
                xml_f.write(f'<i2c_write addr="0x37" count="{count}" nostop="0" radix="16">\n')
                xml_f.write(f'    {hex_data}\n')
                xml_f.write(f'</i2c_write>\n')
                xml_f.write(f'<sleep ms="2"/>\n')
                xml_f.write(f'<i2c_read addr="0x37" count="8"></i2c_read>\n')
                xml_f.write(f'<sleep ms="1"/>\n')
                
                record_data += DDC2BI_header_with_addr_size + chunk_data + crc16_bytes
                non_modified_data += chunk_data
                
            xml_f.write(f'</aardvark>\n')
            print(f"XML has been created as {xml_file}")
            
            with open(non_modified_file_name, 'wb') as non_m_f:
                non_m_f.write(non_modified_data)
                
        elif modetype in ["mtpconfigBin"]:
            offset = physical_addr - 0x13164
            num_chunks = total_size // chunk_size
            remaining_size = total_size % chunk_size

            for i in range(num_chunks):
                chunk_data = original_data[offset + (i * chunk_size) : offset + ((i + 1) * chunk_size)]

                DDC2BI_header_with_addr_size = DDC2BI_header + struct.pack('>I', physical_addr) + struct.pack('>I', chunk_size)

                crc16 = calculate_crc16(chunk_data, len(chunk_data))
                crc16_bytes = bytes([0x00, 0x00]) + struct.pack('>H', crc16)

                count = len(DDC2BI_header_with_addr_size) + len(chunk_data) + len(crc16_bytes)

                data_with_crc = DDC2BI_header_with_addr_size + chunk_data + crc16_bytes
                hex_data = " ".join(f"{b:02X}" for b in data_with_crc)
                xml_f.write(f'<i2c_write addr="0x37" count="{count}" nostop="0" radix="16">\n')
                xml_f.write(f'    {hex_data}\n')
                xml_f.write(f'</i2c_write>\n')
                xml_f.write(f'<sleep ms="2"/>\n')
                xml_f.write(f'<i2c_read addr="0x37" count="8"></i2c_read>\n')
                xml_f.write(f'<sleep ms="1"/>\n')

                record_data += DDC2BI_header_with_addr_size + chunk_data + crc16_bytes
                non_modified_data += chunk_data

                physical_addr += chunk_size

            if remaining_size > 0:
                chunk_data = original_data[offset + (num_chunks * chunk_size):offset + (num_chunks * chunk_size) + remaining_size]

                if len(chunk_data) % 2 != 0:
                    remaining_size = remaining_size + 1
                    chunk_data += bytes([0x00])

                DDC2BI_header_with_addr_size = DDC2BI_header + struct.pack('>I', physical_addr) + struct.pack('>I', remaining_size)

                crc16 = calculate_crc16(chunk_data, len(chunk_data))
                crc16_bytes = bytes([0x00, 0x00]) + struct.pack('>H', crc16)

                count = len(DDC2BI_header_with_addr_size) + len(chunk_data) + len(crc16_bytes)

                data_with_crc = DDC2BI_header_with_addr_size + chunk_data + crc16_bytes
                hex_data = " ".join(f"{b:02X}" for b in data_with_crc)
                xml_f.write(f'<i2c_write addr="0x37" count="{count}" nostop="0" radix="16">\n')
                xml_f.write(f'    {hex_data}\n')
                xml_f.write(f'</i2c_write>\n')
                xml_f.write(f'<sleep ms="2"/>\n')
                xml_f.write(f'<i2c_read addr="0x37" count="8"></i2c_read>\n')
                xml_f.write(f'<sleep ms="1"/>\n')

                record_data += DDC2BI_header_with_addr_size + chunk_data + crc16_bytes
                non_modified_data += chunk_data

            xml_f.write(f'</aardvark>\n')
            print(f"XML has been created as {xml_file}")

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
                    xml_f.write(f'<sleep ms="1"/>\n')
                    xml_f.write(f'</i2c_write>\n')
                    xml_f.write(f'<sleep ms="1"/>\n')
                    
                    batch_data = " ".join(f"0x{byte:02X}" for byte in data_chunk)
                    batch_f.write(f"AppsTest 31 0x{regaddr:02X} 8 {batch_data}\n")
                    
                    record_data += struct.pack('>B', regaddr) + config_data[i * chunk_size:(i + 1) * chunk_size]                
                    regaddr = regaddr + chunk_size
            
            xml_f.write(f'</aardvark>\n')
            print(f"XML has been created as {xml_file}")
            print(f"Batch file has been created as {batch_file_name}")

    with open(output_file, 'wb') as f:
        f.write(record_data)

    print(f"File has been modified and saved to {output_file}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Modify a binary file by adding dynamic byte array and CRC16 checksum based on type.')

    parser.add_argument('params', nargs='+', help='Input parameters in key=value format')
    
    args = parser.parse_args()

    param_dict = {}
    for param in args.params:
        key, value = param.split('=')
        param_dict[key.lower()] = value


    modetype = param_dict.get('type', None)
    input_file = param_dict.get('input', None)
    if input_file is None:
        raise ValueError("Missing required parameter: input")
    
    name = input_file.rsplit('.bin', 1)[0]
    
    output_file = param_dict.get('output', f"{modetype}{name}_converted.bin")
    xml_file = param_dict.get('xml', f"{modetype}{name}_batch.xml")
    slave_addr = int(param_dict.get('slaveaddr', '0x54'), 16)
    chunk_size = int(param_dict.get('payload', '1024'))
    engineer = param_dict.get('eng', 'costomer')

    if not all([modetype, input_file]):
        raise ValueError("Missing required parameter: type")

    modify_binary_file(modetype, input_file, output_file, xml_file, slave_addr, chunk_size, engineer)
