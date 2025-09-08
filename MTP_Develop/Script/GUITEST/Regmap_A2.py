# -*- coding: utf-8 -*-
"""
Created on Tue Jan 14 10:04:01 2025

@author: CC.Cheng
"""

import sqlite3
import pandas as pd
import shutil
import os
import binascii
import xml.etree.ElementTree as ET
from datetime import datetime

#TRACKING TAB
CRC_TRAC_RANGE_START = 0x7C00
CRC_TRAC_RANGE_END = 0x7C09
#DUSR TAB
CRC_DUSR_ID_RANGE_START = 0x7C10
CRC_DUSR_ID_RANGE_END = 0x7C1D
CRC_DUSR_RANGE_START = 0x7C10
CRC_DUSR_RANGE_END = 0x7D49
#CALIBRATION TAB
CRC_CALI_RANGE_START = 0x7D4C
CRC_CALI_RANGE_END = 0x7D5D
#CONFIGURATION TAB
CRC_CONF_RANGE_START = 0x7D64
CRC_CONF_RANGE_END = 0x7DE9


class DatabaseController:
    def __init__(self, filepath):
        self.original_path = filepath
        self.modified_path = self._create_copy()
        
        # Connect to the database
        self.connection = sqlite3.connect(self.modified_path)
        
        # Load registers
        self.dataframe = self._load_map_tab("Configuration")
        self.dusr_dataframe = self._load_map_tab("DUSR")
        self.calibration_dataframe = self._load_map_tab("Calibration")
        self.trac_dataframe = self._load_map_tab("Tracking Data")
        self.prod_dataframe = self._load_map_tab("ACBIN Data")
        self.ovst_dataframe = self._load_map_tab("OVST Flag")
    
        # Load descriptions
        # self.description_dataframe = self._load_description_tab() 
        # self.dusr_description_dataframe = self._load_dusr_description_tab()
        # self.calibration_description_dataframe = self._load_calibration_description_tab()
        # self.trac_description_dataframe = self._load_trac_description_tab()
        # self.prod_description_dataframe = self._load_prod_description_tab()
        # self.ovst_description_dataframe = self._load_ovst_description_tab()
 
    def _create_copy(self):
        """Create a timestamped copy of the original Excel file for manipulation."""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        copy_path = os.path.join(os.path.dirname(self.original_path), f"SiriusA2_MTPRegisterSpecification_{timestamp}.db")
        shutil.copy(self.original_path, copy_path)
        return copy_path
        
 
    def _query_section(self, section_name):
        """Query data for a specific section from the database."""
        query = f"""
        SELECT * FROM SiriusA2_MTPRegisterSpecification
        WHERE section = '{section_name}';
        """
        return pd.read_sql_query(query, self.connection)
    
    def _load_map_tab(self, section):
        """Load the TRACKING data."""
        df = self._query_section(section)
        df.rename(columns={
            df.columns[3]: 'MTP address',
            df.columns[5]: 'Field width',
            df.columns[4]: 'Record Field',
            df.columns[8]: 'Value',
            df.columns[6]: 'Sub Field Name',
            df.columns[9]: 'Description'
        }, inplace=True)
        df.sort_values(by='MTP address', inplace=True)
        return df


    def compute_crc16(self, data: bytes, poly: int = 0x1021, init: int = 0x1021) -> int:
        """
        Compute CRC16 based on bit-by-bit processing logic.
        """
        crc = init
        
        for byte in data:
            for _ in range(8):  # Process each bit in the byte
                msb = (crc >> 15) & 1
                input_bit = (byte >> 7) & 1
                crc = (crc << 1) & 0xFFFF  # Shift CRC and keep it 16-bit
                if msb ^ input_bit:
                    crc ^= poly
                byte = (byte << 1) & 0xFF  # Shift byte for the next bit
        return f"{crc:04X}"

    def calculate_crc_for_conf_range(self):
        """Calculate CRC for a given address range and update CRC fields."""
        def sanitize_address(x):
            try:
                return f"{int(x, 16):04X}"
            except (ValueError, TypeError):
                return None

        self.dataframe['MTP address'] = self.dataframe['MTP address'].apply(sanitize_address)
        filtered_data = self.dataframe[
            self.dataframe['MTP address'].apply(lambda x: x is not None and CRC_CONF_RANGE_START <= int(x, 16) <= CRC_CONF_RANGE_END)
        ]

        # Collect values for CRC calculation
        values = []
        for value in filtered_data['Value']:
            if isinstance(value, str) and all(c in "0123456789ABCDEF" for c in value.upper()):
                values.append(value.upper())

        crc_data = values
        if not crc_data:
            raise ValueError("No valid data found for CRC calculation.")

        crc_data = ''.join(crc_data)
        # Convert the hex string to bytes
        crc_data = bytes.fromhex(crc_data)
        # Calculate CRC16      
        crc_value = self.compute_crc16(crc_data)

        # Update CRC fields accordingly
        self.dataframe.loc[self.dataframe['MTP address'] == '7DEA', 'Value'] = crc_value[:2]
        self.dataframe.loc[self.dataframe['MTP address'] == '7DEB', 'Value'] = crc_value[-2:]
        self.save_changes()
        
    def calculate_crc_for_DUSR_range(self):
        """Calculate CRC for a given address range and update CRC fields."""
        def sanitize_address(x):
            try:
                return f"{int(x, 16):04X}"
            except (ValueError, TypeError):
                return None

        self.dusr_dataframe['MTP address'] = self.dusr_dataframe['MTP address'].apply(sanitize_address)
        filtered_data = self.dusr_dataframe[
            self.dusr_dataframe['MTP address'].apply(lambda x: x is not None and CRC_DUSR_ID_RANGE_START <= int(x, 16) <= CRC_DUSR_ID_RANGE_END)
        ]
        
        # Collect values for CRC calculation
        values = []
        for value in filtered_data['Value']:
            if isinstance(value, str) and all(c in "0123456789ABCDEF" for c in value.upper()):
                values.append(value.upper())

        crc_data = values
        if not crc_data:
            raise ValueError("No valid data found for CRC calculation.")

        crc_data = ''.join(crc_data)
        crc_data = bytes.fromhex(crc_data)
        crc_value = self.compute_crc16(crc_data)
        # Update CRC fields accordingly
        self.dusr_dataframe.loc[self.dusr_dataframe['MTP address'] == '7C1E', 'Value'] = crc_value
        self.save_changes()

        self.dusr_dataframe['MTP address'] = self.dusr_dataframe['MTP address'].apply(sanitize_address)
        filtered_data = self.dusr_dataframe[
            self.dusr_dataframe['MTP address'].apply(lambda x: x is not None and CRC_DUSR_RANGE_START <= int(x, 16) <= CRC_DUSR_RANGE_END)
        ]

        # Collect values for CRC calculation
        values = []
        for value in filtered_data['Value']:
            if isinstance(value, str) and all(c in "0123456789ABCDEF" for c in value.upper()):
                values.append(value.upper())

        crc_data = values
        if not crc_data:
            raise ValueError("No valid data found for CRC calculation.")

        crc_data = ''.join(crc_data)
        crc_data = bytes.fromhex(crc_data)
        # Calculate CRC16      
        crc_value = self.compute_crc16(crc_data)
        # Update CRC fields accordingly
        self.dusr_dataframe.loc[self.dusr_dataframe['MTP address'] == '7D4A', 'Value'] = crc_value
        self.save_changes()
    
    def calculate_crc_for_cali_range(self):
        """Calculate CRC for a given address range and update CRC fields."""
        def sanitize_address(x):
            try:
                return f"{int(x, 16):04X}"
            except (ValueError, TypeError):
                return None

        self.calibration_dataframe['MTP address'] = self.calibration_dataframe['MTP address'].apply(sanitize_address)
        filtered_data = self.calibration_dataframe[
            self.calibration_dataframe['MTP address'].apply(lambda x: x is not None and CRC_CALI_RANGE_START <= int(x, 16) <= CRC_CALI_RANGE_END)
        ]

        # Collect values for CRC calculation
        values = []
        for value in filtered_data['Value']:
            if isinstance(value, str) and all(c in "0123456789ABCDEF" for c in value.upper()):
                values.append(value.upper())

        crc_data = values
        if not crc_data:
            raise ValueError("No valid data found for CRC calculation.")

        crc_data = ''.join(crc_data)
        crc_data = bytes.fromhex(crc_data)
        # Calculate CRC16      
        crc_value = self.compute_crc16(crc_data)

        # Update CRC fields accordingly
        self.calibration_dataframe.loc[self.calibration_dataframe['MTP address'] == '7D5E', 'Value'] = crc_value
        self.save_changes()

    def calculate_crc_for_trac_range(self):
        """Calculate CRC for a given address range and update CRC fields."""
        def sanitize_address(x):
            try:
                return f"{int(x, 16):04X}"
            except (ValueError, TypeError):
                return None

        self.trac_dataframe['MTP address'] = self.trac_dataframe['MTP address'].apply(sanitize_address)
        filtered_data = self.trac_dataframe[
            self.trac_dataframe['MTP address'].apply(lambda x: x is not None and CRC_TRAC_RANGE_START <= int(x, 16) <= CRC_TRAC_RANGE_END)
        ]

        # Collect values for CRC calculation
        values = []
        for value in filtered_data['Value']:
            if isinstance(value, str) and all(c in "0123456789ABCDEF" for c in value.upper()):
                values.append(value.upper())

        crc_data = values
        if not crc_data:
            raise ValueError("No valid data found for CRC calculation.")

        crc_data = ''.join(crc_data)
        # Convert the hex string to bytes
        crc_data = bytes.fromhex(crc_data)
        # Calculate CRC16      
        crc_value = self.compute_crc16(crc_data)

        # Update CRC fields accordingly
        self.trac_dataframe.loc[self.trac_dataframe['MTP address'] == '7C0A', 'Value'] = crc_value[:2] +  crc_value[-2:]
        self.save_changes()

    def clear_crc(self, tab):
        """Clear the CRC fields."""
        if tab == "CONFIGURATION":
            self.dataframe.loc[self.dataframe['MTP address'] == '7DEA', 'Value'] = None
            self.dataframe.loc[self.dataframe['MTP address'] == '7DEB', 'Value'] = None
            self.save_changes()
        if tab == "DUSR":
            self.dusr_dataframe.loc[self.dusr_dataframe['MTP address'] == '7C1E', 'Value'] = None
            self.dusr_dataframe.loc[self.dusr_dataframe['MTP address'] == '7D4A', 'Value'] = None
            self.save_changes()
        if tab == "CALIBRATION":
            self.calibration_dataframe.loc[self.calibration_dataframe['MTP address'] == '7D5E', 'Value'] = None
            self.save_changes()
        if tab == "Tracking":
            self.trac_dataframe.loc[self.trac_dataframe['MTP address'] == '7C0A', 'Value'] = None
            self.save_changes()

    def save_changes(self):
        """Save changes to all relevant Excel sheets."""
        with pd.ExcelWriter(self.modified_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            self.dataframe.to_excel(writer, sheet_name='MTP CONF Map', index=False)
            self.dusr_dataframe.to_excel(writer, sheet_name='DUSR Map', index=False)
            self.calibration_dataframe.to_excel(writer, sheet_name='CALIBRATION Map', index=False)
            self.trac_dataframe.to_excel(writer, sheet_name='TRACKING Map', index=False)
            self.prod_dataframe.to_excel(writer, sheet_name='PROD Map', index=False)
            self.ovst_dataframe.to_excel(writer, sheet_name='OVST Map', index=False)
            # If description sheets or other tabs are needed, write them as well.

    def update_value(self, tab_name, address, new_value):
        """Update a specific MTP address value in the appropriate tab's dataframe."""
        if tab_name == "CONFIGURATION":
            df = self.dataframe
        elif tab_name == "DUSR":
            df = self.dusr_dataframe
        elif tab_name == "CALIBRATION":
            df = self.calibration_dataframe
        elif tab_name == "Tracking":
            df = self.trac_dataframe
        elif tab_name == "Process Monitor Data":
            df = self.prod_dataframe
        elif tab_name == "OVST Flag":
            df = self.ovst_dataframe
        else:
            return  

        # Update the value for the given address
        df.loc[df['MTP address'].str.upper() == address.upper(), 'Value'] = new_value
        self.save_changes()

    def generate_binary_file_mtpconfig(self, output_path):
        """Generate a binary file from the current values."""
        hex_values = self.dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):02X}").tolist()
        # print(hex_values)
        binary_data = binascii.unhexlify(''.join(hex_values))
        # print(binary_data)
        padding = b'\x00' * 356
        with open(output_path, 'wb') as f:
            f.write(padding + binary_data)
            
    def generate_binary_file_mtpfull(self, output_path):
        """Generate a binary file from the current values."""
        trac_hex_values = self.trac_dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):X}".zfill(len(str(x)))).tolist()
        acbi_hex_values = self.prod_dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):X}".zfill(len(str(x)))).tolist()
        conf_hex_values = self.dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):02X}").tolist()
        dusr_hex_values = self.dusr_dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):X}".zfill(len(str(x)))).tolist()
        ovst_hex_values = self.ovst_dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):X}".zfill(len(str(x)))).tolist()
        cali_hex_values = self.calibration_dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):X}".zfill(len(str(x)))).tolist()
        spar_hex_values = ['00'] * 20
        
        # 0x13000 ~ 0x131FF
        first_half_hex_values = trac_hex_values + acbi_hex_values + dusr_hex_values + cali_hex_values + ovst_hex_values + conf_hex_values + spar_hex_values 
        
        # hex_values = first_half_hex_values * 2
        hex_values = first_half_hex_values + ['00'] * 512
        binary_data = binascii.unhexlify(''.join(hex_values))
        with open(output_path, 'wb') as f:
            f.write(binary_data)

    def combine_xml(self,file1_path, Update_File_Path, output_path):
        # Parse the first XML
        # print(file1_path)
        tree1 = ET.parse(file1_path)
        root1 = tree1.getroot()  # The root element of file1
    
        # Parse the second XML
        tree2 = ET.parse(Update_File_Path)
        root2 = tree2.getroot()  # The root element of file2
    
        # Append all children from file2's root into file1's root
        for child in root2:
            root1.append(child)
    
        # Save the updated tree as a new XML
        tree1.write(output_path, encoding="utf-8", xml_declaration=False)
        # print(f"Combined XML written to {output_path}")
    
    
    def close_connection(self):
        """Close the database connection."""
        self.connection.close()
        if os.path.exists(self.modified_path):
            os.remove(self.modified_path)
            

# Example usage
db_path = "C:\WORK\Sirius\MTP\SiriusA2_MTPRegisterSpecification.db"
db_controller = DatabaseController(db_path)

# Access data
# print("Tracking Data:")
# print(db_controller.trac_dataframe)

# print("DUSR Data:")
# print(db_controller.dusr_dataframe)

# Close the connection when done
db_controller.close_connection()
