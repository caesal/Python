# -*- coding: utf-8 -*-
"""
Created on Mon Jan 20 16:36:44 2025

@author: CC.Cheng

Application Name: MTP Alchemist
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import shutil
import os
import sqlite3
import subprocess
from datetime import datetime
import xml.etree.ElementTree as ET
import binascii

# Constants
GUI_VERSION = 'MTP Alchemist A1 Ver0.1.02'
GUI_SIZE = '1920x900'
# GUI_SIZE = '1120x1260'
GUI_THEME = 'clam'  # e.g., 'winnative', 'clam', 'alt', 'default', 'classic', 'vista', 'xpnative'

Update_File_Path = r"MTPDATA_update_execute.xml"
SQL_FILE = r"SiriusA1_MTPRegisterSpecification.db"

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
CRC_CALI_RANGE_END = 0x7D61
#CONFIGURATION TAB
# CRC_ADDRESSES = ('7DEA', '7DEB')
# CRC_CONF_HEADER = ['AA', '55', '00']
CRC_CONF_RANGE1_START = 0x7D64
CRC_CONF_RANGE1_END = 0x7D91
CRC_CONF_RANGE2_START = 0x7D94
CRC_CONF_RANGE2_END = 0x7DAD
CRC_CONF_RANGE3_START = 0x7DB0
CRC_CONF_RANGE3_END = 0x7DDF


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
        # self.ovst_dataframe = self._load_map_tab("OVST Flag")

    def _create_copy(self):
        """Create a timestamped copy of the original Excel file for manipulation."""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        copy_path = os.path.join(os.path.dirname(self.original_path), f"SiriusA1_MTPRegisterSpecification_{timestamp}.db")
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
            df.columns[4]: 'Record Field',
            df.columns[5]: 'Field width',
            df.columns[6]: 'Sub Field Name',
            df.columns[8]: 'Value',
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
            self.dataframe['MTP address'].apply(lambda x: x is not None and CRC_CONF_RANGE1_START <= int(x, 16) <= CRC_CONF_RANGE1_END)
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
        self.dataframe.loc[self.dataframe['MTP address'] == '7D92', 'Value'] = crc_value[-2:]
        self.dataframe.loc[self.dataframe['MTP address'] == '7D93', 'Value'] = crc_value[:2]
        self.save_changes()
        
        self.dataframe['MTP address'] = self.dataframe['MTP address'].apply(sanitize_address)
        filtered_data = self.dataframe[
            self.dataframe['MTP address'].apply(lambda x: x is not None and CRC_CONF_RANGE2_START <= int(x, 16) <= CRC_CONF_RANGE2_END)
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
        self.dataframe.loc[self.dataframe['MTP address'] == '7DAE', 'Value'] = crc_value[-2:]
        self.dataframe.loc[self.dataframe['MTP address'] == '7DAF', 'Value'] = crc_value[:2]
        self.save_changes()
        
        self.dataframe['MTP address'] = self.dataframe['MTP address'].apply(sanitize_address)
        filtered_data = self.dataframe[
            self.dataframe['MTP address'].apply(lambda x: x is not None and CRC_CONF_RANGE3_START <= int(x, 16) <= CRC_CONF_RANGE3_END)
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
        self.dataframe.loc[self.dataframe['MTP address'] == '7DE0', 'Value'] = crc_value[-2:]
        self.dataframe.loc[self.dataframe['MTP address'] == '7DE1', 'Value'] = crc_value[:2]
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
        # print(crc_data)
        if not crc_data:
            raise ValueError("No valid data found for CRC calculation.")

        crc_data = ''.join(crc_data)
        # Convert the hex string to bytes
        crc_data = bytes.fromhex(crc_data)
        # Calculate CRC16      
        crc_value = self.compute_crc16(crc_data)
        # Update CRC fields accordingly
        self.dusr_dataframe.loc[self.dusr_dataframe['MTP address'] == '7C1E', 'Value'] = crc_value[-2:] + crc_value[:2]
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
        # print(crc_data)
        if not crc_data:
            raise ValueError("No valid data found for CRC calculation.")

        crc_data = ''.join(crc_data)
        # Convert the hex string to bytes
        crc_data = bytes.fromhex(crc_data)
        # Calculate CRC16      
        crc_value = self.compute_crc16(crc_data)
        # Update CRC fields accordingly
        self.dusr_dataframe.loc[self.dusr_dataframe['MTP address'] == '7D4A', 'Value'] = crc_value[-2:] + crc_value[:2]
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
        self.calibration_dataframe.loc[self.calibration_dataframe['MTP address'] == '7D62', 'Value'] = crc_value[-2:] + crc_value[:2]
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
        self.trac_dataframe.loc[self.trac_dataframe['MTP address'] == '7C0A', 'Value'] = crc_value[-2:] + crc_value[:2]
        self.save_changes()


    def clear_crc(self, tab):
        """Clear the CRC fields."""
        if tab == "CONFIGURATION":
            self.dataframe.loc[self.dataframe['MTP address'] == '7D92', 'Value'] = None
            self.dataframe.loc[self.dataframe['MTP address'] == '7D93', 'Value'] = None
            self.dataframe.loc[self.dataframe['MTP address'] == '7DAE', 'Value'] = None
            self.dataframe.loc[self.dataframe['MTP address'] == '7DAF', 'Value'] = None
            self.dataframe.loc[self.dataframe['MTP address'] == '7DE0', 'Value'] = None
            self.dataframe.loc[self.dataframe['MTP address'] == '7DE1', 'Value'] = None
            self.save_changes()
        if tab == "DUSR":
            self.dusr_dataframe.loc[self.dusr_dataframe['MTP address'] == '7C1E', 'Value'] = None
            self.dusr_dataframe.loc[self.dusr_dataframe['MTP address'] == '7D4A', 'Value'] = None
            self.save_changes()
        if tab == "CALIBRATION":
            self.calibration_dataframe.loc[self.calibration_dataframe['MTP address'] == '7D62', 'Value'] = None
            self.save_changes()
        if tab == "Tracking":
            self.trac_dataframe.loc[self.trac_dataframe['MTP address'] == '7C0A', 'Value'] = None
            self.save_changes()

    def save_changes(self):
        """Save changes to all relevant database sections."""
        # Save each DataFrame back to its respective section in the database
        sections = {
            "Configuration": self.dataframe,
            "DUSR": self.dusr_dataframe,
            "Calibration": self.calibration_dataframe,
            "Tracking Data": self.trac_dataframe,
            "ACBIN Data": self.prod_dataframe,
            # "OVST Flag": self.ovst_dataframe
        }
        for section_name, df in sections.items():
            self._save_section(df, section_name)
    
    def _save_section(self, df, section_name):
        """Update the database with the contents of a single section."""
        try:
            # Remove rows for the specified section
            self.connection.execute(f"DELETE FROM SiriusA2_MTPRegisterSpecification WHERE section = '{section_name}';")
            self.connection.commit()
            
            # Write the updated data back to the database
            # print(df)
            df.to_sql("SiriusA2_MTPRegisterSpecification", self.connection, if_exists="append", index=False)
        except Exception as e:
            print(f"Error saving section '{section_name}': {e}")

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
    
    def generate_binary_file_mtpproduct(self, output_path):
        """Generate a binary file from the current values."""
        trac_hex_values = self.trac_dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):X}".zfill(len(str(x)))).tolist()
        acbi_hex_values = self.prod_dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):X}".zfill(len(str(x)))).tolist()
        # conf_hex_values = self.dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):02X}").tolist()
        dusr_hex_values = self.dusr_dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):X}".zfill(len(str(x)))).tolist()
        # ovst_hex_values = self.ovst_dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):X}".zfill(len(str(x)))).tolist()
        cali_hex_values = self.calibration_dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):X}".zfill(len(str(x)))).tolist()
        # spar_hex_values = ['00'] * 20
        
        # 0x13000 ~ 0x13163
        first_half_hex_values = trac_hex_values + acbi_hex_values + dusr_hex_values + cali_hex_values
        
        # hex_values = first_half_hex_values * 2
        hex_values = first_half_hex_values
        binary_data = binascii.unhexlify(''.join(hex_values))
        with open(output_path, 'wb') as f:
            f.write(binary_data)

    def generate_binary_file_mtpfull(self, output_path):
        """Generate a binary file from the current values."""
        trac_hex_values = self.trac_dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):X}".zfill(len(str(x)))).tolist()
        acbi_hex_values = self.prod_dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):X}".zfill(len(str(x)))).tolist()
        conf_hex_values = self.dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):02X}").tolist()
        dusr_hex_values = self.dusr_dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):X}".zfill(len(str(x)))).tolist()
        # ovst_hex_values = ['00'] * 4
        cali_hex_values = self.calibration_dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):X}".zfill(len(str(x)))).tolist()
        spar_hex_values = ['00'] * 30
        
        # 0x13000 ~ 0x131FF
        first_half_hex_values = trac_hex_values + acbi_hex_values + dusr_hex_values + cali_hex_values + conf_hex_values + spar_hex_values 
        
        # hex_values = first_half_hex_values * 2
        hex_values = first_half_hex_values
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


class ExcelGUI:
    """
    The main GUI for the MTP register editing application.
    Provides controls to connect/disconnect, calculate/clear CRC, generate binary/XML files, 
    and displays the MTP register data in a tabbed notebook.
    """
    def __init__(self, root):
        self.root = root
        self.root.title(GUI_VERSION)
        self.root.geometry(GUI_SIZE)
        
        # Set an icon if available
        self.root.iconbitmap('snorlax.ico')
        
        # Grid configuration for responsiveness
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=1)

        # self.filepath = SOURCE_EXCEL
        self.filepath = SQL_FILE
        self.controller = None
        self.last_bin_file = None
        self.last_xml_file = None
        # A flag to prevent recursive calls when changing tabs
        self._syncing_tabs = False

        self.create_button_frame()
        self.create_matrix_box()
        self.create_description_box()
        self.create_register_box()
        
        # Bind both notebooks to their respective tab change events
        self.matrix_notebook.bind("<<NotebookTabChanged>>", self.on_matrix_tab_change)
        self.notebook.bind("<<NotebookTabChanged>>", self.on_register_tab_change)

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        style = ttk.Style()
        Font = 'Arial'
        Font_Color = '#000000'
        Win_BG = '#F0F0F0'
        Color = '#edfff8'
        Highlight = '#a0ada8'
        
        style.configure('TButton', font=(Font, 10), background=Color, foreground = Font_Color)
        style.map('TButton', background=[('active', Highlight)])
        style.configure('TMenubutton', font=(Font, 10), background=Color, foreground = Font_Color)
        style.map('TMenubutton', background=[('active', Highlight)])
        style.configure('TNotebook.Tab', font=(Font, 9), background=Color, foreground = Font_Color)
        style.map('TNotebook.Tab', background=[('active', Highlight)])
        style.configure('TLabelframe.Label', font=(Font, 12, 'bold'), background=Win_BG, foreground = Font_Color)
        style.configure('TLabelframe', font=(Font, 10), background=Win_BG, foreground = Font_Color)
        # style.configure('Treeview', font=(Font, 10), background=Color, foreground = Font_Color)
        
    def create_button_frame(self):
        button_frame = ttk.LabelFrame(self.root, text=" Control Panel ")
        button_frame.grid(row=0, column=0, padx=10, pady=10, sticky='nw')

        self.connect_btn = ttk.Button(button_frame, text='Disconnected', command=self.connect)
        self.connect_btn.grid(row=0, column=1, padx=6, pady=26)
        
        self.crc_btn = ttk.Button(button_frame, text='Calculate CRC\n       All tab', command=self.calculate_all_crc)        
        self.crc_btn.grid(row=0, column=2, padx=6, pady=26)

        self.crc_btn = ttk.Button(button_frame, text='Calculate CRC\n  Current tab', command=self.calculate_crc)
        self.crc_btn.grid(row=0, column=3, padx=6, pady=26)

        self.clear_crc_btn = ttk.Button(button_frame, text=' Clear CRC\nCurrent tab', command=self.clear_crc)
        self.clear_crc_btn.grid(row=0, column=4, padx=6, pady=26)
        
        self.option_var = tk.StringVar(value="mtpconfig")
        self.option_menu = ttk.OptionMenu(button_frame, self.option_var, "mtpconfig", "mtpconfig", "mtpfull", "mtpproduct")
        self.option_menu.grid(row=0, column=5, padx=10, pady=10)
        
        self.button_output_frame = ttk.LabelFrame(button_frame, text="Output Files (→ → → →)")
        self.button_output_frame.grid(row=0, column=6, padx=6, pady=8, sticky='nw')
        
        self.generate_btn = ttk.Button(self.button_output_frame, text='Generate Bin', command=self.generate)
        self.generate_btn.grid(row=0, column=0, padx=5, pady=5)

        self.generate_xml_btn = ttk.Button(self.button_output_frame, text='Generate XML', command=self.generate_xml)
        self.generate_xml_btn.grid(row=0, column=1, padx=6, pady=5)
        
        self.help = ttk.Button(button_frame, text='  Help  ', command=self.helpme)
        self.help.grid(row=0, column=7, padx=6, pady=26)
        
        # LED Canvas to indicate connection status
        self.led_canvas = tk.Canvas(button_frame, width=20, height=20, highlightthickness=0)
        self.led_canvas.grid(row=0, column=0, padx=10, pady=10)
        self.led_id = self.led_canvas.create_oval(2, 2, 18, 18, fill='red')

    def create_matrix_box(self):
        """Create the Matrix Box with seven tabs and a vertical scrollbar."""
        matrix_box = ttk.LabelFrame(self.root, text=" Matrix Box")
        matrix_box.grid(row=1, column=0, padx=10, pady=5, sticky='ew')
    
        # Create a Notebook (Tabbed Interface)
        self.matrix_notebook = ttk.Notebook(matrix_box)
        self.matrix_notebook.pack(fill='both', expand=True)
    
        # Tab names
        tab_names = [
            "Tracking",
            "Process Monitor Data",
            "DUSR",
            "CALIBRATION",
            # "OVST Flag",
            "CONFIGURATION",
            # "Redundancy Mirror"
        ]
        self.matrix_tabs = {}
    
        for tab_name in tab_names:
            tab_frame = ttk.Frame(self.matrix_notebook)
            self.matrix_notebook.add(tab_frame, text=tab_name)
            self.matrix_tabs[tab_name] = tab_frame
    
            # Add Treeview and Vertical Scrollbar to each tab
            tree_scrollbar_y = ttk.Scrollbar(tab_frame, orient='vertical')
            tree = ttk.Treeview(
                tab_frame,
                columns=("RAM Address",) + tuple(f"0x{i:02X}" for i in range(16)),
                show='headings',
                height=10,
                yscrollcommand=tree_scrollbar_y.set
            )
            tree.pack(side='left', fill='both', expand=True)
    
            tree_scrollbar_y.pack(side='right', fill='y')
            tree_scrollbar_y.config(command=tree.yview)
    
            # Define columns
            tree.heading("RAM Address", text="RAM Address")
            tree.column("RAM Address", width=100, anchor='center')
            for i in range(16):
                col_name = f"0x{i:02X}"
                tree.heading(col_name, text=col_name)
                tree.column(col_name, width=60, anchor='center')
    
            tab_frame.tree = tree
    
    def on_matrix_tab_change(self, event):
        """Callback when a tab is changed in the Matrix Notebook."""
        self.populate_matrix_tree()

    
    def populate_matrix_tree(self):
        """Populate the Matrix Box with data for the currently selected tab."""
        if not self.controller:
            return
    
        # Get the currently selected tab in the Matrix Notebook
        current_tab_name = self.matrix_notebook.tab(self.matrix_notebook.select(), "text")
        current_tab_frame = self.matrix_tabs[current_tab_name]
        current_tree = current_tab_frame.tree
    
        # Clear the current tree
        current_tree.delete(*current_tree.get_children())
    
        # Define the start and end RAM addresses based on the tab
        tab_ranges = {
            "Tracking": (0x13000, 0x1300B),  # 0x13000 ~ 0x1300B
            "Process Monitor Data": (0x13000, 0x1300F),  # 0x1300C ~ 0x1300F
            "DUSR": (0x13010, 0x1314A),  # 0x13010 ~ 0x1314A
            "CALIBRATION": (0x13140, 0x1315E),  # 0x13140 ~ 0x1315E
            #"OVST Flag": (0x13150, 0x13163),  # 0x1315F ~ 0x13163
            "CONFIGURATION": (0x13160, 0x131EB),  # 0x13164 ~ 0x131EB
            # "Redundancy Mirror": (0x13200, 0x133EB),  # 0x13200 ~ 0x133EB
        }
        start_address, end_address = tab_ranges.get(current_tab_name, (0, 0))
    
        # Create a dictionary for mapping RAM addresses to their data
        matrix_data = {}
        dataframe = None
    
        # Choose the appropriate dataframe based on the tab
        if current_tab_name == "CONFIGURATION":
            dataframe = self.controller.dataframe
        elif current_tab_name == "DUSR":
            dataframe = self.controller.dusr_dataframe
        elif current_tab_name == "CALIBRATION":
            dataframe = self.controller.calibration_dataframe
        elif current_tab_name == "Tracking":
            dataframe = self.controller.trac_dataframe
        elif current_tab_name == "Process Monitor Data":
            dataframe = self.controller.prod_dataframe
    
        if dataframe is not None:
            # Map RAM addresses to values, splitting multi-byte values
            for _, row in dataframe.iterrows():
                try:
                    ram_address = int(row['MTP address'], 16) + (0x13000 - 0x7C00)
                    value = row['Value']
                    if pd.isna(value):
                        continue
                    
                    # Ensure the value is in hexadecimal format
                    hex_value = f"{int(value, 16):0{len(value)}X}"
                    bytes_list = [hex_value[i:i + 2] for i in range(0, len(hex_value), 2)]
    
                    for i, byte in enumerate(bytes_list):
                        column = (ram_address + i) % 16
                        row_key = ((ram_address + i) // 16) * 16
                        if row_key not in matrix_data:
                            matrix_data[row_key] = [""] * 16
                        matrix_data[row_key][column] = byte.upper()
                except (ValueError, TypeError):
                    continue
    
        # Populate the Treeview, including empty rows for missing addresses
        for address in range(start_address, end_address + 0x10, 0x10):
            row_label = f"0x{address:05X}"
            row_values = matrix_data.get(address, [""] * 16)
            current_tree.insert('', 'end', values=[row_label] + row_values)

    def create_description_box(self):
        description_box = ttk.LabelFrame(self.root, text=" Description Detail")
        description_box.grid(row=2, column=0, padx=10, pady=10, sticky='ew')

        description_scrollbar = ttk.Scrollbar(description_box, orient='vertical')

        self.description_tree = ttk.Treeview(description_box, columns=("Sub Field Name", "Description"), show='headings', yscrollcommand=description_scrollbar.set, height=3)
        self.description_tree.heading("Sub Field Name", text="Sub Field Name")
        self.description_tree.column("Sub Field Name", width=120)
        self.description_tree.heading("Description", text="Description")
        self.description_tree.column("Description", width=520, minwidth=200, stretch=True)
        self.description_tree.pack(side='left', fill='both', expand=True)

        description_scrollbar.pack(side='right', fill='y')
        description_box.rowconfigure(0, weight=1)
        description_box.columnconfigure(0, weight=1)
        description_scrollbar.config(command=self.description_tree.yview)

        # Style for the description tree rows
        style = ttk.Style()
        style.configure('Description.Treeview', rowheight=120)
        self.description_tree.configure(style='Description.Treeview')
        
    def create_register_box(self):
        register_box = ttk.LabelFrame(self.root, text=" MTP Register Box")
        register_box.grid(row=0, column=1, rowspan=3, padx=10, pady=10, sticky='nsew')
        self.root.rowconfigure(1, weight=1)
        self.root.columnconfigure(1, weight=1)

        self.notebook = ttk.Notebook(register_box)
        self.notebook.pack(fill='both', expand=True)

        self.tabs = {}
        tab_names = [
            "Tracking",
            "Process Monitor Data",
            "DUSR",
            "CALIBRATION",
            #"OVST Flag",
            "CONFIGURATION",
            # "Redundancy Mirror"
        ]
        for tab_name in tab_names:
            frame = ttk.Frame(self.notebook)
            self.notebook.add(frame, text=tab_name)
            self.tabs[tab_name] = frame
            self._add_treeview_to_tab(frame)

    def on_matrix_tab_change(self, event):
        """Called when the user changes tabs in the Matrix Box Notebook."""
        # Prevent recursion if we're already syncing
        if self._syncing_tabs:
            return
        
        self._syncing_tabs = True
        try:
            # Get the newly selected tab index from matrix_notebook
            tab_index = self.matrix_notebook.index(self.matrix_notebook.select())
            # Programmatically select the same tab in the main register notebook
            self.notebook.select(tab_index)

            # (Optionally) call any method to reload or update content
            self.populate_matrix_tree()

        finally:
            self._syncing_tabs = False

    def on_register_tab_change(self, event):
        """Called when the user changes tabs in the MTP Register Box Notebook."""
        # Prevent recursion if we're already syncing
        if self._syncing_tabs:
            return

        self._syncing_tabs = True
        try:
            # Get the newly selected tab index from main register notebook
            tab_index = self.notebook.index(self.notebook.select())
            # Programmatically select the same tab in the matrix_notebook
            self.matrix_notebook.select(tab_index)

            # (Optionally) call any method to reload or update content
            self.populate_matrix_tree()

        finally:
            self._syncing_tabs = False

    def _add_treeview_to_tab(self, frame):
        tree_scrollbar = ttk.Scrollbar(frame, orient='vertical')
        tree = ttk.Treeview(frame, columns=("Register Name", "MTP Address", "RAM Address", "Byte", "Value"), show='headings', yscrollcommand=tree_scrollbar.set)
        for col in ["Register Name", "MTP Address", "RAM Address", "Byte", "Value"]:
            tree.heading(col, text=col)
            if col == "Register Name": 
                tree.column(col, width=230, anchor='w')
            elif col == "Byte": 
                tree.column(col, width=60, anchor='center')
            else: 
                tree.column(col, width=100, anchor='center')
        tree.pack(side='left', fill='both', expand=True)
        tree_scrollbar.config(command=tree.yview)
        tree_scrollbar.pack(side='right', fill='y')
        frame.tree = tree

        tree.bind("<ButtonRelease-1>", self.on_register_select)
        tree.bind("<Double-1>", self.make_value_editable)

    def on_register_select(self, event):
        selected_item = event.widget.selection()
        if not selected_item or not self.controller:
            return
    
        # Identify which tab is currently selected
        tab_name = self.notebook.tab(self.notebook.select(), "text")
    
        selected_address = event.widget.item(selected_item[0], 'values')[1]  # MTP Address
        if selected_address and selected_address.startswith("0x"):
            address = selected_address[2:]
    
            self.description_tree.delete(*self.description_tree.get_children())
    
            # Select the appropriate description dataframe based on the tab
            if tab_name == "CONFIGURATION":
                description_df = self.controller.dataframe
            elif tab_name == "DUSR":
                description_df = self.controller.dusr_dataframe
            elif tab_name == "CALIBRATION":
                description_df = self.controller.calibration_dataframe
            elif tab_name == "Tracking":
                description_df = self.controller.trac_dataframe
            elif tab_name == "Process Monitor Data":
                description_df = self.controller.prod_dataframe
            else:
                description_df = None
    
            if description_df is not None:
                # Filter by MTP address
                # description = description_df[description_df['MTP address'].str.upper() == address.upper()]
                description = description_df[description_df['MTP address'].str.upper() == address.upper()]
                description = description.sort_index()

                for _, row in description.iterrows():
                    self.description_tree.insert('', 'end', values=(row['Sub Field Name'], row['Description']))

    def on_matrix_select(self, event):
        """When a cell in the matrix is clicked, highlight the corresponding entry in the CONFIGURATION tab."""
        item = self.matrix_tree.identify('item', event.x, event.y)
        column = self.matrix_tree.identify_column(event.x)

        if not item or not column:
            return

        column_index = int(column.strip("#")) - 1
        row_label = self.matrix_tree.item(item, 'values')[0]

        if not row_label or column_index < 0:
            return

        base_address = int(row_label, 16)
        ram_address = base_address + column_index

        # Focus on the corresponding CONFIGURATION tab item
        for child in self.tabs["CONFIGURATION"].tree.get_children():
            register_address = self.tabs["CONFIGURATION"].tree.item(child, 'values')[2]
            if f"0x{ram_address:05X}" == register_address:
                self.tabs["CONFIGURATION"].tree.selection_set(child)
                self.tabs["CONFIGURATION"].tree.see(child)
                break

    def make_value_editable(self, event):
        """
        Double-clicking a cell in the Value column allows editing the value.
        All values in the Value column should be editable and updated in the copied Excel file after editing.
        The maximum editable value depends on the Byte column of this register row:
        In general, if Byte = N, the max length is N bytes and the maximum value is (1 << (N*8)) - 1.
        """
        tree = event.widget
        region = tree.identify_region(event.x, event.y)
        if region != "cell":
            return
    
        item_id = tree.identify_row(event.y)
        column_id = tree.identify_column(event.x)
        column_index = int(column_id.strip("#")) - 1
    
        # Only the Value column (index 4) is editable
        if column_index != 4:
            return
    
        item = tree.item(item_id)
        current_values = item['values']
        current_value = current_values[column_index]
        bbox = tree.bbox(item_id, column_id)
    
        # Identify which tab is currently selected
        tab_name = self.notebook.tab(self.notebook.select(), "text")
    
        if bbox:
            x, y, width, height = bbox
            entry = tk.Entry(tree, justify='center')
            entry.place(x=x, y=y, width=width, height=height)
            entry.insert(0, current_value)
            entry.select_range(0, tk.END)
            entry.focus()
    
            # Determine the allowed maximum value based on the Byte column
            # The 'Byte' value is at index 3 in the tree values
            try:
                Byte_value = int(current_values[3])  # Byte is at index 3
            except (ValueError, TypeError):
                Byte_value = 1  # Default fallback if something is off
    
            # Calculate maximum possible value given the Byte (byte count)
            # Byte_value corresponds to how many bytes. max_val = (1 << (Byte_value*8)) - 1
            max_val = (1 << (Byte_value * 8)) - 1
            # Number of hex digits is Byte_value * 2
            hex_length = Byte_value * 2
    
            def save_value(*args):
                new_value = entry.get().strip().upper()
                # Remove 0x if present for parsing
                if new_value.startswith("0X"):
                    new_value = new_value[2:]
    
                # Validate that new_value is a valid hex and within range
                try:
                    int_value = int(new_value, 16)
                    if int_value > max_val:
                        raise ValueError
                except ValueError:
                    msg = f"Value must be a valid HEX within {hex_length} hex digits (0x{'F'*hex_length})."
                    messagebox.showerror("Invalid Input", msg)
                    entry.destroy()
                    return
    
                # Zero-pad to the required hex length
                formatted_value = f"{int_value:0{hex_length}X}"
    
                tree.item(item_id, values=current_values[:column_index] + [formatted_value] + current_values[column_index+1:])
                if self.controller:
                    address = current_values[1].replace("0x", "")
                    self.controller.update_value(tab_name, address, formatted_value)
                    # Update matrix tab changed
                    if tab_name == "CONFIGURATION" or "DUSR" or "CALIBRATION" or "Tracking" or "Process Monitor Data":
                        self.populate_matrix_tree()
                entry.destroy()
    
            entry.bind("<Return>", save_value)
            entry.bind("<FocusOut>", lambda e: save_value())
      
            
    def connect_controller(self):
        """Connect to the Excel controller and populate the UI."""
        try:
            # self.controller = ExcelController(self.filepath)
            self.controller = DatabaseController(self.filepath)
            self.populate_treeview()
            self.connect_btn.config(text='Connected')
            self.led_canvas.itemconfig(self.led_id, fill='green')
        except Exception as e:
            messagebox.showerror('Error', f'Failed to connect: {e}')

    def disconnect_controller(self):
        """Disconnect from the controller and clear the UI."""
        if self.controller:
            # self.controller.delete_copy()
            self.controller.close_connection()
            self.controller = None
            self.connect_btn.config(text='Disconnected')
            self.led_canvas.itemconfig(self.led_id, fill='red')
    
            # Clear register box data
            for tab in self.tabs.values():
                tab.tree.delete(*tab.tree.get_children())
    
            # Clear description tree
            self.description_tree.delete(*self.description_tree.get_children())
    
            # Clear matrix data for all tabs
            for tab_frame in self.matrix_tabs.values():
                tab_frame.tree.delete(*tab_frame.tree.get_children())


    def connect(self):
        """Toggle connection state."""
        if self.controller is None:
            self.connect_controller()
        else:
            self.disconnect_controller()

    def calculate_crc(self):
        """Calculate CRC and update CRC fields."""
        tab_name = self.notebook.tab(self.notebook.select(), "text")
        if tab_name == "CONFIGURATION":
            if self.controller:
                try:
                    self.controller.calculate_crc_for_conf_range()
                    self.populate_treeview()
                except Exception as e:
                    messagebox.showerror('Error', f'Failed to calculate CRC: {e}')
        if tab_name == "DUSR":
            if self.controller:
                try:
                    self.controller.calculate_crc_for_DUSR_range()
                    self.populate_treeview()
                except Exception as e:
                    messagebox.showerror('Error', f'Failed to calculate CRC: {e}')
        if tab_name == "CALIBRATION":
            if self.controller:
                try:
                    self.controller.calculate_crc_for_cali_range()
                    self.populate_treeview()
                except Exception as e:
                    messagebox.showerror('Error', f'Failed to calculate CRC: {e}')
        if tab_name == "Tracking":
            if self.controller:
                try:
                    self.controller.calculate_crc_for_trac_range()
                    self.populate_treeview()
                except Exception as e:
                    messagebox.showerror('Error', f'Failed to calculate CRC: {e}')

    def calculate_all_crc(self):
        if self.controller:
            try:
                self.controller.calculate_crc_for_conf_range()
                self.controller.calculate_crc_for_DUSR_range()
                self.controller.calculate_crc_for_cali_range()
                self.controller.calculate_crc_for_trac_range()
                
                self.populate_treeview()
            except Exception as e:
                messagebox.showerror('Error', f'Failed to calculate CRC: {e}')
        
    def clear_crc(self):
        """Clear CRC fields."""
        tab_name = self.notebook.tab(self.notebook.select(), "text")
        if tab_name == "CONFIGURATION":
            if self.controller:
                try:
                    self.controller.clear_crc("CONFIGURATION")
                    self.populate_treeview()
                except Exception as e:
                    messagebox.showerror('Error', f'Failed to clear CRC: {e}')
        if tab_name == "DUSR":
            if self.controller:
                try:
                    self.controller.clear_crc("DUSR")
                    self.populate_treeview()
                except Exception as e:
                    messagebox.showerror('Error', f'Failed to clear CRC: {e}')
        if tab_name == "CALIBRATION":
            if self.controller:
                try:
                    self.controller.clear_crc("CALIBRATION")
                    self.populate_treeview()
                except Exception as e:
                    messagebox.showerror('Error', f'Failed to clear CRC: {e}')
        if tab_name == "Tracking":
            if self.controller:
                try:
                    self.controller.clear_crc("Tracking")
                    self.populate_treeview()
                except Exception as e:
                    messagebox.showerror('Error', f'Failed to clear CRC: {e}')

    def helpme(self):
        messagebox.showerror('HELP', 'Please Contact Caesal Cheng immediately')

    def populate_treeview(self):
        """Populate the treeviews for CONFIGURATION, DUSR, and CALIBRATION tabs."""
        if self.controller:
            # Clear all tabs first
            for tab_name, tab in self.tabs.items():
                tab.tree.delete(*tab.tree.get_children())

            # Populate the matrix tree from the CONFIGURATION data
            self.populate_matrix_tree()

            # Populate CONFIGURATION tab
            self._populate_tab_tree("CONFIGURATION", self.controller.dataframe)

            # Populate DUSR tab
            self._populate_tab_tree("DUSR", self.controller.dusr_dataframe)

            # Populate CALIBRATION tab
            self._populate_tab_tree("CALIBRATION", self.controller.calibration_dataframe)

            # Populate Tracking tab
            self._populate_tab_tree("Tracking", self.controller.trac_dataframe)

            # Populate Process Monitor Data tab
            self._populate_tab_tree("Process Monitor Data", self.controller.prod_dataframe)

    def _populate_tab_tree(self, tab_name, dataframe):
        grouped_data = dataframe.groupby('MTP address')
        last_register_name = None
    
        for address, group in grouped_data:
            register_name = group['Record Field'].iloc[0]
    
            # If register name is empty but there's a previous one, reuse it
            if pd.isna(register_name) and last_register_name is not None:
                register_name = last_register_name
            else:
                last_register_name = register_name
    
            # For Value, if empty, show a blank string
            vals = group['Value'].dropna().unique().astype(str)
            value = ', '.join(vals) if len(vals) > 0 else ""
    
            if tab_name == "CONFIGURATION":
                Byte = 1
            else:
                field_width = group['Field width'].iloc[0] if 'Field width' in group.columns and not pd.isna(group['Field width'].iloc[0]) else 8
                # print(field_width)
                Byte = int(field_width) 
    
            mtp_address_display = f"0x{address}"
            address_int = int(address, 16)
            actual_location = address_int + (0x13000 - 0x7C00)
            ram_address_display = f"0x{actual_location:05X}"
    
            self.tabs[tab_name].tree.insert('', 'end', 
                values=(register_name, mtp_address_display, ram_address_display, Byte, value))


    def generate(self):
        """Generate a binary file from the current data."""
        if self.controller:
            output_path = filedialog.asksaveasfilename(
                defaultextension='.bin',
                filetypes=[("Binary files", "*.bin")],
                initialfile="MTP_DATA_1K.bin"
            )
            if output_path:
                try:
                    selected_option = self.option_var.get()
                    # Determine output based on selected option
                    if selected_option == "mtpconfig":
                        self.controller.generate_binary_file_mtpconfig(output_path)
                    elif selected_option == "mtpfull":
                        self.controller.generate_binary_file_mtpfull(output_path)
                    elif selected_option == "mtpproduct":
                        self.controller.generate_binary_file_mtpproduct(output_path)
                    else: # will add WA for different selection
                        self.controller.generate_binary_file_rest(output_path)
                    self.last_bin_file = output_path
                    messagebox.showinfo("Success", "Binary file generated successfully.")
                except Exception as e:
                    messagebox.showerror('Error', f'Failed to generate binary file: {e}')

    def generate_xml(self):
        """Generate XML files based on the selected option and the last generated binary."""
        if self.last_bin_file:
            bin_dir = os.path.dirname(self.last_bin_file)
            selected_option = self.option_var.get()

            # Determine output based on selected option
            if selected_option == "mtpfull":
                output_bin = "MTP_DATA_1K_Converted.bin"
                output_xml = "mtp_batch.xml"
            elif selected_option == "mtpproduct":
                output_bin = "MTP_DATA_1K_ProductConverted.bin"
                output_xml = "mtp_Product_batch.xml"
            else:
                output_bin = "MTP_DATA_1K_ConfigConverted.bin"
                output_xml = "mtp_Config_batch.xml"

            command = ["python", "modify_binary_file_A1.py", selected_option, "MTP_DATA_1K.bin", output_bin, output_xml]

            try:
                subprocess.run(command, cwd=bin_dir, check=True, env={**os.environ, "NO_GUI_REOPEN": "1"})
                # messagebox.showinfo("Success", f"XML file generated successfully with {selected_option}.")
            except subprocess.CalledProcessError as e:
                messagebox.showerror("Error", f"Failed to generate XML: {e}")
        else:
            messagebox.showerror("Error", "Please generate a binary file first.")
        #Combine the XML file and Update file together for Addarvark one-time programming
        if self.controller:
            timestamp = datetime.now().strftime('%m%d%Y_%H%M%S')
            bin_dir = os.path.dirname(self.last_bin_file)
            selected_option = self.option_var.get()

            # Determine output based on selected option
            if selected_option == "mtpfull":
                output_xml = r"mtp_batch.xml"
            elif selected_option == "mtpproduct":
                output_xml = r"mtp_Product_batch.xml"
            else:
                output_xml = r"mtp_Config_batch.xml"
                
            output_path = filedialog.asksaveasfilename(
                defaultextension='.xml',
                filetypes=[("XML files", "*.xml")],
                initialfile=f"Sirius_{selected_option}_A1_{timestamp}.xml"
            )
            if output_path:
                try:
                    self.controller.combine_xml(f"{bin_dir}/{output_xml}", Update_File_Path, output_path)
                    messagebox.showinfo("Success", "Combined XML file generated successfully.")
                except Exception as e:
                    messagebox.showerror('Error', f'Failed to generate Combined XML file: {e}')
        
    def on_close(self):
        """Ensure we disconnect and clean up before closing the GUI."""
        self.disconnect_controller()
        self.root.destroy()


if __name__ == '__main__':
    if os.environ.get("NO_GUI_REOPEN") != "1":
        root = tk.Tk()
        style = ttk.Style(root)
        style.theme_use(GUI_THEME)
        
        app = ExcelGUI(root)
        root.mainloop()
