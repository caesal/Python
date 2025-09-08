# -*- coding: utf-8 -*-
"""
Created on Wed Dec 12 13:51:41 2024

@author: CC.Cheng

Application Name: MTP Alchemist
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import shutil
import os
import subprocess
from datetime import datetime
import binascii

# Constants
GUI_VERSION = 'MTP Alchemist Ver0.0.18'
GUI_SIZE = '1920x900'
GUI_THEME = 'clam'  # e.g., 'winnative', 'clam', 'alt', 'default', 'classic', 'vista', 'xpnative'

SOURCE_EXCEL = r'C:\WORK\Sirius\MTP\Sirius_OTP_Fusemap v23_20241206_PythonSourceExcel.xlsx'
CRC_ADDRESSES = ('7DEA', '7DEB')
CRC_RANGE_START = 0x7D67
CRC_RANGE_END = 0x7DE9
CRC_HEADER = ['AA', '55', '00']


class ExcelController:
    def __init__(self, filepath):
        self.original_path = filepath
        self.modified_path = self._create_copy()
        self.dataframe = self._load_mtp_map_tab()  # CONFIGURATION data
        self.description_dataframe = self._load_description_tab()  # CONF Description
        self.dusr_dataframe = self._load_dusr_map_tab()
        self.calibration_dataframe = self._load_calibration_map_tab()
    
        # Load DUSR and CALIBRATION descriptions
        self.dusr_description_dataframe = self._load_dusr_description_tab()
        self.calibration_description_dataframe = self._load_calibration_description_tab()
        
    def _create_copy(self):
        """Create a timestamped copy of the original Excel file for manipulation."""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        copy_path = os.path.join(os.path.dirname(self.original_path), f"Sirius_OTP_Fusemap_copy_{timestamp}.xlsm")
        shutil.copy(self.original_path, copy_path)
        return copy_path

    def _load_mtp_map_tab(self):
        """Load and prepare the MTP Map tab data."""
        df = pd.read_excel(self.modified_path, sheet_name='MTP CONF Map', header=7)
        df.rename(columns={
            df.columns[0]: 'MTP address',
            df.columns[2]: 'Field width',
            df.columns[4]: 'Record Field',
            df.columns[5]: 'Value'
        }, inplace=True)

        df.sort_values(by='MTP address', inplace=True)
        df = df[df['Value'].notna()]

        # Add CRC placeholder fields
        crc_data = pd.DataFrame({
            'MTP address': CRC_ADDRESSES,
            'Field width': [None, None],
            'Record Field': ['CONFIG0_CRC', 'CONFIG0_CRC'],
            'Value': [None, None]
        })
        df = pd.concat([df, crc_data], ignore_index=True)
        return df

    def _load_dusr_map_tab(self):
        df = pd.read_excel(self.modified_path, sheet_name='DUSR Map', header=7)
        df.rename(columns={
            df.columns[0]: 'MTP address',
            df.columns[2]: 'Field width',
            df.columns[4]: 'Record Field',
            df.columns[5]: 'Value'
        }, inplace=True)
        df.sort_values(by='MTP address', inplace=True)
        # Removed the filtering line to show rows even without values
        return df

    def _load_dusr_description_tab(self):
        df = pd.read_excel(self.modified_path, sheet_name='DUSR Description', header=7)
        df.rename(columns={
            df.columns[6]: 'Sub Field Name',  # G
            df.columns[7]: 'Description',     # H
            df.columns[0]: 'MTP address'      # A
        }, inplace=True)
        df.dropna(subset=['Sub Field Name', 'Description'], inplace=True)
        df['Sub Field Name'].fillna(method='ffill', inplace=True)
        df['Description'].fillna(method='ffill', inplace=True)
        df.reset_index(drop=True, inplace=True)
        return df

    def _load_calibration_map_tab(self):
        df = pd.read_excel(self.modified_path, sheet_name='CALIBRATION Map', header=7)
        df.rename(columns={
            df.columns[0]: 'MTP address',
            df.columns[2]: 'Field width',
            df.columns[4]: 'Record Field',
            df.columns[5]: 'Value'
        }, inplace=True)
        df.sort_values(by='MTP address', inplace=True)
        return df

    def _load_calibration_description_tab(self):
        df = pd.read_excel(self.modified_path, sheet_name='CALIBRATION Description', header=7)
        df.rename(columns={
            df.columns[6]: 'Sub Field Name',
            df.columns[7]: 'Description',
            df.columns[0]: 'MTP address'
        }, inplace=True)
        df.dropna(subset=['Sub Field Name', 'Description'], inplace=True)
        df['Sub Field Name'].fillna(method='ffill', inplace=True)
        df['Description'].fillna(method='ffill', inplace=True)
        df.reset_index(drop=True, inplace=True)
        return df

    def _calculate_crc16(self, values):
        """Calculate a 16-bit CRC for a list of hex string values."""
        crc = 0x1021
        for value in values:
            try:
                data = int(value, 16)
            except ValueError:
                continue

            for _ in range(8):
                flag = data ^ (crc >> 8)
                crc = (crc << 1) & 0xFFFF
                if flag & 0x80:
                    crc = (crc ^ 0x1021) & 0xFFFF
                data <<= 1
        return f"{crc:04X}"

    def calculate_crc_for_range(self):
        """Calculate CRC for a given address range and update CRC fields."""
        def sanitize_address(x):
            try:
                return f"{int(x, 16):04X}"
            except (ValueError, TypeError):
                return None

        self.dataframe['MTP address'] = self.dataframe['MTP address'].apply(sanitize_address)
        filtered_data = self.dataframe[
            self.dataframe['MTP address'].apply(lambda x: x is not None and CRC_RANGE_START <= int(x, 16) <= CRC_RANGE_END)
        ]

        # Collect values for CRC calculation
        values = []
        for value in filtered_data['Value']:
            if isinstance(value, str) and all(c in "0123456789ABCDEF" for c in value.upper()):
                values.append(value.upper())

        crc_data = CRC_HEADER + values
        if not crc_data:
            raise ValueError("No valid data found for CRC calculation.")

        crc_value = self._calculate_crc16(crc_data)

        # Update CRC fields accordingly
        self.dataframe.loc[self.dataframe['MTP address'] == '7DEA', 'Value'] = crc_value[:2]
        self.dataframe.loc[self.dataframe['MTP address'] == '7DEB', 'Value'] = crc_value[-2:]
        self.save_changes()

    def clear_crc(self):
        """Clear the CRC fields."""
        self.dataframe.loc[self.dataframe['MTP address'] == '7DEA', 'Value'] = None
        self.dataframe.loc[self.dataframe['MTP address'] == '7DEB', 'Value'] = None
        self.save_changes()

    def _load_description_tab(self):
        """Load and prepare the Description tab data."""
        df = pd.read_excel(self.modified_path, sheet_name='CONF Description', header=7)
        df.rename(columns={
            df.columns[6]: 'Sub Field Name',
            df.columns[7]: 'Description',
            df.columns[0]: 'MTP address'
        }, inplace=True)

        df.dropna(subset=['Sub Field Name', 'Description'], inplace=True)
        df['Sub Field Name'].fillna(method='ffill', inplace=True)
        df['Description'].fillna(method='ffill', inplace=True)
        df.reset_index(drop=True, inplace=True)
        return df

    def save_changes(self):
        """Save changes to all relevant Excel sheets."""
        with pd.ExcelWriter(self.modified_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            self.dataframe.to_excel(writer, sheet_name='MTP CONF Map', index=False)
            self.dusr_dataframe.to_excel(writer, sheet_name='DUSR Map', index=False)
            self.calibration_dataframe.to_excel(writer, sheet_name='CALIBRATION Map', index=False)
            # If description sheets or other tabs are needed, write them as well.

    def update_value(self, tab_name, address, new_value):
        """Update a specific MTP address value in the appropriate tab's dataframe."""
        if tab_name == "CONFIGURATION":
            df = self.dataframe
        elif tab_name == "DUSR":
            df = self.dusr_dataframe
        elif tab_name == "CALIBRATION":
            df = self.calibration_dataframe
        else:
            return  # No action for unrecognized tabs

        # Update the value for the given address
        df.loc[df['MTP address'].str.upper() == address.upper(), 'Value'] = new_value
        self.save_changes()

    def generate_binary_file_mtpconfig(self, output_path):
        """Generate a binary file from the current values."""
        hex_values = self.dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):02X}").tolist()
        binary_data = binascii.unhexlify(''.join(hex_values))
        padding = b'\x00' * 355
        with open(output_path, 'wb') as f:
            f.write(padding + binary_data)
            
    def generate_binary_file_rest(self, output_path):
        """Generate a binary file from the current values."""
        hex_values = self.dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):02X}").tolist()
        binary_data = binascii.unhexlify(''.join(hex_values))
        padding = b'\x00' * 0
        with open(output_path, 'wb') as f:
            f.write(padding + binary_data)

    def delete_copy(self):
        """Delete the working copy of the Excel file."""
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

        self.filepath = SOURCE_EXCEL
        self.controller = None
        self.last_bin_file = None

        self.create_button_frame()
        self.create_matrix_box()
        self.create_description_box()
        self.create_register_box()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def create_button_frame(self):
        button_frame = ttk.LabelFrame(self.root, text="Control Panel")
        button_frame.grid(row=0, column=0, padx=10, pady=10, sticky='nw')

        self.connect_btn = ttk.Button(button_frame, text='Disconnected', command=self.connect)
        self.connect_btn.grid(row=0, column=1, padx=6, pady=36)

        self.crc_btn = ttk.Button(button_frame, text='Calculate CRC', command=self.calculate_crc)
        self.crc_btn.grid(row=0, column=2, padx=6, pady=36)

        self.clear_crc_btn = ttk.Button(button_frame, text='Clear CRC', command=self.clear_crc)
        self.clear_crc_btn.grid(row=0, column=3, padx=6, pady=36)
        
        self.button_output_frame = ttk.LabelFrame(button_frame, text="Output Files (→)")
        self.button_output_frame.grid(row=0, column=5, padx=10, pady=10, sticky='nw')
        self.generate_btn = ttk.Button(self.button_output_frame, text='Generate Bin', command=self.generate)
        self.generate_btn.grid(row=0, column=0, padx=6, pady=5)

        self.generate_xml_btn = ttk.Button(self.button_output_frame, text='Generate XML', command=self.generate_xml)
        self.generate_xml_btn.grid(row=0, column=1, padx=6, pady=5)

        self.option_var = tk.StringVar(value="mtpconfig")
        self.option_menu = ttk.OptionMenu(button_frame, self.option_var, "mtpconfig", "mtpconfig", "codepatch", "mtpfull", "mtpproduct")
        self.option_menu.grid(row=0, column=4, padx=10, pady=10)

        # LED Canvas to indicate connection status
        self.led_canvas = tk.Canvas(button_frame, width=20, height=20, highlightthickness=0)
        self.led_canvas.grid(row=0, column=0, padx=10, pady=10)
        self.led_id = self.led_canvas.create_oval(2, 2, 18, 18, fill='red')

    def create_matrix_box(self):
        matrix_box = ttk.LabelFrame(self.root, text="Matrix Box")
        matrix_box.grid(row=1, column=0, padx=10, pady=5, sticky='ew')
    
        # Define columns: RAM Address plus 16 hex columns (0x00 through 0x0F)
        columns = ("RAM Address",) + tuple(f"0x{i:02X}" for i in range(16))
        self.matrix_tree = ttk.Treeview(matrix_box, columns=columns, show='headings', height=10)
    
        # Set the heading for RAM Address column
        self.matrix_tree.heading("RAM Address", text="RAM Address")
        self.matrix_tree.column("RAM Address", width=100, anchor='center')
    
        # Set headings for each 0xNN column
        for i in range(16):
            col_name = f"0x{i:02X}"
            self.matrix_tree.heading(col_name, text=col_name)
            self.matrix_tree.column(col_name, width=60, anchor='center')
    
        self.matrix_tree.pack(fill='x', expand=True)
        self.matrix_tree.bind("<ButtonRelease-1>", self.on_matrix_select)


    def create_description_box(self):
        description_box = ttk.LabelFrame(self.root, text="Description Detail")
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
        register_box = ttk.LabelFrame(self.root, text="MTP Register Box")
        register_box.grid(row=0, column=1, rowspan=3, padx=10, pady=10, sticky='nsew')
        self.root.rowconfigure(1, weight=1)
        self.root.columnconfigure(1, weight=1)

        self.notebook = ttk.Notebook(register_box)
        self.notebook.pack(fill='both', expand=True)

        self.tabs = {}
        tab_names = ["Tracking", "Process Monitor Data", "DUSR", "CALIBRATION", "OVST Flag", "CONFIGURATION", "Redundancy Mirror"]
        for tab_name in tab_names:
            frame = ttk.Frame(self.notebook)
            self.notebook.add(frame, text=tab_name)
            self.tabs[tab_name] = frame
            self._add_treeview_to_tab(frame)

    def _add_treeview_to_tab(self, frame):
        tree_scrollbar = ttk.Scrollbar(frame, orient='vertical')
        tree = ttk.Treeview(frame, columns=("Register Name", "MTP Address", "RAM Address", "Bit", "Value"), show='headings', yscrollcommand=tree_scrollbar.set)
        for col in ["Register Name", "MTP Address", "RAM Address", "Bit", "Value"]:
            tree.heading(col, text=col)
            tree.column(col, width=150, anchor='center')
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
                description_df = self.controller.description_dataframe
            elif tab_name == "DUSR":
                description_df = self.controller.dusr_description_dataframe
            elif tab_name == "CALIBRATION":
                description_df = self.controller.calibration_description_dataframe
            else:
                description_df = None
    
            if description_df is not None:
                # Filter by MTP address
                description = description_df[description_df['MTP address'].str.upper() == address.upper()]
    
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
        The maximum editable value depends on the bit column of this register row:
        In general, if bit = N, the max length is N bytes and the maximum value is (1 << (N*8)) - 1.
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
    
            # Determine the allowed maximum value based on the bit column
            # The 'bit' value is at index 3 in the tree values
            try:
                bit_value = int(current_values[3])  # Bit is at index 3
            except (ValueError, TypeError):
                bit_value = 1  # Default fallback if something is off
    
            # Calculate maximum possible value given the bit (byte count)
            # bit_value corresponds to how many bytes. max_val = (1 << (bit_value*8)) - 1
            max_val = (1 << (bit_value * 8)) - 1
            # Number of hex digits is bit_value * 2
            hex_length = bit_value * 2
    
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
                    # Update matrix only if CONFIGURATION tab changed
                    if tab_name == "CONFIGURATION":
                        self.populate_matrix_tree()
                entry.destroy()
    
            entry.bind("<Return>", save_value)
            entry.bind("<FocusOut>", lambda e: save_value())
      
            
    def connect_controller(self):
        """Connect to the Excel controller and populate the UI."""
        try:
            self.controller = ExcelController(self.filepath)
            self.populate_treeview()
            self.connect_btn.config(text='Connected')
            self.led_canvas.itemconfig(self.led_id, fill='green')
        except Exception as e:
            messagebox.showerror('Error', f'Failed to connect: {e}')

    def disconnect_controller(self):
        """Disconnect from the controller and clear the UI."""
        if self.controller:
            self.controller.delete_copy()
            self.controller = None
            self.connect_btn.config(text='Disconnected')
            self.led_canvas.itemconfig(self.led_id, fill='red')
    
            # Clear register box data
            for tab in self.tabs.values():
                tab.tree.delete(*tab.tree.get_children())
    
            # Clear description tree
            self.description_tree.delete(*self.description_tree.get_children())
    
            # Clear matrix data
            self.matrix_tree.delete(*self.matrix_tree.get_children())

    def connect(self):
        """Toggle connection state."""
        if self.controller is None:
            self.connect_controller()
        else:
            self.disconnect_controller()

    def calculate_crc(self):
        """Calculate CRC and update CRC fields."""
        if self.controller:
            try:
                self.controller.calculate_crc_for_range()
                self.populate_treeview()
            except Exception as e:
                messagebox.showerror('Error', f'Failed to calculate CRC: {e}')

    def clear_crc(self):
        """Clear CRC fields."""
        if self.controller:
            try:
                self.controller.clear_crc()
                self.populate_treeview()
            except Exception as e:
                messagebox.showerror('Error', f'Failed to clear CRC: {e}')

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
    
            # For CONFIGURATION tab, bit is always 1.
            # For DUSR and CALIBRATION tabs, bit value from 'Field width'
            if tab_name == "CONFIGURATION":
                bit = 1
            else:
                field_width = group['Field width'].iloc[0] if 'Field width' in group.columns and not pd.isna(group['Field width'].iloc[0]) else 8
                # print(field_width)
                bit = int(field_width) 
    
            mtp_address_display = f"0x{address}"
            address_int = int(address, 16)
            actual_location = address_int + (0x13000 - 0x7C00)
            ram_address_display = f"0x{actual_location:05X}"
    
            # Insert into the tree of the specified tab
            # Format: (Register Name, MTP Address, RAM Address, Bit, Value)
            self.tabs[tab_name].tree.insert('', 'end', 
                values=(register_name, mtp_address_display, ram_address_display, bit, value))

    
    def populate_matrix_tree(self):
        if not self.controller:
            return
    
        self.matrix_tree.delete(*self.matrix_tree.get_children())
        matrix_data = {}
    
        # The code to populate 'matrix_data' remains the same as before
        for _, row in self.controller.dataframe.iterrows():
            ram_address = int(row['MTP address'], 16) + (0x13000 - 0x7C00)
            column = (ram_address % 16) # 16 column + row labeling
            if column < 0:
                column = 15
            row_key = (ram_address // 16) * 16
            if row_key not in matrix_data:
                matrix_data[row_key] = [""] * 16
            matrix_data[row_key][column] = row['Value'] if pd.notna(row['Value']) else ""
    
        # Insert rows into the Treeview
        for row_key in sorted(matrix_data.keys()):
            row_label = f"0x{row_key:05X}"
            # Now we first insert 'row_label' into "RAM Address" column, 
            # followed by the 16 values for 0x00 through 0x0F columns.
            self.matrix_tree.insert('', 'end', values=[row_label] + matrix_data[row_key])

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

            command = ["python", "modify_binary_file.py", selected_option, "MTP_DATA_1K.bin", output_bin, output_xml]

            try:
                subprocess.run(command, cwd=bin_dir, check=True, env={**os.environ, "NO_GUI_REOPEN": "1"})
                messagebox.showinfo("Success", f"XML file generated successfully with {selected_option} option.")
            except subprocess.CalledProcessError as e:
                messagebox.showerror("Error", f"Failed to generate XML: {e}")
        else:
            messagebox.showerror("Error", "Please generate a binary file first.")

    def on_close(self):
        """Ensure we disconnect and clean up before closing the GUI."""
        self.disconnect_controller()
        self.root.destroy()


if __name__ == '__main__':
    if os.environ.get("NO_GUI_REOPEN") != "1":
        root = tk.Tk()
        # Set the theme before creating the GUI
        style = ttk.Style(root)
        style.theme_use(GUI_THEME)
        
        app = ExcelGUI(root)
        root.mainloop()
