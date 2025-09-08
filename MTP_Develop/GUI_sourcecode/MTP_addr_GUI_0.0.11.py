# -*- coding: utf-8 -*-
"""
Created on Thu Dec  5 14:28:38 2024

@author: CC.Cheng
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import shutil
import os
from datetime import datetime
import binascii

# Load the file and create a copy for modifications
class ExcelController:
    def __init__(self, filepath):
        self.original_path = filepath
        self.modified_path = self._create_copy()
        self.dataframe = self._load_mtp_map_tab()
        self.description_dataframe = self._load_description_tab()

    def _create_copy(self):
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        copy_path = os.path.join(os.path.dirname(self.original_path), f"Sirius_OTP_Fusemap_copy_{timestamp}.xlsm")
        shutil.copy(self.original_path, copy_path)
        return copy_path

    def _load_mtp_map_tab(self):
        # Load the MTP Map tab and start reading from row 8 onwards
        df = pd.read_excel(self.modified_path, sheet_name='MTP Map', header=7)
        df.rename(columns={df.columns[0]: 'MTP address', df.columns[2]: 'Field width', df.columns[4]: 'Record Field', df.columns[5]: 'Value'}, inplace=True)
        df.sort_values(by='MTP address', inplace=True)  # Sort by MTP address
        df = df[df['Value'].notna()]  # Remove rows where 'Value' is NaN
        
        # Add CRC fields for addresses 7DEA and 7DEB
        crc_data = pd.DataFrame({'MTP address': ['7DEA', '7DEB'], 'Field width': [None, None], 'Record Field': ['CONFIG0_CRC', 'CONFIG0_CRC'], 'Value': [None, None]})
        df = pd.concat([df, crc_data], ignore_index=True)
        
        return df

    def _calculate_crc16(self, values):
        # Calculate CRC-16 using the provided reference algorithm
        crc = 0x1021
        for value in values:
            try:
                data = int(value, 16)
            except ValueError:
                continue  # Skip invalid values

            for _ in range(8):
                flag = data ^ (crc >> 8)
                crc = (crc << 1) & 0xFFFF
                if flag & 0x80:
                    crc = (crc ^ 0x1021) & 0xFFFF
                data <<= 1
        return f"{crc:04X}"

    def calculate_crc_for_range(self):
        # Calculate CRC-16 for the range from address 7D64 to 7DE9
        values = [v for v in self.dataframe[self.dataframe['MTP address'].between('7D64', '7DE9')]['Value'].dropna() if isinstance(v, str)]
        crc_value = self._calculate_crc16(values)
        self.dataframe.loc[self.dataframe['MTP address'] == '7DEA', 'Value'] = crc_value[-2:]
        self.dataframe.loc[self.dataframe['MTP address'] == '7DEB', 'Value'] = crc_value[:2]
        self.save_changes()

    def clear_crc(self):
        # Clear the CRC values for addresses 7DEA and 7DEB
        self.dataframe.loc[self.dataframe['MTP address'] == '7DEA', 'Value'] = None
        self.dataframe.loc[self.dataframe['MTP address'] == '7DEB', 'Value'] = None
        self.save_changes()

    def _load_description_tab(self):
        # Load the Description tab, starting from row 8 onwards
        df = pd.read_excel(self.modified_path, sheet_name='Description', header=7)
        df.rename(columns={df.columns[6]: 'Sub Field Name', df.columns[7]: 'Description', df.columns[0]: 'MTP address'}, inplace=True)
        df.dropna(subset=['Sub Field Name', 'Description'], inplace=True)  # Remove rows with NaN in both columns
        df['Sub Field Name'].fillna(method='ffill', inplace=True)
        df['Description'].fillna(method='ffill', inplace=True)
        df.reset_index(drop=True, inplace=True)  # Reset index to avoid key errors
        return df

    def save_changes(self):
        with pd.ExcelWriter(self.modified_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            self.dataframe.to_excel(writer, sheet_name='MTP Map', index=False)

    def update_value(self, address, new_value):
        # Update the value in the dataframe
        self.dataframe.loc[self.dataframe['MTP address'] == address, 'Value'] = new_value
        self.save_changes()

    def generate_binary_file(self, output_path):
        hex_values = self.dataframe['Value'].dropna().apply(lambda x: f"{int(str(x), 16):02X}").tolist()
        binary_data = binascii.unhexlify(''.join(hex_values))
        with open(output_path, 'wb') as f:
            f.write(binary_data)

    def delete_copy(self):
        if os.path.exists(self.modified_path):
            os.remove(self.modified_path)


class ExcelGUI:
    def __init__(self, root):
        self.root = root
        self.root = root
        self.root.title('MTP_addr_GUI_0.0.10')
        self.root.geometry('1650x960')
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=1)

        # Filepath to be used
        self.filepath = r'C:\WORK\Sirius\MTP\Sirius_OTP_Fusemap v23_20241206_PythonSourceExcel.xlsx'
        self.controller = None

        # UI Elements
        button_frame = ttk.Frame(root)
        button_frame.grid(row=0, column=0, padx=10, pady=10, sticky='nw')

        self.connect_btn = ttk.Button(button_frame, text='Disconnected', command=self.connect)
        self.connect_btn.grid(row=0, column=0, padx=5, pady=5)

        self.crc_btn = ttk.Button(button_frame, text='Calculate CRC', command=self.calculate_crc)
        self.crc_btn.grid(row=0, column=1, padx=5, pady=5)

        self.clear_crc_btn = ttk.Button(button_frame, text='Clear CRC', command=self.clear_crc)
        self.clear_crc_btn.grid(row=0, column=2, padx=5, pady=5)

        self.generate_btn = ttk.Button(button_frame, text='Generate', command=self.generate)
        self.generate_btn.grid(row=0, column=3, padx=5, pady=5)

        self.led_canvas = tk.Canvas(button_frame, width=20, height=20)
        self.led_canvas.grid(row=1, column=0, padx=5, pady=5)
        self.led_id = self.led_canvas.create_oval(2, 2, 18, 18, fill='red')

        # Add Notebook for tabs
        notebook_frame = ttk.Frame(root)
        notebook_frame.grid(row=0, column=1, rowspan=3, padx=10, pady=10, sticky='nsew')
        self.notebook = ttk.Notebook(notebook_frame)
        self.notebook.pack(fill='both', expand=True)

        # Create tabs
        self.tabs = {}
        tab_names = ["Tracking", "Process Monitor Data", "DUSR", "Calibration", "OVST Flag", "CONFIGURATION", "Minor"]
        for tab_name in tab_names:
            frame = ttk.Frame(self.notebook)
            self.notebook.add(frame, text=tab_name)
            self.tabs[tab_name] = frame
            self._add_treeview_to_tab(frame)

        # Add Treeview for the description box with scrollbar
        description_frame = ttk.Frame(root)
        description_frame.grid(row=1, column=0, padx=20, pady=20, sticky='nsew')
        root.rowconfigure(1, weight=1)
        root.columnconfigure(0, weight=1)
        description_frame.rowconfigure(0, weight=1)
        description_frame.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)
        root.columnconfigure(0, weight=1)
        
        description_scrollbar = ttk.Scrollbar(description_frame, orient='vertical')
        
        # Set a fixed height (number of rows)
        fixed_height = 3
        
        self.description_tree = ttk.Treeview(description_frame, columns=("Sub Field Name", "Description"), show='headings', yscrollcommand=description_scrollbar.set, height=fixed_height)
        self.description_tree.heading("Sub Field Name", text="Sub Field Name")
        self.description_tree.column("Sub Field Name", width=160)
        
        # Set the width of the Description column to 800 pixels for better visibility of long content
        self.description_tree.heading("Description", text="Description")
        self.description_tree.column("Description", width=520, minwidth=200, stretch=True)
        self.description_tree.column("Description", stretch=True)
        
        self.description_tree.pack(side='left', fill='both', expand=True)
        description_scrollbar.pack(side='right', fill='y')
        description_frame.rowconfigure(0, weight=1)
        description_frame.columnconfigure(0, weight=1)
        
        
        description_scrollbar.config(command=self.description_tree.yview)
        description_scrollbar.pack(side='right', fill='y')
        
        # Set a default row height
        style = ttk.Style()
        style.configure('Description.Treeview', rowheight=120)  # Increased row height for better visibility
        self.description_tree.configure(style='Description.Treeview')
        
        # Fix the height of description tree (do not change dynamically)
        
        # Bind treeview selection event to show description
        self.tabs["CONFIGURATION"].tree.bind('<<TreeviewSelect>>', self.on_tree_select)

        # Handle window close event
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _add_treeview_to_tab(self, frame):
        tree_scrollbar = ttk.Scrollbar(frame, orient='vertical')
        tree = ttk.Treeview(frame, columns=("Register Name", "MTP Address", "RAM Address", "Bit", "Value"), show='headings', yscrollcommand=tree_scrollbar.set)
        for col in ["Register Name", "MTP Address", "RAM Address", "Bit", "Value"]:
            tree.heading(col, text=col)
            tree.column(col, width=150, anchor='center')
        tree.pack(side='left', fill='both', expand=True)
        tree_scrollbar.config(command=tree.yview)
        tree_scrollbar.pack(side='right', fill='y')
        
        tree.bind('<Double-1>', lambda e, tree=tree: self.on_double_click_edit(e, tree))
        
        # Store reference to each tree in the dictionary
        frame.tree = tree

    def connect(self):
        if self.controller is None:
            try:
                self.controller = ExcelController(self.filepath)
                self.populate_treeview()
                self.connect_btn.config(text='Connected')
                self.led_canvas.itemconfig(self.led_id, fill='green')
            except Exception as e:
                messagebox.showerror('Error', f'Failed to connect: {str(e)}')
        else:
            # Disconnect
            self.controller.delete_copy()
            self.controller = None
            self.connect_btn.config(text='Disconnected')
            self.led_canvas.itemconfig(self.led_id, fill='red')
            for tab in self.tabs.values():
                tab.tree.delete(*tab.tree.get_children())
            self.description_tree.delete(*self.description_tree.get_children())

    def calculate_crc(self):
        if self.controller:
            try:
                self.controller.calculate_crc_for_range()
                
                self.populate_treeview()
            except Exception as e:
                messagebox.showerror('Error', f'Failed to calculate CRC: {str(e)}')

    def clear_crc(self):
        if self.controller:
            try:
                self.controller.clear_crc()
                
                self.populate_treeview()
            except Exception as e:
                messagebox.showerror('Error', f'Failed to clear CRC: {str(e)}')

    def populate_treeview(self):
        if self.controller:
            # Clear all tabs first
            for tab_name, tab in self.tabs.items():
                tab.tree.delete(*tab.tree.get_children())
            
            # Only populate the CONFIGURATION tab with data from the Excel file
            grouped_data = self.controller.dataframe.groupby('MTP address')
            last_register_name = None
            for address, group in grouped_data:
                register_name = group['Record Field'].iloc[0]
                if pd.isna(register_name) and group['Value'].notna().any() and last_register_name is not None:
                    register_name = last_register_name
                else:
                    last_register_name = register_name
                bit = 8  # Set bit value to always be 8
                address_int = int(address, 16)
                actual_location = address_int + (0x13000 - 0x7C00)
                mtp_address_display = f"0x{address}"
                ram_address_display = f"0x{actual_location:05X}"
                value = ', '.join(group['Value'].dropna().unique().astype(str))  # Concatenate all unique non-NA values
                self.tabs["CONFIGURATION"].tree.insert('', 'end', values=(register_name, mtp_address_display, ram_address_display, bit, value))

    
    def on_tree_select(self, event):
        selected_item = event.widget.selection()
        if selected_item:
            item = event.widget.item(selected_item)
            mtp_address = item['values'][1]
            ram_address = item['values'][2]
        
            # Find all rows matching the selected address in the description dataframe
            description_rows = self.controller.description_dataframe[self.controller.description_dataframe['MTP address'] == mtp_address[2:]]
        
            # Clear the description tree and populate it with relevant rows
            self.description_tree.delete(*self.description_tree.get_children())
            
            for _, row in description_rows.iterrows():
                sub_field_name = row['Sub Field Name']
                description = row['Description']
                if pd.notna(description):  # Check if description is not NaN
                    wrapped_description = description  # Keep the description as it is without wrapping
                else:
                    wrapped_description = 'Check above register description'
        
                # Insert the description into the treeview
                self.description_tree.insert('', 'end', values=(sub_field_name, wrapped_description))

    def on_double_click_edit(self, event, tree):
        item = tree.selection()[0]
        column = tree.identify_column(event.x)
        if column == '#5':  # Value column
            entry = tk.Entry(tree, width=10)
            entry.place(x=event.x_root - tree.winfo_rootx(), y=event.y_root - tree.winfo_rooty())
            entry.insert(0, tree.item(item, 'values')[3])
            entry.selection_range(0, tk.END)
            entry.focus()

            def on_entry_confirm(e):
                new_value = entry.get().strip().upper()
                if len(new_value) <= 2 and all(c in '0123456789ABCDEF' for c in new_value):
                    values = list(tree.item(item, 'values'))
                    values[3] = new_value
                    tree.item(item, values=values)
                    address = values[1][2:]
                    self.controller.update_value(address, new_value)
                    entry.destroy()
                else:
                    messagebox.showerror('Invalid Input', 'Value must be a hexadecimal number between 00 and FF.')
                    entry.destroy()

            entry.bind('<Return>', on_entry_confirm)
            entry.bind('<Return>', lambda e: [on_entry_confirm(e), self.populate_treeview()])
            entry.bind('<Escape>', lambda e: entry.destroy())
            entry.bind('<FocusOut>', lambda e: entry.destroy())
            tree.focus_set()

    def wrap_text(self, text, width):
        return [text[i:i+width] for i in range(0, len(text), width)]

    def generate(self):
        if self.controller:
            output_path = filedialog.asksaveasfilename(defaultextension='.bin', filetypes=[("Binary files", "*.bin")])
            if output_path:
                try:
                    self.controller.generate_binary_file(output_path)
                except Exception as e:
                    messagebox.showerror('Error', f'Failed to generate binary file: {str(e)}')

    def on_close(self):
        if self.controller:
            self.controller.delete_copy()
        self.root.destroy()


if __name__ == '__main__':
    root = tk.Tk()
    app = ExcelGUI(root)
    root.mainloop()
