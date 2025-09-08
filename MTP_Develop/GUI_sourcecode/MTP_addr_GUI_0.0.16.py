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
import subprocess
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
        """
        Calculate CRC-16 for CONFIGURATION tab with a fixed header and values from MTP addresses 0x7D68 to 0x7DE9.
        """
        def sanitize_address(x):
            """
            Ensure the address is a valid 4-digit hexadecimal string.
            """
            try:
                return f"{int(x, 16):04X}"
            except (ValueError, TypeError):
                return None
    
        # Fixed header for CRC calculation
        header = ['AA', '55', '00', '88']
    
        # Sanitize and filter MTP addresses
        self.dataframe['MTP address'] = self.dataframe['MTP address'].apply(sanitize_address)
        filtered_data = self.dataframe[
            self.dataframe['MTP address'].apply(lambda x: x is not None and 0x7D68 <= int(x, 16) <= 0x7DE9)
        ]
    
        # Extract values from the filtered range and ensure valid hexadecimal
        values = []
        for value in filtered_data['Value']:
            try:
                if isinstance(value, str) and all(c in "0123456789ABCDEF" for c in value.upper()):
                    values.append(value.upper())
            except Exception:
                continue
    
        # Combine header and values for CRC calculation
        crc_data = header + values
    
        if not crc_data:
            raise ValueError("No valid data found for CRC calculation.")
    
        # Calculate CRC
        crc_value = self._calculate_crc16(crc_data)
    
        # Update CRC values for 7DEA and 7DEB
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
        padding = b'\x00' * 355
        with open(output_path, 'wb') as f:
            f.write(padding + binary_data)

    def delete_copy(self):
        if os.path.exists(self.modified_path):
            os.remove(self.modified_path)


class ExcelGUI:
    def __init__(self, root):
        self.root = root
        self.root.title('MTP_addr_GUI_0.0.15')
        self.root.geometry('1920x900')
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=1)

        # Filepath to be used
        self.filepath = r'C:\WORK\Sirius\MTP\Sirius_OTP_Fusemap v23_20241206_PythonSourceExcel.xlsx'
        self.controller = None
        self.last_bin_file = None

        # UI Elements
        self.create_button_frame()
        self.create_matrix_box()
        self.create_description_box()
        self.create_register_box()

    def create_button_frame(self):
        button_frame = ttk.LabelFrame(self.root, text="Controls")
        button_frame.grid(row=0, column=0, padx=10, pady=10, sticky='nw')

        self.connect_btn = ttk.Button(button_frame, text='Disconnected', command=self.connect)
        self.connect_btn.grid(row=0, column=0, padx=5, pady=5)

        self.crc_btn = ttk.Button(button_frame, text='Calculate CRC', command=self.calculate_crc)
        self.crc_btn.grid(row=0, column=1, padx=5, pady=5)

        self.clear_crc_btn = ttk.Button(button_frame, text='Clear CRC', command=self.clear_crc)
        self.clear_crc_btn.grid(row=0, column=2, padx=5, pady=5)

        self.generate_btn = ttk.Button(button_frame, text='Generate Bin', command=self.generate)
        self.generate_btn.grid(row=0, column=3, padx=5, pady=5)

        self.generate_xml_btn = ttk.Button(button_frame, text='Generate XML', command=self.generate_xml)
        self.generate_xml_btn.grid(row=0, column=4, padx=5, pady=5)

        self.option_var = tk.StringVar(value="mtpconfig")
        self.option_menu = ttk.OptionMenu(button_frame, self.option_var, "mtpconfig", "mtpconfig", "codepatch", "mtpfull", "mtpproduct")
        self.option_menu.grid(row=0, column=5, padx=5, pady=5)

        self.led_canvas = tk.Canvas(button_frame, width=20, height=20)
        self.led_canvas.grid(row=1, column=0, padx=5, pady=5)
        self.led_id = self.led_canvas.create_oval(2, 2, 18, 18, fill='red')

    def create_matrix_box(self):
        matrix_box = ttk.LabelFrame(self.root, text="Matrix Box")
        matrix_box.grid(row=1, column=0, padx=10, pady=5, sticky='ew')

        self.matrix_tree = ttk.Treeview(matrix_box, columns=tuple(range(16)), show='headings', height=10)
        for i in range(16):
            self.matrix_tree.heading(i, text=f"0x{i:02X}")
            self.matrix_tree.column(i, width=60, anchor='center')
        self.matrix_tree.pack(fill='x', expand=True)

        self.matrix_tree.bind("<ButtonRelease-1>", self.on_matrix_select)

    def create_description_box(self):
        description_box = ttk.LabelFrame(self.root, text="Description Detail")
        description_box.grid(row=2, column=0, padx=10, pady=10, sticky='ew')

        description_scrollbar = ttk.Scrollbar(description_box, orient='vertical')

        self.description_tree = ttk.Treeview(description_box, columns=("Sub Field Name", "Description"), show='headings', yscrollcommand=description_scrollbar.set, height=3)
        self.description_tree.heading("Sub Field Name", text="Sub Field Name")
        self.description_tree.column("Sub Field Name", width=160)
        self.description_tree.heading("Description", text="Description")
        self.description_tree.column("Description", width=520, minwidth=200, stretch=True)
        self.description_tree.pack(side='left', fill='both', expand=True)

        description_scrollbar.pack(side='right', fill='y')
        description_box.rowconfigure(0, weight=1)
        description_box.columnconfigure(0, weight=1)
        description_scrollbar.config(command=self.description_tree.yview)
        # Set a default row height
        style = ttk.Style()
        style.configure('Description.Treeview', rowheight=120)  # Increased row height for better visibility
        self.description_tree.configure(style='Description.Treeview')
        
    def create_register_box(self):
        register_box = ttk.LabelFrame(self.root, text="MTP Register Box")
        register_box.grid(row=0, column=1, rowspan=3, padx=10, pady=10, sticky='nsew')
        self.root.rowconfigure(1, weight=1)
        self.root.columnconfigure(1, weight=1)

        self.notebook = ttk.Notebook(register_box)
        self.notebook.pack(fill='both', expand=True)

        self.tabs = {}
        tab_names = ["Tracking", "Process Monitor Data", "DUSR", "Calibration", "OVST Flag", "CONFIGURATION", "Minor"]
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

        # Bind selection event
        tree.bind("<ButtonRelease-1>", self.on_register_select)
        tree.bind("<Double-1>", self.make_value_editable)

    def on_register_select(self, event):
        selected_item = event.widget.selection()
        if not selected_item or not self.controller:
            return

        selected_address = event.widget.item(selected_item[0], 'values')[1]  # Get MTP Address
        if selected_address and selected_address.startswith("0x"):
            address = selected_address[2:]  # Strip "0x"
            description = self.controller.description_dataframe[
                self.controller.description_dataframe['MTP address'] == address
            ]

            # Clear current contents of the description tree
            self.description_tree.delete(*self.description_tree.get_children())

            # Populate description box with matching entries
            for _, row in description.iterrows():
                self.description_tree.insert('', 'end', values=(row['Sub Field Name'], row['Description']))

    def on_matrix_select(self, event):
        # Identify the item (row) and column clicked
        item = self.matrix_tree.identify('item', event.x, event.y)
        column = self.matrix_tree.identify_column(event.x)

        if not item or not column:
            return

        # Determine the row and column indices
        column_index = int(column.strip("#")) - 1  # Convert to zero-based index
        row_label = self.matrix_tree.item(item, 'values')[0]  # Get row label (e.g., "0x131D0")

        if not row_label or column_index < 0:
            return

        # Calculate the corresponding RAM address
        base_address = int(row_label, 16)
        ram_address = base_address + column_index

        # Highlight the corresponding row in the CONFIGURATION tab
        for tab_name, tab in self.tabs.items():
            if tab_name == "CONFIGURATION":
                for child in tab.tree.get_children():
                    register_address = tab.tree.item(child, 'values')[2]  # RAM Address column
                    if f"0x{ram_address:05X}" == register_address:
                        tab.tree.selection_set(child)
                        tab.tree.see(child)
                        break

    def make_value_editable(self, event):
        tree = event.widget
        region = tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        item_id = tree.identify_row(event.y)
        column_id = tree.identify_column(event.x)
        column_index = int(column_id.strip("#")) - 1

        if column_index != 4:  # Only allow editing the "Value" column
            return

        # Get current value and cell position
        item = tree.item(item_id)
        current_value = item['values'][column_index]
        bbox = tree.bbox(item_id, column_id)

        if bbox:
            x, y, width, height = bbox
            entry = tk.Entry(tree, justify='center')
            entry.place(x=x, y=y, width=width, height=height)
            entry.insert(0, current_value)
            entry.select_range(0, tk.END)
            entry.focus()

            def save_value(*args):
                new_value = entry.get().strip().upper()
                # Ensure value is valid HEX and <= 0xFF
                try:
                    int_value = int(new_value, 16)
                    if int_value > 0xFF:
                        raise ValueError
                except ValueError:
                    messagebox.showerror("Invalid Input", "Value must be a valid 2-byte HEX (0x00 to 0xFF).")
                    entry.destroy()
                    return

                # Update Treeview and controller
                tree.item(item_id, values=item['values'][:column_index] + [f"{int_value:02X}"] + item['values'][column_index+1:])
                if self.controller:
                    address = item['values'][1].replace("0x", "")  # Extract MTP address
                    self.controller.update_value(address, f"{int_value:02X}")
                    self.populate_matrix_tree()
                entry.destroy()

            entry.bind("<Return>", save_value)
            entry.bind("<FocusOut>", lambda e: save_value())

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
            self.populate_matrix_tree()

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
                value = ', '.join(group['Value'].dropna().unique().astype(str))
                self.tabs["CONFIGURATION"].tree.insert('', 'end', values=(register_name, mtp_address_display, ram_address_display, bit, value))

    def populate_matrix_tree(self):
        self.matrix_tree.delete(*self.matrix_tree.get_children())
        matrix_data = {}

        for index, row in self.controller.dataframe.iterrows():
            ram_address = int(row['MTP address'], 16) + (0x13000 - 0x7C00)
            column = (ram_address % 16) - 1  # Adjust for correct alignment
            if column < 0:
                column = 15  # Wrap around to the last column
            row_key = ram_address // 16 * 16
            if row_key not in matrix_data:
                matrix_data[row_key] = [""] * 16
            matrix_data[row_key][column] = row['Value'] if pd.notna(row['Value']) else ""

        for row_key in sorted(matrix_data.keys()):
            row_label = f"0x{row_key:05X}"
            self.matrix_tree.insert('', 'end', values=[row_label] + matrix_data[row_key])


    def generate(self):
        if self.controller:
            output_path = filedialog.asksaveasfilename(defaultextension='.bin', filetypes=[("Binary files", "*.bin")], initialfile="MTP_DATA_1K.bin")
            if output_path:
                try:
                    self.controller.generate_binary_file(output_path)
                    self.last_bin_file = output_path
                    messagebox.showinfo("Success", "Binary file generated successfully.")
                except Exception as e:
                    messagebox.showerror('Error', f'Failed to generate binary file: {str(e)}')

    def generate_xml(self):
        if self.last_bin_file:
            bin_dir = os.path.dirname(self.last_bin_file)
            selected_option = self.option_var.get()
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
                messagebox.showerror("Error", f"Failed to generate XML: {str(e)}")
        else:
            messagebox.showerror("Error", "Please generate a binary file first.")

    def on_close(self):
        if self.controller:
            self.controller.delete_copy()
        self.root.destroy()


if __name__ == '__main__':
    if os.environ.get("NO_GUI_REOPEN") != "1":
        root = tk.Tk()
        app = ExcelGUI(root)
        root.mainloop()
