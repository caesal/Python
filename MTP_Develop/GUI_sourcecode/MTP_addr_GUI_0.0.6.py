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

    def _create_copy(self):
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        copy_path = os.path.join(os.path.dirname(self.original_path), f"Sirius_OTP_Fusemap_copy_{timestamp}.xlsm")
        shutil.copy(self.original_path, copy_path)
        return copy_path

    def _load_mtp_map_tab(self):
        # Load the MTP Map tab and start reading from row 8 onwards
        df = pd.read_excel(self.modified_path, sheet_name='MTP Map', header=7)
        df.rename(columns={df.columns[0]: 'MTP address', df.columns[2]: 'Field width', df.columns[4]: 'Record Field', df.columns[5]: 'Value', df.columns[6]: 'Sub Field Name', df.columns[7]: 'Description'}, inplace=True)
        df.sort_values(by='MTP address', inplace=True)  # Sort by MTP address
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
        self.root.title('Excel MTP Map Controller')
        self.root.geometry('1440x750')
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=1)

        # Filepath to be used
        self.filepath = r'C:\WORK\Sirius\MTP\Sirius_OTP_Fusemap v23_20241206_Orig.xlsm'
        self.controller = None

        # UI Elements
        button_frame = ttk.Frame(root)
        button_frame.grid(row=0, column=0, padx=10, pady=10, sticky='nw')

        self.connect_btn = ttk.Button(button_frame, text='Disconnected', command=self.connect)
        self.connect_btn.grid(row=0, column=0, padx=5, pady=5)

        self.generate_btn = ttk.Button(button_frame, text='Generate', command=self.generate)
        self.generate_btn.grid(row=0, column=1, padx=5, pady=5)

        self.led_canvas = tk.Canvas(button_frame, width=20, height=20)
        self.led_canvas.grid(row=1, column=0, padx=5, pady=5)
        self.led_id = self.led_canvas.create_oval(2, 2, 18, 18, fill='green')

        # Add Treeview with scrollbar for the register box
        tree_frame = ttk.Frame(root)
        tree_frame.grid(row=0, column=1, rowspan=3, padx=10, pady=10, sticky='nsew')

        tree_scrollbar = ttk.Scrollbar(tree_frame, orient='vertical')
        self.tree = ttk.Treeview(tree_frame, columns=("Register Name", "Address", "Bit", "Value"), show='headings', yscrollcommand=tree_scrollbar.set)
        for col in ["Register Name", "Address", "Bit", "Value"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor='center')
        self.tree.pack(side='left', fill='both', expand=True)

        tree_scrollbar.config(command=self.tree.yview)
        tree_scrollbar.pack(side='right', fill='y')

        # Make the value column directly editable
        self.tree.bind('<Double-1>', self.on_double_click_edit)
        self.tree.bind('<Return>', self.on_value_enter)
        

        # Add Treeview for the description box with scrollbar
        description_frame = ttk.Frame(root)
        description_frame.grid(row=1, column=0, padx=10, pady=5, sticky='nsew')

        description_scrollbar = ttk.Scrollbar(description_frame, orient='vertical')
        self.description_tree = ttk.Treeview(description_frame, columns=("Sub Field Name", "Description"), show='headings', yscrollcommand=description_scrollbar.set, height=10)
        self.description_tree.heading("Sub Field Name", text="Sub Field Name")
        self.description_tree.column("Sub Field Name", width=150)
        self.description_tree.heading("Description", text="Description")
        self.description_tree.column("Description", width=350, minwidth=200, stretch=True)
        self.description_tree.pack(side='left', fill='both', expand=True)

        description_scrollbar.config(command=self.description_tree.yview)
        description_scrollbar.pack(side='right', fill='y')

        # Configure resizing behavior
        root.grid_rowconfigure(0, weight=1)
        root.grid_rowconfigure(1, weight=1)
        root.grid_columnconfigure(1, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        description_frame.grid_rowconfigure(0, weight=1)
        description_frame.grid_columnconfigure(0, weight=1)

        # Bind treeview selection event to show description
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)

        # Handle window close event
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def connect(self):
        if self.controller is None:
            try:
                self.controller = ExcelController(self.filepath)
                self.populate_treeview()
                self.connect_btn.config(text='Connected')
                self.led_canvas.itemconfig(self.led_id, fill='red')
            except Exception as e:
                messagebox.showerror('Error', f'Failed to connect: {str(e)}')
        else:
            # Disconnect
            self.controller.delete_copy()
            self.controller = None
            self.connect_btn.config(text='Disconnected')
            self.led_canvas.itemconfig(self.led_id, fill='green')
            self.tree.delete(*self.tree.get_children())
            self.description_tree.delete(*self.description_tree.get_children())

    def populate_treeview(self):
        if self.controller:
            self.tree.delete(*self.tree.get_children())
            grouped_data = self.controller.dataframe.groupby('MTP address')
            last_register_name = None
            for address, group in grouped_data:
                register_name = group['Record Field'].iloc[0]
                if pd.isna(register_name) and group['Value'].notna().any() and last_register_name is not None:
                    register_name = last_register_name
                else:
                    last_register_name = register_name
                bit = 8  # Set bit value to always be 8
                value = ', '.join(group['Value'].dropna().unique().astype(str))  # Concatenate all unique non-NA values
                self.tree.insert('', 'end', values=(register_name, address, bit, value))

    def on_tree_select(self, event):
        selected_item = self.tree.selection()
        if selected_item:
            item = self.tree.item(selected_item)
            address = item['values'][1]

            # Get the index of the selected address
            address_indices = self.controller.dataframe[self.controller.dataframe['MTP address'] == address].index
            start_index = address_indices[0]

            # Find the next address and its index
            next_index = start_index + 1
            while next_index < len(self.controller.dataframe) and pd.isna(self.controller.dataframe.iloc[next_index]['MTP address']):
                next_index += 1

            # If the difference between indices is greater than 1, include all rows between them
            if next_index < len(self.controller.dataframe):
                end_index = next_index
            else:
                end_index = len(self.controller.dataframe)

            description_rows = self.controller.dataframe.iloc[start_index:end_index]

            # Clear the description tree and populate it with relevant rows
            self.description_tree.delete(*self.description_tree.get_children())
            for _, row in description_rows.iterrows():
                sub_field_name = row['Sub Field Name']
                description = row['Description']
                if pd.notna(description):  # Check if description is not NaN
                    wrapped_description = '\n'.join(self.wrap_text(str(description), 50))
                else:
                    wrapped_description = 'Check above register description'
                self.description_tree.insert('', 'end', values=(sub_field_name, wrapped_description))

    def on_double_click_edit(self, event):
        item = self.tree.selection()[0]
        column = self.tree.identify_column(event.x)
        if column == '#4':  # Value column
            entry = tk.Entry(self.tree, width=10)
            entry.place(x=event.x_root - self.tree.winfo_rootx(), y=event.y_root - self.tree.winfo_rooty())
            entry.insert(0, self.tree.item(item, 'values')[3])
            entry.focus()

            def on_entry_confirm(e):
                new_value = entry.get().strip().upper()
                if len(new_value) <= 2 and all(c in '0123456789ABCDEF' for c in new_value):
                    values = list(self.tree.item(item, 'values'))
                    values[3] = new_value
                    self.tree.item(item, values=values)
                    address = values[1]
                    self.controller.update_value(address, new_value)
                    entry.destroy()
                else:
                    messagebox.showerror('Invalid Input', 'Value must be a hexadecimal number between 00 and FF.')
                    entry.destroy()

            entry.bind('<Return>', on_entry_confirm)

    def on_value_enter(self, event):
        item = self.tree.selection()[0]
        column = self.tree.identify_column(event.x)
        if column == '#4':  # Value column
            entry = tk.Entry(self.tree, width=10)
            entry.place(x=event.x_root - self.tree.winfo_rootx(), y=event.y_root - self.tree.winfo_rooty())
            entry.insert(0, self.tree.item(item, 'values')[3])
            entry.focus()

            def on_entry_confirm(e):
                new_value = entry.get().strip().upper()
                if len(new_value) <= 2 and all(c in '0123456789ABCDEF' for c in new_value):
                    values = list(self.tree.item(item, 'values'))
                    values[3] = new_value
                    self.tree.item(item, values=values)
                    address = values[1]
                    self.controller.update_value(address, new_value)
                    entry.destroy()
                else:
                    messagebox.showerror('Invalid Input', 'Value must be a hexadecimal number between 00 and FF.')
                    entry.destroy()

            entry.bind('<Return>', on_entry_confirm)

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
