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
        df.rename(columns={df.columns[0]: 'MTP address', df.columns[2]: 'Field width', df.columns[4]: 'Record Field', df.columns[5]: 'Value'}, inplace=True)
        df.sort_values(by='MTP address', inplace=True)  # Sort by MTP address
        return df

    def save_changes(self):
        with pd.ExcelWriter(self.modified_path, engine='openpyxl', mode='a') as writer:
            self.dataframe.to_excel(writer, sheet_name='MTP Map', index=False)


class ExcelGUI:
    def __init__(self, root):
        self.root = root
        self.root.title('Excel MTP Map Controller')
        self.root.geometry('1000x600')
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=1)

        # Filepath to be used
        self.filepath = r'C:\WORK\Sirius\MTP\Sirius_OTP_Fusemap v23_20241206_Orig.xlsm'
        self.controller = None

        # UI Elements
        button_frame = ttk.Frame(root)
        button_frame.grid(row=0, column=0, padx=10, pady=10, sticky='nw')

        self.connect_btn = ttk.Button(button_frame, text='Connect', command=self.connect)
        self.connect_btn.grid(row=0, column=0, padx=5, pady=5)

        self.generate_btn = ttk.Button(button_frame, text='Generate', command=self.generate)
        self.generate_btn.grid(row=0, column=1, padx=5, pady=5)

        # Add Treeview with scrollbar for the register box
        tree_frame = ttk.Frame(root)
        tree_frame.grid(row=0, column=1, rowspan=3, padx=10, pady=10, sticky='nsew')

        tree_scrollbar = ttk.Scrollbar(tree_frame, orient='vertical')
        self.tree = ttk.Treeview(tree_frame, columns=("Register Name", "Address", "Bit", "Value"), show='headings', yscrollcommand=tree_scrollbar.set)
        for col in ["Register Name", "Address", "Bit", "Value"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)
        self.tree.pack(side='left', fill='both', expand=True)

        tree_scrollbar.config(command=self.tree.yview)
        tree_scrollbar.pack(side='right', fill='y')

        # Add scrollbar to the description text box
        description_frame = ttk.Frame(root)
        description_frame.grid(row=1, column=0, padx=10, pady=5, sticky='nsew')

        description_scrollbar = ttk.Scrollbar(description_frame, orient='vertical')
        self.description_text = tk.Text(description_frame, wrap='word', height=10, width=30, relief='sunken', yscrollcommand=description_scrollbar.set)
        self.description_text.pack(side='left', fill='both', expand=True)

        description_scrollbar.config(command=self.description_text.yview)
        description_scrollbar.pack(side='right', fill='y')

        # Configure resizing behavior
        root.grid_rowconfigure(0, weight=1)
        root.grid_rowconfigure(1, weight=1)
        root.grid_columnconfigure(1, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        description_frame.grid_rowconfigure(0, weight=1)
        description_frame.grid_columnconfigure(0, weight=1)

    def connect(self):
        if self.controller is None:
            try:
                self.controller = ExcelController(self.filepath)
                self.populate_treeview()
                self.populate_description()
                messagebox.showinfo('Connected', f'Excel file connected and copy created at {self.controller.modified_path}.')
            except Exception as e:
                messagebox.showerror('Error', f'Failed to connect: {str(e)}')
        else:
            messagebox.showinfo('Already Connected', 'The Excel file is already connected.')

    def populate_treeview(self):
        if self.controller:
            self.tree.delete(*self.tree.get_children())
            grouped_data = self.controller.dataframe.groupby('MTP address')
            for address, group in grouped_data:
                register_name = group['Record Field'].iloc[0]
                bit = 8  # Set bit value to always be 8
                value = ', '.join(group['Value'].dropna().unique().astype(str))  # Concatenate all unique non-NA values
                self.tree.insert('', 'end', values=(register_name, address, bit, value))

    def populate_description(self):
        if self.controller:
            self.description_text.delete(1.0, tk.END)
            description = self.controller.dataframe.to_string()
            self.description_text.insert(tk.END, description)

    def generate(self):
        messagebox.showinfo('Generate', 'Generate button clicked.')


if __name__ == '__main__':
    root = tk.Tk()
    app = ExcelGUI(root)
    root.mainloop()
