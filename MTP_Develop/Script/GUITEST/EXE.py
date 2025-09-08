# -*- coding: utf-8 -*-
"""
Created on Thu Dec 12 16:25:07 2024

@author: CC.Cheng
"""

import subprocess
import os

# Define the main script and additional files
main_script = "/mnt/data/MTP_Alchemist_0.0.19.py"
additional_files = [
    "/mnt/data/APPSTEST_EEPROM_Updata_script.batch",
    "/mnt/data/MTPCodePatch_DATA_32k.bin",
    "/mnt/data/MTPDATA_update_execute.xml",
    "/mnt/data/MTP_DATA_1K.bin",
    "/mnt/data/MTP_DATA_1K_ConfigConverted.bin",
    "/mnt/data/Sirius_OTP_Fusemap v23_20241206_PythonSourceExcel.xlsx",
    "/mnt/data/snorlax.ico"
]

# Prepare the --add-data arguments
add_data_args = []
for file_path in additional_files:
    base_name = os.path.basename(file_path)
    add_data_args.append(f"--add-data={file_path};.")

# Prepare the PyInstaller command
pyinstaller_command = [
    "pyinstaller",
    "--onefile",
    "--windowed",
    f"--icon={additional_files[-1]}",
    *add_data_args,
    main_script
]

# Run the PyInstaller command
try:
    result = subprocess.run(pyinstaller_command, check=True, capture_output=True, text=True)
    output_message = "Executable created successfully."
except subprocess.CalledProcessError as e:
    output_message = f"Error occurred during packaging:\n{e.stderr}"

output_message
