# -*- coding: utf-8 -*-
"""
Created on Mon Dec  2 15:55:12 2024

@author: CC.Cheng
"""
import os
import time
import keyboard  # For Esc key detection


def load_binary_file(filepath):
    """Load a binary file and return its contents as a byte array."""
    with open(filepath, 'rb') as file:
        return file.read()


def compare_files_byte_difference(reference, current):
    """Compare two binary files byte-by-byte and calculate the number of differing bytes and same bytes."""
    byte_difference = 0
    byte_same = 0
    min_length = min(len(reference), len(current))
    max_length = max(len(reference), len(current))

    # Compare the overlapping portion
    for i in range(min_length):
        if reference[i] != current[i]:
            byte_difference += 1
        else:
            byte_same += 1

    # Account for extra bytes in the longer file (these bytes are considered "different")
    byte_difference += max_length - min_length

    return byte_same, byte_difference


def print_table_header():
    """Print the table header for monitoring results."""
    print(f"{'Files_Compared':<15}{'Consecutive_Same':<20}{'Byte_Same':<15}{'Byte_Difference':<20}")
    print("-" * 70)


def print_table_row(files_compared, consecutive_same_count, byte_same, byte_difference):
    """Print a row of the monitoring table."""
    print(f"{files_compared:<15}{consecutive_same_count:<20}{byte_same:<15}{byte_difference:<20}")


def monitor_and_compare(filepath, interval=15, break_on_no_difference=True, max_consecutive_same=3):
    """
    Monitor and compare binary files in a loop every 15 seconds.

    If no reference file exists, the current file is used as the reference.
    After comparison, the current file becomes the new reference.

    Parameters:
    - filepath (str): Path to the binary file to monitor.
    - interval (int): Time interval in seconds between comparisons.
    - break_on_no_difference (bool): Whether to stop monitoring when files are consistently the same.
    - max_consecutive_same (int): Number of consecutive "all bytes same" comparisons required to break the loop.
    """
    if not os.path.exists(filepath):
        print("File does not exist.")
        return

    print(f"Monitoring for changes every {interval} seconds. Press Esc to exit.")
    reference_data = None
    files_compared = 0
    consecutive_same_count = 0

    print_table_header()

    while True:
        # Check for Esc key press to exit
        if keyboard.is_pressed("esc"):
            print("Esc key pressed. Exiting monitoring.")
            break

        time.sleep(interval)  # Wait for the interval

        # Load the current file
        try:
            current_data = load_binary_file(filepath)
        except FileNotFoundError:
            print("File was deleted or moved. Stopping monitoring.")
            break

        if reference_data is None:
            # If no reference file exists, load the current file as the reference
            reference_data = current_data
            print("Loaded initial reference file.")
        else:
            # Compare the reference file with the current file
            byte_same, byte_difference = compare_files_byte_difference(reference_data, current_data)

            if byte_difference == 0:
                consecutive_same_count += 1
            else:
                consecutive_same_count = 0  # Reset the counter if differences are found
                files_compared += 1  # Increment files compared only when differences exist

            # Print the comparison results in table format
            print_table_row(files_compared, consecutive_same_count, byte_same, byte_difference)

            # Check if monitoring should stop
            if break_on_no_difference and consecutive_same_count >= max_consecutive_same:
                print(f"Files are identical for {max_consecutive_same} consecutive comparisons. Stopping monitoring.")
                break

            # Update the reference file
            reference_data = current_data


# Run the monitor and compare function
file_to_monitor = r"C:\WORK\Sirius\validation_summary_11052024\test.bin"  # Replace with the actual file path
monitor_and_compare(file_to_monitor, interval=16.275, break_on_no_difference=True, max_consecutive_same=3)

