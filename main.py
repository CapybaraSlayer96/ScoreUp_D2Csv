import tkinter as tk
from tkinter import filedialog
import os

from pandas.errors import InvalidIndexError

import data2csv as d2c

# Globals
input_file = ""
tag_value = ""
start_index = 0


def log_message(msg):
    """Append message to the log area in the UI."""
    log_area.insert(tk.END, msg + "\n")
    log_area.see(tk.END)  # auto-scroll


def open_file():
    global input_file
    input_file = filedialog.askopenfilename(
        title="Select a File",
        initialdir="/",
        filetypes=(
            ("Word file", "*.docx"),
            ("Excel file", "*.xlsx"),
            ("All files", "*.*"),
        )
    )
    if input_file:
        log_message(f"File selected: {input_file}")
    else:
        log_message("No file selected.")


def process_file():
    global input_file, tag_value, start_index
    if not input_file:
        log_message("Error: Please open a file first!")
        return

    tag_value = tag_entry.get().strip()
    if not tag_value:
        log_message("Error: Please enter a tag first!")
        return
    if " " in tag_value:
        log_message("Error: Please use '_' instead of spaces in tag!")
        return

    # Lấy starting index
    try:
        start_index = int(start_entry.get().strip())
        if start_index < 1:
            raise ValueError
    except ValueError:
        log_message("Error: Starting index must be a valid integer!")

    # Lấy tên file gốc, đổi extension sang .csv
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    default_name = base_name + ".csv"

    # Save As
    save_path = filedialog.asksaveasfilename(
        title="Save As",
        initialfile=default_name,
        defaultextension=".csv",
        filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
    )

    if not save_path:
        log_message("Processing cancelled (no output file chosen).")
        return

    try:
        # truyền start_index vào document_to_data
        d2c.document_to_data(input_file, save_path, tags=tag_value, start_index=start_index)
        log_message(f"Processed and saved file: {save_path}")
    except Exception as e:
        log_message(f"Processing failed: {e}")


# ---- UI ----
root = tk.Tk()
root.title("ScoreUp 2.0 - Doc2Csv")
root.geometry("800x600")

# Configure grid to expand
root.grid_rowconfigure(4, weight=1)
root.grid_columnconfigure(0, weight=1)

# --- Row 0: Open file ---
open_button = tk.Button(root, text="Open File", width=20, command=open_file)
open_button.grid(row=0, column=0, pady=10)

# --- Row 1: Instruction + Tag input ---
tag_frame = tk.Frame(root)
tag_frame.grid(row=1, column=0, pady=10)

instruction = tk.Label(tag_frame, text="Các từ trong tag cần được ngăn cách bởi dấu '_'", font=("Arial", 12))
instruction.grid(row=0, column=0, columnspan=2, pady=5)

tk.Label(tag_frame, text="Enter tag:").grid(row=1, column=0, padx=5, sticky="e")
tag_entry = tk.Entry(tag_frame, width=20)
tag_entry.grid(row=1, column=1, padx=5, sticky="w")

# --- Row 2: Starting index ---
start_frame = tk.Frame(root)
start_frame.grid(row=2, column=0, pady=10)

tk.Label(start_frame, text="Starting index:").grid(row=0, column=0, padx=5, sticky="e")
start_entry = tk.Entry(start_frame, width=10)
start_entry.insert(0, "1")  # mặc định = 1
start_entry.grid(row=0, column=1, padx=5, sticky="w")

# --- Row 3: Process button ---
process_button = tk.Button(root, text="Process", width=20, command=process_file)
process_button.grid(row=3, column=0, pady=10)

# --- Row 4: Log area ---
log_area = tk.Text(root, wrap="word", font=("Consolas", 11))
log_area.grid(row=4, column=0, sticky="nsew", padx=20, pady=20)

root.mainloop()