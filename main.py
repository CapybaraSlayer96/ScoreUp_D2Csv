import tkinter as tk
from tkinter import filedialog
import shutil
import os
import data2Csv as d2c

# Globals
input_file = ""
output_file = "temp_output.csv"


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
    global input_file
    if not input_file:
        log_message("Error: Please open a file first!")
        return

    save_path = filedialog.asksaveasfilename(
        title="Save As",
        initialfile="processed_output.csv",
        defaultextension=".csv",
        filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
    )
    if not save_path:
        log_message("Processing cancelled (no output file chosen).")
        return

    try:
        d2c.document_to_data(input_file, save_path)
        log_message(f"Processed and saved file: {save_path}")
    except Exception as e:
        log_message(f"Processing failed: {e}")


def export_file():
    global output_file
    if not os.path.exists(output_file):
        log_message("Error: No processed file to export! Run Process first.")
        return

    save_path = filedialog.asksaveasfilename(
        title="Save As",
        initialfile=os.path.basename(output_file),
        defaultextension=".csv",
        filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
    )
    if save_path:
        shutil.copy(output_file, save_path)
        log_message(f"File saved as: {save_path}")
    else:
        log_message("Export cancelled.")


# ---- UI ----
root = tk.Tk()
root.title("ScoreUp 2.0 - Doc2Csv")
root.geometry("800x600")  # full window size

# Configure grid to expand
root.grid_rowconfigure(0, weight=1)
root.grid_rowconfigure(1, weight=3)
root.grid_columnconfigure(0, weight=1)

# Frame for buttons (centered horizontally)
button_frame = tk.Frame(root)
button_frame.grid(row=0, column=0, pady=20)

open_button = tk.Button(button_frame, text="Open File", width=15, command=open_file)
process_button = tk.Button(button_frame, text="Process", width=15, command=process_file)
export_button = tk.Button(button_frame, text="Export File", width=15, command=export_file)

open_button.pack(side="left", padx=10)
process_button.pack(side="left", padx=10)
export_button.pack(side="left", padx=10)

# Log area (fills remaining space)
log_area = tk.Text(root, wrap="word", font=("Consolas", 11))
log_area.grid(row=1, column=0, sticky="nsew", padx=20, pady=20)

root.mainloop()
