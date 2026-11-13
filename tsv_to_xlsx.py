#!/usr/bin/env python3
"""
TSV to XLSX Converter - GUI Version (Handles Inconsistent Columns)
Browse and select TSV files, convert them to XLSX with error handling
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import csv
import sys


def enable_high_dpi():
    """Render the window sharply on Windows displays scaled above 100%.

    Without this, Windows draws the app at 100% and then stretches it, which is
    what makes Tkinter windows look blurry on laptops set to 125% or 150%.
    """
    if sys.platform != "win32":
        return

    try:
        import ctypes
    except ImportError:
        return

    # Per-monitor DPI awareness (Windows 8.1 and newer)
    try:
        if ctypes.windll.shcore.SetProcessDpiAwareness(2) == 0:
            return
    except (AttributeError, OSError):
        pass

    # Fallback for older Windows versions
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except (AttributeError, OSError):
        pass


class TSVtoXLSXConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("TSV to XLSX Converter")

        # Sizes below are written for a 100% display and scaled from here, so
        # the layout keeps its proportions on high-DPI screens.
        self.scale = self.root.winfo_fpixels("1i") / 96.0

        self.root.geometry(f"{self.px(600)}x{self.px(400)}")
        self.root.resizable(False, False)

        # Variables
        self.tsv_file_path = tk.StringVar()
        self.xlsx_file_path = tk.StringVar()

        # Create UI
        self.create_widgets()

    def px(self, value):
        """Convert a design pixel value to real pixels for this display."""
        return int(round(value * self.scale))

    def create_widgets(self):
        # Title
        title_label = tk.Label(
            self.root,
            text="TSV to XLSX Converter",
            font=("Segoe UI", 16, "bold"),
            pady=self.px(20)
        )
        title_label.pack()

        # Input file section
        input_frame = tk.Frame(self.root, padx=self.px(20), pady=self.px(10))
        input_frame.pack(fill="x")

        tk.Label(input_frame, text="Input TSV File:", font=("Segoe UI", 10, "bold")).pack(anchor="w")

        input_path_frame = tk.Frame(input_frame)
        input_path_frame.pack(fill="x", pady=self.px(5))

        tk.Entry(
            input_path_frame,
            textvariable=self.tsv_file_path,
            width=50,
            state="readonly"
        ).pack(side="left", padx=(0, self.px(10)))

        tk.Button(
            input_path_frame,
            text="Browse...",
            command=self.browse_tsv_file,
            width=12
        ).pack(side="left")

        # Output file section
        output_frame = tk.Frame(self.root, padx=self.px(20), pady=self.px(10))
        output_frame.pack(fill="x")

        tk.Label(output_frame, text="Output XLSX File:", font=("Segoe UI", 10, "bold")).pack(anchor="w")

        output_path_frame = tk.Frame(output_frame)
        output_path_frame.pack(fill="x", pady=self.px(5))

        tk.Entry(
            output_path_frame,
            textvariable=self.xlsx_file_path,
            width=50,
            state="readonly"
        ).pack(side="left", padx=(0, self.px(10)))

        tk.Button(
            output_path_frame,
            text="Browse...",
            command=self.browse_xlsx_file,
            width=12
        ).pack(side="left")

        # Convert button
        convert_btn = tk.Button(
            self.root,
            text="Convert TSV to XLSX",
            command=self.convert_file,
            bg="#4CAF50",
            fg="white",
            font=("Segoe UI", 12, "bold"),
            pady=self.px(10),
            width=20
        )
        convert_btn.pack(pady=self.px(30))

        # Status label
        self.status_label = tk.Label(
            self.root,
            text="Select a TSV file to begin",
            font=("Segoe UI", 9),
            fg="gray"
        )
        self.status_label.pack(pady=self.px(10))

    def browse_tsv_file(self):
        """Open file dialog to select TSV file"""
        filename = filedialog.askopenfilename(
            title="Select TSV File",
            filetypes=[
                ("TSV files", "*.tsv"),
                ("Tab-separated files", "*.txt"),
                ("All files", "*.*")
            ]
        )

        if filename:
            self.tsv_file_path.set(filename)
            # Auto-suggest output filename
            base_name = os.path.splitext(filename)[0]
            suggested_output = f"{base_name}.xlsx"
            self.xlsx_file_path.set(suggested_output)
            self.status_label.config(text="Ready to convert", fg="blue")

    def browse_xlsx_file(self):
        """Open file dialog to select output XLSX file location"""
        filename = filedialog.asksaveasfilename(
            title="Save XLSX File As",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("All files", "*.*")
            ]
        )

        if filename:
            self.xlsx_file_path.set(filename)

    def convert_file(self):
        """Perform the TSV to XLSX conversion with error handling for inconsistent columns"""
        tsv_path = self.tsv_file_path.get()
        xlsx_path = self.xlsx_file_path.get()

        # Validation
        if not tsv_path:
            messagebox.showerror("Error", "Please select an input TSV file")
            return

        if not xlsx_path:
            messagebox.showerror("Error", "Please select an output XLSX file location")
            return

        if not os.path.exists(tsv_path):
            messagebox.showerror("Error", "Input TSV file does not exist")
            return

        try:
            # Update status
            self.status_label.config(text="Converting... Please wait", fg="orange")
            self.root.update()

            # Method 1: Read with Python's csv module to handle inconsistent columns
            with open(tsv_path, 'r', encoding='utf-8', errors='replace') as file:
                reader = csv.reader(file, delimiter='\t')
                rows = list(reader)

            # Find the maximum number of columns
            max_cols = max(len(row) for row in rows) if rows else 0

            # Pad rows to have consistent column count
            padded_rows = [row + [''] * (max_cols - len(row)) for row in rows]

            # Convert to DataFrame
            df = pd.DataFrame(padded_rows)

            # If first row looks like headers, use it
            if len(padded_rows) > 0:
                df.columns = [f"Column_{i+1}" if not padded_rows[0][i] else padded_rows[0][i]
                             for i in range(max_cols)]
                df = df[1:]  # Remove the header row from data

            # Write to XLSX
            df.to_excel(xlsx_path, index=False, engine='openpyxl')

            # Success message
            self.status_label.config(
                text=f"✓ Conversion successful! ({len(df)} rows, {len(df.columns)} columns)",
                fg="green"
            )

            messagebox.showinfo(
                "Success",
                f"File converted successfully!\n\nRows: {len(df)}\nColumns: {len(df.columns)}\n\nSaved to:\n{xlsx_path}"
            )

        except Exception as e:
            self.status_label.config(text="✗ Conversion failed", fg="red")
            messagebox.showerror("Conversion Error", f"An error occurred:\n\n{str(e)}")

def main():
    enable_high_dpi()
    root = tk.Tk()
    TSVtoXLSXConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()
