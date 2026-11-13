# TSV to XLSX Converter (GUI)

A small desktop tool built with **Tkinter** that converts tab-separated (`.tsv` / `.txt`) files into Excel workbooks (`.xlsx`). It is designed to handle **rows with inconsistent column counts**, which is where `pandas.read_csv` usually fails.

## Features

- Simple GUI: pick an input TSV file and an output XLSX file with file dialogs.
- Automatically suggests an output path next to the input file (`data.tsv` -> `data.xlsx`).
- Tolerates ragged rows: shorter rows are padded with empty cells up to the widest row.
- Uses the first row as column headers (blank headers become `Column_1`, `Column_2`, ...).
- Reads files as UTF-8 and replaces undecodable bytes instead of crashing.
- Inline status messages plus success/error dialogs.

## Requirements

- Python 3.8 or newer (tested with Python 3.11)
- Packages:

  ```bash
  pip install pandas openpyxl
  ```

  `tkinter` ships with the standard Python installers on Windows and macOS. On Debian/Ubuntu install it with `sudo apt install python3-tk`.

## Usage

Double-click **`run_gui.pyw`** — Windows opens it with `pythonw.exe`, so the app starts **without a black console window**. It is just a two-line launcher for the main script.

From a terminal (console window stays visible):

```bash
python tsv_to_xlsx.py
```

1. Click **Browse...** next to *Input TSV File* and choose your `.tsv` (or `.txt`) file.
2. The *Output XLSX File* field is filled in automatically; change it with **Browse...** if needed.
3. Click **Convert TSV to XLSX**.
4. A dialog reports the number of rows and columns written, and the status bar shows the saved location.

### Why there is no console window

A normal `.py` file is tied to `python.exe`, which always opens a terminal. Double-clicking a `.pyw` file uses `pythonw.exe` instead, which has no console. If you build an EXE, pass `--noconsole` (next section) or the same terminal will reappear.

## Building a windowed EXE

```bash
pip install pyinstaller
pyinstaller --onefile --noconsole --name "TSV to XLSX Converter" tsv_to_xlsx.py
```

The result is `dist/TSV to XLSX Converter.exe`. Without `--noconsole` the EXE opens a command prompt behind the window; rebuild with the flag to remove it. `--onefile` bundles everything into one file but makes startup a little slower (the bundle unpacks to a temp folder on every launch).

## High-DPI screens (blurry text)

The app declares **per-monitor DPI awareness** and scales its geometry and padding to the display, so text stays sharp and the layout keeps its proportions at 125%, 150% or 200% scaling. Without this, Windows renders the window at 100% and stretches the bitmap, which is what makes Tkinter apps look fuzzy on laptops.

If an already-built EXE still looks blurry, rebuild it: DPI awareness has to be requested by the process at startup, so the old binary keeps the blurry behaviour.

## How it works

1. The whole file is read with Python's built-in `csv` module using `delimiter="\t"`.
2. The maximum number of fields across all rows is computed.
3. Every row is padded with empty strings so all rows have equal length.
4. The result is loaded into a `pandas.DataFrame`; the first row becomes the header.
5. The DataFrame is written to `.xlsx` with the `openpyxl` engine (`index=False`).

## Project structure

```
.
|-- tsv_to_xlsx.py    # the application (single file)
|-- run_gui.pyw       # double-click launcher: runs the app without a console
`-- README.md         # this file
```

## Known limitations

- **The first row is always treated as a header.** Files without a header row will lose their first data line.
- **No delimiter auto-detection.** Only tab-separated input is supported (the `.txt` filter still assumes tabs).
- **Excel files are overwritten without warning** when the target path already exists.
- All processing happens on the UI thread, so very large files make the window appear frozen until conversion finishes.
- Duplicate header names are written as-is; pandas/openpyxl may complain about repeated column labels in some versions.
- No `.xls` (legacy) or `.csv` output, and no csv-to-tsv direction.

## Troubleshooting

| Symptom | Fix |
| --- | --- |
| `ModuleNotFoundError: No module named 'pandas'` | Run `pip install pandas openpyxl`. |
| `ModuleNotFoundError: No module named 'tkinter'` | Install the Tk bindings for your Python distribution. |
| All data appears in one column | The file is not actually tab-separated (comma or semicolon delimited). |
| First line of data is missing | The file has no header row; add one or edit the script to skip header handling. |
| `PermissionError` when saving | Close the workbook in Excel and choose a writable location. |
| A black console window opens with the app | Double-click `run_gui.pyw` instead of the `.py`, or rebuild the EXE with `--noconsole`. |
| Window text looks blurry | Fixed in this version (per-monitor DPI awareness). Rebuild any older EXE to pick it up. |

## License

No license file is included. Add one before redistributing.
