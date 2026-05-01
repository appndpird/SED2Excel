# SED Hyperspectral → Excel Converter

A desktop tool for converting Spectral Evolution / NaturaSpec `.SED` field-spectrometer files into structured Excel workbooks. Designed for batch processing of hyperspectral measurements collected during agricultural field trials and remote-sensing ground-truth campaigns.

The tool produces two workbooks per run:

- **`<name>.xlsx`** — a spectral-only workbook with `ID`, `File_ID`, and one column per wavelength (nm).
- **`metadata_<name>.xlsx`** — a full workbook including all SED header fields (date, time, GPS, instrument, integration, etc.) followed by the wavelength columns.

Both workbooks are styled with frozen ID columns, alternating row shading, and a banded header so they are immediately usable for QA and downstream analysis.

---

## Features

- **Batch conversion** of every `.SED` file in a folder (recursive search included).
- **Two output workbooks** in a single run — spectral-only and full metadata + spectral.
- **Sortable numeric `ID`** extracted from the trailing digits of each filename (e.g. `plot_00042.sed` → `42`), with the full filename preserved as `File_ID`.
- **All standard SED header fields** parsed: instrument, detectors, date/time, temperature, integration, foreoptic, radiometric calibration, units, latitude/longitude, altitude, GPS time, satellites, and more.
- **Self-installing Windows launcher** (`.bat`) that locates Python, creates an isolated virtual environment, installs `openpyxl`, and runs the GUI — no manual setup required.
- **Cross-platform Python script** that runs on Windows, macOS, and Linux for users who already have Python.

---

## Requirements

- **Python 3.8 or later** (Windows, macOS, or Linux). On Windows, the launcher will detect any of `py`, `python3`, or `python` on the system `PATH`.
- **`openpyxl`** Python package. The Windows launcher installs this automatically into an isolated virtual environment; manual users can install it with `pip install openpyxl`.
- **Tkinter** — included with the standard Python installer on Windows and macOS. On Debian / Ubuntu Linux, install it with `sudo apt install python3-tk`.

No internet connection is required at runtime once dependencies are installed.

---

## Installation

### Option A — Windows one-click launcher (recommended)

1. Download `SED_to_Excel_Converter.bat` from this repository (or the latest release).
2. Place it in any folder you have write access to. The first run creates a `sed_excel_env` virtual environment alongside the `.bat` file, so avoid placing it in a read-only location like `C:\Program Files`.
3. Make sure Python 3.8+ is installed. If it is not, download it from [python.org/downloads](https://www.python.org/downloads/) and tick **Add Python to PATH** during installation.
4. Double-click `SED_to_Excel_Converter.bat`. On the first run it will:
   - Locate your Python installation.
   - Create an isolated virtual environment in `sed_excel_env\`.
   - Install `openpyxl` into that environment.
   - Launch the GUI.

   Subsequent runs reuse the existing environment and start the GUI in a couple of seconds.

### Option B — Run the Python script directly

Suitable for macOS, Linux, or any Windows user who prefers to manage their own environment.

```bash
# Clone the repository
git clone https://github.com/<your-username>/sed-to-excel-converter.git
cd sed-to-excel-converter

# (Optional but recommended) create a virtual environment
python -m venv .venv
# Windows:  .venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

# Install the single dependency
pip install openpyxl

# Launch the GUI
python sed_to_excel.py
```

---

## Usage

1. **Launch the application** using either method above. The main window opens with three input fields and a Convert button.

2. **Input Folder** — click *Browse* and select the folder containing your `.SED` files. Subfolders are searched automatically. The status bar reports how many `.SED` files were detected.

3. **Output Folder** — defaults to the input folder. Click *Browse* to choose a different location if you want to keep raw data and outputs separate.

4. **Output File Name** — defaults to `hyperspectral_data`. The two output workbooks will be named:
   - `<name>.xlsx` — spectral data
   - `metadata_<name>.xlsx` — metadata + spectral data

   The preview box updates live as you type.

5. **Click Convert.** Progress is shown in the progress bar and live log. If the output files already exist, you will be prompted before they are overwritten.

6. **Review the results.** A dialog confirms when conversion is complete and reports how many files were processed and where they were saved. Files that could not be parsed (missing data block, corrupt, etc.) are listed in the log and skipped.

### Output structure

**`<name>.xlsx` — Spectral Data sheet**

| ID  | File_ID                | 350    | 351    | …  | 2500   |
| --- | ---------------------- | ------ | ------ | -- | ------ |
| 1   | trial_2024_00001       | 0.0421 | 0.0438 | …  | 0.1832 |
| 2   | trial_2024_00002       | 0.0395 | 0.0412 | …  | 0.1811 |

**`metadata_<name>.xlsx` — Full Data sheet**

| ID | File_ID          | Date       | Time     | Latitude   | Longitude   | Integration | … | 350    | 351    | … |
| -- | ---------------- | ---------- | -------- | ---------- | ----------- | ----------- | - | ------ | ------ | - |
| 1  | trial_2024_00001 | 2024-09-12 | 10:32:14 | -31.952200 | 115.861100  | 100         | … | 0.0421 | 0.0438 | … |

The `ID` column is the trailing numeric portion of the filename and is intended for quick numeric sorting / joining. `File_ID` preserves the full original filename stem so each row can always be traced back to its source.

---

## How it works

1. **Header parsing.** Each `.SED` file is read line by line. Lines before the `Data:` marker are interpreted as `key: value` pairs and matched against a controlled list of SED header fields.
2. **Spectral parsing.** Lines after `Data:` are parsed as whitespace-separated `wavelength value` pairs. The first line after `Data:` (the column header) is skipped.
3. **Wavelength alignment.** All unique integer wavelengths across the input batch are collected and sorted; each row is filled column-by-column, leaving blanks where a wavelength is not present in a particular file.
4. **Excel writing.** The two workbooks are written with `openpyxl`, with styled headers, frozen panes (`C2`), and 4-decimal-place number formatting on spectral cells.
5. **Self-installing launcher (Windows only).** The `.bat` file embeds the Python source as a base85+zlib-compressed string. On launch it locates Python, sets up a venv, installs `openpyxl`, decodes the embedded source to a temporary `.py` file, executes it, and deletes the temporary file when the GUI closes.

---

## Troubleshooting

**“Python not found” when launching the `.bat`.** Install Python 3.8+ from [python.org](https://www.python.org/downloads/) and ensure *Add Python to PATH* is selected during installation. Then re-run the launcher.

**“Failed to install openpyxl.”** The launcher needs internet access for the first run only. Check your connection or proxy settings, then re-launch.

**“No `.SED` files found.”** The tool searches recursively, but only matches files with the extension `.sed` or `.SED`. Confirm your files use that extension.

**Some files are skipped with a parse-failure message.** This usually means the file is missing a `Data:` block or its spectral table is malformed. Open the file in a text editor to confirm. The remaining files in the batch are still processed.

**The GUI does not appear on Linux.** Install Tkinter: `sudo apt install python3-tk` (Debian / Ubuntu) or the equivalent for your distribution.

**Excel reports a “number stored as text” warning on metadata columns.** This is expected — SED metadata values like `Date` and `GPS Time` are stored as strings to preserve their original formatting.

---

## Repository contents

```
.
├── sed_to_excel.py                 # Main GUI application (cross-platform)
├── SED_to_Excel_Converter.bat      # Windows self-installing launcher
├── README.md                       # This file
└── docs/
    └── User_Guide.docx             # Full user guide (Word format)
```

---

## Credits

Developed by **Bipul Neupane, PhD** — Research Scientist, DPIRD Node, [Australian Plant Phenomics Network (APPN)](https://www.plantphenomics.org.au/).
Contact: [bipul.neupane@dpird.wa.gov.au](mailto:bipul.neupane@dpird.wa.gov.au)

Developed with [claude.ai](https://claude.ai).

---

## License

Specify your license here (e.g. MIT, Apache-2.0, or DPIRD internal use). Add a `LICENSE` file at the repository root.
