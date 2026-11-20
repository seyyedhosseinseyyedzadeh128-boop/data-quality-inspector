# Excel/CSV Data Quality Inspector

Inspect one local CSV or XLSX worksheet and create a separate UTF-8 text report. The tool runs offline and does not clean input data.

## Setup and use

With Python installed for your Windows account, open PowerShell in the project folder that contains `app.py` and `requirements.txt`. Create a virtual environment, install the dependencies, and launch the app:

```powershell
py -m venv .venv
& ".\.venv\Scripts\python.exe" -m pip install -r requirements.txt
& ".\.venv\Scripts\python.exe" app.py
```

Enter an input path, CSV settings or worksheet number, and a new `.txt` report path. Existing reports are not overwritten.

## Rules and limitations

CSV defaults to UTF-8 with optional BOM and comma; choose an encoding and single-character delimiter explicitly when needed. All CSV fields are text. Empty fields and XLSX blank cells are missing; literal `NULL`, `NA`, `0`, and whitespace-only text are not. Excel empty strings exposed as blank are missing. Text identifiers retain leading zeros; numeric cells displayed with leading zeros remain numeric. Excel display formatting is not preserved. Formulas are inspected as text and never evaluated. XLSM is unsupported. Worksheet used dimensions define the table, so formatted empty columns may cause header errors. Duplicate rows compare exact types and values, with missing cells equal. Edge Unicode whitespace is counted without changing values. Reports omit column names and cell contents. Data is loaded into memory.

Reports omit input filenames, paths, worksheet names, column names, and cell contents; column counts use positions. They include the file type, CSV settings, and aggregate counts, which can still reveal information about a dataset. Worksheet names are displayed during selection. Expected errors use sanitized messages; unexpected errors can still expose diagnostic details. Protect the report destination and review reports and messages before sharing them.

## Tests

Using the virtual environment above, run synthetic tests only when authorized:

```powershell
& ".\.venv\Scripts\python.exe" -B -m unittest discover -s tests -v
```
