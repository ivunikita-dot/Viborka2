# Excel Contract Extractor

This repository contains a single Python script `script.py` that:

- Opens a dialog to select a source Excel file (FIO in column A, file paths in column E).
- For each row, searches the referenced files for the contract signing date (year >= 2022) and
  writes the year into column D ("Год контракта").
- Searches for the regional payment amount (one of 400, 400000, 1300000, 2650000, 3550000)
  and writes it into column F ("Сумма выплаты") or "пусто" if not found.
- Shows a progress window with cancel button.
- Saves a single result Excel file; does not modify the source file.

Requirements
------------
- Python 3.8+
- pandas
- openpyxl

Packaging to .exe with PyInstaller
----------------------------------
Install dependencies and pyinstaller:

```
pip install -r requirements.txt
pip install pyinstaller
```

Build:

```
pyinstaller --onefile --windowed script.py
```

GitHub Actions
--------------
A sample workflow is included at `.github/workflows/build.yml` to build the .exe on Windows runners.
