# Auto-Cad-Excel-Brady-Parser

This repository contains a simple script for processing Excel exports from AutoCAD
(or similar tools) and generating three files for label printing:

- `wire_labels.txt` – counts of wire numbers.
- `cable_labels.txt` – cable identifiers that start with a letter followed by digits.
- `device_labels.txt` – device labels based on a user selected column and regex.

## Usage

```
python parse_excel.py <excel_file> --device-column 2 --device-regex "(\\w+)"
```

Arguments:

- `excel_file` – Path to the Excel document.
- `--device-column` – Index (0-based) of the column containing device text.
- `--device-regex` – Optional regex with one capturing group for the portion of
  the device text to include in `device_labels.txt`. If omitted, the full column
  contents are used.

The script writes the three output files in the current directory.
