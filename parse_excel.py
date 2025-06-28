"""
Parse AutoCAD Excel output to generate label files.
- wire_labels.txt: counts of numeric wire numbers.
- cable_labels.txt: entries starting with a letter followed by digits.
- device_labels.txt: user-selected portion of device text.
"""

import re
import pandas as pd
from collections import Counter
from pathlib import Path


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Parse Excel file for wire, cable and device labels")
    parser.add_argument("excel_file", help="Path to Excel file")
    parser.add_argument("--device-column", type=int, default=None, help="Column index (0-based) that contains device text")
    parser.add_argument("--device-regex", default=None, help="Regex with one capturing group for device label")
    args = parser.parse_args()

    df = pd.read_excel(args.excel_file, header=None)
    df = df.fillna('')
    cells = df.applymap(str).values.ravel()

    wire_pattern = re.compile(r'^\d+$')
    cable_pattern = re.compile(r'^[A-Za-z]+\d+')

    wire_counts = Counter()
    cable_labels = set()

    for cell in cells:
        s = cell.strip()
        if wire_pattern.match(s):
            wire_counts[s] += 1
        elif cable_pattern.match(s):
            cable_labels.add(s)

    # Write wire labels with counts
    with open('wire_labels.txt', 'w') as f:
        for label, count in sorted(wire_counts.items()):
            f.write(f"{label},{count}\n")

    # Write cable labels
    with open('cable_labels.txt', 'w') as f:
        for label in sorted(cable_labels):
            f.write(f"{label}\n")

    # Device labels
    if args.device_column is not None:
        if args.device_column < 0 or args.device_column >= df.shape[1]:
            raise SystemExit("Device column index out of range")
        series = df.iloc[:, args.device_column].astype(str)
        pattern = None
        if args.device_regex:
            pattern = re.compile(args.device_regex)
        device_labels = []
        for item in series:
            item = item.strip()
            if not item:
                continue
            if pattern:
                m = pattern.search(item)
                if m:
                    device_labels.append(m.group(1))
            else:
                device_labels.append(item)
        with open('device_labels.txt', 'w') as f:
            for label in device_labels:
                f.write(f"{label}\n")
    else:
        print("No device column specified, skipping device labels")


if __name__ == "__main__":
    main()
