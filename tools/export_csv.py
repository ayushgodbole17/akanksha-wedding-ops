"""Export each tab of the master workbook to CSV, formatted the way Google
Sheets' CSV link formats them, so the phone page can be tested locally.

Usage:  python tools/export_csv.py wedding-master.xlsx test-data
Then:   python -m http.server 8000   and open  http://localhost:8000/?src=test-data
"""
import csv
import os
import sys
from datetime import date, datetime, time

from openpyxl import load_workbook

src, out = sys.argv[1], sys.argv[2]
os.makedirs(out, exist_ok=True)

def fmt(v):
    if v is None:
        return ""
    if isinstance(v, datetime) and v.time() == time(0):
        v = v.date()
    if isinstance(v, date) and not isinstance(v, datetime):
        return f"{v.day} {v:%b}"
    if isinstance(v, time):
        return f"{(v.hour % 12) or 12}:{v.minute:02d} {'AM' if v.hour < 12 else 'PM'}"
    if isinstance(v, str) and v.startswith("="):
        return ""  # formulas aren't computed outside Excel/Sheets
    return str(v)

wb = load_workbook(src)
for ws in wb.worksheets:
    if ws.title == "How to use":
        continue
    with open(os.path.join(out, f"{ws.title}.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, quoting=csv.QUOTE_ALL)
        for row in ws.iter_rows(values_only=True):
            w.writerow([fmt(c) for c in row])
    print(f"{ws.title}: {ws.max_row - 1} rows")
