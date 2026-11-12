"""Build a test workbook (data/Migration Tracker.xlsx) from the Endpoint Tracker CSV export.

    python seed_workbook.py "path/to/Endpoint Tracker.csv"

Creates three sheets, each holding an Excel Table:
  Endpoint Tracker  (table "Tracker")     - one row per staff endpoint, + Last Updated By/At
  Analysts          (table "Analysts")    - roster: Name, Email, Team, Active, Role
  Activity Log      (table "ActivityLog") - every save made through the app

The app's Workbook page adds the same sheets and columns to a real workbook.
"""
import csv
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

from backend import col_letter
from logic import DEFAULT_ROSTER as ROSTER, LOG_HEADERS, TRACKING_COLS



def add_sheet(wb, title, table, headers, rows, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = title
    ws.append(headers)
    for r in rows:
        ws.append(r)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
    ws.freeze_panes = "A2"
    last = max(ws.max_row, 2)
    if ws.max_row == 1:
        ws.append([""] * len(headers))  # Excel tables need at least one data row
    t = Table(displayName=table, ref=f"A1:{col_letter(len(headers) - 1)}{last}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(t)
    for i, h in enumerate(headers):
        ws.column_dimensions[col_letter(i)].width = min(max(len(h), 12), 28)
    return ws


def main(src, out="data/Migration Tracker.xlsx"):
    with open(src, encoding="utf-8-sig") as f:
        rows = [r for r in csv.reader(f)]
    headers, body = rows[0], [r for r in rows[1:] if r and not r[0].startswith("SAMPLE")]
    headers = headers + TRACKING_COLS
    body = [r + [""] * (len(headers) - len(r)) for r in body]

    wb = Workbook()
    add_sheet(wb, "Endpoint Tracker", "Tracker", headers, body, first=True)
    add_sheet(wb, "Analysts", "Analysts", ["Name", "Email", "Team", "Active", "Role"], ROSTER)
    add_sheet(wb, "Activity Log", "ActivityLog", LOG_HEADERS, [])
    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb.save(out)
    print(f"Wrote {out}: {len(body)} endpoints, {len(ROSTER)} analysts")


if __name__ == "__main__":
    main(*sys.argv[1:])
