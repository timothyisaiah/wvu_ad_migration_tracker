"""Shared helpers and the plain openpyxl backend for the migration tracker.

The backends (here and in excel_file.py) expose the same small interface:

    read_sheet(sheet)                -> pandas.DataFrame (+ "_row" = Excel row number)
    read_row(sheet, row, headers)    -> dict of current values for one row
    write_cells(sheet, updates)      -> updates = [(excel_row, {column: value}), ...]
    append_rows(sheet, table, rows)  -> rows = [[...], ...] appended to an Excel Table

Cells that hold a formula are never overwritten, so any formula columns in the
live workbook keep working.
"""
from __future__ import annotations

import re
import threading
from datetime import datetime

import pandas as pd

def col_letter(idx: int) -> str:
    """0-based column index -> Excel letters (0 -> A, 26 -> AA)."""
    s = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def clean(v):
    """Normalise a raw cell value to a plain string for the app."""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d") if v.time() == datetime.min.time() else v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def frame(values: list[list], first_row: int = 1) -> pd.DataFrame:
    if not values:
        return pd.DataFrame()
    headers = [clean(h) for h in values[0]]
    body = [[clean(c) for c in r] + [""] * (len(headers) - len(r)) for r in values[1:]]
    df = pd.DataFrame(body, columns=headers)
    df["_row"] = range(first_row + 1, first_row + 1 + len(df))
    # drop fully empty rows (e.g. blank space at the bottom of a used range)
    keep = df.drop(columns="_row").apply(lambda r: any(x != "" for x in r), axis=1)
    return df[keep].reset_index(drop=True)


# --------------------------------------------------------------------------- local
class LocalExcelBackend:
    """Reads/writes a local .xlsx with openpyxl. For testing and offline demos."""

    _lock = threading.Lock()

    def __init__(self, path: str):
        self.path = path

    def _open(self, data_only=False):
        from openpyxl import load_workbook
        return load_workbook(self.path, data_only=data_only)

    def read_sheet(self, sheet: str) -> pd.DataFrame:
        with self._lock:
            ws = self._open(data_only=True)[sheet]
            return frame([list(r) for r in ws.iter_rows(values_only=True)])

    def read_row(self, sheet: str, row: int, headers: list[str]) -> dict:
        with self._lock:
            ws = self._open(data_only=True)[sheet]
            hdr = [clean(c.value) for c in ws[1]]
            vals = [clean(c.value) for c in ws[row]]
            return {h: (vals[i] if i < len(vals) else "") for i, h in enumerate(hdr)}

    def write_cells(self, sheet: str, updates: list[tuple[int, dict]], headers=None):
        with self._lock:
            wb = self._open()
            ws = wb[sheet]
            hdr = {clean(c.value): i + 1 for i, c in enumerate(ws[1])}
            for row, values in updates:
                for col, v in values.items():
                    cell = ws.cell(row=row, column=hdr[col])
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        continue  # never overwrite formulas
                    cell.value = v
            self._save(wb)

    def append_rows(self, sheet: str, table: str, rows: list[list]):
        with self._lock:
            wb = self._open()
            ws = wb[sheet]
            for r in rows:
                ws.append(r)
            if table in ws.tables:
                t = ws.tables[table]
                start, end = t.ref.split(":")
                end_col = re.match(r"[A-Z]+", end).group()
                t.ref = f"{start}:{end_col}{ws.max_row}"
            self._save(wb)

    def _save(self, wb):
        try:
            wb.save(self.path)
        except PermissionError as e:
            raise RuntimeError(
                "The workbook is open in Excel on this computer, so it can't be saved. Close it and try again."
            ) from e
