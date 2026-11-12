"""Work on the shared workbook as a file on this PC (kept in sync by OneDrive).

Two engines:

ExcelAppBackend  (recommended, Windows + Excel installed)
    Drives a hidden copy of Microsoft Excel, so the file is edited by Excel itself:
    charts, PivotTables, formatting and formulas are untouched. When the file sits in a
    OneDrive/SharePoint-synced folder and Office is signed in, Excel opens it as a cloud
    document with AutoSave, so each save is merged cell-by-cell with what colleagues are
    doing in Excel Online - the same co-authoring Excel uses between people.
    If co-authoring isn't available, the file is opened only for each read/write and closed
    straight after, so OneDrive is never blocked from syncing it.

OpenpyxlFileBackend  (fallback, no Excel needed)
    Rewrites the whole .xlsx with openpyxl. That would strip charts, PivotTables, slicers and
    pictures, so it refuses to write to a workbook that contains any of them. Use it only for
    a data-only workbook that nobody else edits at the same time.
"""
from __future__ import annotations

import atexit
import concurrent.futures
import os
import queue
import re
import threading
import zipfile
from datetime import datetime
from pathlib import Path

from backend import LocalExcelBackend, clean, col_letter, frame

RICH_PARTS = {"xl/charts/": "charts", "xl/pivotTables/": "PivotTables", "xl/slicers/": "slicers",
              "xl/timelines/": "timelines", "xl/media/": "pictures", "xl/drawings/vmlDrawing": "form controls"}
_NUMERIC_TEXT = re.compile(r"^[+-]?\d[\d.]*([eE][+-]?\d+)?$")
_DATE_TEXT = re.compile(r"^\d{4}-\d{2}-\d{2}( \d{2}:\d{2}(:\d{2})?)?$")


def rich_content(path: str) -> list[str]:
    """Things openpyxl would destroy if it rewrote this workbook."""
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
    except (zipfile.BadZipFile, OSError):
        return []
    return sorted({label for prefix, label in RICH_PARTS.items() if any(n.startswith(prefix) for n in names)})


def excel_available() -> bool:
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    import winreg
    try:
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application")
        return True
    except OSError:
        return False


def onedrive_roots() -> list[Path]:
    """OneDrive folders on this PC, plus synced SharePoint libraries (e.g. 'C:/Users/x/World Vision')."""
    roots = []
    for k in ("OneDriveCommercial", "OneDrive", "OneDriveConsumer"):
        v = os.environ.get(k)
        if v and Path(v).exists() and Path(v) not in roots:
            roots.append(Path(v))
    com = os.environ.get("OneDriveCommercial")
    if com and " - " in Path(com).name:
        org = Path(com).parent / Path(com).name.split(" - ", 1)[1]
        if org.exists():
            roots.append(org)
    return roots


def box_roots() -> list[Path]:
    """Box Drive folder(s) on this PC (normally C:/Users/<you>/Box)."""
    home = Path(os.environ.get("USERPROFILE", Path.home()))
    return [p for p in (home / "Box", home / "Box Sync") if p.exists()]


def sync_roots() -> list[tuple[str, Path]]:
    return [("OneDrive", r) for r in onedrive_roots()] + [("Box", r) for r in box_roots()]


def synced_by(path: str) -> str | None:
    """Which sync service (OneDrive / Box) keeps this file in sync, if any."""
    p = str(Path(path)).lower()
    for name, root in sync_roots():
        if p.startswith(str(root).lower() + os.sep) or p == str(root).lower():
            return name
    return None


def find_workbooks(term: str = "", limit: int = 60) -> list[str]:
    """Search OneDrive/SharePoint folders for .xlsx files whose name contains `term`.
    Only lists names, so cloud-only (Files On-Demand) files are not downloaded."""
    out, term = [], term.lower()
    for _, root in sync_roots():
        for dirpath, dirnames, files in os.walk(root):
            dirnames[:] = [d for d in dirnames if not d.startswith(".") and d not in ("AppData", "node_modules")]
            if dirpath.count(os.sep) - str(root).count(os.sep) > 6:
                dirnames[:] = []
            for f in files:
                if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$") and term in f.lower():
                    out.append(os.path.join(dirpath, f))
                    if len(out) >= limit:
                        return out
    return out


def _cell_value(v):
    """Make text that Excel would mangle (e.g. serial '12E45' -> number) stay text."""
    if isinstance(v, str) and _NUMERIC_TEXT.match(v) and not _DATE_TEXT.match(v) and len(v) > 6:
        return "'" + v
    return v


# --------------------------------------------------------------------------- Excel (COM)
class _ComThread:
    """Excel automation must stay on one thread; every call is queued onto it."""

    def __init__(self):
        self.q: queue.Queue = queue.Queue()
        threading.Thread(target=self._run, daemon=True, name="excel-com").start()

    def _run(self):
        import pythoncom
        pythoncom.CoInitialize()
        while True:
            fn, args, fut = self.q.get()
            try:
                fut.set_result(fn(*args))
            except BaseException as e:  # noqa: BLE001
                fut.set_exception(e)

    def call(self, fn, *args, timeout=180):
        fut = concurrent.futures.Future()
        self.q.put((fn, args, fut))
        return fut.result(timeout=timeout)


class ExcelAppBackend:
    def __init__(self, path: str):
        self.path = str(Path(path).resolve())
        self.com = _ComThread()
        self.xl = None
        self.wb = None
        self.cloud = False
        atexit.register(lambda: self.com.call(self._shutdown, timeout=30))

    # -- runs on the COM thread -------------------------------------------------
    def _excel(self):
        if self.xl is not None:
            try:
                _ = self.xl.Version
                return self.xl
            except Exception:  # Excel was closed or crashed
                self.xl, self.wb = None, None
        import win32com.client
        self.xl = win32com.client.DispatchEx("Excel.Application")  # private, hidden instance
        self.xl.Visible = False
        self.xl.DisplayAlerts = False
        self.xl.ScreenUpdating = False
        self.xl.AskToUpdateLinks = False
        return self.xl

    def _open(self):
        if self.wb is not None:
            try:
                _ = self.wb.Name
                return self.wb
            except Exception:
                self.wb = None
        if not Path(self.path).exists():
            raise FileNotFoundError(f"Workbook not found: {self.path}")
        try:
            wb = self._excel().Workbooks.Open(self.path, UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True)
        except Exception as e:
            raise RuntimeError("Excel couldn't open the workbook. Check it opens normally in Excel on this PC, "
                               "isn't still downloading from OneDrive, and isn't password-protected.") from e
        if wb.ReadOnly:
            wb.Close(SaveChanges=False)
            raise RuntimeError("Excel could only open the workbook read-only. It is probably open and locked on this PC "
                               "(close it in desktop Excel here — colleagues can keep using Excel Online).")
        try:
            if not wb.AutoSaveOn:
                wb.AutoSaveOn = True  # only succeeds for OneDrive/SharePoint cloud documents
            self.cloud = bool(wb.AutoSaveOn)
        except Exception:
            self.cloud = False
        self.wb = wb
        return wb

    def _done(self, saved: bool):
        """After each operation: save; close unless Excel is co-authoring the cloud copy."""
        if self.wb is None:
            return
        if saved:
            self.wb.Save()
        if not self.cloud:
            self.wb.Close(SaveChanges=False)
            self.wb = None

    def _shutdown(self):
        try:
            if self.wb is not None:
                self.wb.Close(SaveChanges=True)
        except Exception:
            pass
        try:
            if self.xl is not None:
                self.xl.Quit()
        except Exception:
            pass
        self.wb = self.xl = None

    def _ws(self, sheet):
        try:
            return self._open().Worksheets(sheet)
        except Exception as e:
            raise RuntimeError(f"Sheet '{sheet}' not found in the workbook.") from e

    def _read(self, sheet):
        try:
            ur = self._ws(sheet).UsedRange
            vals = ur.Value
            first = ur.Row
            if not isinstance(vals, tuple):
                vals = ((vals,),)
            return frame([list(r) for r in vals], first)
        finally:
            self._done(saved=False)

    def _read_row(self, sheet, row, headers):
        try:
            v = self._ws(sheet).Range(f"A{row}:{col_letter(len(headers) - 1)}{row}").Value[0]
            return {h: clean(x) for h, x in zip(headers, v)}
        finally:
            self._done(saved=False)

    def _header_map(self, ws):
        ur = ws.UsedRange
        hdr = ws.Range(ws.Cells(ur.Row, 1), ws.Cells(ur.Row, ur.Column + ur.Columns.Count - 1)).Value[0]
        return {clean(h): i + 1 for i, h in enumerate(hdr) if clean(h)}

    def _write(self, sheet, updates):
        ok = False
        try:
            ws = self._ws(sheet)
            hdr = self._header_map(ws)
            for row, values in updates:
                for col, v in values.items():
                    if col not in hdr:
                        raise RuntimeError(f"Column '{col}' is missing from '{sheet}'. Run Admin → Workbook → Set up.")
                    c = ws.Cells(row, hdr[col])
                    if c.HasFormula:
                        continue  # never overwrite formulas
                    c.Value = _cell_value(v)
            ok = True
        finally:
            self._done(saved=ok)

    def _append(self, sheet, table, rows):
        ok = False
        try:
            ws = self._ws(sheet)
            lo = ws.ListObjects(table)
            for r in rows:
                body = lo.DataBodyRange
                # reuse the single blank placeholder row a new table starts with
                if body is not None and lo.ListRows.Count == 1 and all(x in (None, "") for x in body.Value[0]):
                    target = body
                else:
                    target = lo.ListRows.Add().Range
                target.Value = (tuple(_cell_value(v) for v in r),)
            ok = True
        finally:
            self._done(saved=ok)

    def _info(self):
        try:
            wb = self._open()
            return {"sheets": [ws.Name for ws in wb.Worksheets], "cloud": self.cloud,
                    "location": wb.FullName}
        finally:
            self._done(saved=False)

    def _setup(self, tracker_sheet, add_cols, sheets):
        ok = False
        done = []
        try:
            wb = self._open()
            names = [ws.Name for ws in wb.Worksheets]
            ws = wb.Worksheets(tracker_sheet)
            hdr = self._header_map(ws)
            ur = ws.UsedRange
            lo = ws.ListObjects(1) if ws.ListObjects.Count else None
            for col in add_cols:
                if col in hdr:
                    continue
                if lo is not None:
                    lc = lo.ListColumns.Add()
                    lc.Name = col
                else:
                    last = max(hdr.values()) if hdr else 0
                    ws.Cells(ur.Row, last + 1).Value = col
                    hdr[col] = last + 1
                done.append(f"added column '{col}' to {tracker_sheet}")
            for name, table, headers, rows in sheets:
                if name in names:
                    continue
                new = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
                new.Name = name
                data = [headers] + (rows or [[""] * len(headers)])
                rng = new.Range(new.Cells(1, 1), new.Cells(len(data), len(headers)))
                rng.Value = tuple(tuple(r) for r in data)
                t = new.ListObjects.Add(1, rng, None, 1)  # xlSrcRange, has headers
                t.Name = table
                rng.Columns.AutoFit()
                done.append(f"created sheet '{name}' (table {table})")
            ok = bool(done)
            return done
        finally:
            self._done(saved=ok)

    # -- public interface (any thread) ------------------------------------------
    def read_sheet(self, sheet):
        return self.com.call(self._read, sheet)

    def read_row(self, sheet, row, headers):
        return self.com.call(self._read_row, sheet, row, headers)

    def write_cells(self, sheet, updates, headers=None):
        return self.com.call(self._write, sheet, updates)

    def append_rows(self, sheet, table, rows):
        return self.com.call(self._append, sheet, table, rows)

    def info(self):
        return self.com.call(self._info)

    def setup(self, tracker_sheet, add_cols, sheets):
        return self.com.call(self._setup, tracker_sheet, add_cols, sheets)

    def close(self):
        self.com.call(self._shutdown, timeout=30)


# --------------------------------------------------------------------------- openpyxl
class OpenpyxlFileBackend(LocalExcelBackend):
    """Direct file rewrite. Refuses workbooks with charts/PivotTables/etc."""

    def _guard(self):
        rich = rich_content(self.path)
        if rich:
            raise RuntimeError("This workbook contains " + ", ".join(rich) + ", which a direct file write would "
                               "delete. Switch Admin → Workbook → Engine to 'Microsoft Excel'.")

    def write_cells(self, sheet, updates, headers=None):
        self._guard()
        return super().write_cells(sheet, updates)

    def append_rows(self, sheet, table, rows):
        self._guard()
        return super().append_rows(sheet, table, rows)

    def info(self):
        from openpyxl import load_workbook
        wb = load_workbook(self.path, read_only=True)
        return {"sheets": wb.sheetnames, "cloud": False, "location": self.path}

    def setup(self, tracker_sheet, add_cols, sheets):
        self._guard()
        from openpyxl.worksheet.table import Table, TableStyleInfo
        with self._lock:
            wb = self._open()
            done = []
            ws = wb[tracker_sheet]
            hdr = [clean(c.value) for c in ws[1]]
            last = max((i for i, h in enumerate(hdr) if h), default=-1)
            for col in add_cols:
                if col not in hdr:
                    last += 1
                    ws.cell(row=1, column=last + 1, value=col)
                    hdr.append(col)
                    done.append(f"added column '{col}' to {tracker_sheet}")
            for t in ws.tables.values():  # stretch an existing table over the new columns
                start, end = t.ref.split(":")
                end_row = re.search(r"\d+", end).group()
                t.ref = f"{start}:{col_letter(last)}{end_row}"
            for name, table, headers, rows in sheets:
                if name in wb.sheetnames:
                    continue
                new = wb.create_sheet(name)
                for r in [headers] + (rows or [[""] * len(headers)]):
                    new.append(r)
                tb = Table(displayName=table, ref=f"A1:{col_letter(len(headers) - 1)}{new.max_row}")
                tb.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
                new.add_table(tb)
                done.append(f"created sheet '{name}' (table {table})")
            if done:
                self._save(wb)
            return done


def file_status(path: str) -> dict:
    p = Path(path)
    st_ = p.stat()
    return {"size_kb": round(st_.st_size / 1024), "modified": datetime.fromtimestamp(st_.st_mtime).strftime("%Y-%m-%d %H:%M"),
            "synced_by": synced_by(path),
            "rich": rich_content(path)}
