"""
exporter.py — Fill the Arabic/English supplier Excel template and build CSV.

Excel column mapping (sheet "Supplier Template", data starts at row 3):
  A(1) auto-number | B(2) arabic_name | C(3) english_name | G(7) emirate
  K(11) bank_name_ar | M(13) account_number | N(14) iban | O(15) account title
"""

import copy
import csv
import io
import os
import tempfile
import warnings

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.worksheet.datavalidation import DataValidation

from validator import canonical_emirate

SHEET_NAME = "Supplier Template"
START_ROW = 3

# (column index, record key) — account title (O) reuses english_name.
COLUMN_MAP = [
    (2,  "arabic_name"),
    (3,  "english_name"),
    (7,  "emirate"),
    (11, "bank_name_ar"),
    (13, "account_number"),
    (14, "iban"),
    (15, "english_name"),
]
TEXT_COLS = {13, 14}        # account_number (M), IBAN (N) — keep as literal text
ARABIC_COLS = {2, 11}       # arabic_name (B), bank_name_ar (K) — right-align

# The template's dropdowns live in the modern x14 extension format, which
# openpyxl silently drops on load/save. We re-attach them as classic list
# validations pointing at the same Sheet1 ranges so the saved file keeps them.
LIST_VALIDATIONS = [
    ("Sheet1!$B$2:$B$14", "G3:G251"),   # Emirate
    ("Sheet1!$A$2:$A$3",  "I3:I665"),   # Registration (Registered / Non Registered)
]

CSV_FIELDS = [
    "source_file", "english_name", "arabic_name", "iban", "account_number",
    "swift", "currency", "bank_name_ar", "bank_name_en", "nationality",
    "emirates_id", "date_of_birth", "id_expiry", "emirate", "iban_valid",
]


def _right_align(cell):
    """Right-align with RTL reading order for Arabic cells, keeping vertical."""
    a = cell.alignment
    cell.alignment = Alignment(horizontal="right", vertical=a.vertical, readingOrder=2)


def _reattach_dropdowns(wb, ws):
    """Re-add the list validations openpyxl dropped (best-effort)."""
    if "Sheet1" not in wb.sheetnames:
        return
    try:
        for formula, cells in LIST_VALIDATIONS:
            dv = DataValidation(type="list", formula1=formula, allow_blank=True)
            dv.add(cells)
            ws.add_data_validation(dv)
    except Exception:
        pass        # dropdowns are a convenience; never fail the export over them


def _row_has_data(ws, row) -> bool:
    """A data row counts as occupied if its name columns (B or C) are non-empty."""
    return any(ws.cell(row=row, column=c).value not in (None, "") for c in (2, 3))


def _next_empty_row(ws) -> int:
    """First empty data row at/after START_ROW — where appended records start."""
    row = START_ROW
    while _row_has_data(ws, row):
        row += 1
    return row


def existing_record_count(template_bytes: bytes) -> int:
    """How many records the template already holds (rows filled from row 3)."""
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = load_workbook(io.BytesIO(template_bytes), read_only=True)
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
        return max(0, _next_empty_row(ws) - START_ROW)
    except Exception:
        return 0


def export_to_excel(records: list, template_bytes: bytes, append: bool = True) -> bytes:
    """Return filled .xlsx bytes from the supplier template.

    append=True  → add records after the last existing row (keeps prior data).
    append=False → overwrite starting at row 3.
    """
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        f.write(template_bytes)
        tmp = f.name
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")          # suppress DataValidation warnings
            wb = load_workbook(tmp)
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

        start = _next_empty_row(ws) if append else START_ROW
        for i, rec in enumerate(records):
            row = start + i
            ws.cell(row=row, column=1, value=str(row - START_ROW + 1))   # 1-based running no.
            for col, key in COLUMN_MAP:
                value = rec.get(key, "")
                if col == 7:                         # snap emirate to dropdown value
                    value = canonical_emirate(value)
                cell = ws.cell(row=row, column=col, value=value)
                if col in TEXT_COLS:                 # never let long digits go 1.31E+13
                    cell.number_format = "@"
                if col in ARABIC_COLS:               # RTL-friendly alignment
                    _right_align(cell)

        _reattach_dropdowns(wb, ws)

        out = io.BytesIO()
        wb.save(out)
        return out.getvalue()
    finally:
        os.unlink(tmp)


def _excel_text_guard(value) -> str:
    """Wrap long all-digit identifiers as ="..." so Excel keeps them as text
    on a CSV double-click instead of showing 1.31E+13 and rounding the tail."""
    s = "" if value is None else str(value)
    return f'="{s}"' if s.isdigit() and len(s) >= 12 else s


def export_to_csv(records: list) -> bytes:
    """Return all records as CSV bytes, UTF-8 **with BOM**.

    The BOM (utf-8-sig) is essential: without it, Excel on Windows opens the
    file as Windows-1252 and turns every Arabic character into mojibake
    (e.g. محمد → "Ù…Ø­Ù…Ø¯"). With the BOM, Excel auto-detects UTF-8.
    Long account numbers are text-guarded so they don't become 1.31E+13.
    """
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=CSV_FIELDS, extrasaction="ignore")
    writer.writeheader()
    for rec in records:
        row = dict(rec)                              # don't mutate the caller's record
        row["account_number"] = _excel_text_guard(row.get("account_number", ""))
        writer.writerow(row)
    return buf.getvalue().encode("utf-8-sig")
