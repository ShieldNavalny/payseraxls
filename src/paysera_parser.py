"""
Paysera CSV → Excel converter.

Features:
- Parses the Lithuanian-header Paysera statement CSV
- Highlights rows differently based on transaction presence:
    present     → physical/POS purchase → yellow fill (#FFD700)
    not present → online/card-not-present → default (white/no fill)
- Exports the result as .xlsx via openpyxl
"""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

CSV_ENCODING = "utf-8-sig"          # Paysera exports BOM-UTF-8
CSV_SEPARATOR = ","

# Column names in the CSV (Lithuanian headers)
COL_TYPE       = "Tipas"
COL_DOC_NO     = "Išrašo nr."
COL_TRANSFER   = "Pervedimo nr."
COL_DATETIME   = "Data ir laikas"
COL_PARTY      = "Gavėjas / Mokėtojas"
COL_IBAN       = "EVP / IBAN"
COL_CODE       = "Kodas"
COL_AMOUNT_CCY = "Suma ir valiuta"
COL_CURRENCY   = "Valiutos"
COL_DETAILS    = "Paskirtis"
COL_PMT_CODE   = "Įmokos kodas"
COL_DR_CR      = "Kreditas / Debetas"
COL_BALANCE    = "Likutis"

# Extra column produced by this script
COL_PRESENCE   = "Presence"

# Regex to detect presence type in the Details field
_RE_PRESENT     = re.compile(r"\bnot present\b", re.IGNORECASE)

# Excel fill colours
FILL_PRESENT     = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")  # gold
FILL_NOT_PRESENT = PatternFill(fill_type=None)  # no fill  (white / default)


# ---------------------------------------------------------------------------
# Core helpers
# ---------------------------------------------------------------------------

def _detect_presence(details: str) -> str:
    """Return 'not present' or 'present' based on the Details text."""
    if isinstance(details, str) and _RE_PRESENT.search(details):
        return "not present"
    return "present"


def load_csv(path: str | Path) -> pd.DataFrame:
    """Read a Paysera CSV statement and return a normalised DataFrame."""
    df = pd.read_csv(
        path,
        encoding=CSV_ENCODING,
        sep=CSV_SEPARATOR,
        dtype=str,          # keep everything as strings initially
        keep_default_na=False,
    )

    # Normalise numeric columns
    for col in (COL_AMOUNT_CCY, COL_BALANCE):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].str.replace(",", "."), errors="coerce")

    # Parse datetime column
    if COL_DATETIME in df.columns:
        df[COL_DATETIME] = pd.to_datetime(df[COL_DATETIME], errors="coerce")

    # Derive presence column
    df[COL_PRESENCE] = df[COL_DETAILS].apply(_detect_presence)

    return df


def to_excel(df: pd.DataFrame, output_path: str | Path) -> Path:
    """
    Write the DataFrame to an .xlsx file and apply row colours:
      - 'present'     rows → gold fill (physical purchase)
      - 'not present' rows → no fill  (online purchase)

    Returns the resolved output path.
    """
    output_path = Path(output_path)

    # Write with pandas first (fast)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Statement")

    # Re-open with openpyxl to apply styles
    wb = load_workbook(output_path)
    ws = wb.active

    # Find the index of the Presence column in the sheet (1-based)
    header = [cell.value for cell in ws[1]]
    try:
        presence_col_idx = header.index(COL_PRESENCE) + 1
    except ValueError:
        presence_col_idx = None

    # Apply fills starting from row 2 (skip header)
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        if presence_col_idx is not None:
            presence_value = ws.cell(row=row_idx, column=presence_col_idx).value
        else:
            presence_value = None

        fill = FILL_PRESENT if presence_value == "present" else FILL_NOT_PRESENT
        for cell in row:
            cell.fill = fill

    # Auto-fit column widths (best-effort)
    for col in ws.columns:
        max_len = max(
            (len(str(c.value)) if c.value is not None else 0) for c in col
        )
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)

    wb.save(output_path)
    return output_path


def convert(csv_path: str | Path, output_path: str | Path | None = None) -> Path:
    """
    One-shot helper: load CSV → enrich → export to Excel.

    If *output_path* is omitted, the result is placed next to the source CSV
    with a .xlsx extension.
    """
    csv_path = Path(csv_path)
    if output_path is None:
        output_path = csv_path.with_suffix(".xlsx")

    df = load_csv(csv_path)
    result = to_excel(df, output_path)
    print(f"[paysera_parser] Exported {len(df)} rows → {result}")
    return result
