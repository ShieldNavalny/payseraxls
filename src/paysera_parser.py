"""
paysera_parser.py — парсинг CSV-выписки Paysera → XLSX.

Цветовые пометки:
  - Золотой (#FFD700)    — покупка «в живую» (present в поле Paskirtis)
  - Светло-зелёный       — возврат (Grąžinimas в поле Tipas)
  - Без цвета            — онлайн-покупка, перевод, комиссия

Настраивается через корневой config.py.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Optional

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).parent.parent))
from config import (
    COLOR_PRESENT,
    COLOR_REFUND,
    REFUND_KEYWORDS,
    CARD_KEYWORDS,
)

# ---------------------------------------------------------------------------
# Имена колонок CSV
# ---------------------------------------------------------------------------

COL_TIPAS    = "Tipas"
COL_DATE     = "Data ir laikas"
COL_PAYEE    = "Gavėjas / Mokėtojas"
COL_AMOUNT   = "Suma ir valiuta"
COL_CURRENCY = "Valiutos"
COL_PURPOSE  = "Paskirtis"
COL_DR_CR    = "Kreditas / Debetas"
COL_BALANCE  = "Likutis"
COL_PRESENCE = "Presence"

_RE_NOT_PRESENT = re.compile(r"\bnot present\b", re.IGNORECASE)

FILL_PRESENT = PatternFill(start_color=COLOR_PRESENT, end_color=COLOR_PRESENT, fill_type="solid")
FILL_REFUND  = PatternFill(start_color=COLOR_REFUND,  end_color=COLOR_REFUND,  fill_type="solid")


# ---------------------------------------------------------------------------
# Вспомогательные функции
# ---------------------------------------------------------------------------

def _is_refund(tipas: str) -> bool:
    if not isinstance(tipas, str):
        return False
    return any(kw in tipas for kw in REFUND_KEYWORDS)


def _is_card_tx(tipas: str) -> bool:
    if not isinstance(tipas, str):
        return False
    tl = tipas.lower()
    return any(kw.lower() in tl for kw in CARD_KEYWORDS)


def _detect_presence(purpose: str) -> Optional[str]:
    """'not present' / 'present' / None (если не карточная транзакция)."""
    if not isinstance(purpose, str):
        return None
    return "not present" if _RE_NOT_PRESENT.search(purpose) else "present"


# ---------------------------------------------------------------------------
# Основные функции
# ---------------------------------------------------------------------------

def load_csv(csv_path: str | Path) -> pd.DataFrame:
    """Читает CSV Paysera → DataFrame с колонкой Presence."""
    df = pd.read_csv(
        csv_path,
        sep=",",
        quotechar='"',
        encoding="utf-8-sig",
    )
    df.columns = df.columns.str.strip()

    def _presence(row) -> Optional[str]:
        tipas   = row.get(COL_TIPAS, "")
        purpose = row.get(COL_PURPOSE, "")
        if _is_card_tx(str(tipas)):
            return _detect_presence(str(purpose))
        return None

    df[COL_PRESENCE] = df.apply(_presence, axis=1)
    return df


def to_excel(df: pd.DataFrame, output_path: str | Path) -> None:
    """Сохраняет DataFrame в XLSX с цветовой пометкой строк."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    df.to_excel(output_path, index=False, engine="openpyxl")

    wb = load_workbook(output_path)
    ws = wb.active

    headers = {cell.value: cell.column for cell in ws[1]}
    tipas_col    = headers.get(COL_TIPAS)
    presence_col = headers.get(COL_PRESENCE)

    for row in ws.iter_rows(min_row=2):
        tipas_val    = row[tipas_col - 1].value    if tipas_col    else ""
        presence_val = row[presence_col - 1].value if presence_col else None

        if _is_refund(str(tipas_val or "")):
            fill = FILL_REFUND
        elif presence_val == "present":
            fill = FILL_PRESENT
        else:
            fill = None

        if fill:
            for cell in row:
                cell.fill = fill

    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 60)

    wb.save(output_path)
    wb.close()


def convert(csv_path: str | Path, output_path: str | Path | None = None) -> Path:
    """Публичный API: CSV → XLSX. Возвращает путь к созданному файлу."""
    csv_path = Path(csv_path)
    if output_path is None:
        output_path = csv_path.with_suffix(".xlsx")
    output_path = Path(output_path)
    df = load_csv(csv_path)
    to_excel(df, output_path)
    print(f"✓ Сохранено: {output_path}")
    return output_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    import argparse
    parser = argparse.ArgumentParser(description="Paysera CSV → XLSX")
    parser.add_argument("csv",   help="Входной CSV-файл")
    parser.add_argument("xlsx",  nargs="?", help="Выходной XLSX (необязательно)")
    args = parser.parse_args()
    convert(args.csv, args.xlsx)


if __name__ == "__main__":
    main()
