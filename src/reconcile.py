"""
reconcile.py — сравнение XLS-вывода с CSV-реестром счетов.

Логика:
  1. Читает CSV из корня рабочей директории (documents.csv по умолчанию).
  2. Читает XLSX из output_reports/.
  3. Строит ключи (дата, получатель, |сумма|) для обеих сторон.
  4. Находит записи из CSV, которых НЕТ в XLS.
  5. Копирует недостающие PDF в missing_invoices/<Vendor>/
"""

import csv
import re
import shutil
from pathlib import Path
from typing import List, Dict, Set, Tuple

import openpyxl

from config import OUTPUT_DIR, OUTPUT_FILENAME

# ─── настройки ────────────────────────────────────────────────────────────────
DEFAULT_CSV_NAME = "documents.csv"
MISSING_DIR_NAME = "missing_invoices"


# ─── вспомогательные функции ──────────────────────────────────────────────────

def _normalize_amount(raw: str) -> str:
    """Приводит сумму к строке без знака и EUR, с двумя знаками после точки.
    '-406.96 EUR' → '406.96'
    '406.96'      → '406.96'
    """
    raw = raw.strip().replace(' EUR', '').replace(',', '.')
    try:
        return f"{abs(float(raw)):.2f}"
    except ValueError:
        return raw


def _normalize_vendor(raw: str) -> str:
    return raw.strip().lower()


def _normalize_date_xls(raw: str) -> str:
    """Дата в XLS: строка вида '2025-11-03' или объект datetime."""
    if hasattr(raw, 'strftime'):
        return raw.strftime('%Y-%m-%d')
    raw = str(raw).strip()
    # '03.11.2025' → '2025-11-03'
    m = re.match(r'^(\d{2})\.(\d{2})\.(\d{4})$', raw)
    if m:
        return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
    return raw  # уже ISO


def _make_key(date: str, vendor: str, amount: str) -> Tuple[str, str, str]:
    return (_normalize_date_xls(date), _normalize_vendor(vendor), _normalize_amount(amount))


# ─── чтение XLS ───────────────────────────────────────────────────────────────

def load_xls_keys(xls_path: Path) -> Set[Tuple[str, str, str]]:
    """
    Читает все листы XLS (кроме 'Сводка') и собирает ключи транзакций.
    Ожидаемые колонки листа: Дата(1) Получатель(2) Сумма(3) ...
    """
    keys: Set[Tuple[str, str, str]] = set()
    wb = openpyxl.load_workbook(xls_path, data_only=True)

    for sheet_name in wb.sheetnames:
        if sheet_name in ('Сводка',):
            continue
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2, values_only=True):
            date_val, vendor_val, amount_val = row[0], row[1], row[2]
            if not date_val or not amount_val:
                continue
            keys.add(_make_key(
                str(date_val),
                str(vendor_val or ''),
                str(amount_val)
            ))

    wb.close()
    return keys


# ─── чтение CSV ───────────────────────────────────────────────────────────────

def load_csv_records(csv_path: Path) -> List[Dict]:
    """
    Возвращает список записей из CSV.
    Ожидаемые колонки: Date, Vendor, Amount, Currency, Invoice ID, Type, File
    """
    records = []
    with open(csv_path, newline='', encoding='utf-8-sig') as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            records.append({
                'date':       row.get('Date', '').strip(),
                'vendor':     row.get('Vendor', '').strip(),
                'amount':     row.get('Amount', '').strip(),
                'invoice_id': row.get('Invoice ID', '').strip(),
                'type':       row.get('Type', '').strip(),
                'file':       row.get('File', '').strip(),
            })
    return records


# ─── основная функция ─────────────────────────────────────────────────────────

def reconcile(
    base_dir: Path | None = None,
    csv_name: str = DEFAULT_CSV_NAME,
    xls_name: str | None = None,
    missing_dir_name: str = MISSING_DIR_NAME,
    dry_run: bool = False,
) -> List[Dict]:
    """
    Запускает сверку.

    :param base_dir:         корневая директория (по умолчанию — cwd).
    :param csv_name:         имя CSV-файла в base_dir.
    :param xls_name:         имя XLS-файла (по умолчанию из config.py).
    :param missing_dir_name: имя выходной директории.
    :param dry_run:          если True — только вычисляет, не копирует.
    :return:                 список недостающих записей.
    """
    if base_dir is None:
        base_dir = Path.cwd()

    csv_path = base_dir / csv_name
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV не найден: {csv_path}")

    xls_path = OUTPUT_DIR / (xls_name or OUTPUT_FILENAME)
    if not xls_path.exists():
        raise FileNotFoundError(f"XLS не найден: {xls_path}")

    print(f"📂 CSV:  {csv_path}")
    print(f"📊 XLS:  {xls_path}")

    xls_keys = load_xls_keys(xls_path)
    print(f"   XLS транзакций: {len(xls_keys)}")

    csv_records = load_csv_records(csv_path)
    print(f"   CSV записей:    {len(csv_records)}")

    # ── находим недостающие ────────────────────────────────────────────────────
    missing: List[Dict] = []
    for rec in csv_records:
        key = _make_key(rec['date'], rec['vendor'], rec['amount'])
        if key not in xls_keys:
            missing.append(rec)

    print(f"\n🔍 Недостающих счетов: {len(missing)}")

    if not missing:
        print("✓ Всё совпадает — расхождений нет.")
        return missing

    # ── группируем по вендору для наглядности ─────────────────────────────────
    by_vendor: Dict[str, List[Dict]] = {}
    for rec in missing:
        by_vendor.setdefault(rec['vendor'], []).append(rec)

    for vendor, recs in sorted(by_vendor.items()):
        print(f"   {vendor}: {len(recs)} шт.")

    if dry_run:
        print("\n⚠  Режим dry-run: файлы не скопированы.")
        return missing

    # ── копируем PDF ──────────────────────────────────────────────────────────
    out_root = base_dir / missing_dir_name
    copied = 0
    not_found = 0

    print(f"\n📁 Копируем в: {out_root}")

    for rec in missing:
        if not rec['file']:
            continue

        src = base_dir / rec['file']
        if not src.exists():
            print(f"   ⚠ Не найден файл: {src}")
            not_found += 1
            continue

        # Папка: missing_invoices/<Vendor>/
        vendor_safe = re.sub(r'[\\/:*?"<>|]', '_', rec['vendor'])  # sanitize
        dst_dir = out_root / vendor_safe
        if not dst_dir.exists():
            dst_dir.mkdir(parents=True, exist_ok=True)

        dst = dst_dir / src.name
        shutil.copy2(src, dst)
        copied += 1

    print(f"\n✓ Скопировано: {copied}")
    if not_found:
        print(f"⚠  Не найдено на диске: {not_found}")

    return missing


# ─── CLI ──────────────────────────────────────────────────────────────────────

def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="Сверка XLS-вывода с CSV-реестром счетов"
    )
    parser.add_argument(
        '--csv', default=DEFAULT_CSV_NAME,
        help=f"Имя CSV-файла в корне (по умолчанию: {DEFAULT_CSV_NAME})"
    )
    parser.add_argument(
        '--xls', default=None,
        help="Имя XLS-файла в output_reports/ (по умолчанию из config.py)"
    )
    parser.add_argument(
        '--out', default=MISSING_DIR_NAME,
        help=f"Папка для недостающих счетов (по умолчанию: {MISSING_DIR_NAME})"
    )
    parser.add_argument(
        '--dry-run', action='store_true',
        help="Только показать расхождения, не копировать файлы"
    )
    parser.add_argument(
        '--base-dir', default=None,
        help="Корневая директория (по умолчанию: текущая)"
    )
    args = parser.parse_args()

    base = Path(args.base_dir) if args.base_dir else None

    print("🔄 Payseraxls — Reconcile")
    print("=" * 50)
    reconcile(
        base_dir=base,
        csv_name=args.csv,
        xls_name=args.xls,
        missing_dir_name=args.out,
        dry_run=args.dry_run,
    )
    print("=" * 50)


if __name__ == '__main__':
    main()
