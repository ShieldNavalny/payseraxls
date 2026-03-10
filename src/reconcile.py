"""
reconcile.py — сравнение XLS-вывода с CSV-реестром счетов.

Логика:
  1. Читает CSV из корня рабочей директории (documents.csv по умолчанию).
  2. Читает XLSX из output_reports/.
  3. Строит ключи (дата, canonical_vendor, |сумма|) для обеих сторон.
     - XLS «Получатель» нормализуется через VENDOR_MAP.
     - CSV «Vendor» тоже нормализуется через тот же VENDOR_MAP.
     - Суммы в CSV всегда положительные; в XLS могут быть отрицательными — берём |abs|.
  4. Находит записи из CSV, которых НЕТ в XLS (только расходы — отрицательные в XLS).
  5. Копирует недостающие PDF в missing_invoices/<Vendor>/
  6. Создаёт аннотированный XLS в missing_invoices/ с зелёными строками для найденных.
"""

import csv
import re
import shutil
from pathlib import Path
from typing import List, Dict, Set, Tuple, Optional

import openpyxl
from openpyxl.styles import PatternFill

from config import OUTPUT_DIR, OUTPUT_FILENAME

# ─── настройки ────────────────────────────────────────────────────────────────
DEFAULT_CSV_NAME = "documents.csv"
MISSING_DIR_NAME = "missing_invoices"
REPORT_XLS_NAME = "reconciliation_report.xlsx"

# ─── словарь нормализации вендоров ────────────────────────────────────────────
VENDOR_MAP: List[Tuple[str, str]] = [
    ("sinch mailgun",         "mailgun"),
    ("mailgun",               "mailgun"),
    ("hetzner online",        "hetzner"),
    ("hetzner",               "hetzner"),
    ("amazon",                "amazon"),
    ("allegro",               "allegro"),
    ("aliexpress",            "aliexpress"),
    ("bolt.eu",               "bolt"),
    ("bolt",                  "bolt"),
    ("eu.store.ui.com",       "ui.com"),
    ("ui.com",                "ui.com"),
    ("eurocash",              "eurocash1.lt"),
    ("pirkeu",                "pirkeu.lt"),
    ("sandeliukunuoma",       "sandeliukunuoma.lt"),
]


def _canonical_vendor(raw: str) -> str:
    low = raw.strip().lower()
    for pattern, canonical in VENDOR_MAP:
        if pattern in low:
            return canonical
    return low


def _normalize_amount(raw: str) -> str:
    cleaned = raw.strip().replace(' EUR', '').replace(',', '.').replace('\xa0', '')
    try:
        return f"{abs(float(cleaned)):.2f}"
    except ValueError:
        return cleaned


def _normalize_date(raw) -> str:
    if hasattr(raw, 'strftime'):
        return raw.strftime('%Y-%m-%d')
    raw = str(raw).strip()
    m = re.match(r'^(\d{2})\.(\d{2})\.(\d{4})$', raw)
    if m:
        return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
    return raw


def _make_key(date: str, vendor: str, amount: str) -> Tuple[str, str, str]:
    return (
        _normalize_date(date),
        _canonical_vendor(vendor),
        _normalize_amount(amount),
    )


# ─── чтение XLS — расходные ключи ─────────────────────────────────────────────

def load_xls_keys(xls_path: Path) -> Set[Tuple[str, str, str]]:
    """Берём только расходы (amount < 0)."""
    keys: Set[Tuple[str, str, str]] = set()
    wb = openpyxl.load_workbook(xls_path, data_only=True)

    for sheet_name in wb.sheetnames:
        if sheet_name == 'Сводка':
            continue
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2, values_only=True):
            date_val, vendor_val, amount_val = row[0], row[1], row[2]
            if not date_val or amount_val is None or not vendor_val:
                continue

            amount_str = str(amount_val).replace(' EUR', '').replace(',', '.').strip()
            try:
                amount_float = float(amount_str)
            except ValueError:
                continue

            if amount_float >= 0:
                continue

            keys.add(_make_key(
                str(date_val),
                str(vendor_val),
                str(amount_val),
            ))

    wb.close()
    return keys


# ─── чтение CSV ───────────────────────────────────────────────────────────────

def load_csv_records(csv_path: Path) -> List[Dict]:
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


# ─── CSV ключи — те, что есть в реестре ──────────────────────────────────────────

def build_csv_keys(csv_records: List[Dict]) -> Set[Tuple[str, str, str]]:
    """Строим набор ключей из CSV (дата, vendor, amount)."""
    keys: Set[Tuple[str, str, str]] = set()
    for rec in csv_records:
        keys.add(_make_key(rec['date'], rec['vendor'], rec['amount']))
    return keys


# ─── аннотирование XLS — зелёные строки для найденных ────────────────────────────

def create_annotated_xls(
    original_xls: Path,
    csv_keys: Set[Tuple[str, str, str]],
    output_xls: Path,
) -> None:
    """
    Копирует оригинальный XLS, красит строки, для которых нашлись счета в CSV,
    в зелёный цвет. Остальные остаются красными/фиолетовыми.
    """
    wb = openpyxl.load_workbook(original_xls)
    green_fill = PatternFill(start_color="FF00FF00", end_color="FF00FF00", fill_type="solid")

    for sheet_name in wb.sheetnames:
        if sheet_name == 'Сводка':
            continue
        ws = wb[sheet_name]

        for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            date_cell = row[0]
            vendor_cell = row[1]
            amount_cell = row[2]

            date_val = date_cell.value
            vendor_val = vendor_cell.value
            amount_val = amount_cell.value

            if not date_val or amount_val is None or not vendor_val:
                continue

            # Проверяем, что это расход (< 0)
            amount_str = str(amount_val).replace(' EUR', '').replace(',', '.').strip()
            try:
                amount_float = float(amount_str)
            except ValueError:
                continue

            if amount_float >= 0:
                continue  # возвраты не красим

            key = _make_key(str(date_val), str(vendor_val), str(amount_val))
            if key in csv_keys:
                # Найден счёт — красим строку в зелёный
                for cell in row:
                    cell.fill = green_fill

    output_xls.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_xls)
    wb.close()


# ─── основная функция ─────────────────────────────────────────────────────────

def reconcile(
    base_dir: Optional[Path] = None,
    csv_name: str = DEFAULT_CSV_NAME,
    xls_name: Optional[str] = None,
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
    print(f"   XLS расходных транзакций: {len(xls_keys)}")

    csv_records = load_csv_records(csv_path)
    print(f"   CSV записей:              {len(csv_records)}")

    csv_keys = build_csv_keys(csv_records)
    print(f"   CSV уникальных ключей:       {len(csv_keys)}")

    # ── находим недостающие (транзакции XLS, которых НЕТ в CSV) ────────────────────
    # Логика: xls_keys — расходы в выписке, csv_keys — счета.
    # missing = xls_keys - csv_keys (но мы хотим полные записи, поэтому итерируем по csv_records)
    # Нет, наоборот: нам нужны записи из CSV, которых НЕТ в XLS
    # — это счета, которые не соответствуют транзакциям.
    missing: List[Dict] = []
    for rec in csv_records:
        key = _make_key(rec['date'], rec['vendor'], rec['amount'])
        if key not in xls_keys:
            missing.append(rec)

    matched_count = len(csv_records) - len(missing)
    print(f"\n✓ Найдено счетов:    {matched_count} / {len(csv_records)}")
    print(f"🔍 Недостающих счетов: {len(missing)}")

    if not missing:
        print("✓ Всё совпадает — расхождений нет.")
        # Всё равно создаём аннотированный XLS
        out_root = base_dir / missing_dir_name
        report_path = out_root / REPORT_XLS_NAME
        create_annotated_xls(xls_path, csv_keys, report_path)
        print(f"📄 Аннотированный XLS: {report_path}")
        return missing

    # ── группируем по вендору ───────────────────────────────────────────────────
    by_vendor: Dict[str, List[Dict]] = {}
    for rec in missing:
        by_vendor.setdefault(rec['vendor'], []).append(rec)

    for vendor, recs in sorted(by_vendor.items()):
        print(f"   {vendor}: {len(recs)} шт.")

    if dry_run:
        print("\n⚠  Режим dry-run: файлы не скопированы.")
        _print_missing_table(missing)
        return missing

    # ── копируем PDF ──────────────────────────────────────────────────────────
    out_root = base_dir / missing_dir_name
    copied = 0
    not_found = 0

    print(f"\n📁 Копируем в: {out_root}")

    for rec in missing:
        if not rec['file']:
            print(f"   ⚠ Нет пути к файлу для: {rec['vendor']} {rec['date']} {rec['amount']}")
            continue

        src = base_dir / rec['file']
        if not src.exists():
            print(f"   ⚠ Не найден файл: {src}")
            not_found += 1
            continue

        vendor_safe = re.sub(r'[\\/:*?"<>|]', '_', rec['vendor'])
        dst_dir = out_root / vendor_safe
        dst_dir.mkdir(parents=True, exist_ok=True)

        dst = dst_dir / src.name
        shutil.copy2(src, dst)
        copied += 1

    print(f"\n✓ Скопировано: {copied}")
    if not_found:
        print(f"⚠  Не найдено на диске: {not_found}")

    # ── создаём аннотированный XLS ────────────────────────────────────────────────
    report_path = out_root / REPORT_XLS_NAME
    print(f"\n📊 Создаём аннотированный XLS: {report_path}")
    create_annotated_xls(xls_path, csv_keys, report_path)
    print("✓ Зелёные строчки — найдены счета")
    print("  Красные/фиолетовые — нет счетов")

    return missing


def _print_missing_table(records: List[Dict]) -> None:
    print(f"\n{'Дата':<12} {'Вендор':<25} {'Сумма':>10}  Invoice ID")
    print("-" * 70)
    for r in records:
        print(f"{r['date']:<12} {r['vendor']:<25} {r['amount']:>10}  {r['invoice_id']}")


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
