"""
reconcile.py — сравнение XLS-вывода с CSV-реестром счетов.

Логика:
  1. --csv принимает полный путь. Папка CSV становится base_dir для PDF.
  2. --xls принимает полный путь.
  3. --out папка для вывода: если относительный — относительно cwd; если абсолютный — как есть.
  4. PDF-файлы ищутся относительно папки CSV (base_dir).
  5. Создаёт reconciliation_report.xlsx в --out с зелёными строками для найденных.
"""

import csv
import re
import shutil
from pathlib import Path
from typing import List, Dict, Set, Tuple, Optional

import openpyxl
from openpyxl.styles import PatternFill

# ─── настройки ────────────────────────────────────────────────────────────────
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
    ("wolt",                  "wolt"),
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


# ─── чтение XLS ───────────────────────────────────────────────────────────────

def load_xls_keys(xls_path: Path) -> Set[Tuple[str, str, str]]:
    """Собирает ключи только расходных транзакций (amount < 0)."""
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

            keys.add(_make_key(str(date_val), str(vendor_val), str(amount_val)))

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


def build_csv_keys(csv_records: List[Dict]) -> Set[Tuple[str, str, str]]:
    keys: Set[Tuple[str, str, str]] = set()
    for rec in csv_records:
        keys.add(_make_key(rec['date'], rec['vendor'], rec['amount']))
    return keys


# ─── аннотированный XLS ─────────────────────────────────────────────────────────

def create_annotated_xls(
    original_xls: Path,
    csv_keys: Set[Tuple[str, str, str]],
    output_xls: Path,
) -> None:
    """
    Копирует XLS и красит зелёным строки, для которых нашлись счета.
    """
    wb = openpyxl.load_workbook(original_xls)
    green_fill = PatternFill(start_color="FF00FF00", end_color="FF00FF00", fill_type="solid")

    for sheet_name in wb.sheetnames:
        if sheet_name == 'Сводка':
            continue
        ws = wb[sheet_name]

        for row in ws.iter_rows(min_row=2):
            date_val = row[0].value
            vendor_val = row[1].value
            amount_val = row[2].value

            if not date_val or amount_val is None or not vendor_val:
                continue

            amount_str = str(amount_val).replace(' EUR', '').replace(',', '.').strip()
            try:
                amount_float = float(amount_str)
            except ValueError:
                continue

            if amount_float >= 0:
                continue

            key = _make_key(str(date_val), str(vendor_val), str(amount_val))
            if key in csv_keys:
                for cell in row:
                    cell.fill = green_fill

    output_xls.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_xls)
    wb.close()


# ─── основная функция ─────────────────────────────────────────────────────────

def reconcile(
    csv_path: Path,
    xls_path: Path,
    out_dir: Path,
    dry_run: bool = False,
) -> List[Dict]:
    """
    :param csv_path:  полный путь к CSV. Его папка — base_dir для PDF.
    :param xls_path:  полный путь к XLS.
    :param out_dir:   куда копировать PDF и отчёт.
    :param dry_run:   только статистика, без копирования.
    """
    # PDF ищем относительно папки CSV
    base_dir = csv_path.parent.resolve()

    print(f"📂 CSV:      {csv_path}")
    print(f"📊 XLS:      {xls_path}")
    print(f"📁 PDF база:  {base_dir}")
    print(f"📄 Вывод:    {out_dir}")

    if not csv_path.exists():
        raise FileNotFoundError(f"CSV не найден: {csv_path}")
    if not xls_path.exists():
        raise FileNotFoundError(f"XLS не найден: {xls_path}")

    xls_keys = load_xls_keys(xls_path)
    print(f"   XLS расходных транзакций: {len(xls_keys)}")

    csv_records = load_csv_records(csv_path)
    print(f"   CSV записей:              {len(csv_records)}")

    csv_keys = build_csv_keys(csv_records)

    # ── находим записи из CSV, которых НЕТ в XLS ─────────────────────────────
    missing: List[Dict] = []
    for rec in csv_records:
        key = _make_key(rec['date'], rec['vendor'], rec['amount'])
        if key not in xls_keys:
            missing.append(rec)

    matched_count = len(csv_records) - len(missing)
    print(f"\n✓ Найдено счетов:    {matched_count} / {len(csv_records)}")
    print(f"🔍 Недостающих: {len(missing)}")

    if missing:
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
    if missing:
        print(f"\n📁 Копируем PDF в: {out_dir}")
        copied = 0
        not_found = 0

        for rec in missing:
            if not rec['file']:
                print(f"   ⚠ Нет пути: {rec['vendor']} {rec['date']}")
                continue

            # PDF относительно папки CSV
            src = base_dir / rec['file']
            if not src.exists():
                print(f"   ⚠ Не найден: {src}")
                not_found += 1
                continue

            vendor_safe = re.sub(r'[\\/:*?"<>|]', '_', rec['vendor'])
            dst_dir = out_dir / vendor_safe
            dst_dir.mkdir(parents=True, exist_ok=True)

            shutil.copy2(src, dst_dir / src.name)
            copied += 1

        print(f"✓ Скопировано: {copied}")
        if not_found:
            print(f"⚠  Не найдено на диске: {not_found}")
    else:
        print("✓ Всё совпадает — расхождений нет.")

    # ── аннотированный XLS ────────────────────────────────────────────────
    report_path = out_dir / REPORT_XLS_NAME
    out_dir.mkdir(parents=True, exist_ok=True)
    print(f"\n📊 Отчёт XLS: {report_path}")
    create_annotated_xls(xls_path, csv_keys, report_path)
    print("✓ Зелёные — счёт найден, красные/фиолетовые — нет")

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
        description="Сверка XLS-вывода с CSV-реестром счетов",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Пример:
  python -m src.reconcile \\
    --csv "C:\\INVOICES\\documents.csv" \\
    --xls "C:\\payseraxls\\output_reports\\miss.xlsx" \\
    --out  "C:\\INVOICES\\missing_invoices"

  PDF-файлы ищутся относительно папки с CSV.
"""
    )
    parser.add_argument(
        '--csv',
        required=True,
        help="Полный путь к CSV-файлу. Папка CSV = база для поиска PDF."
    )
    parser.add_argument(
        '--xls',
        required=True,
        help="Полный путь к XLS/XLSX-файлу с выпиской."
    )
    parser.add_argument(
        '--out',
        default=MISSING_DIR_NAME,
        help=f"Папка для недостающих PDF и отчёта (по умолчанию: {MISSING_DIR_NAME} в cwd)"
    )
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help="Только показать расхождения, не копировать файлы"
    )
    args = parser.parse_args()

    csv_path = Path(args.csv).resolve()
    xls_path = Path(args.xls).resolve()
    out_dir  = Path(args.out) if Path(args.out).is_absolute() else Path.cwd() / args.out
    out_dir  = out_dir.resolve()

    print("🔄 Payseraxls — Reconcile")
    print("=" * 50)
    reconcile(
        csv_path=csv_path,
        xls_path=xls_path,
        out_dir=out_dir,
        dry_run=args.dry_run,
    )
    print("=" * 50)


if __name__ == '__main__':
    main()
