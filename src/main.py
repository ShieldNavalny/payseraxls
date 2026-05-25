import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent.parent))

from pdf_parser import PDFParser
from excel_writer import ExcelWriter
from csv_parser import parse_csv_to_excel
from invoice_matcher import run_matching
from config import INPUT_DIR, OUTPUT_DIR, OUTPUT_FILENAME, DEFAULT_OUTPUT_MODE

INVOICES_DIR = Path(__file__).parent.parent / "input_invoices"

def main():
    print("🔍 Paysera Analyzer v4.2")
    print("=" * 50)

    print("\nВыберите режим:")
    print("  1. PDF (выписка в формате PDF)")
    print("  2. CSV (выписка в формате CSV из Paysera)")
    print("  3. Сверка счетов (папка input_invoices/ вс XLS)")

    choice = input("\nВведите 1, 2 или 3 (Enter = 1): ").strip()

    # ------------------------------------------------------------------ CSV
    if choice == '2':
        if not INPUT_DIR.exists():
            INPUT_DIR.mkdir(parents=True)
            print(f"✗ Папка {INPUT_DIR} создана")
            return

        csv_files = list(INPUT_DIR.glob('*.csv'))
        if not csv_files:
            print(f"✗ Нет CSV-файлов в {INPUT_DIR}")
            return

        print(f"📄 Найдено CSV-файлов: {len(csv_files)}")
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

        for csv_file in csv_files:
            out_file = OUTPUT_DIR / (csv_file.stem + ".xlsx")
            print(f"  → {csv_file.name} ...")
            parse_csv_to_excel(str(csv_file), str(out_file))

        print("=" * 50)
        print("✓ Готово!")
        return

    # ------------------------------------------------------------------ Сверка счетов
    if choice == '3':
        xlsx_files = list(OUTPUT_DIR.glob('*.xlsx')) if OUTPUT_DIR.exists() else []
        if not xlsx_files:
            print(f"✗ Нет XLSX-файлов в {OUTPUT_DIR}")
            print("  Пожалуйста сначала создайте XLS из режима 1 или 2.")
            return

        if len(xlsx_files) == 1:
            xls_path = xlsx_files[0]
        else:
            print("\nДоступные XLS:")
            for i, f in enumerate(xlsx_files, 1):
                print(f"  {i}. {f.name}")
            idx = input("Выберите номер: ").strip()
            try:
                xls_path = xlsx_files[int(idx) - 1]
            except (ValueError, IndexError):
                print("✗ Неверный выбор")
                return

        run_matching(
            invoices_dir = INVOICES_DIR,
            xls_path     = xls_path,
            output_dir   = OUTPUT_DIR / "matched",
        )
        return

    # ------------------------------------------------------------------ PDF
    if not INPUT_DIR.exists():
        INPUT_DIR.mkdir(parents=True)
        print(f"✗ Папка {INPUT_DIR} создана")
        return

    pdf_files = list(INPUT_DIR.glob('*.pdf'))
    if not pdf_files:
        print(f"✗ Нет PDF в {INPUT_DIR}")
        return

    print(f"📄 Файлов: {len(pdf_files)}")
    print("\nВыберите режим вывода:")
    print("  1. Один лист для всех файлов")
    print("  2. Отдельные листы для каждого файла (по умолчанию)")

    pdf_mode = input("\nВведите 1 или 2 (Enter = 2): ").strip()
    mode = 'single' if pdf_mode == '1' else 'multiple'

    print(f"Режим: {'один лист' if mode == 'single' else 'разделение по файлам'}")
    print("=" * 50)

    parser = PDFParser()
    transactions = parser.process_directory(str(INPUT_DIR))

    if not transactions:
        print("✗ Не найдено транзакций")
        return

    negative = sum(1 for t in transactions if t['amount'].startswith('-'))
    positive = len(transactions) - negative

    expenses_total = sum(
        float(t['amount'].replace(' EUR', '').replace(',', '.'))
        for t in transactions if t['amount'].startswith('-')
    )
    returns_total = sum(
        float(t['amount'].replace(' EUR', '').replace(',', '.'))
        for t in transactions if not t['amount'].startswith('-')
    )

    print(f"\n✓ Найдено: {len(transactions)}")
    print(f"   - Красные (расходы): {negative} транзакций на {expenses_total:.2f} EUR")
    print(f"   - Фиолетовые (возвраты): {positive} транзакций на {returns_total:.2f} EUR")

    excel_writer = ExcelWriter()
    output_file = OUTPUT_DIR / OUTPUT_FILENAME
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    excel_writer.write_transactions(transactions, str(output_file), mode)

    print("=" * 50)
    print("✓ Готово!")

if __name__ == '__main__':
    main()
