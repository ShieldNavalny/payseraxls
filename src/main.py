import sys
from pathlib import Path

# Allow 'python src/main.py' from project root AND 'python main.py' from src/
sys.path.insert(0, str(Path(__file__).parent))          # adds src/
sys.path.insert(0, str(Path(__file__).parent.parent))   # adds project root (for config)

from pdf_parser import PDFParser
from excel_writer import ExcelWriter
from csv_parser import parse_csv_to_excel
from config import INPUT_DIR, OUTPUT_DIR, OUTPUT_FILENAME, DEFAULT_OUTPUT_MODE

def main():
    print("🔍 Paysera Analyzer v4.1")
    print("=" * 50)

    print("\nВыберите режим входных данных:")
    print("  1. PDF (выписка в формате PDF)")
    print("  2. CSV (выписка в формате CSV из Paysera)")

    source_choice = input("\nВведите 1 или 2 (Enter = 1): ").strip()

    # ------------------------------------------------------------------ CSV
    if source_choice == '2':
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

    choice = input("\nВведите 1 или 2 (Enter = 2): ").strip()
    mode = 'single' if choice == '1' else 'multiple'

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
