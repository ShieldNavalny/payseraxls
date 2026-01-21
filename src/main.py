from pathlib import Path
from src.pdf_parser import PDFParser
from src.excel_writer import ExcelWriter
from config import INPUT_DIR, OUTPUT_DIR, OUTPUT_FILENAME, DEFAULT_OUTPUT_MODE

def main():
    print("🔍 Paysera PDF Analyzer v4.0")
    print("=" * 50)
    
    if not INPUT_DIR.exists():
        INPUT_DIR.mkdir(parents=True)
        print(f"✗ Папка {INPUT_DIR} создана")
        return
    
    pdf_files = list(INPUT_DIR.glob('*.pdf'))
    if not pdf_files:
        print(f"✗ Нет PDF в {INPUT_DIR}")
        return
    
    print(f"📄 Файлов: {len(pdf_files)}")
    
    # ВЫБОР РЕЖИМА
    print("\nВыберите режим вывода:")
    print("  1. Один лист для всех файлов")
    print("  2. Отдельные листы для каждого файла (по умолчанию)")
    
    choice = input("\nВведите 1 или 2 (Enter = 2): ").strip()
    
    if choice == '1':
        mode = 'single'
    else:
        mode = 'multiple'
    
    print(f"Режим: {'один лист' if mode == 'single' else 'разделение по файлам'}")
    print("=" * 50)
    
    # Парсинг
    parser = PDFParser()
    transactions = parser.process_directory(str(INPUT_DIR))
    
    if not transactions:
        print("✗ Не найдено транзакций")
        return
    
    # Статистика
    negative = sum(1 for t in transactions if t['amount'].startswith('-'))
    positive = len(transactions) - negative
    
    # Подсчет сумм
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
    
    # Excel
    excel_writer = ExcelWriter()
    output_file = OUTPUT_DIR / OUTPUT_FILENAME
    
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    excel_writer.write_transactions(transactions, str(output_file), mode)
    
    print("=" * 50)
    print("✓ Готово!")

if __name__ == '__main__':
    main()
