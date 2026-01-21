from pathlib import Path
from src.pdf_parser import PDFParser
from src.excel_writer import ExcelWriter
from config import INPUT_DIR, OUTPUT_DIR, OUTPUT_FILENAME

def main():
    print("🔍 Paysera PDF Analyzer v1.0")
    print("=" * 50)
    
    # Проверяем наличие входной папки
    if not INPUT_DIR.exists():
        INPUT_DIR.mkdir(parents=True)
        print(f"✗ Папка {INPUT_DIR} создана. Поместите PDF файлы туда.")
        return
    
    pdf_files = list(INPUT_DIR.glob('*.pdf'))
    if not pdf_files:
        print(f"✗ Нет PDF файлов в {INPUT_DIR}")
        return
    
    print(f"📄 Найдено файлов: {len(pdf_files)}")
    
    # Парсим PDF
    parser = PDFParser()
    transactions = parser.process_directory(str(INPUT_DIR))
    
    if not transactions:
        print("✗ Не найдено проблемных транзакций")
        return
    
    print(f"⚠️  Найдено проблемных транзакций: {len(transactions)}")
    
    # Статистика
    yellow_count = sum(1 for t in transactions if t['color_type'] == 'yellow')
    green_count = sum(1 for t in transactions if t['color_type'] == 'green')
    
    print(f"   - Красные (недостаток счетов): {yellow_count}")
    print(f"   - Фиолетовые (возвраты без документов): {green_count}")
    
    # Создаем Excel
    excel_writer = ExcelWriter()
    output_file = OUTPUT_DIR / OUTPUT_FILENAME
    
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    excel_writer.write_transactions(transactions, str(output_file))
    
    print("=" * 50)
    print("✓ Обработка завершена!")

if __name__ == '__main__':
    main()
