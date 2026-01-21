from pathlib import Path
from src.pdf_parser import PDFParser
from src.excel_writer import ExcelWriter
from config import INPUT_DIR, OUTPUT_DIR, OUTPUT_FILENAME

def main():
    print("🔍 Paysera PDF Analyzer v3.1")
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
    
    parser = PDFParser()
    transactions = parser.process_directory(str(INPUT_DIR))
    
    if not transactions:
        print("✗ Не найдено транзакций")
        return
    
    # Статистика
    negative = sum(1 for t in transactions if t['amount'].startswith('-'))
    positive = len(transactions) - negative
    
    print(f"\n✓ Найдено: {len(transactions)}")
    print(f"   - Красные (расходы): {negative}")
    print(f"   - Фиолетовые (возвраты): {positive}")
    
    excel_writer = ExcelWriter()
    output_file = OUTPUT_DIR / OUTPUT_FILENAME
    
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    excel_writer.write_transactions(transactions, str(output_file))
    
    print("=" * 50)
    print("✓ Готово!")

if __name__ == '__main__':
    main()
