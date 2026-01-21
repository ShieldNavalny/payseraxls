# debug_pdf.py - отладочный скрипт

import pdfplumber
import PyPDF2

pdf_path = "example.pdf"

print("=" * 60)
print("АНАЛИЗ СТРУКТУРЫ PDF")
print("=" * 60)

# Метод 1: PyPDF2 для аннотаций
print("\n### PyPDF2: Аннотации ###")
with open(pdf_path, 'rb') as f:
    pdf_reader = PyPDF2.PdfReader(f)
    page = pdf_reader.pages[0]
    
    if '/Annots' in page:
        annots = page['/Annots']
        print(f"Найдено аннотаций: {len(annots)}")
        
        for i, annot in enumerate(annots[:5]):
            annot_obj = annot.get_object()
            print(f"\nАннотация {i}:")
            print(f"  Subtype: {annot_obj.get('/Subtype')}")
            print(f"  Rect: {annot_obj.get('/Rect')}")
            print(f"  C (color): {annot_obj.get('/C')}")
            print(f"  Contents: {annot_obj.get('/Contents')}")
            print(f"  Все ключи: {list(annot_obj.keys())}")
    else:
        print("Аннотаций не найдено")

# Метод 2: pdfplumber для текста и таблиц
print("\n### pdfplumber: Текст и таблицы ###")
with pdfplumber.open(pdf_path) as pdf:
    page = pdf.pages[0]
    
    text = page.extract_text()
    print("\nПервые 20 строк текста:")
    for i, line in enumerate(text.split('\n')[:20]):
        print(f"  {i}: {line[:80]}")
    
    tables = page.extract_tables()
    print(f"\nТаблиц найдено: {len(tables)}")
    if tables:
        print(f"Первая таблица, строк: {len(tables[0])}")
        for i, row in enumerate(tables[0][:5]):
            print(f"  Строка {i}: {row}")
