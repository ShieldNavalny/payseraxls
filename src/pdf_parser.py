import pdfplumber
from pathlib import Path
from typing import List, Dict, Tuple
import re

class PDFParser:
    def __init__(self):
        self.transactions = []
    
    def extract_colored_text(self, pdf_path: str) -> List[Dict]:
        """
        Извлекает текст с информацией о цвете фона из PDF
        """
        transactions = []
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    # Получаем raw данные страницы
                    text = page.extract_text()
                    rects = page.rects
                    
                    # Извлекаем таблицы (основной источник данных)
                    tables = page.extract_tables()
                    
                    if tables:
                        for table in tables:
                            for row_idx, row in enumerate(table):
                                # Проверяем цвет фона строки через rects
                                color_type = self._detect_row_color(
                                    page, row_idx, rects
                                )
                                
                                if color_type in ['yellow', 'green']:
                                    transaction = self._parse_transaction_row(
                                        row, color_type, page_num
                                    )
                                    if transaction:
                                        transactions.append(transaction)
        
        except Exception as e:
            print(f"Ошибка при парсинге {pdf_path}: {e}")
        
        return transactions
    
    def _detect_row_color(self, page, row_idx: int, rects: List) -> str:
        """
        Определяет цвет фона строки (желтый или зеленый)
        """
        # Paysera использует заливку ячеек - анализируем объекты заливки
        for rect in rects:
            if rect.get('fill'):
                # Получаем RGB значение заливки
                fill_color = rect.get('fill')
                if self._is_yellow(fill_color):
                    return 'yellow'
                elif self._is_green(fill_color):
                    return 'green'
        
        return 'none'
    
    def _is_yellow(self, color: Tuple) -> bool:
        """Проверка желтого цвета (FFFF00 или близкие оттенки)"""
        if not color or len(color) < 3:
            return False
        r, g, b = color[:3]
        # Желтый: R высокий, G высокий, B низкий
        return r > 200 and g > 200 and b < 100
    
    def _is_green(self, color: Tuple) -> bool:
        """Проверка зеленого цвета"""
        if not color or len(color) < 3:
            return False
        r, g, b = color[:3]
        # Зеленый: R низкий, G высокий, B низкий
        return r < 100 and g > 150 and b < 100
    
    def _parse_transaction_row(self, row: List, color_type: str, 
                               page_num: int) -> Dict:
        """
        Парсит строку транзакции и извлекает:
        - Получателя
        - Сумму
        - Дату
        - Последние 4 цифры карты
        """
        if len(row) < 4:
            return None
        
        try:
            # Типовая структура Paysera: Дата | Получатель | Сумма | Карта
            transaction = {
                'date': row[0] if row[0] else 'N/A',
                'recipient': row[1] if row[1] else 'N/A',
                'amount': row[2] if row[2] else '0',
                'card_last_four': self._extract_card_last_four(
                    row[-1] if row[-1] else ''
                ),
                'color_type': color_type,
                'page': page_num + 1
            }
            
            return transaction
        
        except (IndexError, ValueError):
            return None
    
    def _extract_card_last_four(self, card_str: str) -> str:
        """Извлекает последние 4 цифры карты из строки"""
        digits = re.findall(r'\d{4}', card_str)
        return digits[-1] if digits else 'Unknown'
    
    def process_directory(self, directory: str) -> List[Dict]:
        """Обрабатывает все PDF файлы в директории"""
        pdf_files = list(Path(directory).glob('*.pdf'))
        all_transactions = []
        
        for pdf_file in pdf_files:
            print(f"Обработка: {pdf_file.name}")
            transactions = self.extract_colored_text(str(pdf_file))
            all_transactions.extend(transactions)
        
        return all_transactions
