import PyPDF2
import pdfplumber
from pathlib import Path
from typing import List, Dict, Optional
import re

class PDFParser:
    def __init__(self):
        self.transactions = []
    
    def _extract_month_year_from_filename(self, filename: str) -> str:
        """
        Извлекает месяц и год из имени файла
        EVP0510014895802_2025-01-01_2025-01-31- truksta dok -> 01_2025
        """
        # Ищем паттерн ГГГГ-ММ-ДД
        match = re.search(r'(\d{4})-(\d{2})-\d{2}', filename)
        
        if match:
            year = match.group(1)
            month = match.group(2)
            return f"{month}_{year}"
        
        # Fallback: если не нашли дату, возвращаем первые 30 символов
        return filename[:30]

    def extract_colored_text(self, pdf_path: str) -> List[Dict]:
        """Извлекает транзакции с желтым Highlight"""
        transactions = []

        # Получаем имя файла для group by
        file_name_full = Path(pdf_path).stem
        file_name = self._extract_month_year_from_filename(file_name_full)
        
        try:
            yellow_rects = self._get_yellow_highlights(pdf_path)
            
            if not yellow_rects:
                print("  Нет желтых highlights")
                return transactions
            
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    
                    if not tables:
                        continue
                    
                    table = tables[0]
                    page_yellows = yellow_rects.get(page_num, [])
                    
                    if not page_yellows:
                        continue
                    
                    words = page.extract_words()
                    
                    
                    for row_idx, row in enumerate(table):
                        if self._row_has_yellow_by_coords(row, page_yellows, words):
                            transaction = self._parse_transaction_row(
                                row, page_num, table, row_idx
                            )
                            if transaction:
                                transaction['source_file'] = file_name 
                                transactions.append(transaction)
        
        except Exception as e:
            print(f"Ошибка: {e}")
            import traceback
            traceback.print_exc()
        
        return self._deduplicate(transactions)
    
    def _get_yellow_highlights(self, pdf_path: str) -> Dict[int, List]:
        """Извлекает координаты желтых Highlight"""
        highlights = {}
        
        with open(pdf_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            
            for page_num, page in enumerate(reader.pages):
                if '/Annots' not in page:
                    continue
                
                page_highlights = []
                annots = page['/Annots']
                
                for annot in annots:
                    annot_obj = annot.get_object()
                    
                    if annot_obj.get('/Subtype') == '/Highlight':
                        color = annot_obj.get('/C')
                        if color and self._is_yellow(color):
                            rect = annot_obj.get('/Rect')
                            if rect:
                                page_height = float(page.mediabox.height)
                                
                                pdf_y0 = float(rect[1])
                                pdf_y1 = float(rect[3])
                                
                                plumber_y0 = page_height - pdf_y1
                                plumber_y1 = page_height - pdf_y0
                                
                                page_highlights.append({
                                    'x0': float(rect[0]),
                                    'y0': plumber_y0,
                                    'x1': float(rect[2]),
                                    'y1': plumber_y1
                                })
                
                if page_highlights:
                    highlights[page_num] = page_highlights
                    print(f"  Страница {page_num}: {len(page_highlights)} желтых")
        
        return highlights
    
    def _is_yellow(self, color) -> bool:
        """Желтый: [1, 0.918, 0]"""
        if len(color) >= 3:
            r, g, b = float(color[0]), float(color[1]), float(color[2])
            return r > 0.9 and g > 0.8 and b < 0.1
        return False
    
    def _row_has_yellow_by_coords(self, row: List, highlights: List, words: List) -> bool:
        """Проверяет пересечение строки с желтым highlight"""
        if not highlights or not row:
            return False
        
        row_text = ' '.join([str(cell) for cell in row if cell])
        
        if not row_text or len(row_text) < 10:
            return False
        
        # Должна содержать сумму EUR
        if not re.search(r'-?\d+\.\d{2}\s*EUR', row_text):
            return False
        
        # Исключаем итоги
        if self._is_summary_row(row_text):
            return False
        
        # Ключевые слова
        key_words = self._extract_key_words(row_text)
        
        if not key_words:
            return False
        
        # Координаты слов
        row_y_coords = []
        for word_obj in words:
            if any(kw in word_obj['text'] for kw in key_words):
                row_y_coords.append({
                    'y0': word_obj['top'],
                    'y1': word_obj['bottom'],
                    'x0': word_obj['x0'],
                    'x1': word_obj['x1']
                })
        
        if not row_y_coords:
            return False
        
        # Проверка пересечения
        for word_coord in row_y_coords:
            for hl in highlights:
                if self._boxes_intersect(word_coord, hl):
                    return True
        
        return False
    
    def _is_summary_row(self, text: str) -> bool:
        """Итоговые строки"""
        keywords = [
            'Likutis pradžioje',
            'Likutis pabaigoje',
            'Debeto apyvarta',
            'Kredito apyvarta',
            'Komisinis mokestis',
            'Pervedimo mokestis'
        ]
        return any(k in text for k in keywords)
    
    def _extract_key_words(self, text: str) -> List[str]:
        """Уникальные слова для поиска"""
        words = []
        
        # Номер 10+ цифр
        m = re.search(r'\d{10}', text)
        if m:
            words.append(m.group(0))
        
        # Заглавные 3+ букв
        m = re.search(r'[A-Z]{3,}', text)
        if m:
            words.append(m.group(0))
        
        # Сумма
        m = re.search(r'-?\d+\.\d{2}', text)
        if m:
            words.append(m.group(0))
        
        return words
    
    def _boxes_intersect(self, box1: Dict, box2: Dict) -> bool:
        """Пересечение bbox"""
        y_overlap = not (box1['y1'] < box2['y0'] or box2['y1'] < box1['y0'])
        x_overlap = not (box1['x1'] < box2['x0'] or box2['x1'] < box1['x0'])
        return y_overlap and x_overlap
    
    def _parse_transaction_row(self, row: List, page_num: int, 
                           table: List, row_idx: int) -> Optional[Dict]:
        """Парсит строку"""
        if not row or len(row) < 3:
            return None
        
        full_text = '\n'.join([str(cell) for cell in row if cell])
        
        try:
            date = self._extract_date(full_text)
            amount = self._extract_amount(full_text)
            recipient = self._extract_recipient(full_text, row)

            card = self._extract_card(full_text)
            if card == 'N/A' and row_idx + 1 < len(table):
                next_row = table[row_idx + 1]
                next_text = '\n'.join([str(cell) for cell in next_row if cell])
                if 'Paskirtis' in next_text or 'card' in next_text.lower():
                    card = self._extract_card(next_text)
            
            if not amount or amount == '0 EUR':
                return None
            
            return {
                'date': date,
                'recipient': recipient,
                'amount': amount,
                'card_last_four': card,
                'page': page_num + 1,
                'comment': ''
            }
        
        except:
            return None
    
    def _extract_date(self, text: str) -> str:
        """02.01.2025"""
        m = re.search(r'(\d{4})-(\d{2})-(\d{2})\s+\d{2}:\d{2}:\d{2}', text)
        if m:
            year, month, day = m.groups()
            return f"{day}.{month}.{year}"
        return 'N/A'
    
    def _extract_amount(self, text: str) -> str:
        """Сумма EUR"""
        m = re.search(r'(-?\d+\.\d{2})\s*EUR', text)
        if m:
            return f"{m.group(1)} EUR"
        return '0 EUR'
    
    def _extract_recipient(self, text: str, row: List) -> str:
        """Получатель"""
        if len(row) >= 3 and row[2]:
            recipient = str(row[2]).strip()
            recipient = re.sub(r'\(\d+\)', '', recipient)
            recipient = re.sub(r'kortelės\s+\d+\s+', '', recipient)
            
            # IBAN
            iban = re.search(r'([A-Z]{2}\d{18,})', recipient)
            if iban:
                return iban.group(1)
            
            # Домен
            domain = re.search(r'([A-Z0-9][A-Z0-9.-]+\.[A-Z]{2,})', recipient, re.IGNORECASE)
            if domain:
                return domain.group(1)
            
            return recipient.strip()[:50]
        
        return 'Unknown'
    
    def _extract_card(self, text: str) -> str:
        """*NNNN из Paskirtis"""
        # Паттерн 1: card *4569, not present
        m = re.search(r'card\s*\*(\d{4})', text, re.IGNORECASE)
        if m:
            return m.group(1)
        
        # Паттерн 2: просто *4569
        m = re.search(r'\*(\d{4})', text)
        if m:
            return m.group(1)
        
        return 'N/A'
    
    def _deduplicate(self, transactions: List[Dict]) -> List[Dict]:
        """Дубликаты"""
        seen = set()
        unique = []
        
        for t in transactions:
            key = (t['date'], t['amount'], t['recipient'])
            if key not in seen:
                seen.add(key)
                unique.append(t)
        
        return unique
    
    def process_directory(self, directory: str) -> List[Dict]:
        """Все PDF"""
        pdf_files = list(Path(directory).glob('*.pdf'))
        all_transactions = []
        
        for pdf_file in pdf_files:
            print(f"Обработка: {pdf_file.name}")
            transactions = self.extract_colored_text(str(pdf_file))
            all_transactions.extend(transactions)
            print(f"  Найдено: {len(transactions)}")
        
        return all_transactions
