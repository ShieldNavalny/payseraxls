from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from pathlib import Path
from typing import List, Dict
from collections import defaultdict

class ExcelWriter:
    def __init__(self):
        self.workbook = Workbook()
        # Удаляем дефолтный лист
        if 'Sheet' in self.workbook.sheetnames:
            del self.workbook['Sheet']
        
        self.red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
        self.purple_fill = PatternFill(start_color="FF9933FF", end_color="FF9933FF", fill_type="solid")
        self.yellow_fill = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
        self.white_font = Font(color="FFFFFFFF", bold=True, size=11)
        self.black_font = Font(color="FF000000", bold=True, size=11)
        self.border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
    
    def write_transactions(self, transactions: List[Dict], output_path: str, mode: str = 'single'):
        """
        Записывает транзакции в Excel
        mode: 'single' - все в один лист, 'multiple' - по файлам
        """
        if mode == 'multiple':
            self._write_multiple_sheets(transactions, output_path)
        else:
            self._write_single_sheet(transactions, output_path, "Все транзакции")
    
    def _write_multiple_sheets(self, transactions: List[Dict], output_path: str):
        """Создаёт отдельный лист для каждого PDF файла"""
        # Группируем по страницам (номер страницы = номер PDF если у вас один PDF на страницу)
        # Лучше добавить поле 'source_file' в транзакцию
        
        # Временное решение: группируем все
        sheets = defaultdict(list)
        
        for t in transactions:
            # Если есть поле source_file, используем его
            sheet_name = t.get('source_file', 'Файл 1')
            sheets[sheet_name].append(t)
        
        # Создаём лист для каждой группы
        for idx, (sheet_name, sheet_transactions) in enumerate(sheets.items()):
            # Обрезаем имя до 31 символа (лимит Excel)
            safe_name = sheet_name[:31]
            
            worksheet = self.workbook.create_sheet(title=safe_name)
            self._populate_sheet(worksheet, sheet_transactions)
        
        # Добавляем сводный лист в начало
        summary_sheet = self.workbook.create_sheet(title="Сводка", index=0)
        self._create_summary_sheet(summary_sheet, sheets)
        
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        self.workbook.save(output_path)
        print(f"✓ Отчет: {output_path}")
    
    def _write_single_sheet(self, transactions: List[Dict], output_path: str, sheet_name: str):
        """Все транзакции в один лист"""
        worksheet = self.workbook.create_sheet(title=sheet_name)
        self._populate_sheet(worksheet, transactions)
        
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        self.workbook.save(output_path)
        print(f"✓ Отчет: {output_path}")
    
    def _populate_sheet(self, worksheet, transactions: List[Dict]):
        """Заполняет лист данными"""
        headers = ['Дата', 'Получатель', 'Сумма', 'Карта', 'Комментарий']
        
        # Заголовки
        for col_num, header in enumerate(headers, 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.value = header
            cell.font = Font(bold=True, color="FFFFFFFF", size=12)
            cell.fill = PatternFill(start_color="FF2C3E50", end_color="FF2C3E50", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = self.border
        
        # Данные
        for row_num, t in enumerate(transactions, 2):
            is_negative = t['amount'].startswith('-')
            fill = self.red_fill if is_negative else self.purple_fill
            
            row_data = [
                t['date'],
                t['recipient'],
                t['amount'],
                t['card_last_four'] if t['card_last_four'] != 'N/A' else '',
                t['comment']
            ]
            
            for col_num, value in enumerate(row_data, 1):
                cell = worksheet.cell(row=row_num, column=col_num)
                cell.value = value
                cell.fill = fill
                cell.font = self.white_font
                cell.border = self.border
                cell.alignment = Alignment(
                    horizontal="left" if col_num in [2, 5] else "center",
                    vertical="center"
                )
        
        # ДОБАВЛЯЕМ ИТОГИ
        last_row = len(transactions) + 2
        self._add_totals(worksheet, transactions, last_row)
        
        # Ширина колонок
        worksheet.column_dimensions['A'].width = 13
        worksheet.column_dimensions['B'].width = 40
        worksheet.column_dimensions['C'].width = 15
        worksheet.column_dimensions['D'].width = 10
        worksheet.column_dimensions['E'].width = 30
        
        worksheet.freeze_panes = "A2"
    
    def _add_totals(self, worksheet, transactions: List[Dict], start_row: int):
        """Добавляет итоговые строки с суммами"""
        # Подсчет
        expenses_total = 0.0
        returns_total = 0.0
        
        for t in transactions:
            amount_str = t['amount'].replace(' EUR', '').replace(',', '.')
            amount = float(amount_str)
            
            if amount < 0:
                expenses_total += amount
            else:
                returns_total += amount
        
        # Пустая строка
        start_row += 1
        
        # Траты
        cell_label = worksheet.cell(row=start_row, column=2)
        cell_label.value = "Итого траты:"
        cell_label.font = self.black_font
        cell_label.alignment = Alignment(horizontal="right")
        
        cell_value = worksheet.cell(row=start_row, column=3)
        cell_value.value = f"{expenses_total:.2f} EUR"
        cell_value.fill = self.yellow_fill
        cell_value.font = self.black_font
        cell_value.border = self.border
        cell_value.alignment = Alignment(horizontal="center")
        
        # Возвраты
        start_row += 1
        cell_label = worksheet.cell(row=start_row, column=2)
        cell_label.value = "Итого возвраты:"
        cell_label.font = self.black_font
        cell_label.alignment = Alignment(horizontal="right")
        
        cell_value = worksheet.cell(row=start_row, column=3)
        cell_value.value = f"{returns_total:.2f} EUR"
        cell_value.fill = self.yellow_fill
        cell_value.font = self.black_font
        cell_value.border = self.border
        cell_value.alignment = Alignment(horizontal="center")
    
    def _create_summary_sheet(self, worksheet, sheets_data: Dict):
        """Создаёт сводный лист с общей статистикой"""
        headers = ['Файл', 'Траты', 'Возвраты', 'Всего транзакций']
        
        # Заголовки
        for col_num, header in enumerate(headers, 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.value = header
            cell.font = Font(bold=True, color="FFFFFFFF", size=12)
            cell.fill = PatternFill(start_color="FF2C3E50", end_color="FF2C3E50", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = self.border
        
        # Данные по файлам
        row_num = 2
        total_expenses = 0.0
        total_returns = 0.0
        total_count = 0
        
        for file_name, transactions in sheets_data.items():
            expenses = sum(1 for t in transactions if t['amount'].startswith('-'))
            returns = len(transactions) - expenses
            
            expenses_sum = sum(
                float(t['amount'].replace(' EUR', '').replace(',', '.'))
                for t in transactions if t['amount'].startswith('-')
            )
            returns_sum = sum(
                float(t['amount'].replace(' EUR', '').replace(',', '.'))
                for t in transactions if not t['amount'].startswith('-')
            )
            
            worksheet.cell(row=row_num, column=1).value = file_name
            worksheet.cell(row=row_num, column=2).value = f"{expenses_sum:.2f} EUR ({expenses})"
            worksheet.cell(row=row_num, column=3).value = f"{returns_sum:.2f} EUR ({returns})"
            worksheet.cell(row=row_num, column=4).value = len(transactions)
            
            total_expenses += expenses_sum
            total_returns += returns_sum
            total_count += len(transactions)
            
            row_num += 1
        
        # Итоговая строка
        row_num += 1
        worksheet.cell(row=row_num, column=1).value = "ИТОГО:"
        worksheet.cell(row=row_num, column=1).font = Font(bold=True, size=12)
        
        cell = worksheet.cell(row=row_num, column=2)
        cell.value = f"{total_expenses:.2f} EUR"
        cell.fill = self.yellow_fill
        cell.font = self.black_font
        cell.border = self.border
        
        cell = worksheet.cell(row=row_num, column=3)
        cell.value = f"{total_returns:.2f} EUR"
        cell.fill = self.yellow_fill
        cell.font = self.black_font
        cell.border = self.border
        
        cell = worksheet.cell(row=row_num, column=4)
        cell.value = total_count
        cell.fill = self.yellow_fill
        cell.font = self.black_font
        cell.border = self.border
        
        # Ширина колонок
        worksheet.column_dimensions['A'].width = 40
        worksheet.column_dimensions['B'].width = 25
        worksheet.column_dimensions['C'].width = 25
        worksheet.column_dimensions['D'].width = 20
