from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from pathlib import Path
from typing import List, Dict

class ExcelWriter:
    def __init__(self):
        self.workbook = Workbook()
        self.worksheet = self.workbook.active
        self.worksheet.title = "Проблемные транзакции"
        
        # Определяем стили
        self.red_fill = PatternFill(
            start_color="FFFF0000", 
            end_color="FFFF0000", 
            fill_type="solid"
        )
        self.purple_fill = PatternFill(
            start_color="FF800080", 
            end_color="FF800080", 
            fill_type="solid"
        )
        self.white_font = Font(color="FFFFFFFF", bold=True)
        self.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
    
    def write_transactions(self, transactions: List[Dict], 
                          output_path: str) -> None:
        """
        Записывает транзакции в Excel с цветовой разметкой
        """
        # Заголовки
        headers = ['Дата', 'Получатель', 'Сумма', 'Карта (*****)', 
                   'Тип проблемы', 'Страница PDF']
        
        for col_num, header in enumerate(headers, 1):
            cell = self.worksheet.cell(row=1, column=col_num)
            cell.value = header
            cell.font = Font(bold=True, color="FFFFFFFF")
            cell.fill = PatternFill(
                start_color="FF404040", 
                end_color="FF404040", 
                fill_type="solid"
            )
            cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # Данные транзакций
        for row_num, transaction in enumerate(transactions, 2):
            row_cells = [
                transaction['date'],
                transaction['recipient'],
                transaction['amount'],
                f"****{transaction['card_last_four']}",
                'Недостаток счетов' if transaction['color_type'] == 'yellow' 
                    else 'Возврат без документов',
                transaction['page']
            ]
            
            # Выбираем цвет в зависимости от типа
            fill_color = (self.red_fill 
                         if transaction['color_type'] == 'yellow' 
                         else self.purple_fill)
            
            for col_num, value in enumerate(row_cells, 1):
                cell = self.worksheet.cell(row=row_num, column=col_num)
                cell.value = value
                cell.fill = fill_color
                cell.font = self.white_font
                cell.border = self.border
                cell.alignment = Alignment(horizontal="left", vertical="center")
        
        # Автоширина колонок
        column_widths = [12, 30, 12, 15, 20, 12]
        for col_num, width in enumerate(column_widths, 1):
            self.worksheet.column_dimensions[
                chr(64 + col_num)
            ].width = width
        
        # Фиксируем заголовок
        self.worksheet.freeze_panes = "A2"
        
        # Сохраняем
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        self.workbook.save(output_path)
        print(f"✓ Отчет сохранен: {output_path}")
