import csv
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def translate_paysera_value(column_name, value):
    """Переводит литовские термины Paysera на русский язык и чистит текст."""
    if not value:
        return ""
    val_str = str(value).strip()
    
    if column_name == "Tipas":
        if "Mokėjimo kortelės transakcija" in val_str:
            return "Транзакция по карте"
        if "Pervedimas" in val_str:
            return "Перевод"
            
    if column_name == "Kreditas / Debetas":
        if val_str == "D": return "Д (Списание)"
        if val_str == "K": return "К (Зачисление)"
        
    if column_name == "Paskirtis":
        replacements = {
            "Pirkinys": "Покупка",
            "Pirkinio grąžinimas": "Возврат покупки",
            "palaikymo mokestis": "комиссия за обслуживание",
            "Automatiškai sugeneruotas mokėjimo nurodymas": "Автоматически сгенерированное поручение",
            "Nr.": "№"
        }
        for lt, ru in replacements.items():
            val_str = val_str.replace(lt, ru)
            
    return val_str

def parse_csv_to_excel(input_csv_path, output_excel_path):
    # Задаем желаемый порядок оригинальных столбцов (ключей из CSV)
    # Приоритет: Дата, Получатель, Сумма. Менее важное убрано в конец.
    column_order_keys = [
        "Data ir laikas", 
        "Gavėjas / Mokėtojas", 
        "Suma ir valiuta", 
        "Tipas",
        "Paskirtis",
        "Valiutos",
        "Kreditas / Debetas",
        "Likutis",
        "EVP / IBAN",
        "Kodas",
        "Įmokos kodas",
        "Išrašo nr.",
        "Pervedimo nr."
    ]

    # Соответствующие русские заголовки
    headers_ru = [
        "Дата и время", 
        "Получатель / Плательщик", 
        "Сумма", 
        "Тип транзакции",
        "Назначение платежа",
        "Валюта",
        "Д/К",
        "Остаток",
        "EVP / IBAN",
        "Код получателя",
        "Код взноса",
        "№ выписки",
        "№ перевода"
    ]

    # Инициализация Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Выписка Paysera"
    ws.views.sheetView[0].showGridLines = True

    # Стили
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    refund_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    data_font = Font(name="Segoe UI", size=10, color="000000")
    expense_font = Font(name="Segoe UI", size=10, color="C00000") 

    thin_border = Side(border_style="thin", color="D9D9D9")
    data_border = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # Пишем заголовки
    ws.append(headers_ru)
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center

    # Читаем CSV
    with open(input_csv_path, mode='r', encoding='utf-8') as f:
        reader = csv.reader(f, delimiter=';')
        # Получаем оригинальные заголовки CSV из первой строки
        csv_headers = [h.replace('"', '').strip() for h in next(reader)]

        row_idx = 2
        for row in reader:
            # Создаем словарь {Оригинальный заголовок: Значение} для текущей строки
            row_dict = {csv_headers[i]: val.replace('"', '').strip() for i, val in enumerate(row) if i < len(csv_headers)}
            
            # Логика определения возврата (остается неизменной)
            is_card_trans = "kortelės" in row_dict.get("Tipas", "").lower()
            is_credit = row_dict.get("Kreditas / Debetas", "") == "K"
            try:
                amt = float(row_dict.get("Suma ir valiuta", "0"))
            except ValueError:
                amt = 0.0
            is_refund = is_card_trans and (is_credit or amt > 0)
            
            # Формируем строку строго в новом порядке column_order_keys
            translated_row = []
            for key in column_order_keys:
                val = row_dict.get(key, "")
                
                # Конвертируем числа, если это столбцы сумм
                if key in ["Suma ir valiuta", "Likutis"]:
                    try:
                        translated_row.append(float(val))
                    except ValueError:
                        translated_row.append(val)
                else:
                    translated_row.append(translate_paysera_value(key, val))
            
            ws.append(translated_row)
            
            # Применяем форматирование к ячейкам новой строки
            for col_idx, cell in enumerate(ws[row_idx], start=1):
                cell.font = data_font
                cell.border = data_border
                
                if is_refund:
                    cell.fill = refund_fill
                    
                current_key = column_order_keys[col_idx-1]
                
                # Форматирование в зависимости от типа данных
                if current_key in ["Išrašo nr.", "Pervedimo nr.", "Kodas", "Įmokos kodas"]:
                    cell.alignment = align_center
                    cell.number_format = "@" 
                elif current_key in ["Data ir laikas", "Kreditas / Debetas", "Valiutos"]:
                    cell.alignment = align_center
                elif current_key in ["Suma ir valiuta", "Likutis"]:
                    cell.alignment = align_right
                    cell.number_format = "#,##0.00"
                    # Красный цвет для трат (отрицательные суммы)
                    if current_key == "Suma ir valiuta" and not is_refund and isinstance(cell.value, (int, float)) and cell.value < 0:
                        cell.font = expense_font
                else:
                    cell.alignment = align_left
                    
            row_idx += 1

    # Автоматическая ширина столбцов
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        # Ограничиваем максимальную ширину столбца, чтобы Назначение платежа не было бесконечным
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 60)

    # Закрепляем шапку
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers_ru))}{row_idx-1}"

    wb.save(output_excel_path)
    print(f"Файл успешно сохранен как {output_excel_path}")

# Пример использования
if __name__ == "__main__":
    parse_csv_to_excel("input.csv", "paysera_ru.xlsx")