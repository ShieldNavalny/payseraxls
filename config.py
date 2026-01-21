import os
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent
INPUT_DIR = PROJECT_ROOT / "input_pdfs"
OUTPUT_DIR = PROJECT_ROOT / "output_reports"

# Цвета для выделения в XLSX (RGB)
COLOR_RED = "FFFF0000"      # Красный для расходов без счетов
COLOR_PURPLE = "FF800080"   # Фиолетовый для возвратов без документов
COLOR_GRAY = "FFF0F0F0"     # Серый для остальных

# Пороговые значения для обнаружения цветов в PDF (HSV диапазоны)
YELLOW_HUE_RANGE = (40, 70)      # Желтый диапазон
GREEN_HUE_RANGE = (80, 140)      # Зеленый диапазон

# Настройки вывода
OUTPUT_FILENAME = "miss.xlsx"

# Режим вывода: 'single' (один лист) или 'multiple' (по файлам)
DEFAULT_OUTPUT_MODE = 'multiple'