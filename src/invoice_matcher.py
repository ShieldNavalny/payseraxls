"""
invoice_matcher.py
Матчинг PDF-счетов из input_invoices/ против транзакций в XLS.

Стратегия матчинга (достаточно любого одного):
  A) Точная сумма + месяц
  B) Вендор (fuzzy) + месяц
  C) Приближенная сумма (±2%) + месяц

Результат:
  - XLS: совпавшие строки подсвечиваются тёмно-зелёным, добавляется колонка "Счёт (файл)"
  - PDF: совпавшая сумма подсвечивается жёлтым прямоугольником

Требует:
  pip install pdfplumber pymupdf pytesseract Pillow rapidfuzz
  + Tesseract OCR на системе (https://github.com/UB-Mannheim/tesseract/wiki для Windows)
"""

from __future__ import annotations

import io
import re
import sys
from pathlib import Path
from typing import Dict, List, NamedTuple, Optional, Tuple

# ставим src/ в sys.path чтобы работал и как модуль и напрямую
sys.path.insert(0, str(Path(__file__).parent.parent))

import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

try:
    import pdfplumber
except ImportError:
    pdfplumber = None  # type: ignore

try:
    import fitz  # pymupdf
except ImportError:
    fitz = None  # type: ignore

try:
    import pytesseract
    from PIL import Image
except ImportError:
    pytesseract = None  # type: ignore
    Image = None  # type: ignore

try:
    from rapidfuzz import fuzz
except ImportError:
    fuzz = None  # type: ignore


# ---------------------------------------------------------------------------
# Константы
# ---------------------------------------------------------------------------

AMOUNT_TOLERANCE   = 0.02   # 2%
FUZZY_VENDOR_SCORE = 70      # порог rapidfuzz partial_ratio
HIGHLIGHT_COLOR    = (1, 1, 0)   # RGB 0–1 — жёлтый в PDF

FILL_MATCHED = PatternFill(start_color="1E7B34", end_color="1E7B34", fill_type="solid")  # тёмно-зелёный
FONT_MATCHED = Font(name="Segoe UI", size=10, color="FFFFFF", bold=True)

# Регекспы для извлечения данных из PDF
_RE_AMOUNT = re.compile(
    r"(?:EUR|USD|GBP|\$|\u20ac|\u00a3)?\s*"
    r"(\d{1,6}[.,]\d{2})"
    r"(?:\s*(?:EUR|USD|GBP|\$|\u20ac|\u00a3))?",
    re.IGNORECASE,
)
_RE_DATE = re.compile(
    r"(?:"
    r"(\d{4})[.\-/](\d{2})[.\-/](\d{2})"   # YYYY-MM-DD
    r"|(\d{2})[.\-/](\d{2})[.\-/](\d{4})"   # DD.MM.YYYY / DD-MM-YYYY
    r")"
)


# ---------------------------------------------------------------------------
# Структуры
# ---------------------------------------------------------------------------

class PdfRecord(NamedTuple):
    pdf_path:  Path
    page_num:  int          # 0-based
    amount:    float
    year_month: str         # 'YYYY-MM'
    raw_text:  str          # весь текст страницы
    amount_bbox: Optional[Tuple[float, float, float, float]]  # для подсветки PDF


class XlsRow(NamedTuple):
    row_idx:    int
    sheet_name: str
    date:       str    # ISO
    year_month: str    # 'YYYY-MM'
    vendor:     str
    amount:     float  # отрицательное число


# ---------------------------------------------------------------------------
# Извлечение текста из PDF
# ---------------------------------------------------------------------------

def _extract_text_pdfplumber(pdf_path: Path) -> List[Tuple[int, str]]:
    """Возвращает [(page_idx, text), ...]"""
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            pages.append((i, text))
    return pages


def _extract_text_ocr(pdf_path: Path) -> List[Tuple[int, str]]:
    """Фоллбэк через pytesseract для сканов."""
    if fitz is None:
        raise ImportError("pymupdf не установлен")
    if pytesseract is None:
        raise ImportError("pytesseract не установлен")

    doc = fitz.open(pdf_path)
    pages = []
    for i, page in enumerate(doc):
        mat  = fitz.Matrix(2, 2)   # 2x для качества OCR
        pix  = page.get_pixmap(matrix=mat, alpha=False)
        img  = Image.open(io.BytesIO(pix.tobytes("png")))
        text = pytesseract.image_to_string(img, lang="eng+lit+rus")
        pages.append((i, text))
    doc.close()
    return pages


def extract_pdf_pages(pdf_path: Path) -> List[Tuple[int, str]]:
    """Сначала pdfplumber, если текста мало — OCR."""
    if pdfplumber is None:
        return _extract_text_ocr(pdf_path)

    pages = _extract_text_pdfplumber(pdf_path)
    # если среднее кол-во символов на странице < 30 — скан
    avg_len = sum(len(t) for _, t in pages) / max(len(pages), 1)
    if avg_len < 30 and pytesseract is not None:
        print(f"  [{pdf_path.name}] мало текста — переключаюсь на OCR")
        return _extract_text_ocr(pdf_path)
    return pages


# ---------------------------------------------------------------------------
# Парсинг данных из текста страницы
# ---------------------------------------------------------------------------

def _parse_year_month(text: str) -> Optional[str]:
    """'YYYY-MM' из первой найденной даты."""
    m = _RE_DATE.search(text)
    if not m:
        return None
    if m.group(1):   # YYYY-MM-DD
        return f"{m.group(1)}-{m.group(2)}"
    else:            # DD.MM.YYYY
        return f"{m.group(6)}-{m.group(5)}"


def _parse_amounts(text: str) -> List[float]:
    amounts = []
    for match in _RE_AMOUNT.finditer(text):
        raw = match.group(1).replace(',', '.')
        try:
            amounts.append(float(raw))
        except ValueError:
            pass
    return amounts


def parse_pdf_records(pdf_path: Path) -> List[PdfRecord]:
    """Возвращает все PdfRecord из одного PDF-файла."""
    records: List[PdfRecord] = []
    pages = extract_pdf_pages(pdf_path)

    for page_num, text in pages:
        year_month = _parse_year_month(text)
        if not year_month:
            continue

        amounts = _parse_amounts(text)
        for amount in amounts:
            if amount < 0.5:   # игнорируем копеечные суммы
                continue
            records.append(PdfRecord(
                pdf_path   = pdf_path,
                page_num   = page_num,
                amount     = amount,
                year_month = year_month,
                raw_text   = text,
                amount_bbox = None,  # заполняется при подсветке
            ))
    return records


# ---------------------------------------------------------------------------
# Чтение XLS
# ---------------------------------------------------------------------------

def _year_month_from_cell(val) -> str:
    if hasattr(val, 'strftime'):
        return val.strftime('%Y-%m')
    s = str(val).strip()
    m = re.match(r'^(\d{4})-(\d{2})', s)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    m = re.match(r'^(\d{2})\.(\d{2})\.(\d{4})', s)
    if m:
        return f"{m.group(3)}-{m.group(2)}"
    return s


def load_xls_rows(xls_path: Path) -> List[XlsRow]:
    wb = openpyxl.load_workbook(xls_path, data_only=True)
    rows: List[XlsRow] = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            date_val, vendor_val, amount_val = row[0], row[1], row[2]
            if not date_val or amount_val is None:
                continue
            try:
                amt = float(str(amount_val).replace(',', '.').replace('\xa0', '').strip())
            except ValueError:
                continue
            if amt >= 0:
                continue   # только расходы
            rows.append(XlsRow(
                row_idx    = row_idx,
                sheet_name = sheet_name,
                date       = str(date_val),
                year_month = _year_month_from_cell(date_val),
                vendor     = str(vendor_val or "").strip(),
                amount     = amt,
            ))
    wb.close()
    return rows


# ---------------------------------------------------------------------------
# Матчинг
# ---------------------------------------------------------------------------

def _amounts_close(a: float, b: float) -> bool:
    """True если разница не больше AMOUNT_TOLERANCE."""
    if b == 0:
        return False
    return abs(a - b) / abs(b) <= AMOUNT_TOLERANCE


def _vendor_in_text(vendor: str, text: str) -> bool:
    """True если вендор fuzzy-совпадает с частью текста."""
    if not vendor or len(vendor) < 3:
        return False
    if fuzz is not None:
        score = fuzz.partial_ratio(vendor.lower(), text.lower())
        return score >= FUZZY_VENDOR_SCORE
    # фоллбэк без rapidfuzz
    return vendor.lower() in text.lower()


def match_records(
    xls_rows: List[XlsRow],
    pdf_records: List[PdfRecord],
) -> Dict[int, List[PdfRecord]]:
    """
    Возвращает {xls_row_idx: [matched PdfRecord, ...]}.
    Совпадение по любому из трёх критериев (+ одинаковый месяц).
    """
    result: Dict[int, List[PdfRecord]] = {}

    for xls in xls_rows:
        xls_abs = abs(xls.amount)
        matched: List[PdfRecord] = []

        for pdf in pdf_records:
            if pdf.year_month != xls.year_month:
                continue   # месяц обязателен

            a_exact  = abs(pdf.amount - xls_abs) < 0.01          # A
            a_fuzzy  = _amounts_close(pdf.amount, xls_abs)        # C
            v_match  = _vendor_in_text(xls.vendor, pdf.raw_text)  # B

            if a_exact or (a_fuzzy and not a_exact) or v_match:
                matched.append(pdf)

        if matched:
            result[xls.row_idx] = matched

    return result


# ---------------------------------------------------------------------------
# Подсветка XLS
# ---------------------------------------------------------------------------

def annotate_xls(
    xls_path: Path,
    matches: Dict[int, List[PdfRecord]],
    output_path: Path,
) -> None:
    """
    Добавляет колонку 'Счёт (файл)' и подсвечивает строки тёмно-зелёным.
    """
    wb = openpyxl.load_workbook(xls_path)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        # Находим или создаём колонку 'Счёт (файл)'
        header_row = [cell.value for cell in ws[1]]
        if "Счёт (файл)" in header_row:
            invoice_col = header_row.index("Счёт (файл)") + 1
        else:
            invoice_col = ws.max_column + 1
            ws.cell(row=1, column=invoice_col, value="Счёт (файл)")
            # копируем стиль заголовка
            hdr_src = ws.cell(row=1, column=1)
            ws.cell(row=1, column=invoice_col).font = hdr_src.font
            ws.cell(row=1, column=invoice_col).fill = hdr_src.fill
            ws.cell(row=1, column=invoice_col).alignment = hdr_src.alignment
            ws.column_dimensions[get_column_letter(invoice_col)].width = 40

        for row_idx, pdf_list in matches.items():
            # имена файлов через '; '
            file_names = "; ".join(
                sorted({p.pdf_path.name for p in pdf_list})
            )

            row_cells = list(ws.iter_rows(
                min_row=row_idx, max_row=row_idx,
                min_col=1, max_col=invoice_col
            ))[0]

            # подсветка + шрифт
            for cell in row_cells:
                cell.fill = FILL_MATCHED
                cell.font = FONT_MATCHED

            # записываем имя файла
            ws.cell(row=row_idx, column=invoice_col).value = file_names

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    wb.close()
    print(f"✓ XLS сохранён: {output_path}")


# ---------------------------------------------------------------------------
# Подсветка PDF
# ---------------------------------------------------------------------------

def _find_amount_rects(
    page: "fitz.Page",
    amount: float,
) -> List["fitz.Rect"]:
    """Ищем все вхождения суммы (XXXX.XX / XX,XX) на странице."""
    rects = []
    patterns = [
        f"{amount:.2f}",
        f"{amount:.2f}".replace('.', ','),
    ]
    # также без десятичных, если сумма целая
    if amount == int(amount):
        patterns += [f"{int(amount)}", f"{int(amount)}.00", f"{int(amount)},00"]
    for pat in patterns:
        hits = page.search_for(pat)
        rects.extend(hits)
    return rects


def highlight_pdfs(
    matches: Dict[int, List[PdfRecord]],
    output_dir: Path,
) -> None:
    """
    Сохраняет отмеченные PDF в output_dir/highlighted_pdfs/.
    """
    if fitz is None:
        print("⚠  pymupdf не установлен — подсветка PDF пропущена")
        return

    # группируем все совпадения по файлу
    pdf_hits: Dict[Path, Dict[int, List[float]]] = {}  # pdf_path -> {page: [amounts]}
    for pdf_list in matches.values():
        for rec in pdf_list:
            pdf_hits.setdefault(rec.pdf_path, {}).setdefault(
                rec.page_num, []
            ).append(rec.amount)

    out_pdf_dir = output_dir / "highlighted_pdfs"
    out_pdf_dir.mkdir(parents=True, exist_ok=True)

    for pdf_path, pages_dict in pdf_hits.items():
        doc = fitz.open(pdf_path)
        for page_num, amounts in pages_dict.items():
            page = doc[page_num]
            for amount in amounts:
                rects = _find_amount_rects(page, amount)
                for rect in rects:
                    # жёлтая полупрозрачная подсветка
                    annot = page.add_highlight_annot(rect)
                    annot.set_colors(stroke=HIGHLIGHT_COLOR)
                    annot.update()

        dst = out_pdf_dir / pdf_path.name
        doc.save(dst, garbage=4, deflate=True)
        doc.close()
        print(f"✓ PDF отмечен: {dst}")


# ---------------------------------------------------------------------------
# Главная функция
# ---------------------------------------------------------------------------

def run_matching(
    invoices_dir: Path,
    xls_path: Path,
    output_dir: Path,
) -> None:
    print("🔍 Invoice Matcher")
    print("=" * 50)
    print(f"  PDF:    {invoices_dir}")
    print(f"  XLS:    {xls_path}")
    print(f"  Вывод:  {output_dir}")
    print()

    if not invoices_dir.exists():
        print(f"✗ Папка не найдена: {invoices_dir}")
        return
    if not xls_path.exists():
        print(f"✗ XLS не найден: {xls_path}")
        return

    # 1. Читаем XLS
    xls_rows = load_xls_rows(xls_path)
    print(f"✓ XLS: {len(xls_rows)} строк расходов")

    # 2. Читаем PDFс
    pdf_files = sorted(invoices_dir.glob("*.pdf"))
    print(f"✓ PDF: {len(pdf_files)} файлов")
    print()

    all_pdf_records: List[PdfRecord] = []
    for pdf_path in pdf_files:
        print(f"  → читаю {pdf_path.name} ...", end=" ", flush=True)
        try:
            recs = parse_pdf_records(pdf_path)
            all_pdf_records.extend(recs)
            print(f"{len(recs)} записей")
        except Exception as e:
            print(f"⚠ ошибка: {e}")

    print(f"\n✓ Итого записей из PDF: {len(all_pdf_records)}")

    # 3. Матчинг
    matches = match_records(xls_rows, all_pdf_records)
    print(f"✓ Совпадений: {len(matches)} строк XLS")

    # 4. Аннотация XLS
    out_xls = output_dir / (xls_path.stem + "_matched.xlsx")
    annotate_xls(xls_path, matches, out_xls)

    # 5. Подсветка PDF
    highlight_pdfs(matches, output_dir)

    print("=" * 50)
    print("✓ Готово!")
    print(f"  XLS с отметками: {out_xls}")
    print(f"  PDF с подсветкой: {output_dir / 'highlighted_pdfs'}")


# ---------------------------------------------------------------------------
# CLI и вход через main.py
# ---------------------------------------------------------------------------

def main_cli() -> None:
    import argparse
    parser = argparse.ArgumentParser(description="Invoice matcher: PDF vs XLS")
    parser.add_argument("--invoices", required=True, help="Папка с PDF-счетами")
    parser.add_argument("--xls",      required=True, help="Путь к XLS/XLSX-файлу")
    parser.add_argument("--out",      default="output", help="Папка вывода")
    args = parser.parse_args()
    run_matching(
        invoices_dir = Path(args.invoices).resolve(),
        xls_path     = Path(args.xls).resolve(),
        output_dir   = Path(args.out).resolve(),
    )


if __name__ == "__main__":
    main_cli()
