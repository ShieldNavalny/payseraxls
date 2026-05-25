"""
invoice_matcher.py
Матчинг PDF-счетов из input_invoices/ против транзакций в XLS.

Стратегия матчинга (достаточно любого одного + одинаковый месяц):
  A) Точная сумма  (разница < 0.01)
  B) Вендор fuzzy  (rapidfuzz.partial_ratio ≥ FUZZY_VENDOR_SCORE)
  C) Приблизительная сумма (отклонение ≤ AMOUNT_TOLERANCE)

Результат:
  - XLS: совпавшие строки подсвечиваются тёмно-зелёным,
         добавляется колонка «Счёт (файл)» с именем PDF-файла
  - PDF: совпавшая сумма подсвечивается жёлтым прямоугольником

Настраивается через корневой config.py.
"""

from __future__ import annotations

import io
import re
import sys
from pathlib import Path
from typing import Dict, List, NamedTuple, Optional, Tuple

sys.path.insert(0, str(Path(__file__).parent.parent))
from config import (
    AMOUNT_TOLERANCE,
    FUZZY_VENDOR_SCORE,
    MIN_AMOUNT,
    PDF_HIGHLIGHT_COLOR,
    COLOR_MATCHED,
    FONT_MATCHED_COLOR,
    FONT_MATCHED_BOLD,
)

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
# Стили (берутся из config)
# ---------------------------------------------------------------------------

FILL_MATCHED = PatternFill(
    start_color=COLOR_MATCHED, end_color=COLOR_MATCHED, fill_type="solid"
)
FONT_MATCHED = Font(
    name="Segoe UI", size=10,
    color=FONT_MATCHED_COLOR,
    bold=FONT_MATCHED_BOLD,
)

# Регексп для извлечения сумм и дат из текста PDF
_RE_AMOUNT = re.compile(
    r"(?:EUR|USD|GBP|\$|\u20ac|\u00a3)?\s*(\d{1,6}[.,]\d{2})(?:\s*(?:EUR|USD|GBP|\$|\u20ac|\u00a3))?",
    re.IGNORECASE,
)
_RE_DATE = re.compile(
    r"(?:"
    r"(\d{4})[.\-/](\d{2})[.\-/](\d{2})"
    r"|(\d{2})[.\-/](\d{2})[.\-/](\d{4})"
    r")"
)


# ---------------------------------------------------------------------------
# Структуры данных
# ---------------------------------------------------------------------------

class PdfRecord(NamedTuple):
    pdf_path:   Path
    page_num:   int
    amount:     float
    year_month: str
    raw_text:   str
    amount_bbox: Optional[Tuple[float, float, float, float]]


class XlsRow(NamedTuple):
    row_idx:    int
    sheet_name: str
    date:       str
    year_month: str
    vendor:     str
    amount:     float


# ---------------------------------------------------------------------------
# Извлечение текста из PDF
# ---------------------------------------------------------------------------

def _extract_text_pdfplumber(pdf_path: Path) -> List[Tuple[int, str]]:
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            pages.append((i, text))
    return pages


def _extract_text_ocr(pdf_path: Path) -> List[Tuple[int, str]]:
    if fitz is None:
        raise ImportError("pymupdf не установлен")
    if pytesseract is None:
        raise ImportError("pytesseract не установлен")
    doc = fitz.open(pdf_path)
    pages = []
    for i, page in enumerate(doc):
        mat = fitz.Matrix(2, 2)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        text = pytesseract.image_to_string(img, lang="eng+lit+rus")
        pages.append((i, text))
    doc.close()
    return pages


def extract_pdf_pages(pdf_path: Path) -> List[Tuple[int, str]]:
    """Сначала pdfplumber, если текста мало — OCR."""
    if pdfplumber is None:
        return _extract_text_ocr(pdf_path)
    pages = _extract_text_pdfplumber(pdf_path)
    avg_len = sum(len(t) for _, t in pages) / max(len(pages), 1)
    if avg_len < 30 and pytesseract is not None:
        print(f"  [{pdf_path.name}] мало текста — переключаюсь на OCR")
        return _extract_text_ocr(pdf_path)
    return pages


# ---------------------------------------------------------------------------
# Парсинг данных из текста страницы
# ---------------------------------------------------------------------------

def _parse_year_month(text: str) -> Optional[str]:
    m = _RE_DATE.search(text)
    if not m:
        return None
    if m.group(1):   # YYYY-MM-DD
        return f"{m.group(1)}-{m.group(2)}"
    return f"{m.group(6)}-{m.group(5)}"  # DD.MM.YYYY


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
    records: List[PdfRecord] = []
    for page_num, text in extract_pdf_pages(pdf_path):
        year_month = _parse_year_month(text)
        if not year_month:
            continue
        for amount in _parse_amounts(text):
            if amount < MIN_AMOUNT:
                continue
            records.append(PdfRecord(
                pdf_path=pdf_path,
                page_num=page_num,
                amount=amount,
                year_month=year_month,
                raw_text=text,
                amount_bbox=None,
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
                continue
            rows.append(XlsRow(
                row_idx=row_idx,
                sheet_name=sheet_name,
                date=str(date_val),
                year_month=_year_month_from_cell(date_val),
                vendor=str(vendor_val or "").strip(),
                amount=amt,
            ))
    wb.close()
    return rows


# ---------------------------------------------------------------------------
# Матчинг
# ---------------------------------------------------------------------------

def _amounts_close(a: float, b: float) -> bool:
    if b == 0:
        return False
    return abs(a - b) / abs(b) <= AMOUNT_TOLERANCE


def _vendor_in_text(vendor: str, text: str) -> bool:
    if not vendor or len(vendor) < 3:
        return False
    if fuzz is not None:
        return fuzz.partial_ratio(vendor.lower(), text.lower()) >= FUZZY_VENDOR_SCORE
    return vendor.lower() in text.lower()


def match_records(
    xls_rows: List[XlsRow],
    pdf_records: List[PdfRecord],
) -> Dict[int, List[PdfRecord]]:
    result: Dict[int, List[PdfRecord]] = {}
    for xls in xls_rows:
        xls_abs = abs(xls.amount)
        matched: List[PdfRecord] = []
        for pdf in pdf_records:
            if pdf.year_month != xls.year_month:
                continue
            a_exact = abs(pdf.amount - xls_abs) < 0.01
            a_fuzzy = _amounts_close(pdf.amount, xls_abs)
            v_match = _vendor_in_text(xls.vendor, pdf.raw_text)
            if a_exact or a_fuzzy or v_match:
                matched.append(pdf)
        if matched:
            result[xls.row_idx] = matched
    return result


# ---------------------------------------------------------------------------
# Аннотация XLS
# ---------------------------------------------------------------------------

def annotate_xls(
    xls_path: Path,
    matches: Dict[int, List[PdfRecord]],
    output_path: Path,
) -> None:
    wb = openpyxl.load_workbook(xls_path)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        header_row = [cell.value for cell in ws[1]]
        if "\u0421\u0447\u0451\u0442 (\u0444\u0430\u0439\u043b)" in header_row:
            invoice_col = header_row.index("\u0421\u0447\u0451\u0442 (\u0444\u0430\u0439\u043b)") + 1
        else:
            invoice_col = ws.max_column + 1
            ws.cell(row=1, column=invoice_col, value="\u0421\u0447\u0451\u0442 (\u0444\u0430\u0439\u043b)")
            ws.column_dimensions[get_column_letter(invoice_col)].width = 40

        for row_idx, pdf_list in matches.items():
            # Показываем относительный путь от invoices_dir для читаемости
            file_names = "; ".join(sorted({str(p.pdf_path.name) for p in pdf_list}))
            row_cells = list(ws.iter_rows(
                min_row=row_idx, max_row=row_idx,
                min_col=1, max_col=invoice_col
            ))[0]
            for cell in row_cells:
                cell.fill = FILL_MATCHED
                cell.font = FONT_MATCHED
            ws.cell(row=row_idx, column=invoice_col).value = file_names

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    wb.close()
    print(f"\u2713 XLS \u0441\u043e\u0445\u0440\u0430\u043d\u0451\u043d: {output_path}")


# ---------------------------------------------------------------------------
# Подсветка PDF
# ---------------------------------------------------------------------------

def _find_amount_rects(page: "fitz.Page", amount: float) -> List["fitz.Rect"]:
    rects = []
    patterns = [f"{amount:.2f}", f"{amount:.2f}".replace('.', ',')]
    if amount == int(amount):
        patterns += [f"{int(amount)}", f"{int(amount)}.00", f"{int(amount)},00"]
    for pat in patterns:
        rects.extend(page.search_for(pat))
    return rects


def highlight_pdfs(
    matches: Dict[int, List[PdfRecord]],
    output_dir: Path,
) -> None:
    if fitz is None:
        print("\u26a0  pymupdf \u043d\u0435 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d \u2014 \u043f\u043e\u0434\u0441\u0432\u0435\u0442\u043a\u0430 PDF \u043f\u0440\u043e\u043f\u0443\u0449\u0435\u043d\u0430")
        return

    pdf_hits: Dict[Path, Dict[int, List[float]]] = {}
    for pdf_list in matches.values():
        for rec in pdf_list:
            pdf_hits.setdefault(rec.pdf_path, {}).setdefault(rec.page_num, []).append(rec.amount)

    out_pdf_dir = output_dir / "highlighted_pdfs"
    out_pdf_dir.mkdir(parents=True, exist_ok=True)

    for pdf_path, pages_dict in pdf_hits.items():
        doc = fitz.open(pdf_path)
        for page_num, amounts in pages_dict.items():
            page = doc[page_num]
            for amount in amounts:
                for rect in _find_amount_rects(page, amount):
                    annot = page.add_highlight_annot(rect)
                    annot.set_colors(stroke=PDF_HIGHLIGHT_COLOR)
                    annot.update()
        dst = out_pdf_dir / pdf_path.name
        doc.save(dst, garbage=4, deflate=True)
        doc.close()
        print(f"\u2713 PDF \u043e\u0442\u043c\u0435\u0447\u0435\u043d: {dst}")


# ---------------------------------------------------------------------------
# Главная функция
# ---------------------------------------------------------------------------

def run_matching(
    invoices_dir: Path,
    xls_path: Path,
    output_dir: Path,
) -> None:
    print("\U0001f50d Invoice Matcher")
    print("=" * 50)
    print(f"  PDF:    {invoices_dir}")
    print(f"  XLS:    {xls_path}")
    print(f"  \u0412\u044b\u0432\u043e\u0434:  {output_dir}")
    print()

    if not invoices_dir.exists():
        print(f"\u2717 \u041f\u0430\u043f\u043a\u0430 \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d\u0430: {invoices_dir}")
        return
    if not xls_path.exists():
        print(f"\u2717 XLS \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d: {xls_path}")
        return

    xls_rows = load_xls_rows(xls_path)

    # rglob — рекурсивный обход всех подпапок
    pdf_files = sorted(invoices_dir.rglob("*.pdf"))

    print(f"\u2713 XLS: {len(xls_rows)} \u0441\u0442\u0440\u043e\u043a \u0440\u0430\u0441\u0445\u043e\u0434\u043e\u0432")
    print(f"\u2713 PDF: {len(pdf_files)} \u0444\u0430\u0439\u043b\u043e\u0432 (\u0440\u0435\u043a\u0443\u0440\u0441\u0438\u0432\u043d\u043e)")
    print()

    all_pdf_records: List[PdfRecord] = []
    for pdf_path in pdf_files:
        # Показываем путь относительно invoices_dir
        rel = pdf_path.relative_to(invoices_dir)
        print(f"  \u2192 \u0447\u0438\u0442\u0430\u044e {rel} ...", end=" ", flush=True)
        try:
            recs = parse_pdf_records(pdf_path)
            all_pdf_records.extend(recs)
            print(f"{len(recs)} \u0437\u0430\u043f\u0438\u0441\u0435\u0439")
        except Exception as e:
            print(f"\u26a0 \u043e\u0448\u0438\u0431\u043a\u0430: {e}")

    print(f"\n\u2713 \u0418\u0442\u043e\u0433\u043e \u0437\u0430\u043f\u0438\u0441\u0435\u0439 \u0438\u0437 PDF: {len(all_pdf_records)}")

    matches = match_records(xls_rows, all_pdf_records)
    print(f"\u2713 \u0421\u043e\u0432\u043f\u0430\u0434\u0435\u043d\u0438\u0439: {len(matches)} \u0441\u0442\u0440\u043e\u043a XLS")

    out_xls = output_dir / (xls_path.stem + "_matched.xlsx")
    annotate_xls(xls_path, matches, out_xls)
    highlight_pdfs(matches, output_dir)

    print("=" * 50)
    print("\u2713 \u0413\u043e\u0442\u043e\u0432\u043e!")
    print(f"  XLS \u0441 \u043e\u0442\u043c\u0435\u0442\u043a\u0430\u043c\u0438:  {out_xls}")
    print(f"  PDF \u0441 \u043f\u043e\u0434\u0441\u0432\u0435\u0442\u043a\u043e\u0439: {output_dir / 'highlighted_pdfs'}")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main_cli() -> None:
    import argparse
    parser = argparse.ArgumentParser(description="Invoice matcher: PDF vs XLS")
    parser.add_argument("--invoices", required=True, help="\u041f\u0430\u043f\u043a\u0430 \u0441 PDF-\u0441\u0447\u0435\u0442\u0430\u043c\u0438")
    parser.add_argument("--xls",      required=True, help="\u041f\u0443\u0442\u044c \u043a XLS/XLSX-\u0444\u0430\u0439\u043b\u0443")
    parser.add_argument("--out",      default="output", help="\u041f\u0430\u043f\u043a\u0430 \u0432\u044b\u0432\u043e\u0434\u0430")
    args = parser.parse_args()
    run_matching(
        invoices_dir=Path(args.invoices).resolve(),
        xls_path=Path(args.xls).resolve(),
        output_dir=Path(args.out).resolve(),
    )


if __name__ == "__main__":
    main_cli()
