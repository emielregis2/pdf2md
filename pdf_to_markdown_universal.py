#!/usr/bin/env python3
"""
pdf_to_markdown_universal.py
------------------------------
Uniwersalny konwerter PDF -> Markdown, zoptymalizowany pod kątem
wysyłania dokumentów do modeli AI (oszczędność tokenów).

Obsługuje w JEDNYM przebiegu:
  * PDF-y "cyfrowe" (z warstwą tekstową)      -> ekstrakcja tekstu + tabel
  * PDF-y będące skanami (same obrazy)        -> OCR (Tesseract)
  * PDF-y mieszane (część stron tekst, część skan) -> automatyczna detekcja per strona
  * Wierną rekonstrukcję TABEL (także ze skanów) jako prawdziwe tabele Markdown,
    a nie "zupę tekstową" wymieszanych liczb.

Optymalizacje pod tokeny AI:
  * usuwanie powtarzających się nagłówków/stopek (np. numeracja stron,
    "Poufne", nazwa firmy powtarzana na każdej stronie),
  * kompresja nadmiarowych białych znaków i pustych linii,
  * obrazy NIE są osadzane w treści (tylko krótki placeholder) - zdjęcia/loga
    nie niosą wartości analitycznej, a kosztują tokeny.

WYMAGANIA SYSTEMOWE (poza pip install):
    # Silnik OCR:
    sudo apt-get install tesseract-ocr tesseract-ocr-pol tesseract-ocr-eng
    # (dla innych języków: tesseract-ocr-<kod_iso>, np. tesseract-ocr-deu)

    # Renderowanie PDF -> obraz (poppler), zwykle już jest w systemie z PyMuPDF,
    # ale dla pdf2image jako fallback:
    sudo apt-get install poppler-utils

WYMAGANIA PYTHON (pip):
    pip install pymupdf4llm pymupdf img2table pytesseract pandas tabulate pillow \
        --break-system-packages

Użycie z linii poleceń:
    python pdf_to_markdown_universal.py raport.pdf
    python pdf_to_markdown_universal.py raport.pdf -o raport.md
    python pdf_to_markdown_universal.py *.pdf --output-dir ./markdown_output
    python pdf_to_markdown_universal.py skan.pdf --lang pol+eng --dpi 300
    python pdf_to_markdown_universal.py raport.pdf --force-ocr        # wymuś OCR na każdej stronie
    python pdf_to_markdown_universal.py raport.pdf --keep-headers     # nie usuwaj powtarzających się nagłówków/stopek
    python pdf_to_markdown_universal.py raport.pdf --stats            # pokaż statystyki oszczędności tokenów

Użycie jako moduł:
    from pdf_to_markdown_universal import convert_pdf
    convert_pdf("raport.pdf", "raport.md")
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

# --- Zależności zewnętrzne (z czytelnym komunikatem, gdy czegoś brakuje) ---
_MISSING = []
try:
    import fitz  # PyMuPDF
except ImportError:
    _MISSING.append("pymupdf")
try:
    import pymupdf4llm
except ImportError:
    _MISSING.append("pymupdf4llm")
try:
    import pandas as pd
except ImportError:
    _MISSING.append("pandas")
try:
    from img2table.document import Image as Img2TableImage
    from img2table.ocr import TesseractOCR
except ImportError:
    _MISSING.append("img2table")
try:
    import pytesseract
    from PIL import Image
except ImportError:
    _MISSING.append("pytesseract / pillow")

if _MISSING:
    print(
        "Brakuje wymaganych bibliotek: " + ", ".join(_MISSING) + "\n"
        "Zainstaluj je poleceniem:\n"
        "    pip install pymupdf4llm pymupdf img2table pytesseract pandas "
        "tabulate pillow --break-system-packages\n"
        "Wymagany jest też silnik systemowy Tesseract OCR:\n"
        "    sudo apt-get install tesseract-ocr tesseract-ocr-pol tesseract-ocr-eng",
        file=sys.stderr,
    )
    sys.exit(1)


def _ensure_tesseract_available() -> None:
    """
    Na Windows instalator Tesseract (UB Mannheim) domyślnie NIE dopisuje ścieżki
    do zmiennej PATH, przez co `pytesseract`/`img2table` nie znajdują silnika OCR
    mimo poprawnej instalacji. Sprawdzamy typowe lokalizacje i ustawiamy ścieżkę
    jawnie, zanim cokolwiek spróbuje wywołać tesseract.
    """
    import os
    import shutil

    if shutil.which("tesseract"):
        return  # jest w PATH, nic nie trzeba robić

    candidates = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"),
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            pytesseract.pytesseract.tesseract_cmd = candidate
            os.environ["TESSDATA_PREFIX"] = str(Path(candidate).parent / "tessdata")
            return

    print(
        "UWAGA: Nie znaleziono silnika Tesseract OCR ani w PATH, ani w typowych "
        "lokalizacjach instalacji na Windows. Konwersja zeskanowanych PDF-ów nie "
        "zadziała, dopóki nie zainstalujesz Tesseract OCR:\n"
        "  https://github.com/UB-Mannheim/tesseract/wiki\n"
        "Podczas instalacji zaznacz dodatkowe pakiety językowe (np. Polish).",
        file=sys.stderr,
    )


_ensure_tesseract_available()


# ============================================================================
# Konfiguracja / struktury danych
# ============================================================================

@dataclass
class ConversionConfig:
    lang: str = "pol+eng"          # języki OCR (kody Tesseract, łączone znakiem '+')
    dpi: int = 300                 # rozdzielczość renderowania stron-skanów
    force_ocr: bool = False        # wymuś OCR nawet gdy strona ma warstwę tekstową
    min_text_chars: int = 30       # próg: poniżej tylu znaków strona uznawana za skan
    strip_repeated_headers: bool = True   # usuwaj linie powtarzające się na wielu stronach
    header_footer_zone_ratio: float = 0.12  # % wysokości strony uznawany za strefę nagłówka/stopki
    repeated_line_threshold: float = 0.4  # linia uznana za "powtarzalną", gdy występuje na >=40% stron
    collapse_whitespace: bool = True
    table_min_confidence: bool = True  # img2table: implicit_rows/columns detection


@dataclass
class PageResult:
    page_number: int
    is_scanned: bool
    markdown: str
    header_footer_candidates: list[str] = field(default_factory=list)  # linie ze stref brzegowych strony


# ============================================================================
# Detekcja typu strony
# ============================================================================

def _page_has_text_layer(page: "fitz.Page", min_chars: int) -> bool:
    """Sprawdza, czy strona ma sensowną warstwę tekstową, czy to praktycznie sam obraz."""
    text = page.get_text("text").strip()
    return len(text) >= min_chars


def _header_footer_zone_candidates(page: "fitz.Page", zone_ratio: float) -> list[str]:
    """
    Zwraca linie tekstu leżące WYŁĄCZNIE w górnej lub dolnej strefie strony
    (np. górne/dolne 12% wysokości). Tylko takie linie mogą być kandydatami do
    usunięcia jako powtarzalny nagłówek/stopka - dzięki temu prawdziwa treść
    dokumentu (np. "Rozdział 1", "Rozdział 2"...) nigdy nie zostanie omyłkowo
    potraktowana jako nagłówek, nawet jeśli po normalizacji cyfr wygląda podobnie
    do innych stron.
    """
    page_h = page.rect.height
    top_limit = page_h * zone_ratio
    bottom_limit = page_h * (1 - zone_ratio)

    candidates = []
    text_dict = page.get_text("dict")
    for block in text_dict.get("blocks", []):
        for line in block.get("lines", []):
            y0 = line["bbox"][1]
            if y0 <= top_limit or y0 >= bottom_limit:
                line_text = "".join(span["text"] for span in line.get("spans", [])).strip()
                if line_text:
                    candidates.append(line_text)
    return candidates


def _header_footer_zone_candidates_ocr(
    elements: list[tuple[float, str]], page_height_px: float, zone_ratio: float
) -> list[str]:
    """Wariant powyższej funkcji dla stron OCR-owanych (mamy tylko listę (y, tekst))."""
    top_limit = page_height_px * zone_ratio
    bottom_limit = page_height_px * (1 - zone_ratio)
    return [text for y, text in elements if (y <= top_limit or y >= bottom_limit)]


def _optimal_render_dpi(
    page: "fitz.Page",
    requested_dpi: int,
    min_usable_dpi: int = 200,
    max_dpi: int = 450,
) -> int:
    """
    Wiele "skanów" w PDF-ie to w rzeczywistości pojedynczy obraz rozciągnięty na całą stronę.
    Renderowanie strony w DPI RÓŻNYM od natywnej rozdzielczości tego obrazu wymaga interpolacji
    (bilinear/bicubic), która rozmywa krawędzie linii tabel i pogarsza detekcję siatki oraz OCR.

    Dlatego zamiast sztywno narzucać `requested_dpi`, ta funkcja wykrywa taki przypadek i
    renderuje stronę DOKŁADNIE w natywnej rozdzielczości obrazu źródłowego (bez przeskalowania) -
    chyba że jest ona zbyt niska dla sensownego OCR (`min_usable_dpi`), wtedy dopiero skalujemy w górę.
    """
    try:
        images = page.get_images(full=True)
        if len(images) != 1:
            return requested_dpi

        xref = images[0][0]
        img_info = page.parent.extract_image(xref)
        img_w, img_h = img_info["width"], img_info["height"]

        page_w_pt, page_h_pt = page.rect.width, page.rect.height
        if page_w_pt <= 0 or page_h_pt <= 0:
            return requested_dpi

        native_dpi_x = img_w / (page_w_pt / 72)
        native_dpi_y = img_h / (page_h_pt / 72)
        native_dpi = round(min(native_dpi_x, native_dpi_y))  # ostrożnie: mniejsza z osi

        if native_dpi < min_usable_dpi:
            # obraz źródłowy ma zbyt niską rozdzielczość - trzeba go i tak przeskalować w górę,
            # więc korzystamy z wyższej z dwóch wartości (i tak stracimy na ostrości linii)
            return min(max(requested_dpi, min_usable_dpi), max_dpi)

        return min(native_dpi, max_dpi)
    except Exception:  # noqa: BLE001
        return requested_dpi


# ============================================================================
# Obsługa stron CYFROWYCH (mają warstwę tekstu)
# ============================================================================

def _convert_digital_page(doc: "fitz.Document", page_index: int) -> str:
    """
    Konwertuje pojedynczą stronę cyfrową na Markdown przy pomocy pymupdf4llm,
    który wykorzystuje natywny, wektorowy detektor tabel PyMuPDF (page.find_tables()) -
    dokładny dla PDF-ów generowanych programowo (raporty, wyciągi, faktury).
    """
    md = pymupdf4llm.to_markdown(doc, pages=[page_index])
    return md.strip()


# ============================================================================
# Obsługa stron ZESKANOWANYCH (OCR + detekcja tabel z obrazu)
# ============================================================================

def _dataframe_to_markdown_table(df: "pd.DataFrame") -> str:
    """Zamienia DataFrame (wynik img2table) na tabelę Markdown."""
    df = df.fillna("")

    # Jeśli img2table nie ustawiło nagłówków (kolumny to domyślny RangeIndex 0,1,2...),
    # a pierwszy wiersz wygląda jak nagłówek (same niepuste, krótkie wartości tekstowe),
    # użyj go jako właściwego nagłówka tabeli - czytelniejsze i krótsze niż "0 | 1 | 2".
    if list(df.columns) == list(range(len(df.columns))) and len(df) > 0:
        first_row = df.iloc[0]
        if all(str(v).strip() != "" for v in first_row):
            df = df.iloc[1:].copy()
            df.columns = [str(v).strip() for v in first_row]

    df = df.astype(str).replace(r"^\s*$", "", regex=True)
    try:
        return df.to_markdown(index=False)
    except ImportError:
        # fallback bez biblioteki tabulate (nie powinno się zdarzyć, ale na wszelki wypadek)
        header = "| " + " | ".join(df.columns.astype(str)) + " |"
        sep = "| " + " | ".join(["---"] * len(df.columns)) + " |"
        rows = [
            "| " + " | ".join(row.astype(str)) + " |"
            for _, row in df.iterrows()
        ]
        return "\n".join([header, sep] + rows)


def _convert_scanned_page(
    page: "fitz.Page",
    page_index: int,
    config: ConversionConfig,
    ocr_engine: "TesseractOCR",
    tmp_dir: Path,
) -> tuple[str, list[str]]:
    """
    Konwertuje zeskanowaną stronę:
      1. renderuje stronę do obrazu w wysokiej rozdzielczości,
      2. wykrywa i ekstrahuje tabele (img2table) -> Markdown,
      3. maskuje obszary tabel na obrazie,
      4. resztę strony poddaje zwykłemu OCR (pytesseract),
      5. składa wynik w kolejności odpowiadającej układowi strony (góra -> dół).

    Zwraca (markdown_strony, lista_kandydatów_na_nagłówek_stopkę).
    """
    effective_dpi = _optimal_render_dpi(page, config.dpi)
    zoom = effective_dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    img_path = tmp_dir / f"page_{page_index}.png"
    pix.save(str(img_path))

    elements: list[tuple[float, str]] = []  # (pozycja_y, markdown_fragment)

    # --- 1) Detekcja tabel ---
    table_bboxes = []
    try:
        img2table_doc = Img2TableImage(src=str(img_path))
        extracted_tables = img2table_doc.extract_tables(
            ocr=ocr_engine,
            implicit_rows=True,
            implicit_columns=True,
            borderless_tables=True,
            min_confidence=50,
        )
        for table in extracted_tables:
            df = table.df
            if df is None or df.empty:
                continue
            md_table = _dataframe_to_markdown_table(df)
            top_y = table.bbox.y1
            elements.append((top_y, md_table))
            table_bboxes.append(table.bbox)
    except Exception as exc:  # noqa: BLE001
        print(f"  [ostrzeżenie] detekcja tabel nieudana na str. {page_index + 1}: {exc}", file=sys.stderr)

    # --- 2) Maskowanie obszarów tabel przed zwykłym OCR (unikamy duplikacji treści) ---
    pil_img = Image.open(img_path).convert("RGB")
    if table_bboxes:
        from PIL import ImageDraw
        draw = ImageDraw.Draw(pil_img)
        for bbox in table_bboxes:
            draw.rectangle([bbox.x1, bbox.y1, bbox.x2, bbox.y2], fill="white")

    # --- 3) OCR pozostałego tekstu, z pozycjami (do zachowania kolejności) ---
    ocr_data = pytesseract.image_to_data(
        pil_img, lang=config.lang, output_type=pytesseract.Output.DICT
    )
    # Grupujemy słowa w linie na podstawie (block_num, par_num, line_num)
    lines: dict[tuple, dict] = {}
    n = len(ocr_data["text"])
    for i in range(n):
        word = ocr_data["text"][i].strip()
        if not word:
            continue
        key = (ocr_data["block_num"][i], ocr_data["par_num"][i], ocr_data["line_num"][i])
        if key not in lines:
            lines[key] = {"top": ocr_data["top"][i], "words": []}
        lines[key]["words"].append(word)

    for key, data in lines.items():
        line_text = " ".join(data["words"])
        elements.append((float(data["top"]), line_text))

    # --- 4) Sortowanie wg pozycji pionowej i składanie wyniku ---
    elements.sort(key=lambda e: e[0])
    parts = [e[1] for e in elements]

    # --- 5) Kandydaci na nagłówek/stopkę - linie leżące w strefie brzegowej strony ---
    header_footer_candidates = _header_footer_zone_candidates_ocr(
        elements, pix.height, config.header_footer_zone_ratio
    )

    return "\n\n".join(parts).strip(), header_footer_candidates


# ============================================================================
# Post-processing: oszczędność tokenów
# ============================================================================

def _collapse_whitespace(text: str) -> str:
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = "\n".join(line.rstrip() for line in text.split("\n"))
    return text.strip()


def _strip_repeated_headers_footers(pages: list[PageResult], config: ConversionConfig) -> list[PageResult]:
    """
    Wykrywa linie, które powtarzają się (niemal) identycznie na dużej części stron
    (typowo: nazwa firmy, "Poufne", numeracja "Strona X z Y") i usuwa je z treści -
    nie wnoszą informacji, a kosztują tokeny przy każdej stronie.

    WAŻNE: kandydatami są WYŁĄCZNIE linie leżące w górnej/dolnej strefie strony
    (`header_footer_candidates`, patrz `_header_footer_zone_candidates*`). Dzięki temu
    prawdziwa treść dokumentu (np. "Rozdział 1", "Rozdział 2"...) nigdy nie zostanie
    omyłkowo usunięta, nawet jeśli po normalizacji cyfr wygląda podobnie na każdej stronie.
    """
    if len(pages) < 3:
        return pages  # za mało stron, by wiarygodnie wykryć wzorzec

    line_counter: Counter[str] = Counter()
    for p in pages:
        unique_candidates = set(l.strip() for l in p.header_footer_candidates if l.strip())
        for line in unique_candidates:
            # normalizujemy numery stron np. "Strona 3 z 12" -> "Strona # z #", żeby też je złapać
            normalized = re.sub(r"\d+", "#", line)
            line_counter[normalized] += 1

    threshold = max(2, int(len(pages) * config.repeated_line_threshold))
    repeated_patterns = {line for line, count in line_counter.items() if count >= threshold}

    if not repeated_patterns:
        return pages

    cleaned_pages = []
    for p in pages:
        new_lines = []
        for line in p.markdown.split("\n"):
            normalized = re.sub(r"\d+", "#", line.strip())
            if normalized in repeated_patterns and normalized != "":
                continue  # pomijamy powtarzalną linię (nagłówek/stopka)
            new_lines.append(line)
        cleaned_pages.append(
            PageResult(p.page_number, p.is_scanned, "\n".join(new_lines), p.header_footer_candidates)
        )
    return cleaned_pages


# ============================================================================
# Główna funkcja konwersji
# ============================================================================

def convert_pdf(
    pdf_path: str | Path,
    output_path: str | Path | None = None,
    config: ConversionConfig | None = None,
    show_stats: bool = False,
) -> Path:
    """
    Konwertuje plik PDF (dowolny: cyfrowy / skan / mieszany) na plik Markdown.

    Zwraca ścieżkę do zapisanego pliku .md
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"Nie znaleziono pliku: {pdf_path}")

    config = config or ConversionConfig()

    if output_path is None:
        output_path = pdf_path.with_suffix(".md")
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    tmp_dir = output_path.parent / f".{output_path.stem}_tmp"
    tmp_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(str(pdf_path))
    ocr_engine = TesseractOCR(lang=config.lang)

    page_results: list[PageResult] = []
    n_scanned = 0

    print(f"Przetwarzam: {pdf_path.name} ({len(doc)} str.)")
    for i, page in enumerate(doc):
        has_text = _page_has_text_layer(page, config.min_text_chars)
        is_scanned = config.force_ocr or not has_text

        if is_scanned:
            n_scanned += 1
            md, hf_candidates = _convert_scanned_page(page, i, config, ocr_engine, tmp_dir)
            print(f"  str. {i + 1}/{len(doc)}: SKAN -> OCR")
        else:
            md = _convert_digital_page(doc, i)
            hf_candidates = _header_footer_zone_candidates(page, config.header_footer_zone_ratio)
            print(f"  str. {i + 1}/{len(doc)}: tekst cyfrowy")

        # Usuwamy odwołania do osadzonych obrazów - nie chcemy ich w markdownie
        md = re.sub(r"!\[.*?\]\(.*?\)", "[obraz]", md)

        page_results.append(PageResult(i + 1, is_scanned, md, hf_candidates))

    doc.close()

    # Sprzątanie plików tymczasowych
    for f in tmp_dir.glob("*"):
        f.unlink()
    tmp_dir.rmdir()

    # Usuwanie powtarzalnych nagłówków/stopek
    if config.strip_repeated_headers:
        page_results = _strip_repeated_headers_footers(page_results, config)

    # Składanie finalnego dokumentu
    final_parts = []
    for p in page_results:
        content = p.markdown.strip()
        if content:
            final_parts.append(content)

    final_md = "\n\n".join(final_parts)

    if config.collapse_whitespace:
        final_md = _collapse_whitespace(final_md)

    output_path.write_text(final_md, encoding="utf-8")

    print(f"Zapisano: {output_path}")
    if show_stats:
        _print_stats(pdf_path, final_md, n_scanned, len(page_results))

    return output_path


def _print_stats(pdf_path: Path, final_md: str, n_scanned: int, n_pages: int) -> None:
    pdf_size_kb = pdf_path.stat().st_size / 1024
    n_chars = len(final_md)
    # bardzo zgrubne oszacowanie tokenów (dla modeli GPT/Claude ~4 znaki/token dla języka angielskiego,
    # dla polskiego zwykle nieco mniej znaków/token ze względu na diakrytyki - przyjmujemy ~3.3)
    est_tokens = int(n_chars / 3.3)
    print("\n--- Statystyki ---")
    print(f"Stron ogółem:        {n_pages}")
    print(f"Stron przez OCR:     {n_scanned}")
    print(f"Rozmiar PDF:         {pdf_size_kb:.1f} KB")
    print(f"Znaków w Markdown:   {n_chars:,}".replace(",", " "))
    print(f"Szac. liczba tokenów: ~{est_tokens:,}".replace(",", " "))


def convert_many(pdf_paths: list[Path], output_dir: Path | None, config: ConversionConfig) -> None:
    ok, failed = [], []
    for pdf_path in pdf_paths:
        try:
            out_path = (output_dir / (pdf_path.stem + ".md")) if output_dir else None
            convert_pdf(pdf_path, out_path, config)
            ok.append(pdf_path)
        except Exception as exc:  # noqa: BLE001
            print(f"BŁĄD przy pliku {pdf_path.name}: {exc}", file=sys.stderr)
            failed.append(pdf_path)

    print("\n--- Podsumowanie ---")
    print(f"Poprawnie skonwertowane: {len(ok)}")
    if failed:
        print(f"Nieudane: {len(failed)}")
        for f in failed:
            print(f"  - {f}")


# ============================================================================
# CLI
# ============================================================================

def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Uniwersalny konwerter PDF -> Markdown (tekst cyfrowy + OCR skanów + "
            "wierne tabele), zoptymalizowany pod tokeny AI."
        ),
    )
    parser.add_argument("pdf_files", nargs="+", help="Plik(i) PDF (obsługiwane wildcardy, np. *.pdf)")
    parser.add_argument("-o", "--output", help="Ścieżka pliku wyjściowego .md (tylko dla 1 pliku)")
    parser.add_argument("--output-dir", help="Katalog wyjściowy dla wielu plików")
    parser.add_argument("--lang", default="pol+eng", help="Języki OCR wg kodów Tesseract (domyślnie: pol+eng)")
    parser.add_argument("--dpi", type=int, default=300, help="Rozdzielczość renderowania skanów (domyślnie: 300)")
    parser.add_argument("--force-ocr", action="store_true", help="Wymuś OCR na WSZYSTKICH stronach")
    parser.add_argument("--keep-headers", action="store_true",
                         help="Nie usuwaj powtarzających się nagłówków/stopek")
    parser.add_argument("--stats", action="store_true", help="Pokaż statystyki (rozmiar, szac. liczba tokenów)")
    return parser


def main() -> None:
    parser = build_arg_parser()
    args = parser.parse_args()

    pdf_paths = [Path(p) for p in args.pdf_files]
    missing = [p for p in pdf_paths if not p.exists()]
    if missing:
        for m in missing:
            print(f"Nie znaleziono pliku: {m}", file=sys.stderr)
        sys.exit(1)

    config = ConversionConfig(
        lang=args.lang,
        dpi=args.dpi,
        force_ocr=args.force_ocr,
        strip_repeated_headers=not args.keep_headers,
    )

    if len(pdf_paths) == 1 and not args.output_dir:
        convert_pdf(pdf_paths[0], args.output, config, show_stats=args.stats)
    else:
        output_dir = Path(args.output_dir) if args.output_dir else None
        if output_dir:
            output_dir.mkdir(parents=True, exist_ok=True)
        convert_many(pdf_paths, output_dir, config)


if __name__ == "__main__":
    main()
