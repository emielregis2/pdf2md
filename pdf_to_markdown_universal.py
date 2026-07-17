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
import datetime
import os
import re
import sys
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

# --- Kodowanie stdout/stderr -----------------------------------------------
# Pod pythonw.exe (uruchamianie BEZ konsoli - dokładnie tak działa wywołanie
# z menu kontekstowego Eksploratora) sys.stdout/sys.stderr bywają None, a
# zwykłe print() na None się wywala. Tam, gdzie konsola jednak istnieje
# (uruchomienie z cmd/PowerShell), domyślne kodowanie konsoli Windows
# (cp852/cp1250) potrafi zamieniać polskie znaki (ą, ć, ę, ł, ń, ó, ś, ż, ź)
# w "krzaki". Ustawiamy to najwcześniej, jak się da - zanim padnie pierwszy
# print() gdziekolwiek w tym module.
for _stream_name in ("stdout", "stderr"):
    _stream = getattr(sys, _stream_name)
    if _stream is None:
        setattr(sys, _stream_name, open(os.devnull, "w", encoding="utf-8"))
    elif hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:  # noqa: BLE001
            pass

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
    Konfiguruje ścieżkę do silnika Tesseract OCR ORAZ, niezależnie od niej,
    ścieżkę do katalogu z danymi językowymi (tessdata). To dwa osobne, częste
    źródła problemów na Windows:
      1. tesseract.exe bywa zainstalowany, ale nie dopisany do PATH (domyślne
         zachowanie instalatora UB Mannheim),
      2. nawet gdy tesseract.exe JEST w PATH, zainstalowany pakiet językowy
         może nie obejmować potrzebnego języka (np. nie zaznaczono "Polish"
         przy instalacji) - wtedy img2table/pytesseract zgłaszają błąd
         "trained data cannot be located", mimo że sam silnik działa poprawnie.
    Dlatego szukamy tessdata ZAWSZE, również gdy tesseract.exe jest już w PATH -
    z priorytetem dla katalogu "tessdata" leżącego obok tego skryptu (jeśli
    ktoś ręcznie dołożył tam brakujące pakiety językowe).
    """
    import os
    import shutil

    # --- 1) silnik tesseract.exe ---
    if not shutil.which("tesseract"):
        candidates = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"),
        ]
        for candidate in candidates:
            if candidate and Path(candidate).exists():
                pytesseract.pytesseract.tesseract_cmd = candidate
                break
        else:
            print(
                "UWAGA: Nie znaleziono silnika Tesseract OCR ani w PATH, ani w typowych "
                "lokalizacjach instalacji na Windows. Konwersja zeskanowanych PDF-ów nie "
                "zadziała, dopóki nie zainstalujesz Tesseract OCR:\n"
                "  https://github.com/UB-Mannheim/tesseract/wiki\n"
                "Podczas instalacji zaznacz dodatkowe pakiety językowe (np. Polish).",
                file=sys.stderr,
            )

    # --- 2) katalog z danymi językowymi (tessdata) ---
    # UWAGA: nie ufamy ślepo już ustawionej zmiennej TESSDATA_PREFIX (np. przez
    # system, wcześniejszy instalator albo ręczną, błędną konfigurację) - jeśli
    # wskazuje na folder, który nie istnieje albo nie ma w nim plików .traineddata,
    # traktujemy to tak, jakby zmienna nie była ustawiona wcale, i szukamy dalej.
    def _tessdata_dir_is_valid(path_str: str | None) -> bool:
        if not path_str:
            return False
        try:
            p = Path(path_str)
            return p.is_dir() and any(p.glob("*.traineddata"))
        except OSError:
            return False

    if not _tessdata_dir_is_valid(os.environ.get("TESSDATA_PREFIX")):
        local_tessdata = Path(__file__).resolve().parent / "tessdata"
        tessdata_candidates = [
            local_tessdata,  # priorytet: ręcznie dołożony pakiet obok skryptu
            Path(r"C:\Program Files\Tesseract-OCR\tessdata"),
            Path(r"C:\Program Files (x86)\Tesseract-OCR\tessdata"),
        ]
        for td in tessdata_candidates:
            if _tessdata_dir_is_valid(str(td)):
                os.environ["TESSDATA_PREFIX"] = str(td)
                break
        else:
            print(
                "UWAGA: Nie znaleziono katalogu tessdata (danych językowych OCR) - "
                "ani obok tego skryptu, ani w typowej lokalizacji instalacji Tesseract. "
                "Rozpoznawanie tekstu ze skanów może się nie powieść z błędem "
                "'trained data cannot be located'. Umieść potrzebne pliki .traineddata "
                f"(np. pol.traineddata, eng.traineddata) w katalogu: {local_tessdata}",
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
    xlsx_max_rows_per_sheet: int = 300     # twardy limit wierszy na arkusz (ochrona tokenów)
    xlsx_skip_hidden_sheets: bool = True   # pomijaj arkuszy ukryte (zwykle robocze/pomocnicze)


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

    # Sprzątanie plików tymczasowych. Może się nie udać, jeśli folder leży w
    # OneDrive (lub podobnej usłudze synchronizacji), która chwilowo blokuje
    # pliki podczas indeksowania - to NIE powinno przerywać konwersji, skoro
    # dokument został już poprawnie przetworzony, więc łapiemy błąd i tylko
    # ostrzegamy, zamiast wywalać cały proces.
    try:
        for f in tmp_dir.glob("*"):
            f.unlink()
        tmp_dir.rmdir()
    except OSError as exc:
        print(
            f"  [ostrzeżenie] nie udało się usunąć folderu tymczasowego {tmp_dir}: {exc}\n"
            "  (prawdopodobnie folder jest chwilowo zablokowany przez OneDrive - "
            "można go usunąć ręcznie później, konwersja i tak się dokończy)",
            file=sys.stderr,
        )

    # Zrzut treści PRZED czyszczeniem (nagłówki/stopki, białe znaki) - potrzebny
    # tylko do statystyk, żeby pokazać, ile dała sama optymalizacja.
    raw_parts = [p.markdown.strip() for p in page_results if p.markdown.strip()]
    raw_md = "\n\n".join(raw_parts)

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
        _print_stats(pdf_path, raw_md, final_md, n_scanned, len(page_results))

    return output_path


def _estimate_tokens(text: str) -> int:
    """Bardzo zgrubne oszacowanie liczby tokenów (dla PL ~3.3 znaku/token)."""
    return int(len(text) / 3.3)


def _fmt_int(n: int) -> str:
    return f"{n:,}".replace(",", " ")


def _print_stats(pdf_path: Path, raw_md: str, final_md: str, n_scanned: int, n_pages: int) -> None:
    pdf_size_kb = pdf_path.stat().st_size / 1024

    raw_tokens = _estimate_tokens(raw_md)
    final_chars = len(final_md)
    final_tokens = _estimate_tokens(final_md)

    # Wg oficjalnej dokumentacji Claude API (platform.claude.com/docs/en/build-with-claude/pdf-support):
    # gdy PDF trafia do modelu BEZPOŚREDNIO jako dokument (a nie jako Markdown),
    # każda strona jest liczona jako 1500-3000 tokenów (tekst + reprezentacja
    # obrazu strony), niezależnie od tego, jak mało faktycznego tekstu zawiera.
    pdf_tokens_low = n_pages * 1500
    pdf_tokens_high = n_pages * 3000

    print("\n--- Statystyki ---")
    print(f"Stron ogółem:            {n_pages}")
    print(f"Stron przez OCR:         {n_scanned}")
    print(f"Rozmiar PDF:             {pdf_size_kb:.1f} KB")
    print()
    print("Wynikowy plik Markdown:")
    print(f"  Znaków:                {_fmt_int(final_chars)}")
    print(f"  Szac. liczba tokenów:  ~{_fmt_int(final_tokens)}")
    print()
    print("Dla porównania - ten sam PDF wysłany bezpośrednio do Claude*:")
    print(f"  Szac. liczba tokenów:  ~{_fmt_int(pdf_tokens_low)} - {_fmt_int(pdf_tokens_high)}")

    if pdf_tokens_low > final_tokens:
        saved_low = pdf_tokens_low - final_tokens
        saved_high = pdf_tokens_high - final_tokens
        pct_low = 100 * saved_low / pdf_tokens_low if pdf_tokens_low else 0
        pct_high = 100 * saved_high / pdf_tokens_high if pdf_tokens_high else 0
        print()
        print("Oszacowana oszczędność dzięki konwersji na Markdown:")
        print(f"  ~{_fmt_int(saved_low)} - {_fmt_int(saved_high)} tokenów  "
              f"(ok. {pct_low:.0f}% - {pct_high:.0f}% mniej)")

    if raw_tokens > final_tokens:
        cleanup_saved = raw_tokens - final_tokens
        print(f"\n(Samo czyszczenie nagłówków/stopek i białych znaków w tym pliku")
        print(f" zaoszczędziło dodatkowo ~{_fmt_int(cleanup_saved)} tokenów)")

    print("\n* wg dokumentacji Anthropic: platform.claude.com/docs/en/build-with-claude/pdf-support")
    print("  (każda strona PDF-a liczona jako 1500-3000 tokenów: tekst + obraz strony)")


# ============================================================================
# Obsługa dokumentów WORD (.docx oraz stary format .doc)
# ============================================================================

def _convert_docx_to_markdown(docx_path: Path) -> str:
    """
    Konwertuje plik .docx na Markdown przez mammoth (docx -> HTML, z zachowaniem
    nagłówków, list, pogrubień, tabel na podstawie stylów Worda) + markdownify
    (HTML -> Markdown). Obie biblioteki są czysto pythonowe (bez zewnętrznych
    programów), instalowane osobno od reszty, bo nie każdy użytkownik pdf2md
    potrzebuje obsługi Worda.
    """
    try:
        import mammoth
    except ImportError as exc:
        raise RuntimeError(
            "Do konwersji plików Word (.docx) potrzebna jest biblioteka 'mammoth'.\n"
            "Zainstaluj ją poleceniem:\n"
            "    pip install mammoth markdownify"
        ) from exc
    try:
        from markdownify import markdownify as _html_to_md
    except ImportError as exc:
        raise RuntimeError(
            "Do konwersji plików Word (.docx) potrzebna jest biblioteka 'markdownify'.\n"
            "Zainstaluj ją poleceniem:\n"
            "    pip install mammoth markdownify"
        ) from exc

    with open(docx_path, "rb") as f:
        # Domyślna mapa stylów mammoth rozpoznaje wbudowane style Worda po ich
        # wewnętrznym ID (zwykle angielskojęzycznym), ale część dokumentów -
        # zwłaszcza pisanych w polskiej wersji Worda albo wyeksportowanych z
        # innych narzędzi - ma nagłówki nazwane po polsku ("Nagłówek 1" itd.).
        # Dodajemy jawną mapę dla najczęstszych wariantów, żeby takie nagłówki
        # też trafiały do Markdown jako #, ## itd., a nie jako zwykły tekst.
        style_map = """
p[style-name='Heading 1'] => h1:fresh
p[style-name='Heading 2'] => h2:fresh
p[style-name='Heading 3'] => h3:fresh
p[style-name='Heading 4'] => h4:fresh
p[style-name='Title'] => h1:fresh
p[style-name='Nagłówek 1'] => h1:fresh
p[style-name='Nagłówek 2'] => h2:fresh
p[style-name='Nagłówek 3'] => h3:fresh
p[style-name='Nagłówek 4'] => h4:fresh
p[style-name='Tytuł'] => h1:fresh
"""
        result = mammoth.convert_to_html(f, style_map=style_map)

    for msg in result.messages:
        # mammoth zgłasza np. niestandardowe style, które pominął - to nie błąd,
        # ale warto o tym wiedzieć, więc logujemy na stderr, nie przerywając.
        print(f"  [mammoth] {msg}", file=sys.stderr)

    md = _html_to_md(result.value, heading_style="ATX", bullets="-")
    return md.strip()


def _convert_legacy_doc_to_docx(doc_path: Path, tmp_dir: Path) -> Path:
    """
    Stary format .doc to binarny format OLE (nie ZIP+XML jak .docx), więc mammoth
    go nie obsłuży. Najwierniejszy sposób konwersji to automatyzacja Microsoft Word
    przez pywin32 (jeśli Word jest zainstalowany na tym komputerze) - otwieramy
    dokument i zapisujemy go od razu jako .docx, a dalej lecimy tym samym potokiem
    co dla zwykłych plików .docx.
    """
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "Stare pliki .doc wymagają automatycznej konwersji przez Microsoft Word, "
            "a biblioteka 'pywin32' nie jest zainstalowana.\n"
            "Zainstaluj ją poleceniem: pip install pywin32\n"
            "Alternatywnie: otwórz ten plik ręcznie w Wordzie, zapisz jako .docx "
            "('Zapisz jako' -> typ pliku 'Dokument Word (*.docx)') i przekonwertuj "
            "ten nowy plik."
        ) from exc

    out_docx = tmp_dir / (doc_path.stem + "_converted.docx")
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        wd_doc = word.Documents.Open(str(doc_path), ReadOnly=True)
        wd_doc.SaveAs(str(out_docx), FileFormat=16)  # 16 = wdFormatXMLDocument (.docx)
        wd_doc.Close(False)
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(
            "Nie udało się automatycznie przekonwertować pliku .doc przez Microsoft Word "
            f"(błąd: {exc}).\n"
            "Upewnij się, że Word jest zainstalowany na tym komputerze, albo otwórz "
            "plik ręcznie i zapisz jako .docx."
        ) from exc
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:  # noqa: BLE001
                pass

    if not out_docx.exists():
        raise RuntimeError("Konwersja .doc -> .docx przez Word nie powiodła się (brak pliku wynikowego).")
    return out_docx


def convert_docx(
    doc_path: str | Path,
    output_path: str | Path | None = None,
    config: ConversionConfig | None = None,
    show_stats: bool = False,
) -> Path:
    """
    Konwertuje plik Word (.docx lub stary .doc) na plik Markdown.
    Zwraca ścieżkę do zapisanego pliku .md
    """
    doc_path = Path(doc_path)
    if not doc_path.exists():
        raise FileNotFoundError(f"Nie znaleziono pliku: {doc_path}")

    config = config or ConversionConfig()

    if output_path is None:
        output_path = doc_path.with_suffix(".md")
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    tmp_dir = output_path.parent / f".{output_path.stem}_tmp"
    tmp_dir.mkdir(parents=True, exist_ok=True)

    try:
        if doc_path.suffix.lower() == ".doc":
            print(f"Przetwarzam: {doc_path.name} (stary format .doc -> konwertuję przez Word)")
            actual_docx_path = _convert_legacy_doc_to_docx(doc_path, tmp_dir)
        else:
            print(f"Przetwarzam: {doc_path.name}")
            actual_docx_path = doc_path

        raw_md = _convert_docx_to_markdown(actual_docx_path)
    finally:
        # Sprzątanie - tak samo jak przy PDF, nie przerywamy konwersji, jeśli
        # folder tymczasowy akurat jest zablokowany (np. przez OneDrive).
        try:
            for f in tmp_dir.glob("*"):
                f.unlink()
            tmp_dir.rmdir()
        except OSError as exc:
            print(
                f"  [ostrzeżenie] nie udało się usunąć folderu tymczasowego {tmp_dir}: {exc}",
                file=sys.stderr,
            )

    # Usuwamy odwołania do osadzonych obrazów - tak samo jak przy PDF
    final_md = re.sub(r"!\[.*?\]\(.*?\)", "[obraz]", raw_md)

    if config.collapse_whitespace:
        final_md = _collapse_whitespace(final_md)

    output_path.write_text(final_md, encoding="utf-8")
    print(f"Zapisano: {output_path}")

    if show_stats:
        _print_stats_docx(doc_path, raw_md, final_md)

    return output_path


def _print_stats_docx(doc_path: Path, raw_md: str, final_md: str) -> None:
    file_size_kb = doc_path.stat().st_size / 1024
    raw_tokens = _estimate_tokens(raw_md)
    final_chars = len(final_md)
    final_tokens = _estimate_tokens(final_md)

    print("\n--- Statystyki ---")
    print(f"Rozmiar pliku:           {file_size_kb:.1f} KB")
    print()
    print("Wynikowy plik Markdown:")
    print(f"  Znaków:                {_fmt_int(final_chars)}")
    print(f"  Szac. liczba tokenów:  ~{_fmt_int(final_tokens)}")

    if raw_tokens > final_tokens:
        cleanup_saved = raw_tokens - final_tokens
        print(f"\n(Czyszczenie białych znaków w tym pliku zaoszczędziło dodatkowo "
              f"~{_fmt_int(cleanup_saved)} tokenów)")

    print(
        "\n* Uwaga: w odróżnieniu od PDF, pliki Word NIE są w Claude renderowane "
        "jako obrazy stron - ich koszt tokenowy jest zbliżony do zwykłego tekstu. "
        "Konwersja na Markdown daje tu przede wszystkim porządek i przewidywalny "
        "format, a nie tak dużą oszczędność tokenów jak przy PDF."
    )


# ============================================================================
# Obsługa arkuszy EXCEL (.xlsx oraz stary format .xls)
# ============================================================================

def _trim_empty_edges(rows: list[tuple]) -> list[tuple]:
    """
    Usuwa całkowicie puste wiersze i kolumny na brzegach danych. Excel często
    "pamięta" znacznie większy zakres komórek niż faktyczne dane (np. po
    wcześniejszym formatowaniu usuniętych już wierszy) - to jest pierwszy i
    najprostszy sposób na uniknięcie wysyłania samych pustych komórek do AI.
    """
    def is_row_empty(row: tuple) -> bool:
        return all(c is None or (isinstance(c, str) and c.strip() == "") for c in row)

    start = 0
    while start < len(rows) and is_row_empty(rows[start]):
        start += 1
    end = len(rows)
    while end > start and is_row_empty(rows[end - 1]):
        end -= 1
    rows = rows[start:end]
    if not rows:
        return []

    n_cols = max(len(r) for r in rows)
    rows = [tuple(r) + (None,) * (n_cols - len(r)) for r in rows]

    def is_col_empty(col_idx: int) -> bool:
        return all(
            r[col_idx] is None or (isinstance(r[col_idx], str) and r[col_idx].strip() == "")
            for r in rows
        )

    col_start = 0
    while col_start < n_cols and is_col_empty(col_start):
        col_start += 1
    col_end = n_cols
    while col_end > col_start and is_col_empty(col_end - 1):
        col_end -= 1

    return [r[col_start:col_end] for r in rows]


def _cell_to_str(value) -> str:
    """Formatuje pojedynczą komórkę do tekstu w komórce tabeli Markdown."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TAK" if value else "NIE"
    if isinstance(value, float):
        # unikamy "3.0000000000000004" itp. - jeśli liczba całkowita, pokaż bez ".0"
        text = str(int(value)) if value.is_integer() else f"{value:g}"
    elif isinstance(value, datetime.datetime):
        text = value.strftime("%Y-%m-%d") if value.time() == datetime.time(0, 0) else value.strftime("%Y-%m-%d %H:%M")
    elif isinstance(value, datetime.date):
        text = value.strftime("%Y-%m-%d")
    else:
        text = str(value)
    # w komórce tabeli Markdown nie mogą być "|" ani znaki nowej linii
    return text.replace("|", "\\|").replace("\n", " ").strip()


def _rows_to_markdown_table(rows: list[tuple]) -> str:
    """Zamienia listę krotek (pierwszy wiersz = nagłówek) na tabelę Markdown."""
    if not rows:
        return ""
    header, body = rows[0], rows[1:]

    def fmt_row(row: tuple) -> str:
        return "| " + " | ".join(_cell_to_str(v) for v in row) + " |"

    lines = [fmt_row(header), "| " + " | ".join(["---"] * len(header)) + " |"]
    lines.extend(fmt_row(r) for r in body)
    return "\n".join(lines)


def _convert_xlsx_to_markdown(xlsx_path: Path, config: ConversionConfig) -> tuple[str, list[dict]]:
    """
    Konwertuje skoroszyt .xlsx na Markdown, arkusz po arkuszu. Zwraca
    (markdown, lista_raportów_per_arkusz) - ta druga wartość jest potrzebna
    tylko do statystyk (--stats), żeby było widać co i dlaczego zostało
    pominięte albo obcięte.
    """
    try:
        import openpyxl
    except ImportError as exc:
        raise RuntimeError(
            "Do konwersji arkuszy Excel (.xlsx) potrzebna jest biblioteka 'openpyxl'.\n"
            "Zainstaluj ją poleceniem:\n"
            "    pip install openpyxl"
        ) from exc

    wb = openpyxl.load_workbook(str(xlsx_path), data_only=True, read_only=True)

    parts: list[str] = []
    sheet_reports: list[dict] = []

    for ws in wb.worksheets:
        if config.xlsx_skip_hidden_sheets and ws.sheet_state != "visible":
            sheet_reports.append({"name": ws.title, "skipped": "ukryty arkusz"})
            continue

        rows = _trim_empty_edges(list(ws.iter_rows(values_only=True)))
        if not rows:
            sheet_reports.append({"name": ws.title, "skipped": "pusty arkusz"})
            continue

        n_rows_total = len(rows)
        n_cols = len(rows[0])
        truncated = n_rows_total > config.xlsx_max_rows_per_sheet
        if truncated:
            rows = rows[: config.xlsx_max_rows_per_sheet]

        section = f"## {ws.title}\n\n"
        if truncated:
            section += (
                f"*(pokazano pierwsze {config.xlsx_max_rows_per_sheet} z {n_rows_total} "
                f"wierszy - pełne dane w oryginalnym pliku)*\n\n"
            )
        section += _rows_to_markdown_table(rows)
        parts.append(section)

        sheet_reports.append(
            {"name": ws.title, "rows": n_rows_total, "cols": n_cols, "truncated": truncated}
        )

    wb.close()
    return "\n\n".join(parts).strip(), sheet_reports


def _convert_legacy_xls_to_xlsx(xls_path: Path, tmp_dir: Path) -> Path:
    """
    Stary format .xls (binarny) konwertujemy do .xlsx przez automatyzację
    Microsoft Excel (pywin32) - analogicznie do obsługi starych plików .doc.
    """
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "Stare pliki .xls wymagają automatycznej konwersji przez Microsoft Excel, "
            "a biblioteka 'pywin32' nie jest zainstalowana.\n"
            "Zainstaluj ją poleceniem: pip install pywin32\n"
            "Alternatywnie: otwórz ten plik ręcznie w Excelu i zapisz jako .xlsx."
        ) from exc

    out_xlsx = tmp_dir / (xls_path.stem + "_converted.xlsx")
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(xls_path), ReadOnly=True)
        wb.SaveAs(str(out_xlsx), FileFormat=51)  # 51 = xlOpenXMLWorkbook (.xlsx)
        wb.Close(False)
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(
            "Nie udało się automatycznie przekonwertować pliku .xls przez Microsoft Excel "
            f"(błąd: {exc}).\n"
            "Upewnij się, że Excel jest zainstalowany na tym komputerze, albo otwórz "
            "plik ręcznie i zapisz jako .xlsx."
        ) from exc
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:  # noqa: BLE001
                pass

    if not out_xlsx.exists():
        raise RuntimeError("Konwersja .xls -> .xlsx przez Excel nie powiodła się (brak pliku wynikowego).")
    return out_xlsx


def convert_xlsx(
    path: str | Path,
    output_path: str | Path | None = None,
    config: ConversionConfig | None = None,
    show_stats: bool = False,
) -> Path:
    """
    Konwertuje arkusz Excel (.xlsx lub stary .xls) na plik Markdown.
    Zwraca ścieżkę do zapisanego pliku .md
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Nie znaleziono pliku: {path}")

    config = config or ConversionConfig()

    if output_path is None:
        output_path = path.with_suffix(".md")
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    tmp_dir = output_path.parent / f".{output_path.stem}_tmp"
    tmp_dir.mkdir(parents=True, exist_ok=True)

    try:
        if path.suffix.lower() == ".xls":
            print(f"Przetwarzam: {path.name} (stary format .xls -> konwertuję przez Excel)")
            actual_xlsx_path = _convert_legacy_xls_to_xlsx(path, tmp_dir)
        else:
            print(f"Przetwarzam: {path.name}")
            actual_xlsx_path = path

        raw_md, sheet_reports = _convert_xlsx_to_markdown(actual_xlsx_path, config)
    finally:
        try:
            for f in tmp_dir.glob("*"):
                f.unlink()
            tmp_dir.rmdir()
        except OSError as exc:
            print(
                f"  [ostrzeżenie] nie udało się usunąć folderu tymczasowego {tmp_dir}: {exc}",
                file=sys.stderr,
            )

    final_md = raw_md
    if config.collapse_whitespace:
        final_md = _collapse_whitespace(final_md)

    output_path.write_text(final_md, encoding="utf-8")
    print(f"Zapisano: {output_path}")

    if show_stats:
        _print_stats_xlsx(path, final_md, sheet_reports, config)

    return output_path


def _print_stats_xlsx(
    path: Path, final_md: str, sheet_reports: list[dict], config: ConversionConfig
) -> None:
    file_size_kb = path.stat().st_size / 1024
    final_chars = len(final_md)
    final_tokens = _estimate_tokens(final_md)

    print("\n--- Statystyki ---")
    print(f"Rozmiar pliku:           {file_size_kb:.1f} KB")
    print(f"Arkuszy w pliku:         {len(sheet_reports)}")
    for s in sheet_reports:
        if "skipped" in s:
            print(f"  - {s['name']}: pominięty ({s['skipped']})")
        else:
            trunc = "  [OBCIĘTY]" if s.get("truncated") else ""
            print(f"  - {s['name']}: {s['rows']} wierszy x {s['cols']} kolumn{trunc}")
    print()
    print("Wynikowy plik Markdown:")
    print(f"  Znaków:                {_fmt_int(final_chars)}")
    print(f"  Szac. liczba tokenów:  ~{_fmt_int(final_tokens)}")
    print(
        f"\n* Uwaga: podobnie jak Word, arkusze Excel nie są w Claude renderowane jako\n"
        f"  obrazy stron - realna oszczędność tokenów bierze się głównie z pomijania\n"
        f"  ukrytych arkuszy, pustych wierszy/kolumn i obcinania bardzo dużych tabel\n"
        f"  (limit: {config.xlsx_max_rows_per_sheet} wierszy/arkusz), a nie z samego formatu Markdown."
    )


def convert_document(
    path: str | Path,
    output_path: str | Path | None = None,
    config: ConversionConfig | None = None,
    show_stats: bool = False,
) -> Path:
    """
    Rozpoznaje typ pliku po rozszerzeniu (.pdf / .docx / .doc / .xlsx / .xls) i
    wywołuje właściwy konwerter. To wspólny punkt wejścia dla CLI i launchera
    GUI, dzięki czemu oba nie muszą znać szczegółów obsługi każdego formatu.
    """
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return convert_pdf(path, output_path, config, show_stats)
    if suffix in (".docx", ".doc"):
        return convert_docx(path, output_path, config, show_stats)
    if suffix in (".xlsx", ".xls"):
        return convert_xlsx(path, output_path, config, show_stats)
    raise ValueError(f"Nieobsługiwany typ pliku: {suffix} (obsługiwane: .pdf, .docx, .doc, .xlsx, .xls)")


def convert_many(doc_paths: list[Path], output_dir: Path | None, config: ConversionConfig) -> None:
    ok, failed = [], []
    for doc_path in doc_paths:
        try:
            out_path = (output_dir / (doc_path.stem + ".md")) if output_dir else None
            convert_document(doc_path, out_path, config)
            ok.append(doc_path)
        except Exception as exc:  # noqa: BLE001
            print(f"BŁĄD przy pliku {doc_path.name}: {exc}", file=sys.stderr)
            failed.append(doc_path)

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
            "Uniwersalny konwerter PDF/DOCX/DOC/XLSX/XLS -> Markdown (tekst cyfrowy + OCR skanów + "
            "wierne tabele, dokumenty Word przez mammoth, arkusze Excel przez openpyxl), "
            "zoptymalizowany pod tokeny AI."
        ),
    )
    parser.add_argument("input_files", nargs="+",
                         help="Plik(i) PDF/DOCX/DOC/XLSX/XLS (obsługiwane wildcardy, np. *.pdf)")
    parser.add_argument("-o", "--output", help="Ścieżka pliku wyjściowego .md (tylko dla 1 pliku)")
    parser.add_argument("--output-dir", help="Katalog wyjściowy dla wielu plików")
    parser.add_argument("--lang", default="pol+eng", help="Języki OCR wg kodów Tesseract (domyślnie: pol+eng, dotyczy tylko PDF)")
    parser.add_argument("--dpi", type=int, default=300, help="Rozdzielczość renderowania skanów (domyślnie: 300, dotyczy tylko PDF)")
    parser.add_argument("--force-ocr", action="store_true", help="Wymuś OCR na WSZYSTKICH stronach (dotyczy tylko PDF)")
    parser.add_argument("--keep-headers", action="store_true",
                         help="Nie usuwaj powtarzających się nagłówków/stopek (dotyczy tylko PDF)")
    parser.add_argument("--max-rows", type=int, default=300,
                         help="Limit wierszy na arkusz Excel (domyślnie: 300, dotyczy tylko XLSX/XLS)")
    parser.add_argument("--stats", action="store_true", help="Pokaż statystyki (rozmiar, szac. liczba tokenów)")
    return parser


def main() -> None:
    parser = build_arg_parser()
    args = parser.parse_args()

    input_paths = [Path(p) for p in args.input_files]
    missing = [p for p in input_paths if not p.exists()]
    if missing:
        for m in missing:
            print(f"Nie znaleziono pliku: {m}", file=sys.stderr)
        sys.exit(1)

    supported = (".pdf", ".docx", ".doc", ".xlsx", ".xls")
    unsupported = [p for p in input_paths if p.suffix.lower() not in supported]
    if unsupported:
        for u in unsupported:
            print(f"Nieobsługiwany typ pliku: {u} (obsługiwane: {', '.join(supported)})", file=sys.stderr)
        sys.exit(1)

    config = ConversionConfig(
        lang=args.lang,
        dpi=args.dpi,
        force_ocr=args.force_ocr,
        strip_repeated_headers=not args.keep_headers,
        xlsx_max_rows_per_sheet=args.max_rows,
    )

    if len(input_paths) == 1 and not args.output_dir:
        convert_document(input_paths[0], args.output, config, show_stats=args.stats)
    else:
        output_dir = Path(args.output_dir) if args.output_dir else None
        if output_dir:
            output_dir.mkdir(parents=True, exist_ok=True)
        convert_many(input_paths, output_dir, config)


if __name__ == "__main__":
    main()
