# pdf2md

Uniwersalny konwerter **PDF → Markdown** z integracją menu kontekstowego
Eksploratora Windows. Obsługuje zarówno dokumenty cyfrowe (z warstwą tekstu),
jak i skany (OCR), z naciskiem na wierne odtwarzanie **tabel** — szczególnie
przydatne przy raportach finansowych/księgowych przed wysłaniem ich do modeli AI.

## Funkcje

- **Automatyczna detekcja typu strony** (tekst cyfrowy vs. skan) — działa
  nawet dla PDF-ów mieszanych (część stron tekstowych, część skanów).
- **OCR** (Tesseract, domyślnie `pol+eng`) dla zeskanowanych stron.
- **Wierna rekonstrukcja tabel** — zarówno z PDF-ów cyfrowych (natywny
  detektor wektorowy PyMuPDF), jak i ze skanów (`img2table`), z automatycznym
  dopasowaniem rozdzielczości renderowania do natywnego DPI obrazu źródłowego
  (unika rozmycia linii tabeli przez niepotrzebną interpolację).
- **Optymalizacja pod tokeny AI**: usuwanie powtarzających się
  nagłówków/stopek (np. numeracja stron, nazwa firmy na każdej stronie),
  kompresja białych znaków, bez osadzania obrazów w treści.
- **Integracja z Eksploratorem Windows**: prawy klawisz na pliku `.pdf` →
  „Konwertuj do Markdown (pdf2md)" → plik `.md` powstaje w tym samym katalogu.

## Szybki start (Windows + integracja z Eksploratorem)

Zobacz pełną instrukcję: **[INSTALACJA.md](INSTALACJA.md)**

Skrót:
1. Zainstaluj [Python](https://www.python.org/downloads/) (zaznacz „Add to PATH”)
   i [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) (z pakietem
   językowym Polish).
2. `pip install pymupdf4llm pymupdf img2table pytesseract pandas tabulate pillow`
3. Sklonuj/pobierz to repo, uruchom `install_pdf2md.bat`.
4. Gotowe — kliknij prawym klawiszem na dowolny plik `.pdf`.

## Użycie z linii poleceń (bez integracji z Eksploratorem)

```bash
python pdf_to_markdown_universal.py raport.pdf --stats
python pdf_to_markdown_universal.py *.pdf --output-dir ./markdown_output
python pdf_to_markdown_universal.py skan.pdf --lang pol+eng --dpi 300
python pdf_to_markdown_universal.py raport.pdf --force-ocr        # wymuś OCR na każdej stronie
python pdf_to_markdown_universal.py raport.pdf --keep-headers     # nie czyść nagłówków/stopek
```

## Użycie jako moduł Python

```python
from pdf_to_markdown_universal import convert_pdf, ConversionConfig

config = ConversionConfig(lang="pol+eng", dpi=300)
convert_pdf("raport.pdf", "raport.md", config=config, show_stats=True)
```

## Struktura repozytorium

| Plik | Rola |
|---|---|
| `pdf_to_markdown_universal.py` | Silnik konwersji — CLI i moduł do importu |
| `pdf2md_launcher.pyw` | Okienko GUI uruchamiane z menu kontekstowego (bez konsoli) |
| `install_pdf2md.bat` | Instalator wpisu w menu kontekstowym Eksploratora |
| `uninstall_pdf2md.bat` | Usuwa wpis z menu kontekstowego |
| `INSTALACJA.md` | Pełna instrukcja instalacji i rozwiązywania problemów |

## Wymagania

- Python 3.10+
- [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) (z potrzebnymi
  pakietami językowymi, np. `pol`)
- Biblioteki Python: `pymupdf4llm`, `pymupdf`, `img2table`, `pytesseract`,
  `pandas`, `tabulate`, `pillow`

## Znane ograniczenia

- Detekcja tabel ze skanów nie jest w 100% niezawodna przy bardzo gęsto
  upakowanych kolumnach liczbowych lub niskiej jakości skanach — dla
  krytycznych dokumentów finansowych warto zweryfikować wynik.
- Kolejność elementów na zeskanowanej stronie wielokolumnowej jest uproszczona
  (sortowanie góra→dół po współrzędnej Y).

## Licencja

MIT — patrz [LICENSE](LICENSE).
