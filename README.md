# pdf2md

Uniwersalny konwerter **PDF / DOCX / DOC / XLSX / XLS → Markdown** z integracją
menu kontekstowego Eksploratora Windows. Zoptymalizowany pod kątem wysyłania
dokumentów do modeli AI (Claude i inne) — chodzi o to, żeby zamiast surowego
pliku binarnego wysłać zwarty, czytelny tekst, który kosztuje mniej tokenów
i jest łatwiejszy do przetworzenia.

## Funkcje

### PDF
- **Automatyczna detekcja typu strony** (tekst cyfrowy vs. skan) — działa
  nawet dla PDF-ów mieszanych (część stron tekstowych, część skanów).
- **OCR** (Tesseract, domyślnie `pol+eng`) dla zeskanowanych stron.
- **Wierna rekonstrukcja tabel** — zarówno z PDF-ów cyfrowych (natywny
  detektor wektorowy PyMuPDF), jak i ze skanów (`img2table`), z automatycznym
  dopasowaniem rozdzielczości renderowania do natywnego DPI obrazu źródłowego
  (unika rozmycia linii tabeli przez niepotrzebną interpolację).
- **Usuwanie powtarzających się nagłówków/stopek** (numeracja stron, nazwa
  firmy na każdej stronie) i kompresja białych znaków.
- **Statystyki oszczędności tokenów** (`--stats`) — porównanie: ile by
  kosztowało wysłanie tego samego PDF-a bezpośrednio do Claude (wg
  dokumentacji Anthropic każda strona PDF-a to 1500-3000 tokenów: tekst +
  obraz strony) vs. ile kosztuje wynikowy Markdown.

### Word (.docx / .doc)
- Konwersja przez `mammoth` (docx -> HTML z zachowaniem stylów) +
  `markdownify` (HTML -> Markdown), z mapą stylów rozpoznającą polskie i
  angielskie nazwy nagłówków ("Heading 1" / "Nagłówek 1" -> `#` itd.).
- Stary format `.doc` (binarny OLE) jest automatycznie konwertowany do
  `.docx` przez Microsoft Word w tle (COM/`pywin32`), zanim trafi do tego
  samego potoku co zwykłe `.docx`.

### Excel (.xlsx / .xls)
- Konwersja arkusz po arkuszu na tabele Markdown (`openpyxl`).
- **Pomijanie ukrytych arkuszy** - zwykle to dane robocze/pomocnicze, które
  i tak nie interesują AI.
- **Przycinanie pustych wierszy/kolumn na brzegach** - Excel często
  "pamięta" znacznie większy zakres niż faktyczne dane.
- **Twardy limit wierszy na arkusz** (domyślnie 300, `--max-rows`) - chroni
  przed tym, żeby jeden ogromny arkusz nie zjadł całego budżetu tokenów;
  obcięcie jest zawsze jasno opisane w wyniku.
- Stary format `.xls` konwertowany automatycznie do `.xlsx` przez Microsoft
  Excel w tle (COM/`pywin32`), tak samo jak `.doc` -> `.docx`.

### Wspólne dla wszystkich formatów
- **Integracja z Eksploratorem Windows**: prawy klawisz na pliku -> „pdf2md"
  -> plik `.md` powstaje w tym samym katalogu.
- Odporność na chwilowe blokady plików przez OneDrive/synchronizację przy
  sprzątaniu plików tymczasowych (ostrzeżenie zamiast przerwania konwersji).
- Poprawna obsługa polskich znaków niezależnie od tego, czy program jest
  uruchamiany z konsoli, czy z menu kontekstowego (bez okna konsoli).

## Jak to działa

```
plik wejściowy (.pdf/.docx/.doc/.xlsx/.xls)
        |
        v
  convert_document()  --> rozpoznaje typ po rozszerzeniu
        |
        +-- .pdf         -> convert_pdf()   -> pymupdf4llm (tekst) + img2table/Tesseract (skany)
        +-- .docx / .doc -> convert_docx()  -> mammoth + markdownify (.doc konwertowany przez Word)
        +-- .xlsx / .xls -> convert_xlsx()  -> openpyxl (.xls konwertowany przez Excel)
        |
        v
  wspólne czyszczenie: kompresja białych znaków, usuwanie obrazów z treści
        |
        v
  plik .md obok oryginału + opcjonalne statystyki (--stats)
```

Menu kontekstowe działa przez wpis w rejestrze Windows
(`HKEY_CURRENT_USER\...\SystemFileAssociations\<rozszerzenie>\shell\pdf2md`),
który uruchamia `pdf2md_launcher.pyw` przez `pythonw.exe` (bez okna konsoli)
z pełną ścieżką klikniętego pliku jako argumentem.

> **Uwaga dla użytkowników [Nilesoft Shell](https://nilesoft.org/):** jeśli
> masz zainstalowane to (bardzo popularne) narzędzie do personalizacji menu
> kontekstowego, ono **przechwytuje budowanie menu i nie zawsze pokazuje
> wpisy z rejestru automatycznie**, nawet po restarcie Eksploratora czy
> całego systemu. W takim wypadku najpewniejsze jest dopisanie wpisu wprost
> do `shell.nss` (zwykle `C:\Program Files\Nilesoft Shell\shell.nss`), np.:
> ```
> item(mode=mode.single type='file' where=sel.file.ext=='.pdf' title='pdf2md' cmd='C:\SCIEZKA\DO\pythonw.exe' arg='"C:\SCIEZKA\DO\pdf2md_launcher.pyw" "@sel.path"')
> ```
> (osobna linia dla każdego rozszerzenia: `.pdf`, `.docx`, `.doc`, `.xlsx`,
> `.xls`) - a następnie restart `explorer.exe`, bo Nilesoft Shell wczytuje
> `shell.nss` przy starcie powłoki, nie na żywo przy każdym menu.

## Szybki start (Windows + integracja z Eksploratorem)

Zobacz pełną instrukcję: **[INSTALACJA.md](INSTALACJA.md)**

Skrót:
1. Zainstaluj [Python](https://www.python.org/downloads/) (zaznacz „Add to PATH").
2. Do PDF: [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki)
   (z pakietem językowym Polish) + `pip install pymupdf4llm pymupdf img2table pytesseract pandas tabulate pillow`
3. Do Worda: `pip install mammoth markdownify`
4. Do Excela: `pip install openpyxl`
5. Do starych `.doc`/`.xls` (opcjonalnie, wymaga zainstalowanego Office):
   `pip install pywin32`
6. Sklonuj/pobierz to repo, uruchom `install_pdf2md.bat`.
7. Gotowe - kliknij prawym klawiszem na dowolny obsługiwany plik.

## Użycie z linii poleceń (bez integracji z Eksploratorem)

```bash
python pdf_to_markdown_universal.py raport.pdf --stats
python pdf_to_markdown_universal.py umowa.docx --stats
python pdf_to_markdown_universal.py budzet.xlsx --stats --max-rows 500
python pdf_to_markdown_universal.py *.pdf --output-dir ./markdown_output
python pdf_to_markdown_universal.py skan.pdf --lang pol+eng --dpi 300
python pdf_to_markdown_universal.py raport.pdf --force-ocr        # wymuś OCR na każdej stronie
python pdf_to_markdown_universal.py raport.pdf --keep-headers     # nie czyść nagłówków/stopek
```

## Użycie jako moduł Python

```python
from pdf_to_markdown_universal import convert_document, ConversionConfig

# Uniwersalnie - sam rozpozna typ pliku po rozszerzeniu:
convert_document("raport.pdf", show_stats=True)
convert_document("umowa.docx", show_stats=True)
convert_document("budzet.xlsx", show_stats=True)

# Albo bezpośrednio, z pełną kontrolą nad konfiguracją:
from pdf_to_markdown_universal import convert_pdf, convert_docx, convert_xlsx

config = ConversionConfig(lang="pol+eng", dpi=300, xlsx_max_rows_per_sheet=500)
convert_pdf("raport.pdf", "raport.md", config=config, show_stats=True)
convert_docx("umowa.docx", "umowa.md", show_stats=True)
convert_xlsx("budzet.xlsx", "budzet.md", config=config, show_stats=True)
```

## Struktura repozytorium

| Plik | Rola |
|---|---|
| `pdf_to_markdown_universal.py` | Silnik konwersji (PDF + DOCX/DOC + XLSX/XLS) - CLI i moduł do importu |
| `pdf2md_launcher.pyw` | Okienko GUI uruchamiane z menu kontekstowego (bez konsoli) |
| `install_pdf2md.bat` | Instalator wpisu w menu kontekstowym Eksploratora (wszystkie obsługiwane rozszerzenia) |
| `uninstall_pdf2md.bat` | Usuwa wpisy z menu kontekstowego |
| `INSTALACJA.md` | Pełna instrukcja instalacji i rozwiązywania problemów |

## Wymagania

- Python 3.10+
- **Do PDF:** [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki)
  (z potrzebnymi pakietami językowymi, np. `pol`) + biblioteki `pymupdf4llm`,
  `pymupdf`, `img2table`, `pytesseract`, `pandas`, `tabulate`, `pillow`
- **Do Worda (.docx):** biblioteki `mammoth`, `markdownify`
- **Do Excela (.xlsx):** biblioteka `openpyxl`
- **Do starych `.doc`/`.xls`:** dodatkowo zainstalowany Microsoft
  Word/Excel + biblioteka `pywin32` (automatycznie konwertuje w tle do
  nowego formatu przed dalszym przetwarzaniem; bez Office trzeba ręcznie
  zapisać plik w nowym formacie)

## O oszczędności tokenów

Format docelowy (Markdown) nie jest magicznie tańszy sam w sobie dla
każdego typu pliku:

- **PDF** wysłany bezpośrednio do Claude jest wyjątkowo drogi - każda
  strona liczona jest jako 1500-3000 tokenów (tekst + reprezentacja obrazu
  strony), niezależnie od gęstości tekstu. Tu konwersja na Markdown daje
  zwykle **60-90% oszczędności**.
- **Word i Excel** nie są renderowane jako obrazy stron, więc ich koszt
  tokenowy jest bliższy zwykłemu tekstowi. Tu realna oszczędność bierze się
  głównie z **usuwania szumu**: powtarzalnych nagłówków/stopek, pustych
  komórek, ukrytych arkuszy i obcinania nieproporcjonalnie dużych tabel -
  a nie z samego faktu bycia plikiem `.md`.

Narzędzie mówi to wprost w statystykach (`--stats`) dla każdego formatu,
zamiast obiecywać jednakowe oszczędności wszędzie.

## Znane ograniczenia

- Detekcja tabel ze skanów PDF nie jest w 100% niezawodna przy bardzo gęsto
  upakowanych kolumnach liczbowych lub niskiej jakości skanach.
- Kolejność elementów na zeskanowanej stronie wielokolumnowej jest
  uproszczona (sortowanie góra-dół po współrzędnej Y).
- Konwersja Worda zakłada, że nagłówki używają standardowych stylów
  ("Heading N" / "Nagłówek N") - dokumenty z całkowicie niestandardowym
  nazewnictwem stylów wyjdą jako zwykły tekst, bez struktury `#`/`##`.
- Konwersja Excela nie odtwarza formuł (tylko wyliczone wartości) ani
  scalonych komórek, wykresów czy formatowania warunkowego.
- Jeśli masz zainstalowany Nilesoft Shell lub inne narzędzie
  personalizujące menu kontekstowe, może ono wymagać osobnej konfiguracji
  - patrz sekcja "Jak to działa" wyżej.

## Licencja

MIT - patrz [LICENSE](LICENSE).
