# pdf2md — konwersja PDF/DOCX/DOC → Markdown z menu kontekstowego Eksploratora

Po instalacji: **prawy klawisz na pliku `.pdf`/`.docx`/`.doc` → "pdf2md"**
→ w tym samym katalogu powstaje plik `.md`.

Działa zarówno w natywnym Eksploratorze Windows, jak i w aplikacji **Files**
(files.community), bo obie korzystają z tego samego mechanizmu rejestru Windows
(dokładnie tak, jak widoczne u Ciebie wpisy WinRAR, 7-Zip, Balabolka).

---

## Co jest w paczce

| Plik | Rola |
|---|---|
| `pdf_to_markdown_universal.py` | Silnik konwersji (logika: OCR, tabele, itd.) |
| `pdf2md_launcher.pyw` | Okienko uruchamiane z menu kontekstowego (bez konsoli) |
| `install_pdf2md.bat` | Rejestruje wpis w menu kontekstowym |
| `uninstall_pdf2md.bat` | Usuwa wpis z menu kontekstowego |

**Wszystkie 4 pliki muszą leżeć w jednym katalogu** — np. `C:\Narzedzia\pdf2md\`.
Nie przenoś ich osobno; `install_pdf2md.bat` sam wykrywa swoją lokalizację
i zapisuje ją w rejestrze jako ścieżkę do `pdf2md_launcher.pyw`.

---

## Instalacja krok po kroku

### 1. Python

Jeśli nie masz Pythona: pobierz z [python.org/downloads](https://www.python.org/downloads/)
i **podczas instalacji zaznacz "Add python.exe to PATH"** — to ważne, bez tego
skrypt instalacyjny nie znajdzie interpretera.

Sprawdź w terminalu (PowerShell / cmd):
```
python --version
```

### 2. Tesseract OCR (silnik rozpoznawania tekstu ze skanów)

Pobierz instalator z: **https://github.com/UB-Mannheim/tesseract/wiki**
(oficjalna dystrybucja Tesseract dla Windows).

Podczas instalacji, w sekcji **"Additional language data"** zaznacz co najmniej:
- Polish
- (English jest zwykle domyślnie zaznaczony)

Domyślna ścieżka instalacji to `C:\Program Files\Tesseract-OCR\` — skrypt
konwertera sam ją wykryje, nawet jeśli instalator nie dopisał jej do PATH
(to częsty problem na Windows — już to obsłużyliśmy w kodzie).

### 3. Biblioteki Pythona

Otwórz terminal (PowerShell) i uruchom:
```
pip install pymupdf4llm pymupdf img2table pytesseract pandas tabulate pillow
```

Jeśli chcesz też konwertować pliki Word (`.docx`), dodatkowo:
```
pip install mammoth markdownify
```

Do starych plików `.doc` (nie `.docx`) potrzebny jest zainstalowany Microsoft
Word oraz `pip install pywin32` — bez tego stare `.doc` nie zostaną obsłużone
(trzeba je wtedy ręcznie zapisać w Wordzie jako `.docx`).

### 4. Umieść pliki i uruchom instalator

1. Skopiuj wszystkie 4 pliki z tej paczki do jednego katalogu, np. `C:\Narzedzia\pdf2md\`
2. Kliknij dwukrotnie **`install_pdf2md.bat`**
3. Skrypt sam:
   - znajdzie `pythonw.exe`,
   - sprawdzi, czy biblioteki są zainstalowane (ostrzeże, jeśli nie),
   - doda wpis do rejestru **HKEY_CURRENT_USER** (bez uprawnień administratora)
4. Gotowe — kliknij prawym klawiszem na dowolny plik `.pdf`.

Jeśli używasz aplikacji **Files**, może być potrzebne jej ponowne uruchomienie,
żeby wczytała nowy wpis z rejestru.

---

## Odinstalowanie

Uruchom `uninstall_pdf2md.bat` — usuwa tylko wpis rejestru, nic więcej
(pliki `.py`/`.pyw` możesz wtedy skasować ręcznie).

---

## Jak to działa "pod maską"

`install_pdf2md.bat` dopisuje do rejestru (klucz `HKEY_CURRENT_USER`, więc bez
uprawnień admina) coś w rodzaju:

```
[HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations\.pdf\shell\pdf2md]
@="pdf2md"

[HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations\.pdf\shell\pdf2md\command]
@="\"C:\\Python312\\pythonw.exe\" \"C:\\Narzedzia\\pdf2md\\pdf2md_launcher.pyw\" \"%1\""
```

`SystemFileAssociations\.pdf\shell\...` to standardowy, udokumentowany
mechanizm Windows do dodawania akcji dla konkretnego typu pliku — działa
niezależnie od tego, jaki program jest domyślnie skojarzony z PDF-ami
(u Ciebie to Chrome, ale to nie ma znaczenia).

Po kliknięciu Windows uruchamia `pythonw.exe` (wersja **bez** okna konsoli)
ze ścieżką klikniętego pliku jako argumentem. `pdf2md_launcher.pyw` otwiera
małe okienko z logiem konwersji na żywo i przyciskiem "Otwórz plik wynikowy"
po zakończeniu.

---

## Rozwiązywanie problemów

**Menu pokazuje pozycję, ale nic się nie dzieje po kliknięciu**
→ Sprawdź, czy `pythonw.exe` i wszystkie 4 pliki nadal są w tej samej
lokalizacji, co podczas instalacji. Jeśli przeniosłeś katalog — uruchom
`install_pdf2md.bat` ponownie (nadpisze stary wpis nową ścieżką).

**Okienko pokazuje błąd o brakujących bibliotekach**
→ Wróć do kroku 3 (`pip install ...`). Upewnij się, że instalujesz do tej
samej instalacji Pythona, którą wykrył `install_pdf2md.bat` (widoczna
w oknie terminala podczas instalacji).

**Konwersja skanów nie działa / błąd "tesseract not found"**
→ Zainstaluj Tesseract OCR (krok 2). Jeśli zainstalowałeś w niestandardowej
lokalizacji, dopisz ją ręcznie w `pdf_to_markdown_universal.py`, w funkcji
`_ensure_tesseract_available()` (lista `candidates`).

**Chcę zmienić język OCR (np. dodać niemiecki)**
→ W `pdf2md_launcher.pyw` znajdź linię:
```python
config = ConversionConfig(lang="pol+eng", dpi=300)
```
i zmień np. na `lang="pol+eng+deu"` (wymaga doinstalowania pakietu
`tesseract-ocr-deu` — analogicznie jak dla polskiego, przez instalator
UB Mannheim z zaznaczoną opcją German).

**Chcę, żeby plik .md powstawał w innym katalogu niż PDF (np. zawsze na Pulpicie)**
→ Daj znać, dopiszę taką opcję do launchera (obecnie zawsze zapisuje obok
oryginału, zgodnie z tym, o co prosiłeś).
