# pdf2md — instalacja od zera na nowym komputerze

Po instalacji: **prawy klawisz na pliku `.pdf`/`.docx`/`.doc`/`.xlsx`/`.xls` → "pdf2md"**
→ w tym samym katalogu powstaje plik `.md`.

Ten przewodnik zakłada, że zaczynasz **całkowicie od zera** na nowym
komputerze - nic wcześniej nie jest zainstalowane ani skonfigurowane.

---

## Skrócona checklista (dla doświadczonych)

1. Python 3.10+ z zaznaczonym "Add python.exe to PATH"
2. `pip install pymupdf4llm pymupdf img2table pytesseract pandas tabulate pillow mammoth markdownify openpyxl`
3. (opcjonalnie) Tesseract OCR z pakietem `Polish` - do PDF-ów-skanów
4. (opcjonalnie) `pip install pywin32` + zainstalowany MS Office - do starych `.doc`/`.xls`
5. Sklonuj/pobierz repo do jednego katalogu, np. `C:\Narzedzia\pdf2md\`
6. Uruchom `install_pdf2md.bat`
7. Kliknij prawym klawiszem na dowolny obsługiwany plik

Jeśli coś nie zadziała od razu - sekcja **Rozwiązywanie problemów** niżej
pokrywa wszystkie realne przypadki, na jakie trafiliśmy podczas rozwoju
tego narzędzia.

---

## Co jest w paczce

| Plik | Rola |
|---|---|
| `pdf_to_markdown_universal.py` | Silnik konwersji (PDF/DOCX/DOC/XLSX/XLS) - CLI i moduł do importu |
| `pdf2md_launcher.pyw` | Okienko uruchamiane z menu kontekstowego (bez konsoli) |
| `install_pdf2md.bat` | Rejestruje wpisy w menu kontekstowym dla wszystkich obsługiwanych rozszerzeń |
| `uninstall_pdf2md.bat` | Usuwa wpisy z menu kontekstowego |
| `README.md` | Ogólny opis projektu i architektury |

**Te 4 pliki wykonywalne muszą leżeć razem, w jednym katalogu** - np.
`C:\Narzedzia\pdf2md\`. Nie przenoś ich osobno; `install_pdf2md.bat` sam
wykrywa swoją lokalizację i zapisuje ją w rejestrze jako ścieżkę do
`pdf2md_launcher.pyw`. Jeśli przeniesiesz katalog później, uruchom
`install_pdf2md.bat` ponownie.

---

## Instalacja krok po kroku

### 1. Python

Pobierz z [python.org/downloads](https://www.python.org/downloads/)
(wersja 3.10 lub nowsza) i **podczas instalacji zaznacz "Add python.exe to
PATH"** - to ważne, bez tego skrypt instalacyjny nie znajdzie interpretera.

Sprawdź w terminalu (PowerShell lub cmd):
```
python --version
```

### 2. Biblioteki Pythona - podstawa (PDF)

```
pip install pymupdf4llm pymupdf img2table pytesseract pandas tabulate pillow
```

### 3. Tesseract OCR - tylko jeśli chcesz konwertować zeskanowane PDF-y

Jeśli Twoje PDF-y mają zawsze warstwę tekstową (nie są skanami), możesz
pominąć ten krok - konwersja tekstu cyfrowego zadziała bez Tesseracta.

Pobierz instalator z: **https://github.com/UB-Mannheim/tesseract/wiki**
(oficjalna dystrybucja Tesseract dla Windows).

Podczas instalacji, w sekcji **"Additional language data"** zaznacz co
najmniej **Polish** (English jest zwykle domyślnie zaznaczony).

Domyślna ścieżka instalacji to `C:\Program Files\Tesseract-OCR\` - skrypt
sam ją wykryje, nawet jeśli instalator nie dopisał jej do PATH (to częsty
problem na Windows - już to obsłużyliśmy w kodzie).

### 4. Biblioteki Pythona - Word i Excel (opcjonalnie)

Do konwersji plików `.docx`:
```
pip install mammoth markdownify
```

Do konwersji arkuszy `.xlsx`:
```
pip install openpyxl
```

Do starych formatów `.doc` i `.xls` (nie `.docx`/`.xlsx`) potrzebny jest
**zainstalowany Microsoft Word/Excel** (prawdziwy Office, nie tylko
biblioteka) oraz:
```
pip install pywin32
```
Program automatycznie otworzy plik w Wordzie/Excelu w tle i zapisze go od
razu w nowym formacie. Bez zainstalowanego Office to się nie uda - trzeba
wtedy ręcznie otworzyć plik i zapisać jako `.docx`/`.xlsx`.

### 5. Umieść pliki i uruchom instalator

1. Skopiuj/sklonuj repo do jednego katalogu, np. `C:\Narzedzia\pdf2md\`
2. Kliknij dwukrotnie **`install_pdf2md.bat`**
3. Skrypt sam:
   - znajdzie `pythonw.exe`,
   - sprawdzi, które biblioteki są zainstalowane (ostrzeże, jeśli czegoś
     brakuje - ale i tak doda wpis do menu, więc możesz doinstalować
     brakujące pakiety później bez ponownej instalacji),
   - doda wpisy do rejestru **HKEY_CURRENT_USER** (bez uprawnień
     administratora) dla wszystkich pięciu rozszerzeń.
4. Gotowe - kliknij prawym klawiszem na dowolny obsługiwany plik.

---

## Odinstalowanie

Uruchom `uninstall_pdf2md.bat` - usuwa wpisy rejestru dla wszystkich
rozszerzeń, nic więcej (pliki `.py`/`.pyw` możesz wtedy skasować ręcznie).

---

## Jak to działa "pod maską"

`install_pdf2md.bat` dopisuje do rejestru (klucz `HKEY_CURRENT_USER`, więc
bez uprawnień admina) po jednym wpisie na każde rozszerzenie, w rodzaju:

```
[HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations\.pdf\shell\pdf2md]
@="pdf2md"

[HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations\.pdf\shell\pdf2md\command]
@="\"C:\\Python312\\pythonw.exe\" \"C:\\Narzedzia\\pdf2md\\pdf2md_launcher.pyw\" \"%1\""
```

(analogicznie dla `.docx`, `.doc`, `.xlsx`, `.xls`)

`SystemFileAssociations\<rozszerzenie>\shell\...` to standardowy,
udokumentowany mechanizm Windows do dodawania akcji dla konkretnego typu
pliku - działa niezależnie od tego, jaki program jest domyślnie skojarzony
z danym rozszerzeniem.

Po kliknięciu Windows uruchamia `pythonw.exe` (wersja **bez** okna konsoli)
ze ścieżką klikniętego pliku jako argumentem. `pdf2md_launcher.pyw` otwiera
małe okienko z logiem konwersji na żywo, statystykami i przyciskiem "Otwórz
plik wynikowy" po zakończeniu (okno zamyka się samo po 10 sekundach przy
sukcesie).

---

## Rozwiązywanie problemów

### Menu w ogóle nie pokazuje pozycji "pdf2md"

Najpierw sprawdź, czy wpis faktycznie jest w rejestrze - otwórz
`regedit`, przejdź do:
```
HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations\.pdf\shell\pdf2md
```
Jeśli wpisu nie ma - uruchom `install_pdf2md.bat` jeszcze raz.

**Jeśli wpis JEST w rejestrze, ale menu i tak go nie pokazuje** (nawet po
restarcie Eksploratora czy całego komputera) - najprawdopodobniej masz
zainstalowane narzędzie do personalizacji menu kontekstowego, które
**przechwytuje budowanie menu i ignoruje niektóre wpisy z rejestru**. Znane
przypadki:

- **[Nilesoft Shell](https://nilesoft.org/)** - bardzo popularne, darmowe
  narzędzie. Ono samo czyta rejestr dla "zwykłych", statycznie
  zarejestrowanych aplikacji, ale własne, ręcznie dopisane reguły trzeba
  dodać wprost do jego pliku konfiguracyjnego `shell.nss` (zwykle
  `C:\Program Files\Nilesoft Shell\shell.nss`), w sekcji `dynamic { ... }`:
  ```
  item(mode=mode.single type='file' where=sel.file.ext=='.pdf' title='pdf2md' cmd='C:\SCIEZKA\DO\pythonw.exe' arg='"C:\SCIEZKA\DO\pdf2md_launcher.pyw" "@sel.path"')
  ```
  (osobna linia dla każdego rozszerzenia: `.pdf`, `.docx`, `.doc`, `.xlsx`,
  `.xls`, z tymi samymi ścieżkami co w kroku instalacji). **Ważne:**
  Nilesoft Shell wczytuje `shell.nss` przy starcie `explorer.exe`, a nie na
  żywo przy każdym otwarciu menu - po edycji trzeba zrestartować Eksplorator
  (Menedżer zadań → znajdź `explorer.exe` → Uruchom ponownie zadanie, albo
  po prostu zrestartuj komputer).
- **Inne podobne narzędzia** (np. różne "Context Menu Manager") działają na
  podobnej zasadzie - sprawdź ich dokumentację/plik konfiguracyjny.
- Jeśli używasz **aplikacji "Files"** (files.community) zamiast natywnego
  Eksploratora, może być potrzebne jej ponowne uruchomienie (zamknij i
  otwórz od nowa), żeby wczytała nowy wpis z rejestru.

### Menu pokazuje starą nazwę / stary wpis mimo aktualizacji

To ten sam mechanizm co wyżej - jeśli masz Nilesoft Shell lub podobne
narzędzie, sprawdź czy w jego pliku konfiguracyjnym nie ma osobnej,
ręcznie dopisanej reguły z inną nazwą, którą trzeba poprawić ręcznie.
Sama aktualizacja rejestru przez `install_pdf2md.bat` tego nie naprawi.

### Okienko pokazuje błąd o brakujących bibliotekach

Wróć do kroków 2/4 (`pip install ...`) - upewnij się, że instalujesz do tej
samej instalacji Pythona, którą wykrył `install_pdf2md.bat` (widoczna
w oknie terminala podczas instalacji, np. `python --version` powinno
pokazać tę samą wersję).

### Błąd "Tesseract 'pol' trained data cannot be located"

To znaczy, że silnik Tesseract jest zainstalowany, ale nie może znaleźć
pakietu językowego. Najczęstsza przyczyna: zmienna środowiskowa
`TESSDATA_PREFIX` jest już ustawiona w systemie (przez inny program albo
starą, ręczną konfigurację) i wskazuje na zły/nieistniejący katalog - a
skrypt domyślnie nie nadpisuje już ustawionej zmiennej.

Sprawdź w PowerShell:
```powershell
[Environment]::GetEnvironmentVariable("TESSDATA_PREFIX","User")
```
Jeśli wskazuje na coś innego niż faktyczny katalog `tessdata` Tesseracta
(zwykle `C:\Program Files\Tesseract-OCR\tessdata`), popraw ją:
```powershell
[Environment]::SetEnvironmentVariable("TESSDATA_PREFIX", "C:\Program Files\Tesseract-OCR\tessdata", "User")
```
i uruchom konwersję ponownie (może być potrzebny nowy terminal/restart
Eksploratora, żeby nowy proces odziedziczył poprawną wartość).

### Konwersja skanów nie działa / błąd "tesseract not found"

Zainstaluj Tesseract OCR (krok 3). Jeśli zainstalowałeś w niestandardowej
lokalizacji, dopisz ją ręcznie w `pdf_to_markdown_universal.py`, w funkcji
`_ensure_tesseract_available()` (lista `candidates`).

### Błąd przy sprzątaniu folderu tymczasowego (np. w OneDrive)

To nie jest błąd krytyczny - jeśli plik leży w folderze zsynchronizowanym
z OneDrive (lub podobną usługą), która chwilowo blokuje pliki podczas
indeksowania, zobaczysz ostrzeżenie w logu, ale **konwersja i tak się
dokończy poprawnie**. Pusty folder tymczasowy możesz usunąć ręcznie
później albo zignorować.

### Błąd przy konwersji starego `.doc`/`.xls`

Sprawdź, czy masz zainstalowanego prawdziwego Microsoft Word/Excel (nie
tylko `pywin32`) - automatyczna konwersja `.doc`→`.docx` i `.xls`→`.xlsx`
wymaga faktycznie działającej aplikacji Office w tle. Jeśli jej nie masz,
otwórz plik ręcznie w dowolnym edytorze obsługującym stare formaty i
zapisz jako `.docx`/`.xlsx`, a potem przekonwertuj ten nowy plik.

### Chcę zmienić język OCR (np. dodać niemiecki)

W `pdf2md_launcher.pyw` znajdź linię:
```python
config = ConversionConfig(lang="pol+eng", dpi=300)
```
i zmień np. na `lang="pol+eng+deu"` (wymaga doinstalowania pakietu
`tesseract-ocr-deu` - analogicznie jak dla polskiego, przez instalator
UB Mannheim z zaznaczoną opcją German).

### Chcę zmienić limit wierszy dla Excela

W tej samej linii w `pdf2md_launcher.pyw` dodaj parametr
`xlsx_max_rows_per_sheet`, np.:
```python
config = ConversionConfig(lang="pol+eng", dpi=300, xlsx_max_rows_per_sheet=1000)
```

### Chcę, żeby plik .md powstawał w innym katalogu niż oryginał

Obecnie zawsze zapisuje obok oryginału. Jeśli potrzebujesz innego
zachowania, zmień wywołanie `convert_document()` w `pdf2md_launcher.pyw`,
podając własną ścieżkę `output_path`.
