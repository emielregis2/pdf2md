@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
title pdf2md - instalacja menu kontekstowego

echo ============================================================
echo   pdf2md - instalacja w menu kontekstowym Eksploratora
echo ============================================================
echo.

rem --- 1) Znajdz pythonw.exe (wersja bez konsoli) ---
set "PYTHONW="
for /f "delims=" %%i in ('where pythonw 2^>nul') do (
    if not defined PYTHONW set "PYTHONW=%%i"
)

if not defined PYTHONW (
    echo [BLAD] Nie znaleziono pythonw.exe w PATH.
    echo         Zainstaluj Python z https://www.python.org/downloads/
    echo         i podczas instalacji zaznacz "Add python.exe to PATH".
    echo.
    pause
    exit /b 1
)

echo Znaleziono pythonw.exe: %PYTHONW%

rem --- 2) Ustal katalog, w ktorym leza skrypty (ten sam, co install_pdf2md.bat) ---
set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
set "LAUNCHER=%SCRIPT_DIR%\pdf2md_launcher.pyw"

if not exist "%LAUNCHER%" (
    echo [BLAD] Nie znaleziono pliku: %LAUNCHER%
    echo         Upewnij sie, ze install_pdf2md.bat jest w tym samym
    echo         katalogu co pdf2md_launcher.pyw i pdf_to_markdown_universal.py
    echo.
    pause
    exit /b 1
)

echo Katalog skryptow: %SCRIPT_DIR%

rem --- 3) Sprawdz zaleznosci Pythona (tylko informacyjnie) ---
set "PYTHON_EXE="
for /f "delims=" %%i in ('where python 2^>nul') do (
    if not defined PYTHON_EXE set "PYTHON_EXE=%%i"
)

echo.
echo Sprawdzam zainstalowane biblioteki Pythona...
if defined PYTHON_EXE (
    "%PYTHON_EXE%" -c "import fitz, pymupdf4llm, img2table, pytesseract, pandas" 2>nul
    if errorlevel 1 (
        echo [OSTRZEZENIE] Brakuje niektorych bibliotek Pythona.
        echo               Zainstaluj je poleceniem:
        echo               pip install pymupdf4llm pymupdf img2table pytesseract pandas tabulate pillow
        echo.
        echo Instalacja menu kontekstowego bedzie kontynuowana, ale konwersja
        echo nie zadziala, dopoki nie doinstalujesz brakujacych bibliotek.
        echo.
        pause
    )
) else (
    echo [OSTRZEZENIE] Nie znaleziono python.exe w PATH - pomijam sprawdzenie bibliotek.
)

rem --- 4) Zbuduj plik .reg z poprawnymi, wyescapowanymi sciezkami ---
set "REG_FILE=%TEMP%\pdf2md_install.reg"

rem W plikach .reg backslash trzeba podwajac, a cudzyslow poprzedzac \"
set "PYTHONW_ESC=%PYTHONW:\=\\%"
set "LAUNCHER_ESC=%LAUNCHER:\=\\%"

> "%REG_FILE%" (
    echo Windows Registry Editor Version 5.00
    echo.
    echo [HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations\.pdf\shell\pdf2md]
    echo @="Konwertuj do Markdown ^(pdf2md^)"
    echo "Icon"="%PYTHONW_ESC%,0"
    echo.
    echo [HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations\.pdf\shell\pdf2md\command]
    echo @="\"%PYTHONW_ESC%\" \"%LAUNCHER_ESC%\" \"%%1\""
)

echo.
echo Rejestruje wpis w rejestrze ^(HKEY_CURRENT_USER - bez uprawnien admina^)...
reg import "%REG_FILE%"

if errorlevel 1 (
    echo [BLAD] Nie udalo sie zaimportowac wpisu do rejestru.
    del "%REG_FILE%" >nul 2>&1
    pause
    exit /b 1
)

del "%REG_FILE%" >nul 2>&1

echo.
echo ============================================================
echo   GOTOWE! Kliknij prawym klawiszem na dowolny plik .pdf
echo   i wybierz "Konwertuj do Markdown (pdf2md)"
echo ============================================================
echo.
echo Uwaga: jesli uzywasz aplikacji "Files" (files.community) zamiast
echo natywnego Eksploratora, moze byc potrzebne jej ponowne uruchomienie,
echo zeby wczytala nowy wpis z rejestru.
echo.
pause
