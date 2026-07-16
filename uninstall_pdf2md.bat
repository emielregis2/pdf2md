@echo off
chcp 65001 >nul
title pdf2md - odinstalowanie

echo ============================================================
echo   pdf2md - usuwanie wpisu z menu kontekstowego
echo ============================================================
echo.

reg delete "HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations\.pdf\shell\pdf2md" /f

if errorlevel 1 (
    echo.
    echo Nie znaleziono wpisu w rejestrze - byc moze pdf2md nie byl zainstalowany
    echo albo zostal juz usuniety.
) else (
    echo.
    echo Wpis "Konwertuj do Markdown (pdf2md)" zostal usuniety z menu kontekstowego.
)

echo.
pause
