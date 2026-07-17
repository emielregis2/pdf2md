@echo off
chcp 65001 >nul
title pdf2md - odinstalowanie

echo ============================================================
echo   pdf2md - usuwanie wpisu z menu kontekstowego
echo ============================================================
echo.

reg delete "HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations\.pdf\shell\pdf2md" /f >nul 2>&1
set "PDF_RESULT=%ERRORLEVEL%"
reg delete "HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations\.docx\shell\pdf2md" /f >nul 2>&1
set "DOCX_RESULT=%ERRORLEVEL%"
reg delete "HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations\.doc\shell\pdf2md" /f >nul 2>&1
set "DOC_RESULT=%ERRORLEVEL%"
reg delete "HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations\.xlsx\shell\pdf2md" /f >nul 2>&1
set "XLSX_RESULT=%ERRORLEVEL%"
reg delete "HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations\.xls\shell\pdf2md" /f >nul 2>&1
set "XLS_RESULT=%ERRORLEVEL%"

if %PDF_RESULT%==0 (
    echo Wpis dla .pdf zostal usuniety.
) else (
    echo Wpis dla .pdf nie zostal znaleziony ^(byc moze juz usuniety^).
)
if %DOCX_RESULT%==0 (
    echo Wpis dla .docx zostal usuniety.
) else (
    echo Wpis dla .docx nie zostal znaleziony ^(byc moze juz usuniety^).
)
if %DOC_RESULT%==0 (
    echo Wpis dla .doc zostal usuniety.
) else (
    echo Wpis dla .doc nie zostal znaleziony ^(byc moze juz usuniety^).
)
if %XLSX_RESULT%==0 (
    echo Wpis dla .xlsx zostal usuniety.
) else (
    echo Wpis dla .xlsx nie zostal znaleziony ^(byc moze juz usuniety^).
)
if %XLS_RESULT%==0 (
    echo Wpis dla .xls zostal usuniety.
) else (
    echo Wpis dla .xls nie zostal znaleziony ^(byc moze juz usuniety^).
)

echo.
pause
