@echo off
chcp 65001 >nul
title Nepremičnine GUI – Zaženjalnik
color 0A

echo.
echo  ══════════════════════════════════════════════════════
echo     🏠  Nepremičnine GUI  –  Zaženjalnik
echo  ══════════════════════════════════════════════════════
echo.

:: Poišči Python
where py >nul 2>&1
if %errorlevel% == 0 (
    set PY=py
    goto :check_libs
)
where python >nul 2>&1
if %errorlevel% == 0 (
    set PY=python
    goto :check_libs
)
where python3 >nul 2>&1
if %errorlevel% == 0 (
    set PY=python3
    goto :check_libs
)

echo  ✗  Python ni najden!
echo.
echo     Namesti Python z: https://www.python.org/downloads/
echo     (Označi "Add Python to PATH" med namestitvijo!)
echo.
pause
exit /b 1

:check_libs
echo  Python: %PY%
echo.
echo  Preverjam in nameščam potrebne knjižnice ...
echo.

%PY% -m pip install --upgrade DrissionPage beautifulsoup4 lxml python-docx

if %errorlevel% neq 0 (
    echo.
    echo  ✗  Napaka pri namestitvi knjižnic!
    pause
    exit /b 1
)

echo.
echo  ✓  Vse knjižnice so nameščene.
echo.
echo  ▶  Zaganjam gui.py ...
echo.

%PY% -X utf8 "%~dp0gui.py"

if %errorlevel% neq 0 (
    echo.
    echo  ✗  Napaka pri zagonu gui.py (koda: %errorlevel%)
    pause
)

