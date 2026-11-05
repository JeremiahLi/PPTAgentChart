@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
title PPT-Agent dependency installer

echo ============================================================
echo   PPT-Agent renderer - dependency installer
echo ============================================================
echo.

REM ---- locate Python ----
set "PY=python"
where python >nul 2>nul
if errorlevel 1 (
    set "PY=py -3"
    where py >nul 2>nul
    if errorlevel 1 (
        echo [ERROR] Python 3.10+ not found.
        echo         Install it from https://www.python.org/downloads/
        echo         and tick "Add Python to PATH" during setup.
        echo.
        pause
        exit /b 1
    )
)
%PY% --version >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python command failed. Please check PATH.
    pause
    exit /b 1
)
echo [1/3] Python found:
%PY% --version
echo.

REM ---- offline wheels present? then never touch the network ----
REM Wheels live in per-version subfolders (py311/ py312/ py313/ py314/):
REM pip's --find-links does NOT recurse, so passing the top folder alone finds
REM nothing on 3.12/3.13 (the old script checked python_pptx-*.whl at the top
REM level and silently fell through to the online branch).
set "WL=%~dp0.pi\wheels"
if exist "%WL%\*.whl" goto choose_wheels
if exist "%WL%\py311\*.whl" goto choose_wheels
goto online

:choose_wheels
set "PYMINOR="
for /f "delims=" %%V in ('%PY% -c "import sys;print(sys.version_info[1])"') do set "PYMINOR=%%V"
if defined PYMINOR if exist "%WL%\py3!PYMINOR!\*.whl" set "WL=%WL%\py3!PYMINOR!"
echo       wheels folder: %WL%
goto offline

:online
echo [2/3] No local wheels found - installing from PyPI (needs network)...
REM ---- optional mirror (uncomment and edit if PyPI is slow/blocked) ----
REM set "PIP_INDEX=-i https://pypi.tuna.tsinghua.edu.cn/simple"
%PY% -m pip install --upgrade pip
%PY% -m pip install %PIP_INDEX% -r "%~dp0requirements.txt"
goto verify

:offline
echo [2/3] Offline mode: installing from the local wheels folder
echo       (no network access required)
%PY% -m pip install --no-index --find-links="%WL%" -r "%~dp0requirements.txt"
if errorlevel 1 (
    echo.
    echo [HINT] If pip is too old to read the wheels, install it from the folder too:
    echo        %PY% -m pip install --no-index --find-links="%WL%" pip
)

:verify
echo.
echo [3/3] Verifying imports...
%PY% -c "import pptx, lxml, pydantic, matplotlib, numpy, PIL, yaml; print('ALL OK')"
if errorlevel 1 (
    echo [ERROR] Verification failed. See messages above.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   DONE. All dependencies installed and verified.
echo.
echo   Render a deck:
echo     %PY% .pi\renderer\render.py .pi\work\deck.json -o .pi\out\deck.pptx --qc
echo   With source-coverage check (seminar decks):
echo     %PY% .pi\renderer\render.py .pi\work\deck.json -o .pi\out\deck.pptx --qc --source .pi\work\analysis.md
echo   Optional extras (docx / pdf / xlsx skills):
echo     %PY% -m pip install --no-index --find-links "%WL%" -r requirements-full.txt
echo ============================================================
pause
