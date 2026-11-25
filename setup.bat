@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
title PPT-Agent dependency installer

echo ============================================================
echo   PPT-Agent renderer - dependency installer
echo ============================================================
echo.

REM ---- locate Python: prefer 3.12, the version the offline wheels target ----
REM .pi/wheels/py312 holds cp312 binaries. A 3.13/3.14 interpreter cannot
REM install them, so resolve 3.12 first and only then fall back.
set "PY="
py -3.12 -c "import sys" >nul 2>nul
if not errorlevel 1 ( set "PY=py -3.12" & goto pyfound )
py -3 -c "import sys" >nul 2>nul
if not errorlevel 1 ( set "PY=py -3" & goto pyfound )
where python >nul 2>nul
if not errorlevel 1 ( set "PY=python" & goto pyfound )
goto nopython

:pyfound
%PY% --version >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python command failed. Please check PATH.
    pause
    exit /b 1
)
echo [1/3] Python found:
%PY% --version
%PY% -c "import sys;print('      minor version: %d.%d' % sys.version_info[:2])"
echo.

REM ---- offline wheels present? then never touch the network ----
REM Wheels live in per-version subfolders (py312/ etc.): pip's --find-links does
REM NOT recurse, so the top .pi\wheels folder alone finds nothing.
set "WL=%~dp0.pi\wheels"
if exist "%WL%\*.whl" goto choose_wheels
if exist "%WL%\py312\*.whl" goto choose_wheels
goto online

:choose_wheels
set "PYMINOR="
for /f "delims=" %%V in ('%PY% -c "import sys;print(sys.version_info[1])"') do set "PYMINOR=%%V"
if defined PYMINOR if exist "%WL%\py3!PYMINOR!\*.whl" set "WL=%WL%\py3!PYMINOR!"
echo       wheels folder: %WL%
goto offline

:online
echo [2/3] No local wheels found - installing from PyPI (needs network)...
REM ---- internal mirror first, public PyPI as fallback ----
set "PIP_INDEX=--index-url http://172.30.39.34:18080/simple --trusted-host 172.30.39.34"
%PY% -m pip install --upgrade pip
%PY% -m pip install --disable-pip-version-check %PIP_INDEX% -r "%~dp0requirements.txt"
if errorlevel 1 (
    echo.
    echo [HINT] Mirror unreachable, falling back to public PyPI...
    %PY% -m pip install --disable-pip-version-check -r "%~dp0requirements.txt"
)
goto verify

:offline
echo [2/3] Offline mode: installing from the local wheels folder
echo       (no network access required)
%PY% -m pip install --disable-pip-version-check --no-index --find-links="%WL%" -r "%~dp0requirements.txt"
if errorlevel 1 (
    echo.
    echo [HINT] Usually means the interpreter is not the one the wheels target.
    echo        Check the "minor version" line above against the folder name ^(py312^).
    echo        A 3.13/3.14 interpreter needs its own wheel set - see .pi\wheels\README.md
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
echo   Render a deck (run from the repo root; engine/ is a package):
echo     %PY% -m engine.build .pi\work\slides -o .pi\out\deck.pptx --source .pi\work\analysis.md
echo   pi checks these dependencies automatically on startup
echo   (.pi\extensions\deps-bootstrap.ts); run this script to fix a broken env.
echo   Need an extra package? See .pi\wheels\README.md for the mirror command.
echo ============================================================
pause

goto :eof

:nopython
echo [ERROR] Python not found.
echo         Install CPython 3.12 (x64) from https://www.python.org/downloads/
echo         and tick "Add Python to PATH" during setup.
echo.
pause
exit /b 1
