@echo off
cd /d "%~dp0"

echo ================================================================
echo BUILD EXE - ANNUAIRE STATISTIQUE DE MEKNES 2025 HCP
echo ================================================================
echo.

if exist ".venv\Scripts\activate.bat" (
    call ".venv\Scripts\activate.bat"
)

python -m pip install --upgrade pyinstaller
if errorlevel 1 goto :error

python "scripts\build_exe_windows.py"
if errorlevel 1 goto :error

echo.
echo EXE cree dans le dossier dist.
pause
exit /b 0

:error
echo.
echo ERREUR pendant la construction.
pause
exit /b 1
