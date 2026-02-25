@echo off
setlocal
echo ============================================================
echo           AUTOFORM PRO BUILDER (v1.2.9)
echo ============================================================

echo [1/4] Detecting project location...
set "BASE_DIR=%~dp0"
cd /d "%BASE_DIR%"

echo [2/4] Activating Virtual Environment...
set "VENV_GLOBAL=D:\CODING\venv"
set "VENV_LOCAL=%BASE_DIR%venv"

if exist "%VENV_GLOBAL%\Scripts\activate.bat" (
    echo [INFO] Using Global Venv at %VENV_GLOBAL%
    call "%VENV_GLOBAL%\Scripts\activate.bat"
) else (
    if not exist "%VENV_LOCAL%\Scripts\activate.bat" (
        echo [INFO] Global Venv not found. Creating Local Venv...
        python -m venv "%VENV_LOCAL%"
        call "%VENV_LOCAL%\Scripts\activate.bat"
        echo [INFO] Installing dependencies from requirements.txt...
        python -m pip install --upgrade pip
        pip install -r requirements.txt
    ) else (
        echo [INFO] Using Local Venv at %VENV_LOCAL%
        call "%VENV_LOCAL%\Scripts\activate.bat"
    )
)

echo [3/4] Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist AutoForm.spec del /q AutoForm.spec

echo [4/4] Generating Executable (PyInstaller)...
echo --onefile: Single EXE output
echo --noconsole: Hide terminal
echo --collect-all customtkinter: Include CTK assets
echo --icon: Use icon.png

python -m PyInstaller --noconsole --onefile --clean --noconfirm ^
    --collect-all customtkinter ^
    --icon="icon.ico" ^
    --name "AutoFormPro" ^
    ui.py

echo ============================================================
echo BUILD COMPLETE! Check dist/ folder for AutoFormPro.exe
echo ============================================================
pause
