@echo off
REM === Build ActivityLogger.exe using PyInstaller ===

REM Ensure you have pyinstaller installed:
REM    pip install pyinstaller

REM Clean previous build if needed
if exist ActivityLogger.spec del ActivityLogger.spec

REM === Only rebuild if sources changed ===
set "SRC=ActivityLogger.py"
set "EXE=dist\ActivityLogger.exe"

REM Check if EXE exists and if source is newer
if exist "%EXE%" (
    powershell -Command ^
        "$src=(Get-Item '%SRC%').LastWriteTimeUtc; $exe=(Get-Item '%EXE%').LastWriteTimeUtc; if ($src -le $exe) { exit 1 }"
    if %errorlevel%==1 (
        echo ActivityLogger.exe is up to date. Skipping build.
        goto end
    )
)

REM Clean previous build
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build

REM Build the executable (one file, windowed)
pyinstaller --noconfirm --windowed --onefile --name ActivityLogger ActivityLogger.py

:end
echo === Build complete! ===
