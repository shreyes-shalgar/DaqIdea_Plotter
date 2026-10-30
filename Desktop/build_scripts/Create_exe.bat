@echo off
setlocal

REM Resolve the desktop app directory relative to this script.
for %%I in ("%~dp0..") do set "APP_DIR=%%~fI"
set "OUTPUT_DIR=%APP_DIR%\dist"

if not exist "%APP_DIR%\daqidea_plotter.py" (
	echo ERROR: Could not find "%APP_DIR%\daqidea_plotter.py".
	exit /b 1
)

python -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
	echo ERROR: PyInstaller is not installed for the active Python interpreter.
	echo Install it with: python -m pip install pyinstaller
	exit /b 1
)

python -m PyInstaller ^
	--noconfirm ^
	--clean ^
	--onefile ^
	--windowed ^
	--name DaqIdeaPlotter ^
	--collect-all tkinterdnd2 ^
	--distpath "%OUTPUT_DIR%" ^
	--workpath "%APP_DIR%\build" ^
	--specpath "%APP_DIR%\build" ^
	"%APP_DIR%\daqidea_plotter.py"

if errorlevel 1 (
	echo ERROR: PyInstaller failed to build DaqIdeaPlotter.exe.
	exit /b 1
)

if not exist "%OUTPUT_DIR%\DaqIdeaPlotter.exe" (
	echo ERROR: Build completed without producing DaqIdeaPlotter.exe.
	exit /b 1
)

echo Build successful: "%OUTPUT_DIR%\DaqIdeaPlotter.exe"
exit /b 0