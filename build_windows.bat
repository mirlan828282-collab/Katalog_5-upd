@echo off
setlocal
cd /d "%~dp0"
python -m pip install -r requirements.txt
if errorlevel 1 goto :error
python download_models.py
if errorlevel 1 goto :error
python -m PyInstaller --noconfirm --clean --onefile --windowed --name PhotoArchiveCatalog ^
  --add-data "models;models" ^
  --add-data "assets;assets" ^
  --hidden-import cv2 ^
  --hidden-import numpy ^
  --hidden-import openpyxl ^
  --hidden-import docx ^
  --hidden-import pyodbc ^
  --collect-all openpyxl ^
  main.py
if errorlevel 1 goto :error
echo.
echo READY: dist\PhotoArchiveCatalog.exe
exit /b 0
:error
echo.
echo BUILD ERROR
exit /b 1
