@echo off
setlocal enabledelayedexpansion

where python >nul 2>nul
if errorlevel 1 (
  echo Python not found. Install Python 3.10+ from https://www.python.org/downloads/
  pause
  exit /b 1
)

python bank_pdf_to_excel_ocr.py
echo.
echo (Completed. Press any key to close.)
pause >nul
