@echo off
call .venv\Scripts\activate.bat
pyinstaller ^
  --noconfirm ^
  --windowed ^
  --name CumminsInvoiceStudio ^
  --icon assets\app_icon.ico ^
  --add-data "assets;assets" ^
  --add-data "sample_data;sample_data" ^
  qt_app.py

echo.
echo Windows build completed.
echo Output: dist\CumminsInvoiceStudio
