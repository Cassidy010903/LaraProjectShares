@echo off
python -m pip install openpyxl pyinstaller
python -m PyInstaller --onefile --windowed --name Sully --clean app.py
echo.
echo The program is in the dist folder: dist\Sully.exe
pause
