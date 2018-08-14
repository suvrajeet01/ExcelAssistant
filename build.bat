@echo off

echo "Requires Python 3. Tested with Python 3.6."
pause

pip install openpyxl pyinstaller
pyinstaller -F --icon=icon.ico main.py
del /q dist\ExcelAssistant.exe
rename dist\main.exe ExcelAssistant.exe

pause