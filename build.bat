@echo off
cd /d "%~dp0"

if not exist venv (
  py -3.12 -m venv venv
)

call venv\Scripts\activate

py -3.12 -m pip install --upgrade pip
py -3.12 -m pip install pandas openpyxl pyinstaller tkcalendar

rem If PyInstaller >= 6: use --collect-data
rem If PyInstaller < 6: change to --collect-all tkcalendar
py -3.12 -m PyInstaller --onefile --noconsole --clean --name FilepulldownQC ^
  --hidden-import openpyxl.cell._writer ^
  --hidden-import tkcalendar ^
  --collect-data tkcalendar ^
  "%~dp0FilepulldownQC.py"

echo.
echo ✅ Build complete!
echo EXE created at: "%~dp0dist\FilepulldownQC.exe"
pause
