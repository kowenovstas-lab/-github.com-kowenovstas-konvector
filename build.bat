@echo off
setlocal

cd /d %~dp0

REM Create venv
if not exist venv (
  python -m venv venv
)

call venv\Scripts\activate

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

REM Build exe
python -m PyInstaller --onefile --windowed --name BatchConverter batch_convert_gui.py

echo.
echo Done. Your EXE is here:
echo %cd%\dist\BatchConverter.exe
echo.
pause
