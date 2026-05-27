@echo off
setlocal

set TARGET=C:\Build\konvector

if not exist %TARGET% (
  mkdir %TARGET%
)

cd /d %TARGET%

REM Create venv in target
if not exist venv (
  python -m venv venv
)

call venv\Scripts\activate

python -m pip install --upgrade pip
python -m pip install -r "%~dp0requirements.txt"

REM Copy source files into target (so PyInstaller uses clean paths)
copy /Y "%~dp0batch_convert_gui.py" "%TARGET%\batch_convert_gui.py" >nul
copy /Y "%~dp0requirements.txt" "%TARGET%\requirements.txt" >nul

REM Build exe
python -m PyInstaller --onefile --windowed --name BatchConverter "%TARGET%\batch_convert_gui.py"

echo.
echo Done. Your EXE is here:
echo %TARGET%\dist\BatchConverter.exe
echo.
pause
