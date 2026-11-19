@echo off
rem Builds dist\OutlookArchiver.exe on a Windows computer with Python installed.
cd /d "%~dp0"
python -m pip install --upgrade pywin32 pyinstaller || goto :error
python -m PyInstaller --noconfirm --onefile --windowed --name OutlookArchiver ^
  --hidden-import win32timezone --add-data "archiver/web;archiver/web" OutlookArchiver.py || goto :error
echo.
echo Done: dist\OutlookArchiver.exe
pause
exit /b 0
:error
echo Build failed.
pause
exit /b 1
