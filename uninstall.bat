@echo off
echo Removing FileConverter from the right-click menu...
"%LOCALAPPDATA%\FileConverter\venv\Scripts\python.exe" "%~dp0convert.py" --uninstall
echo.
echo To remove everything, also delete the folder %LOCALAPPDATA%\FileConverter
pause
