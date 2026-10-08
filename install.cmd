@echo off
setlocal
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" goto deps
py -3 --version >nul 2>&1
if errorlevel 1 goto pythonfallback
py -3 -m venv .venv
if errorlevel 1 goto failed
goto deps
:pythonfallback
python --version >nul 2>&1
if errorlevel 1 goto missing
python -m venv .venv
if errorlevel 1 goto failed
:deps
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto failed
echo Installation complete. Run start.cmd.
pause
exit /b 0
:missing
echo Install Python 3.11 or newer, then run this file again.
echo https://www.python.org/downloads/
pause
exit /b 1
:failed
echo Installation failed. Check the error above.
pause
exit /b 1
