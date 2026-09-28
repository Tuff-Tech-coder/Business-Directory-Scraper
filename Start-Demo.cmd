@echo off
setlocal
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" goto install
where py >nul 2>nul
if not errorlevel 1 goto usepy
where python >nul 2>nul
if not errorlevel 1 goto usepython
if exist "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" goto usebundled
echo Python 3.11 or newer is required. See README.md for setup instructions.
pause
exit /b 1
:usepy
py -3 -m venv .venv
goto checkenv
:usepython
python -m venv .venv
goto checkenv
:usebundled
"%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" -m venv .venv
:checkenv
if errorlevel 1 goto failed
:install
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto failed
echo.
echo Open http://127.0.0.1:8765 in your browser after the server starts.
echo Press Ctrl+C in this window to stop the demo.
".venv\Scripts\python.exe" directory_scraper.py --serve
if errorlevel 1 goto failed
exit /b 0
:failed
echo.
echo The demo could not start. Check the error above and README.md.
pause
exit /b 1
