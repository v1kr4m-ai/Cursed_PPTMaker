@echo off
rem Starts the Cursed PPTMaker window (no console). For the old console version use "Make PPT.bat".
cd /d "%~dp0"
where pythonw >nul 2>nul && (start "" pythonw "%~dp0app.py" & exit /b)
rem pythonw not found: fall back to python, keeping this window open for errors
python "%~dp0app.py"
if errorlevel 1 pause
