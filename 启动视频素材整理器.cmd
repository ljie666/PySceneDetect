@echo off
setlocal
if not exist "%~dp0.venv\Scripts\pythonw.exe" (
    echo Project Python environment is missing.
    pause
    exit /b 1
)
if not exist "%~dp0video_organizer.py" (
    echo video_organizer.py is missing.
    pause
    exit /b 1
)
start "" /D "%~dp0." "%~dp0.venv\Scripts\pythonw.exe" "%~dp0video_organizer.py"
