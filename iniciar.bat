@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist "myenv\Scripts\pythonw.exe" (
    start "" "myenv\Scripts\pythonw.exe" app.py
) else (
    start "" pythonw app.py
)
