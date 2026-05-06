@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist "myenv\Scripts\python.exe" (
    start "Gerador de Certificados - Servidor" "myenv\Scripts\python.exe" app_web.py
) else (
    start "Gerador de Certificados - Servidor" python app_web.py
)

timeout /t 2 /nobreak >nul
start "" http://localhost:5000
