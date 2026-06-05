@echo off
cd /d "%~dp0"
start "EnglishReader" cmd /c "npm run dev"
timeout /t 5 /nobreak
start http://127.0.0.1:3000
