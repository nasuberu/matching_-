@echo off
start "マッチングアプリ" /D "%~dp0" cmd /k node --use-system-ca server.js
timeout /t 2 /nobreak >nul
start "" http://localhost:4001
