@echo off
title Theos - Log Chat/Update & Sync to GitHub
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0log_chat_update.ps1" %*
echo.
pause
