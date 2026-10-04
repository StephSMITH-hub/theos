@echo off
title Theos - GitHub Auto-Sync & Push
cd /d "%~dp0"
echo ==========================================================
echo Starting Auto-Sync to GitHub (https://github.com/StephSMITH-hub/theos)
echo ==========================================================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0auto_sync_push.ps1" %*
echo.
pause
