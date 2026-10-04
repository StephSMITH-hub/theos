@echo off
title Theos - Pull Updates from GitHub
cd /d "%~dp0"
echo ==========================================================
echo Pulling latest updates from GitHub (https://github.com/StephSMITH-hub/theos)
echo ==========================================================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0pull_latest_repo.ps1"
echo.
pause
