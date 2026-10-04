@echo off
title Theos - Continuous Auto-Sync Watcher
cd /d "%~dp0"
echo ==========================================================
echo Starting Continuous Background Sync Watcher
echo Checking for changes every 2 minutes and pushing to GitHub
echo Press Ctrl+C at any time to stop
echo ==========================================================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0auto_sync_push.ps1" -Watch -IntervalSeconds 120
pause
