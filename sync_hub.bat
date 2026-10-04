@echo off
title Theos - GitHub Sync & Chat Hub
cd /d "%~dp0"

:menu
cls
echo =====================================================================
echo                THEOS REPOSITORY - GITHUB SYNC HUB                    
echo          Repo: https://github.com/StephSMITH-hub/theos               
echo =====================================================================
echo.
echo  [1] SYNC & PUSH: Automatically commit and push all progress to GitHub
echo  [2] PULL UPDATES: Pull latest updates and files from GitHub
echo  [3] LOG CHAT/UPDATE: Add chat notes or progress update and sync to GitHub
echo  [4] AUTO-WATCHER: Start continuous background auto-sync (every 2 mins)
echo  [5] VIEW LOG: View recent Chat and Progress Updates log
echo  [6] STATUS: Check current Git status and branch info
echo  [0] EXIT
echo.
echo =====================================================================
set /p choice="Enter your choice (0-6): "

if "%choice%"=="1" goto sync_push
if "%choice%"=="2" goto pull_repo
if "%choice%"=="3" goto log_chat
if "%choice%"=="4" goto auto_watcher
if "%choice%"=="5" goto view_log
if "%choice%"=="6" goto git_status
if "%choice%"=="0" goto end

echo Invalid option, please try again.
pause
goto menu

:sync_push
cls
echo [Running: Auto-Sync & Push to GitHub]
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0auto_sync_push.ps1"
pause
goto menu

:pull_repo
cls
echo [Running: Pull Latest Updates from GitHub]
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0pull_latest_repo.ps1"
pause
goto menu

:log_chat
cls
echo [Running: Log Chat/Update & Sync to GitHub]
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0log_chat_update.ps1"
pause
goto menu

:auto_watcher
cls
echo [Running: Continuous Auto-Sync Watcher]
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0auto_sync_push.ps1" -Watch -IntervalSeconds 120
pause
goto menu

:view_log
cls
echo =====================================================================
echo                     RECENT CHAT AND UPDATE LOG                       
echo =====================================================================
if exist "%~dp0CHAT_AND_UPDATES_LOG.md" (
    powershell -NoProfile -Command "Get-Content '%~dp0CHAT_AND_UPDATES_LOG.md' -Tail 35"
) else (
    echo Log file not found yet.
)
echo.
pause
goto menu

:git_status
cls
echo =====================================================================
echo                         GIT REPOSITORY STATUS                        
echo =====================================================================
git status
echo.
echo Remote info:
git remote -v
pause
goto menu

:end
exit /b 0
