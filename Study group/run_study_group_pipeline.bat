@echo off
setlocal
chcp 65001 >nul

echo ==========================================================
echo    SAINTS COMMUNITY CHURCH - STUDY GROUP ENGINE RUNNER
echo    Location: Study group
echo ==========================================================
echo.

cd /d "%~dp0"

echo [*] Executing study_group_engine.py...
python study_group_engine.py
if errorlevel 1 (
    echo [!] Error executing study_group_engine.py
    pause
    exit /b 1
)

echo.
echo [*] Triggering repository auto-sync to GitHub...
cd ..
if exist "auto_sync_push.bat" (
    call auto_sync_push.bat
) else (
    git add -A
    git commit -m "Auto-sync study group pipeline progress"
    git push origin main
)

echo.
echo ==========================================================
echo    ALL STUDY GROUP PIPELINES COMPLETED & SYNCED!
echo ==========================================================
echo.
pause
