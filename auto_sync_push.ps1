# ==============================================================================
# auto_sync_push.ps1 - Automatically syncs and pushes all progress to GitHub
# Repository: https://github.com/StephSMITH-hub/theos
# ==============================================================================

param(
    [string]$CommitMessage = "",
    [switch]$Watch,
    [int]$IntervalSeconds = 120
)

# Ensure UTF-8 output encoding
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -Path $RepoRoot

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "   THEOS REPOSITORY - GITHUB AUTO-SYNC & PUSH" -ForegroundColor Cyan
Write-Host "   Repo: https://github.com/StephSMITH-hub/theos" -ForegroundColor DarkCyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""

# 1. Enable long paths to prevent Windows MAX_PATH errors with deep folders
git config core.longpaths true

function Sync-Progress {
    param([string]$CustomMsg = "")
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] Checking repository status..." -ForegroundColor Yellow
    
    # Refresh index and inspect git status
    git update-index --refresh 2>$null | Out-Null
    $status = git status --porcelain
    
    if (-not $status) {
        Write-Host "[$timestamp] No local changes detected in working tree." -ForegroundColor Green
        
        # Check if local branch is ahead of remote
        $aheadCount = (git rev-list --count origin/main..HEAD 2>$null)
        if ($aheadCount -and [int]$aheadCount -gt 0) {
            Write-Host "[$timestamp] Local branch is ahead by $aheadCount commit(s). Pushing to GitHub..." -ForegroundColor Yellow
            $pushResult = git push origin main 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Host "[$timestamp] Successfully pushed to https://github.com/StephSMITH-hub/theos" -ForegroundColor Green
            } else {
                Write-Host "[$timestamp] Push failed: $pushResult" -ForegroundColor Red
            }
        } else {
            Write-Host "[$timestamp] Local workspace is fully in sync with GitHub." -ForegroundColor Green
        }
        return
    }

    Write-Host "[$timestamp] Local changes detected! Staging all files..." -ForegroundColor Yellow
    git add -A

    # Determine commit message
    if ([string]::IsNullOrWhiteSpace($CustomMsg)) {
        $msg = "Auto-sync progress: $timestamp"
    } else {
        $msg = "$CustomMsg (Synced: $timestamp)"
    }

    Write-Host "[$timestamp] Committing: '$msg'" -ForegroundColor Cyan
    $commitResult = git commit -m "$msg" 2>&1
    Write-Host $commitResult -ForegroundColor Gray

    # Pull any new remote commits with rebase to ensure clean push
    Write-Host "[$timestamp] Fetching latest remote commits..." -ForegroundColor Yellow
    git fetch origin main 2>$null | Out-Null
    
    $rebaseResult = git pull --rebase origin main 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[$timestamp] Notice during pull: $rebaseResult" -ForegroundColor Yellow
    }

    # Push to GitHub
    Write-Host "[$timestamp] Pushing changes to GitHub (origin/main)..." -ForegroundColor Cyan
    $pushResult = git push origin main 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "[$timestamp] SUCCESS: All progress has been synced to GitHub directly!" -ForegroundColor Green
        Write-Host "View at: https://github.com/StephSMITH-hub/theos" -ForegroundColor Cyan
    } else {
        Write-Host "[$timestamp] PUSH ERROR: $pushResult" -ForegroundColor Red
    }
}

if ($Watch) {
    Write-Host "Starting continuous auto-sync watcher..." -ForegroundColor Magenta
    Write-Host "Watching '$RepoRoot' every $IntervalSeconds seconds for changes. Press Ctrl+C to stop." -ForegroundColor Gray
    Write-Host ""
    while ($true) {
        try {
            Sync-Progress -CustomMsg "Continuous watcher progress"
        } catch {
            Write-Host "Error during sync: $_" -ForegroundColor Red
        }
        Start-Sleep -Seconds $IntervalSeconds
    }
} else {
    Sync-Progress -CustomMsg $CommitMessage
}

Write-Host ""
Write-Host "Done!" -ForegroundColor Green
