# ==============================================================================
# pull_latest_repo.ps1 - Pulls every update from GitHub repo to local folder
# Repository: https://github.com/StephSMITH-hub/theos
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -Path $RepoRoot

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "   THEOS REPOSITORY - PULL LATEST UPDATES FROM GITHUB" -ForegroundColor Cyan
Write-Host "   Repo: https://github.com/StephSMITH-hub/theos" -ForegroundColor DarkCyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""

# Ensure Git long paths is active
git config core.longpaths true

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "[$timestamp] Fetching updates from GitHub..." -ForegroundColor Yellow

# Fetch from origin
$fetchResult = git fetch origin 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "Fetch warning: $fetchResult" -ForegroundColor Yellow
}

# Check if local changes exist before pulling
$localChanges = git status --porcelain
$stashed = $false

if ($localChanges) {
    Write-Host "[$timestamp] Local uncommitted changes detected. Stashing temporarily to prevent merge conflicts..." -ForegroundColor Yellow
    git stash push -m "Auto-stashed before pull at $timestamp" 2>&1 | Out-Null
    $stashed = $true
}

# Pull latest commits
Write-Host "[$timestamp] Pulling latest changes from origin/main..." -ForegroundColor Cyan
$pullOutput = git pull origin main 2>&1
Write-Host $pullOutput -ForegroundColor Gray

# If stashed earlier, pop back
if ($stashed) {
    Write-Host "[$timestamp] Restoring your uncommitted local changes..." -ForegroundColor Yellow
    git stash pop 2>&1 | Out-Null
}

Write-Host ""
Write-Host "---------------- CURRENT REPOSITORY STATUS ----------------" -ForegroundColor DarkCyan
$latestCommit = git log -1 --pretty=format:"Latest Commit: %h%nAuthor:        %an <%ae>%nDate:          %ad%nMessage:       %s"
Write-Host $latestCommit -ForegroundColor White
Write-Host "-----------------------------------------------------------" -ForegroundColor DarkCyan

Write-Host ""
Write-Host "SUCCESS: Your local repository is up to date with GitHub!" -ForegroundColor Green
Write-Host "Repository: https://github.com/StephSMITH-hub/theos" -ForegroundColor Cyan
Write-Host ""
