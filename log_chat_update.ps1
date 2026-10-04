# ==============================================================================
# log_chat_update.ps1 - Records chat/progress updates and syncs to GitHub
# Repository: https://github.com/StephSMITH-hub/theos
# ==============================================================================

param(
    [string]$Topic = "",
    [string]$Notes = "",
    [switch]$NonInteractive
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -Path $RepoRoot

$LogFile = Join-Path $RepoRoot "CHAT_AND_UPDATES_LOG.md"

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "   THEOS - CHAT & UPDATE LOGGER -> GITHUB SYNC" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""

# Display recent updates from log
if (Test-Path $LogFile) {
    Write-Host "--- RECENT UPDATES ON THIS FOLDER ---" -ForegroundColor DarkGray
    $lines = Get-Content -Path $LogFile -Encoding UTF8
    $recent = $lines | Select-Object -Last 20
    Write-Host ($recent -join [Environment]::NewLine) -ForegroundColor Gray
    Write-Host "-------------------------------------" -ForegroundColor DarkGray
    Write-Host ""
}

# If arguments not provided and interactive, ask the user
if (-not $NonInteractive) {
    if ([string]::IsNullOrWhiteSpace($Topic)) {
        $Topic = Read-Host "Enter Topic / Summary of what was discussed or done"
    }
    if ([string]::IsNullOrWhiteSpace($Notes)) {
        $Notes = Read-Host "Enter Chat Details / Key Notes (optional, press Enter to skip)"
    }
}

# Fallbacks if still empty
if ([string]::IsNullOrWhiteSpace($Topic)) {
    $Topic = "Folder progress and transcript review"
}
if ([string]::IsNullOrWhiteSpace($Notes)) {
    $Notes = "Progress update on folder activities."
}

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Detect changed files in repo
git update-index --refresh 2>$null | Out-Null
$statusLines = git status --porcelain
$changedFiles = @()
if ($statusLines) {
    foreach ($line in $statusLines) {
        if ($line.Length -gt 3) {
            $changedFiles += $line.Substring(3).Trim('"')
        }
    }
}

$filesSummary = ""
if ($changedFiles.Count -gt 0) {
    $formattedList = @()
    $takeCount = [Math]::Min($changedFiles.Count, 10)
    for ($i = 0; $i -lt $takeCount; $i++) {
        $formattedList += "  - ``$($changedFiles[$i])``"
    }
    $filesSummary = [Environment]::NewLine + "- **Files Modified / Added:**" + [Environment]::NewLine + ($formattedList -join [Environment]::NewLine)
    if ($changedFiles.Count -gt 10) {
        $extra = $changedFiles.Count - 10
        $filesSummary += [Environment]::NewLine + "  - ... and $extra more file(s)"
    }
} else {
    $filesSummary = [Environment]::NewLine + "- **Files Modified / Added:** None (Chat / Log update only)"
}

# Format log entry
$entryLines = @(
    "",
    "### [$timestamp] - $Topic",
    "- **Date & Time:** $timestamp",
    "- **Chat / Discussion Notes:** $Notes$filesSummary",
    "- **Status:** Synced with GitHub",
    "",
    "---"
)
$newEntryText = $entryLines -join [Environment]::NewLine

# Append to CHAT_AND_UPDATES_LOG.md
Add-Content -Path $LogFile -Value $newEntryText -Encoding UTF8
Write-Host "[$timestamp] Added entry to CHAT_AND_UPDATES_LOG.md" -ForegroundColor Green

# Ensure Git longpaths
git config core.longpaths true

# Stage, commit and push to GitHub
Write-Host "[$timestamp] Staging all files and log..." -ForegroundColor Yellow
git add -A

$commitMsg = "Chat/Update: $Topic ($timestamp)"
Write-Host "[$timestamp] Committing: '$commitMsg'..." -ForegroundColor Cyan
$commitOutput = git commit -m "$commitMsg" 2>&1
Write-Host $commitOutput -ForegroundColor Gray

Write-Host "[$timestamp] Pulling remote updates with rebase..." -ForegroundColor Yellow
git fetch origin main 2>$null | Out-Null
git pull --rebase origin main 2>&1 | Out-Null

Write-Host "[$timestamp] Pushing to GitHub (https://github.com/StephSMITH-hub/theos)..." -ForegroundColor Cyan
$pushResult = git push origin main 2>&1

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "==========================================================" -ForegroundColor Green
    Write-Host " SUCCESS: Chat update and folder progress synced to GitHub!" -ForegroundColor Green
    Write-Host " Repository Link: https://github.com/StephSMITH-hub/theos" -ForegroundColor Cyan
    Write-Host " Log File: CHAT_AND_UPDATES_LOG.md" -ForegroundColor Cyan
    Write-Host "==========================================================" -ForegroundColor Green
} else {
    Write-Host "Push output: $pushResult" -ForegroundColor Red
}
Write-Host ""
