# Theos - GitHub Sync & Chat Tracker Guide

This repository is connected directly to:
👉 **[https://github.com/StephSMITH-hub/theos](https://github.com/StephSMITH-hub/theos)**

All scripts are pre-configured to handle long filenames on Windows (`core.longpaths true`) and temporary Word lock files (`~$*.docx`).

---

## What Each File Does

### 1. `auto_sync_push.bat` (or `auto_sync_push.ps1`)
- **Purpose:** Automatically syncs and pushes all progress on the folder directly to GitHub.
- **How to use:** Double-click `auto_sync_push.bat` anytime you want to save and push your latest work.
- **Continuous Watcher:** Double-click `auto_sync_watcher.bat` to keep a background process running that checks for changes every 2 minutes and auto-pushes them to GitHub.

### 2. `log_chat_update.bat` (and `CHAT_AND_UPDATES_LOG.md`)
- **Purpose:** Allows you to track the chat discussions and progress you are having on this folder per time, and automatically syncs the notes directly to GitHub.
- **How to use:** Double-click `log_chat_update.bat`.
  - It will display recent entries from `CHAT_AND_UPDATES_LOG.md`.
  - Type the topic/summary (e.g., "Updated Series 2 transcripts").
  - Optionally enter details or notes.
  - It automatically adds your update to `CHAT_AND_UPDATES_LOG.md`, stages all files, commits with your message, and pushes to GitHub.

### 3. `pull_latest_repo.bat` (or `pull_latest_repo.ps1`)
- **Purpose:** Pulls every update from the GitHub repository into your local folder.
- **How to use:** Double-click `pull_latest_repo.bat`.
  - Safely stashes any uncommitted work so merge conflicts don't occur.
  - Pulls down new commits and files from GitHub.
  - Displays the latest commit info and confirms everything is up to date.

### 4. `sync_hub.bat`
- **Master Control Center:** Double-click `sync_hub.bat` to open an interactive menu with numbers `[1]` to `[6]` allowing you to push, pull, log chats, view logs, or start the watcher in one click.

---

## Script List Summary

| File | Type | What it does |
|------|------|--------------|
| `auto_sync_push.bat` | Batch script | 1-click Push all changes to GitHub |
| `auto_sync_watcher.bat` | Batch script | Runs in background; pushes changes every 2 minutes |
| `log_chat_update.bat` | Batch script | Prompts for chat/update notes, logs to `CHAT_AND_UPDATES_LOG.md`, and pushes to GitHub |
| `pull_latest_repo.bat` | Batch script | 1-click Pull latest files/changes from GitHub |
| `sync_hub.bat` | Batch script | All-in-one interactive menu |
| `CHAT_AND_UPDATES_LOG.md` | Markdown | Running history of chats and progress updates |
| `.gitignore` | Git config | Ignores temporary Word locks (`~$*.docx`) and OS junk |
