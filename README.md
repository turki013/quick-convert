# Quick Convert

Fast, lightweight, **100% local & offline** file conversion and PDF utility engine. Built with a decoupled `core/` engine powering:
- **Windows 11/10 Right-Click Context Menu** (`cli.py` + `shell_extension/` Sparse Package, `pythonw.exe` + Windows toast notifications).
- **24/7 Cloud-Ready Telegram Bot** (`bot.py` using `aiogram` v3 + `aiohttp` health-check/webhook server).

## Supported Features
- **Documents:** `.docx` ↔ `.md`, `.docx` → `.pdf` (MS Word COM on Windows / LibreOffice headless on Linux), `.md` → `.pdf`, `.pdf` → `.docx` / `.md` / `.txt`
- **PDF Utilities:** Compress (`low`, `medium`, `high`), Merge (multi-select, folder, or `/merge` in Telegram), Split (page ranges or burst), Extract Embedded Images
- **Images:** `.png` ↔ `.jpg` ↔ `.webp` ↔ `.bmp`, Compress Image (`low`, `medium`, `high`), Images ↔ PDF (`/img2pdf` in Telegram)
- **Tabular Data:** `.xlsx` ↔ `.csv` ↔ `.json` with `utf-8-sig` CSV encoding for full Arabic text compatibility

## Local Windows Setup

```powershell
cd C:\Users\HUAWEI\quick_convert
py -3.11 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\pytest.exe -v
```

### Install Windows 11/10 Context Menu
Enable **Developer Mode** in Windows 11 (*Settings → System → For developers*) so it appears directly in the primary modern right-click menu, then run:

```powershell
.\.venv\Scripts\python.exe scripts\install_context_menu.py
```

To remove it:

```powershell
.\.venv\Scripts\python.exe scripts\uninstall_context_menu.py
```

## 24/7 Cloud Deployment (Koyeb / Render)

The included `Dockerfile` uses `python:3.11-slim` with `libreoffice-writer` for `.docx` → `.pdf` on Linux, and `bot.py` exposes `/health` on `PORT` (default `8000`) for free-tier health checks / UptimeRobot.

1. Push this repository to GitHub (`.gitignore` excludes `.venv` and temp files).
2. Create a new **Web Service** on **Koyeb** or **Render** using **Docker** (`Dockerfile`).
3. Set environment variables:
   - `BOT_TOKEN`: Your Telegram Bot token from `@BotFather`
   - `ALLOWED_USER_IDS` *(optional)*: Comma-separated Telegram user IDs for private access (e.g. `123456789`)
   - `PORT` *(optional)*: `8000` (Render sets `PORT` automatically)
   - `WEBHOOK_URL` *(optional)*: Your public `https://your-app.koyeb.app` URL to use Webhook mode instead of Polling.
4. Set Health Check path to `/health`.
