# Mail2GPT

Windows desktop tool to convert Outlook emails into a ChatGPT-ready folder (Markdown transcript + extracted attachments). Offline-first.

## What it does
- Drag & drop a `.eml` / `.msg` file into the app, or use the watched-folder mode
- Automatically creates a timestamped folder:
  - `YYYY-MM-DD_HHMMSS_Sender_Subject`
- Inside the folder:
  - one unique `.md` file with header + instructions + cleaned body
  - all attachments extracted as separate files

## Why
Uploading raw `.eml/.msg` to a chat can lead to incomplete parsing. Mail2GPT prepares a clean and structured folder so you can drag the whole folder into ChatGPT with fewer ambiguities.

## Outlook note (important)
Some versions of the **new Outlook** (Chromium/WebView based) do not expose email contents through drag&drop (they may only provide links). In that case, use the **watched folder**:
- drag the email into a Windows folder (Outlook generates a `.eml`)
- Mail2GPT detects the new file and processes it automatically

## Requirements
- Windows 10/11
- Python 3.10+ recommended

## Install & run (dev)
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python app.py
```

## Build a single `.exe` (PyInstaller)
```powershell
python -m pip install pyinstaller
python -m PyInstaller --noconsole --onefile --windowed app.py
```
Output: `dist\app.exe`

## Usage
- **Drag a file**: drop a `.eml`/`.msg` file into the app window.
- **Watched folder**:
  - turn on “Watching ON”
  - choose a folder
  - drop emails from Outlook into that folder (Outlook will create `.eml` on disk)

## Tech
- GUI: PySide6
- `.eml` parsing: Python standard library `email`
- `.msg` parsing: `extract-msg`
- HTML → text: BeautifulSoup

## License
MIT (see LICENSE).

