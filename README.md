# Enterprise Automation: Oracle Fusion & OpenText Reporting Pipeline

![Python 3.12+](https://img.shields.io/badge/Python-3.12%2B-blue.svg)
![Playwright](https://img.shields.io/badge/Automation-Playwright-2b3137.svg)
![SharePoint Sync](https://img.shields.io/badge/Sync-MS%20Graph%20API-0078d4.svg)
![Windows Task Scheduler](https://img.shields.io/badge/Scheduler-Windows%20Native-00a4ef.svg)

This repository contains automated pipelines to download, filter, and sync business reports and employee exit documents. It automates two primary workflows: **NDC Process Request Status Reports** (Oracle Fusion ERP) and **F&F Exit Documents** (OpenText Content Server), uploading all results to Microsoft SharePoint Online.

## Features at a Glance
| Module | Capability | Description |
|---|---|---|
| **Oracle Fusion** | Report Extraction | Bypasses SSO, searches Business Units, and downloads Excel reports. |
| **OpenText** | F&F Exit Sync | Scrapes OpenText directories for PDF F&F calculation sheets. |
| **SharePoint** | Graph API Uploads | Syncs files reliably via MSAL Authentication and clears local caches. |
| **Authentication** | RSA MFA Handling | Caches browser sessions and handles mobile push notifications automatically. |

Detailed system block diagrams, sequence diagrams, and multi-stage authentication workflows are located in the unified documentation:
**[docs/architecture.md](file:///d:/Projects/NDC-Tracking-Automation/docs/architecture.md)**

---

## 1. Prerequisites

Before setting up the project, verify that the host machine has:
1. **Python 3.12+** installed.
2. **Google Chrome** installed (the scripts use the host's native Chrome installation to bypass SSO).
3. **[uv](https://github.com/astral-sh/uv)** (recommended Python package manager) or standard `pip`.

---

## 2. Installation & Setup

### Step 1: Configure Credentials
Copy `.env.sample` to a new file named `.env` in the root directory:
```powershell
copy .env.sample .env
```

Open `.env` in a text editor and fill in the required fields:
* **Oracle Fusion URL** and corporate credentials (`ORACLE_EMAIL` / `ORACLE_PASSWORD`).
* **OpenText Content Server** URL link.
* **Microsoft Azure App Registration** credentials (`SHAREPOINT_TENANT_ID`, `SHAREPOINT_CLIENT_ID`, `SHAREPOINT_CLIENT_SECRET`).
* **SharePoint Destination Folder** path and site URL.
* Set `HEADLESS=true` for background VM execution or `false` to open the browser window and watch.
* Set `START_MAXIMIZED=true` (default) to launch the browser in full screen when `HEADLESS=false`.

### Step 2: Install Project Dependencies
Run the following commands in your terminal to initialize the environment:

```powershell
# Sync Python environment and install package dependencies
uv sync

# Install Playwright browser dependencies
uv run playwright install chromium
```

*(Optional)* If you are using standard `pip` and Python virtual environments:
```powershell
python -m venv .venv
.venv\Scripts\activate
pip install .
playwright install chromium
```

---

## 3. How to Run the Pipelines

The project can be run manually from a terminal or triggered automatically via background tasks.

### 1. Manual Execution (Terminal)

```powershell
# Run the NDC Report Pipeline (Oracle scraper -> local BU filter -> SharePoint upload)
uv run scheduler/run_pipeline.py

# Run the F&F Active Document Pipeline (SharePoint list check -> OpenText scraper -> real-time upload)
uv run scripts/process_fnf_closed_reports.py

# Run F&F Document Downloader manually for specific employee IDs
uv run scripts/download_f&f_report.py

# Run Pre-Joining & Talent Acquisition Document Downloader manually
uv run scripts/download_prejoining_report.py

# Upload ONLY NDC reports to SharePoint (skipping other folders)
uv run scripts/upload_to_sharepoint.py --type ndc

# Upload ONLY F&F documents to SharePoint
uv run scripts/upload_to_sharepoint.py --type ff

# Upload ONLY Pre-Joining documents to SharePoint
uv run scripts/upload_to_sharepoint.py --type prejoining

# Upload a specific file directly to SharePoint
uv run scripts/upload_to_sharepoint.py --file "C:\path\to\your\file.pdf"
```

---

## 4. Background Task Scheduling

The system uses Windows Task Scheduler daily to run in windowless mode (`pythonw.exe`).

### Setup Schedules (Run Once)
To register the background tasks on the VM, open **PowerShell** (no Administrator rights required) and run:
```powershell
powershell -ExecutionPolicy Bypass -File scheduler\setup_scheduler.ps1
```

### Verify Active Schedules
Confirm all daily tasks are registered and ready:
```powershell
Get-ScheduledTask -TaskName "*_Pipeline_*" | Select-Object TaskName, State
```

### Pause Background Tasks
To temporarily disable background triggers (e.g. during target system maintenance):
```powershell
Get-ScheduledTask -TaskName "*_Pipeline_*" | Disable-ScheduledTask
```

### Resume Background Tasks
To re-enable background triggers:
```powershell
Get-ScheduledTask -TaskName "*_Pipeline_*" | Enable-ScheduledTask
```

### Remove Background Tasks
To permanently delete all registered tasks from the Windows host:
```powershell
powershell -ExecutionPolicy Bypass -File scheduler\setup_scheduler.ps1 -Uninstall
```

---

## 5. Directory Layout & Logs

* **`scheduler/`**: Orchestration scripts and background Task Scheduler installations.
* **`scripts/`**: Automation scripts for browser scraping, filtering, and MS Graph API syncing.
* **`uploads/`**: Local cache directories where reports and PDFs are temporarily stored.
* **`logs/`**: Operations log directory. Inspect `logs/pipeline.log` and `logs/process_fnf_closed_reports.log` to audit recent runs.
* **`chrome_automation_profile/`**: User profile directory caching session cookies to bypass Microsoft SSO.
