# Enterprise Automation: Oracle Fusion & OpenText Reporting Pipeline

![Python 3.12+](https://img.shields.io/badge/Python-3.12%2B-blue.svg)
![FastAPI](https://img.shields.io/badge/FastAPI-REST%20API-009688.svg)
![Playwright](https://img.shields.io/badge/Automation-Playwright-2b3137.svg)
![SharePoint Sync](https://img.shields.io/badge/Sync-MS%20Graph%20API-0078d4.svg)
![Windows Task Scheduler](https://img.shields.io/badge/Scheduler-Windows%20Native-00a4ef.svg)

An enterprise automation suite and asynchronous REST API to extract, post-process, filter, and synchronize business reports and employee HR operations documents across **Oracle Fusion ERP**, **OpenText Content Server**, and **Microsoft SharePoint Online**.

---

## Features at a Glance

| Module | Pipeline | Description |
|---|---|---|
| **FastAPI REST API** | Server & OpenAPI | REST endpoints with interactive Swagger UI (`/docs`) for manual triggers and status tracking. |
| **Oracle Fusion** | NDC Reports | Bypasses SSO, selects Business Unit parameters, downloads, filters, and uploads NDC reports. |
| **Oracle Fusion** | ELCM Reports | Downloads New Joinees report, filters BU, initializes/merges `Email Sent` RPA tracking, and syncs to SharePoint. |
| **OpenText** | F&F Exit Docs | Scans active SharePoint lists or employee IDs to scrape PDF/Excel F&F exit settlement sheets. |
| **OpenText** | Pre-Joining Docs | Multi-tab parallel extraction of Offer, BGV, and Appointment letters with Master Excel generation. |
| **SharePoint Online**| MS Graph API | Automated file synchronization and cleanup via Microsoft Graph API and MSAL authentication. |
| **Task Scheduler** | Windows Native | Headless execution (`pythonw.exe`) with automatic catch-up on system reboot. |

Detailed architecture diagrams and workflows are located in: **[docs/architecture.md](docs/architecture.md)**.

---

## 1. Prerequisites

Before running the project, verify that the host machine has:
1. **Python 3.12+** installed.
2. **Google Chrome** installed (used to bypass SSO and handle corporate authentication).
3. **[uv](https://github.com/astral-sh/uv)** (recommended Python package manager) or standard `pip`.

---

## 2. Installation & Setup

### Step 1: Configure Environment Variables
Copy `.env.sample` to a new `.env` file in the project root:
```powershell
copy .env.sample .env
```

Open `.env` and configure your credentials:
* **Oracle Fusion**: `ORACLE_URL`, `ORACLE_EMAIL`, `ORACLE_PASSWORD`
* **OpenText Content Server**: `CONTENT_SERVER_URL`
* **Microsoft Azure Graph API**: `SHAREPOINT_TENANT_ID`, `SHAREPOINT_CLIENT_ID`, `SHAREPOINT_CLIENT_SECRET`
* **SharePoint Destinations**: `SHAREPOINT_SITE_URL`, `SHAREPOINT_TARGET_FOLDER`, `SHAREPOINT_TARGET_FOLDER_ELCM`
* **Runtime Flags**: `HEADLESS=true` (or `false` to view browser), `START_MAXIMIZED=true`

### Step 2: Install Dependencies
```powershell
# Sync Python environment and install all packages
uv sync

# Install Playwright browser binaries
uv run playwright install chromium
```

*(Alternative using standard pip)*:
```powershell
python -m venv .venv
.venv\Scripts\activate
pip install .
playwright install chromium
```

---

## 3. How to Run

You can run the automation pipelines either via the **FastAPI REST API Server** or by running individual **Python scripts directly**.

### Option A: Run the FastAPI REST API Server (Recommended)

Start the server using `uv`:
```powershell
uv run .\main.py
```

* **API Server**: `http://localhost:8000`
* **Interactive Swagger UI**: [http://localhost:8000/docs](http://localhost:8000/docs)
* **ReDoc Documentation**: [http://localhost:8000/redoc](http://localhost:8000/redoc)

#### Key Endpoints

| Method | Endpoint | Description |
|---|---|---|
| `POST` | `/api/v1/automation/ndc/download` | Triggers the NDC report download, BU filtering, and SharePoint push. |
| `POST` | `/api/v1/automation/elcm/download` | Triggers the ELCM New Joinees report download, BU filtering, and SharePoint sync. |
| `POST` | `/api/v1/automation/fnf/process-closed` | Reads `.xlsx` from SharePoint `F&F_Active_Report`, downloads F&F docs, and uploads. |
| `POST` | `/api/v1/automation/fnf/download-by-ids` | Downloads F&F documents for specific employee numbers provided in JSON body. |
| `POST` | `/api/v1/automation/prejoining/download` | Downloads Pre-Joining letters (Offer, BGV, Appointment) for provided employee IDs. |
| `POST` | `/api/v1/automation/sharepoint/upload` | Uploads local files to SharePoint (`target_pipeline`: `all`, `ndc`, `ff`, `prejoining`). |
| `GET` | `/api/v1/automation/jobs` | Lists all background automation jobs and their current status. |
| `GET` | `/api/v1/automation/jobs/{job_id}` | Checks status, execution time, and errors for a specific job ID. |
| `GET` | `/health` | Application health check endpoint. |

---

### Option B: Standalone Script Execution (Terminal / CLI)

Each script can also be executed independently from the command line:

```powershell
# 1. NDC Report Pipeline (Oracle scraper -> Adani Green BU filter -> SharePoint sync)
uv run app/scripts/automation_scripts/download_ndc_report.py

# 2. ELCM New Joinees Pipeline (Oracle scraper -> Adani Green filter -> merge Email Sent tracking -> SharePoint sync)
uv run app/scripts/automation_scripts/elcm_report_downloader.py

# 3. F&F Active Document Pipeline (Reads active SharePoint list -> OpenText scraper -> SharePoint upload)
uv run app/scripts/processing_scripts/process_fnf_closed_reports.py

# 4. F&F Document Downloader (Manual prompt for specific employee IDs)
uv run "app/scripts/automation_scripts/download_f&f_report.py"

# 5. Pre-Joining Document Downloader (Manual prompt for employee IDs -> Master Excel -> SharePoint)
uv run app/scripts/automation_scripts/download_prejoining_report.py

# 6. SharePoint Uploader (Direct synchronization of local uploads/ to SharePoint)
uv run app/scripts/processing_scripts/upload_to_sharepoint.py
```

---

## 4. Background Task Scheduling (Windows Task Scheduler)

The repository includes a PowerShell configuration script to register background tasks with Windows Task Scheduler. Tasks run silently using `pythonw.exe` without interrupting active desktop sessions, with automated failover and boot catch-up enabled.

### Default Schedule Timings (Configured in `app/core/config.py`)
* **NDC Pipeline**: `10:00 AM`, `1:00 PM`, `3:00 PM`, `5:00 PM`, `7:00 PM`
* **F&F Pipeline**: `11:00 AM`, `4:00 PM`
* **ELCM Pipeline**: `10:30 AM`

### Register Scheduled Tasks
Run from PowerShell (no Administrator privileges required):
```powershell
powershell -ExecutionPolicy Bypass -File app\scheduler\setup_scheduler.ps1
```

### Check Active Scheduled Tasks
```powershell
Get-ScheduledTask -TaskName "*_Pipeline_*" | Select-Object TaskName, State
```

### Pause / Resume Tasks
```powershell
# Pause all automation tasks
Get-ScheduledTask -TaskName "*_Pipeline_*" | Disable-ScheduledTask

# Resume all automation tasks
Get-ScheduledTask -TaskName "*_Pipeline_*" | Enable-ScheduledTask
```

### Uninstall Scheduled Tasks
```powershell
powershell -ExecutionPolicy Bypass -File app\scheduler\setup_scheduler.ps1 -Uninstall
```

---

## 5. Directory Structure & Logs

```
NDC-Tracking-Automation/
├── app/
│   ├── api/
│   │   └── routes.py                         # FastAPI endpoint definitions
│   ├── core/
│   │   ├── config.py                         # Central paths, directories, and schedule times
│   │   ├── logger.py                         # Dual-mode logging & console output filtering
│   │   └── process.py                        # Windows subprocess window suppression
│   ├── dto/
│   │   └── schemas.py                        # Pydantic models & request/response schemas
│   ├── helpers/
│   │   └── auth.py                           # Microsoft SSO & browser session handler
│   ├── scheduler/
│   │   └── setup_scheduler.ps1               # Automated Windows Task Scheduler registration
│   ├── scripts/
│   │   ├── automation_scripts/
│   │   │   ├── download_ndc_report.py        # Oracle Fusion NDC report automation
│   │   │   ├── elcm_report_downloader.py     # Oracle Fusion ELCM report automation
│   │   │   ├── download_f&f_report.py        # OpenText F&F exit document scraper
│   │   │   └── download_prejoining_report.py # OpenText Pre-Joining letters scraper
│   │   └── processing_scripts/
│   │       ├── process_fnf_closed_reports.py # SharePoint batch processor for F&F
│   │       └── upload_to_sharepoint.py       # Microsoft Graph API uploader
│   └── services/
│       └── automation_service.py             # Background task management & job tracker
├── chrome_automation_profile/                # Persistent Chrome user data and SSO cookies
├── errors/                                   # Automation failure screenshots (generated on error)
├── logs/
│   ├── server_logs.log                       # Consolidated FastAPI server and background job log
│   ├── download_ndc_report.log               # Dedicated NDC standalone script log
│   ├── download_elcm_report.log              # Dedicated ELCM standalone script log
│   ├── process_fnf_closed_reports.log        # Dedicated F&F batch processing log
│   └── download_ff_report.log                # Dedicated F&F download script log
├── uploads/                                  # Local storage for reports before SharePoint sync
├── docs/
│   └── architecture.md                       # Architecture design and data flow diagrams
├── main.py                                   # FastAPI entry point & Uvicorn server launcher
├── pyproject.toml                            # Project metadata and dependencies
└── README.md
```
