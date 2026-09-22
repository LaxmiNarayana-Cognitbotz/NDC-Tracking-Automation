# Downloads F&F active reports from SharePoint, extracts employee numbers,
# runs F&F extraction, and uploads results back to SharePoint.

import asyncio
import importlib
import sys
import traceback
from datetime import datetime
from pathlib import Path
from typing import List

import httpx
import pandas as pd
from dotenv import load_dotenv

# Add project root to sys.path so we can import from scripts.*
sys.path.insert(0, str(Path(__file__).parent.parent))

from scripts.upload_to_sharepoint import SharePointService, SharePointUtils

# Dynamically import the download module since its filename contains an ampersand (&)
download_module = importlib.import_module("scripts.download_f&f_report")
download_ff_reports = download_module.download_ff_reports

# Initialize logging directory and writers
LOG_DIR = Path(__file__).parent.parent / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "process_fnf_closed_reports.log"


# Logger Writer
class LoggerWriter:
    """Redirects stdout and stderr streams to both console and a dedicated log file."""

    def __init__(self, file_path: Path):
        self.terminal = sys.stdout
        self.file_path = file_path

    def write(self, message: str):
        try:
            with open(self.file_path, "a", encoding="utf-8") as f:
                f.write(message)
        except Exception:
            pass

        if self.terminal is not None:
            try:
                self.terminal.write(message)
            except UnicodeEncodeError:
                try:
                    self.terminal.write(
                        message.encode("cp1252", errors="replace").decode("cp1252")
                    )
                except Exception:
                    pass
            except Exception:
                pass

    def flush(self):
        if self.terminal is not None:
            try:
                self.terminal.flush()
            except Exception:
                pass


sys.stdout = LoggerWriter(LOG_FILE)
sys.stderr = LoggerWriter(LOG_FILE)

load_dotenv(verbose=True)


# Application Configuration
class FNFProcessorConfig:
    """Encapsulates directories and configuration for F&F Active Reports processing."""

    def __init__(self):
        try:
            self.uploads_dir: Path = Path(__file__).parent.parent / "uploads"
            self.local_excel_dir: Path = self.uploads_dir / "FNF_Active_Reports"
            self.ff_reports_dir: Path = self.uploads_dir / "FF_Reports"

            self.local_excel_dir.mkdir(parents=True, exist_ok=True)
            self.ff_reports_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] FNFProcessorConfig Init Error: {e}")


# Automation Utilities
class FNFProcessorUtils:
    """Helper utilities for timestamps and logging."""

    @staticmethod
    def ts() -> str:
        """Returns the current formatted timestamp string."""
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# F&F Closed Reports Processor Orchestrator
class FNFClosedReportsProcessor:
    """Downloads active Excel files from SharePoint, extracts employee numbers, and triggers F&F download/upload."""

    def __init__(self, config: FNFProcessorConfig):
        self.config = config
        self.sp_service = SharePointService()

    def extract_person_numbers(self, excel_path: Path) -> List[str]:
        """Parses the Excel file, locates Person Number / Employee ID column, and returns unique IDs."""
        try:
            df = pd.read_excel(excel_path)
            person_col = None
            for col in df.columns:
                col_str = str(col).strip().lower()
                if "person number" in col_str or "employee id" in col_str:
                    person_col = col
                    break

            if not person_col:
                print(
                    f"[{FNFProcessorUtils.ts()}] [ERROR] Could not find 'Person Number' column in {excel_path.name}. Columns: {list(df.columns)}"
                )
                return []

            # Extract and sanitize string IDs
            raw_ids = (
                df[person_col]
                .dropna()
                .astype(str)
                .str.replace(r"\.0$", "", regex=True)
                .str.strip()
                .tolist()
            )
            clean_ids = [e for e in raw_ids if e and e.lower() != "nan"]

            # Deduplicate preserving order
            unique_employees = list(dict.fromkeys(clean_ids))
            print(
                f"[{FNFProcessorUtils.ts()}] Found {len(unique_employees)} unique employee(s) in {excel_path.name}: {unique_employees}"
            )
            return unique_employees
        except Exception as e:
            print(f"[{FNFProcessorUtils.ts()}] [ERROR] Error reading Excel '{excel_path.name}': {e}\n{traceback.format_exc()}")
            return []

    def create_upload_callback(
        self, client: httpx.AsyncClient, site_id: str, drive_id: str, target_base_url: str
    ):
        """Creates an async callback to upload downloaded employee documents directly to SharePoint."""

        async def on_employee_downloaded(emp_id: str, downloaded_paths: list):
            try:
                emp_target_url = f"{target_base_url}/{emp_id}"
                for file_path_str in downloaded_paths:
                    f_path = Path(file_path_str)
                    if f_path.exists():
                        try:
                            await self.sp_service.upload_file(
                                client, site_id, drive_id, emp_target_url, f_path
                            )
                            print(f"[{FNFProcessorUtils.ts()}] Uploaded successfully: {f_path.name} for employee {emp_id}")
                            try:
                                if f_path.exists():
                                    f_path.unlink()
                                    print(f"[{FNFProcessorUtils.ts()}]   [CLEANUP] Deleted local file: {f_path.name}")
                            except Exception as del_err:
                                print(f"[{FNFProcessorUtils.ts()}]   [WARN] Could not delete local file {f_path.name}: {del_err}")
                        except Exception as e:
                            print(f"[{FNFProcessorUtils.ts()}] [ERROR] Failed to upload {f_path.name} for {emp_id}: {e}")

                # Clean up employee directory if empty
                emp_dir = self.config.ff_reports_dir / emp_id
                if emp_dir.exists() and emp_dir.is_dir():
                    try:
                        remaining = list(emp_dir.iterdir())
                        if not remaining:
                            emp_dir.rmdir()
                            print(f"[{FNFProcessorUtils.ts()}]   [CLEANUP] Removed empty folder: {emp_dir.name}")
                    except Exception:
                        pass
            except Exception as e:
                print(f"[{FNFProcessorUtils.ts()}] [ERROR] Upload callback error for employee {emp_id}: {e}")

        return on_employee_downloaded

    async def process(self):
        """Main execution flow for processing F&F active reports from SharePoint."""
        print(f"[{FNFProcessorUtils.ts()}] Starting F&F Active Reports processing...")

        async with httpx.AsyncClient(timeout=60.0, verify=False) as client:
            try:
                print(f"[{FNFProcessorUtils.ts()}] Connecting to SharePoint...")
                site_id = await self.sp_service.get_site_id(client)
                drive_id, folder_path = await self.sp_service.get_drive_details(client, site_id)
                token = await self.sp_service.get_access_token()
                headers = {"Authorization": f"Bearer {token}"}

                # 1. List files in F&F_Active_Report folder on SharePoint
                fnf_folder_in_drive = f"{folder_path}/F&F_Active_Report".strip("/")
                encoded_path = SharePointUtils.encode_path_segments(fnf_folder_in_drive)

                url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}:/children"
                print(f"[{FNFProcessorUtils.ts()}] Checking SharePoint directory: {fnf_folder_in_drive}")

                response = await client.get(url, headers=headers)
                if response.status_code != 200:
                    print(
                        f"[{FNFProcessorUtils.ts()}] [ERROR] Failed to list SharePoint folder: {response.status_code} - {response.text}"
                    )
                    return

                children = response.json().get("value", [])
                excel_files = []
                for item in children:
                    if "file" in item:
                        name = item.get("name", "")
                        if name.lower().endswith(".xlsx"):
                            excel_files.append(item)

                print(f"[{FNFProcessorUtils.ts()}] Found {len(excel_files)} Excel file(s) in '{fnf_folder_in_drive}'.")

                if not excel_files:
                    return

                for file_item in excel_files:
                    file_name = file_item.get("name")
                    download_url = file_item.get("@microsoft.graph.downloadUrl")
                    if not download_url:
                        print(f"[{FNFProcessorUtils.ts()}] [ERROR] File '{file_name}' does not have a download URL.")
                        continue

                    local_file_path = self.config.local_excel_dir / file_name
                    print(f"[{FNFProcessorUtils.ts()}] Downloading '{file_name}'...")

                    try:
                        dl_response = await client.get(download_url)
                        if dl_response.status_code != 200:
                            print(
                                f"[{FNFProcessorUtils.ts()}] [ERROR] Failed to download '{file_name}' (status {dl_response.status_code})."
                            )
                            continue

                        with open(local_file_path, "wb") as f:
                            f.write(dl_response.content)

                        print(
                            f"[{FNFProcessorUtils.ts()}] Successfully downloaded '{file_name}'. Extracting Person Numbers..."
                        )

                        unique_employees = self.extract_person_numbers(local_file_path)
                        if not unique_employees:
                            print(f"[{FNFProcessorUtils.ts()}] No valid employee IDs found in {file_name}.")
                            continue

                        target_base_url = f"{folder_path}/F&F_Documents"
                        upload_cb = self.create_upload_callback(client, site_id, drive_id, target_base_url)

                        print(
                            f"[{FNFProcessorUtils.ts()}] Initiating F&F document download for {len(unique_employees)} employee(s)..."
                        )
                        await download_ff_reports(unique_employees, upload_callback=upload_cb)

                        # Clean up local Excel file
                        if local_file_path.exists():
                            local_file_path.unlink()
                            print(f"[{FNFProcessorUtils.ts()}] [CLEANUP] Deleted local Excel file: {file_name}")

                    except Exception as e:
                        print(f"[{FNFProcessorUtils.ts()}] [ERROR] Error processing '{file_name}': {e}\n{traceback.format_exc()}")

            except Exception as e:
                print(f"[{FNFProcessorUtils.ts()}] [ERROR] SharePoint operation failed: {e}\n{traceback.format_exc()}")

        print(f"[{FNFProcessorUtils.ts()}] Finished F&F Active Reports processing.")


# Main Execution Entrypoint
async def process_fnf_active_reports():
    """Main execution entrypoint for processing F&F active reports."""
    config = FNFProcessorConfig()
    processor = FNFClosedReportsProcessor(config)
    await processor.process()


if __name__ == "__main__":
    asyncio.run(process_fnf_active_reports())
