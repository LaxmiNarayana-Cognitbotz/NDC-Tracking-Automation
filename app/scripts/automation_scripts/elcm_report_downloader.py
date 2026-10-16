import sys
from pathlib import Path

# Add project root to sys.path so the script can be run directly or via VS Code Run
ROOT_DIR = Path(__file__).resolve().parents[3]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import asyncio
import io
import os
import re
import traceback
from datetime import datetime
from typing import List, Optional

import httpx
import pandas as pd
from dotenv import load_dotenv
from playwright.async_api import Frame, Page, TimeoutError as PlaywrightTimeoutError, async_playwright

import app.scripts.processing_scripts.upload_to_sharepoint as sp_uploader
from app.helpers.auth import MicrosoftSSOHandler
from app.core.config import AUTOMATION_PROFILE_DIR, ELCM_ERRORS_DIR, ELCM_REPORTS_DIR
from app.core.logger import setup_logger
from app.core.process import suppress_subprocess_window

# ==== Suppress Chrome subprocess window on Windows ====
suppress_subprocess_window()

# ==== Logger Setup ====
LOG_FILE = setup_logger("download_elcm_report.log")

load_dotenv(verbose=True)


# Application Configuration
class AppConfig:
    """Manages environment variables, credentials, paths, and runtime settings."""

    def __init__(self):
        try:
            # Oracle Fusion credentials and URL
            self.oracle_url: str = (os.getenv("ORACLE_URL") or "").strip()
            self.oracle_email: str = (os.getenv("ORACLE_EMAIL") or "").strip()
            self.oracle_password: str = (os.getenv("ORACLE_PASSWORD") or "").strip()

            # SharePoint target folder for ELCM report
            self.sharepoint_target_folder_elcm: str = (
                os.getenv("SHAREPOINT_TARGET_FOLDER_ELCM") or ""
            ).strip()

            # Output and storage directories
            self.download_dir: Path = ELCM_REPORTS_DIR
            self.automation_profile_dir: Path = AUTOMATION_PROFILE_DIR
            self.error_dir: Path = ELCM_ERRORS_DIR

            # Browser configuration
            self.headless: bool = self._parse_bool(os.getenv("HEADLESS"), default=True)
            self.start_maximized: bool = self._parse_bool(os.getenv("START_MAXIMIZED"), default=True)

            # Report parameters
            self.business_unit_search: str = "Adani Green"
            self.assignment_status_value: str = "ACTIVE"
            self.download_timeout_ms: int = 1_200_000  # 20 minutes for large ELCM Excel export

            # Required base columns to keep from Oracle ELCM report
            self.required_columns: List[str] = [
                "Employee Code",
                "Name of Employee",
                "BusinessUnitasperOraclestring",
                "Location Name",
                "Work Email",
            ]

            # Target locations for post-download filtering
            self.target_locations: List[str] = [
                "CT-1 Inspire Business Park Ahmedabad",
                "ACH Building Ahmedabad",
                "Ahmedabad",
            ]

            # Default tracking column for RPA team email status
            self.default_tracking_column: str = "Email Sent"

            # Ensure download directory exists
            self.download_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Configuration Initialization Error: {e}")

    @staticmethod
    def _parse_bool(val: Optional[str], default: bool = True) -> bool:
        """Safely parses string values into boolean flags."""
        try:
            if val is None:
                return default
            return str(val).strip().lower() in ("true", "1", "yes", "t", "y")
        except Exception:
            return default


# Automation Utilities
class AutomationUtils:
    """Utility helper functions for timing, screenshots, locks, and DOM actions."""

    @staticmethod
    def ts() -> str:
        """Returns the current formatted timestamp string."""
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    @staticmethod
    def report_save_stem() -> str:
        """Generates a timestamped filename stem for the downloaded ELCM report."""
        try:
            now = datetime.now()
            return (
                f"ELCM_Report_{now.day}_"
                f"{now.strftime('%B')}_{now.year}_{now.strftime('%I.%M%p')}"
            )
        except Exception:
            return f"ELCM_Report_{int(datetime.now().timestamp())}"

    @staticmethod
    def normalize_emp_code(val) -> str:
        """Normalizes Employee Code strings (e.g. '10002440.0' -> '10002440') for reliable matching."""
        if pd.isna(val):
            return ""
        s = str(val).strip()
        s = re.sub(r"\.0$", "", s)
        return s

    @staticmethod
    def cleanup_profile_locks(profile_dir: Path):
        """Cleans up Chromium profile lock files to avoid browser startup crashes."""
        try:
            if not profile_dir.exists():
                return
            lock_files = ["SingletonLock", "SingletonSocket", "SingletonCookie", "lockfile"]
            for lock_name in lock_files:
                lock_path = profile_dir / lock_name
                try:
                    if lock_path.exists():
                        lock_path.unlink()
                except OSError:
                    pass
        except Exception as e:
            print(f"[{AutomationUtils.ts()}] [CLEANUP] Lock cleanup warning: {e}")

    @staticmethod
    async def save_debug_screenshot(page: Page, error_dir: Path, label: str) -> str:
        """Captures a full-page diagnostic screenshot upon errors formatted with source file, date and time."""
        try:
            error_dir.mkdir(parents=True, exist_ok=True)
            now_str = datetime.now().strftime("%d_%b_%Y_(%I.%M%p)")
            path = error_dir / f"elcm_{label}_{now_str}.png"
            await page.screenshot(path=str(path), full_page=True)
            print(f"[{AutomationUtils.ts()}]   Screenshot saved: {path.name}")
            return str(path)
        except Exception as e:
            print(f"[{AutomationUtils.ts()}]   Could not save screenshot: {e}")
            return ""

    @staticmethod
    async def safe_click(locator, description: str, timeout: int = 15_000):
        """Scrolls into view, waits for visibility, and clicks element safely."""
        try:
            await locator.wait_for(state="visible", timeout=timeout)
            await locator.scroll_into_view_if_needed()
            await locator.click(timeout=timeout)
        except PlaywrightTimeoutError:
            raise RuntimeError(f"Could not click '{description}' within {timeout}ms")
        except Exception as e:
            raise RuntimeError(f"Error clicking '{description}': {e}")

    @staticmethod
    async def click_with_fallback(
        page_or_frame, selectors: List[str], description: str, timeout_each: int = 10_000
    ):
        """Attempts to click a list of fallback selectors until one succeeds."""
        errors = []
        for sel in selectors:
            try:
                loc = page_or_frame.locator(sel).first
                await loc.wait_for(state="visible", timeout=timeout_each)
                await loc.scroll_into_view_if_needed()
                await loc.click()
                return
            except Exception as e:
                errors.append(f"    {sel!r}: {e}")
        raise RuntimeError(
            f"No working selector for '{description}'.\nTried:\n" + "\n".join(errors)
        )

    @staticmethod
    async def wait_for_loading(frame: Frame, timeout_ms: int = 90_000):
        """Waits for Oracle ADF loading modal masks to disappear."""
        try:
            mask = frame.locator("div.modalMask")
            if await mask.count() == 0:
                return
            if not await mask.first.is_visible(timeout=2000):
                return
            await mask.first.wait_for(state="hidden", timeout=timeout_ms)
        except Exception:
            pass


# ==== Report Post-Filter & RPA Tracking Merger ====
class ReportPostFilter:
    """
    Handles parsing, Business Unit and Location filtering, keeping required columns,
    and merging previous RPA email-sent tracking data by Employee Code.
    """

    REQUIRED_COLUMNS = [
        "Employee Code",
        "Name of Employee",
        "BusinessUnitasperOraclestring",
        "Location Name",
        "Work Email",
    ]

    TARGET_LOCATIONS = [
        "CT-1 Inspire Business Park Ahmedabad",
        "ACH Building Ahmedabad",
        "Ahmedabad",
    ]

    @staticmethod
    def _save_dataframe_to_excel(df: pd.DataFrame, file_path: Path):
        """Writes a DataFrame cleanly to .xls or .xlsx."""
        if file_path.suffix.lower() == ".xls":
            df.to_excel(file_path, index=False)
        else:
            df.to_excel(file_path, index=False, engine="openpyxl")

    @staticmethod
    def process_elcm_excel(
        file_path: Path,
        required_columns: Optional[List[str]] = None,
        target_locations: Optional[List[str]] = None,
        default_tracking_col: str = "Email Sent",
    ) -> Optional[pd.DataFrame]:
        """
        Processes the downloaded ELCM Excel report:
          1. Locates the data table and header row
          2. Filters rows so 'BusinessUnitasperOraclestring' starts with 'Adani Green'
          3. Filters rows so 'Location Name' matches target locations (Ahmedabad offices)
          4. Retains required columns + initializes the default tracking column ('Email Sent')
          5. Saves the cleaned Excel file to disk and returns the DataFrame.
        """
        if required_columns is None:
            required_columns = ReportPostFilter.REQUIRED_COLUMNS
        if target_locations is None:
            target_locations = ReportPostFilter.TARGET_LOCATIONS

        try:
            print(f"[{AutomationUtils.ts()}] Processing ELCM Excel report: {file_path.name}...")

            is_html = False
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                head = f.read(1000).lower()
                if "<html" in head or "<table" in head:
                    is_html = True

            if is_html:
                dfs = pd.read_html(str(file_path))
                raw_df = max(dfs, key=len)
            else:
                raw_df = pd.read_excel(file_path, header=None)

            # Find the row index containing the column headers (e.g. 'Employee Code' & 'Work Email')
            header_row_idx = None
            for r_idx in range(min(30, len(raw_df))):
                row_vals = [str(x).strip() for x in raw_df.iloc[r_idx].tolist()]
                if "Employee Code" in row_vals and "Work Email" in row_vals:
                    header_row_idx = r_idx
                    break

            if header_row_idx is None:
                raise RuntimeError(
                    f"Could not find header row containing {required_columns} in {file_path.name}"
                )

            headers = [str(x).strip() for x in raw_df.iloc[header_row_idx].tolist()]
            data_df = raw_df.iloc[header_row_idx + 1 :].copy()
            data_df.columns = headers

            total_initial_cols = len(data_df.columns)
            total_initial_rows = len(data_df)

            col_map = {str(c).strip().lower(): str(c).strip() for c in data_df.columns}
            if "location name" in col_map and "Location Name" not in data_df.columns:
                data_df.rename(columns={col_map["location name"]: "Location Name"}, inplace=True)
            elif "location" in col_map and "Location Name" not in data_df.columns:
                data_df.rename(columns={col_map["location"]: "Location Name"}, inplace=True)

            missing_cols = [c for c in required_columns if c not in data_df.columns]
            if missing_cols:
                raise RuntimeError(
                    f"Missing required columns in Excel: {missing_cols}. Found columns: {list(data_df.columns)}"
                )

            # Filter rows to Adani Green if any non-Adani Green rows exist
            if "BusinessUnitasperOraclestring" in data_df.columns:
                data_df = data_df[
                    data_df["BusinessUnitasperOraclestring"]
                    .astype(str)
                    .str.strip()
                    .str.lower()
                    .str.startswith("adani green")
                ]

            # Filter rows by Location Name
            if "Location Name" in data_df.columns and target_locations:
                norm_loc = data_df["Location Name"].astype(str).str.strip().str.lower()
                target_locs_lower = [loc.strip().lower() for loc in target_locations if loc.strip()]
                loc_mask = norm_loc.isin(target_locs_lower) | norm_loc.str.contains("ahmedabad", na=False)
                before_loc_count = len(data_df)
                data_df = data_df[loc_mask]
                print(
                    f"[{AutomationUtils.ts()}] Location filter applied: "
                    f"retained {len(data_df)}/{before_loc_count} records matching Ahmedabad locations."
                )

            # Preserve any existing tracking columns if already present in file, plus the required columns
            extra_existing_cols = [
                c for c in data_df.columns if c not in required_columns and c == default_tracking_col
            ]
            cleaned_df = data_df[required_columns + extra_existing_cols].copy()
            cleaned_df["Employee Code"] = cleaned_df["Employee Code"].apply(
                AutomationUtils.normalize_emp_code
            )

            if default_tracking_col and default_tracking_col not in cleaned_df.columns:
                cleaned_df[default_tracking_col] = ""

            cleaned_df.reset_index(drop=True, inplace=True)
            ReportPostFilter._save_dataframe_to_excel(cleaned_df, file_path)

            print(
                f"[{AutomationUtils.ts()}] Excel processed successfully: "
                f"Kept {len(cleaned_df)}/{total_initial_rows} rows and "
                f"base columns ({', '.join(required_columns)}) + '{default_tracking_col}' "
                f"(removed {total_initial_cols - len(required_columns)} unnecessary columns)."
            )
            return cleaned_df

        except Exception as e:
            print(f"[{AutomationUtils.ts()}] Error processing ELCM Excel file: {e}\n{traceback.format_exc()}")
            return None

    @staticmethod
    def merge_rpa_tracking_from_previous_df(
        new_df: pd.DataFrame,
        old_df: pd.DataFrame,
        file_path: Path,
        required_columns: List[str],
        default_tracking_col: str = "Email Sent",
    ) -> pd.DataFrame:
        """
        Merges RPA email tracking status from the previous SharePoint Excel (old_df)
        into the newly downloaded ELCM report (new_df) matched by 'Employee Code'.
        Preserves both the default tracking column ('Email Sent') and any additional
        tracking columns added/updated by the RPA team on SharePoint.
        """
        try:
            if old_df is None or old_df.empty or "Employee Code" not in old_df.columns:
                print(
                    f"[{AutomationUtils.ts()}]   No previous Employee Code records to merge; keeping fresh report."
                )
                return new_df

            old_df = old_df.copy()
            old_df.columns = [str(c).strip() for c in old_df.columns]
            old_df["_norm_emp_code"] = old_df["Employee Code"].apply(AutomationUtils.normalize_emp_code)
            old_df = old_df[old_df["_norm_emp_code"] != ""].drop_duplicates(
                subset=["_norm_emp_code"], keep="first"
            )

            # Identify all RPA tracking columns in old SharePoint sheet (anything beyond the 4 base columns)
            rpa_tracking_cols = [
                c
                for c in old_df.columns
                if c not in required_columns and c != "_norm_emp_code" and not c.startswith("Unnamed:")
            ]

            if default_tracking_col and default_tracking_col not in rpa_tracking_cols:
                rpa_tracking_cols.insert(0, default_tracking_col)
                if default_tracking_col not in old_df.columns:
                    old_df[default_tracking_col] = ""

            new_df = new_df.copy()
            new_df["_norm_emp_code"] = new_df["Employee Code"].apply(AutomationUtils.normalize_emp_code)

            preserved_records_count = 0
            for col in rpa_tracking_cols:
                if col in old_df.columns:
                    # Build lookup map {emp_code: tracking_val} for non-empty values
                    lookup = {}
                    for _, row in old_df.iterrows():
                        emp_key = row["_norm_emp_code"]
                        val = row[col]
                        if pd.notna(val) and str(val).strip() != "" and str(val).strip().lower() != "nan":
                            lookup[emp_key] = str(val).strip()

                    new_df[col] = new_df["_norm_emp_code"].map(lambda k: lookup.get(k, ""))
                    if col == rpa_tracking_cols[0]:
                        preserved_records_count = sum(1 for v in new_df[col] if v != "")
                else:
                    if col not in new_df.columns:
                        new_df[col] = ""

            new_df.drop(columns=["_norm_emp_code"], inplace=True, errors="ignore")
            ReportPostFilter._save_dataframe_to_excel(new_df, file_path)

            print(
                f"[{AutomationUtils.ts()}]   [MERGE] Merged RPA tracking columns ({', '.join(rpa_tracking_cols)}) "
                f"from previous SharePoint file: {preserved_records_count} employee(s) already had status recorded, "
                f"{len(new_df) - preserved_records_count} employee(s) pending."
            )
            return new_df

        except Exception as e:
            print(f"[{AutomationUtils.ts()}]   [MERGE] Warning while merging RPA tracking data: {e}")
            return new_df


# SharePoint Sync & Uploader for ELCM
class ELCMSharePointUploader:
    """
    Synchronizes the ELCM Excel report with SharePoint folder SHAREPOINT_TARGET_FOLDER_ELCM:
      1. Fetches the existing Excel file from SharePoint (if any) into memory
      2. Merges RPA email-sent tracking statuses by Employee Code into the new report
      3. Deletes the old Excel file(s) from SharePoint
      4. Uploads the updated new Excel report to SharePoint
    """

    def __init__(self, config: AppConfig):
        self.config = config

    async def _fetch_existing_sharepoint_excel(
        self, client: httpx.AsyncClient, site_id: str, drive_id: str, folder_path: str
    ) -> Optional[pd.DataFrame]:
        """Downloads the existing ELCM Excel report from SharePoint into memory before deletion."""
        try:
            token = await sp_uploader.get_access_token()
            headers = {"Authorization": f"Bearer {token}"}
            encoded_path = sp_uploader.encode_path_segments(folder_path)

            url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}:/children"
            response = await client.get(url, headers=headers)

            if response.status_code != 200:
                return None

            children = response.json().get("value", [])
            excel_items = [
                item
                for item in children
                if "file" in item and item.get("name", "").lower().endswith((".xlsx", ".xls"))
            ]

            if not excel_items:
                print(f"[{AutomationUtils.ts()}]   No existing Excel file found in SharePoint folder '{folder_path}'.")
                return None

            # Sort by lastModifiedDateTime descending to pick the most recent SharePoint file
            excel_items.sort(key=lambda x: x.get("lastModifiedDateTime", ""), reverse=True)
            latest_item = excel_items[0]
            file_name = latest_item.get("name", "existing.xlsx")
            download_url = latest_item.get("@microsoft.graph.downloadUrl")

            if not download_url:
                item_id = latest_item.get("id")
                content_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{item_id}/content"
                dl_res = await client.get(content_url, headers=headers, follow_redirects=True)
            else:
                dl_res = await client.get(download_url, follow_redirects=True)

            if dl_res.status_code != 200:
                print(
                    f"[{AutomationUtils.ts()}]   Warning: Could not download existing SharePoint file '{file_name}' ({dl_res.status_code})."
                )
                return None

            print(
                f"[{AutomationUtils.ts()}]   Fetched existing SharePoint report '{file_name}' to preserve RPA email tracking."
            )
            old_df = pd.read_excel(io.BytesIO(dl_res.content), dtype=str)
            return old_df

        except Exception as e:
            print(f"[{AutomationUtils.ts()}]   Warning fetching existing SharePoint Excel: {e}")
            return None

    async def sync_and_upload_report(self, local_file_path: Path, new_df: Optional[pd.DataFrame] = None) -> bool:
        """
        1. Connects to SharePoint (SHAREPOINT_TARGET_FOLDER_ELCM)
        2. Reads existing SharePoint report and merges RPA email tracking into local_file_path
        3. Deletes old file(s) on SharePoint
        4. Uploads the newly merged ELCM report to SharePoint
        """
        if not local_file_path.exists():
            raise FileNotFoundError(f"ELCM report file not found for upload: {local_file_path}")

        if not self.config.sharepoint_target_folder_elcm:
            raise ValueError("SHAREPOINT_TARGET_FOLDER_ELCM is not set in .env")

        print(f"[{AutomationUtils.ts()}] Connecting to SharePoint for ELCM sync & upload...")
        sp_uploader.TARGET_FOLDER = self.config.sharepoint_target_folder_elcm
        sp_uploader._DRIVE_ID = None

        async with httpx.AsyncClient(timeout=120.0, verify=False) as client:
            site_id = await sp_uploader.get_site_id(client)
            drive_id, folder_path = await sp_uploader.get_drive_details(client, site_id)

            print(f"[{AutomationUtils.ts()}] Target SharePoint folder: {self.config.sharepoint_target_folder_elcm}")
            print(f"[{AutomationUtils.ts()}] Resolved drive folder path: {folder_path}")

            # ==== Fetch existing SharePoint Excel before deleting to preserve tracking ====
            old_sp_df = await self._fetch_existing_sharepoint_excel(client, site_id, drive_id, folder_path)
            if old_sp_df is not None:
                if new_df is None:
                    new_df = pd.read_excel(local_file_path, dtype=str)
                ReportPostFilter.merge_rpa_tracking_from_previous_df(
                    new_df=new_df,
                    old_df=old_sp_df,
                    file_path=local_file_path,
                    required_columns=self.config.required_columns,
                    default_tracking_col=self.config.default_tracking_column,
                )

            # ==== Delete old ELCM report files from SharePoint folder ====
            await sp_uploader.cleanup_sharepoint_folder_files(client, site_id, drive_id, folder_path)

            # ==== Upload the new merged ELCM report to SharePoint ====
            print(f"[{AutomationUtils.ts()}] Uploading {local_file_path.name} to SharePoint...")
            await sp_uploader.upload_file(client, site_id, drive_id, folder_path, local_file_path)
            print(f"[{AutomationUtils.ts()}] SUCCESS! Uploaded {local_file_path.name} to SharePoint ({folder_path}).")
            return True


# Oracle ELCM Report Downloader & Orchestrator
class ELCMOracleReportDownloader:
    """Coordinates browser navigation, parameter selection, report download, Excel column filtering, and SharePoint sync."""

    def __init__(self, app_config: AppConfig):
        self.config = app_config
        self.sso_handler = MicrosoftSSOHandler(
            oracle_email=app_config.oracle_email,
            oracle_password=app_config.oracle_password,
            error_dir=app_config.error_dir,
        )
        self.sp_uploader = ELCMSharePointUploader(app_config)

    async def wait_for_oracle_home(self, page: Page, timeout_s: int = 300):
        """Waits for Oracle Fusion springboard / home page to be visible."""
        try:
            print(f"[{AutomationUtils.ts()}] Waiting for Oracle home page (up to {timeout_s}s)...")
            home_selectors = [
                'text="HR SPOC Reports"',
                'a:has-text("HR SPOC Reports")',
                "#c_64db9d54bc7c456d958c987c0cc83fcf",
                ".fuse-welcome-page",
                ".springboard",
            ]
            deadline = asyncio.get_event_loop().time() + timeout_s

            while asyncio.get_event_loop().time() < deadline:
                current_url = page.url.lower()

                # Handle SSO authentication if login URL is active
                if any(
                    x in current_url
                    for x in [
                        "login.microsoftonline",
                        "login.microsoft",
                        "adfs",
                        "saml",
                        "securid.com",
                        "adani.auth",
                    ]
                ):
                    await self.sso_handler.authenticate(page, timeout_s=300)

                # Check if Oracle Home selector is visible
                for sel in home_selectors:
                    try:
                        if await page.locator(sel).first.is_visible(timeout=1500):
                            print(f"[{AutomationUtils.ts()}]   Home page ready (matched: {sel!r})")
                            await page.wait_for_timeout(2000)
                            return
                    except Exception:
                        pass

                await page.wait_for_timeout(1000)

            await AutomationUtils.save_debug_screenshot(page, self.config.error_dir, "home_load_timeout")
            raise RuntimeError(f"Oracle home page did not load within {timeout_s}s.\nCurrent URL: {page.url}")
        except Exception as e:
            if not isinstance(e, RuntimeError):
                await AutomationUtils.save_debug_screenshot(page, self.config.error_dir, "home_load_exception")
            raise

    def _is_report_page(self, url: str) -> bool:
        """Determines if the given URL corresponds to the BI Publisher report page."""
        try:
            return "bipublisherEntry" in url or "saw.dll" in url
        except Exception:
            return False

    async def wait_for_report_tab(self, context, timeout_s: int = 90) -> Page:
        """Waits for the newly opened BI Publisher report browser tab."""
        try:
            print(f"[{AutomationUtils.ts()}] Waiting for BI Publisher report tab...")
            deadline = asyncio.get_event_loop().time() + timeout_s

            while asyncio.get_event_loop().time() < deadline:
                for p in context.pages:
                    if self._is_report_page(p.url):
                        await p.bring_to_front()
                        try:
                            await p.wait_for_load_state("domcontentloaded", timeout=10_000)
                        except Exception:
                            pass
                        print(f"[{AutomationUtils.ts()}]   Report tab found: {p.url[:80]}...")
                        return p
                await asyncio.sleep(1.0)

            urls = [p.url for p in context.pages]
            raise RuntimeError(
                f"BI Publisher report tab not found within {timeout_s}s.\nOpen tabs:\n  " + "\n  ".join(urls)
            )
        except Exception:
            raise

    async def get_bipublisher_frame(self, report_page: Page, timeout_s: int = 90) -> Frame:
        """Locates the parameter iframe inside the BI Publisher report page."""
        try:
            status_input = "#xdo\\:xdo\\:_paramsP_Assignment_Status_div_input"
            bu_input = "#xdo\\:xdo\\:_paramsP_BU_div_input"
            param_loose = '[id*="_paramsP_Assignment_Status_div_input"], [id*="_paramsP_BU_div_input"]'
            apply_btn = "#reportViewApply"
            deadline = asyncio.get_event_loop().time() + timeout_s

            while asyncio.get_event_loop().time() < deadline:
                children = report_page.main_frame.child_frames
                if children:
                    frame = children[0]
                    try:
                        s_loc = frame.locator(status_input)
                        b_loc = frame.locator(bu_input)
                        if (await s_loc.count() > 0 and await s_loc.first.is_visible(timeout=1500)) or (
                            await b_loc.count() > 0 and await b_loc.first.is_visible(timeout=1500)
                        ):
                            return frame
                    except Exception:
                        pass

                for frame in report_page.frames:
                    try:
                        loc = frame.locator(param_loose)
                        if await loc.count() > 0 and await loc.first.is_visible(timeout=1000):
                            return frame
                    except Exception:
                        pass

                try:
                    await report_page.locator(apply_btn).first.wait_for(state="visible", timeout=1500)
                except Exception:
                    pass

                await asyncio.sleep(1.0)

            await AutomationUtils.save_debug_screenshot(report_page, self.config.error_dir, "report_frame_not_found")
            raise RuntimeError(
                f"BI Publisher parameter iframe not found within {timeout_s}s.\nPage URL: {report_page.url}"
            )
        except Exception:
            raise

    async def select_assignment_status(self, report_page: Page, report_frame: Frame):
        """
        Selects the Assignment Status parameter ('ACTIVE') as recorded in recording.txt:
          1. Clicks #xdo:xdo:_paramsP_Assignment_Status_div_input
          2. Selects #xdo:xdo:_paramsP_Assignment_Status_div_cb_0 / _li_0 span ('ACTIVE')
          3. Closes the dropdown menu
        """
        print(f"[{AutomationUtils.ts()}] Selecting Assignment Status: {self.config.assignment_status_value}...")
        status_input_selectors = [
            "#xdo\\:xdo\\:_paramsP_Assignment_Status_div_input",
            'xpath=//*[@id="xdo:xdo:_paramsP_Assignment_Status_div_input"]',
            '[id*="_paramsP_Assignment_Status_div_input"]',
        ]
        await AutomationUtils.click_with_fallback(
            report_frame, status_input_selectors, "Assignment Status field"
        )
        await report_page.wait_for_timeout(1500)

        # Uncheck "All" if checked so only ACTIVE is selected
        try:
            all_cb = report_frame.locator("#xdo\\:xdo\\:_paramsP_Assignment_Status_div_cb_all").first
            if await all_cb.count() > 0 and await all_cb.is_visible(timeout=1000):
                if await all_cb.is_checked():
                    await all_cb.click()
                    await report_page.wait_for_timeout(500)
        except Exception:
            pass

        # Ensure ACTIVE (#xdo:xdo:_paramsP_Assignment_Status_div_cb_0) is checked
        cb_0 = report_frame.locator("#xdo\\:xdo\\:_paramsP_Assignment_Status_div_cb_0").first
        li_0_span = report_frame.locator("#xdo\\:xdo\\:_paramsP_Assignment_Status_div_li_0 span").first

        status_selected = False
        try:
            if await cb_0.count() > 0 and await cb_0.is_visible(timeout=3000):
                if not await cb_0.is_checked():
                    await cb_0.click()
                    await report_page.wait_for_timeout(500)
                if not await cb_0.is_checked() and await li_0_span.count() > 0:
                    await li_0_span.click()
                    await report_page.wait_for_timeout(500)
                status_selected = True
                print(f"[{AutomationUtils.ts()}]   Assignment Status '{self.config.assignment_status_value}' selected.")
        except Exception:
            pass

        if not status_selected:
            await AutomationUtils.click_with_fallback(
                report_frame,
                [
                    "#xdo\\:xdo\\:_paramsP_Assignment_Status_div_li_0 span",
                    'xpath=//*[@id="xdo:xdo:_paramsP_Assignment_Status_div_li_0"]/label/span',
                    "#xdo\\:xdo\\:_paramsP_Assignment_Status_div_cb_0",
                    'xpath=//*[@id="xdo:xdo:_paramsP_Assignment_Status_div_cb_0"]',
                ],
                f"Assignment Status '{self.config.assignment_status_value}'",
            )

        await report_page.wait_for_timeout(1000)

        # Dismiss the dropdown by clicking outside (matches recording.txt)
        try:
            if report_frame.child_frames:
                await report_frame.child_frames[0].locator("html").click(
                    position={"x": 833, "y": 83}, timeout=2000
                )
            else:
                await report_frame.locator("html").click(position={"x": 10, "y": 10}, timeout=2000)
        except Exception:
            try:
                await report_page.keyboard.press("Escape")
            except Exception:
                pass

        await report_page.wait_for_timeout(1000)

    async def select_business_unit(self, report_page: Page, report_frame: Frame):
        """
        Selects the Business Unit parameter ('Adani Green') as recorded in recording.txt:
          1. Clicks #xdo:xdo:_paramsP_BU_div_input
          2. Unchecks #xdo:xdo:_paramsP_BU_div_li_all ('All') if checked
          3. Clicks #xdo:xdo:_paramsP_BU_div_b span.dialogFloatLeft (Search icon)
          4. Fills #xdo:xdo:_paramsP_BU_div_input_searchDialog_input with 'Adani Green'
          5. Clicks #xdo:xdo:_paramsP_BU_div_input_searchDialog_button (Search)
          6. Clicks #xdo:xdo:_paramsP_BU_div_input_searchDialog_moveAllImg (Move All)
          7. Clicks OK button
        """
        print(f"[{AutomationUtils.ts()}] Selecting Business Unit: {self.config.business_unit_search}...")
        bu_input = report_frame.locator("#xdo\\:xdo\\:_paramsP_BU_div_input").first
        await AutomationUtils.safe_click(bu_input, "Business Unit field")
        await report_page.wait_for_timeout(1000)

        # Uncheck "All" (#xdo:xdo:_paramsP_BU_div_li_all) if checked
        try:
            bu_cb_all = report_frame.locator("#xdo\\:xdo\\:_paramsP_BU_div_cb_all").first
            bu_li_all = report_frame.locator("#xdo\\:xdo\\:_paramsP_BU_div_li_all").first
            if await bu_cb_all.count() > 0 and await bu_cb_all.is_visible(timeout=1500):
                if await bu_cb_all.is_checked():
                    await bu_li_all.click()
                    print(f"[{AutomationUtils.ts()}]   Unchecked 'All' in Business Unit dropdown.")
                    await report_page.wait_for_timeout(1000)
        except Exception:
            pass

        # Ensure dropdown is open before clicking Search icon
        bu_search_selectors = [
            "#xdo\\:xdo\\:_paramsP_BU_div_b span.dialogFloatLeft",
            'xpath=//*[@id="xdo:xdo:_paramsP_BU_div_s"]/span/span[1]',
            "#xdo\\:xdo\\:_paramsP_BU_div_s",
        ]
        if not await report_frame.locator(bu_search_selectors[0]).first.is_visible(timeout=1000):
            await AutomationUtils.safe_click(bu_input, "Business Unit field")
            await report_page.wait_for_timeout(1000)

        await AutomationUtils.click_with_fallback(
            report_frame, bu_search_selectors, "Business Unit search icon"
        )
        await report_page.wait_for_timeout(2000)

        dialog_input = report_frame.locator("#xdo\\:xdo\\:_paramsP_BU_div_input_searchDialog_input")
        await dialog_input.wait_for(state="visible", timeout=15_000)
        await dialog_input.fill(self.config.business_unit_search)

        await AutomationUtils.safe_click(
            report_frame.locator("#xdo\\:xdo\\:_paramsP_BU_div_input_searchDialog_button"),
            "Search button",
        )
        await report_page.wait_for_timeout(5000)
        await AutomationUtils.wait_for_loading(report_frame, timeout_ms=30_000)

        await AutomationUtils.safe_click(
            report_frame.locator("#xdo\\:xdo\\:_paramsP_BU_div_input_searchDialog_moveAllImg"),
            "Move All button",
        )
        await report_page.wait_for_timeout(2000)

        ok_selectors = [
            'xpath=//*[@id="xdo:xdo:_paramsP_BU_div_input_searchDialog_dialogTable"]/tbody/tr[3]/td[2]/table/tbody/tr/td[2]/button[1]',
            "button.button_mo",
        ]
        await AutomationUtils.click_with_fallback(report_frame, ok_selectors, "OK button")
        await report_page.wait_for_timeout(2000)

    def _download_extension(self, suggested_filename: str) -> str:
        """Extracts valid file extension from suggested filename or defaults to .xlsx."""
        try:
            ext = Path(suggested_filename).suffix.lower() if suggested_filename else ""
            if ext in (".xls", ".xlsx", ".xlsm", ".csv"):
                return ext
            return ".xlsx"
        except Exception:
            return ".xlsx"

    async def open_format_menu_and_click_excel(self, report_page: Page, report_frame: Frame):
        """
        Opens the View Format menu (#xdo:viewFormatLink) and selects the Excel export option
        (#_xdoFMenu1 li:nth-of-type(6) div.itemTxt) as recorded in recording.txt.
        """
        format_dropdown_selectors = [
            "#xdo\\:viewFormatLink > img:nth-of-type(2)",
            'xpath=//*[@id="xdo:viewFormatLink"]/img[2]',
            "#xdo\\:viewFormatLink",
        ]
        excel_menu_item_selectors = [
            "#_xdoFMenu1 li:nth-of-type(6) div.itemTxt",
            'xpath=//*[@id="_xdoFMenu1"]/div/div/ul/li[6]/div/a/div[2]',
            "#_xdoFMenu1 li:nth-of-type(6)",
            'xpath=//*[@id="_xdoFMenu1"]//li[contains(.,"Excel")]',
        ]

        for attempt in range(1, 5):
            print(f"[{AutomationUtils.ts()}] Opening View Format menu (Attempt {attempt}/4)...")
            await AutomationUtils.click_with_fallback(
                report_frame, format_dropdown_selectors, "View Format dropdown"
            )
            await report_page.wait_for_timeout(1500)

            try:
                if await report_frame.locator("#_xdoFMenu1").first.is_visible(timeout=3000):
                    break
            except Exception:
                pass
            await report_page.wait_for_timeout(1500)

        print(f"[{AutomationUtils.ts()}] Selecting Excel format (#_xdoFMenu1 li:nth-of-type(6))...")
        await AutomationUtils.click_with_fallback(
            report_frame, excel_menu_item_selectors, "Excel export menu item"
        )

    async def click_apply_and_wait_for_download(
        self, report_page: Page, report_frame: Frame, timeout_ms: int = 1_200_000
    ) -> Optional[str]:
        """
        Triggers report generation via Apply button, selects Excel format from View Format menu,
        and waits for the long-running ELCM Excel download to complete.
        """
        download_started_event = asyncio.Event()

        def _on_download_start(dl):
            print(f"[{AutomationUtils.ts()}]   Browser download event triggered: {dl.suggested_filename or 'ELCM_Report'}")
            download_started_event.set()

        report_page.on("download", _on_download_start)

        try:
            print(
                f"[{AutomationUtils.ts()}] Applying report & exporting Excel "
                f"(waiting up to {timeout_ms // 60_000} minutes)..."
            )
            try:
                async with report_page.expect_download(timeout=timeout_ms) as dl_info:
                    # ==== Click Apply button ====
                    await AutomationUtils.safe_click(
                        report_frame.locator("#reportViewApply"), "Apply button"
                    )
                    await report_page.wait_for_timeout(2000)
                    await AutomationUtils.wait_for_loading(report_frame, timeout_ms=180_000)

                    # ==== Open View Format menu and select Excel format ====
                    if not download_started_event.is_set():
                        await self.open_format_menu_and_click_excel(report_page, report_frame)

                    # ==== Poll while Oracle BI Publisher generates the large ELCM Excel file ====
                    start_wait = asyncio.get_event_loop().time()
                    deadline = start_wait + (timeout_ms / 1000.0)
                    last_log_time = start_wait
                    report_completed_logged = False

                    while not download_started_event.is_set() and asyncio.get_event_loop().time() < deadline:
                        now = asyncio.get_event_loop().time()
                        elapsed = int(now - start_wait)

                        if not report_completed_logged:
                            for frame in report_page.frames:
                                try:
                                    loc = frame.get_by_text("Report Completed", exact=True)
                                    if await loc.count() > 0 and await loc.first.is_visible(timeout=300):
                                        print(f"[{AutomationUtils.ts()}]   Report Completed ({elapsed}s elapsed).")
                                        report_completed_logged = True
                                        break
                                except Exception:
                                    pass

                        if now - last_log_time >= 15:
                            print(
                                f"[{AutomationUtils.ts()}]   Still generating/downloading ELCM Excel report... "
                                f"({elapsed}s elapsed)"
                            )
                            last_log_time = now

                        await asyncio.sleep(1.0)

                # ==== Wait for the full file stream transfer to finish writing to disk ====
                download = await dl_info.value
                suggested = download.suggested_filename or ""
                ext = self._download_extension(suggested)
                save_path = self.config.download_dir / f"{AutomationUtils.report_save_stem()}{ext}"

                print(f"[{AutomationUtils.ts()}] Saving downloaded stream to {save_path.name}...")
                await download.save_as(save_path)

                failure = await download.failure()
                if failure:
                    raise RuntimeError(f"Browser download failed: {failure}")

                size = save_path.stat().st_size
                if size < 512:
                    raise RuntimeError(f"Downloaded file is too small ({size} bytes) — likely incomplete.")

                print(f"[{AutomationUtils.ts()}] Saved: {save_path.name} ({size:,} bytes)")
                return str(save_path)

            except PlaywrightTimeoutError:
                await AutomationUtils.save_debug_screenshot(
                    report_page, self.config.error_dir, "download_timeout"
                )
                return None
        except Exception as e:
            print(f"[{AutomationUtils.ts()}] Error during report download: {e}")
            await AutomationUtils.save_debug_screenshot(
                report_page, self.config.error_dir, "download_exception"
            )
            return None
        finally:
            try:
                report_page.remove_listener("download", _on_download_start)
            except Exception:
                pass

    async def download(self):
        """Main execution flow for downloading, filtering columns, merging RPA tracking, and uploading to SharePoint."""
        print(
            f"[{AutomationUtils.ts()}] Starting ELCM report download ({'headless' if self.config.headless else 'visible'})"
        )

        # Clean up old reports locally before starting
        try:
            for f in self.config.download_dir.iterdir():
                if f.is_file():
                    try:
                        f.unlink()
                        print(f"[{AutomationUtils.ts()}] [CLEANUP] Deleted old local file: {f.name}")
                    except Exception as e:
                        print(f"[{AutomationUtils.ts()}] [CLEANUP] Warning: could not delete {f.name}: {e}")
        except Exception as e:
            print(f"[{AutomationUtils.ts()}] [CLEANUP] Local cleanup error: {e}")

        AutomationUtils.cleanup_profile_locks(self.config.automation_profile_dir)

        context = None
        result = None
        try:
            async with async_playwright() as p:
                print(f"[{AutomationUtils.ts()}] Launching browser...")
                chrome_args = [
                    "--no-first-run",
                    "--no-default-browser-check",
                    "--disable-blink-features=AutomationControlled",
                    "--disable-session-crashed-bubble",
                    "--hide-crash-restore-bubble",
                ]

                if self.config.headless:
                    launch_kwargs = {
                        "headless": True,
                        "viewport": {"width": 1920, "height": 1080},
                    }
                else:
                    chrome_args.append("--start-maximized")
                    launch_kwargs = {
                        "headless": False,
                        "no_viewport": True,
                    }

                context = await p.chromium.launch_persistent_context(
                    user_data_dir=str(self.config.automation_profile_dir),
                    args=chrome_args,
                    **launch_kwargs,
                )

                page = context.pages[0] if context.pages else await context.new_page()
                if not self.config.headless:
                    try:
                        await page.bring_to_front()
                    except Exception:
                        pass

                print(f"[{AutomationUtils.ts()}] Navigating to Oracle Fusion...")
                await page.goto(self.config.oracle_url, wait_until="domcontentloaded", timeout=90_000)
                await page.wait_for_timeout(3000)

                await self.wait_for_oracle_home(page, timeout_s=300)
                await page.wait_for_timeout(2000)

                print(f"[{AutomationUtils.ts()}] Opening ELCM report...")
                await AutomationUtils.click_with_fallback(
                    page,
                    [
                        'a:has-text("HR SPOC Reports")',
                        'text="HR SPOC Reports"',
                        "#c_64db9d54bc7c456d958c987c0cc83fcf",
                        'xpath=//*[@id="c_64db9d54bc7c456d958c987c0cc83fcf"]',
                        'xpath=//a[normalize-space()="HR SPOC Reports"]',
                        'xpath=//*[contains(text(),"HR SPOC Reports")]',
                    ],
                    "HR SPOC Reports",
                )
                await page.wait_for_timeout(3000)

                elcm_selectors = [
                    "#c_b1f5de02298446578e3e2cad15bc0285",
                    'xpath=//*[@id="c_b1f5de02298446578e3e2cad15bc0285"]',
                    "#yourapps_c_64db9d54bc7c456d958c987c0cc83fcf > div:nth-of-type(1) path.svg-outline",
                    'xpath=//*[@id="c_b1f5de02298446578e3e2cad15bc0285"]/svg/path[8]',
                    "#yourapps_c_64db9d54bc7c456d958c987c0cc83fcf > div:nth-of-type(1)",
                    'a:has-text("ELCM")',
                    'text="ELCM"',
                    'xpath=//*[contains(text(),"ELCM")]',
                ]
                try:
                    async with context.expect_page(timeout=60_000) as new_page_info:
                        await AutomationUtils.click_with_fallback(page, elcm_selectors, "ELCM Report")
                    report_page = await new_page_info.value
                except PlaywrightTimeoutError:
                    report_page = await self.wait_for_report_tab(context, timeout_s=30)

                await report_page.bring_to_front()
                try:
                    await report_page.wait_for_load_state("networkidle", timeout=30_000)
                except PlaywrightTimeoutError:
                    await report_page.wait_for_load_state("domcontentloaded", timeout=15_000)
                await report_page.wait_for_timeout(3000)

                report_frame = await self.get_bipublisher_frame(report_page, timeout_s=90)
                await AutomationUtils.wait_for_loading(report_frame, timeout_ms=30_000)
                print(f"[{AutomationUtils.ts()}] Report parameters loaded")

                # Set Assignment Status (ACTIVE)
                await self.select_assignment_status(report_page, report_frame)

                # Set Business Unit (Adani Green)
                await self.select_business_unit(report_page, report_frame)

                # Apply report, select Excel format, and wait for download
                result = await self.click_apply_and_wait_for_download(
                    report_page, report_frame, timeout_ms=self.config.download_timeout_ms
                )

        except Exception as e:
            print(f"[{AutomationUtils.ts()}] FATAL: {e}\n{traceback.format_exc()}")
            raise
        finally:
            if context:
                try:
                    for pg in context.pages:
                        try:
                            await pg.close()
                        except Exception:
                            pass
                    await context.close()
                except Exception:
                    pass

        # Post-process downloaded Excel, merge RPA tracking from existing SharePoint file,
        # delete old SharePoint file, and upload the updated report
        if result:
            file_path = Path(result)
            cleaned_df = ReportPostFilter.process_elcm_excel(
                file_path,
                required_columns=self.config.required_columns,
                target_locations=self.config.target_locations,
                default_tracking_col=self.config.default_tracking_column,
            )
            print(f"\nSUCCESS! Report saved locally to:\n   {file_path}")
            await self.sp_uploader.sync_and_upload_report(file_path, new_df=cleaned_df)
        else:
            print(f"\nDownload did not complete in time. Check screenshot in {self.config.error_dir}")


# Main entry point for the script
async def download_elcm_report():
    """Main entrypoint function for ELCM report downloader."""
    setup_logger("download_elcm_report.log")
    app_config = AppConfig()
    downloader = ELCMOracleReportDownloader(app_config)
    await downloader.download()


if __name__ == "__main__":
    asyncio.run(download_elcm_report())
