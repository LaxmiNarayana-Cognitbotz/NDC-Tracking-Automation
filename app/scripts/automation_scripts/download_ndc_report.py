"""
Oracle Fusion - NDC Process Request Status Report Downloader
"""

import sys
from pathlib import Path

# Add project root to sys.path so the script can be run directly or via VS Code Run
ROOT_DIR = Path(__file__).resolve().parents[3]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import asyncio
import os
import traceback
from datetime import datetime
from typing import List, Optional

import pandas as pd
from dotenv import load_dotenv
from playwright.async_api import Frame, Page, TimeoutError as PlaywrightTimeoutError, async_playwright

from app.helpers.auth import MicrosoftSSOHandler
from app.core.config import AUTOMATION_PROFILE_DIR, NDC_ERRORS_DIR, NDC_REPORTS_DIR
from app.core.logger import setup_logger
from app.core.process import suppress_subprocess_window
import app.scripts.processing_scripts.upload_to_sharepoint as sp_uploader

# ==== Suppress Chrome subprocess window on Windows ====
suppress_subprocess_window()

# ==== Logger Setup ====
LOG_FILE = setup_logger("download_ndc_report.log")

load_dotenv(verbose=True)


# ==== Application Configuration ====
class AppConfig:
    """Manages environment variables, credentials, paths, and runtime settings."""

    def __init__(self):
        try:
            # Oracle Fusion credentials and URL
            self.oracle_url: str = (os.getenv("ORACLE_URL") or "").strip()
            self.oracle_email: str = (os.getenv("ORACLE_EMAIL") or "").strip()
            self.oracle_password: str = (os.getenv("ORACLE_PASSWORD") or "").strip()

            # Output and storage directories
            self.download_dir: Path = NDC_REPORTS_DIR
            self.automation_profile_dir: Path = AUTOMATION_PROFILE_DIR
            self.error_dir: Path = NDC_ERRORS_DIR

            # Browser configuration
            self.headless: bool = self._parse_bool(os.getenv("HEADLESS"), default=True)
            self.start_maximized: bool = self._parse_bool(os.getenv("START_MAXIMIZED"), default=True)

            # Report parameters
            self.business_unit_search: str = "Adani Green"
            self.from_date_value: str = "02-09-2026"

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


# ==== Automation Utilities ====
class AutomationUtils:
    """Utility helper functions for timing, screenshots, locks, and DOM actions."""

    @staticmethod
    def ts() -> str:
        """Returns the current formatted timestamp string."""
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    @staticmethod
    def report_save_stem() -> str:
        """Generates a timestamped filename stem for the downloaded report."""
        try:
            now = datetime.now()
            return (
                f"NDC_Process_Request_Status_{now.day}_"
                f"{now.strftime('%B')}_{now.year}_{now.strftime('%I.%M%p')}"
            )
        except Exception:
            return f"NDC_Process_Request_Status_{int(datetime.now().timestamp())}"

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
        """Captures a full-page diagnostic screenshot upon errors."""
        try:
            error_dir.mkdir(parents=True, exist_ok=True)
            now_str = datetime.now().strftime("%d_%b_%Y_(%I.%M%p)")
            path = error_dir / f"ndc_{label}_{now_str}.png"
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


# ==== Report Post-Filter ====
class ReportPostFilter:
    """Handles parsing and post-filtering of downloaded NDC Excel reports."""

    @staticmethod
    def filter_for_adani_green(file_path: Path):
        """
        If the downloaded report's Business Unit metadata contains 'All',
        filters the data rows to only keep Business Units starting with 'Adani Green'.
        """
        try:
            print(f"[{AutomationUtils.ts()}] Checking if report requires post-download filtering...")
            dfs = []

            # Parse file either as HTML table or binary Excel
            try:
                is_html = False
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    head = f.read(1000)
                    if "<html" in head.lower() or "<table" in head.lower():
                        is_html = True

                if is_html:
                    dfs = pd.read_html(str(file_path))
                else:
                    df = pd.read_excel(file_path, header=None)
                    dfs = [df]
            except Exception as e:
                print(f"[{AutomationUtils.ts()}]   Error parsing Excel file {file_path.name}: {e}")
                return

            # Search for Business Unit metadata row and table
            bu_meta_tbl_idx = None
            bu_meta_row = None
            bu_meta_col = None
            bu_meta_val = None

            for idx, df in enumerate(dfs):
                for r_idx in range(min(30, len(df))):
                    row_vals = [str(x).strip() for x in df.iloc[r_idx].tolist()]
                    col_names = [str(c).strip() for c in df.columns]

                    for c_idx, val in enumerate(row_vals):
                        if "Business Unit" in val:
                            if ":" in val and len(val.split(":", 1)[1].strip()) > 0:
                                bu_meta_tbl_idx = idx
                                bu_meta_row = r_idx
                                bu_meta_col = c_idx
                                bu_meta_val = val.split(":", 1)[1].strip()
                                break
                            else:
                                for next_col_idx, next_val in enumerate(row_vals[c_idx + 1 :], start=c_idx + 1):
                                    if next_val and next_val != "nan":
                                        bu_meta_tbl_idx = idx
                                        bu_meta_row = r_idx
                                        bu_meta_col = next_col_idx
                                        bu_meta_val = next_val
                                        break
                            if bu_meta_tbl_idx is not None:
                                break

                    if bu_meta_tbl_idx is not None:
                        break

                    for c_idx, val in enumerate(col_names):
                        if "Business Unit" in val:
                            if ":" in val and len(val.split(":", 1)[1].strip()) > 0:
                                bu_meta_tbl_idx = idx
                                bu_meta_row = -1
                                bu_meta_col = c_idx
                                bu_meta_val = val.split(":", 1)[1].strip()
                                break
                            else:
                                if len(df) > 0:
                                    bu_meta_tbl_idx = idx
                                    bu_meta_row = -1
                                    bu_meta_col = c_idx
                                    bu_meta_val = str(df.iloc[0, c_idx]).strip()
                                    break
                    if bu_meta_tbl_idx is not None:
                        break

            if bu_meta_tbl_idx is None:
                print(f"[{AutomationUtils.ts()}]   Could not locate 'Business Unit' metadata cell in tables.")
                return

            # Check if value is "All"
            is_all = False
            if bu_meta_val:
                cleaned_val = bu_meta_val.lower().replace(":", "").strip()
                if cleaned_val == "all":
                    is_all = True

            if not is_all:
                print(f"[{AutomationUtils.ts()}]   Business Unit already filtered ('{bu_meta_val}'). No local filtering needed.")
                return

            # Locate data table and headers
            data_tbl_idx = None
            header_row_idx = None
            bu_col_idx = None

            for idx, df in enumerate(dfs):
                start_row = 5 if idx == bu_meta_tbl_idx else 0
                for r_idx in range(start_row, len(df)):
                    row_vals = [str(x).strip() for x in df.iloc[r_idx].tolist()]
                    if "Business Unit" in row_vals:
                        non_empty = sum(1 for x in row_vals if x and x != "nan")
                        if non_empty >= 4:
                            data_tbl_idx = idx
                            header_row_idx = r_idx
                            bu_col_idx = row_vals.index("Business Unit")
                            break
                if data_tbl_idx is not None:
                    break

            if data_tbl_idx is None:
                print(f"[{AutomationUtils.ts()}]   Error: Could not locate table header row containing 'Business Unit' column.")
                return

            # Perform filtering
            target_df = dfs[data_tbl_idx]
            metadata_df = target_df.iloc[:header_row_idx].copy()
            header_row = target_df.iloc[header_row_idx].copy()
            data_df = target_df.iloc[header_row_idx + 1 :].copy()

            original_row_count = len(data_df)

            def should_keep(row):
                val = str(row.iloc[bu_col_idx]).strip()
                if not val or val == "nan":
                    return True
                return val.lower().startswith("adani green")

            filtered_data_df = data_df[data_df.apply(should_keep, axis=1)]
            filtered_row_count = len(filtered_data_df)

            new_meta_val = (
                ":[Adani Green > Renewable > Hydro O&M,Adani Green > Renewable > Hydro Projects,"
                "Adani Green > Renewable > Solar O&M,Adani Green > Renewable > Solar Projects,"
                "Adani Green > Renewable > Wind O&M,Adani Green > Renewable > Wind Projects] (Filtered)"
            )

            meta_df = dfs[bu_meta_tbl_idx]
            if bu_meta_row == -1:
                new_cols = list(meta_df.columns)
                new_cols[bu_meta_col] = "Business Unit " + new_meta_val
                meta_df.columns = new_cols
            else:
                meta_df.iloc[bu_meta_row, bu_meta_col] = new_meta_val

            # Reconstruct and write back
            try:
                if len(dfs) == 1:
                    final_df = pd.concat(
                        [metadata_df, pd.DataFrame([header_row]), filtered_data_df],
                        ignore_index=True,
                    )
                    if file_path.suffix.lower() == ".xls":
                        final_df.to_excel(file_path, header=False, index=False, engine="openpyxl")
                    else:
                        final_df.to_excel(file_path, header=False, index=False, engine="openpyxl")
                else:
                    dfs[data_tbl_idx] = pd.concat(
                        [metadata_df, pd.DataFrame([header_row]), filtered_data_df],
                        ignore_index=True,
                    )
                    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                        start_row = 0
                        for df_to_write in dfs:
                            df_to_write.to_excel(
                                writer, sheet_name="Sheet1", startrow=start_row, index=False, header=False
                            )
                            start_row += len(df_to_write) + 1

                print(
                    f"[{AutomationUtils.ts()}] SUCCESS: Excel filter applied! Kept {filtered_row_count}/{original_row_count} rows starting with 'Adani Green'."
                )
            except Exception as e:
                print(f"[{AutomationUtils.ts()}]   Error saving filtered Excel file: {e}")

        except Exception as e:
            print(f"[{AutomationUtils.ts()}] Exception in filter_for_adani_green: {e}")


# ==== Oracle Report Downloader & Orchestrator ====
class NDCOracleReportDownloader:
    """Coordinates browser navigation, parameter selection, report download, and file processing."""

    def __init__(self, app_config: AppConfig):
        self.config = app_config
        self.sso_handler = MicrosoftSSOHandler(
            oracle_email=app_config.oracle_email,
            oracle_password=app_config.oracle_password,
            error_dir=app_config.error_dir,
        )

    async def wait_for_oracle_home(self, page: Page, timeout_s: int = 300):
        """Waits for Oracle Fusion springboard / home page to be visible."""
        try:
            print(f"[{AutomationUtils.ts()}] Waiting for Oracle home page (up to {timeout_s}s)...")
            home_selectors = [
                'text="HR SPOC Reports"',
                'a:has-text("HR SPOC Reports")',
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
                    await self.sso_handler.authenticate(page, timeout_s=120)

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
            return "bipublisherEntry" in url or ("saw.dll" in url and "NDC" in url)
        except Exception:
            return False

    async def wait_for_report_tab(self, context, timeout_s: int = 90) -> Page:
        """Waits for the newly opened BI Publisher report browser tab."""
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

    async def get_bipublisher_frame(self, report_page: Page, timeout_s: int = 90) -> Frame:
        """Locates the parameter iframe inside the BI Publisher report page."""
        bu_input = "#xdo\\:xdo\\:_paramsP_Business_Unit_div_input"
        bu_input_loose = '[id*="Business_Unit_div_input"]'
        apply_btn = "#reportViewApply"
        deadline = asyncio.get_event_loop().time() + timeout_s

        while asyncio.get_event_loop().time() < deadline:
            children = report_page.main_frame.child_frames
            if children:
                frame = children[0]
                try:
                    loc = frame.locator(bu_input)
                    if await loc.count() > 0 and await loc.first.is_visible(timeout=2000):
                        return frame
                except Exception:
                    pass

            for frame in report_page.frames:
                try:
                    loc = frame.locator(bu_input_loose)
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

    async def wait_for_report_completed(
        self, report_page: Page, report_frame: Frame, timeout_ms: int = 300_000
    ) -> bool:
        """Polls until the report area shows 'Report Completed'."""
        try:
            print(f"[{AutomationUtils.ts()}] Waiting for 'Report Completed' (up to {timeout_ms // 1000}s)...")
            deadline = asyncio.get_event_loop().time() + timeout_ms / 1000
            while asyncio.get_event_loop().time() < deadline:
                for frame in report_page.frames:
                    try:
                        loc = frame.get_by_text("Report Completed", exact=True)
                        if await loc.count() > 0 and await loc.first.is_visible(timeout=500):
                            print(f"[{AutomationUtils.ts()}]   Report Completed.")
                            return True
                    except Exception:
                        pass
                await asyncio.sleep(2)
            print(f"[{AutomationUtils.ts()}]   'Report Completed' not seen — still waiting for download...")
            return False
        except Exception as e:
            print(f"[{AutomationUtils.ts()}]   Warning while waiting for report completion: {e}")
            return False

    def _download_extension(self, suggested_filename: str) -> str:
        """Extracts valid file extension from suggested filename or defaults to .xls."""
        try:
            ext = Path(suggested_filename).suffix.lower() if suggested_filename else ""
            if ext in (".xls", ".xlsx", ".xlsm", ".csv"):
                return ext
            return ".xls"
        except Exception:
            return ".xls"

    async def click_apply_and_wait_for_download(
        self, report_page: Page, report_frame: Frame, timeout_ms: int = 300_000
    ) -> Optional[str]:
        """Triggers report generation via Apply button and captures downloaded file."""
        try:
            print(f"[{AutomationUtils.ts()}] Applying report & downloading...")
            try:
                async with report_page.expect_download(timeout=timeout_ms) as dl_info:
                    await AutomationUtils.safe_click(report_frame.locator("#reportViewApply"), "Apply button")
                    await AutomationUtils.wait_for_loading(report_frame, timeout_ms=90_000)
                    await self.wait_for_report_completed(report_page, report_frame, timeout_ms=timeout_ms)

                download = await dl_info.value
                failure = await download.failure()
                if failure:
                    raise RuntimeError(f"Browser download failed: {failure}")

                suggested = download.suggested_filename or ""
                ext = self._download_extension(suggested)
                save_path = self.config.download_dir / f"{AutomationUtils.report_save_stem()}{ext}"

                await download.save_as(save_path)
                size = save_path.stat().st_size
                if size < 512:
                    raise RuntimeError(f"Downloaded file is too small ({size} bytes) — likely incomplete.")

                print(f"[{AutomationUtils.ts()}] Saved: {save_path.name} ({size:,} bytes)")
                return str(save_path)

            except PlaywrightTimeoutError:
                await AutomationUtils.save_debug_screenshot(report_page, self.config.error_dir, "download_timeout")
                return None
        except Exception as e:
            print(f"[{AutomationUtils.ts()}] Error during report download: {e}")
            await AutomationUtils.save_debug_screenshot(report_page, self.config.error_dir, "download_exception")
            return None

    async def download(self):
        """Main execution flow for downloading and processing the NDC report."""
        setup_logger("download_ndc_report.log")
        print(
            f"[{AutomationUtils.ts()}] Starting NDC report download ({'headless' if self.config.headless else 'visible'})"
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
                    accept_downloads=True,
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

                await self.wait_for_oracle_home(page, timeout_s=120)
                await page.wait_for_timeout(2000)

                print(f"[{AutomationUtils.ts()}] Opening NDC Process Request Status report...")
                await AutomationUtils.click_with_fallback(
                    page,
                    [
                        'a:has-text("HR SPOC Reports")',
                        'text="HR SPOC Reports"',
                        "#c_64db9d54bc7c456d958c987c0cc83fcf",
                        'xpath=//a[normalize-space()="HR SPOC Reports"]',
                        'xpath=//*[contains(text(),"HR SPOC Reports")]',
                    ],
                    "HR SPOC Reports",
                )
                await page.wait_for_timeout(3000)

                ndc_selectors = [
                    'a:has-text("NDC Process Request Status")',
                    'text="NDC Process Request Status"',
                    "#c_fc8798b561de4e0bb05f5b982200a547_0",
                    'text="NDC Process Request"',
                    'xpath=//*[contains(text(),"NDC Process Request")]',
                ]
                try:
                    async with context.expect_page(timeout=60_000) as new_page_info:
                        await AutomationUtils.click_with_fallback(page, ndc_selectors, "NDC Process Request Status")
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

                # Set From Date
                from_date_selectors = [
                    "#_paramsP_FROMDATE",
                    'xpath=//*[@id="_paramsP_FROMDATE"]',
                    'input[id="_paramsP_FROMDATE"]',
                ]
                from_date_filled = False
                for sel in from_date_selectors:
                    try:
                        loc = report_frame.locator(sel).first
                        if await loc.is_visible(timeout=5000):
                            await loc.click(click_count=3)
                            await report_page.wait_for_timeout(500)
                            await loc.fill(self.config.from_date_value)
                            await report_page.wait_for_timeout(500)
                            await loc.click()
                            await report_page.wait_for_timeout(1000)
                            from_date_filled = True
                            break
                    except Exception:
                        continue

                if from_date_filled:
                    print(f"[{AutomationUtils.ts()}] From Date set: {self.config.from_date_value}")
                else:
                    await AutomationUtils.save_debug_screenshot(
                        report_page, self.config.error_dir, "from_date_not_found"
                    )
                    print(f"[{AutomationUtils.ts()}] WARNING: Could not fill From Date field!")
                await report_page.wait_for_timeout(2000)

                # Set Business Unit
                print(f"[{AutomationUtils.ts()}] Selecting Business Unit: {self.config.business_unit_search}...")
                await AutomationUtils.safe_click(
                    report_frame.locator("#xdo\\:xdo\\:_paramsP_Business_Unit_div_input"),
                    "Business Unit field",
                )
                await AutomationUtils.safe_click(
                    report_frame.locator("#xdo\\:xdo\\:_paramsP_Business_Unit_div_b span.dialogFloatLeft"),
                    "Business Unit search icon",
                )
                await report_page.wait_for_timeout(2000)

                dialog_input = report_frame.locator("#xdo\\:xdo\\:_paramsP_Business_Unit_div_input_searchDialog_input")
                await dialog_input.wait_for(state="visible", timeout=15_000)
                await dialog_input.fill(self.config.business_unit_search)

                await AutomationUtils.safe_click(
                    report_frame.locator("#xdo\\:xdo\\:_paramsP_Business_Unit_div_input_searchDialog_button"),
                    "Search button",
                )
                await report_page.wait_for_timeout(5000)

                await AutomationUtils.safe_click(
                    report_frame.locator("#xdo\\:xdo\\:_paramsP_Business_Unit_div_input_searchDialog_moveAllImg"),
                    "Move All button",
                )
                await report_page.wait_for_timeout(2000)

                ok_selectors = [
                    'xpath=//*[@id="xdo:xdo:_paramsP_Business_Unit_div_input_searchDialog_dialogTable"]/tbody/tr[3]/td[2]/table/tbody/tr/td[2]/button[1]',
                    "button.button_mo",
                ]
                await AutomationUtils.click_with_fallback(report_frame, ok_selectors, "OK button")
                await report_page.wait_for_timeout(2000)

                # Download report and apply post-filter
                result = await self.click_apply_and_wait_for_download(
                    report_page, report_frame, timeout_ms=300_000
                )

                if result:
                    ReportPostFilter.filter_for_adani_green(Path(result))
                    print(f"\nSUCCESS! Report saved to:\n   {result}")

                    # ==== Push directly to SharePoint ====
                    try:
                        print(f"[{AutomationUtils.ts()}] Pushing NDC report directly to SharePoint...")
                        target_sp_folder = os.getenv("SHAREPOINT_TARGET_FOLDER")
                        sp_uploader.TARGET_FOLDER = target_sp_folder
                        sp_uploader._DRIVE_ID = None
                        async with sp_uploader.httpx.AsyncClient(verify=False) as client:
                            site_id = await sp_uploader.get_site_id(client)
                            drive_id, folder_path = await sp_uploader.get_drive_details(client, site_id, target_sp_folder)
                            target_folder_url = f"{folder_path}/NDC_Reports"

                            await sp_uploader.cleanup_sharepoint_folder_files(client, site_id, drive_id, target_folder_url)
                            local_file = Path(result)
                            if local_file.exists():
                                await sp_uploader.upload_file(client, site_id, drive_id, target_folder_url, local_file)
                                print(f"[{AutomationUtils.ts()}] SharePoint sync completed successfully: {local_file.name}")
                            else:
                                await sp_uploader.upload_ndc_reports(client, site_id, drive_id, folder_path)
                                print(f"[{AutomationUtils.ts()}] SharePoint sync completed successfully.")
                    except Exception as sp_err:
                        print(f"[{AutomationUtils.ts()}] [WARNING] SharePoint upload failed: {sp_err}\n{traceback.format_exc()}")
                else:
                    print(f"\nDownload did not complete in time. Check screenshot in {self.config.download_dir}")

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


# ==== Main Entry Point ====
async def download_ndc_report():
    """Main entrypoint function for NDC report downloader."""
    app_config = AppConfig()
    downloader = NDCOracleReportDownloader(app_config)
    await downloader.download()


if __name__ == "__main__":
    asyncio.run(download_ndc_report())
