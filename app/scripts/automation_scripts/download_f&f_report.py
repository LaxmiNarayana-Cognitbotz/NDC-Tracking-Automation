"""
OpenText Content Server - F&F (Full & Final) Exit Document Downloader
"""

import sys
from pathlib import Path

# Add project root to sys.path so the script can be run directly or via VS Code Run
ROOT_DIR = Path(__file__).resolve().parents[3]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import asyncio
import glob
import os
import re
import traceback
from datetime import datetime
from typing import Dict, List, Optional, Tuple

from dotenv import load_dotenv
from playwright.async_api import BrowserContext, Page, TimeoutError as PlaywrightTimeoutError, async_playwright

from app.helpers.auth import MicrosoftSSOHandler
from app.core.config import AUTOMATION_PROFILE_DIR, FF_ERRORS_DIR, FF_REPORTS_DIR
from app.core.logger import setup_logger
from app.core.process import suppress_subprocess_window
import app.scripts.processing_scripts.upload_to_sharepoint as sp_uploader

# ==== Suppress Chrome subprocess window on Windows ====
suppress_subprocess_window()

# ==== Logger Setup ====
LOG_FILE = setup_logger("download_ff_report.log")

load_dotenv(verbose=True)


# ==== Result Status Constants ====
STATUS_DOWNLOADED = "DOWNLOADED"        # F&F sheet found and saved
STATUS_NO_FF_DOCS = "NO_FF_DOCS"        # Folder found, but no F&F sheet uploaded
STATUS_NOT_FOUND = "NOT_FOUND_0_HITS"   # Content Server returned 0 search results
STATUS_ERROR = "ERROR"                  # Search timeout, network failure, etc.


# ==== Application Configuration ====
class AppConfig:
    """Manages environment variables, paths, timeouts, and browser settings."""

    def __init__(self):
        try:
            # Credentials & URLs
            self.oracle_email: str = (os.getenv("ORACLE_EMAIL") or "").strip()
            self.oracle_password: str = (os.getenv("ORACLE_PASSWORD") or "").strip()
            self.content_server_url: str = (os.getenv("CONTENT_SERVER_URL") or "").strip()

            # Output and storage directories
            self.download_dir: Path = FF_REPORTS_DIR
            self.automation_profile_dir: Path = AUTOMATION_PROFILE_DIR
            self.error_dir: Path = FF_ERRORS_DIR

            # Browser configuration
            self.headless: bool = self._parse_bool(os.getenv("HEADLESS"), default=True)
            self.start_maximized: bool = self._parse_bool(os.getenv("START_MAXIMIZED"), default=True)

            # Timeouts & limits
            self.element_timeout: int = 15_000
            self.nav_timeout: int = 90_000
            self.max_concurrent_tabs: int = 5
            self.search_results_wait: int = 45_000

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
    def cleanup_profile_locks(profile_dir: Path):
        """Remove Chrome lock files on the automation profile only."""
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
    def cleanup_error_screenshots(error_dir: Path, employee_number: str):
        """Delete all debug error screenshots for the given employee from errors/FF_Errors/."""
        if not error_dir.exists():
            return
        patterns = [
            str(error_dir / f"ff_*_{employee_number}_*.png"),
            str(error_dir / f"download_ff_report_*_{employee_number}_*.png"),
            str(error_dir / f"debug_*_{employee_number}*.png"),
        ]
        for pattern in patterns:
            for path in glob.glob(pattern):
                try:
                    os.remove(path)
                    print(f"[{AutomationUtils.ts()}]   Cleaned up error screenshot: {Path(path).name}")
                except OSError:
                    pass

    @staticmethod
    def extract_employee_id_from_debug_screenshot(path: Path) -> Optional[str]:
        """Extract employee ID from ff_{label}_{employee_id}_{timestamp}.png screenshots."""
        stem = path.stem
        # Matches ff_{label}_{employee_id}_{date_time} or legacy download_ff_report_...
        match = re.search(r"(?:ff|download_ff_report)_[a-z_]+_(\d{6,10})_", stem, flags=re.IGNORECASE)
        if match:
            return match.group(1)

        match_any = re.search(r"(?<!\d)(\d{6,10})(?!\d)", stem)
        if match_any:
            return match_any.group(1)

        return None

    @staticmethod
    def debug_screenshot_path(error_dir: Path, label: str, employee_number: Optional[str] = None) -> Path:
        safe_label = re.sub(r"[^A-Za-z_]+", "_", label).strip("_").lower()
        now_str = datetime.now().strftime("%d_%b_%Y_(%I.%M%p)")
        if employee_number:
            return error_dir / f"ff_{safe_label}_{employee_number}_{now_str}.png"
        return error_dir / f"ff_{safe_label}_{now_str}.png"

    @staticmethod
    async def save_debug_screenshot(
        page: Page, error_dir: Path, label: str, employee_number: Optional[str] = None
    ) -> str:
        """Save a full-page diagnostic screenshot formatted with source script, label, employee ID, date, and time."""
        error_dir.mkdir(parents=True, exist_ok=True)
        path = AutomationUtils.debug_screenshot_path(error_dir, label, employee_number)
        try:
            await page.screenshot(path=str(path), full_page=True)
            print(f"[{AutomationUtils.ts()}]   Screenshot saved: {path.name}")
        except Exception as e:
            print(f"[{AutomationUtils.ts()}]   Could not save screenshot: {e}")
        return str(path)

    @staticmethod
    def get_failed_ids_from_error_folder(error_dir: Path) -> List[str]:
        """
        Scan errors/FF_Errors/ for debug screenshots and extract unique
        employee IDs from the filenames.
        """
        if not error_dir.exists():
            return []
        seen = set()
        ids = []
        for path in sorted(error_dir.glob("*.png")):
            emp_id = AutomationUtils.extract_employee_id_from_debug_screenshot(path)
            if emp_id and emp_id not in seen:
                seen.add(emp_id)
                ids.append(emp_id)
        return ids

    @staticmethod
    async def safe_click(locator, description: str, timeout: int = 15_000):
        """Wait for an element to be visible, scroll into view, then click it."""
        try:
            await locator.wait_for(state="visible", timeout=timeout)
            await locator.scroll_into_view_if_needed()
            await locator.click(timeout=timeout)
        except PlaywrightTimeoutError:
            raise RuntimeError(f"Could not click '{description}' within {timeout}ms")

    @staticmethod
    async def click_with_fallback(
        page_or_frame, selectors: list, description: str, timeout_each: int = 10_000
    ):
        """Try multiple selectors until one clicks successfully."""
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
    async def check_and_dismiss_error_dialog(page: Page) -> bool:
        """
        Detect and dismiss Content Server error dialogs (e.g. 'Network communication failed.').
        Returns True if an error dialog was found and dismissed.
        """
        try:
            error_selectors = [
                'div.binf-modal-dialog:has-text("Error")',
                'div.binf-modal-content:has-text("Error")',
                'div.csui-messagepanel:has-text("Error")',
                'div.binf-modal-content:has-text("Network communication failed")',
            ]
            for sel in error_selectors:
                try:
                    dialog = page.locator(sel).first
                    if await dialog.is_visible(timeout=1500):
                        error_text = (await dialog.inner_text()).strip().replace("\n", " ")[:200]
                        print(f"[{AutomationUtils.ts()}] [WARN] Content Server error dialog detected: {error_text}")

                        close_selectors = [
                            'button:has-text("Close")',
                            'button.binf-btn:has-text("Close")',
                            'a.binf-close',
                            'button.csui-btn-close',
                            'button[aria-label="Close"]',
                        ]
                        for close_sel in close_selectors:
                            try:
                                close_btn = page.locator(close_sel).first
                                if await close_btn.is_visible(timeout=2000):
                                    await close_btn.click(force=True)
                                    print(f"[{AutomationUtils.ts()}]   Dismissed error dialog")
                                    await page.wait_for_timeout(1500)
                                    return True
                            except Exception:
                                continue

                        await page.keyboard.press("Escape")
                        print(f"[{AutomationUtils.ts()}]   Dismissed error dialog via Escape key")
                        await page.wait_for_timeout(1500)
                        return True
                except Exception:
                    continue
        except Exception:
            pass
        return False


# ==== F&F Document Filter ====
class FFDocumentFilter:
    """Regex pattern and filter helpers for identifying F&F sheet documents."""

    FF_DOCUMENT_PATTERN = re.compile(
        r"(?i)"
        r"("
        r"F\s*AND\s*F"
        r"|F\s*&\s*F"
        r"|F[_\-\s]*N[_\-\s]*F"
        r"|F_AND_F"
        r"|FULL\s*AND\s*FINAL"
        r"|FULL[_\-]AND[_\-]FINAL"
        r")"
    )

    @classmethod
    def is_ff_document(cls, name: str) -> bool:
        """Return True only if the document name matches the F&F pattern."""
        return bool(cls.FF_DOCUMENT_PATTERN.search(name))


# ==== Content Server Navigator & Downloader ====
class ContentServerClient:
    """Handles page interactions, navigation, search, and downloading on Content Server."""

    def __init__(self, config: AppConfig):
        self.config = config
        self.sso_handler = MicrosoftSSOHandler(
            oracle_email=config.oracle_email,
            oracle_password=config.oracle_password,
            error_dir=config.error_dir,
        )

    async def wait_for_home(self, page: Page, timeout_s: int = 300):
        """
        Wait for Content Server home page. Handles Microsoft SSO login if redirected,
        RSA/MFA retry (up to 3 attempts), and 'Stay signed in?' prompt.
        """
        print(f"[{AutomationUtils.ts()}] Waiting for Content Server home page (up to {timeout_s}s)...")
        home_selectors = [
            "div.csui-search a",
            'a[title="Search in Content Server"]',
            "div.csui-home-item",
            ".csui-nodestable",
        ]

        login_handled = False
        rsa_retry_count = 0
        deadline = asyncio.get_event_loop().time() + timeout_s

        while asyncio.get_event_loop().time() < deadline:
            current_url = page.url.lower()
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

            # Handle RSA / SecurID Rejection
            try:
                if "securid.com" in current_url or "adani.auth" in current_url:
                    retry_btn = page.locator('button:has-text("Retry")').first
                    if await retry_btn.is_visible(timeout=1000):
                        rsa_retry_count += 1
                        print(f"[{AutomationUtils.ts()}] [WARN] RSA/MFA Authentication failed or rejected (Attempt {rsa_retry_count}/3).")
                        if rsa_retry_count >= 3:
                            await AutomationUtils.save_debug_screenshot(page, self.config.error_dir, "rsa_failed")
                            raise RuntimeError("RSA/MFA Authentication failed 3 times.")

                        print(f"[{AutomationUtils.ts()}]   Clicking 'Retry' for RSA...")
                        await retry_btn.click()
                        await page.wait_for_timeout(3000)
                        print(f"[{AutomationUtils.ts()}] Waiting for RSA/MFA approval...")
                        continue
            except RuntimeError:
                raise
            except Exception:
                pass

            # Handle "Stay signed in?" prompt
            try:
                stay_signed = page.locator('#idSIButton9, input[value="Yes"]').first
                if await stay_signed.is_visible(timeout=1000):
                    await stay_signed.click()
                    print(f"[{AutomationUtils.ts()}]   Clicked 'Stay signed in? -> Yes'")
                    await page.wait_for_timeout(3000)
                    continue
            except Exception:
                pass

            # Check for Content Server home page
            for sel in home_selectors:
                try:
                    if await page.locator(sel).first.is_visible(timeout=2000):
                        print(f"[{AutomationUtils.ts()}] Content Server home page ready")
                        return
                except Exception:
                    pass
            await page.wait_for_timeout(2000)

        await AutomationUtils.save_debug_screenshot(page, self.config.error_dir, "home_load_timeout")
        raise RuntimeError(
            f"Content Server home page did not load within {timeout_s}s.\n"
            f"Current URL: {page.url}"
        )

    async def navigate_home(self, page: Page):
        """Click the Home icon to return to the Content Server home page."""
        await AutomationUtils.click_with_fallback(
            page,
            [
                "div.csui-home-item div",
                'a[title="Home page"]',
                "div.csui-home-item a",
                'xpath=/html/body/div[2]/nav/div/div[1]/div[2]/a/div',
            ],
            "Home page icon",
            timeout_each=self.config.element_timeout,
        )
        await page.wait_for_timeout(3000)

    async def search_employee(self, page: Page, employee_number: str):
        """
        Perform a search in Content Server for the given employee number.
        Flow: Click search icon -> type in search box -> click search button.
        """
        print(f"[{AutomationUtils.ts()}] Searching for employee: {employee_number}")

        # Step 1: Click the Search icon in the navbar
        await AutomationUtils.click_with_fallback(
            page,
            [
                'a[title="Search in Content Server"]',
                "div.csui-search a",
                'xpath=/html/body/div[2]/nav/div/div[2]/div[2]/div/div[2]/a',
            ],
            "Search icon",
            timeout_each=self.config.element_timeout,
        )
        await page.wait_for_timeout(1500)

        # Step 2: Click the search input to focus it
        search_input_selectors = [
            'input[placeholder="Enter your search term"]',
            "body > div:nth-of-type(2) input",
            'xpath=/html/body/div[2]/nav/div/div[2]/div[2]/div/div[1]/div/div/input',
        ]
        search_input = None
        for sel in search_input_selectors:
            try:
                loc = page.locator(sel).first
                if await loc.is_visible(timeout=5000):
                    search_input = loc
                    break
            except Exception:
                continue

        if not search_input:
            raise RuntimeError("Could not find the search input field")

        await search_input.click()
        await page.wait_for_timeout(500)

        # Step 3: Clear any existing text and type the employee number
        await search_input.fill("")
        await page.wait_for_timeout(300)
        await search_input.fill(employee_number)
        await page.wait_for_timeout(300)

        # Press Enter on the search input to trigger search
        await search_input.press("Enter")
        await page.wait_for_timeout(500)

        # Also click the Search submit button as fallback
        try:
            search_btn = page.locator('a[title="Search in Content Server"], div.csui-search a').first
            if await search_btn.is_visible(timeout=1500):
                await search_btn.click()
        except Exception:
            pass

        # Give Content Server time to dispatch query and detach old DOM
        await page.wait_for_timeout(3500)

        # Step 5: Wait for search results to appear
        if await AutomationUtils.check_and_dismiss_error_dialog(page):
            print(f"[{AutomationUtils.ts()}]   Error dialog dismissed -- retrying search for {employee_number}")
            await self.navigate_home(page)
            await page.wait_for_timeout(2000)
            await self.search_employee(page, employee_number)
            return

        try:
            try:
                spinner = page.locator(
                    ".csui-loading-parent-wrapper, .binf-list-group-loading, .load-container, .cs-indicator"
                ).first
                await spinner.wait_for(state="hidden", timeout=self.config.search_results_wait)
            except Exception:
                pass

            await page.locator("#results, div.binf-list-group, .csui-search-results").first.wait_for(
                state="visible", timeout=self.config.search_results_wait
            )
        except PlaywrightTimeoutError:
            if await AutomationUtils.check_and_dismiss_error_dialog(page):
                print(f"[{AutomationUtils.ts()}]   Error dialog appeared during search -- skipping {employee_number}")
                raise RuntimeError(
                    f"Content Server error during search for employee {employee_number}"
                )

            try:
                zero_hits_loc = page.locator('text=/About 0 hits/i').first
                if await zero_hits_loc.is_visible(timeout=3000):
                    link_count = await page.locator(f'a:has-text("{employee_number}")').count()
                    if link_count == 0:
                        print(f"[{AutomationUtils.ts()}]   Employee {employee_number} not found on Content Server (0 hits) -- skipping")
                        raise RuntimeError(
                            f"Employee {employee_number} not found on Content Server (0 search results)"
                        )
            except RuntimeError:
                raise
            except Exception:
                pass

            await AutomationUtils.save_debug_screenshot(page, self.config.error_dir, "no_results", employee_number)
            raise RuntimeError(
                f"Search results did not appear for employee {employee_number} (timeout)"
            )

    async def click_employee_master_folder(self, page: Page, employee_number: str):
        """Finds and clicks the employee's master folder in the search results."""
        print(f"[{AutomationUtils.ts()}] Looking for exact master folder match: {employee_number}")

        for _poll in range(8):
            try:
                loc = page.get_by_role("link", name=employee_number, exact=True).first
                if await loc.is_visible(timeout=1000):
                    break
            except Exception:
                pass

            try:
                loc = page.locator(
                    f'a[title="{employee_number}"], '
                    f'a[aria-label="{employee_number}"], '
                    f'a[data-binf-original-title="{employee_number}"]'
                ).first
                if await loc.is_visible(timeout=1000):
                    break
            except Exception:
                pass

            if _poll >= 3:
                try:
                    zero_hits_loc = page.locator('text=/About 0 hits/i').first
                    if await zero_hits_loc.is_visible(timeout=1000):
                        has_link = await page.locator(f'a:has-text("{employee_number}")').count()
                        if has_link == 0:
                            raise RuntimeError(
                                f"Employee {employee_number} not found on Content Server (0 search results)"
                            )
                except RuntimeError:
                    raise
                except Exception:
                    pass

            await page.wait_for_timeout(2500)

        # Strategy 1: Exact text matching using Playwright role locator
        try:
            loc = page.get_by_role("link", name=employee_number, exact=True).first
            if await loc.is_visible(timeout=5000):
                await loc.scroll_into_view_if_needed()
                await loc.click()
                await page.wait_for_timeout(3000)
                return
        except Exception:
            pass

        # Strategy 2: Match by exact title / aria-label / data-binf-original-title
        try:
            loc = page.locator(
                f'a[title="{employee_number}"], '
                f'a[aria-label="{employee_number}"], '
                f'a[data-binf-original-title="{employee_number}"]'
            ).first
            if await loc.is_visible(timeout=3000):
                await loc.scroll_into_view_if_needed()
                await loc.click()
                await page.wait_for_timeout(3000)
                return
        except Exception:
            pass

        # Strategy 3: Iterate and check exact inner_text
        selectors = [
            ".csui-search-item a",
            "div.csui-search-item a",
            "div.binf-list-group a",
            "#results a",
        ]

        for sel in selectors:
            try:
                locators = page.locator(sel)
                count = await locators.count()
                for i in range(count):
                    loc = locators.nth(i)
                    text = await loc.inner_text()
                    if text and text.strip() == employee_number:
                        await loc.scroll_into_view_if_needed()
                        await loc.click()
                        await page.wait_for_timeout(3000)
                        return
            except Exception:
                pass

        raise RuntimeError(
            f"Employee folder '{employee_number}' not found in search results -- skipping"
        )

    async def click_adani_exit_documents(self, page: Page, employee_number: Optional[str] = None):
        """Click the 'Adani Exit Documents' folder link inside the employee folder."""
        try:
            await page.locator("td.csui-table-cell-name").first.wait_for(
                state="visible", timeout=self.config.element_timeout
            )
        except PlaywrightTimeoutError:
            await AutomationUtils.save_debug_screenshot(page, self.config.error_dir, "no_folder_table", employee_number)
            raise RuntimeError("Employee folder table did not load")

        await page.wait_for_timeout(1500)

        try:
            exact_folder = page.get_by_role("link", name="Adani Exit Documents", exact=True).first
            if await exact_folder.is_visible(timeout=5000):
                await exact_folder.scroll_into_view_if_needed()
                await exact_folder.click()
                await page.wait_for_timeout(3000)
                return
        except Exception:
            pass

        await AutomationUtils.click_with_fallback(
            page,
            [
                'a[title="Adani Exit Documents"]',
                'a[aria-label="Adani Exit Documents"]',
                'text="Adani Exit Documents"',
                'xpath=//a[text()="Adani Exit Documents"]',
                'a:has-text("Adani Exit Documents")',
                "tr.csui-state-hover > td.csui-table-cell-name a > span",
                'xpath=//span[contains(text(),"Adani Exit Documents")]',
                'xpath=//a[contains(.,"Adani Exit Documents")]',
            ],
            "Adani Exit Documents folder",
            timeout_each=self.config.element_timeout,
        )
        await page.wait_for_timeout(3000)

    async def get_document_count(self, page: Page) -> int:
        """Count the number of document rows in the current folder table."""
        try:
            await page.locator("tbody tr").first.wait_for(
                state="visible", timeout=self.config.element_timeout
            )
            await page.wait_for_timeout(1000)

            rows = page.locator("tbody tr td.csui-table-cell-name")
            count = await rows.count()
            print(f"[{AutomationUtils.ts()}] Found {count} document(s)")
            return count
        except PlaywrightTimeoutError:
            print(f"[{AutomationUtils.ts()}]   No document rows found in the folder")
            return 0

    async def download_document_at_row(
        self, page: Page, row_index: int, total_rows: int, download_dir: Path, employee_number: str
    ) -> Optional[str]:
        """Download a single F&F document from the folder table by its 1-based row index."""
        row_selector = f"tbody tr:nth-of-type({row_index})"
        try:
            row = page.locator(row_selector).first

            try:
                await row.scroll_into_view_if_needed(timeout=5000)
            except Exception:
                pass

            await row.wait_for(state="visible", timeout=self.config.element_timeout)

            row_name = f"Row {row_index}"
            name_cell = row.locator("td.csui-table-cell-name a, td.csui-table-cell-name").first
            if await name_cell.count() > 0:
                row_name = (await name_cell.inner_text()).strip()

            is_folder = False
            icon_cell = row.locator(
                "td:nth-of-type(2), td.csui-table-cell-icon, td.csui-table-cell-type, td:nth-child(2)"
            ).first
            if await icon_cell.count() > 0:
                html = await icon_cell.inner_html()
                html_lower = html.lower()
                if "folder" in html_lower or "csui-icon-node-folder" in html_lower:
                    is_folder = True

            if is_folder:
                print(f"[{AutomationUtils.ts()}] Skipping folder: '{row_name}' (row {row_index}/{total_rows})")
                return None

            if not FFDocumentFilter.is_ff_document(row_name):
                print(f"[{AutomationUtils.ts()}] Skipping non-F&F document: '{row_name}' (row {row_index}/{total_rows})")
                return None

            download_dir.mkdir(parents=True, exist_ok=True)
            print(f"[{AutomationUtils.ts()}] Downloading F&F document: '{row_name}' (row {row_index}/{total_rows})...")
            await row.hover()
            await page.wait_for_timeout(1000)
        except Exception as e:
            print(f"[{AutomationUtils.ts()}]   Could not analyze or hover row {row_index}: {e}")
            return None

        more_actions_selectors = [
            f"tbody tr:nth-of-type({row_index}) li.binf-dropdown > a",
            f'xpath=//tbody/tr[{row_index}]//li[contains(@class,"binf-dropdown")]/a',
            f'xpath=//tbody/tr[{row_index}]//a[@title="More actions"]',
        ]

        try:
            await AutomationUtils.click_with_fallback(
                page,
                more_actions_selectors,
                f"More actions (row {row_index})",
                timeout_each=8000,
            )
        except RuntimeError:
            try:
                more_btn = page.locator('a[title="More actions"]').first
                if await more_btn.is_visible(timeout=5000):
                    await more_btn.click()
                    print(f"[{AutomationUtils.ts()}]   Clicked 'More actions' via fallback")
            except Exception as e2:
                print(f"[{AutomationUtils.ts()}]   Could not click 'More actions' for row {row_index}: {e2}")
                await AutomationUtils.save_debug_screenshot(page, self.config.error_dir, "more_actions_fail", employee_number)
                return None

        await page.wait_for_timeout(1500)

        try:
            async with page.expect_download(timeout=60_000) as dl_info:
                download_clicked = False

                download_selectors = [
                    'a[title="Download"]',
                    'ul.binf-dropdown-menu a[title="Download"]',
                    'li.binf-dropdown ul.binf-dropdown-menu a[title="Download"]',
                    f"li.binf-dropdown li:nth-of-type(1) > a",
                    'xpath=//ul[contains(@class,"binf-dropdown-menu")]//a[@title="Download"]',
                ]
                for sel in download_selectors:
                    try:
                        loc = page.locator(sel).first
                        if await loc.is_visible(timeout=3000):
                            await loc.scroll_into_view_if_needed()
                            await loc.click()
                            download_clicked = True
                            break
                    except Exception:
                        continue

                if not download_clicked:
                    force_selectors = [
                        'a[title="Download"]',
                        'xpath=//a[@title="Download"]',
                        'xpath=//a[contains(.,"Download")]',
                    ]
                    for sel in force_selectors:
                        try:
                            loc = page.locator(sel).first
                            if await loc.count() > 0:
                                await loc.click(force=True, timeout=8000)
                                download_clicked = True
                                print(f"[{AutomationUtils.ts()}]   Used force-click on Download button")
                                break
                        except Exception:
                            continue

                if not download_clicked:
                    try:
                        await page.evaluate("""
                            () => {
                                const link = document.querySelector('a[title="Download"]');
                                if (link) { link.click(); return true; }
                                return false;
                            }
                        """)
                        download_clicked = True
                        print(f"[{AutomationUtils.ts()}]   Used JavaScript click on Download button")
                    except Exception:
                        pass

                if not download_clicked:
                    raise RuntimeError(
                        f"Could not click Download for row {row_index} after all attempts"
                    )

            download = await dl_info.value
            failure = await download.failure()
            if failure:
                print(f"[{AutomationUtils.ts()}]   Download failed for row {row_index}: {failure}")
                return None

            suggested = download.suggested_filename or f"ff_document_{row_index}"
            suggested_clean = suggested.strip()
            match = re.search(r"\.\s*(\d+)$", suggested_clean)
            if match:
                suffix_num = match.group(1)
                base = suggested_clean[: match.start()].strip().rstrip(".")
                suggested_clean = f"{base}_{suffix_num}.pdf"
            else:
                ext = Path(suggested_clean).suffix.lower()
                if ext not in [".pdf", ".xlsx", ".xls", ".doc", ".docx", ".png", ".jpg"]:
                    suggested_clean = suggested_clean.rstrip(".")
                    suggested_clean = f"{suggested_clean}.pdf"

            save_path = download_dir / suggested_clean

            if save_path.exists():
                stem = save_path.stem
                suffix = save_path.suffix
                counter = 1
                while save_path.exists():
                    save_path = download_dir / f"{stem}_{counter}{suffix}"
                    counter += 1

            await download.save_as(save_path)
            size = save_path.stat().st_size
            print(f"[{AutomationUtils.ts()}] Saved: {save_path.name} ({size:,} bytes)")
            return str(save_path)

        except PlaywrightTimeoutError:
            print(f"[{AutomationUtils.ts()}]   Download timed out for row {row_index}")
            await AutomationUtils.save_debug_screenshot(page, self.config.error_dir, "download_timeout", employee_number)
            return None
        except Exception as e:
            print(f"[{AutomationUtils.ts()}]   Error downloading row {row_index}: {e}")
            return None

    async def download_ff_for_employee(
        self, page: Page, employee_number: str, download_dir: Path
    ) -> Tuple[str, List[str]]:
        """Full flow to download all F&F documents for one employee."""
        downloaded_files = []
        status = STATUS_ERROR

        try:
            # Step 1: Search
            await self.search_employee(page, employee_number)

            # Step 2: Click the exact employee folder matching the ID
            await self.click_employee_master_folder(page, employee_number)

            # Step 3: Navigate into "Adani Exit Documents"
            try:
                await self.click_adani_exit_documents(page, employee_number)
            except RuntimeError as e:
                if "Employee folder table did not load" in str(e) or "Adani Exit Documents" in str(e):
                    print(f"[{AutomationUtils.ts()}] [WARN] Failed to load folder table or find Adani Exit Documents. Clicking Home and retrying once...")
                    await self.navigate_home(page)
                    await page.wait_for_timeout(3000)
                    await self.search_employee(page, employee_number)
                    await self.click_employee_master_folder(page, employee_number)
                    await self.click_adani_exit_documents(page, employee_number)
                else:
                    raise

            # Step 4: Get document count and download each
            doc_count = await self.get_document_count(page)

            if doc_count == 0:
                print(f"[{AutomationUtils.ts()}]   No documents found for employee {employee_number}")
                status = STATUS_NO_FF_DOCS
            else:
                emp_dir = download_dir / employee_number

                for row_idx in range(1, doc_count + 1):
                    await AutomationUtils.check_and_dismiss_error_dialog(page)

                    result = await self.download_document_at_row(
                        page, row_idx, doc_count, emp_dir, employee_number
                    )
                    if result:
                        downloaded_files.append(result)
                    await page.wait_for_timeout(2000)

                if downloaded_files:
                    status = STATUS_DOWNLOADED
                else:
                    status = STATUS_NO_FF_DOCS

            AutomationUtils.cleanup_error_screenshots(self.config.error_dir, employee_number)

        except RuntimeError as e:
            err_msg = str(e)
            if "0 search results" in err_msg or "0 hits" in err_msg:
                print(f"[{AutomationUtils.ts()}] Employee {employee_number} not found on Content Server (0 hits)")
                status = STATUS_NOT_FOUND
                AutomationUtils.cleanup_error_screenshots(self.config.error_dir, employee_number)
            else:
                print(f"[{AutomationUtils.ts()}] Error processing employee {employee_number}: {e}")
                await AutomationUtils.save_debug_screenshot(page, self.config.error_dir, "error", employee_number)
                status = STATUS_ERROR
        except Exception as e:
            print(f"[{AutomationUtils.ts()}] Error processing employee {employee_number}: {e}")
            await AutomationUtils.save_debug_screenshot(page, self.config.error_dir, "error", employee_number)
            status = STATUS_ERROR

        return status, downloaded_files


# ==== F&F Document Orchestrator ====
class FFReportDownloader:
    """Orchestrates browser setup, batch downloads, retries, and reporting."""

    def __init__(self, config: Optional[AppConfig] = None):
        self.config = config or AppConfig()
        self.client = ContentServerClient(self.config)

    @staticmethod
    def prompt_employee_numbers() -> List[str]:
        """Interactive terminal prompt -- collects employee ID(s) from the user."""
        terminal = sys.__stdout__

        terminal.write("\n")
        terminal.write("  ╔═══════════════════════════════════════════════════════╗\n")
        terminal.write("  ║        F&F Exit Document Downloader                  ║\n")
        terminal.write("  ╚═══════════════════════════════════════════════════════╝\n")
        terminal.write("\n")
        terminal.flush()

        try:
            terminal.write("  Enter Employee ID (or multiple comma-separated): ")
            terminal.flush()
            raw = sys.__stdin__.readline()
        except (EOFError, KeyboardInterrupt):
            terminal.write("\n")
            return []

        raw = raw.strip()
        if not raw:
            terminal.write("\n  -> No employee numbers entered.\n\n")
            terminal.flush()
            return []

        collected = []
        for token in raw.replace(",", " ").split():
            token = token.strip()
            if token:
                collected.append(token)

        seen = set()
        unique = []
        for num in collected:
            if num not in seen:
                seen.add(num)
                unique.append(num)

        if unique:
            terminal.write(f"\n  -> {len(unique)} employee(s) queued: {', '.join(unique)}\n\n")
            terminal.flush()
        else:
            terminal.write("\n  -> No employee numbers entered.\n\n")
            terminal.flush()

        return unique

    @staticmethod
    def print_categorized_summary(all_results: Dict[str, Tuple[str, List[str]]]):
        """Print a clear, categorized download summary."""
        downloaded = {k: v for k, (s, v) in all_results.items() if s == STATUS_DOWNLOADED}
        no_ff = {k: v for k, (s, v) in all_results.items() if s == STATUS_NO_FF_DOCS}
        not_found = {k: v for k, (s, v) in all_results.items() if s == STATUS_NOT_FOUND}
        errors = {k: v for k, (s, v) in all_results.items() if s == STATUS_ERROR}

        total_files = sum(len(v) for v in downloaded.values())
        total_emps = len(all_results)

        print()
        print(f"[{AutomationUtils.ts()}] ═══════════════════════════════════════════════════════════")
        print(f"[{AutomationUtils.ts()}] DOWNLOAD SUMMARY")
        print(f"[{AutomationUtils.ts()}] ═══════════════════════════════════════════════════════════")

        if downloaded:
            print(f"[{AutomationUtils.ts()}]")
            print(f"[{AutomationUtils.ts()}]   [OK] DOWNLOADED ({len(downloaded)} employee(s), {total_files} file(s)):")
            for emp, files in downloaded.items():
                print(f"[{AutomationUtils.ts()}]      {emp}: {len(files)} file(s)")

        if no_ff:
            print(f"[{AutomationUtils.ts()}]")
            print(f"[{AutomationUtils.ts()}]   [INFO] NO F&F DOCUMENT ({len(no_ff)} employee(s) -- folder found, no F&F sheet uploaded):")
            for emp in no_ff:
                print(f"[{AutomationUtils.ts()}]      {emp}")

        if not_found:
            print(f"[{AutomationUtils.ts()}]")
            print(f"[{AutomationUtils.ts()}]   [--] NOT FOUND ON CONTENT SERVER ({len(not_found)} employee(s) -- 0 search hits):")
            for emp in not_found:
                print(f"[{AutomationUtils.ts()}]      {emp}")

        if errors:
            print(f"[{AutomationUtils.ts()}]")
            print(f"[{AutomationUtils.ts()}]   [FAIL] FAILED / ERRORS ({len(errors)} employee(s) -- search timeout or network failure):")
            for emp in errors:
                print(f"[{AutomationUtils.ts()}]      {emp}")

        print(f"[{AutomationUtils.ts()}] ───────────────────────────────────────────────────────────")
        print(f"[{AutomationUtils.ts()}]   Total: {total_files} file(s) downloaded for {total_emps} employee(s)")
        if errors:
            print(f"[{AutomationUtils.ts()}]   [WARN] {len(errors)} employee(s) still have unresolved errors -- check errors/ff_error/ for screenshots")
        print(f"[{AutomationUtils.ts()}] ═══════════════════════════════════════════════════════════")

    async def _run_batch(
        self,
        p_context: BrowserContext,
        emp_list: List[str],
        num_tabs: int,
        nav_lock: asyncio.Lock,
        total_label: int,
        all_results: Dict[str, Tuple[str, List[str]]],
        start_offset: int = 0,
        upload_cb=None,
    ):
        """Process a list of employees using the given browser context and tab count."""
        first_page = p_context.pages[0] if p_context.pages else await p_context.new_page()
        worker_pages = [first_page]
        for i in range(1, num_tabs):
            try:
                new_tab = await p_context.new_page()
                worker_pages.append(new_tab)
                print(f"[{AutomationUtils.ts()}]   Opened tab {i + 1}/{num_tabs}")
                await asyncio.sleep(0.5)
            except Exception as e:
                print(f"[{AutomationUtils.ts()}]   Could not open tab {i + 1}: {e}")
                break

        actual_tabs = len(worker_pages)
        print(f"[{AutomationUtils.ts()}] {actual_tabs} tab(s) ready. Starting downloads...\n")

        async def _worker(worker_page: Page, emp_num: str, idx: int):
            try:
                print(f"[{AutomationUtils.ts()}] ── [Tab] Employee {idx}/{total_label}: {emp_num} ──")

                nav_success = False
                for nav_attempt in range(2):
                    try:
                        async with nav_lock:
                            await worker_page.goto(
                                self.config.content_server_url,
                                wait_until="domcontentloaded",
                                timeout=self.config.nav_timeout,
                            )
                            await worker_page.wait_for_timeout(2000)
                            await self.client.wait_for_home(worker_page, timeout_s=30)
                        nav_success = True
                        break
                    except Exception as nav_err:
                        if nav_attempt == 0:
                            print(f"[{AutomationUtils.ts()}]   [Tab] Navigation timeout for {emp_num} -- retrying once...")
                            await worker_page.wait_for_timeout(3000)
                        else:
                            raise RuntimeError(
                                f"Content Server navigation failed after 2 attempts for {emp_num}: {nav_err}"
                            )

                if not nav_success:
                    raise RuntimeError(f"Could not navigate to Content Server for {emp_num}")

                status, downloaded = await self.client.download_ff_for_employee(
                    worker_page, emp_num, self.config.download_dir
                )
                all_results[emp_num] = (status, downloaded)

                if downloaded and upload_cb:
                    try:
                        print(f"[{AutomationUtils.ts()}] [Tab] Triggering upload for {emp_num}...")
                        await upload_cb(emp_num, downloaded)
                    except Exception as e:
                        print(f"[{AutomationUtils.ts()}] [Tab] Upload callback failed for {emp_num}: {e}")

            except Exception as e:
                print(f"[{AutomationUtils.ts()}] [Tab] Error for {emp_num}: {e}")
                await AutomationUtils.save_debug_screenshot(worker_page, self.config.error_dir, "error", emp_num)
                all_results[emp_num] = (STATUS_ERROR, [])

        for batch_start in range(0, len(emp_list), actual_tabs):
            batch = emp_list[batch_start : batch_start + actual_tabs]
            batch_tasks = []
            for i, emp_num in enumerate(batch):
                page = worker_pages[i]
                idx = start_offset + batch_start + i + 1
                batch_tasks.append(_worker(page, emp_num, idx))

            await asyncio.gather(*batch_tasks)
            await asyncio.sleep(1.5)

        for wp in worker_pages[1:]:
            try:
                await wp.close()
            except Exception:
                pass

    async def download_reports(
        self, employee_numbers: List[str], upload_callback=None
    ) -> Dict[str, Tuple[str, List[str]]]:
        """
        Main runner. Downloads F&F exit documents for one or more employees.
        Uses concurrent browser tabs for parallel processing and retries failed records.
        """
        self.config.download_dir.mkdir(parents=True, exist_ok=True)

        if not employee_numbers:
            print(f"[{AutomationUtils.ts()}] No employee numbers provided. Exiting.")
            return {}

        if not self.config.content_server_url:
            print(f"[{AutomationUtils.ts()}] [FATAL ERROR] CONTENT_SERVER_URL is completely blank in your .env file!")
            print(f"[{AutomationUtils.ts()}] Please open the .env file and add the URL like: CONTENT_SERVER_URL=https://...")
            return {}

        concurrency = min(self.config.max_concurrent_tabs, len(employee_numbers))
        print(f"[{AutomationUtils.ts()}] Starting F&F report download ({'headless' if self.config.headless else 'visible'})")
        print(f"[{AutomationUtils.ts()}] Concurrency: {concurrency} tab(s) for {len(employee_numbers)} employee(s)")

        AutomationUtils.cleanup_profile_locks(self.config.automation_profile_dir)

        all_results: Dict[str, Tuple[str, List[str]]] = {}

        try:
            async with async_playwright() as p:
                print(f"[{AutomationUtils.ts()}] Launching browser...")

                # Build Chromium args cleanly based on START_MAXIMIZED and HEADLESS
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

                # Step 1: Establish login / SSO on the first tab
                first_page = context.pages[0] if context.pages else await context.new_page()

                print(f"[{AutomationUtils.ts()}] Navigating to Content Server...")
                await first_page.goto(
                    self.config.content_server_url,
                    wait_until="domcontentloaded",
                    timeout=self.config.nav_timeout,
                )
                await first_page.wait_for_timeout(3000)

                await self.client.wait_for_home(first_page, timeout_s=120)
                await first_page.wait_for_timeout(2000)
                print(f"[{AutomationUtils.ts()}] Login confirmed. Preparing {concurrency} tab(s)...\n")

                nav_lock = asyncio.Lock()

                # Step 2: Primary batch
                await self._run_batch(
                    context,
                    employee_numbers,
                    concurrency,
                    nav_lock,
                    len(employee_numbers),
                    all_results,
                    upload_cb=upload_callback,
                )

                # Step 3: Retry loop
                MAX_RETRY_ROUNDS = 5
                initial_not_found = [
                    emp for emp, (st, _) in all_results.items() if st == STATUS_NOT_FOUND
                ]

                for retry_round in range(1, MAX_RETRY_ROUNDS + 1):
                    folder_failed_ids = AutomationUtils.get_failed_ids_from_error_folder(
                        self.config.error_dir
                    )

                    if retry_round == 1 and initial_not_found:
                        combined_ids = list(folder_failed_ids)
                        for nf_id in initial_not_found:
                            if nf_id not in combined_ids:
                                combined_ids.append(nf_id)
                        current_retry_ids = combined_ids
                    else:
                        current_retry_ids = folder_failed_ids

                    if not current_retry_ids:
                        print(f"[{AutomationUtils.ts()}] errors/ff_error/ is empty -- no retries needed.")
                        break

                    retry_concurrency = min(2, len(current_retry_ids))
                    print()
                    print(f"[{AutomationUtils.ts()}] ═══════════════════════════════════════════════════════════")
                    print(f"[{AutomationUtils.ts()}] RETRY ROUND {retry_round}/{MAX_RETRY_ROUNDS}: {len(current_retry_ids)} employee(s) queued for retry")
                    print(f"[{AutomationUtils.ts()}]   IDs: {', '.join(current_retry_ids)}")
                    print(f"[{AutomationUtils.ts()}]   Retrying with {retry_concurrency} tab(s) for reliability...")
                    print(f"[{AutomationUtils.ts()}] ═══════════════════════════════════════════════════════════")
                    print()

                    try:
                        remaining_pages = context.pages
                        if remaining_pages:
                            retry_page = remaining_pages[0]
                            await retry_page.goto(
                                self.config.content_server_url,
                                wait_until="domcontentloaded",
                                timeout=self.config.nav_timeout,
                            )
                            await retry_page.wait_for_timeout(2000)
                            await self.client.wait_for_home(retry_page, timeout_s=30)
                    except Exception:
                        pass

                    await self._run_batch(
                        context,
                        current_retry_ids,
                        retry_concurrency,
                        nav_lock,
                        len(current_retry_ids),
                        all_results,
                        upload_cb=upload_callback,
                    )

                    resolved_count = 0
                    for emp_num in current_retry_ids:
                        retry_status, _ = all_results.get(emp_num, (STATUS_ERROR, []))
                        if retry_status != STATUS_ERROR:
                            print(f"[{AutomationUtils.ts()}]   [OK] {emp_num} resolved on retry (status: {retry_status})")
                            AutomationUtils.cleanup_error_screenshots(self.config.error_dir, emp_num)
                            resolved_count += 1

                    print(f"[{AutomationUtils.ts()}]   Round {retry_round} complete: {resolved_count}/{len(current_retry_ids)} resolved")

                    remaining = AutomationUtils.get_failed_ids_from_error_folder(
                        self.config.error_dir
                    )
                    if not remaining:
                        print(f"[{AutomationUtils.ts()}] errors/ff_error/ is now empty -- all retries succeeded!")
                        break
                    elif retry_round < MAX_RETRY_ROUNDS:
                        print(f"[{AutomationUtils.ts()}]   {len(remaining)} still remaining, will retry next round...")
                    else:
                        print(f"[{AutomationUtils.ts()}]   {len(remaining)} still remaining after {MAX_RETRY_ROUNDS} retry rounds.")

                for pg in context.pages:
                    try:
                        await pg.close()
                    except Exception:
                        pass
                try:
                    await context.close()
                except Exception:
                    pass

        except Exception as e:
            print(f"[{AutomationUtils.ts()}] FATAL: {e}\n{traceback.format_exc()}")

        self.print_categorized_summary(all_results)
        return all_results


# ==== Main Entry Point ====
prompt_employee_numbers = FFReportDownloader.prompt_employee_numbers
_print_categorized_summary = FFReportDownloader.print_categorized_summary


async def download_ff_reports(employee_numbers: List[str], upload_callback=None):
    """Entry point for scripts.process_fnf_closed_reports and external tools."""
    setup_logger("download_ff_report.log")
    downloader = FFReportDownloader()

    # If no custom callback provided, upload downloaded files directly to SharePoint
    if upload_callback is None:
        async def default_sp_upload(emp_id: str, downloaded_paths: list):
            try:
                async with sp_uploader.httpx.AsyncClient(verify=False) as client:
                    site_id = await sp_uploader.get_site_id(client)
                    drive_id, folder_path = await sp_uploader.get_drive_details(client, site_id)
                    emp_target_url = f"{folder_path}/F&F_Documents/{emp_id}"
                    for file_path_str in downloaded_paths:
                        f_path = Path(file_path_str)
                        if f_path.exists():
                            await sp_uploader.upload_file(client, site_id, drive_id, emp_target_url, f_path)
                            print(f"[{AutomationUtils.ts()}] [SP-SYNC] Uploaded {f_path.name} directly to SharePoint for {emp_id}")
                            try:
                                f_path.unlink()
                                print(f"[{AutomationUtils.ts()}] [CLEANUP] Deleted local file: {f_path.name}")
                            except Exception as del_err:
                                print(f"[{AutomationUtils.ts()}] [CLEANUP] Warning: Could not delete {f_path.name}: {del_err}")

                    # Clean up employee directory if empty
                    emp_dir = FF_REPORTS_DIR / emp_id
                    if emp_dir.exists() and emp_dir.is_dir():
                        try:
                            emp_dir.rmdir()
                        except Exception:
                            pass
            except Exception as e:
                print(f"[{AutomationUtils.ts()}] [WARNING] SharePoint direct upload failed for {emp_id}: {e}")

        upload_callback = default_sp_upload

    return await downloader.download_reports(employee_numbers, upload_callback=upload_callback)


if __name__ == "__main__":
    numbers = prompt_employee_numbers()
    if numbers:
        asyncio.run(download_ff_reports(numbers))
    else:
        print("No employee numbers entered. Exiting.")
