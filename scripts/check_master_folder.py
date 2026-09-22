# OpenText Content Server - Employee Master Folder Existence Checker
# Searches each employee ID on Content Server and checks whether the
# employee master folder (named exactly as the employee ID) is visible in search results.

import asyncio
import os
import re
import subprocess
import sys
import traceback
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
from dotenv import load_dotenv
from playwright.async_api import Page, TimeoutError as PlaywrightTimeoutError, async_playwright

# Suppress Chrome subprocess console window on Windows
if sys.platform == "win32":
    _original_popen_init = subprocess.Popen.__init__

    def _patched_popen_init(self, *args, **kwargs):
        if "creationflags" not in kwargs:
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        else:
            kwargs["creationflags"] |= subprocess.CREATE_NO_WINDOW
        _original_popen_init(self, *args, **kwargs)

    subprocess.Popen.__init__ = _patched_popen_init


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


# Initialize logging directory and writers
LOG_DIR = Path(__file__).parent.parent / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "check_master_folder.log"

sys.stdout = LoggerWriter(LOG_FILE)
sys.stderr = LoggerWriter(LOG_FILE)

load_dotenv(verbose=True)


# Application Configuration
class MasterFolderConfig:
    """Manages credentials, paths, timeouts, and browser execution parameters."""

    def __init__(self):
        try:
            # Credentials & Target URL
            self.oracle_email: str = (os.getenv("ORACLE_EMAIL") or "").strip()
            self.oracle_password: str = (os.getenv("ORACLE_PASSWORD") or "").strip()
            self.content_server_url: str = (os.getenv("CONTENT_SERVER_URL") or "").strip()

            # Output Paths & Directories
            self.output_dir: Path = Path(__file__).parent.parent / "uploads" / "Master_Folder_Check"
            self.output_excel_path: Path = self.output_dir / "Master_Folder_Check_Report.xlsx"
            self.automation_profile_dir: Path = Path(__file__).parent.parent / "chrome_automation_profile"
            self.error_screenshot_dir: Path = Path(__file__).parent.parent / "errors" / "master_folder_check_error"

            # Browser Runtime Options
            self.headless: bool = self._parse_bool(os.getenv("HEADLESS"), default=True)
            self.start_maximized: bool = self._parse_bool(os.getenv("START_MAXIMIZED"), default=True)
            self.window_position: str = (os.getenv("WINDOW_POSITION") or "").strip()
            self.browser_channel: Optional[str] = (os.getenv("BROWSER_CHANNEL") or "").strip() or None

            # Timeouts & Retry Limits
            self.element_timeout: int = 15_000         # ms
            self.nav_timeout: int = 90_000             # ms
            self.search_results_wait: int = 45_000     # ms
            self.max_concurrent_tabs: int = 3
            self.max_retries_per_employee: int = 3

            # Status Constants
            self.status_yes: str = "Yes"
            self.status_no: str = "No"

            # Create output directories
            self.output_dir.mkdir(parents=True, exist_ok=True)
            self.error_screenshot_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Configuration initialization error: {e}")

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
class MasterFolderUtils:
    """Utility helper methods for timestamps, file locking, screenshots, and dialogs."""

    @staticmethod
    def ts() -> str:
        """Returns the current formatted timestamp string."""
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    @staticmethod
    def cleanup_profile_locks(profile_dir: Path):
        """Removes leftover Chrome lock files from the automation profile."""
        try:
            if not profile_dir.exists():
                return
            for name in ["SingletonLock", "SingletonSocket", "SingletonCookie", "lockfile"]:
                p = profile_dir / name
                try:
                    if p.exists():
                        p.unlink()
                except OSError:
                    pass
        except Exception as e:
            print(f"[{MasterFolderUtils.ts()}] [CLEANUP] Lock cleanup warning: {e}")

    @staticmethod
    async def save_debug_screenshot(
        page: Page, error_dir: Path, label: str, employee_number: Optional[str] = None
    ) -> str:
        """Captures a full-page diagnostic screenshot upon errors."""
        try:
            error_dir.mkdir(parents=True, exist_ok=True)
            safe = re.sub(r"[^A-Za-z_]+", "_", label).strip("_").lower()
            if employee_number:
                clean = re.sub(r"\D", "", str(employee_number))
                path = error_dir / f"debug_{safe}_{clean}.png"
            else:
                path = error_dir / f"debug_{safe}.png"

            await page.screenshot(path=str(path), full_page=True)
            print(f"[{MasterFolderUtils.ts()}]   Screenshot: {path.name}")
            return str(path)
        except Exception as e:
            print(f"[{MasterFolderUtils.ts()}]   Screenshot failed: {e}")
            return ""

    @staticmethod
    async def dismiss_error_dialog(page: Page) -> bool:
        """Detects and dismisses Content Server error dialogs. Returns True if one was found."""
        try:
            dialog_selectors = [
                'div.binf-modal-dialog:has-text("Error")',
                'div.binf-modal-content:has-text("Error")',
                'div.csui-messagepanel:has-text("Error")',
                'div.binf-modal-content:has-text("Network communication failed")',
            ]
            for sel in dialog_selectors:
                try:
                    dlg = page.locator(sel).first
                    if await dlg.is_visible(timeout=1500):
                        txt = (await dlg.inner_text()).strip().replace("\n", " ")[:200]
                        print(f"[{MasterFolderUtils.ts()}] [WARN] CS error dialog: {txt}")

                        close_selectors = [
                            'button:has-text("Close")',
                            "a.binf-close",
                            'button[aria-label="Close"]',
                        ]
                        for csel in close_selectors:
                            try:
                                btn = page.locator(csel).first
                                if await btn.is_visible(timeout=2000):
                                    await btn.click(force=True)
                                    await page.wait_for_timeout(1500)
                                    return True
                            except Exception:
                                continue

                        await page.keyboard.press("Escape")
                        await page.wait_for_timeout(1500)
                        return True
                except Exception:
                    continue
        except Exception:
            pass
        return False


# Microsoft SSO Handler
class MicrosoftSSOHandler:
    """Manages Microsoft Entra / SSO authentication state machine, KMSI, and MFA approvals."""

    @staticmethod
    async def check_sso_errors(page: Page):
        """Inspects the page for SSO validation errors and raises a RuntimeError if found."""
        for eid in ["usernameError", "passwordError", "error"]:
            try:
                loc = page.locator(f"#{eid}").first
                if await loc.is_visible(timeout=1000):
                    txt = (await loc.inner_text()).strip().replace("\n", " ")
                    if txt:
                        raise RuntimeError(f"SSO Error: {txt}")
            except RuntimeError:
                raise
            except Exception:
                pass

        try:
            loc = page.locator('[role="alert"]').first
            if await loc.is_visible(timeout=1000):
                txt = (await loc.inner_text()).strip().replace("\n", " ")
                if txt:
                    raise RuntimeError(f"SSO Alert: {txt}")
        except RuntimeError:
            raise
        except Exception:
            pass

    @staticmethod
    async def handle_stay_signed_in(page: Page) -> bool:
        """
        Handles Microsoft Entra 'Stay signed in?' (KMSI) prompt.
        Ensures 'Don't show this again' is checked so Microsoft issues a persistent session token.
        """
        try:
            kmsi_header = page.locator(
                '#loginHeader:has-text("Stay signed in?"), '
                'div[role="heading"]:has-text("Stay signed in?"), '
                'div:has-text("Stay signed in?"), '
                'div:has-text("Do this to reduce the number of times")'
            ).first
            kmsi_box = page.locator(
                '#KmsiCheckboxField, input[name="DontShowAgain"], input[type="checkbox"]'
            ).first

            if not (await kmsi_header.count() > 0 and await kmsi_header.is_visible(timeout=500)) and not (
                await kmsi_box.count() > 0 and await kmsi_box.is_visible(timeout=500)
            ):
                return False

            if await kmsi_box.count() > 0 and await kmsi_box.is_visible(timeout=1000):
                try:
                    if not await kmsi_box.is_checked():
                        await kmsi_box.check()
                        print(f"[{MasterFolderUtils.ts()}]   Checked 'Don't show this again' (KMSI)")
                        await page.wait_for_timeout(300)
                except Exception:
                    pass

            stay_signed_btn = page.locator(
                '#idSIButton9, input[value="Yes"], input[type="submit"][value="Yes"], button:has-text("Yes")'
            ).first
            if await stay_signed_btn.is_visible(timeout=2000):
                await stay_signed_btn.click(force=True, timeout=5000)
                print(f"[{MasterFolderUtils.ts()}]   Clicked 'Stay signed in? -> Yes'")
                await page.wait_for_timeout(3000)
                return True
        except Exception:
            pass
        return False

    @staticmethod
    async def handle_rsa_entra_selection(page: Page) -> bool:
        """If Microsoft presents the 'Verify your identity' screen, select RSAEntraMFA."""
        try:
            rsa_btn = page.locator('div[role="button"], button, a').filter(has_text="RSAEntraMFA").first
            if await rsa_btn.is_visible(timeout=1000):
                print(f"[{MasterFolderUtils.ts()}]   'Verify your identity' screen detected.")
                await rsa_btn.click(force=True, timeout=5000)
                print(f"[{MasterFolderUtils.ts()}]   Clicked 'Approve with RSAEntraMFA'")
                await page.wait_for_timeout(3000)
                return True
        except Exception:
            pass
        return False

    @staticmethod
    async def handle_microsoft_sso(page: Page, email: str, password: str) -> bool:
        """
        Executes automated Microsoft SSO credential input:
          1. Pick account or enter email -> click Next
          2. Enter password -> click Sign In
          3. Select RSAEntraMFA if prompt appears
          4. Handle 'Stay signed in?' (KMSI) prompt
          5. Wait for phone push notification approval
        """
        if not email or not password:
            print(f"[{MasterFolderUtils.ts()}] [WARN] No credentials in .env -- waiting for manual login...")
            return False

        print(f"[{MasterFolderUtils.ts()}] SSO login detected -- filling credentials automatically...")
        await page.wait_for_timeout(2000)
        await MicrosoftSSOHandler.check_sso_errors(page)

        email_filled = False
        password_filled = False
        rsa_clicked = False

        for _ in range(25):
            url = page.url.lower()
            if not any(x in url for x in ["login.microsoftonline", "login.microsoft", "adfs", "saml"]):
                print(f"[{MasterFolderUtils.ts()}]   SSO completed.")
                return True

            # Check for Stay signed in (KMSI) prompt
            if await MicrosoftSSOHandler.handle_stay_signed_in(page):
                continue

            # Check for RSAEntraMFA selection button
            if not rsa_clicked and await MicrosoftSSOHandler.handle_rsa_entra_selection(page):
                rsa_clicked = True
                break

            # Pick an account screen
            if not email_filled and not password_filled:
                try:
                    hdr = page.locator(
                        'div#loginHeader:has-text("Pick an account"), '
                        'div[role="heading"]:has-text("Pick an account")'
                    ).first
                    if await hdr.is_visible(timeout=1000):
                        target = email.strip().lower()
                        found = False
                        account_selectors = [
                            'div.table[role="button"]',
                            'div[data-test-id]',
                            'div.row.tile',
                            'li.table',
                        ]
                        for sel in account_selectors:
                            try:
                                tiles = page.locator(sel)
                                for i in range(await tiles.count()):
                                    tile = tiles.nth(i)
                                    if target in (await tile.inner_text()).lower():
                                        await tile.click(force=True, timeout=5000)
                                        found = True
                                        break
                                if found:
                                    break
                            except Exception:
                                continue

                        if not found:
                            try:
                                ua = page.locator(
                                    'div#otherTile, div:has-text("Use another account")'
                                ).last
                                if await ua.is_visible(timeout=3000):
                                    await ua.click(force=True, timeout=5000)
                            except Exception:
                                pass
                        await page.wait_for_timeout(3000)
                        continue
                except Exception:
                    pass

            # Password screen
            if not password_filled:
                try:
                    pw = page.locator(
                        'input[name="passwd"], input[id="i0118"], input[type="password"]'
                    ).first
                    ind = page.locator(
                        'input[type="submit"][value="Sign in"], #idSIButton9, '
                        'div:has-text("Enter password"), #identityBanner'
                    ).first
                    if await pw.is_visible(timeout=1000) and (
                        email_filled or await ind.is_visible(timeout=1000)
                    ):
                        print(f"[{MasterFolderUtils.ts()}]   Password screen -- filling...")
                        for attempt in range(1, 4):
                            await pw.click(force=True, timeout=3000)
                            await pw.fill("", force=True, timeout=3000)
                            await page.wait_for_timeout(200)
                            await pw.type(password.strip(), delay=50)
                            await page.wait_for_timeout(500)
                            await page.locator(
                                'input[type="submit"], #idSIButton9'
                            ).first.click(force=True, timeout=5000)
                            await page.wait_for_timeout(3000)
                            try:
                                await MicrosoftSSOHandler.check_sso_errors(page)
                                password_filled = email_filled = True
                                break
                            except RuntimeError as e:
                                if attempt == 3:
                                    raise
                                print(f"[{MasterFolderUtils.ts()}]   {e} - Retrying...")
                        continue
                except RuntimeError:
                    raise
                except Exception:
                    pass

            # Email screen
            if not email_filled and not password_filled:
                try:
                    em = page.locator(
                        'input[name="loginfmt"], input[type="email"], input[id="i0116"]'
                    ).first
                    if await em.is_visible(timeout=1000):
                        print(f"[{MasterFolderUtils.ts()}]   Email screen -- filling...")
                        for attempt in range(1, 4):
                            await em.click(force=True, timeout=3000)
                            await em.fill("", force=True, timeout=3000)
                            await page.wait_for_timeout(200)
                            await em.type(email.strip(), delay=50)
                            await page.wait_for_timeout(500)
                            await page.locator(
                                'input[type="submit"], #idSIButton9'
                            ).first.click(force=True, timeout=5000)
                            await page.wait_for_timeout(3000)
                            try:
                                await MicrosoftSSOHandler.check_sso_errors(page)
                                email_filled = True
                                break
                            except RuntimeError as e:
                                if attempt == 3:
                                    raise
                                print(f"[{MasterFolderUtils.ts()}]   {e} - Retrying...")
                        continue
                except RuntimeError:
                    raise
                except Exception:
                    pass

            # Check if post-password screens appear
            if password_filled:
                if not rsa_clicked and await MicrosoftSSOHandler.handle_rsa_entra_selection(page):
                    rsa_clicked = True
                    break
                if await MicrosoftSSOHandler.handle_stay_signed_in(page):
                    continue
                if any(x in url for x in ["securid.com", "adani.auth"]) or rsa_clicked:
                    break

            await page.wait_for_timeout(1000)

        print(f"[{MasterFolderUtils.ts()}] Waiting for RSA/MFA approval (push notification sent to phone)...")
        return True


# Master Folder Existence Checker Core Class
class MasterFolderChecker:
    """Orchestrates OpenText Content Server navigation, searches, existence checks, and Excel generation."""

    def __init__(self, config: MasterFolderConfig):
        self.config = config

    async def wait_for_home(self, page: Page, timeout_s: int = 300):
        """Waits for Content Server home page, handling SSO and MFA redirects as needed."""
        try:
            print(f"[{MasterFolderUtils.ts()}] Waiting for Content Server home (up to {timeout_s}s)...")
            home_sels = [
                "div.csui-search a",
                'a[title="Search in Content Server"]',
                "div.csui-home-item",
                ".csui-nodestable",
            ]
            login_handled = False
            rsa_retries = 0
            deadline = asyncio.get_event_loop().time() + timeout_s

            while asyncio.get_event_loop().time() < deadline:
                url = page.url.lower()

                if any(x in url for x in ["login.microsoftonline", "login.microsoft", "adfs", "saml"]):
                    if await MicrosoftSSOHandler.handle_stay_signed_in(page):
                        continue
                    if await MicrosoftSSOHandler.handle_rsa_entra_selection(page):
                        continue
                    if not login_handled:
                        login_handled = await MicrosoftSSOHandler.handle_microsoft_sso(
                            page, self.config.oracle_email, self.config.oracle_password
                        )

                try:
                    if "securid.com" in url or "adani.auth" in url:
                        retry_btn = page.locator('button:has-text("Retry")').first
                        if await retry_btn.is_visible(timeout=1000):
                            rsa_retries += 1
                            if rsa_retries >= 3:
                                raise RuntimeError("RSA/MFA failed 3 times.")
                            await retry_btn.click()
                            await page.wait_for_timeout(3000)
                            continue
                except RuntimeError:
                    raise
                except Exception:
                    pass

                # Handle KMSI prompt
                if await MicrosoftSSOHandler.handle_stay_signed_in(page):
                    continue

                for sel in home_sels:
                    try:
                        if await page.locator(sel).first.is_visible(timeout=2000):
                            print(f"[{MasterFolderUtils.ts()}] Content Server home page ready")
                            await page.wait_for_timeout(2000)
                            return
                    except Exception:
                        pass
                await page.wait_for_timeout(2000)

            await MasterFolderUtils.save_debug_screenshot(page, self.config.error_screenshot_dir, "home_load_timeout")
            raise RuntimeError(f"Home page did not load within {timeout_s}s. URL: {page.url}")
        except Exception as e:
            if not isinstance(e, RuntimeError):
                print(f"[{MasterFolderUtils.ts()}] Error waiting for home page: {e}")
            raise

    async def navigate_home(self, page: Page):
        """Clicks the Home icon to return to the Content Server root page."""
        try:
            home_selectors = [
                "div.csui-home-item div",
                'a[title="Home page"]',
                "div.csui-home-item a",
                'xpath=/html/body/div[2]/nav/div/div[1]/div[2]/a/div',
            ]
            for sel in home_selectors:
                try:
                    loc = page.locator(sel).first
                    if await loc.is_visible(timeout=5000):
                        await loc.click()
                        await page.wait_for_timeout(3000)
                        return
                except Exception:
                    continue
        except Exception as e:
            print(f"[{MasterFolderUtils.ts()}] Navigate home warning: {e}")

    async def search_employee(self, page: Page, employee_number: str):
        """Clicks search icon -> types employee ID -> presses Enter -> waits for results container."""
        try:
            print(f"[{MasterFolderUtils.ts()}] Searching: {employee_number}")

            search_icon_selectors = [
                'a[title="Search in Content Server"]',
                "div.csui-search a",
                'xpath=/html/body/div[2]/nav/div/div[2]/div[2]/div/div[2]/a',
            ]
            for sel in search_icon_selectors:
                try:
                    loc = page.locator(sel).first
                    if await loc.is_visible(timeout=self.config.element_timeout):
                        await loc.click()
                        break
                except Exception:
                    continue
            await page.wait_for_timeout(1500)

            search_inputs = [
                'input[placeholder="Enter your search term"]',
                "body > div:nth-of-type(2) input",
                'xpath=/html/body/div[2]/nav/div/div[2]/div[2]/div/div[1]/div/div/input',
            ]
            inp = None
            for sel in search_inputs:
                try:
                    loc = page.locator(sel).first
                    if await loc.is_visible(timeout=5000):
                        inp = loc
                        break
                except Exception:
                    continue

            if not inp:
                raise RuntimeError("Search input not found")

            await inp.click()
            await page.wait_for_timeout(300)
            await inp.fill("")
            await page.wait_for_timeout(200)
            await inp.fill(employee_number)
            await page.wait_for_timeout(300)
            await inp.press("Enter")
            await page.wait_for_timeout(3500)

            # Dismiss error dialog if it appeared
            if await MasterFolderUtils.dismiss_error_dialog(page):
                print(f"[{MasterFolderUtils.ts()}]   CS dialog dismissed -- retrying search for {employee_number}")
                await self.navigate_home(page)
                await page.wait_for_timeout(2000)
                await self.search_employee(page, employee_number)
                return

            # Wait for search results container
            try:
                try:
                    await page.locator(
                        ".csui-loading-parent-wrapper, .binf-list-group-loading, .load-container, .cs-indicator"
                    ).first.wait_for(state="hidden", timeout=self.config.search_results_wait)
                except Exception:
                    pass
                await page.locator(
                    "#results, div.binf-list-group, .csui-search-results"
                ).first.wait_for(state="visible", timeout=self.config.search_results_wait)
            except PlaywrightTimeoutError:
                try:
                    if await page.locator("text=/About 0 hits/i").first.is_visible(timeout=2000):
                        raise RuntimeError(f"0 hits for {employee_number}")
                except RuntimeError:
                    raise
                except Exception:
                    pass

                if await MasterFolderUtils.dismiss_error_dialog(page):
                    raise RuntimeError(f"CS error during search for {employee_number}")
                await MasterFolderUtils.save_debug_screenshot(
                    page, self.config.error_screenshot_dir, "no_results", employee_number
                )
                raise RuntimeError(f"Search results timeout for {employee_number}")
        except Exception as e:
            if not isinstance(e, RuntimeError):
                print(f"[{MasterFolderUtils.ts()}] Search error for {employee_number}: {e}")
            raise

    async def folder_visible(self, page: Page, employee_number: str) -> bool:
        """
        Polls up to 8 times to check whether a folder named exactly <employee_number>
        is visible in search results.
        """
        try:
            for poll in range(8):
                # Role link — exact name match
                try:
                    if await page.get_by_role(
                        "link", name=employee_number, exact=True
                    ).first.is_visible(timeout=1000):
                        return True
                except Exception:
                    pass

                # Attribute checks (title, aria-label, data-binf-original-title)
                try:
                    loc = page.locator(
                        f'a[title="{employee_number}"], '
                        f'a[aria-label="{employee_number}"], '
                        f'a[data-binf-original-title="{employee_number}"]'
                    ).first
                    if await loc.is_visible(timeout=1000):
                        return True
                except Exception:
                    pass

                # After 3 polls, check for genuine 0-hits
                if poll >= 3:
                    try:
                        if await page.locator("text=/About 0 hits/i").first.is_visible(timeout=1000):
                            if await page.locator(f'a:has-text("{employee_number}")').count() == 0:
                                return False
                    except Exception:
                        pass

                await page.wait_for_timeout(2500)

            # Final pass: iterate through all result links for exact inner_text
            for sel in [".csui-search-item a", "div.binf-list-group a", "#results a"]:
                try:
                    links = page.locator(sel)
                    for i in range(await links.count()):
                        if (await links.nth(i).inner_text()).strip() == employee_number:
                            return True
                except Exception:
                    pass

            return False
        except Exception as e:
            print(f"[{MasterFolderUtils.ts()}] Folder visibility check error for {employee_number}: {e}")
            return False

    async def check_one_employee(self, page: Page, employee_number: str) -> str:
        """
        Full verification flow with built-in retries for a single employee.
        Returns STATUS_YES or STATUS_NO.
        """
        last_error = None
        for attempt in range(1, self.config.max_retries_per_employee + 1):
            try:
                if attempt > 1:
                    wait_s = attempt * 5
                    print(
                        f"[{MasterFolderUtils.ts()}]   Retry {attempt - 1}/{self.config.max_retries_per_employee - 1}"
                        f" for {employee_number} — waiting {wait_s}s before retrying..."
                    )
                    await page.wait_for_timeout(wait_s * 1000)
                    try:
                        await self.navigate_home(page)
                        await page.wait_for_timeout(2000)
                    except Exception:
                        pass

                await self.search_employee(page, employee_number)
                found = await self.folder_visible(page, employee_number)

                if found:
                    print(f"[{MasterFolderUtils.ts()}]   {employee_number} -> YES  (master folder found)")
                    return self.config.status_yes
                else:
                    print(f"[{MasterFolderUtils.ts()}]   {employee_number} -> NO   (master folder not found)")
                    return self.config.status_no

            except RuntimeError as e:
                # Genuine 0-hits is definitive — return No immediately without retry
                if "0 hits" in str(e) or "0 search results" in str(e):
                    print(f"[{MasterFolderUtils.ts()}]   {employee_number} -> NO   (0 search hits)")
                    return self.config.status_no
                last_error = str(e)
                print(
                    f"[{MasterFolderUtils.ts()}]   Attempt {attempt}/{self.config.max_retries_per_employee} "
                    f"failed for {employee_number}: {last_error}"
                )
                if attempt == self.config.max_retries_per_employee:
                    await MasterFolderUtils.save_debug_screenshot(
                        page, self.config.error_screenshot_dir, "all_retries_exhausted", employee_number
                    )

            except Exception as e:
                last_error = str(e)
                print(
                    f"[{MasterFolderUtils.ts()}]   Attempt {attempt}/{self.config.max_retries_per_employee} "
                    f"failed for {employee_number}: {last_error}"
                )
                if attempt == self.config.max_retries_per_employee:
                    await MasterFolderUtils.save_debug_screenshot(
                        page, self.config.error_screenshot_dir, "all_retries_exhausted", employee_number
                    )

        print(
            f"[{MasterFolderUtils.ts()}]   {employee_number} -> NO   "
            f"(could not confirm after {self.config.max_retries_per_employee} attempts — last error: {last_error})"
        )
        return self.config.status_no

    def save_excel(self, results: Dict[str, str]):
        """Writes current results to Excel report with professional column widths."""
        try:
            self.config.output_dir.mkdir(parents=True, exist_ok=True)
            df = pd.DataFrame(
                [{"Employee ID": eid, "Master Folder Exists": st} for eid, st in results.items()],
                columns=["Employee ID", "Master Folder Exists"],
            )
            with pd.ExcelWriter(str(self.config.output_excel_path), engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Master_Folder_Check")
                ws = writer.sheets["Master_Folder_Check"]
                for ci, col in enumerate(df.columns, 1):
                    max_w = max(
                        len(str(col)),
                        df[col].astype(str).str.len().max() if not df.empty else 0,
                    )
                    ws.column_dimensions[ws.cell(1, ci).column_letter].width = max(max_w + 5, 22)
            print(f"[{MasterFolderUtils.ts()}]   Excel updated: {self.config.output_excel_path.name}  ({len(df)} row(s))")
        except Exception as e:
            print(f"[{MasterFolderUtils.ts()}] [ERROR] Excel write failed: {e}")

    def print_summary(self, results: Dict[str, str]):
        """Prints formatted completion summary to console and log file."""
        try:
            yes_list = [e for e, s in results.items() if s == self.config.status_yes]
            no_list = [e for e, s in results.items() if s == self.config.status_no]

            print()
            print(f"[{MasterFolderUtils.ts()}] ═══════════════════════════════════════════════════════")
            print(f"[{MasterFolderUtils.ts()}] MASTER FOLDER EXISTENCE CHECK — SUMMARY")
            print(f"[{MasterFolderUtils.ts()}] ═══════════════════════════════════════════════════════")
            if yes_list:
                print(f"[{MasterFolderUtils.ts()}]   [YES] Folder EXISTS  ({len(yes_list)}): {', '.join(yes_list)}")
            if no_list:
                print(f"[{MasterFolderUtils.ts()}]   [NO]  Folder MISSING ({len(no_list)}): {', '.join(no_list)}")
            print(f"[{MasterFolderUtils.ts()}] ───────────────────────────────────────────────────────")
            print(f"[{MasterFolderUtils.ts()}]   Total: {len(results)}  |  Yes: {len(yes_list)}  |  No: {len(no_list)}")
            print(f"[{MasterFolderUtils.ts()}]   Report saved to: {self.config.output_excel_path}")
            print(f"[{MasterFolderUtils.ts()}] ═══════════════════════════════════════════════════════")
        except Exception as e:
            print(f"[{MasterFolderUtils.ts()}] Summary print error: {e}")

    async def run(self, employee_numbers: List[str]):
        """Runs parallel employee master folder verification across browser tabs."""
        try:
            self.config.output_dir.mkdir(parents=True, exist_ok=True)

            if not employee_numbers:
                print(f"[{MasterFolderUtils.ts()}] No employee numbers provided. Exiting.")
                return

            if not self.config.content_server_url:
                print(f"[{MasterFolderUtils.ts()}] [FATAL] CONTENT_SERVER_URL is not set in .env!")
                return

            concurrency = min(self.config.max_concurrent_tabs, len(employee_numbers))
            print(
                f"[{MasterFolderUtils.ts()}] Master Folder Checker  |  "
                f"{'headless' if self.config.headless else 'visible'}  |  "
                f"{concurrency} tab(s)  |  {len(employee_numbers)} employee(s)  |  "
                f"max {self.config.max_retries_per_employee} retries/employee"
            )
            MasterFolderUtils.cleanup_profile_locks(self.config.automation_profile_dir)

            all_results: Dict[str, str] = {}
            excel_lock = asyncio.Lock()

            async def run_batch(ctx, emp_list, num_tabs, nav_lock, total):
                first = ctx.pages[0] if ctx.pages else await ctx.new_page()
                workers = [first]
                for i in range(1, num_tabs):
                    try:
                        tab = await ctx.new_page()
                        workers.append(tab)
                        print(f"[{MasterFolderUtils.ts()}]   Opened tab {i + 1}/{num_tabs}")
                        await asyncio.sleep(0.5)
                    except Exception as e:
                        print(f"[{MasterFolderUtils.ts()}]   Could not open tab {i + 1}: {e}")
                        break

                print(f"[{MasterFolderUtils.ts()}] {len(workers)} tab(s) ready.\n")

                async def worker(pg, emp, idx):
                    print(f"[{MasterFolderUtils.ts()}] ── [{idx}/{total}] {emp} ──")
                    try:
                        for nav_attempt in range(2):
                            try:
                                async with nav_lock:
                                    await pg.goto(
                                        self.config.content_server_url,
                                        wait_until="domcontentloaded",
                                        timeout=self.config.nav_timeout,
                                    )
                                    await pg.wait_for_timeout(2000)
                                    await self.wait_for_home(pg, timeout_s=30)
                                break
                            except Exception as e:
                                if nav_attempt == 1:
                                    raise RuntimeError(f"Navigation failed for {emp}: {e}")
                                await pg.wait_for_timeout(3000)

                        status = await self.check_one_employee(pg, emp)
                        all_results[emp] = status
                        async with excel_lock:
                            self.save_excel(all_results)

                    except Exception as e:
                        print(f"[{MasterFolderUtils.ts()}]   Navigation completely failed for {emp}: {e}")
                        await MasterFolderUtils.save_debug_screenshot(
                            pg, self.config.error_screenshot_dir, "nav_fail", emp
                        )
                        all_results[emp] = self.config.status_no
                        async with excel_lock:
                            self.save_excel(all_results)

                for i in range(0, len(emp_list), len(workers)):
                    batch = emp_list[i : i + len(workers)]
                    await asyncio.gather(
                        *[worker(workers[j], emp, i + j + 1) for j, emp in enumerate(batch)]
                    )
                    await asyncio.sleep(1.5)

                for wp in workers[1:]:
                    try:
                        await wp.close()
                    except Exception:
                        pass

            ctx = None
            try:
                async with async_playwright() as p:
                    print(f"[{MasterFolderUtils.ts()}] Launching browser...")
                    chrome_args = [
                        "--no-first-run",
                        "--no-default-browser-check",
                        "--disable-blink-features=AutomationControlled",
                        "--disable-session-crashed-bubble",
                        "--hide-crash-restore-bubble",
                    ]
                    if self.config.window_position:
                        chrome_args.append(f"--window-position={self.config.window_position}")

                    if self.config.headless:
                        chrome_args.append("--window-size=1920,1080")
                        launch_kwargs = {
                            "headless": True,
                            "viewport": {"width": 1920, "height": 1080},
                        }
                    else:
                        if self.config.start_maximized:
                            chrome_args.append("--start-maximized")
                            launch_kwargs = {
                                "headless": False,
                                "no_viewport": True,
                            }
                        else:
                            chrome_args.append("--window-size=1517,900")
                            launch_kwargs = {
                                "headless": False,
                                "viewport": {"width": 1517, "height": 900},
                            }

                    ctx = await p.chromium.launch_persistent_context(
                        user_data_dir=str(self.config.automation_profile_dir),
                        channel=self.config.browser_channel,
                        accept_downloads=False,
                        args=chrome_args,
                        **launch_kwargs,
                    )

                    first = ctx.pages[0] if ctx.pages else await ctx.new_page()
                    if not self.config.headless:
                        try:
                            await first.bring_to_front()
                        except Exception:
                            pass
                    print(f"[{MasterFolderUtils.ts()}] Navigating to Content Server...")
                    await first.goto(
                        self.config.content_server_url,
                        wait_until="domcontentloaded",
                        timeout=self.config.nav_timeout,
                    )
                    await first.wait_for_timeout(3000)
                    await self.wait_for_home(first, timeout_s=120)
                    await first.wait_for_timeout(2000)
                    print(f"[{MasterFolderUtils.ts()}] Login confirmed.\n")

                    nav_lock = asyncio.Lock()
                    await run_batch(
                        ctx, employee_numbers, concurrency, nav_lock, len(employee_numbers)
                    )

            except Exception as e:
                print(f"[{MasterFolderUtils.ts()}] FATAL: {e}\n{traceback.format_exc()}")
            finally:
                if ctx:
                    for pg in ctx.pages:
                        try:
                            await pg.close()
                        except Exception:
                            pass
                    try:
                        await ctx.close()
                        print(f"[{MasterFolderUtils.ts()}] Browser context closed cleanly.")
                    except Exception as e:
                        print(f"[{MasterFolderUtils.ts()}] Warning closing browser context: {e}")

            self.print_summary(all_results)
            self.save_excel(all_results)
        except Exception as e:
            print(f"[{MasterFolderUtils.ts()}] Execution failed: {e}\n{traceback.format_exc()}")


# Input Prompt Handler
class InputPromptHandler:
    """Handles terminal user interactions and employee list loading from Excel/CSV files."""

    @staticmethod
    def load_from_file(file_path: Path) -> List[str]:
        """Reads employee IDs from an Excel (.xlsx, .xls, .xlsm) or CSV file."""
        try:
            ext = file_path.suffix.lower()
            if ext in [".xlsx", ".xlsm"]:
                df = pd.read_excel(file_path, dtype=str, engine="openpyxl")
            elif ext == ".xls":
                df = pd.read_excel(file_path, dtype=str)
            elif ext == ".csv":
                df = pd.read_csv(file_path, dtype=str)
            else:
                raise ValueError(f"Unsupported format: {ext}")

            if df.empty:
                return []

            col = None
            priority_patterns = [
                r"employee[_\s\-]*code",
                r"emp[_\s\-]*code",
                r"employee[_\s\-]*id",
                r"emp[_\s\-]*id",
                r"employee[_\s\-]*no",
                r"employee[_\s\-]*number",
                r"emp[_\s\-]*no",
                r"user[_\s\-]*id",
                r"perner",
                r"id",
            ]
            for pat in priority_patterns:
                for c in df.columns:
                    if re.search(pat, str(c).strip(), re.IGNORECASE):
                        col = c
                        break
                if col:
                    break

            if not col:
                col = df.columns[0]
                print(f"[{MasterFolderUtils.ts()}] No employee column found; using first column: '{col}'")
            else:
                print(f"[{MasterFolderUtils.ts()}] Employee column: '{col}'")

            ids, seen = [], set()
            for val in df[col].dropna():
                v = str(val).strip()
                if v.endswith(".0"):
                    v = v[:-2]
                v = v.strip()
                if v and v.lower() != "nan" and v not in seen:
                    seen.add(v)
                    ids.append(v)
            return ids
        except Exception as e:
            print(f"[{MasterFolderUtils.ts()}] File load error: {e}")
            return []

    @staticmethod
    def prompt_input(max_retries: int = 3) -> List[str]:
        """Prompts the user interactively in the terminal for file paths or employee IDs."""
        try:
            t = sys.__stdout__
            t.write("\n")
            t.write("  ╔══════════════════════════════════════════════════════════════════╗\n")
            t.write("  ║        Employee Master Folder Existence Checker                 ║\n")
            t.write("  ╚══════════════════════════════════════════════════════════════════╝\n")
            t.write("\n")
            t.write("  Checks whether each employee master folder exists on Content Server.\n")
            t.write("  Output -> uploads/Master_Folder_Check/Master_Folder_Check_Report.xlsx\n")
            t.write("  Columns: Employee ID | Master Folder Exists  (Yes / No only)\n")
            t.write(f"  Auto-retries transient errors up to {max_retries}x per employee.\n")
            t.write("\n")
            t.write("  Enter Excel/CSV file path  OR  Employee ID(s) (comma/space separated):\n")
            t.write("  > ")
            t.flush()

            try:
                raw = sys.__stdin__.readline().strip().strip("'\"")
            except (EOFError, KeyboardInterrupt):
                return []

            if not raw:
                t.write("  No input entered.\n\n")
                t.flush()
                return []

            p = Path(raw)
            if not p.exists():
                alt = Path(__file__).parent.parent / raw
                if alt.exists():
                    p = alt

            if p.exists() and p.suffix.lower() in [".xlsx", ".xls", ".xlsm", ".csv"]:
                try:
                    t.write(f"  -> Reading from file: {p.name}...\n")
                    t.flush()
                    ids = InputPromptHandler.load_from_file(p)
                    sample = ", ".join(ids[:6]) + ("..." if len(ids) > 6 else "")
                    t.write(f"  -> Loaded {len(ids):,} employee ID(s). Sample: {sample}\n\n")
                    t.flush()
                    return ids
                except Exception as e:
                    t.write(f"  [ERROR] {e}\n\n")
                    t.flush()
                    return []

            seen, ids = set(), []
            for tok in raw.replace(",", " ").split():
                tok = tok.strip().strip("'\"")
                if tok and tok not in seen:
                    seen.add(tok)
                    ids.append(tok)

            if ids:
                sample = ", ".join(ids[:10]) + ("..." if len(ids) > 10 else "")
                t.write(f"  -> {len(ids)} employee(s): {sample}\n\n")
                t.flush()
            else:
                t.write("  No employee IDs found.\n\n")
                t.flush()
            return ids
        except Exception as e:
            print(f"[{MasterFolderUtils.ts()}] Terminal prompt error: {e}")
            return []


# Main execution entrypoint
async def check_master_folders(employee_numbers: List[str]):
    """Main execution entrypoint for checking employee master folders."""
    config = MasterFolderConfig()
    checker = MasterFolderChecker(config)
    await checker.run(employee_numbers)


if __name__ == "__main__":
    ids = InputPromptHandler.prompt_input()
    if ids:
        asyncio.run(check_master_folders(ids))
    else:
        print("No employee numbers entered. Exiting.")

