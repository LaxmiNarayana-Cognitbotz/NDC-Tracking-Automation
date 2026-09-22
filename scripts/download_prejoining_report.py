# OpenText Content Server - Pre-Joining & Talent Acquisition Document Downloader
# Downloads onboarding documents: Appointment Letter, Offer Letter, Background Verification (BGV),
# Pre Talent Acquisition, Offer of Appointment, LOI, etc.

import asyncio
import glob
import os
import re
import subprocess
import sys
import traceback
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

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
LOG_FILE = LOG_DIR / "download_prejoining_report.log"

sys.stdout = LoggerWriter(LOG_FILE)
sys.stderr = LoggerWriter(LOG_FILE)

load_dotenv(verbose=True)


# Application Configuration
class PreJoiningConfig:
    """Manages credentials, file paths, timeouts, and execution parameters for Pre-Joining downloader."""

    def __init__(self):
        try:
            # SSO credentials & Target URL
            self.oracle_email: str = (os.getenv("ORACLE_EMAIL") or "").strip()
            self.oracle_password: str = (os.getenv("ORACLE_PASSWORD") or "").strip()
            self.content_server_url: str = (os.getenv("CONTENT_SERVER_URL") or "").strip()

            # Output paths
            self.download_dir: Path = Path(__file__).parent.parent / "uploads" / "Pre_Joining_Reports"
            self.master_excel_path: Path = self.download_dir / "Master_Pre_Joining_Summary.xlsx"
            self.automation_profile_dir: Path = Path(__file__).parent.parent / "chrome_automation_profile"
            self.error_screenshot_dir: Path = Path(__file__).parent.parent / "errors" / "prejoining_error"

            # Browser execution settings
            self.headless: bool = self._parse_bool(os.getenv("HEADLESS"), default=True)
            self.start_maximized: bool = self._parse_bool(os.getenv("START_MAXIMIZED"), default=True)
            self.window_position: str = (os.getenv("WINDOW_POSITION") or "").strip()
            self.browser_channel: Optional[str] = (os.getenv("BROWSER_CHANNEL") or "").strip() or None

            # Timeouts & Limits
            self.element_timeout: int = 15_000        # ms
            self.nav_timeout: int = 90_000            # ms
            self.search_results_wait: int = 45_000    # ms
            self.max_concurrent_tabs: int = 5
            self.max_retry_rounds: int = 5

            # Result Status Constants
            self.status_downloaded: str = "DOWNLOADED"
            self.status_no_docs: str = "NO_PREJOINING_DOCS"
            self.status_not_found: str = "NOT_FOUND_0_HITS"
            self.status_error: str = "ERROR"

            # Ensure directories exist
            self.download_dir.mkdir(parents=True, exist_ok=True)
            self.error_screenshot_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] PreJoiningConfig error: {e}")

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
class PreJoiningUtils:
    """Utility helper functions for timing, screenshots, lock cleanup, regex matching, and dialogs."""

    # ── Document Matching Regex Patterns ─────────────────────────────────────
    PREJOINING_DOCUMENT_PATTERN = re.compile(
        r"(?i)("
        r"appointment[_\s\-]*letter"
        r"|letter[_\s\-]*of[_\s\-]*appointment"
        r"|offer[_\s\-]*of[_\s\-]*appointment"
        r"|offer[_\s\-]*&[_\s\-]*appointment"
        r"|offer[_\s\-]*letter(?:[_\s\-]*or[_\s\-]*letter[_\s\-]*(?:of|on)[_\s\-]*intent)?"
        r"|letter[_\s\-]*(?:of|on)[_\s\-]*intent"
        r"|(?<![a-zA-Z])loi(?![a-zA-Z])"
        r"|background[_\s\-]*verification(?:[_\s\-]*report)?"
        r"|(?<![a-zA-Z])bgv(?:[_\s\-]*report)?(?![a-zA-Z])"
        r"|background[_\s\-]*check"
        r"|pre[_\s\-]*talent[_\s\-]*acquisition"
        r"|talent[_\s\-]*acquisition"
        r"|pre[_\s\-]*joining"
        r")"
    )

    OFFER_LETTER_PATTERN = re.compile(
        r"(?i)("
        r"offer[_\s\-]*letter(?:[_\s\-]*or[_\s\-]*letter[_\s\-]*(?:of|on)[_\s\-]*intent)?"
        r"|letter[_\s\-]*(?:of|on)[_\s\-]*intent"
        r"|(?<![a-zA-Z])loi(?![a-zA-Z])"
        r"|offer[_\s\-]*of[_\s\-]*appointment"
        r"|offer[_\s\-]*&[_\s\-]*appointment"
        r")"
    )

    BGV_PATTERN = re.compile(
        r"(?i)("
        r"background[_\s\-]*verification(?:[_\s\-]*report)?"
        r"|(?<![a-zA-Z])bgv(?:[_\s\-]*report)?(?![a-zA-Z])"
        r"|background[_\s\-]*check"
        r")"
    )

    APPOINTMENT_LETTER_PATTERN = re.compile(
        r"(?i)("
        r"appointment[_\s\-]*letter"
        r"|letter[_\s\-]*of[_\s\-]*appointment"
        r")"
    )

    PREVIOUS_DOC_PATTERN = re.compile(
        r"(?i)(?:previous|prior[_\s\-]*employ|past[_\s\-]*employ)",
    )

    @staticmethod
    def ts() -> str:
        """Returns current timestamp string."""
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    @staticmethod
    def cleanup_profile_locks(profile_dir: Path):
        """Cleans up Chromium profile lock files to prevent startup collisions."""
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
            print(f"[{PreJoiningUtils.ts()}] [CLEANUP] Profile lock cleanup warning: {e}")

    @staticmethod
    def cleanup_error_screenshots(employee_number: str, error_dir: Path):
        """Deletes debug error screenshots for the given employee upon successful completion."""
        try:
            if not error_dir.exists():
                return
            clean_id = re.sub(r"\D", "", str(employee_number))
            patterns = [
                str(error_dir / f"debug_*_{clean_id}.png"),
                str(error_dir / f"debug_*_{clean_id}_*.png"),
            ]
            for pattern in patterns:
                for path in glob.glob(pattern):
                    try:
                        os.remove(path)
                        print(f"[{PreJoiningUtils.ts()}]   Cleaned up error screenshot: {Path(path).name}")
                    except OSError:
                        pass
        except Exception as e:
            print(f"[{PreJoiningUtils.ts()}] Cleanup error screenshot warning: {e}")

    @staticmethod
    def extract_employee_id_from_debug_screenshot(path: Path) -> Optional[str]:
        """Extracts employee ID from a debug screenshot filename (debug_{type}_{id}.png)."""
        try:
            stem = path.stem
            match = re.match(r"^debug_[a-z_]+_(\d+)$", stem, flags=re.IGNORECASE)
            if match:
                return match.group(1)

            match2 = re.search(r"(?<!\d)(\d{6,10})(?!\d)", stem)
            if match2:
                return match2.group(1)
        except Exception:
            pass
        return None

    @staticmethod
    def debug_screenshot_path(label: str, error_dir: Path, employee_number: Optional[str] = None) -> Path:
        """Constructs canonical debug screenshot path."""
        safe_label = re.sub(r"[^A-Za-z_]+", "_", label).strip("_").lower()
        if employee_number:
            clean_id = re.sub(r"\D", "", str(employee_number))
            return error_dir / f"debug_{safe_label}_{clean_id}.png"
        return error_dir / f"debug_{safe_label}.png"

    @staticmethod
    def get_failed_ids_from_error_folder(error_dir: Path) -> List[str]:
        """Scans the error directory for debug screenshots and extracts unique employee IDs."""
        try:
            if not error_dir.exists():
                return []
            seen = set()
            ids = []
            for path in sorted(error_dir.glob("debug_*.png")):
                emp_id = PreJoiningUtils.extract_employee_id_from_debug_screenshot(path)
                if emp_id and emp_id not in seen:
                    seen.add(emp_id)
                    ids.append(emp_id)
            return ids
        except Exception as e:
            print(f"[{PreJoiningUtils.ts()}] Failed IDs extraction warning: {e}")
            return []

    @staticmethod
    async def save_debug_screenshot(
        page: Page, error_dir: Path, label: str, employee_number: Optional[str] = None
    ) -> str:
        """Captures a full-page diagnostic screenshot."""
        try:
            error_dir.mkdir(parents=True, exist_ok=True)
            path = PreJoiningUtils.debug_screenshot_path(label, error_dir, employee_number)
            await page.screenshot(path=str(path), full_page=True)
            print(f"[{PreJoiningUtils.ts()}]   Screenshot: {path}")
            return str(path)
        except Exception as e:
            print(f"[{PreJoiningUtils.ts()}]   Could not save screenshot: {e}")
            return ""

    @staticmethod
    async def check_and_dismiss_error_dialog(page: Page) -> bool:
        """Detects and dismisses Content Server error modal dialogs."""
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
                        print(f"[{PreJoiningUtils.ts()}] [WARN] Content Server error dialog detected: {error_text}")

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
                                    print(f"[{PreJoiningUtils.ts()}]   Dismissed error dialog")
                                    await page.wait_for_timeout(1500)
                                    return True
                            except Exception:
                                continue

                        await page.keyboard.press("Escape")
                        print(f"[{PreJoiningUtils.ts()}]   Dismissed error dialog via Escape key")
                        await page.wait_for_timeout(1500)
                        return True
                except Exception:
                    continue
        except Exception:
            pass
        return False

    @staticmethod
    async def click_with_fallback(
        page_or_frame, selectors: List[str], description: str, timeout_each: int = 10_000
    ):
        """Attempts to click a list of fallback selectors sequentially until one succeeds."""
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
    def classify_document(filename: str) -> List[str]:
        """
        Classifies a downloaded file into one or more categories:
        'Offer Letter', 'Background Verification', 'Appointment Letter'.
        Excludes previous employment records.
        """
        try:
            if PreJoiningUtils.PREVIOUS_DOC_PATTERN.search(filename):
                return []

            categories = []
            if PreJoiningUtils.OFFER_LETTER_PATTERN.search(filename):
                categories.append("Offer Letter")
            if PreJoiningUtils.BGV_PATTERN.search(filename):
                categories.append("Background Verification")
            if PreJoiningUtils.APPOINTMENT_LETTER_PATTERN.search(filename):
                categories.append("Appointment Letter")
            return categories
        except Exception:
            return []

    @staticmethod
    def is_prejoining_document(name: str) -> bool:
        """Returns True if document matches onboarding regex and is NOT from past employment."""
        try:
            if PreJoiningUtils.PREVIOUS_DOC_PATTERN.search(name):
                return False
            return bool(PreJoiningUtils.PREJOINING_DOCUMENT_PATTERN.search(name))
        except Exception:
            return False


# Microsoft SSO Handler
class MicrosoftSSOHandler:
    """Manages Microsoft Entra / SSO authentication state machine, KMSI, and MFA approvals."""

    @staticmethod
    async def check_for_sso_errors(page: Page):
        """Inspects the page for SSO validation errors and raises a RuntimeError if found."""
        for error_id in ["usernameError", "passwordError", "error"]:
            try:
                loc = page.locator(f"#{error_id}").first
                if await loc.is_visible(timeout=1000):
                    err_text = await loc.inner_text()
                    err_text = err_text.strip().replace("\n", " ")
                    if err_text:
                        raise RuntimeError(f"Microsoft SSO Error: {err_text}")
            except RuntimeError:
                raise
            except Exception:
                pass

        try:
            alert_loc = page.locator('[role="alert"]').first
            if await alert_loc.is_visible(timeout=1000):
                text = (await alert_loc.inner_text()).strip().replace("\n", " ")
                if text:
                    raise RuntimeError(f"Microsoft SSO Alert: {text}")
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
                        print(f"[{PreJoiningUtils.ts()}]   Checked 'Don't show this again' (KMSI)")
                        await page.wait_for_timeout(300)
                except Exception:
                    pass

            stay_signed_btn = page.locator(
                '#idSIButton9, input[value="Yes"], input[type="submit"][value="Yes"], button:has-text("Yes")'
            ).first
            if await stay_signed_btn.is_visible(timeout=2000):
                await stay_signed_btn.click(force=True, timeout=5000)
                print(f"[{PreJoiningUtils.ts()}]   Clicked 'Stay signed in? -> Yes'")
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
                print(f"[{PreJoiningUtils.ts()}]   'Verify your identity' screen detected.")
                await rsa_btn.click(force=True, timeout=5000)
                print(f"[{PreJoiningUtils.ts()}]   Clicked 'Approve with RSAEntraMFA'")
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
          3. Select RSAEntraMFA if 'Verify your identity' prompt appears
          4. Handle 'Stay signed in?' (KMSI) prompt
          5. Wait for RSA/MFA push notification approval
        """
        if not email or not password:
            print(f"[{PreJoiningUtils.ts()}] [WARN] No credentials in .env -- waiting for manual login...")
            print(f"[{PreJoiningUtils.ts()}]   Create a .env file with ORACLE_EMAIL and ORACLE_PASSWORD")
            return False

        print(f"[{PreJoiningUtils.ts()}] SSO login detected -- filling credentials automatically...")
        await page.wait_for_timeout(2000)
        await MicrosoftSSOHandler.check_for_sso_errors(page)

        email_filled = False
        password_filled = False
        rsa_clicked = False

        for iteration in range(25):
            current_url = page.url.lower()
            if not any(x in current_url for x in ["login.microsoftonline", "login.microsoft", "adfs", "saml"]):
                print(f"[{PreJoiningUtils.ts()}]   SSO login completed -- redirected away from login screen.")
                return True

            # Check for Stay signed in (KMSI) prompt
            if await MicrosoftSSOHandler.handle_stay_signed_in(page):
                continue

            # Check for RSAEntraMFA selection button
            if not rsa_clicked and await MicrosoftSSOHandler.handle_rsa_entra_selection(page):
                rsa_clicked = True
                break

            # 1. "Pick an account" screen
            if not email_filled and not password_filled:
                try:
                    pick_header = page.locator(
                        'div#loginHeader:has-text("Pick an account"), '
                        'div[role="heading"]:has-text("Pick an account")'
                    ).first
                    if await pick_header.is_visible(timeout=1000):
                        print(f"[{PreJoiningUtils.ts()}]   'Pick an account' screen detected.")
                        target_email = email.strip().lower()
                        account_found = False

                        account_selectors = [
                            'div.table[role="button"]',
                            'div[data-test-id]',
                            'div.row.tile',
                            'li.table',
                        ]
                        for sel in account_selectors:
                            try:
                                accounts = page.locator(sel)
                                count = await accounts.count()
                                for i in range(count):
                                    tile = accounts.nth(i)
                                    tile_text = (await tile.inner_text()).strip().lower()
                                    if target_email in tile_text:
                                        print(f"[{PreJoiningUtils.ts()}]   Found matching account tile: {email}")
                                        await tile.click(force=True, timeout=5000)
                                        account_found = True
                                        break
                                if account_found:
                                    break
                            except Exception:
                                continue

                        if not account_found:
                            try:
                                el = page.locator(f'small:has-text("{email}"), div:has-text("{email}")').first
                                if await el.is_visible(timeout=2000):
                                    parent = el.locator(
                                        'xpath=ancestor::div[@role="button" or contains(@class,"table") or contains(@class,"tile")]'
                                    ).first
                                    if await parent.count() > 0:
                                        await parent.click(force=True, timeout=5000)
                                    else:
                                        await el.click(force=True, timeout=5000)
                                    account_found = True
                                    print(f"[{PreJoiningUtils.ts()}]   Clicked account tile via email match: {email}")
                            except Exception:
                                pass

                        if not account_found:
                            try:
                                use_another = page.locator('div#otherTile, div:has-text("Use another account")').last
                                if await use_another.is_visible(timeout=3000):
                                    await use_another.click(force=True, timeout=5000)
                                    print(f"[{PreJoiningUtils.ts()}]   Account not in list -- clicked 'Use another account'")
                            except Exception:
                                print(f"[{PreJoiningUtils.ts()}]   Could not find account or 'Use another account' button")

                        await page.wait_for_timeout(3000)
                        continue
                except Exception:
                    pass

            # 2. Password input screen
            if not password_filled:
                try:
                    pw_input = page.locator(
                        'input[name="passwd"], input[id="i0118"], input[type="password"]'
                    ).first
                    pw_indicator = page.locator(
                        'input[type="submit"][value="Sign in"], #idSIButton9[value="Sign in"], '
                        '#idSIButton9:has-text("Sign in"), div:has-text("Enter password"), '
                        'div:has-text("enter password"), #identityBanner, #idA_PWD_SwitchToCredPicker'
                    ).first
                    if await pw_input.is_visible(timeout=1000) and (
                        email_filled or await pw_indicator.is_visible(timeout=1000)
                    ):
                        print(f"[{PreJoiningUtils.ts()}]   Password screen detected.")
                        for attempt in range(1, 4):
                            await pw_input.click(force=True, timeout=3000)
                            await pw_input.fill("", force=True, timeout=3000)
                            await page.wait_for_timeout(200)
                            await pw_input.type(password.strip(), delay=50)
                            print(f"[{PreJoiningUtils.ts()}]   Entered password (Attempt {attempt}/3)")
                            await page.wait_for_timeout(500)

                            signin_btn = page.locator('input[type="submit"], #idSIButton9').first
                            await signin_btn.click(force=True, timeout=5000)
                            print(f"[{PreJoiningUtils.ts()}]   Clicked Sign In")
                            await page.wait_for_timeout(3000)

                            try:
                                await MicrosoftSSOHandler.check_for_sso_errors(page)
                                password_filled = True
                                email_filled = True
                                break
                            except RuntimeError as e:
                                if attempt == 3:
                                    raise
                                print(f"[{PreJoiningUtils.ts()}]   {e} - Retrying...")
                        continue
                except RuntimeError:
                    raise
                except Exception:
                    pass

            # 3. Email input screen
            if not email_filled and not password_filled:
                try:
                    email_input = page.locator(
                        'input[name="loginfmt"], input[type="email"], input[id="i0116"]'
                    ).first
                    if await email_input.is_visible(timeout=1000):
                        print(f"[{PreJoiningUtils.ts()}]   Email screen detected.")
                        for attempt in range(1, 4):
                            await email_input.click(force=True, timeout=3000)
                            await email_input.fill("", force=True, timeout=3000)
                            await page.wait_for_timeout(200)
                            await email_input.type(email.strip(), delay=50)
                            print(f"[{PreJoiningUtils.ts()}]   Entered email (Attempt {attempt}/3)")
                            await page.wait_for_timeout(500)

                            next_btn = page.locator('input[type="submit"], #idSIButton9').first
                            await next_btn.click(force=True, timeout=5000)
                            print(f"[{PreJoiningUtils.ts()}]   Clicked Next")
                            await page.wait_for_timeout(3000)

                            try:
                                await MicrosoftSSOHandler.check_for_sso_errors(page)
                                email_filled = True
                                break
                            except RuntimeError as e:
                                if attempt == 3:
                                    raise
                                print(f"[{PreJoiningUtils.ts()}]   {e} - Retrying...")
                        continue
                except RuntimeError:
                    raise
                except Exception:
                    pass

            # If password was filled, check subsequent screens
            if password_filled:
                if not rsa_clicked and await MicrosoftSSOHandler.handle_rsa_entra_selection(page):
                    rsa_clicked = True
                    break
                if await MicrosoftSSOHandler.handle_stay_signed_in(page):
                    continue
                if any(x in current_url for x in ["securid.com", "adani.auth"]) or rsa_clicked:
                    break

            await page.wait_for_timeout(1000)

        print(f"[{PreJoiningUtils.ts()}] Waiting for RSA/MFA approval...")
        print()
        print("  +------------------------------------------------------------+")
        print("  |  RSA push notification sent to the user's phone.           |")
        print("  |  Please ask them to APPROVE it.                            |")
        print("  |  Waiting up to 5 minutes...                                |")
        print("  +------------------------------------------------------------+")
        print()

        return True


# Pre-Joining Report Downloader Core Engine
class PreJoiningReportDownloader:
    """Orchestrates OpenText Content Server navigation, searches, Pre-Joining folder traversal, and downloads."""

    def __init__(self, config: PreJoiningConfig):
        self.config = config

    async def wait_for_content_server_home(self, page: Page, timeout_s: int = 300):
        """
        Waits for Content Server home page. Handles Microsoft SSO login if redirected,
        RSA/MFA retry (up to 3 attempts), and 'Stay signed in?' prompt.
        """
        try:
            print(f"[{PreJoiningUtils.ts()}] Waiting for Content Server home page (up to {timeout_s}s)...")
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
                if any(x in current_url for x in ["login.microsoftonline", "login.microsoft", "adfs", "saml"]):
                    if await MicrosoftSSOHandler.handle_stay_signed_in(page):
                        continue
                    if await MicrosoftSSOHandler.handle_rsa_entra_selection(page):
                        continue
                    if not login_handled:
                        login_handled = await MicrosoftSSOHandler.handle_microsoft_sso(
                            page, self.config.oracle_email, self.config.oracle_password
                        )

                try:
                    if "securid.com" in current_url or "adani.auth" in current_url:
                        retry_btn = page.locator('button:has-text("Retry")').first
                        if await retry_btn.is_visible(timeout=1000):
                            rsa_retry_count += 1
                            print(
                                f"[{PreJoiningUtils.ts()}] [WARN] RSA/MFA Authentication failed or rejected "
                                f"(Attempt {rsa_retry_count}/3)."
                            )
                            if rsa_retry_count >= 3:
                                await PreJoiningUtils.save_debug_screenshot(
                                    page, self.config.error_screenshot_dir, "rsa_failed"
                                )
                                raise RuntimeError("RSA/MFA Authentication failed 3 times.")

                            print(f"[{PreJoiningUtils.ts()}]   Clicking 'Retry' for RSA...")
                            await retry_btn.click()
                            await page.wait_for_timeout(3000)
                            print(f"[{PreJoiningUtils.ts()}] Waiting for RSA/MFA approval...")
                            continue
                except RuntimeError:
                    raise
                except Exception:
                    pass

                # Handle KMSI prompt
                if await MicrosoftSSOHandler.handle_stay_signed_in(page):
                    continue

                for sel in home_selectors:
                    try:
                        if await page.locator(sel).first.is_visible(timeout=2000):
                            print(f"[{PreJoiningUtils.ts()}] Content Server home page ready")
                            await page.wait_for_timeout(2000)
                            return
                    except Exception:
                        pass
                await page.wait_for_timeout(2000)

            await PreJoiningUtils.save_debug_screenshot(
                page, self.config.error_screenshot_dir, "home_load_timeout"
            )
            raise RuntimeError(
                f"Content Server home page did not load within {timeout_s}s.\nCurrent URL: {page.url}"
            )
        except Exception as e:
            if not isinstance(e, RuntimeError):
                print(f"[{PreJoiningUtils.ts()}] Error waiting for Content Server home: {e}")
            raise

    async def navigate_home(self, page: Page):
        """Clicks the Home icon to return to Content Server home page."""
        try:
            await PreJoiningUtils.click_with_fallback(
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
        except Exception as e:
            print(f"[{PreJoiningUtils.ts()}] Navigate home warning: {e}")

    async def search_employee(self, page: Page, employee_number: str):
        """Performs a search in Content Server for the given employee number."""
        try:
            print(f"[{PreJoiningUtils.ts()}] Searching for employee: {employee_number}")

            # Click search icon
            await PreJoiningUtils.click_with_fallback(
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

            # Focus search input
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
            await search_input.fill("")
            await page.wait_for_timeout(300)
            await search_input.fill(employee_number)
            await page.wait_for_timeout(300)
            await search_input.press("Enter")
            await page.wait_for_timeout(500)

            # Fallback search button click
            try:
                search_btn = page.locator('a[title="Search in Content Server"], div.csui-search a').first
                if await search_btn.is_visible(timeout=1500):
                    await search_btn.click()
            except Exception:
                pass

            await page.wait_for_timeout(3500)

            # Check error dialog
            if await PreJoiningUtils.check_and_dismiss_error_dialog(page):
                print(f"[{PreJoiningUtils.ts()}]   Error dialog dismissed -- retrying search for {employee_number}")
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
                try:
                    zero_hits_loc = page.locator("text=/About 0 hits/i").first
                    if await zero_hits_loc.is_visible(timeout=2000):
                        print(f"[{PreJoiningUtils.ts()}]   Employee {employee_number} not found on Content Server (0 hits) -- skipping")
                        raise RuntimeError(
                            f"Employee {employee_number} not found on Content Server (0 search results)"
                        )
                except RuntimeError:
                    raise
                except Exception:
                    pass

                if await PreJoiningUtils.check_and_dismiss_error_dialog(page):
                    print(f"[{PreJoiningUtils.ts()}]   Error dialog appeared during search -- skipping {employee_number}")
                    raise RuntimeError(
                        f"Content Server error during search for employee {employee_number}"
                    )

                await PreJoiningUtils.save_debug_screenshot(
                    page, self.config.error_screenshot_dir, "no_results", employee_number
                )
                raise RuntimeError(
                    f"Search results did not appear for employee {employee_number} (timeout)"
                )
        except Exception as e:
            if not isinstance(e, RuntimeError):
                print(f"[{PreJoiningUtils.ts()}] Search exception for {employee_number}: {e}")
            raise

    async def click_employee_master_folder(self, page: Page, employee_number: str):
        """Finds and clicks the employee's master folder in search results."""
        try:
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
                        zero_hits_loc = page.locator("text=/About 0 hits/i").first
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

            # Strategy 1: Exact role match
            try:
                loc = page.get_by_role("link", name=employee_number, exact=True).first
                if await loc.is_visible(timeout=5000):
                    await loc.scroll_into_view_if_needed()
                    await loc.click()
                    await page.wait_for_timeout(3000)
                    return
            except Exception:
                pass

            # Strategy 2: Exact attribute match
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

            # Strategy 3: Iterate search items
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
        except Exception as e:
            if not isinstance(e, RuntimeError):
                print(f"[{PreJoiningUtils.ts()}] Click employee master folder error: {e}")
            raise

    async def click_prejoining_folder(self, page: Page, employee_number: str = ""):
        """Clicks 'Pre Joining and Talent Acquisition Stage' folder link inside employee folder."""
        try:
            try:
                await page.locator("td.csui-table-cell-name").first.wait_for(
                    state="visible", timeout=self.config.element_timeout
                )
            except PlaywrightTimeoutError:
                await PreJoiningUtils.save_debug_screenshot(
                    page, self.config.error_screenshot_dir, "no_folder_table", employee_number
                )
                raise RuntimeError("Employee folder table did not load")

            await page.wait_for_timeout(1500)

            folder_regex = re.compile(r"(?:pre[_\s\-]*joining|talent[_\s\-]*acquisition)", re.I)

            # Strategy 1: Regex / Role locator
            try:
                exact_folder = page.get_by_role("link", name=folder_regex).first
                if await exact_folder.is_visible(timeout=5000):
                    await exact_folder.scroll_into_view_if_needed()
                    await exact_folder.click()
                    await page.wait_for_timeout(3000)
                    return
            except Exception:
                pass

            # Strategy 2: Specific selectors
            selectors = [
                'a[title="Pre Joining and Talent Acquisition Stage"]',
                'a[aria-label="Pre Joining and Talent Acquisition Stage"]',
                'text="Pre Joining and Talent Acquisition Stage"',
                'xpath=//a[contains(.,"Pre Joining") and contains(.,"Talent Acquisition")]',
                'xpath=//span[contains(.,"Pre Joining") and contains(.,"Talent Acquisition")]',
                'xpath=//a[contains(translate(., "PRE JOINING", "pre joining"), "pre joining")]',
                'xpath=//a[contains(translate(., "TALENT ACQUISITION", "talent acquisition"), "talent acquisition")]',
                'a:has-text("Pre Joining")',
                'a:has-text("Talent Acquisition")',
                "tr.csui-state-hover > td.csui-table-cell-name a > span",
            ]

            try:
                await PreJoiningUtils.click_with_fallback(
                    page,
                    selectors,
                    "Pre Joining and Talent Acquisition Stage folder",
                    timeout_each=self.config.element_timeout,
                )
                await page.wait_for_timeout(3000)
                return
            except RuntimeError:
                pass

            # Strategy 3: Iterate table rows
            try:
                rows = page.locator("tbody tr td.csui-table-cell-name a, tbody tr td.csui-table-cell-name span")
                count = await rows.count()
                for i in range(count):
                    loc = rows.nth(i)
                    text = (await loc.inner_text()).strip()
                    if folder_regex.search(text):
                        print(f"[{PreJoiningUtils.ts()}]   Matched Pre-Joining folder row text: '{text}'")
                        await loc.scroll_into_view_if_needed()
                        await loc.click()
                        await page.wait_for_timeout(3000)
                        return
            except Exception:
                pass

            raise RuntimeError("Could not find 'Pre Joining and Talent Acquisition Stage' folder")
        except Exception as e:
            if not isinstance(e, RuntimeError):
                print(f"[{PreJoiningUtils.ts()}] Pre-Joining folder navigation error: {e}")
            raise

    async def get_document_count(self, page: Page) -> int:
        """Counts the number of document rows in the folder table."""
        try:
            await page.locator("tbody tr").first.wait_for(
                state="visible", timeout=self.config.element_timeout
            )
            await page.wait_for_timeout(1000)

            rows = page.locator("tbody tr td.csui-table-cell-name")
            count = await rows.count()
            print(f"[{PreJoiningUtils.ts()}] Found {count} item(s) in folder")
            return count
        except PlaywrightTimeoutError:
            print(f"[{PreJoiningUtils.ts()}]   No items found in the folder")
            return 0
        except Exception as e:
            print(f"[{PreJoiningUtils.ts()}] Document count error: {e}")
            return 0

    async def download_document_at_row(
        self, page: Page, row_index: int, total_rows: int, download_dir: Path, employee_number: str = ""
    ) -> Optional[str]:
        """Downloads a single document from folder table by 1-based row index."""
        row_selector = f"tbody tr:nth-of-type({row_index})"
        try:
            row = page.locator(row_selector).first

            try:
                await row.scroll_into_view_if_needed(timeout=5000)
            except Exception:
                pass

            await row.wait_for(state="visible", timeout=self.config.element_timeout)

            # Check row title against regex
            row_name = f"Row {row_index}"
            name_cell = row.locator("td.csui-table-cell-name a, td.csui-table-cell-name").first
            if await name_cell.count() > 0:
                row_name = (await name_cell.inner_text()).strip()

            if not PreJoiningUtils.is_prejoining_document(row_name):
                print(f"[{PreJoiningUtils.ts()}] Skipping non-matching document: '{row_name}' (row {row_index}/{total_rows})")
                return None

            download_dir.mkdir(parents=True, exist_ok=True)
            print(f"[{PreJoiningUtils.ts()}] Downloading matching document: '{row_name}' (row {row_index}/{total_rows})...")
            await row.hover()
            await page.wait_for_timeout(1000)
        except Exception as e:
            print(f"[{PreJoiningUtils.ts()}]   Could not analyze or hover row {row_index}: {e}")
            return None

        # Click More Actions (...) dropdown
        more_actions_selectors = [
            f"tbody tr:nth-of-type({row_index}) li.binf-dropdown > a",
            f'xpath=//tbody/tr[{row_index}]//li[contains(@class,"binf-dropdown")]/a',
            f'xpath=//tbody/tr[{row_index}]//a[@title="More actions"]',
        ]

        try:
            await PreJoiningUtils.click_with_fallback(
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
                    print(f"[{PreJoiningUtils.ts()}]   Clicked 'More actions' via fallback")
            except Exception as e2:
                print(f"[{PreJoiningUtils.ts()}]   Could not click 'More actions' for row {row_index}: {e2}")
                await PreJoiningUtils.save_debug_screenshot(
                    page, self.config.error_screenshot_dir, "more_actions_fail", employee_number
                )
                return None

        await page.wait_for_timeout(1500)

        # Click Download from dropdown menu
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
                                print(f"[{PreJoiningUtils.ts()}]   Used force-click on Download button")
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
                        print(f"[{PreJoiningUtils.ts()}]   Used JavaScript click on Download button")
                    except Exception:
                        pass

                if not download_clicked:
                    raise RuntimeError(f"Could not click Download for row {row_index} after all attempts")

            download = await dl_info.value
            failure = await download.failure()
            if failure:
                print(f"[{PreJoiningUtils.ts()}]   Download failed for row {row_index}: {failure}")
                return None

            suggested = download.suggested_filename or f"doc_{row_index}"

            # Sanitize filename and correct extension
            suggested_clean = suggested.strip()
            match = re.search(r"\.\s*(\d+)$", suggested_clean)
            if match:
                suffix_num = match.group(1)
                base = suggested_clean[:match.start()].strip().rstrip(".")
                suggested_clean = f"{base}_{suffix_num}.pdf"
            else:
                ext = Path(suggested_clean).suffix.lower()
                if ext not in [".pdf", ".xlsx", ".xls", ".doc", ".docx", ".png", ".jpg", ".zip"]:
                    suggested_clean = suggested_clean.rstrip(".")
                    suggested_clean = f"{suggested_clean}.pdf"

            save_path = download_dir / suggested_clean

            # Handle duplicate filenames
            if save_path.exists():
                stem = save_path.stem
                suffix = save_path.suffix
                counter = 1
                while save_path.exists():
                    save_path = download_dir / f"{stem}_{counter}{suffix}"
                    counter += 1

            await download.save_as(save_path)
            size = save_path.stat().st_size
            print(f"[{PreJoiningUtils.ts()}] Saved: {save_path.name} ({size:,} bytes)")
            return str(save_path)

        except PlaywrightTimeoutError:
            print(f"[{PreJoiningUtils.ts()}]   Download timed out for row {row_index}")
            await PreJoiningUtils.save_debug_screenshot(
                page, self.config.error_screenshot_dir, "download_timeout", employee_number
            )
            return None
        except Exception as e:
            print(f"[{PreJoiningUtils.ts()}]   Error downloading row {row_index}: {e}")
            return None

    async def download_prejoining_for_employee(
        self, page: Page, employee_number: str, download_dir: Path
    ) -> Tuple[str, List[str]]:
        """Executes full download flow for a single employee."""
        downloaded_files = []
        emp_dir = None
        status = self.config.status_error

        try:
            # Step 1: Search employee
            await self.search_employee(page, employee_number)

            # Step 2: Click employee master folder
            await self.click_employee_master_folder(page, employee_number)

            # Step 3: Navigate into Pre-Joining folder
            try:
                await self.click_prejoining_folder(page, employee_number)
            except RuntimeError as e:
                if "table did not load" in str(e).lower() or "pre joining" in str(e).lower():
                    print(f"[{PreJoiningUtils.ts()}] [WARN] Folder load failed. Clicking Home and retrying once...")
                    await self.navigate_home(page)
                    await page.wait_for_timeout(3000)
                    await self.search_employee(page, employee_number)
                    await self.click_employee_master_folder(page, employee_number)
                    await self.click_prejoining_folder(page, employee_number)
                else:
                    raise

            # Step 4: Count and download documents
            doc_count = await self.get_document_count(page)

            if doc_count == 0:
                print(f"[{PreJoiningUtils.ts()}]   No documents found in Pre-Joining folder for employee {employee_number}")
                status = self.config.status_no_docs
            else:
                emp_dir = download_dir / employee_number

                for row_idx in range(1, doc_count + 1):
                    await PreJoiningUtils.check_and_dismiss_error_dialog(page)

                    result = await self.download_document_at_row(
                        page, row_idx, doc_count, emp_dir, employee_number
                    )
                    if result:
                        downloaded_files.append(result)
                    await page.wait_for_timeout(2000)

                if downloaded_files:
                    status = self.config.status_downloaded
                else:
                    status = self.config.status_no_docs

            # Clean up error screenshots on success
            PreJoiningUtils.cleanup_error_screenshots(employee_number, self.config.error_screenshot_dir)

        except RuntimeError as e:
            err_msg = str(e)
            if "0 search results" in err_msg or "0 hits" in err_msg:
                print(f"[{PreJoiningUtils.ts()}] Employee {employee_number} not found on Content Server (0 hits)")
                status = self.config.status_not_found
                PreJoiningUtils.cleanup_error_screenshots(employee_number, self.config.error_screenshot_dir)
            else:
                print(f"[{PreJoiningUtils.ts()}] Error processing employee {employee_number}: {e}")
                await PreJoiningUtils.save_debug_screenshot(
                    page, self.config.error_screenshot_dir, "error", employee_number
                )
                status = self.config.status_error
        except Exception as e:
            print(f"[{PreJoiningUtils.ts()}] Error processing employee {employee_number}: {e}")
            await PreJoiningUtils.save_debug_screenshot(
                page, self.config.error_screenshot_dir, "error", employee_number
            )
            status = self.config.status_error

        return status, downloaded_files

    def load_existing_master_excel(self) -> Dict[str, Dict[str, str]]:
        """Loads existing employee records from Master_Pre_Joining_Summary.xlsx."""
        try:
            self.config.download_dir.mkdir(parents=True, exist_ok=True)
            if not self.config.master_excel_path.exists():
                return {}

            df = pd.read_excel(
                str(self.config.master_excel_path), sheet_name="Master_Summary", dtype={"Employee ID": str}
            )
            records = {}
            for _, row in df.iterrows():
                eid = str(row.get("Employee ID", "")).strip()
                if eid and eid != "nan":
                    records[eid] = {
                        "Employee ID": eid,
                        "Offer Letter": str(row.get("Offer Letter", "No")).strip(),
                        "Background Verification": str(row.get("Background Verification", "No")).strip(),
                        "Appointment Letter": str(row.get("Appointment Letter", "No")).strip(),
                    }
            print(f"[{PreJoiningUtils.ts()}] Loaded {len(records):,} existing employee record(s) from Master Excel.")
            return records
        except Exception as e:
            print(f"[{PreJoiningUtils.ts()}] [WARN] Could not read existing Master Excel ({e}). A fresh one will be started.")
            return {}

    def record_employee_to_master_excel(
        self,
        master_records: Dict[str, Dict[str, str]],
        emp_id: str,
        status: str,
        file_paths: List[str],
    ):
        """Updates Master_Pre_Joining_Summary.xlsx incrementally in real-time."""
        try:
            emp_str = str(emp_id).strip()
            has_offer = "No"
            has_bgv = "No"
            has_appointment = "No"

            prev = master_records.get(emp_str, {})

            for fpath in file_paths:
                fname = Path(fpath).name
                categories = PreJoiningUtils.classify_document(fname)
                if "Offer Letter" in categories:
                    has_offer = "Yes"
                if "Background Verification" in categories:
                    has_bgv = "Yes"
                if "Appointment Letter" in categories:
                    has_appointment = "Yes"

            # Retain prior Yes
            if has_offer == "No" and prev.get("Offer Letter") == "Yes":
                has_offer = "Yes"
            if has_bgv == "No" and prev.get("Background Verification") == "Yes":
                has_bgv = "Yes"
            if has_appointment == "No" and prev.get("Appointment Letter") == "Yes":
                has_appointment = "Yes"

            master_records[emp_str] = {
                "Employee ID": emp_str,
                "Offer Letter": has_offer,
                "Background Verification": has_bgv,
                "Appointment Letter": has_appointment,
            }

            df = pd.DataFrame(list(master_records.values()))
            cols = ["Employee ID", "Offer Letter", "Background Verification", "Appointment Letter"]
            df = df.reindex(columns=cols)

            self.config.download_dir.mkdir(parents=True, exist_ok=True)
            with pd.ExcelWriter(str(self.config.master_excel_path), engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Master_Summary")

                worksheet = writer.sheets["Master_Summary"]
                for col_idx, col_name in enumerate(df.columns, start=1):
                    max_len = max(
                        len(str(col_name)),
                        df[col_name].astype(str).str.len().max() if not df.empty else 0,
                    )
                    col_letter = worksheet.cell(row=1, column=col_idx).column_letter
                    worksheet.column_dimensions[col_letter].width = max(max_len + 5, 16)

            print(
                f"[{PreJoiningUtils.ts()}]   [Master Excel Live] {emp_str} -> "
                f"Offer: {has_offer} | BGV: {has_bgv} | Appointment: {has_appointment} "
                f"({len(master_records)} total in master)"
            )
        except Exception as e:
            print(f"[{PreJoiningUtils.ts()}]   [ERROR] Failed real-time Master Excel write for {emp_id}: {e}")

    def generate_excel_summary(self, all_results: Dict[str, Tuple[str, List[str]]]):
        """Verifies and confirms final Master Excel summary file."""
        try:
            if self.config.master_excel_path.exists():
                df = pd.read_excel(str(self.config.master_excel_path), sheet_name="Master_Summary", dtype=str)
                print(f"[{PreJoiningUtils.ts()}] Final Master Excel verified: {self.config.master_excel_path} ({len(df):,} total employee(s))")
        except Exception as e:
            print(f"[{PreJoiningUtils.ts()}] Excel verification warning: {e}")

    def print_categorized_summary(self, all_results: Dict[str, Tuple[str, List[str]]]):
        """Prints formatted console and log summary of downloaded files."""
        try:
            downloaded = {k: v for k, (s, v) in all_results.items() if s == self.config.status_downloaded}
            no_docs = {k: v for k, (s, v) in all_results.items() if s == self.config.status_no_docs}
            not_found = {k: v for k, (s, v) in all_results.items() if s == self.config.status_not_found}
            errors = {k: v for k, (s, v) in all_results.items() if s == self.config.status_error}

            total_files = sum(len(v) for v in downloaded.values())
            total_emps = len(all_results)

            print()
            print(f"[{PreJoiningUtils.ts()}] ═══════════════════════════════════════════════════════════")
            print(f"[{PreJoiningUtils.ts()}] PRE-JOINING DOWNLOAD SUMMARY")
            print(f"[{PreJoiningUtils.ts()}] ═══════════════════════════════════════════════════════════")

            if downloaded:
                print(f"[{PreJoiningUtils.ts()}]")
                print(f"[{PreJoiningUtils.ts()}]   [OK] DOWNLOADED ({len(downloaded)} employee(s), {total_files} file(s)):")
                for emp, files in downloaded.items():
                    print(f"[{PreJoiningUtils.ts()}]      {emp}: {len(files)} file(s)")

            if no_docs:
                print(f"[{PreJoiningUtils.ts()}]")
                print(f"[{PreJoiningUtils.ts()}]   [INFO] NO MATCHING PRE-JOINING DOCUMENTS ({len(no_docs)} employee(s) -- folder found, no matching document uploaded):")
                for emp in no_docs:
                    print(f"[{PreJoiningUtils.ts()}]      {emp}")

            if not_found:
                print(f"[{PreJoiningUtils.ts()}]")
                print(f"[{PreJoiningUtils.ts()}]   [--] NOT FOUND ON CONTENT SERVER ({len(not_found)} employee(s) -- 0 search hits):")
                for emp in not_found:
                    print(f"[{PreJoiningUtils.ts()}]      {emp}")

            if errors:
                print(f"[{PreJoiningUtils.ts()}]")
                print(f"[{PreJoiningUtils.ts()}]   [FAIL] FAILED / ERRORS ({len(errors)} employee(s) -- search timeout or network failure):")
                for emp in errors:
                    print(f"[{PreJoiningUtils.ts()}]      {emp}")

            print(f"[{PreJoiningUtils.ts()}] ───────────────────────────────────────────────────────────")
            print(f"[{PreJoiningUtils.ts()}]   Total: {total_files} file(s) downloaded for {total_emps} employee(s)")
            if errors:
                print(f"[{PreJoiningUtils.ts()}]   [WARN] {len(errors)} employee(s) still have unresolved errors -- check errors/prejoining_error/ for screenshots")
            print(f"[{PreJoiningUtils.ts()}] ═══════════════════════════════════════════════════════════")
        except Exception as e:
            print(f"[{PreJoiningUtils.ts()}] Summary print error: {e}")

    async def download_prejoining_reports(
        self, employee_numbers: List[str], upload_callback=None
    ):
        """Main execution engine. Concurrently downloads documents across browser tabs with retry passes."""
        try:
            self.config.download_dir.mkdir(parents=True, exist_ok=True)

            if not employee_numbers:
                print(f"[{PreJoiningUtils.ts()}] No employee numbers provided. Exiting.")
                return

            if not self.config.content_server_url:
                print(f"[{PreJoiningUtils.ts()}] [FATAL ERROR] CONTENT_SERVER_URL is blank in your .env file!")
                print(f"[{PreJoiningUtils.ts()}] Please open .env and add: CONTENT_SERVER_URL=https://...")
                return

            concurrency = min(self.config.max_concurrent_tabs, len(employee_numbers))
            print(f"[{PreJoiningUtils.ts()}] Starting Pre-Joining document download ({'headless' if self.config.headless else 'visible'})")
            print(f"[{PreJoiningUtils.ts()}] Concurrency: {concurrency} tab(s) for {len(employee_numbers)} employee(s)")

            PreJoiningUtils.cleanup_profile_locks(self.config.automation_profile_dir)

            all_results: Dict[str, Tuple[str, List[str]]] = {}
            master_records: Dict[str, Dict[str, str]] = self.load_existing_master_excel()
            excel_lock = asyncio.Lock()

            async def _run_batch(
                p_context,
                emp_list: List[str],
                num_tabs: int,
                nav_lock: asyncio.Lock,
                total_label: int,
                start_offset: int = 0,
                upload_cb=None,
            ):
                first_page = p_context.pages[0] if p_context.pages else await p_context.new_page()
                worker_pages = [first_page]
                for i in range(1, num_tabs):
                    try:
                        new_tab = await p_context.new_page()
                        worker_pages.append(new_tab)
                        print(f"[{PreJoiningUtils.ts()}]   Opened tab {i + 1}/{num_tabs}")
                        await asyncio.sleep(0.5)
                    except Exception as e:
                        print(f"[{PreJoiningUtils.ts()}]   Could not open tab {i + 1}: {e}")
                        break

                actual_tabs = len(worker_pages)
                print(f"[{PreJoiningUtils.ts()}] {actual_tabs} tab(s) ready. Starting downloads...\n")

                async def _worker(worker_page, emp_num: str, idx: int):
                    try:
                        print(f"[{PreJoiningUtils.ts()}] ── [Tab] Employee {idx}/{total_label}: {emp_num} ──")

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
                                    await self.wait_for_content_server_home(worker_page, timeout_s=30)
                                nav_success = True
                                break
                            except Exception as nav_err:
                                if nav_attempt == 0:
                                    print(f"[{PreJoiningUtils.ts()}]   [Tab] Navigation timeout for {emp_num} -- retrying once...")
                                    await worker_page.wait_for_timeout(3000)
                                else:
                                    raise RuntimeError(
                                        f"Content Server navigation failed after 2 attempts for {emp_num}: {nav_err}"
                                    )

                        if not nav_success:
                            raise RuntimeError(f"Could not navigate to Content Server for {emp_num}")

                        status, downloaded = await self.download_prejoining_for_employee(
                            worker_page, emp_num, self.config.download_dir
                        )
                        all_results[emp_num] = (status, downloaded)

                        async with excel_lock:
                            self.record_employee_to_master_excel(master_records, emp_num, status, downloaded)

                        if downloaded and upload_cb:
                            try:
                                print(f"[{PreJoiningUtils.ts()}] [Tab] Triggering upload for {emp_num}...")
                                await upload_cb(emp_num, downloaded)
                            except Exception as e:
                                print(f"[{PreJoiningUtils.ts()}] [Tab] Upload callback failed for {emp_num}: {e}")

                    except Exception as e:
                        print(f"[{PreJoiningUtils.ts()}] [Tab] Error for {emp_num}: {e}")
                        all_results[emp_num] = (self.config.status_error, [])
                        async with excel_lock:
                            self.record_employee_to_master_excel(master_records, emp_num, self.config.status_error, [])

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

            context = None
            try:
                async with async_playwright() as p:
                    print(f"[{PreJoiningUtils.ts()}] Launching browser...")
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

                    context = await p.chromium.launch_persistent_context(
                        user_data_dir=str(self.config.automation_profile_dir),
                        channel=self.config.browser_channel,
                        accept_downloads=True,
                        args=chrome_args,
                        **launch_kwargs,
                    )

                    first_page = context.pages[0] if context.pages else await context.new_page()
                    if not self.config.headless:
                        try:
                            await first_page.bring_to_front()
                        except Exception:
                            pass

                    print(f"[{PreJoiningUtils.ts()}] Navigating to Content Server...")
                    await first_page.goto(
                        self.config.content_server_url, wait_until="domcontentloaded", timeout=self.config.nav_timeout
                    )
                    await first_page.wait_for_timeout(3000)

                    await self.wait_for_content_server_home(first_page, timeout_s=120)
                    await first_page.wait_for_timeout(2000)
                    print(f"[{PreJoiningUtils.ts()}] Login confirmed. Preparing {concurrency} tab(s)...\n")

                    nav_lock = asyncio.Lock()

                    # Primary batch
                    await _run_batch(
                        context,
                        employee_numbers,
                        concurrency,
                        nav_lock,
                        len(employee_numbers),
                        upload_cb=upload_callback,
                    )

                    # Retry loop
                    initial_not_found = [
                        emp for emp, (st, _) in all_results.items() if st == self.config.status_not_found
                    ]

                    for retry_round in range(1, self.config.max_retry_rounds + 1):
                        folder_failed_ids = PreJoiningUtils.get_failed_ids_from_error_folder(
                            self.config.error_screenshot_dir
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
                            print(f"[{PreJoiningUtils.ts()}] errors/prejoining_error/ is empty -- no retries needed.")
                            break

                        retry_concurrency = min(2, len(current_retry_ids))
                        print()
                        print(f"[{PreJoiningUtils.ts()}] ═══════════════════════════════════════════════════════════")
                        print(
                            f"[{PreJoiningUtils.ts()}] RETRY ROUND {retry_round}/{self.config.max_retry_rounds}: "
                            f"{len(current_retry_ids)} employee(s) queued for retry"
                        )
                        print(f"[{PreJoiningUtils.ts()}]   IDs: {', '.join(current_retry_ids)}")
                        print(f"[{PreJoiningUtils.ts()}]   Retrying with {retry_concurrency} tab(s) for reliability...")
                        print(f"[{PreJoiningUtils.ts()}] ═══════════════════════════════════════════════════════════")
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
                                await self.wait_for_content_server_home(retry_page, timeout_s=30)
                        except Exception:
                            pass

                        await _run_batch(
                            context,
                            current_retry_ids,
                            retry_concurrency,
                            nav_lock,
                            len(current_retry_ids),
                            upload_cb=upload_callback,
                        )

                        resolved_count = 0
                        for emp_num in current_retry_ids:
                            retry_status, _ = all_results.get(emp_num, (self.config.status_error, []))
                            if retry_status != self.config.status_error:
                                print(f"[{PreJoiningUtils.ts()}]   [OK] {emp_num} resolved on retry (status: {retry_status})")
                                PreJoiningUtils.cleanup_error_screenshots(emp_num, self.config.error_screenshot_dir)
                                resolved_count += 1

                        print(
                            f"[{PreJoiningUtils.ts()}]   Round {retry_round} complete: "
                            f"{resolved_count}/{len(current_retry_ids)} resolved"
                        )

                        remaining = PreJoiningUtils.get_failed_ids_from_error_folder(
                            self.config.error_screenshot_dir
                        )
                        if not remaining:
                            print(f"[{PreJoiningUtils.ts()}] errors/prejoining_error/ is now empty -- all retries succeeded!")
                            break
                        elif retry_round < self.config.max_retry_rounds:
                            print(f"[{PreJoiningUtils.ts()}]   {len(remaining)} still remaining, will retry next round...")
                        else:
                            print(f"[{PreJoiningUtils.ts()}]   {len(remaining)} still remaining after {self.config.max_retry_rounds} retry rounds.")

            except Exception as e:
                print(f"[{PreJoiningUtils.ts()}] FATAL: {e}\n{traceback.format_exc()}")
            finally:
                if context:
                    for pg in context.pages:
                        try:
                            await pg.close()
                        except Exception:
                            pass
                    try:
                        await context.close()
                        print(f"[{PreJoiningUtils.ts()}] Browser context closed cleanly.")
                    except Exception as e:
                        print(f"[{PreJoiningUtils.ts()}] Warning closing browser context: {e}")

            self.print_categorized_summary(all_results)
            self.generate_excel_summary(all_results)
        except Exception as e:
            print(f"[{PreJoiningUtils.ts()}] Download process fatal error: {e}\n{traceback.format_exc()}")


# Input Prompt Handler
class InputPromptHandler:
    """Handles terminal user interaction and employee ID extraction from files."""

    @staticmethod
    def load_employee_ids_from_file(file_path: Path) -> List[str]:
        """Reads employee IDs from Excel (.xlsx, .xls, .xlsm) or CSV file."""
        try:
            ext = file_path.suffix.lower()
            if ext in [".xlsx", ".xlsm"]:
                df = pd.read_excel(file_path, dtype=str, engine="openpyxl")
            elif ext == ".xls":
                df = pd.read_excel(file_path, dtype=str)
            elif ext == ".csv":
                df = pd.read_csv(file_path, dtype=str)
            else:
                raise ValueError(f"Unsupported file format '{ext}'. Please provide a .xlsx, .xls, or .csv file.")

            if df.empty:
                return []

            target_col = None
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
                for col in df.columns:
                    if re.search(pat, str(col).strip(), re.IGNORECASE):
                        target_col = col
                        break
                if target_col:
                    break

            if not target_col:
                target_col = df.columns[0]
                print(f"[{PreJoiningUtils.ts()}] No explicit 'Employee Code' header found; using first column: '{target_col}'")
            else:
                print(f"[{PreJoiningUtils.ts()}] Identified employee column: '{target_col}'")

            ids = []
            seen = set()
            for val in df[target_col].dropna():
                clean_val = str(val).strip()
                if clean_val.endswith(".0"):
                    clean_val = clean_val[:-2]
                clean_val = clean_val.strip()
                if clean_val and clean_val.lower() != "nan" and clean_val not in seen:
                    seen.add(clean_val)
                    ids.append(clean_val)

            return ids
        except Exception as e:
            print(f"[{PreJoiningUtils.ts()}] File load error: {e}")
            return []

    @staticmethod
    def prompt_employee_numbers(error_dir: Path) -> List[str]:
        """Interactively collects employee IDs or files from terminal input."""
        try:
            terminal = sys.__stdout__

            terminal.write("\n")
            terminal.write("  ╔══════════════════════════════════════════════════════════════════════╗\n")
            terminal.write("  ║       Pre-Joining & Talent Acquisition Document Downloader           ║\n")
            terminal.write("  ╚══════════════════════════════════════════════════════════════════════╝\n")
            terminal.write("\n")
            terminal.flush()

            failed_in_folder = PreJoiningUtils.get_failed_ids_from_error_folder(error_dir)
            if failed_in_folder:
                terminal.write(f"  [NOTE] {len(failed_in_folder)} failed employee screenshot(s) found in errors/prejoining_error/:\n")
                terminal.write(f"         {', '.join(failed_in_folder[:15])}{'...' if len(failed_in_folder) > 15 else ''}\n")
                terminal.write("         (Press Enter without typing to automatically retry these failed employees)\n\n")
                terminal.flush()

            try:
                terminal.write("  Enter Excel/CSV file path OR Employee ID(s): ")
                terminal.flush()
                raw = sys.__stdin__.readline()
            except (EOFError, KeyboardInterrupt):
                terminal.write("\n")
                return []

            raw = raw.strip().strip("'\"")
            if not raw:
                if failed_in_folder:
                    terminal.write(f"\n  -> Retrying {len(failed_in_folder)} failed employee(s) from error folder: {', '.join(failed_in_folder)}\n\n")
                    terminal.flush()
                    return failed_in_folder
                terminal.write("\n  -> No input entered.\n\n")
                terminal.flush()
                return []

            possible_path = Path(raw)
            if not possible_path.exists():
                alt_path = Path(__file__).parent.parent / raw
                if alt_path.exists():
                    possible_path = alt_path

            if possible_path.exists() and possible_path.suffix.lower() in [".xlsx", ".xls", ".xlsm", ".csv"]:
                try:
                    terminal.write(f"\n  -> Reading employee codes from file: {possible_path.name}...\n")
                    terminal.flush()
                    extracted_ids = InputPromptHandler.load_employee_ids_from_file(possible_path)
                    if extracted_ids:
                        terminal.write(f"  -> Successfully loaded {len(extracted_ids):,} unique employee ID(s) from {possible_path.name}!\n")
                        terminal.write(f"  -> Sample: {', '.join(extracted_ids[:8])}{'...' if len(extracted_ids) > 8 else ''}\n\n")
                        terminal.flush()
                        return extracted_ids
                    else:
                        terminal.write(f"  [WARN] No valid employee IDs found in {possible_path.name}.\n\n")
                        terminal.flush()
                        return []
                except Exception as e:
                    terminal.write(f"  [ERROR] Failed to read file {possible_path.name}: {e}\n\n")
                    terminal.flush()
                    return []

            collected = []
            for token in raw.replace(",", " ").split():
                token = token.strip().strip("'\"")
                if token:
                    collected.append(token)

            seen = set()
            unique = []
            for num in collected:
                if num not in seen:
                    seen.add(num)
                    unique.append(num)

            if unique:
                terminal.write(f"\n  -> {len(unique):,} employee(s) queued: {', '.join(unique[:10])}{'...' if len(unique) > 10 else ''}\n\n")
                terminal.flush()
            else:
                terminal.write("\n  -> No employee numbers entered.\n\n")
                terminal.flush()

            return unique
        except Exception as e:
            print(f"[{PreJoiningUtils.ts()}] Terminal prompt error: {e}")
            return []


# Main execution entrypoint
async def download_prejoining_reports(employee_numbers: List[str], upload_callback=None):
    """Main execution entrypoint for downloading Pre-Joining reports."""
    config = PreJoiningConfig()
    downloader = PreJoiningReportDownloader(config)
    await downloader.download_prejoining_reports(employee_numbers, upload_callback=upload_callback)


if __name__ == "__main__":
    config = PreJoiningConfig()
    numbers = InputPromptHandler.prompt_employee_numbers(config.error_screenshot_dir)
    if numbers:
        asyncio.run(download_prejoining_reports(numbers))
    else:
        print("No employee numbers entered. Exiting.")

