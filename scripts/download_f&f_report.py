# OpenText Content Server - F&F (Full & Final) Exit Document Downloader

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

from dotenv import load_dotenv
from playwright.async_api import Page, TimeoutError as PlaywrightTimeoutError
from playwright.async_api import async_playwright

# Suppress Chrome background subprocess console window on Windows
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
    """Redirects stdout and stderr streams to both console and a log file."""

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
LOG_FILE = LOG_DIR / "download_ff_report.log"

sys.stdout = LoggerWriter(LOG_FILE)
sys.stderr = LoggerWriter(LOG_FILE)

load_dotenv(verbose=True)


# Application Configuration
class FFConfig:
    """Encapsulates environment variables, URLs, paths, timeouts, and constants."""

    # Result status constants
    STATUS_DOWNLOADED = "DOWNLOADED"
    STATUS_NO_FF_DOCS = "NO_FF_DOCS"
    STATUS_NOT_FOUND = "NOT_FOUND_0_HITS"
    STATUS_ERROR = "ERROR"

    def __init__(self):
        try:
            # SSO credentials and Content Server URL
            self.oracle_email: str = (os.getenv("ORACLE_EMAIL") or "").strip()
            self.oracle_password: str = (os.getenv("ORACLE_PASSWORD") or "").strip()
            self.content_server_url: str = (os.getenv("CONTENT_SERVER_URL") or "").strip()

            # Output and automation profile directories
            self.download_dir: Path = Path(__file__).parent.parent / "uploads" / "FF_Reports"
            self.automation_profile_dir: Path = Path(__file__).parent.parent / "chrome_automation_profile"
            self.error_dir: Path = Path(__file__).parent.parent / "errors" / "ff_error"

            # Browser execution settings
            self.headless: bool = self._parse_bool(os.getenv("HEADLESS"), default=True)
            self.start_maximized: bool = self._parse_bool(os.getenv("START_MAXIMIZED"), default=True)
            self.window_position: str = (os.getenv("WINDOW_POSITION") or "").strip()
            self.browser_channel: Optional[str] = (os.getenv("BROWSER_CHANNEL") or "").strip() or None

            # Timeouts and concurrency
            self.element_timeout: int = 15_000
            self.nav_timeout: int = 90_000
            self.max_concurrent_tabs: int = 5
            self.search_results_wait: int = 45_000
            self.max_retry_rounds: int = 5

            # Ensure directories exist
            self.download_dir.mkdir(parents=True, exist_ok=True)
            self.error_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Configuration Initialization Error: {e}")

    @staticmethod
    def _parse_bool(val: Optional[str], default: bool = True) -> bool:
        """Parses boolean strings (true/false/1/0/yes/no) safely."""
        try:
            if val is None:
                return default
            return str(val).strip().lower() in ("true", "1", "yes", "t", "y")
        except Exception:
            return default


# Automation Utilities
class FFAutomationUtils:
    """Helper utilities for timestamps, profile locks, screenshots, dialogs, and DOM interactions."""

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

    @staticmethod
    def ts() -> str:
        """Returns the current formatted timestamp string."""
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    @staticmethod
    def cleanup_profile_locks(profile_dir: Path):
        """Removes stale Chromium profile lock files to ensure clean browser startup."""
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
            print(f"[{FFAutomationUtils.ts()}] [CLEANUP] Lock cleanup warning: {e}")

    @staticmethod
    def cleanup_error_screenshots(error_dir: Path, employee_number: str):
        """Deletes all debug error screenshots for the specified employee from the error directory."""
        try:
            if not error_dir.exists():
                return
            patterns = [
                str(error_dir / f"debug_*_{employee_number}.png"),
                str(error_dir / f"debug_error_{employee_number}_*.png"),
            ]
            for pattern in patterns:
                for path in glob.glob(pattern):
                    try:
                        os.remove(path)
                        print(f"[{FFAutomationUtils.ts()}]   Cleaned up error screenshot: {Path(path).name}")
                    except OSError:
                        pass
        except Exception as e:
            print(f"[{FFAutomationUtils.ts()}]   Warning cleaning up screenshots for {employee_number}: {e}")

    @staticmethod
    def extract_employee_id_from_debug_screenshot(path: Path) -> Optional[str]:
        """Extracts the employee ID from debug_{label}_{employee_id}.png filenames."""
        try:
            match = re.match(r"^debug_[a-z_]+_(\d+)$", path.stem, flags=re.IGNORECASE)
            if match:
                return match.group(1)

            # Backward compatibility for timestamped screenshots
            match = re.match(r"^debug_error_(\d+)_\d+$", path.stem, flags=re.IGNORECASE)
            if match:
                return match.group(1)
        except Exception:
            pass
        return None

    @staticmethod
    def debug_screenshot_path(error_dir: Path, label: str, employee_number: Optional[str] = None) -> Path:
        """Constructs a sanitized debug screenshot path."""
        try:
            safe_label = re.sub(r"[^A-Za-z_]+", "_", label).strip("_").lower()
            if employee_number:
                return error_dir / f"debug_{safe_label}_{employee_number}.png"
            return error_dir / f"debug_{safe_label}.png"
        except Exception:
            return error_dir / f"debug_{int(datetime.now().timestamp())}.png"

    @staticmethod
    async def save_debug_screenshot(
        page: Page, error_dir: Path, label: str, employee_number: Optional[str] = None
    ) -> str:
        """Captures and saves a full-page diagnostic screenshot on failure."""
        try:
            error_dir.mkdir(parents=True, exist_ok=True)
            path = FFAutomationUtils.debug_screenshot_path(error_dir, label, employee_number)
            await page.screenshot(path=str(path), full_page=True)
            print(f"[{FFAutomationUtils.ts()}]   Screenshot saved: {path.name}")
            return str(path)
        except Exception as e:
            print(f"[{FFAutomationUtils.ts()}]   Could not save screenshot: {e}")
            return ""

    @staticmethod
    def get_failed_ids_from_error_folder(error_dir: Path) -> List[str]:
        """Scans the error directory and extracts unique employee IDs from screenshot filenames."""
        try:
            if not error_dir.exists():
                return []
            seen = set()
            ids = []
            for path in error_dir.glob("debug_*.png"):
                emp_id = FFAutomationUtils.extract_employee_id_from_debug_screenshot(path)
                if emp_id and emp_id not in seen:
                    seen.add(emp_id)
                    ids.append(emp_id)
            return ids
        except Exception as e:
            print(f"[{FFAutomationUtils.ts()}]   Error scanning failed IDs from error folder: {e}")
            return []

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
                        print(f"[{FFAutomationUtils.ts()}] [WARN] Content Server error dialog detected: {error_text}")

                        close_selectors = [
                            'button:has-text("Close")',
                            'button.binf-btn:has-text("Close")',
                            "a.binf-close",
                            "button.csui-btn-close",
                            'button[aria-label="Close"]',
                        ]
                        for close_sel in close_selectors:
                            try:
                                close_btn = page.locator(close_sel).first
                                if await close_btn.is_visible(timeout=2000):
                                    await close_btn.click(force=True)
                                    print(f"[{FFAutomationUtils.ts()}]   Dismissed error dialog")
                                    await page.wait_for_timeout(1500)
                                    return True
                            except Exception:
                                continue

                        await page.keyboard.press("Escape")
                        print(f"[{FFAutomationUtils.ts()}]   Dismissed error dialog via Escape key")
                        await page.wait_for_timeout(1500)
                        return True
                except Exception:
                    continue
        except Exception:
            pass
        return False

    @staticmethod
    def is_ff_document(name: str) -> bool:
        """Returns True only if the document name matches the F&F pattern regex."""
        try:
            return bool(FFAutomationUtils.FF_DOCUMENT_PATTERN.search(name))
        except Exception:
            return False

    @staticmethod
    async def safe_click(locator, description: str, timeout: int = 15_000):
        """Waits for element visibility, scrolls into view, and clicks safely."""
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
            f"No working selector for '{description}'. Tried:\n" + "\n".join(errors)
        )


# Microsoft SSO Handler
class MicrosoftSSOHandler:
    """Manages Microsoft Entra SSO, RSA MFA push verification, and KMSI prompt."""

    def __init__(self, config: FFConfig):
        self.config = config
        self.email_filled = False
        self.password_filled = False
        self.rsa_clicked = False
        self.rsa_retry_count = 0
        self.max_rsa_retries = 3

    async def check_for_sso_errors(self, page: Page):
        """Checks if Microsoft SSO displays validation or security error messages."""
        try:
            for error_id in ["usernameError", "passwordError", "error"]:
                try:
                    loc = page.locator(f"#{error_id}").first
                    if await loc.count() > 0 and await loc.is_visible(timeout=500):
                        err_text = (await loc.inner_text()).strip().replace("\n", " ")
                        if err_text:
                            raise RuntimeError(f"Microsoft SSO Error: {err_text}")
                except RuntimeError:
                    raise
                except Exception:
                    pass

            try:
                alert_loc = page.locator('[role="alert"]').first
                if await alert_loc.count() > 0 and await alert_loc.is_visible(timeout=500):
                    text = (await alert_loc.inner_text()).strip().replace("\n", " ")
                    if text:
                        raise RuntimeError(f"Microsoft SSO Alert: {text}")
            except RuntimeError:
                raise
            except Exception:
                pass
        except RuntimeError:
            raise
        except Exception:
            pass

    async def detect_screen(self, page: Page) -> str:
        """Identifies the active Microsoft SSO authentication screen."""
        try:
            current_url = page.url.lower()

            if not any(
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
                return "REDIRECTED"

            # 1. KMSI ("Stay signed in?") prompt
            try:
                kmsi_header = page.locator(
                    '#loginHeader:has-text("Stay signed in?"), '
                    'div[role="heading"]:has-text("Stay signed in?"), '
                    'div:has-text("Stay signed in?"), '
                    'div:has-text("Do this to reduce the number of times")'
                ).first
                kmsi_checkbox = page.locator(
                    '#KmsiCheckboxField, input[name="DontShowAgain"]'
                ).first
                if (await kmsi_header.count() > 0 and await kmsi_header.is_visible(timeout=300)) or (
                    await kmsi_checkbox.count() > 0 and await kmsi_checkbox.is_visible(timeout=300)
                ):
                    return "KMSI"
            except Exception:
                pass

            # 2. RSAEntraMFA method selection
            try:
                rsa_btn = page.locator('div[role="button"], button, a').filter(
                    has_text="RSAEntraMFA"
                ).first
                if await rsa_btn.count() > 0 and await rsa_btn.is_visible(timeout=300):
                    return "RSA_SELECTION"
            except Exception:
                pass

            # 3. RSA phone push approval waiting
            if any(x in current_url for x in ["securid.com", "adani.auth"]) or self.rsa_clicked:
                return "RSA_WAIT"

            # 4. Password input screen
            try:
                pw_input = page.locator(
                    'input[name="passwd"], input[id="i0118"], input[type="password"]'
                ).first
                if await pw_input.count() > 0 and await pw_input.is_visible(timeout=300):
                    return "PASSWORD"
            except Exception:
                pass

            # 5. Account picker screen
            try:
                pick_header = page.locator(
                    '#loginHeader:has-text("Pick an account"), div[role="heading"]:has-text("Pick an account")'
                ).first
                if await pick_header.count() > 0 and await pick_header.is_visible(timeout=300):
                    return "PICK_ACCOUNT"
            except Exception:
                pass

            # 6. Email input screen
            try:
                email_input = page.locator(
                    'input[name="loginfmt"], input[type="email"], input[id="i0116"]'
                ).first
                if await email_input.count() > 0 and await email_input.is_visible(timeout=300):
                    return "EMAIL"
            except Exception:
                pass

            return "UNKNOWN"
        except Exception:
            return "UNKNOWN"

    async def _handle_pick_account(self, page: Page):
        """Selects the target account tile if listed on account picker screen."""
        try:
            target_email = self.config.oracle_email.strip().lower()
            account_found = False

            account_selectors = [
                'div.table[role="button"]',
                "div[data-test-id]",
                "div.row.tile",
                "li.table",
            ]
            for sel in account_selectors:
                try:
                    accounts = page.locator(sel)
                    count = await accounts.count()
                    for i in range(count):
                        tile = accounts.nth(i)
                        tile_text = (await tile.inner_text()).strip().lower()
                        if target_email in tile_text:
                            print(f"[{FFAutomationUtils.ts()}]   Selecting account: {self.config.oracle_email}")
                            await tile.click(force=True, timeout=5000)
                            account_found = True
                            break
                    if account_found:
                        break
                except Exception:
                    continue

            if not account_found:
                try:
                    el = page.locator(
                        f'small:has-text("{self.config.oracle_email}"), div:has-text("{self.config.oracle_email}")'
                    ).first
                    if await el.is_visible(timeout=2000):
                        parent = el.locator(
                            'xpath=ancestor::div[@role="button" or contains(@class,"table") or contains(@class,"tile")]'
                        ).first
                        if await parent.count() > 0:
                            await parent.click(force=True, timeout=5000)
                        else:
                            await el.click(force=True, timeout=5000)
                        account_found = True
                        print(f"[{FFAutomationUtils.ts()}]   Selected account: {self.config.oracle_email}")
                except Exception:
                    pass

            if not account_found:
                try:
                    use_another = page.locator('div#otherTile, div:has-text("Use another account")').last
                    if await use_another.is_visible(timeout=3000):
                        await use_another.click(force=True, timeout=5000)
                except Exception:
                    pass

            await page.wait_for_timeout(2000)
        except Exception as e:
            print(f"[{FFAutomationUtils.ts()}]   Warning during account pick: {e}")

    async def _handle_email(self, page: Page):
        """Fills the email field and submits."""
        try:
            email_input = page.locator(
                'input[name="loginfmt"], input[type="email"], input[id="i0116"]'
            ).first
            if await email_input.is_visible(timeout=2000):
                for attempt in range(1, 4):
                    try:
                        await email_input.click(force=True, timeout=3000)
                        await email_input.fill("", force=True, timeout=3000)
                        await page.wait_for_timeout(200)
                        await email_input.type(self.config.oracle_email.strip(), delay=40)
                        print(f"[{FFAutomationUtils.ts()}]   Entered email ({self.config.oracle_email})")
                        await page.wait_for_timeout(300)

                        next_btn = page.locator('input[type="submit"], #idSIButton9').first
                        await next_btn.click(force=True, timeout=5000)
                        await page.wait_for_timeout(2500)

                        await self.check_for_sso_errors(page)
                        self.email_filled = True
                        break
                    except RuntimeError:
                        if attempt == 3:
                            raise
                        print(f"[{FFAutomationUtils.ts()}]   Email submission retry {attempt}/3...")
                        await page.wait_for_timeout(1000)
        except RuntimeError:
            raise
        except Exception as e:
            print(f"[{FFAutomationUtils.ts()}]   Warning during email input: {e}")

    async def _handle_password(self, page: Page):
        """Fills the password field and submits."""
        try:
            pw_input = page.locator(
                'input[name="passwd"], input[id="i0118"], input[type="password"]'
            ).first
            if await pw_input.is_visible(timeout=2000):
                for attempt in range(1, 4):
                    try:
                        await pw_input.click(force=True, timeout=3000)
                        await pw_input.fill("", force=True, timeout=3000)
                        await page.wait_for_timeout(200)
                        await pw_input.type(self.config.oracle_password.strip(), delay=40)
                        print(f"[{FFAutomationUtils.ts()}]   Entered password (Attempt {attempt}/3)")
                        await page.wait_for_timeout(300)

                        signin_btn = page.locator('input[type="submit"], #idSIButton9').first
                        await signin_btn.click(force=True, timeout=5000)
                        await page.wait_for_timeout(2500)

                        await self.check_for_sso_errors(page)
                        self.password_filled = True
                        break
                    except RuntimeError:
                        if attempt == 3:
                            raise
                        print(f"[{FFAutomationUtils.ts()}]   Password submission retry {attempt}/3...")
                        await page.wait_for_timeout(1000)
        except RuntimeError:
            raise
        except Exception as e:
            print(f"[{FFAutomationUtils.ts()}]   Warning during password input: {e}")

    async def _handle_rsa_selection(self, page: Page):
        """Clicks the RSAEntraMFA button to trigger approval push notification."""
        try:
            rsa_btn = page.locator('div[role="button"], button, a').filter(
                has_text="RSAEntraMFA"
            ).first
            if await rsa_btn.is_visible(timeout=2000):
                await rsa_btn.click(force=True, timeout=5000)
                self.rsa_clicked = True
                await page.wait_for_timeout(2500)
                print()
                print("  +------------------------------------------------------------+")
                print("  |  RSA push notification sent to phone.                      |")
                print("  |  Please APPROVE the push on your device.                   |")
                print("  |  Waiting up to 5 minutes...                                |")
                print("  +------------------------------------------------------------+")
                print()
        except Exception as e:
            print(f"[{FFAutomationUtils.ts()}]   Warning during RSA selection: {e}")

    async def _handle_kmsi(self, page: Page):
        """Handles 'Stay signed in?' prompt by ticking checkbox and clicking Yes."""
        try:
            kmsi_box = page.locator('#KmsiCheckboxField, input[name="DontShowAgain"]').first
            try:
                if await kmsi_box.count() > 0 and await kmsi_box.is_visible(timeout=1000):
                    if not await kmsi_box.is_checked():
                        await kmsi_box.check()
                        await page.wait_for_timeout(200)
            except Exception:
                pass

            stay_signed_btn = page.locator(
                '#idSIButton9, input[value="Yes"], button:has-text("Yes")'
            ).first
            if await stay_signed_btn.is_visible(timeout=2000):
                await stay_signed_btn.click(force=True, timeout=5000)
                print(f"[{FFAutomationUtils.ts()}]   Confirmed 'Stay signed in? -> Yes'")
                await page.wait_for_timeout(2500)
        except Exception as e:
            print(f"[{FFAutomationUtils.ts()}]   Warning during KMSI confirmation: {e}")

    async def authenticate(self, page: Page, timeout_s: int = 300) -> bool:
        """Coordinates the complete Microsoft SSO authentication lifecycle."""
        try:
            if not self.config.oracle_email or not self.config.oracle_password:
                print(f"[{FFAutomationUtils.ts()}] [WARN] No credentials in .env -- waiting for manual login...")
                return False

            print(f"[{FFAutomationUtils.ts()}] SSO login detected -- completing authentication...")
            deadline = asyncio.get_event_loop().time() + timeout_s

            while asyncio.get_event_loop().time() < deadline:
                await self.check_for_sso_errors(page)
                screen = await self.detect_screen(page)

                if screen == "REDIRECTED":
                    return True

                elif screen == "PICK_ACCOUNT":
                    await self._handle_pick_account(page)

                elif screen == "EMAIL":
                    await self._handle_email(page)

                elif screen == "PASSWORD":
                    await self._handle_password(page)

                elif screen == "RSA_SELECTION":
                    await self._handle_rsa_selection(page)

                elif screen == "RSA_WAIT":
                    try:
                        retry_btn = page.locator('button:has-text("Retry")').first
                        if await retry_btn.count() > 0 and await retry_btn.is_visible(timeout=1000):
                            self.rsa_retry_count += 1
                            print(
                                f"[{FFAutomationUtils.ts()}] [WARN] RSA MFA failed or rejected (Attempt {self.rsa_retry_count}/{self.max_rsa_retries})."
                            )
                            if self.rsa_retry_count >= self.max_rsa_retries:
                                await FFAutomationUtils.save_debug_screenshot(
                                    page, self.config.error_dir, "rsa_failed"
                                )
                                raise RuntimeError("RSA/MFA Authentication failed 3 times.")
                            print(f"[{FFAutomationUtils.ts()}]   Clicking 'Retry' for RSA MFA...")
                            await retry_btn.click()
                            await page.wait_for_timeout(3000)
                            continue
                    except RuntimeError:
                        raise
                    except Exception:
                        pass

                    current_url = page.url.lower()
                    if not any(
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
                        return True

                elif screen == "KMSI":
                    await self._handle_kmsi(page)

                # Check if Content Server home selectors are visible
                for sel in [
                    "div.csui-search a",
                    'a[title="Search in Content Server"]',
                    "div.csui-home-item",
                    ".csui-nodestable",
                ]:
                    try:
                        if await page.locator(sel).first.is_visible(timeout=1000):
                            return True
                    except Exception:
                        pass

                await page.wait_for_timeout(1000)

            return False
        except Exception as e:
            print(f"[{FFAutomationUtils.ts()}] SSO Authentication Exception: {e}")
            raise


# Content Server Downloader & Orchestrator
class FFReportDownloader:
    """Manages searching, navigating folders, downloading F&F documents, and concurrent processing."""

    def __init__(self, config: FFConfig):
        self.config = config
        self.sso_handler = MicrosoftSSOHandler(config)

    async def wait_for_content_server_home(self, page: Page, timeout_s: int = 300):
        """Waits for Content Server home page to load, handling SSO if needed."""
        try:
            print(f"[{FFAutomationUtils.ts()}] Waiting for Content Server home page (up to {timeout_s}s)...")
            home_selectors = [
                "div.csui-search a",
                'a[title="Search in Content Server"]',
                "div.csui-home-item",
                ".csui-nodestable",
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

                for sel in home_selectors:
                    try:
                        if await page.locator(sel).first.is_visible(timeout=1500):
                            print(f"[{FFAutomationUtils.ts()}] Content Server home page ready")
                            await page.wait_for_timeout(2000)
                            return
                    except Exception:
                        pass

                await page.wait_for_timeout(1000)

            await FFAutomationUtils.save_debug_screenshot(page, self.config.error_dir, "home_load_timeout")
            raise RuntimeError(f"Content Server home page did not load within {timeout_s}s. URL: {page.url}")
        except Exception as e:
            if not isinstance(e, RuntimeError):
                await FFAutomationUtils.save_debug_screenshot(page, self.config.error_dir, "home_load_exception")
            raise

    async def navigate_home(self, page: Page):
        """Clicks the Home icon in the navigation bar to return to Content Server home."""
        try:
            await FFAutomationUtils.click_with_fallback(
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
            print(f"[{FFAutomationUtils.ts()}]   Warning navigating home: {e}")

    async def search_employee(self, page: Page, employee_number: str):
        """Executes a search query for the specific employee number."""
        try:
            print(f"[{FFAutomationUtils.ts()}] Searching for employee: {employee_number}")

            # Step 1: Click Search icon in the navbar
            await FFAutomationUtils.click_with_fallback(
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

            # Step 2: Locate search input field
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
            await page.wait_for_timeout(300)

            # Step 3: Type employee number and press Enter
            await search_input.fill("")
            await page.wait_for_timeout(200)
            await search_input.fill(employee_number)
            await page.wait_for_timeout(200)
            await search_input.press("Enter")
            await page.wait_for_timeout(500)

            # Fallback submit button click
            try:
                search_btn = page.locator('a[title="Search in Content Server"], div.csui-search a').first
                if await search_btn.is_visible(timeout=1500):
                    await search_btn.click()
            except Exception:
                pass

            await page.wait_for_timeout(3500)

            # Step 4: Handle modal error dialogs if any
            if await FFAutomationUtils.check_and_dismiss_error_dialog(page):
                print(f"[{FFAutomationUtils.ts()}]   Error dialog dismissed -- retrying search for {employee_number}")
                await self.navigate_home(page)
                await page.wait_for_timeout(2000)
                await self.search_employee(page, employee_number)
                return

            # Step 5: Wait for search results container
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
                if await FFAutomationUtils.check_and_dismiss_error_dialog(page):
                    raise RuntimeError(f"Content Server error during search for employee {employee_number}")

                # Check for "About 0 hits"
                try:
                    zero_hits_loc = page.locator("text=/About 0 hits/i").first
                    if await zero_hits_loc.is_visible(timeout=3000):
                        link_count = await page.locator(f'a:has-text("{employee_number}")').count()
                        if link_count == 0:
                            print(
                                f"[{FFAutomationUtils.ts()}]   Employee {employee_number} not found on Content Server (0 hits) -- skipping"
                            )
                            raise RuntimeError(
                                f"Employee {employee_number} not found on Content Server (0 search results)"
                            )
                except RuntimeError:
                    raise
                except Exception:
                    pass

                await FFAutomationUtils.save_debug_screenshot(page, self.config.error_dir, "no_results", employee_number)
                raise RuntimeError(f"Search results did not appear for employee {employee_number} (timeout)")

        except Exception as e:
            if not isinstance(e, RuntimeError):
                await FFAutomationUtils.save_debug_screenshot(page, self.config.error_dir, "search_error", employee_number)
            raise

    async def click_employee_master_folder(self, page: Page, employee_number: str):
        """Locates and clicks the exact employee master folder in search results."""
        try:
            print(f"[{FFAutomationUtils.ts()}] Looking for exact master folder match: {employee_number}")

            # Polling wait to ensure results are rendered
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

            # Strategy 1: Exact role locator
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

            # Strategy 3: Iterate inner_text
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

            raise RuntimeError(f"Employee folder '{employee_number}' not found in search results -- skipping")

        except Exception as e:
            if not isinstance(e, RuntimeError):
                await FFAutomationUtils.save_debug_screenshot(page, self.config.error_dir, "master_folder_error", employee_number)
            raise

    async def click_adani_exit_documents(self, page: Page, employee_number: Optional[str] = None):
        """Clicks the 'Adani Exit Documents' subfolder link inside the employee folder."""
        try:
            try:
                await page.locator("td.csui-table-cell-name").first.wait_for(
                    state="visible", timeout=self.config.element_timeout
                )
            except PlaywrightTimeoutError:
                await FFAutomationUtils.save_debug_screenshot(page, self.config.error_dir, "no_folder_table", employee_number)
                raise RuntimeError("Employee folder table did not load")

            await page.wait_for_timeout(1500)

            # Try exact role locator first
            try:
                exact_folder = page.get_by_role("link", name="Adani Exit Documents", exact=True).first
                if await exact_folder.is_visible(timeout=5000):
                    await exact_folder.scroll_into_view_if_needed()
                    await exact_folder.click()
                    await page.wait_for_timeout(3000)
                    return
            except Exception:
                pass

            # Fallback selectors
            await FFAutomationUtils.click_with_fallback(
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

        except Exception as e:
            if not isinstance(e, RuntimeError):
                await FFAutomationUtils.save_debug_screenshot(page, self.config.error_dir, "exit_docs_error", employee_number)
            raise

    async def get_document_count(self, page: Page) -> int:
        """Counts the number of document rows in the current folder table."""
        try:
            await page.locator("tbody tr").first.wait_for(state="visible", timeout=self.config.element_timeout)
            await page.wait_for_timeout(1000)
            rows = page.locator("tbody tr td.csui-table-cell-name")
            count = await rows.count()
            print(f"[{FFAutomationUtils.ts()}] Found {count} document(s)")
            return count
        except PlaywrightTimeoutError:
            print(f"[{FFAutomationUtils.ts()}]   No document rows found in the folder")
            return 0
        except Exception as e:
            print(f"[{FFAutomationUtils.ts()}]   Error counting documents: {e}")
            return 0

    async def download_document_at_row(
        self, page: Page, row_index: int, total_rows: int, download_dir: Path, employee_number: str
    ) -> Optional[str]:
        """Downloads a single F&F document from the table by row index."""
        try:
            row_selector = f"tbody tr:nth-of-type({row_index})"
            row = page.locator(row_selector).first

            try:
                await row.scroll_into_view_if_needed(timeout=5000)
            except Exception:
                pass

            await row.wait_for(state="visible", timeout=self.config.element_timeout)

            # Get row name
            row_name = f"Row {row_index}"
            name_cell = row.locator("td.csui-table-cell-name a, td.csui-table-cell-name").first
            if await name_cell.count() > 0:
                row_name = (await name_cell.inner_text()).strip()

            # Check if row is a subfolder
            is_folder = False
            icon_cell = row.locator(
                "td:nth-of-type(2), td.csui-table-cell-icon, td.csui-table-cell-type, td:nth-child(2)"
            ).first
            if await icon_cell.count() > 0:
                html = (await icon_cell.inner_html()).lower()
                if "folder" in html or "csui-icon-node-folder" in html:
                    is_folder = True

            if is_folder:
                print(f"[{FFAutomationUtils.ts()}] Skipping folder: '{row_name}' (row {row_index}/{total_rows})")
                return None

            # Filter for F&F documents only
            if not FFAutomationUtils.is_ff_document(row_name):
                print(f"[{FFAutomationUtils.ts()}] Skipping non-F&F document: '{row_name}' (row {row_index}/{total_rows})")
                return None

            download_dir.mkdir(parents=True, exist_ok=True)
            print(f"[{FFAutomationUtils.ts()}] Downloading F&F document: '{row_name}' (row {row_index}/{total_rows})...")
            await row.hover()
            await page.wait_for_timeout(1000)

            # Click 'More actions' (...)
            more_actions_selectors = [
                f"tbody tr:nth-of-type({row_index}) li.binf-dropdown > a",
                f'xpath=//tbody/tr[{row_index}]//li[contains(@class,"binf-dropdown")]/a',
                f'xpath=//tbody/tr[{row_index}]//a[@title="More actions"]',
            ]
            try:
                await FFAutomationUtils.click_with_fallback(
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
                except Exception as e2:
                    print(f"[{FFAutomationUtils.ts()}]   Could not click 'More actions' for row {row_index}: {e2}")
                    await FFAutomationUtils.save_debug_screenshot(page, self.config.error_dir, "more_actions_fail", employee_number)
                    return None

            await page.wait_for_timeout(1500)

            # Click Download from dropdown menu
            async with page.expect_download(timeout=60_000) as dl_info:
                download_clicked = False
                download_selectors = [
                    'a[title="Download"]',
                    'ul.binf-dropdown-menu a[title="Download"]',
                    'li.binf-dropdown ul.binf-dropdown-menu a[title="Download"]',
                    "li.binf-dropdown li:nth-of-type(1) > a",
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
                                break
                        except Exception:
                            continue

                if not download_clicked:
                    try:
                        await page.evaluate(
                            """() => {
                                const link = document.querySelector('a[title="Download"]');
                                if (link) { link.click(); return true; }
                                return false;
                            }"""
                        )
                        download_clicked = True
                    except Exception:
                        pass

                if not download_clicked:
                    raise RuntimeError(f"Could not click Download for row {row_index} after all attempts")

            download = await dl_info.value
            failure = await download.failure()
            if failure:
                print(f"[{FFAutomationUtils.ts()}]   Download failed for row {row_index}: {failure}")
                return None

            suggested = download.suggested_filename or f"ff_document_{row_index}"
            suggested_clean = suggested.strip()

            # Fix numbering suffix (e.g. '. 2' -> '_2.pdf')
            match = re.search(r"\.\s*(\d+)$", suggested_clean)
            if match:
                suffix_num = match.group(1)
                base = suggested_clean[: match.start()].strip().rstrip(".")
                suggested_clean = f"{base}_{suffix_num}.pdf"
            else:
                ext = Path(suggested_clean).suffix.lower()
                if ext not in [".pdf", ".xlsx", ".xls", ".doc", ".docx", ".png", ".jpg"]:
                    suggested_clean = f"{suggested_clean.rstrip('.')}.pdf"

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
            print(f"[{FFAutomationUtils.ts()}] Saved: {save_path.name} ({size:,} bytes)")
            return str(save_path)

        except PlaywrightTimeoutError:
            print(f"[{FFAutomationUtils.ts()}]   Download timed out for row {row_index}")
            await FFAutomationUtils.save_debug_screenshot(page, self.config.error_dir, "download_timeout", employee_number)
            return None
        except Exception as e:
            print(f"[{FFAutomationUtils.ts()}]   Error downloading row {row_index}: {e}")
            return None

    async def download_ff_for_employee(
        self, page: Page, employee_number: str, download_dir: Path
    ) -> Tuple[str, List[str]]:
        """Coordinates full download workflow for a single employee."""
        downloaded_files: List[str] = []
        status = FFConfig.STATUS_ERROR

        try:
            await self.search_employee(page, employee_number)
            await self.click_employee_master_folder(page, employee_number)

            try:
                await self.click_adani_exit_documents(page, employee_number)
            except RuntimeError as e:
                if "Employee folder table did not load" in str(e) or "Adani Exit Documents" in str(e):
                    print(
                        f"[{FFAutomationUtils.ts()}] [WARN] Folder table issue for {employee_number}. Retrying once from home..."
                    )
                    await self.navigate_home(page)
                    await page.wait_for_timeout(3000)
                    await self.search_employee(page, employee_number)
                    await self.click_employee_master_folder(page, employee_number)
                    await self.click_adani_exit_documents(page, employee_number)
                else:
                    raise

            doc_count = await self.get_document_count(page)
            if doc_count == 0:
                print(f"[{FFAutomationUtils.ts()}]   No documents found for employee {employee_number}")
                status = FFConfig.STATUS_NO_FF_DOCS
            else:
                emp_dir = download_dir / employee_number
                for row_idx in range(1, doc_count + 1):
                    await FFAutomationUtils.check_and_dismiss_error_dialog(page)
                    result = await self.download_document_at_row(
                        page, row_idx, doc_count, emp_dir, employee_number
                    )
                    if result:
                        downloaded_files.append(result)
                    await page.wait_for_timeout(2000)

                if downloaded_files:
                    status = FFConfig.STATUS_DOWNLOADED
                else:
                    status = FFConfig.STATUS_NO_FF_DOCS

            FFAutomationUtils.cleanup_error_screenshots(self.config.error_dir, employee_number)

        except RuntimeError as e:
            err_msg = str(e)
            if "0 search results" in err_msg or "0 hits" in err_msg:
                print(f"[{FFAutomationUtils.ts()}] Employee {employee_number} not found on Content Server (0 hits)")
                status = FFConfig.STATUS_NOT_FOUND
                FFAutomationUtils.cleanup_error_screenshots(self.config.error_dir, employee_number)
            else:
                print(f"[{FFAutomationUtils.ts()}] Error processing employee {employee_number}: {e}")
                await FFAutomationUtils.save_debug_screenshot(page, self.config.error_dir, "error", employee_number)
                status = FFConfig.STATUS_ERROR
        except Exception as e:
            print(f"[{FFAutomationUtils.ts()}] Error processing employee {employee_number}: {e}")
            await FFAutomationUtils.save_debug_screenshot(page, self.config.error_dir, "error", employee_number)
            status = FFConfig.STATUS_ERROR

        return status, downloaded_files

    async def _run_batch(
        self,
        p_context,
        emp_list: List[str],
        num_tabs: int,
        nav_lock: asyncio.Lock,
        total_label: int,
        all_results: Dict[str, Tuple[str, List[str]]],
        start_offset: int = 0,
        upload_cb=None,
    ):
        """Processes a batch of employees across concurrent browser tabs."""
        first_page = p_context.pages[0] if p_context.pages else await p_context.new_page()
        worker_pages = [first_page]
        for i in range(1, num_tabs):
            try:
                new_tab = await p_context.new_page()
                worker_pages.append(new_tab)
                print(f"[{FFAutomationUtils.ts()}]   Opened tab {i + 1}/{num_tabs}")
                await asyncio.sleep(0.5)
            except Exception as e:
                print(f"[{FFAutomationUtils.ts()}]   Could not open tab {i + 1}: {e}")
                break

        actual_tabs = len(worker_pages)
        print(f"[{FFAutomationUtils.ts()}] {actual_tabs} tab(s) ready. Starting downloads...\n")

        async def _worker(worker_page, emp_num: str, idx: int):
            try:
                print(f"[{FFAutomationUtils.ts()}] -- [Tab] Employee {idx}/{total_label}: {emp_num} --")
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
                            print(f"[{FFAutomationUtils.ts()}]   [Tab] Navigation timeout for {emp_num} -- retrying once...")
                            await worker_page.wait_for_timeout(3000)
                        else:
                            raise RuntimeError(
                                f"Content Server navigation failed after 2 attempts for {emp_num}: {nav_err}"
                            )

                if not nav_success:
                    raise RuntimeError(f"Could not navigate to Content Server for {emp_num}")

                status, downloaded = await self.download_ff_for_employee(
                    worker_page, emp_num, self.config.download_dir
                )
                all_results[emp_num] = (status, downloaded)

                if downloaded and upload_cb:
                    try:
                        print(f"[{FFAutomationUtils.ts()}] [Tab] Triggering upload for {emp_num}...")
                        await upload_cb(emp_num, downloaded)
                    except Exception as e:
                        print(f"[{FFAutomationUtils.ts()}] [Tab] Upload callback failed for {emp_num}: {e}")

            except Exception as e:
                print(f"[{FFAutomationUtils.ts()}] [Tab] Error for {emp_num}: {e}")
                await FFAutomationUtils.save_debug_screenshot(worker_page, self.config.error_dir, "error", emp_num)
                all_results[emp_num] = (FFConfig.STATUS_ERROR, [])

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

    def _print_categorized_summary(self, all_results: Dict[str, Tuple[str, List[str]]]):
        """Prints a structured summary of download results."""
        downloaded = {k: v for k, (s, v) in all_results.items() if s == FFConfig.STATUS_DOWNLOADED}
        no_ff = {k: v for k, (s, v) in all_results.items() if s == FFConfig.STATUS_NO_FF_DOCS}
        not_found = {k: v for k, (s, v) in all_results.items() if s == FFConfig.STATUS_NOT_FOUND}
        errors = {k: v for k, (s, v) in all_results.items() if s == FFConfig.STATUS_ERROR}

        total_files = sum(len(v) for v in downloaded.values())
        total_emps = len(all_results)

        print()
        print(f"[{FFAutomationUtils.ts()}] ===========================================================")
        print(f"[{FFAutomationUtils.ts()}] DOWNLOAD SUMMARY")
        print(f"[{FFAutomationUtils.ts()}] ===========================================================")

        if downloaded:
            print(f"[{FFAutomationUtils.ts()}]")
            print(f"[{FFAutomationUtils.ts()}]   [OK] DOWNLOADED ({len(downloaded)} employee(s), {total_files} file(s)):")
            for emp, files in downloaded.items():
                print(f"[{FFAutomationUtils.ts()}]      {emp}: {len(files)} file(s)")

        if no_ff:
            print(f"[{FFAutomationUtils.ts()}]")
            print(
                f"[{FFAutomationUtils.ts()}]   [INFO] NO F&F DOCUMENT ({len(no_ff)} employee(s) -- folder found, no F&F sheet uploaded):"
            )
            for emp in no_ff:
                print(f"[{FFAutomationUtils.ts()}]      {emp}")

        if not_found:
            print(f"[{FFAutomationUtils.ts()}]")
            print(
                f"[{FFAutomationUtils.ts()}]   [--] NOT FOUND ON CONTENT SERVER ({len(not_found)} employee(s) -- 0 search hits):"
            )
            for emp in not_found:
                print(f"[{FFAutomationUtils.ts()}]      {emp}")

        if errors:
            print(f"[{FFAutomationUtils.ts()}]")
            print(
                f"[{FFAutomationUtils.ts()}]   [FAIL] FAILED / ERRORS ({len(errors)} employee(s) -- search timeout or network failure):"
            )
            for emp in errors:
                print(f"[{FFAutomationUtils.ts()}]      {emp}")

        print(f"[{FFAutomationUtils.ts()}] -----------------------------------------------------------")
        print(f"[{FFAutomationUtils.ts()}]   Total: {total_files} file(s) downloaded for {total_emps} employee(s)")
        if errors:
            print(
                f"[{FFAutomationUtils.ts()}]   [WARN] {len(errors)} employee(s) still have unresolved errors -- check errors/ff_error/ for screenshots"
            )
        print(f"[{FFAutomationUtils.ts()}] ===========================================================")

    async def execute(self, employee_numbers: List[str], upload_callback=None):
        """Executes the full F&F download batch and retry cycle."""
        if not employee_numbers:
            print(f"[{FFAutomationUtils.ts()}] No employee numbers provided. Exiting.")
            return

        if not self.config.content_server_url:
            print(f"[{FFAutomationUtils.ts()}] [FATAL ERROR] CONTENT_SERVER_URL is completely blank in your .env file!")
            return

        concurrency = min(self.config.max_concurrent_tabs, len(employee_numbers))
        print(
            f"[{FFAutomationUtils.ts()}] Starting F&F report download ({'headless' if self.config.headless else 'visible'})"
        )
        print(f"[{FFAutomationUtils.ts()}] Concurrency: {concurrency} tab(s) for {len(employee_numbers)} employee(s)")

        FFAutomationUtils.cleanup_profile_locks(self.config.automation_profile_dir)
        all_results: Dict[str, Tuple[str, List[str]]] = {}

        context = None
        try:
            async with async_playwright() as p:
                print(f"[{FFAutomationUtils.ts()}] Launching browser...")
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

                print(f"[{FFAutomationUtils.ts()}] Navigating to Content Server...")
                await first_page.goto(
                    self.config.content_server_url,
                    wait_until="domcontentloaded",
                    timeout=self.config.nav_timeout,
                )
                await first_page.wait_for_timeout(3000)

                await self.wait_for_content_server_home(first_page, timeout_s=120)
                await first_page.wait_for_timeout(2000)
                print(f"[{FFAutomationUtils.ts()}] Login confirmed. Preparing {concurrency} tab(s)...\n")

                nav_lock = asyncio.Lock()

                # Primary batch execution
                await self._run_batch(
                    context,
                    employee_numbers,
                    concurrency,
                    nav_lock,
                    len(employee_numbers),
                    all_results,
                    upload_cb=upload_callback,
                )

                # Retry loop
                initial_not_found = [
                    emp for emp, (st, _) in all_results.items() if st == FFConfig.STATUS_NOT_FOUND
                ]

                for retry_round in range(1, self.config.max_retry_rounds + 1):
                    folder_failed_ids = FFAutomationUtils.get_failed_ids_from_error_folder(self.config.error_dir)

                    if retry_round == 1 and initial_not_found:
                        combined_ids = list(folder_failed_ids)
                        for nf_id in initial_not_found:
                            if nf_id not in combined_ids:
                                combined_ids.append(nf_id)
                        current_retry_ids = combined_ids
                    else:
                        current_retry_ids = folder_failed_ids

                    if not current_retry_ids:
                        print(f"[{FFAutomationUtils.ts()}] errors/ff_error/ is empty -- no retries needed.")
                        break

                    retry_concurrency = min(2, len(current_retry_ids))
                    print()
                    print(f"[{FFAutomationUtils.ts()}] ===========================================================")
                    print(
                        f"[{FFAutomationUtils.ts()}] RETRY ROUND {retry_round}/{self.config.max_retry_rounds}: {len(current_retry_ids)} employee(s) queued for retry"
                    )
                    print(f"[{FFAutomationUtils.ts()}]   IDs: {', '.join(current_retry_ids)}")
                    print(f"[{FFAutomationUtils.ts()}]   Retrying with {retry_concurrency} tab(s) for reliability...")
                    print(f"[{FFAutomationUtils.ts()}] ===========================================================")
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
                        retry_status, _ = all_results.get(emp_num, (FFConfig.STATUS_ERROR, []))
                        if retry_status != FFConfig.STATUS_ERROR:
                            print(f"[{FFAutomationUtils.ts()}]   [OK] {emp_num} resolved on retry (status: {retry_status})")
                            FFAutomationUtils.cleanup_error_screenshots(self.config.error_dir, emp_num)
                            resolved_count += 1

                    print(
                        f"[{FFAutomationUtils.ts()}]   Round {retry_round} complete: {resolved_count}/{len(current_retry_ids)} resolved"
                    )

                    remaining = FFAutomationUtils.get_failed_ids_from_error_folder(self.config.error_dir)
                    if not remaining:
                        print(f"[{FFAutomationUtils.ts()}] errors/ff_error/ is now empty -- all retries succeeded!")
                        break

        except Exception as e:
            print(f"[{FFAutomationUtils.ts()}] FATAL: {e}\n{traceback.format_exc()}")
        finally:
            if context:
                for pg in context.pages:
                    try:
                        await pg.close()
                    except Exception:
                        pass
                try:
                    await context.close()
                    print(f"[{FFAutomationUtils.ts()}] Browser context closed cleanly.")
                except Exception as e:
                    print(f"[{FFAutomationUtils.ts()}] Warning closing browser context: {e}")

        self._print_categorized_summary(all_results)


# Terminal Interactive Prompt Handler
class TerminalPromptHandler:
    """Manages CLI interactive prompt for collecting employee numbers."""

    @staticmethod
    def prompt_employee_numbers() -> List[str]:
        """Interactive terminal prompt collecting employee IDs from standard input."""
        try:
            terminal = sys.__stdout__
            terminal.write("\n")
            terminal.write("  =======================================================\n")
            terminal.write("          F&F Exit Document Downloader                   \n")
            terminal.write("  =======================================================\n")
            terminal.write("\n")
            terminal.flush()

            terminal.write("  Enter Employee ID (or multiple comma-separated): ")
            terminal.flush()
            raw = sys.__stdin__.readline()

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
        except (EOFError, KeyboardInterrupt):
            return []
        except Exception as e:
            print(f"[{FFAutomationUtils.ts()}] Prompt Error: {e}")
            return []


# Main Entrypoints and Helpers
def prompt_employee_numbers() -> List[str]:
    """Top-level wrapper for terminal prompt collection."""
    return TerminalPromptHandler.prompt_employee_numbers()


async def download_ff_reports(employee_numbers: List[str], upload_callback=None):
    """Entrypoint function for downloading F&F reports for the specified employees."""
    config = FFConfig()
    downloader = FFReportDownloader(config)
    await downloader.execute(employee_numbers, upload_callback=upload_callback)


if __name__ == "__main__":
    numbers = TerminalPromptHandler.prompt_employee_numbers()
    if numbers:
        asyncio.run(download_ff_reports(numbers))
    else:
        print("No employee numbers entered. Exiting.")
