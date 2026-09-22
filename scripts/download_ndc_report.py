# Oracle Fusion - NDC Process Request Status Report Downloader

import asyncio
import os
import subprocess
import sys
import traceback
from datetime import datetime
from pathlib import Path
from typing import List, Optional

import pandas as pd
from dotenv import load_dotenv
from playwright.async_api import Frame, Page, TimeoutError as PlaywrightTimeoutError, async_playwright

# Suppress Chrome subprocess window on Windows
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
LOG_FILE = LOG_DIR / "download_ndc_report.log"

sys.stdout = LoggerWriter(LOG_FILE)
sys.stderr = LoggerWriter(LOG_FILE)

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

            # Output and storage directories
            self.download_dir: Path = Path(__file__).parent.parent / "uploads" / "NDC_Reports"
            self.automation_profile_dir: Path = Path(__file__).parent.parent / "chrome_automation_profile"
            self.error_dir: Path = Path(__file__).parent.parent / "errors" / "ndc_error"

            # Browser configuration
            self.headless: bool = self._parse_bool(os.getenv("HEADLESS"), default=True)
            self.start_maximized: bool = self._parse_bool(os.getenv("START_MAXIMIZED"), default=True)
            self.window_position: str = (os.getenv("WINDOW_POSITION") or "").strip()
            self.browser_channel: Optional[str] = (os.getenv("BROWSER_CHANNEL") or "").strip() or None

            # Report parameters
            self.business_unit_search: str = "Adani Green"
            self.from_date_value: str = "02-09-2026"

            # Ensure required directories exist
            self.download_dir.mkdir(parents=True, exist_ok=True)
            self.error_dir.mkdir(parents=True, exist_ok=True)
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
            path = error_dir / f"debug_{label}_{datetime.now().strftime('%H%M%S')}.png"
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


# Microsoft SSO Handler
class MicrosoftSSOHandler:
    """Manages Microsoft Entra SSO, MFA push verification, and KMSI prompt."""

    def __init__(self, config: AppConfig):
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

            # Check if redirected away from login screens
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
                            print(f"[{AutomationUtils.ts()}]   Selecting account: {self.config.oracle_email}")
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
                        print(f"[{AutomationUtils.ts()}]   Selected account: {self.config.oracle_email}")
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
            print(f"[{AutomationUtils.ts()}]   Warning during account pick: {e}")

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
                        print(f"[{AutomationUtils.ts()}]   Entered email ({self.config.oracle_email})")
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
                        print(f"[{AutomationUtils.ts()}]   Email submission retry {attempt}/3...")
                        await page.wait_for_timeout(1000)
        except RuntimeError:
            raise
        except Exception as e:
            print(f"[{AutomationUtils.ts()}]   Warning during email input: {e}")

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
                        print(f"[{AutomationUtils.ts()}]   Entered password (Attempt {attempt}/3)")
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
                        print(f"[{AutomationUtils.ts()}]   Password submission retry {attempt}/3...")
                        await page.wait_for_timeout(1000)
        except RuntimeError:
            raise
        except Exception as e:
            print(f"[{AutomationUtils.ts()}]   Warning during password input: {e}")

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
            print(f"[{AutomationUtils.ts()}]   Warning during RSA selection: {e}")

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
                print(f"[{AutomationUtils.ts()}]   Confirmed 'Stay signed in? -> Yes'")
                await page.wait_for_timeout(2500)
        except Exception as e:
            print(f"[{AutomationUtils.ts()}]   Warning during KMSI confirmation: {e}")

    async def authenticate(self, page: Page, timeout_s: int = 300) -> bool:
        """Coordinates the complete Microsoft SSO authentication lifecycle."""
        try:
            if not self.config.oracle_email or not self.config.oracle_password:
                print(f"[{AutomationUtils.ts()}] ⚠ No credentials in .env — waiting for manual login...")
                return False

            print(f"[{AutomationUtils.ts()}] SSO login detected — completing authentication...")
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
                                f"[{AutomationUtils.ts()}] ⚠ RSA MFA failed or rejected (Attempt {self.rsa_retry_count}/{self.max_rsa_retries})."
                            )
                            if self.rsa_retry_count >= self.max_rsa_retries:
                                await AutomationUtils.save_debug_screenshot(
                                    page, self.config.error_dir, "rsa_failed_3_times"
                                )
                                raise RuntimeError("RSA/MFA Authentication failed 3 times.")
                            print(f"[{AutomationUtils.ts()}]   Clicking 'Retry' for RSA MFA...")
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

                # Check if Oracle Home is already visible
                for sel in [
                    'text="HR SPOC Reports"',
                    'a:has-text("HR SPOC Reports")',
                    ".fuse-welcome-page",
                    ".springboard",
                ]:
                    try:
                        if await page.locator(sel).first.is_visible(timeout=1000):
                            return True
                    except Exception:
                        pass

                await page.wait_for_timeout(1000)

            return False
        except Exception as e:
            print(f"[{AutomationUtils.ts()}] SSO Authentication Exception: {e}")
            raise


# Report Post-Filter
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
                        final_df.to_excel(file_path, header=False, index=False)
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


# Oracle Report Downloader & Orchestrator
class NDCOracleReportDownloader:
    """Coordinates browser navigation, parameter selection, report download, and file processing."""

    def __init__(self, app_config: AppConfig):
        self.config = app_config
        self.sso_handler = MicrosoftSSOHandler(app_config)

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
        except Exception as e:
            raise

    async def get_bipublisher_frame(self, report_page: Page, timeout_s: int = 90) -> Frame:
        """Locates the parameter iframe inside the BI Publisher report page."""
        try:
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
        except Exception as e:
            raise

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


# Main entry point for the script
async def download_ndc_report():
    """Main entrypoint function for NDC report downloader."""
    app_config = AppConfig()
    downloader = NDCOracleReportDownloader(app_config)
    await downloader.download()


if __name__ == "__main__":
    asyncio.run(download_ndc_report())
