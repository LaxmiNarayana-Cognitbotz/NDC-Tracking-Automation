"""
OpenText Content Server - F&F (Full & Final) Exit Document Downloader
"""

import asyncio
import glob
import os
import re
import sys
import traceback
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from dotenv import load_dotenv
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError

# ── Logging Setup ────────────────────────────────────────────────────────────
LOG_DIR = Path(__file__).parent.parent / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "download_ff_report.log"


class LoggerWriter:
    def __init__(self, file_path):
        self.terminal = sys.stdout
        self.file_path = file_path

    def write(self, message):
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
                    self.terminal.write(message.encode("cp1252", errors="replace").decode("cp1252"))
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

# ── Configuration ────────────────────────────────────────────────────────────

# SSO credentials (same as NDC script)
ORACLE_EMAIL = os.getenv("ORACLE_EMAIL")
ORACLE_PASSWORD = os.getenv("ORACLE_PASSWORD")

# Content Server URL
CONTENT_SERVER_URL = os.getenv("CONTENT_SERVER_URL")

# Output directory for downloaded F&F documents
DOWNLOAD_DIR = Path(__file__).parent.parent / "uploads" / "FF_Reports"

# Standalone automation profile (persists browser session between runs)
AUTOMATION_PROFILE_DIR = Path(__file__).parent.parent / "chrome_automation_profile"

# Headless mode
HEADLESS = os.getenv("HEADLESS", "true").lower()

# Default timeout for individual element interactions (ms)
ELEMENT_TIMEOUT = 15_000

# Page load / navigation timeout (ms)
NAV_TIMEOUT = 90_000

# Maximum number of concurrent browser tabs for parallel downloads
MAX_CONCURRENT_TABS = 5

# How long to wait for the search results page to fully render (ms)
SEARCH_RESULTS_WAIT = 30_000

# ── Result Status Constants ──────────────────────────────────────────────────
# Used to distinguish different outcomes per employee so the summary is clear.
STATUS_DOWNLOADED = "DOWNLOADED"          # F&F sheet found and saved
STATUS_NO_FF_DOCS = "NO_FF_DOCS"          # Folder found, but no F&F sheet uploaded
STATUS_NOT_FOUND  = "NOT_FOUND_0_HITS"    # Content Server returned 0 search results
STATUS_ERROR      = "ERROR"               # Search timeout, network failure, etc.

# Error screenshot directory
ERROR_SCREENSHOT_DIR = Path(__file__).parent.parent / "errors" / "ff_error"


def ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _cleanup_profile_locks():
    """Remove Chrome lock files on the automation profile only."""
    lock_files = ["SingletonLock", "SingletonSocket", "SingletonCookie", "lockfile"]
    if not AUTOMATION_PROFILE_DIR.exists():
        return
    for lock_name in lock_files:
        lock_path = AUTOMATION_PROFILE_DIR / lock_name
        try:
            if lock_path.exists():
                lock_path.unlink()
        except OSError:
            pass


def cleanup_error_screenshots(employee_number: str):
    """Delete all debug error screenshots for the given employee from errors/ff_error/."""
    if not ERROR_SCREENSHOT_DIR.exists():
        return
    patterns = [
        str(ERROR_SCREENSHOT_DIR / f"debug_*_{employee_number}.png"),
        str(ERROR_SCREENSHOT_DIR / f"debug_error_{employee_number}_*.png"),
    ]
    for pattern in patterns:
        for path in glob.glob(pattern):
            try:
                os.remove(path)
                print(f"[{ts()}]   Cleaned up error screenshot: {Path(path).name}")
            except OSError:
                pass


def _extract_employee_id_from_debug_screenshot(path: Path) -> Optional[str]:
    """Extract employee ID from debug_{label}_{employee_id}.png screenshots."""
    match = re.match(r"^debug_[a-z_]+_(\d+)$", path.stem, flags=re.IGNORECASE)
    if match:
        return match.group(1)

    # Backward compatibility for older timestamped error screenshots.
    match = re.match(r"^debug_error_(\d+)_\d+$", path.stem, flags=re.IGNORECASE)
    if match:
        return match.group(1)

    return None


def _debug_screenshot_path(label: str, employee_number: Optional[str] = None) -> Path:
    safe_label = re.sub(r"[^A-Za-z_]+", "_", label).strip("_").lower()
    if employee_number:
        return ERROR_SCREENSHOT_DIR / f"debug_{safe_label}_{employee_number}.png"
    return ERROR_SCREENSHOT_DIR / f"debug_{safe_label}.png"


async def save_debug_screenshot(page, label: str, employee_number: Optional[str] = None) -> str:
    """Save a full-page screenshot for debugging."""
    ERROR_SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
    path = _debug_screenshot_path(label, employee_number)
    try:
        await page.screenshot(path=str(path), full_page=True)
        print(f"[{ts()}]   Screenshot: {path}")
    except Exception as e:
        print(f"[{ts()}]   Could not save screenshot: {e}")
    return str(path)


def get_failed_ids_from_error_folder() -> List[str]:
    """
    Scan errors/ff_error/ for debug screenshots and extract unique
    employee IDs from the filenames.

    Filename pattern: debug_{label}_{employee_id}.png
    Example: debug_error_30087487.png  ->  "30087487"
    """
    if not ERROR_SCREENSHOT_DIR.exists():
        return []
    seen = set()
    ids = []
    for path in ERROR_SCREENSHOT_DIR.glob("debug_*.png"):
        emp_id = _extract_employee_id_from_debug_screenshot(path)
        if emp_id and emp_id not in seen:
            seen.add(emp_id)
            ids.append(emp_id)
    return ids


async def check_and_dismiss_error_dialog(page) -> bool:
    """
    Detect and dismiss Content Server error dialogs (e.g. 'Network communication failed.').
    Returns True if an error dialog was found and dismissed.
    """
    try:
        # Look for common Content Server error modals
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
                    print(f"[{ts()}] [WARN] Content Server error dialog detected: {error_text}")

                    # Try to click the Close button
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
                                print(f"[{ts()}]   Dismissed error dialog")
                                await page.wait_for_timeout(1500)
                                return True
                        except Exception:
                            continue

                    # If no Close button, press Escape
                    await page.keyboard.press("Escape")
                    print(f"[{ts()}]   Dismissed error dialog via Escape key")
                    await page.wait_for_timeout(1500)
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


# ── SSO Handling (ported from download_ndc_report.py) ────────────────────────

async def check_for_sso_errors(page):
    """Check if Microsoft SSO shows a validation error, and raise an exception if found."""
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


async def _handle_microsoft_sso(page):
    """
    Automatically fill the Microsoft SSO login form dynamically.
    """
    if not ORACLE_EMAIL or not ORACLE_PASSWORD:
        print(f"[{ts()}] [WARN] No credentials in .env -- waiting for manual login...")
        print(f"[{ts()}]   Create a .env file with ORACLE_EMAIL and ORACLE_PASSWORD")
        return False

    print(f"[{ts()}] SSO login detected -- filling credentials automatically...")

    # ── Dynamic SSO State Machine ─────────────────────────────────────────
    # Since Microsoft login screens can appear in unpredictable order (skipping
    # email, asking for account pick first, or direct password), we dynamically
    # detect which form element is currently active and handle it.
    email_filled = False
    password_filled = False
    rsa_clicked = False

    for iteration in range(15):
        # 0. Check if we have been redirected away from Microsoft SSO
        current_url = page.url.lower()
        if not any(x in current_url for x in ["login.microsoftonline", "login.microsoft", "adfs", "saml"]):
            print(f"[{ts()}]   SSO login completed -- redirected away from login screen.")
            return True

        # 1. "Pick an account" screen (check this first before touching inputs)
        if not email_filled and not password_filled:
            try:
                pick_header = page.locator('div#loginHeader:has-text("Pick an account"), div[role="heading"]:has-text("Pick an account")').first
                if await pick_header.is_visible(timeout=1000):
                    print(f"[{ts()}]   'Pick an account' screen detected.")
                    target_email = ORACLE_EMAIL.strip().lower()
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
                                    print(f"[{ts()}]   Found matching account tile: {ORACLE_EMAIL}")
                                    await tile.click(force=True, timeout=5000)
                                    account_found = True
                                    break
                            if account_found:
                                break
                        except Exception:
                            continue

                    if not account_found:
                        try:
                            el = page.locator(f'small:has-text("{ORACLE_EMAIL}"), div:has-text("{ORACLE_EMAIL}")').first
                            if await el.is_visible(timeout=2000):
                                parent = el.locator('xpath=ancestor::div[@role="button" or contains(@class,"table") or contains(@class,"tile")]').first
                                if await parent.count() > 0:
                                    await parent.click(force=True, timeout=5000)
                                else:
                                    await el.click(force=True, timeout=5000)
                                account_found = True
                                print(f"[{ts()}]   Clicked account tile via email text match: {ORACLE_EMAIL}")
                        except Exception:
                            pass

                    if not account_found:
                        try:
                            use_another = page.locator('div#otherTile, div:has-text("Use another account")').last
                            if await use_another.is_visible(timeout=3000):
                                await use_another.click(force=True, timeout=5000)
                                print(f"[{ts()}]   Account not in list -- clicked 'Use another account'")
                        except Exception:
                            print(f"[{ts()}]   Could not find matching account or 'Use another account' button")

                    await page.wait_for_timeout(3000)
                    continue
            except Exception:
                pass

        # 2. Password input screen
        # To avoid false positives on the email screen (where password box exists in DOM),
        # we require EITHER that email was already filled in step 1, OR that explicit password
        # screen indicators are visible ("Enter password" heading, "Sign in" button, or back-arrow banner).
        if not password_filled:
            try:
                pw_input = page.locator('input[name="passwd"], input[id="i0118"], input[type="password"]').first
                pw_indicator = page.locator('input[type="submit"][value="Sign in"], #idSIButton9[value="Sign in"], #idSIButton9:has-text("Sign in"), div:has-text("Enter password"), div:has-text("enter password"), #identityBanner, #idA_PWD_SwitchToCredPicker').first
                if await pw_input.is_visible(timeout=1000) and (email_filled or await pw_indicator.is_visible(timeout=1000)):
                    print(f"[{ts()}]   Password screen detected.")
                    for attempt in range(1, 4):
                        await pw_input.click(force=True, timeout=3000)
                        await pw_input.fill("", force=True, timeout=3000)
                        await page.wait_for_timeout(200)
                        await pw_input.type(ORACLE_PASSWORD.strip(), delay=50)
                        print(f"[{ts()}]   Entered password (Attempt {attempt}/3)")
                        await page.wait_for_timeout(500)

                        signin_btn = page.locator('input[type="submit"], #idSIButton9').first
                        await signin_btn.click(force=True, timeout=5000)
                        print(f"[{ts()}]   Clicked Sign In")
                        await page.wait_for_timeout(3000)

                        try:
                            await check_for_sso_errors(page)
                            password_filled = True
                            email_filled = True  # Email step wasn't needed
                            break
                        except RuntimeError as e:
                            if attempt == 3:
                                raise
                            print(f"[{ts()}]   {e} - Retrying...")
                    continue
            except RuntimeError:
                raise
            except Exception:
                pass

        # 3. Email input screen (only if password not filled and email not filled)
        if not email_filled and not password_filled:
            try:
                email_input = page.locator('input[name="loginfmt"], input[type="email"], input[id="i0116"]').first
                if await email_input.is_visible(timeout=1000):
                    print(f"[{ts()}]   Email screen detected.")
                    for attempt in range(1, 4):
                        await email_input.click(force=True, timeout=3000)
                        await email_input.fill("", force=True, timeout=3000)
                        await page.wait_for_timeout(200)
                        await email_input.type(ORACLE_EMAIL.strip(), delay=50)
                        print(f"[{ts()}]   Entered email (Attempt {attempt}/3)")
                        await page.wait_for_timeout(500)

                        next_btn = page.locator('input[type="submit"], #idSIButton9').first
                        await next_btn.click(force=True, timeout=5000)
                        print(f"[{ts()}]   Clicked Next")
                        await page.wait_for_timeout(3000)

                        try:
                            await check_for_sso_errors(page)
                            email_filled = True
                            break
                        except RuntimeError as e:
                            if attempt == 3:
                                raise
                            print(f"[{ts()}]   {e} - Retrying...")
                    continue
            except RuntimeError:
                raise
            except Exception:
                pass

        # 4. RSA / MFA "Verify your identity" screen
        if not rsa_clicked:
            try:
                rsa_btn = page.locator('div[role="button"], button').filter(has_text="RSAEntraMFA").first
                if await rsa_btn.is_visible(timeout=1500):
                    print(f"[{ts()}]   'Verify your identity' screen detected.")
                    await rsa_btn.click(force=True, timeout=5000)
                    print(f"[{ts()}]   Clicked 'Approve with RSAEntraMFA'")
                    await page.wait_for_timeout(3000)
                    rsa_clicked = True
                    break  # Break loop and wait for push approval
            except Exception:
                pass

        # If password was already filled and we are just waiting for MFA/redirect or KMSI
        if password_filled:
            break

        await page.wait_for_timeout(1000)

    # ── Step 4: RSA / MFA approval ────────────────────────────────────────
    print(f"[{ts()}] Waiting for RSA/MFA approval...")
    print()
    print("  +------------------------------------------------------------+")
    print("  |  RSA push notification sent to the user's phone.           |")
    print("  |  Please ask them to APPROVE it.                            |")
    print("  |  Waiting up to 5 minutes...                                |")
    print("  +------------------------------------------------------------+")
    print()

    return True


# ── Helpers ──────────────────────────────────────────────────────────────────

async def safe_click(locator, description: str, timeout: int = ELEMENT_TIMEOUT):
    """Wait for an element to be visible, scroll into view, then click it."""
    try:
        await locator.wait_for(state="visible", timeout=timeout)
        await locator.scroll_into_view_if_needed()
        await locator.click(timeout=timeout)
    except PlaywrightTimeoutError:
        raise RuntimeError(f"Could not click '{description}' within {timeout}ms")


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


async def wait_for_content_server_home(page, timeout_s: int = 300):
    """
    Wait for Content Server home page. Handles Microsoft SSO login if redirected,
    RSA/MFA retry (up to 3 attempts), and 'Stay signed in?' prompt.
    """
    print(f"[{ts()}] Waiting for Content Server home page (up to {timeout_s}s)...")
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

        # ── Check if we're on a login/SSO page ────────────────────────────
        current_url = page.url.lower()
        if any(x in current_url for x in ["login.microsoftonline", "login.microsoft", "adfs", "saml"]):
            if not login_handled:
                login_handled = await _handle_microsoft_sso(page)

        # ── Handle RSA / SecurID Rejection ────────────────────────────────
        try:
            if "securid.com" in current_url or "adani.auth" in current_url:
                retry_btn = page.locator('button:has-text("Retry")').first
                if await retry_btn.is_visible(timeout=1000):
                    rsa_retry_count += 1
                    print(f"[{ts()}] [WARN] RSA/MFA Authentication failed or rejected (Attempt {rsa_retry_count}/3).")
                    if rsa_retry_count >= 3:
                        await save_debug_screenshot(page, "rsa_failed")
                        raise RuntimeError("RSA/MFA Authentication failed 3 times.")

                    print(f"[{ts()}]   Clicking 'Retry' for RSA...")
                    await retry_btn.click()
                    await page.wait_for_timeout(3000)
                    print(f"[{ts()}] Waiting for RSA/MFA approval...")
                    continue
        except RuntimeError:
            raise
        except Exception:
            pass

        # ── Handle "Stay signed in?" prompt ───────────────────────────────
        try:
            stay_signed = page.locator('#idSIButton9, input[value="Yes"]').first
            if await stay_signed.is_visible(timeout=1000):
                await stay_signed.click()
                print(f"[{ts()}]   Clicked 'Stay signed in? -> Yes'")
                await page.wait_for_timeout(3000)
                continue
        except Exception:
            pass

        # ── Check for Content Server home page ────────────────────────────
        for sel in home_selectors:
            try:
                if await page.locator(sel).first.is_visible(timeout=2000):
                    print(f"[{ts()}] Content Server home page ready")
                    return
            except Exception:
                pass
        await page.wait_for_timeout(2000)

    await save_debug_screenshot(page, "home_load_timeout")
    raise RuntimeError(
        f"Content Server home page did not load within {timeout_s}s.\n"
        f"Current URL: {page.url}"
    )


async def search_employee(page, employee_number: str):
    """
    Perform a search in Content Server for the given employee number.
    Flow: Click search icon -> type in search box -> click search button.
    """
    print(f"[{ts()}] Searching for employee: {employee_number}")

    # Step 1: Click the Search icon in the navbar
    await click_with_fallback(
        page,
        [
            'a[title="Search in Content Server"]',
            "div.csui-search a",
            'xpath=/html/body/div[2]/nav/div/div[2]/div[2]/div/div[2]/a',
        ],
        "Search icon",
        timeout_each=ELEMENT_TIMEOUT,
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
    await page.wait_for_timeout(500)

    # Step 4: Click the Search submit button
    await click_with_fallback(
        page,
        [
            'a[title="Search in Content Server"]',
            "div.csui-search a",
            'xpath=/html/body/div[2]/nav/div/div[2]/div[2]/div/div[2]/a',
        ],
        "Search submit button",
        timeout_each=ELEMENT_TIMEOUT,
    )
    await page.wait_for_timeout(3000)

    # Step 5: Wait for search results to appear (with smart "0 hits" detection)
    # Check for error dialogs first (e.g. "Network communication failed")
    if await check_and_dismiss_error_dialog(page):
        print(f"[{ts()}]   Error dialog dismissed -- retrying search for {employee_number}")
        await navigate_home(page)
        await page.wait_for_timeout(2000)
        # Retry the search once after dismissing the error
        await search_employee(page, employee_number)
        return

    try:
        # First, explicitly wait for the loading spinner to disappear
        try:
            spinner = page.locator(".csui-loading-parent-wrapper, .binf-list-group-loading, .load-container, .cs-indicator").first
            await spinner.wait_for(state="hidden", timeout=SEARCH_RESULTS_WAIT)
        except Exception:
            pass

        # Then wait for the results container to be visible (broader selectors)
        await page.locator("#results, div.binf-list-group, .csui-search-results").first.wait_for(
            state="visible", timeout=SEARCH_RESULTS_WAIT
        )
    except PlaywrightTimeoutError:
        # Check if the search page shows "About 0 hits" -- meaning the employee
        # genuinely doesn't exist on Content Server (not a timeout issue)
        try:
            zero_hits_loc = page.locator('text=/About 0 hits/i').first
            if await zero_hits_loc.is_visible(timeout=2000):
                print(f"[{ts()}]   Employee {employee_number} not found on Content Server (0 hits) -- skipping")
                raise RuntimeError(
                    f"Employee {employee_number} not found on Content Server (0 search results)"
                )
        except RuntimeError:
            raise
        except Exception:
            pass

        # Check for error dialog that appeared during search
        if await check_and_dismiss_error_dialog(page):
            print(f"[{ts()}]   Error dialog appeared during search -- skipping {employee_number}")
            raise RuntimeError(
                f"Content Server error during search for employee {employee_number}"
            )

        # Genuine timeout -- the server is too slow
        await save_debug_screenshot(page, "no_results", employee_number)
        raise RuntimeError(
            f"Search results did not appear for employee {employee_number} (timeout)"
        )


async def click_employee_master_folder(page, employee_number: str):
    """
    Finds and clicks the employee's master folder in the search results.
    The folder's name should match the employee number exactly.

    Includes a polling wait (up to 15s) to ensure search results have fully
    rendered in the DOM before checking -- this avoids false negatives caused
    by concurrent tab searches overwhelming Content Server.
    """
    print(f"[{ts()}] Looking for exact master folder match: {employee_number}")

    # ── Polling wait: ensure search results are actually rendered ──────────
    # When multiple tabs search simultaneously, the DOM may not have rendered
    # results yet even though the page looks ready. Poll up to 15s.
    for _poll in range(6):  # 6 x 2.5s = 15s max
        # Check for "0 hits" early to avoid waiting the full timeout
        try:
            zero_hits_loc = page.locator('text=/About 0 hits/i').first
            if await zero_hits_loc.is_visible(timeout=500):
                raise RuntimeError(
                    f"Employee {employee_number} not found on Content Server (0 search results)"
                )
        except RuntimeError:
            raise
        except Exception:
            pass

        # Check if any search result link matching the employee number exists
        try:
            loc = page.get_by_role("link", name=employee_number, exact=True).first
            if await loc.is_visible(timeout=500):
                break  # Results are rendered, proceed to click
        except Exception:
            pass

        # Also check title attribute
        try:
            loc = page.locator(f'a[title="{employee_number}"]').first
            if await loc.is_visible(timeout=500):
                break
        except Exception:
            pass

        # Wait before next poll
        await page.wait_for_timeout(2500)

    # ── Strategy 1: Exact text matching using Playwright role locator ─────
    try:
        loc = page.get_by_role("link", name=employee_number, exact=True).first
        if await loc.is_visible(timeout=5000):
            await loc.scroll_into_view_if_needed()
            await loc.click()
            await page.wait_for_timeout(3000)
            return
    except Exception:
        pass

    # ── Strategy 2: Match by exact title / aria-label / data-binf-original-title
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

    # ── Strategy 3: Iterate and check exact inner_text ────────────────────
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

    # Do NOT click the first search result -- it may belong to a different employee.
    # This was causing the script to navigate into wrong folders and fail.
    raise RuntimeError(
        f"Employee folder '{employee_number}' not found in search results -- skipping"
    )


async def click_adani_exit_documents(page, employee_number: Optional[str] = None):
    """
    Click the 'Adani Exit Documents' folder link inside the employee folder.
    This navigates into the subfolder containing the F&F documents.
    """
    # Wait for the table to render first
    try:
        await page.locator("td.csui-table-cell-name").first.wait_for(
            state="visible", timeout=ELEMENT_TIMEOUT
        )
    except PlaywrightTimeoutError:
        await save_debug_screenshot(page, "no_folder_table", employee_number)
        raise RuntimeError("Employee folder table did not load")

    await page.wait_for_timeout(1500)

    # Click the "Adani Exit Documents" link (exact match to avoid clicking files)
    # First try exact role locator
    try:
        exact_folder = page.get_by_role("link", name="Adani Exit Documents", exact=True).first
        if await exact_folder.is_visible(timeout=5000):
            await exact_folder.scroll_into_view_if_needed()
            await exact_folder.click()
            await page.wait_for_timeout(3000)
            return
    except Exception:
        pass

    # Fallback to standard locators with exact matching priority
    await click_with_fallback(
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
        timeout_each=ELEMENT_TIMEOUT,
    )
    await page.wait_for_timeout(3000)


async def get_document_count(page) -> int:
    """Count the number of document rows in the current folder table."""
    try:
        # Wait for the table to be visible
        await page.locator("tbody tr").first.wait_for(
            state="visible", timeout=ELEMENT_TIMEOUT
        )
        await page.wait_for_timeout(1000)

        # Count rows that have a name cell (actual document rows)
        rows = page.locator("tbody tr td.csui-table-cell-name")
        count = await rows.count()
        print(f"[{ts()}] Found {count} document(s)")
        return count
    except PlaywrightTimeoutError:
        print(f"[{ts()}]   No document rows found in the folder")
        return 0


# ── F&F Document Name Filter ─────────────────────────────────────────────────
# Strong regex to match ONLY F&F (Full & Final) sheet documents.
# Matches:  "F AND F", "F&F", "FNF", "F_AND_F", "FULL AND FINAL", "FULL_AND_FINAL"
# Case-insensitive. Rejects experience letters, relieving letters, resignation
# letters, NDC forms, undertakings, and any other non-F&F documents.
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


def _is_ff_document(name: str) -> bool:
    """Return True only if the document name matches the F&F pattern."""
    return bool(FF_DOCUMENT_PATTERN.search(name))


async def download_document_at_row(
    page, row_index: int, total_rows: int, download_dir: Path, employee_number: str
) -> Optional[str]:
    """
    Download a single document from the folder table by its 1-based row index.
    Flow: hover row -> check if folder -> check F&F name -> click "More actions" -> click "Download".

    Only downloads documents whose name matches the F&F regex pattern.
    The row_index is 1-based matching the <tr> nth-of-type in the DOM.
    """
    row_selector = f"tbody tr:nth-of-type({row_index})"
    try:
        row = page.locator(row_selector).first

        # Try scrolling the row into view first to handle lazy-loaded or off-screen rows
        try:
            await row.scroll_into_view_if_needed(timeout=5000)
        except Exception:
            pass

        await row.wait_for(state="visible", timeout=ELEMENT_TIMEOUT)

        # Get the row name for logging
        row_name = f"Row {row_index}"
        name_cell = row.locator("td.csui-table-cell-name a, td.csui-table-cell-name").first
        if await name_cell.count() > 0:
            row_name = (await name_cell.inner_text()).strip()

        # Check if the row represents a folder instead of a file
        is_folder = False
        # Get the second column (icon column)
        icon_cell = row.locator("td:nth-of-type(2), td.csui-table-cell-icon, td.csui-table-cell-type, td:nth-child(2)").first
        if await icon_cell.count() > 0:
            html = await icon_cell.inner_html()
            html_lower = html.lower()
            if "folder" in html_lower or "csui-icon-node-folder" in html_lower:
                is_folder = True

        if is_folder:
            print(f"[{ts()}] Skipping folder: '{row_name}' (row {row_index}/{total_rows})")
            return None

        # ── F&F Filter: Only download F&F sheet documents ─────────────────
        if not _is_ff_document(row_name):
            print(f"[{ts()}] Skipping non-F&F document: '{row_name}' (row {row_index}/{total_rows})")
            return None

        # Create the local folder only now that we know we have a valid F&F document
        download_dir.mkdir(parents=True, exist_ok=True)

        print(f"[{ts()}] Downloading F&F document: '{row_name}' (row {row_index}/{total_rows})...")
        await row.hover()
        await page.wait_for_timeout(1000)
    except Exception as e:
        print(f"[{ts()}]   Could not analyze or hover row {row_index}: {e}")
        return None

    # Click "More actions" (the ... dropdown menu) for this row
    more_actions_selectors = [
        f"tbody tr:nth-of-type({row_index}) li.binf-dropdown > a",
        f'xpath=//tbody/tr[{row_index}]//li[contains(@class,"binf-dropdown")]/a',
        f'xpath=//tbody/tr[{row_index}]//a[@title="More actions"]',
    ]

    try:
        await click_with_fallback(
            page,
            more_actions_selectors,
            f"More actions (row {row_index})",
            timeout_each=8000,
        )
    except RuntimeError:
        # Fallback: try the generic "More actions" locator after hovering
        try:
            more_btn = page.locator('a[title="More actions"]').first
            if await more_btn.is_visible(timeout=5000):
                await more_btn.click()
                print(f"[{ts()}]   Clicked 'More actions' via fallback")
        except Exception as e2:
            print(f"[{ts()}]   Could not click 'More actions' for row {row_index}: {e2}")
            await save_debug_screenshot(page, "more_actions_fail", employee_number)
            return None

    await page.wait_for_timeout(1500)

    # ── Click "Download" from the dropdown menu ──────────────────────────
    # The dropdown items can be hidden until the menu is fully expanded.
    # Strategy: try visible click first, then force-click on hidden elements.
    try:
        async with page.expect_download(timeout=60_000) as dl_info:
            download_clicked = False

            # Attempt 1: Try standard visible selectors
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

            # Attempt 2: Force-click on hidden Download link (common in this UI)
            if not download_clicked:
                force_selectors = [
                    'a[title="Download"]',
                    'xpath=//a[@title="Download"]',
                    'xpath=//a[contains(.,"Download")]',
                ]
                for sel in force_selectors:
                    try:
                        loc = page.locator(sel).first
                        # Check element exists in DOM even if hidden
                        if await loc.count() > 0:
                            await loc.click(force=True, timeout=8000)
                            download_clicked = True
                            print(f"[{ts()}]   Used force-click on Download button")
                            break
                    except Exception:
                        continue

            if not download_clicked:
                # Attempt 3: Use keyboard shortcut or dispatch click event via JS
                try:
                    await page.evaluate("""
                        () => {
                            const link = document.querySelector('a[title="Download"]');
                            if (link) { link.click(); return true; }
                            return false;
                        }
                    """)
                    download_clicked = True
                    print(f"[{ts()}]   Used JavaScript click on Download button")
                except Exception:
                    pass

            if not download_clicked:
                raise RuntimeError(
                    f"Could not click Download for row {row_index} after all attempts"
                )

        download = await dl_info.value
        failure = await download.failure()
        if failure:
            print(f"[{ts()}]   Download failed for row {row_index}: {failure}")
            return None

        suggested = download.suggested_filename or f"ff_document_{row_index}"
        
        # ── Sanitize filename and correct/force extensions ───────────────────
        suggested_clean = suggested.strip()
        # Pattern matching: ". <digits>" suffix (e.g. ". 2")
        match = re.search(r"\.\s*(\d+)$", suggested_clean)
        if match:
            suffix_num = match.group(1)
            base = suggested_clean[:match.start()].strip().rstrip(".")
            suggested_clean = f"{base}_{suffix_num}.pdf"
        else:
            ext = Path(suggested_clean).suffix.lower()
            if ext not in [".pdf", ".xlsx", ".xls", ".doc", ".docx", ".png", ".jpg"]:
                suggested_clean = suggested_clean.rstrip(".")
                suggested_clean = f"{suggested_clean}.pdf"

        save_path = download_dir / suggested_clean

        # Handle duplicate filenames by appending a counter
        if save_path.exists():
            stem = save_path.stem
            suffix = save_path.suffix
            counter = 1
            while save_path.exists():
                save_path = download_dir / f"{stem}_{counter}{suffix}"
                counter += 1

        await download.save_as(save_path)
        size = save_path.stat().st_size
        print(f"[{ts()}] Saved: {save_path.name} ({size:,} bytes)")
        return str(save_path)

    except PlaywrightTimeoutError:
        print(f"[{ts()}]   Download timed out for row {row_index}")
        await save_debug_screenshot(page, "download_timeout", employee_number)
        return None
    except Exception as e:
        print(f"[{ts()}]   Error downloading row {row_index}: {e}")
        return None


async def navigate_home(page):
    """Click the Home icon to return to the Content Server home page."""
    await click_with_fallback(
        page,
        [
            "div.csui-home-item div",
            'a[title="Home page"]',
            "div.csui-home-item a",
            'xpath=/html/body/div[2]/nav/div/div[1]/div[2]/a/div',
        ],
        "Home page icon",
        timeout_each=ELEMENT_TIMEOUT,
    )
    await page.wait_for_timeout(3000)


async def download_ff_for_employee(
    page, employee_number: str, download_dir: Path
) -> Tuple[str, List[str]]:
    """
    Full flow to download all F&F documents for one employee:
      1. Search by employee number
      2. Click first search result
      3. Click "Adani Exit Documents" folder
      4. Download each document
      5. Navigate back home

    Returns a tuple of (status, downloaded_file_paths):
      - (STATUS_DOWNLOADED, [paths])  -- F&F files were found and saved.
      - (STATUS_NO_FF_DOCS, [])       -- Folder found, but no F&F sheet uploaded.
      - (STATUS_NOT_FOUND, [])        -- Content Server returned 0 search hits.
      - (STATUS_ERROR, [])            -- Search timeout, network failure, etc.
    """
    downloaded_files = []
    emp_dir = None
    status = STATUS_ERROR  # Default; overridden on success

    try:
        # Step 1: Search
        await search_employee(page, employee_number)

        # Step 2: Click the exact employee folder matching the ID
        await click_employee_master_folder(page, employee_number)

        # Step 3: Navigate into "Adani Exit Documents"
        try:
            await click_adani_exit_documents(page, employee_number)
        except RuntimeError as e:
            if "Employee folder table did not load" in str(e) or "Adani Exit Documents" in str(e):
                print(f"[{ts()}] [WARN] Failed to load folder table or find Adani Exit Documents. Clicking Home and retrying once...")
                await navigate_home(page)
                await page.wait_for_timeout(3000)
                await search_employee(page, employee_number)
                await click_employee_master_folder(page, employee_number)
                await click_adani_exit_documents(page, employee_number)
            else:
                raise

        # Step 4: Get document count and download each
        doc_count = await get_document_count(page)

        if doc_count == 0:
            print(f"[{ts()}]   No documents found for employee {employee_number}")
            status = STATUS_NO_FF_DOCS
        else:
            # Define the subfolder path, but do NOT create it yet.
            # It will be created in download_document_at_row ONLY if a valid document is found.
            emp_dir = download_dir / employee_number

            for row_idx in range(1, doc_count + 1):
                # Check for error dialogs between downloads (server may show errors under load)
                await check_and_dismiss_error_dialog(page)

                result = await download_document_at_row(
                    page, row_idx, doc_count, emp_dir, employee_number
                )
                if result:
                    downloaded_files.append(result)
                # Small delay between downloads
                await page.wait_for_timeout(2000)

            # Folder was found and iterated -- set status based on whether any F&F file existed
            if downloaded_files:
                status = STATUS_DOWNLOADED
            else:
                # Documents existed (e.g. Resignation Letter, NDC Form) but none matched F&F
                status = STATUS_NO_FF_DOCS

    except RuntimeError as e:
        err_msg = str(e)
        if "0 search results" in err_msg or "0 hits" in err_msg:
            print(f"[{ts()}] Employee {employee_number} not found on Content Server (0 hits)")
            status = STATUS_NOT_FOUND
        else:
            print(f"[{ts()}] Error processing employee {employee_number}: {e}")
            await save_debug_screenshot(page, "error", employee_number)
            status = STATUS_ERROR
    except Exception as e:
        print(f"[{ts()}] Error processing employee {employee_number}: {e}")
        await save_debug_screenshot(page, "error", employee_number)
        status = STATUS_ERROR

    return status, downloaded_files


# ─── MAIN ────────────────────────────────────────────────────────────────────

def prompt_employee_numbers() -> List[str]:
    """
    Interactive terminal prompt -- collects employee ID(s) from the user on a single line.
    Strips any unwanted leading and trailing whitespaces.
    """
    # Use the raw terminal stream so the prompt is visible even with LoggerWriter
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

    # Strip leading/trailing whitespaces from the whole input line
    raw = raw.strip()
    if not raw:
        terminal.write("\n  -> No employee numbers entered.\n\n")
        terminal.flush()
        return []

    # Split by comma or whitespace, and strip each individual token
    collected = []
    for token in raw.replace(",", " ").split():
        token = token.strip()
        if token:
            collected.append(token)

    # De-duplicate while preserving order
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


async def download_ff_reports(employee_numbers: List[str], upload_callback=None):
    """
    Main entry point. Downloads F&F exit documents for one or more employees.
    Uses concurrent browser tabs (up to MAX_CONCURRENT_TABS) for parallel processing.
    After the primary batch, automatically retries employees that had search/network
    errors (STATUS_ERROR) with lower concurrency for reliability.

    Args:
        employee_numbers: List of employee numbers to process.
        upload_callback:  Optional async callback(emp_num, file_list) for uploading.
    """
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

    if not employee_numbers:
        print(f"[{ts()}] No employee numbers provided. Exiting.")
        return

    if not CONTENT_SERVER_URL:
        print(f"[{ts()}] [FATAL ERROR] CONTENT_SERVER_URL is completely blank in your .env file!")
        print(f"[{ts()}] Please open the .env file and add the URL like: CONTENT_SERVER_URL=https://...")
        return

    concurrency = min(MAX_CONCURRENT_TABS, len(employee_numbers))
    print(f"[{ts()}] Starting F&F report download ({'headless' if HEADLESS == 'true' else 'visible'})")
    print(f"[{ts()}] Concurrency: {concurrency} tab(s) for {len(employee_numbers)} employee(s)")

    _cleanup_profile_locks()

    # Structured results: emp_num -> (status, [file_paths])
    all_results: Dict[str, Tuple[str, List[str]]] = {}

    async def _run_batch(
        p_context, emp_list: List[str], num_tabs: int,
        nav_lock: asyncio.Lock, total_label: int, start_offset: int = 0,
        upload_cb=None,
    ):
        """
        Process a list of employees using the given browser context and tab count.
        Updates all_results in-place.
        """
        # Pre-create worker tabs (reuse first page for tab 1)
        first_page = p_context.pages[0] if p_context.pages else await p_context.new_page()
        worker_pages = [first_page]
        for i in range(1, num_tabs):
            try:
                new_tab = await p_context.new_page()
                worker_pages.append(new_tab)
                print(f"[{ts()}]   Opened tab {i + 1}/{num_tabs}")
                await asyncio.sleep(0.5)
            except Exception as e:
                print(f"[{ts()}]   Could not open tab {i + 1}: {e}")
                break

        actual_tabs = len(worker_pages)
        print(f"[{ts()}] {actual_tabs} tab(s) ready. Starting downloads...\n")

        async def _worker(worker_page, emp_num: str, idx: int):
            """Process a single employee in its own browser tab."""
            try:
                print(f"[{ts()}] ── [Tab] Employee {idx}/{total_label}: {emp_num} ──")

                # Navigate to Content Server home (serialised via lock)
                nav_success = False
                for nav_attempt in range(2):
                    try:
                        async with nav_lock:
                            await worker_page.goto(
                                CONTENT_SERVER_URL,
                                wait_until="domcontentloaded",
                                timeout=NAV_TIMEOUT,
                            )
                            await worker_page.wait_for_timeout(2000)
                            await wait_for_content_server_home(worker_page, timeout_s=30)
                        nav_success = True
                        break
                    except Exception as nav_err:
                        if nav_attempt == 0:
                            print(f"[{ts()}]   [Tab] Navigation timeout for {emp_num} -- retrying once...")
                            await worker_page.wait_for_timeout(3000)
                        else:
                            raise RuntimeError(
                                f"Content Server navigation failed after 2 attempts for {emp_num}: {nav_err}"
                            )

                if not nav_success:
                    raise RuntimeError(f"Could not navigate to Content Server for {emp_num}")

                # Download F&F docs for this employee
                status, downloaded = await download_ff_for_employee(
                    worker_page, emp_num, DOWNLOAD_DIR
                )
                all_results[emp_num] = (status, downloaded)

                if downloaded and upload_cb:
                    try:
                        print(f"[{ts()}] [Tab] Triggering upload for {emp_num}...")
                        await upload_cb(emp_num, downloaded)
                    except Exception as e:
                        print(f"[{ts()}] [Tab] Upload callback failed for {emp_num}: {e}")

            except Exception as e:
                print(f"[{ts()}] [Tab] Error for {emp_num}: {e}")
                all_results[emp_num] = (STATUS_ERROR, [])

        # Process employees in batches matching available tabs
        for batch_start in range(0, len(emp_list), actual_tabs):
            batch = emp_list[batch_start : batch_start + actual_tabs]
            batch_tasks = []
            for i, emp_num in enumerate(batch):
                page = worker_pages[i]
                idx = start_offset + batch_start + i + 1
                batch_tasks.append(_worker(page, emp_num, idx))

            await asyncio.gather(*batch_tasks)

            # Small staggered delay between batches to avoid overwhelming Content Server
            await asyncio.sleep(1.5)

        # Close extra tabs (keep first page for potential retry reuse)
        for wp in worker_pages[1:]:
            try:
                await wp.close()
            except Exception:
                pass

    try:
        async with async_playwright() as p:
            print(f"[{ts()}] Launching browser...")
            chrome_args = [
                "--no-first-run",
                "--no-default-browser-check",
                "--disable-blink-features=AutomationControlled",
                "--disable-session-crashed-bubble",
                "--hide-crash-restore-bubble",
                "--window-size=1517,900",
                "--window-position=-32000,-32000",
            ]
            is_headless = HEADLESS == "true"

            context = await p.chromium.launch_persistent_context(
                user_data_dir=str(AUTOMATION_PROFILE_DIR),
                # channel="chrome",
                headless=is_headless,
                accept_downloads=True,
                args=chrome_args,
            )

            # ── Step 1: Establish login / SSO on the first tab ────────────
            first_page = context.pages[0] if context.pages else await context.new_page()

            print(f"[{ts()}] Navigating to Content Server...")
            await first_page.goto(
                CONTENT_SERVER_URL, wait_until="domcontentloaded", timeout=NAV_TIMEOUT
            )
            await first_page.wait_for_timeout(3000)

            await wait_for_content_server_home(first_page, timeout_s=120)
            await first_page.wait_for_timeout(2000)
            print(f"[{ts()}] Login confirmed. Preparing {concurrency} tab(s)...\n")

            nav_lock = asyncio.Lock()

            # ── Step 2: Primary batch ─────────────────────────────────────
            await _run_batch(
                context, employee_numbers, concurrency,
                nav_lock, len(employee_numbers),
                upload_cb=upload_callback,
            )

            # -- Step 3: Retry loop -- keep retrying until errors/ff_error/ is empty --
            # Scans the error screenshot folder for employee IDs and retries them.
            # This also picks up leftover screenshots from previous runs.
            MAX_RETRY_ROUNDS = 3
            for retry_round in range(1, MAX_RETRY_ROUNDS + 1):
                # Scan the error folder for remaining failed IDs
                folder_failed_ids = get_failed_ids_from_error_folder()
                if not folder_failed_ids:
                    print(f"[{ts()}] errors/ff_error/ is empty -- no retries needed.")
                    break

                retry_concurrency = min(2, len(folder_failed_ids))  # Lower concurrency for stability
                print()
                print(f"[{ts()}] ═══════════════════════════════════════════════════════════")
                print(f"[{ts()}] RETRY ROUND {retry_round}/{MAX_RETRY_ROUNDS}: {len(folder_failed_ids)} error screenshot(s) found")
                print(f"[{ts()}]   IDs: {', '.join(folder_failed_ids)}")
                print(f"[{ts()}]   Retrying with {retry_concurrency} tab(s) for reliability...")
                print(f"[{ts()}] ═══════════════════════════════════════════════════════════")
                print()

                # Navigate first page back home before retry
                try:
                    remaining_pages = context.pages
                    if remaining_pages:
                        retry_page = remaining_pages[0]
                        await retry_page.goto(
                            CONTENT_SERVER_URL, wait_until="domcontentloaded", timeout=NAV_TIMEOUT
                        )
                        await retry_page.wait_for_timeout(2000)
                        await wait_for_content_server_home(retry_page, timeout_s=30)
                except Exception:
                    pass

                await _run_batch(
                    context, folder_failed_ids, retry_concurrency,
                    nav_lock, len(folder_failed_ids),
                    upload_cb=upload_callback,
                )

                # Clean up error screenshots for employees that succeeded on retry
                resolved_count = 0
                for emp_num in folder_failed_ids:
                    retry_status, _ = all_results.get(emp_num, (STATUS_ERROR, []))
                    if retry_status != STATUS_ERROR:
                        print(f"[{ts()}]   [OK] {emp_num} resolved on retry (status: {retry_status})")
                        cleanup_error_screenshots(emp_num)
                        resolved_count += 1

                print(f"[{ts()}]   Round {retry_round} complete: {resolved_count}/{len(folder_failed_ids)} resolved")

                # Check if error folder is now empty
                remaining = get_failed_ids_from_error_folder()
                if not remaining:
                    print(f"[{ts()}] errors/ff_error/ is now empty -- all retries succeeded!")
                    break
                elif retry_round < MAX_RETRY_ROUNDS:
                    print(f"[{ts()}]   {len(remaining)} still remaining, will retry next round...")
                else:
                    print(f"[{ts()}]   {len(remaining)} still remaining after {MAX_RETRY_ROUNDS} retry rounds.")

            # ── Close browser ─────────────────────────────────────────────
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
        print(f"[{ts()}] FATAL: {e}\n{traceback.format_exc()}")

    # ── Categorized Summary ───────────────────────────────────────────────────
    _print_categorized_summary(all_results)


def _print_categorized_summary(all_results: Dict[str, Tuple[str, List[str]]]):
    """Print a clear, categorized download summary."""
    downloaded = {k: v for k, (s, v) in all_results.items() if s == STATUS_DOWNLOADED}
    no_ff      = {k: v for k, (s, v) in all_results.items() if s == STATUS_NO_FF_DOCS}
    not_found  = {k: v for k, (s, v) in all_results.items() if s == STATUS_NOT_FOUND}
    errors     = {k: v for k, (s, v) in all_results.items() if s == STATUS_ERROR}

    total_files = sum(len(v) for v in downloaded.values())
    total_emps  = len(all_results)

    print()
    print(f"[{ts()}] ═══════════════════════════════════════════════════════════")
    print(f"[{ts()}] DOWNLOAD SUMMARY")
    print(f"[{ts()}] ═══════════════════════════════════════════════════════════")

    if downloaded:
        print(f"[{ts()}]")
        print(f"[{ts()}]   [OK] DOWNLOADED ({len(downloaded)} employee(s), {total_files} file(s)):")
        for emp, files in downloaded.items():
            print(f"[{ts()}]      {emp}: {len(files)} file(s)")

    if no_ff:
        print(f"[{ts()}]")
        print(f"[{ts()}]   [INFO] NO F&F DOCUMENT ({len(no_ff)} employee(s) -- folder found, no F&F sheet uploaded):")
        for emp in no_ff:
            print(f"[{ts()}]      {emp}")

    if not_found:
        print(f"[{ts()}]")
        print(f"[{ts()}]   [--] NOT FOUND ON CONTENT SERVER ({len(not_found)} employee(s) -- 0 search hits):")
        for emp in not_found:
            print(f"[{ts()}]      {emp}")

    if errors:
        print(f"[{ts()}]")
        print(f"[{ts()}]   [FAIL] FAILED / ERRORS ({len(errors)} employee(s) -- search timeout or network failure):")
        for emp in errors:
            print(f"[{ts()}]      {emp}")

    print(f"[{ts()}] ───────────────────────────────────────────────────────────")
    print(f"[{ts()}]   Total: {total_files} file(s) downloaded for {total_emps} employee(s)")
    if errors:
        print(f"[{ts()}]   [WARN] {len(errors)} employee(s) still have unresolved errors -- check errors/ff_error/ for screenshots")
    print(f"[{ts()}] ═══════════════════════════════════════════════════════════")


if __name__ == "__main__":
    numbers = prompt_employee_numbers()
    if numbers:
        asyncio.run(download_ff_reports(numbers))
    else:
        print("No employee numbers entered. Exiting.")
