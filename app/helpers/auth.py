"""
Microsoft Entra ID / SSO authentication and MFA push approval state machine.
"""

import asyncio
from datetime import datetime
from pathlib import Path
from typing import Optional

from playwright.async_api import Page

from app.core.config import ERRORS_DIR


class MicrosoftSSOHandler:
    """Manages Microsoft Entra SSO, MFA push verification, and KMSI prompt."""

    def __init__(
        self,
        oracle_email: str,
        oracle_password: str,
        error_dir: Optional[Path] = None,
        max_rsa_retries: int = 3,
    ):
        self.oracle_email = oracle_email
        self.oracle_password = oracle_password
        self.error_dir = error_dir or (ERRORS_DIR / "sso_error")
        self.email_filled = False
        self.password_filled = False
        self.rsa_clicked = False
        self.rsa_retry_count = 0
        self.max_rsa_retries = max_rsa_retries

    @staticmethod
    def _ts() -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    async def _save_debug_screenshot(self, page: Page, label: str) -> str:
        try:
            self.error_dir.mkdir(parents=True, exist_ok=True)
            path = self.error_dir / f"debug_{label}_{datetime.now().strftime('%H%M%S')}.png"
            await page.screenshot(path=str(path), full_page=True)
            print(f"[{self._ts()}]   Screenshot saved: {path.name}")
            return str(path)
        except Exception as e:
            print(f"[{self._ts()}]   Could not save screenshot: {e}")
            return ""

    async def check_for_sso_errors(
        self,
        page: Page,
        check_password_error: Optional[bool] = None,
        check_username_error: Optional[bool] = None,
    ):
        """Checks if Microsoft SSO displays validation or security error messages."""
        try:
            if check_password_error is None:
                check_password_error = self.password_filled
            if check_username_error is None:
                check_username_error = self.email_filled

            error_ids = ["error"]
            if check_username_error:
                error_ids.append("usernameError")
            if check_password_error:
                error_ids.append("passwordError")

            for error_id in error_ids:
                try:
                    loc = page.locator(f"#{error_id}").first
                    if await loc.count() > 0 and await loc.is_visible(timeout=500):
                        err_text = (await loc.inner_text()).strip().replace("\n", " ")
                        if err_text:
                            # Skip harmless validation prompts if submission has not yet occurred
                            if not check_password_error and "password" in err_text.lower():
                                continue
                            if not check_username_error and ("email" in err_text.lower() or "skype" in err_text.lower()):
                                continue
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
                        if not check_password_error and "password" in text.lower():
                            pass
                        elif not check_username_error and ("email" in text.lower() or "skype" in text.lower()):
                            pass
                        else:
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

            # ==== KMSI Prompt Check ====
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

            # ==== RSA Entra MFA Method Selection Check ====
            try:
                rsa_btn = page.locator('div[role="button"], button, a').filter(
                    has_text="RSAEntraMFA"
                ).first
                if await rsa_btn.count() > 0 and await rsa_btn.is_visible(timeout=300):
                    return "RSA_SELECTION"
            except Exception:
                pass

            # ==== RSA Phone Push Approval Waiting Check ====
            if any(x in current_url for x in ["securid.com", "adani.auth"]) or self.rsa_clicked:
                return "RSA_WAIT"

            # ==== Microsoft Authenticator Number Match Check ====
            try:
                auth_sign = page.locator(
                    '#idRichContext_DisplaySign, div.displaySign, '
                    '#loginHeader:has-text("Approve sign in request"), '
                    'div[role="heading"]:has-text("Approve sign in request")'
                ).first
                if await auth_sign.count() > 0 and await auth_sign.is_visible(timeout=300):
                    return "MFA_NUMBER"
            except Exception:
                pass

            # Reset email_filled if we are explicitly on the Sign In (email) screen
            try:
                sign_in_header = page.locator(
                    '#loginHeader:has-text("Sign in"), div[role="heading"]:has-text("Sign in")'
                ).first
                if await sign_in_header.count() > 0 and await sign_in_header.is_visible(timeout=200):
                    self.email_filled = False
            except Exception:
                pass

            # ==== Account Picker Screen Check ====
            try:
                pick_header = page.locator(
                    '#loginHeader:has-text("Pick an account"), '
                    'div[role="heading"]:has-text("Pick an account"), '
                    '#loginHeader:has-text("Choose an account"), '
                    '#otherTile, div#otherTile, #tilesHolder'
                ).first
                pw_input_chk = page.locator('input[name="passwd"], input[type="password"]').first
                if await pick_header.count() > 0 and await pick_header.is_visible(timeout=300):
                    if not (await pw_input_chk.count() > 0 and await pw_input_chk.is_visible(timeout=200)):
                        return "PICK_ACCOUNT"
            except Exception:
                pass

            # ==== Password Input Screen Check ====
            try:
                pw_input = page.locator(
                    'input[name="passwd"], input[id="i0118"], input[type="password"]'
                ).first
                pw_header = page.locator(
                    '#loginHeader:has-text("Enter password"), '
                    'div[role="heading"]:has-text("Enter password"), '
                    '#loginHeader:has-text("password"), '
                    '#idSIButton9[value="Sign in"]'
                ).first
                if await pw_input.count() > 0 and await pw_input.is_visible(timeout=300):
                    return "PASSWORD"
                if await pw_header.count() > 0 and await pw_header.is_visible(timeout=300):
                    return "PASSWORD"
            except Exception:
                pass

            # ==== Email Input Screen Check ====
            try:
                email_input = page.locator(
                    'input[name="loginfmt"], input[type="email"], input[id="i0116"]'
                ).first
                if not self.email_filled and await email_input.count() > 0 and await email_input.is_visible(timeout=300):
                    return "EMAIL"
            except Exception:
                pass

            return "UNKNOWN"
        except Exception:
            return "UNKNOWN"

    async def _handle_pick_account(self, page: Page):
        """Selects the target account tile if listed on account picker screen."""
        try:
            target_email = self.oracle_email.strip().lower()
            target_user = target_email.split("@")[0] if "@" in target_email else target_email
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
                        if target_email in tile_text or (target_user and target_user in tile_text):
                            print(f"[{self._ts()}]   Selecting account: {self.oracle_email}")
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
                        f'small:has-text("{self.oracle_email}"), div:has-text("{self.oracle_email}")'
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
                        print(f"[{self._ts()}]   Selected account: {self.oracle_email}")
                except Exception:
                    pass

            if not account_found:
                try:
                    use_another_selectors = [
                        "#otherTile",
                        "div#otherTile",
                        'div:has-text("Use another account")',
                        'div[role="button"]:has-text("Use another account")',
                        'button:has-text("Use another account")',
                        '[data-test-id="otherTile"]',
                        "#tilesHolder div.table",
                    ]
                    for u_sel in use_another_selectors:
                        try:
                            use_another = page.locator(u_sel).first
                            if await use_another.count() > 0 and await use_another.is_visible(timeout=1000):
                                print(f"[{self._ts()}]   Clicking 'Use another account'...")
                                await use_another.click(force=True, timeout=5000)
                                self.email_filled = False
                                self.password_filled = False
                                break
                        except Exception:
                            continue
                except Exception:
                    pass

            await page.wait_for_timeout(2000)
        except Exception as e:
            print(f"[{self._ts()}]   Warning during account pick: {e}")

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
                        try:
                            await page.keyboard.press("Control+A")
                            await page.keyboard.press("Backspace")
                        except Exception:
                            pass
                        await email_input.fill("", force=True, timeout=3000)
                        await page.wait_for_timeout(200)
                        await email_input.type(self.oracle_email.strip(), delay=40)
                        print(f"[{self._ts()}]   Entered email ({self.oracle_email})")
                        await page.wait_for_timeout(300)

                        next_btn = page.locator('input[type="submit"], #idSIButton9').first
                        await next_btn.click(force=True, timeout=5000)
                        await page.wait_for_timeout(2500)

                        await self.check_for_sso_errors(page, check_password_error=False, check_username_error=True)
                        self.email_filled = True
                        break
                    except RuntimeError:
                        if attempt == 3:
                            raise
                        print(f"[{self._ts()}]   Email submission retry {attempt}/3...")
                        await page.wait_for_timeout(1000)
        except RuntimeError:
            raise
        except Exception as e:
            print(f"[{self._ts()}]   Warning during email input: {e}")

    async def _handle_password(self, page: Page):
        """Fills the password field and submits."""
        try:
            target_email = self.oracle_email.strip().lower()
            target_user = target_email.split("@")[0] if "@" in target_email else target_email

            # Verify if the displayed account matches target Oracle email
            try:
                identity_locs = [
                    "#displayName",
                    "div.identity",
                    ".identity-banner",
                    "#loginHeader + div",
                    'div[role="heading"] + div',
                    'div:has-text("@")',
                    'small:has-text("@")',
                ]
                current_account = ""
                for sel in identity_locs:
                    try:
                        loc = page.locator(sel).first
                        if await loc.count() > 0 and await loc.is_visible(timeout=300):
                            text = (await loc.inner_text()).strip().lower()
                            title = (await loc.get_attribute("title") or "").strip().lower()
                            for candidate in [text, title]:
                                if candidate and candidate != "enter password":
                                    current_account = candidate
                                    break
                            if current_account:
                                break
                    except Exception:
                        continue

                # If a different account is active on the password screen, click Back to switch
                is_target_account = (
                    target_email in current_account
                    or (target_user and target_user in current_account)
                )
                if current_account and not is_target_account:
                    print(
                        f"[{self._ts()}]   Previous session with different account detected ({current_account}) — switching to {self.oracle_email}"
                    )
                    back_btn_selectors = [
                        "#idBtn_Back",
                        'button[id="idBtn_Back"]',
                        'button[aria-label="Back"]',
                        'input[id="idBtn_Back"]',
                        "div.backButton",
                        "button.backButton",
                        'button:has-text("Back")',
                        'div[role="button"][aria-label="Back"]',
                        "#idA_PWD_SwitchToCredPicker",
                        'a:has-text("Sign-in options")',
                    ]
                    for b_sel in back_btn_selectors:
                        try:
                            b_loc = page.locator(b_sel).first
                            if await b_loc.count() > 0 and await b_loc.is_visible(timeout=500):
                                await b_loc.click(force=True, timeout=3000)
                                print(f"[{self._ts()}]   Clicked Back button to switch account")
                                await page.wait_for_timeout(2000)
                                self.email_filled = False
                                self.password_filled = False
                                return
                        except Exception:
                            continue
            except Exception as e:
                print(f"[{self._ts()}]   Warning verifying active account: {e}")

            pw_input = page.locator(
                'input[name="passwd"], input[id="i0118"], input[type="password"]'
            ).first
            if await pw_input.is_visible(timeout=5000):
                for attempt in range(1, 4):
                    try:
                        await pw_input.click(force=True, timeout=3000)
                        try:
                            await page.keyboard.press("Control+A")
                            await page.keyboard.press("Backspace")
                        except Exception:
                            pass
                        await pw_input.fill("", force=True, timeout=3000)
                        await page.wait_for_timeout(200)
                        await pw_input.type(self.oracle_password.strip(), delay=40)
                        print(f"[{self._ts()}]   Entered password (Attempt {attempt}/3)")
                        await page.wait_for_timeout(300)

                        signin_btn = page.locator('input[type="submit"], #idSIButton9').first
                        await signin_btn.click(force=True, timeout=5000)
                        await page.wait_for_timeout(2500)

                        self.password_filled = True
                        await self.check_for_sso_errors(page, check_password_error=True)
                        break
                    except RuntimeError:
                        if attempt == 3:
                            raise
                        print(f"[{self._ts()}]   Password submission retry {attempt}/3...")
                        await page.wait_for_timeout(1000)
        except RuntimeError:
            raise
        except Exception as e:
            print(f"[{self._ts()}]   Warning during password input: {e}")

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
            print(f"[{self._ts()}]   Warning during RSA selection: {e}")

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
                print(f"[{self._ts()}]   Confirmed 'Stay signed in? -> Yes'")
                await page.wait_for_timeout(2500)
        except Exception as e:
            print(f"[{self._ts()}]   Warning during KMSI confirmation: {e}")

    async def authenticate(self, page: Page, timeout_s: int = 300) -> bool:
        """Coordinates the complete Microsoft SSO authentication lifecycle."""
        try:
            if not self.oracle_email or not self.oracle_password:
                print(f"[{self._ts()}] ⚠ No credentials in .env — waiting for manual login...")
                return False

            self.email_filled = False
            self.password_filled = False
            self.rsa_clicked = False
            self.rsa_retry_count = 0

            print(f"[{self._ts()}] SSO login detected — completing authentication...")
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
                                f"[{self._ts()}] ⚠ RSA MFA failed or rejected (Attempt {self.rsa_retry_count}/{self.max_rsa_retries})."
                            )
                            if self.rsa_retry_count >= self.max_rsa_retries:
                                await self._save_debug_screenshot(page, "rsa_failed_3_times")
                                raise RuntimeError("RSA/MFA Authentication failed 3 times.")
                            print(f"[{self._ts()}]   Clicking 'Retry' for RSA MFA...")
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

                elif screen == "MFA_NUMBER":
                    try:
                        num_el = page.locator('#idRichContext_DisplaySign, div.displaySign').first
                        num_val = ""
                        if await num_el.count() > 0 and await num_el.is_visible(timeout=500):
                            num_val = (await num_el.inner_text()).strip()
                        if num_val:
                            print(f"[{self._ts()}] [MFA] Microsoft Authenticator prompt: Enter number [{num_val}] on your Authenticator app.")
                        else:
                            print(f"[{self._ts()}] [MFA] Microsoft Authenticator prompt: Please approve the request on your Authenticator app.")
                    except Exception:
                        pass
                    await page.wait_for_timeout(3000)

                elif screen == "KMSI":
                    await self._handle_kmsi(page)

                await page.wait_for_timeout(1000)

            return False
        except Exception as e:
            print(f"[{self._ts()}] SSO Authentication Exception: {e}")
            raise

