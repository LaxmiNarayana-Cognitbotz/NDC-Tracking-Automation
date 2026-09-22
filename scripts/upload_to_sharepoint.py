# SharePoint File Upload Script using Microsoft Graph API

import argparse
import asyncio
import os
import re
import sys
import traceback
import urllib.parse
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple

import httpx
import msal
import requests
import urllib3
from dotenv import load_dotenv

# Suppress insecure HTTPS request warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


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
LOG_FILE = LOG_DIR / "upload_to_sharepoint.log"

sys.stdout = LoggerWriter(LOG_FILE)
sys.stderr = LoggerWriter(LOG_FILE)

load_dotenv(verbose=True)


# Application Configuration
class SharePointConfig:
    """Encapsulates SharePoint Graph API credentials, paths, and cached IDs."""

    def __init__(self):
        try:
            self.tenant_id: str = (os.getenv("SHAREPOINT_TENANT_ID") or "").strip()
            self.site_url: str = (os.getenv("SHAREPOINT_SITE_URL") or "").strip()
            self.client_id: str = (os.getenv("SHAREPOINT_CLIENT_ID") or "").strip()
            self.client_secret: str = (os.getenv("SHAREPOINT_CLIENT_SECRET") or "").strip()
            self.target_folder: str = (os.getenv("SHAREPOINT_TARGET_FOLDER") or "").strip()

            self.authority: str = (
                f"https://login.microsoftonline.com/{self.tenant_id}" if self.tenant_id else ""
            )
            self.scopes = ["https://graph.microsoft.com/.default"]

            # Local directories
            self.uploads_dir: Path = Path(__file__).parent.parent / "uploads"
            self.ndc_reports_dir: Path = self.uploads_dir / "NDC_Reports"
            self.ff_reports_dir: Path = self.uploads_dir / "FF_Reports"

            self.ndc_reports_dir.mkdir(parents=True, exist_ok=True)
            self.ff_reports_dir.mkdir(parents=True, exist_ok=True)

            # In-memory ID cache
            self.cached_site_id: Optional[str] = None
            self.cached_drive_id: Optional[str] = None
        except Exception as e:
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] SharePointConfig Init Error: {e}")


# Automation Utilities
class SharePointUtils:
    """Helper utilities for timestamps, path encoding, and local file operations."""

    @staticmethod
    def ts() -> str:
        """Returns the current formatted timestamp string."""
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    @staticmethod
    def encode_path_segments(path: str) -> str:
        """URL-encodes each path segment individually while preserving slashes."""
        try:
            segments = [urllib.parse.quote(s) for s in path.split("/") if s]
            return "/".join(segments)
        except Exception:
            return path

    @staticmethod
    def get_latest_report(directory: Path, pattern: str = "*.xls") -> Path:
        """Finds the most recently modified report file in the specified directory."""
        try:
            if not directory.exists():
                raise FileNotFoundError(f"Local reports directory '{directory}' does not exist.")

            files = list(directory.glob(pattern))
            if not files:
                files = list(directory.glob("*.xlsx"))
                if not files:
                    raise FileNotFoundError(
                        f"No files matching '{pattern}' or '*.xlsx' found in '{directory}'."
                    )

            files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            return files[0]
        except Exception:
            raise

    @staticmethod
    def cleanup_old_ndc_reports(latest_file: Path, directory: Path):
        """Removes older local report files from the NDC reports directory."""
        try:
            latest_str = str(latest_file.resolve()).lower()
            for f in directory.iterdir():
                if f.is_file() and str(f.resolve()).lower() != latest_str:
                    try:
                        f.unlink()
                        print(f"[{SharePointUtils.ts()}] [CLEANUP] Deleted local file: {f.name}")
                    except Exception as e:
                        print(f"[{SharePointUtils.ts()}] [CLEANUP] Warning: could not delete {f.name}: {e}")
        except Exception as e:
            print(f"[{SharePointUtils.ts()}] [CLEANUP] Local cleanup error: {e}")


# SharePoint Authentication Manager
class SharePointAuthManager:
    """Handles MSAL client credentials authentication and token acquisition."""

    def __init__(self, config: SharePointConfig):
        self.config = config
        self._session = requests.Session()
        self._session.verify = False

    async def get_access_token(self) -> str:
        """Obtains a Microsoft Graph API access token using client credentials flow."""
        try:
            app = msal.ConfidentialClientApplication(
                self.config.client_id,
                authority=self.config.authority,
                client_credential=self.config.client_secret,
                http_client=self._session,
            )

            result = app.acquire_token_silent(self.config.scopes, account=None)
            if not result:
                result = await asyncio.to_thread(
                    app.acquire_token_for_client, scopes=self.config.scopes
                )

            if "access_token" in result:
                return result["access_token"]
            else:
                error_msg = (
                    result.get("error_description") or result.get("error") or "Unknown authentication error"
                )
                raise RuntimeError(f"SharePoint authentication failed: {error_msg}")
        except Exception as e:
            if not isinstance(e, RuntimeError):
                raise RuntimeError(f"SharePoint token acquisition exception: {e}")
            raise


# SharePoint Operations Service
class SharePointService:
    """Coordinates Microsoft Graph API site lookup, drive resolution, file uploads, and cleanup."""

    def __init__(self, config: Optional[SharePointConfig] = None):
        self.config = config or SharePointConfig()
        self.auth = SharePointAuthManager(self.config)

    async def get_access_token(self) -> str:
        """Convenience method to retrieve Graph API access token."""
        return await self.auth.get_access_token()

    async def get_site_id(self, client: httpx.AsyncClient) -> str:
        """Resolves the SharePoint Site ID from the configured SITE_URL."""
        try:
            if self.config.cached_site_id:
                return self.config.cached_site_id

            parsed_url = urllib.parse.urlparse(self.config.site_url)
            hostname = parsed_url.netloc
            relative_path = parsed_url.path.rstrip("/")

            token = await self.auth.get_access_token()
            headers = {"Authorization": f"Bearer {token}"}

            url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{relative_path}"
            response = await client.get(url, headers=headers)
            if response.status_code != 200:
                raise RuntimeError(f"Failed to resolve site ID: {response.status_code} - {response.text}")

            site_data = response.json()
            site_id = site_data.get("id")
            if not site_id:
                raise RuntimeError("Site details resolved but 'id' field is missing.")

            self.config.cached_site_id = site_id
            return site_id
        except Exception as e:
            if not isinstance(e, RuntimeError):
                raise RuntimeError(f"Error resolving SharePoint site ID: {e}")
            raise

    async def get_drive_details(self, client: httpx.AsyncClient, site_id: str) -> Tuple[str, str]:
        """Resolves the Drive ID and relative folder path within that drive."""
        try:
            if self.config.cached_drive_id:
                return self.config.cached_drive_id, self.config.target_folder

            parsed_site = urllib.parse.urlparse(self.config.site_url)
            site_path = parsed_site.path.rstrip("/")

            folder = self.config.target_folder
            if folder.startswith(site_path):
                folder = folder[len(site_path) :]
            folder = folder.strip("/")

            token = await self.auth.get_access_token()
            headers = {"Authorization": f"Bearer {token}"}

            url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
            response = await client.get(url, headers=headers)
            if response.status_code != 200:
                raise RuntimeError(f"Failed to list drives: {response.status_code} - {response.text}")

            drives = response.json().get("value", [])
            drive_id = None
            folder_path_in_drive = ""

            for d in drives:
                drive_web_url = d.get("webUrl", "")
                parsed_drive = urllib.parse.urlparse(drive_web_url)
                drive_path = parsed_drive.path.rstrip("/")

                drive_rel_path = ""
                if drive_path.startswith(site_path):
                    drive_rel_path = drive_path[len(site_path) :].strip("/")

                drive_name = d.get("name", "")

                candidates = []
                if drive_rel_path:
                    candidates.append(drive_rel_path)
                    candidates.append(urllib.parse.unquote(drive_rel_path))
                    candidates.append(urllib.parse.quote(drive_rel_path))
                if drive_name:
                    candidates.append(drive_name)

                for candidate in candidates:
                    candidate_clean = candidate.strip("/")
                    if not candidate_clean:
                        continue

                    if folder == candidate_clean:
                        drive_id = d.get("id")
                        folder_path_in_drive = ""
                        break
                    elif folder.startswith(candidate_clean + "/"):
                        drive_id = d.get("id")
                        folder_path_in_drive = folder[len(candidate_clean) :].strip("/")
                        break

                if drive_id:
                    break

            if not drive_id:
                default_drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
                res = await client.get(default_drive_url, headers=headers)
                if res.status_code == 200:
                    drive_id = res.json().get("id")
                    folder_path_in_drive = folder
                else:
                    raise RuntimeError(f"Failed to resolve site default drive: {res.text}")

            try:
                encoded = SharePointUtils.encode_path_segments(folder_path_in_drive)
                check_url = (
                    f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{encoded}"
                )
                check_res = await client.get(check_url, headers=headers)
                if check_res.status_code == 404:
                    normalized = re.sub(r"\s+", " ", folder_path_in_drive)
                    if normalized != folder_path_in_drive:
                        norm_encoded = SharePointUtils.encode_path_segments(normalized)
                        norm_res = await client.get(
                            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{norm_encoded}",
                            headers=headers,
                        )
                        if norm_res.status_code == 200:
                            folder_path_in_drive = normalized
            except Exception:
                pass

            self.config.cached_drive_id = drive_id
            return drive_id, folder_path_in_drive
        except Exception as e:
            if not isinstance(e, RuntimeError):
                raise RuntimeError(f"Error resolving SharePoint drive details: {e}")
            raise

    async def upload_file(
        self,
        client: httpx.AsyncClient,
        site_id: str,
        drive_id: str,
        target_folder_url: str,
        local_file_path: Path,
    ):
        """Uploads a local file to the designated SharePoint folder path."""
        try:
            file_name = local_file_path.name
            file_size = local_file_path.stat().st_size

            clean_folder = "/".join([p for p in target_folder_url.split("/") if p])
            target_file_path = f"{clean_folder}/{file_name}"
            encoded_path = SharePointUtils.encode_path_segments(target_file_path)

            token = await self.auth.get_access_token()

            with open(local_file_path, "rb") as f:
                file_data = f.read()

            if file_size <= 4 * 1024 * 1024:  # 4MB
                url = (
                    f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}:/content"
                )
                headers = {
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/octet-stream",
                }
                response = await client.put(url, headers=headers, content=file_data)
                if response.status_code not in [200, 201]:
                    raise RuntimeError(f"Upload failed (PUT): {response.status_code} - {response.text}")
            else:
                session_url = (
                    f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}:/createUploadSession"
                )
                session_payload = {"item": {"@microsoft.graph.conflictBehavior": "replace"}}
                session_response = await client.post(
                    session_url,
                    headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                    json=session_payload,
                )
                if session_response.status_code != 200:
                    raise RuntimeError(
                        f"Failed to create upload session: {session_response.status_code} - {session_response.text}"
                    )

                upload_url = session_response.json().get("uploadUrl")
                chunk_size = 3276800  # Multiples of 327,680 bytes
                for i in range(0, file_size, chunk_size):
                    chunk = file_data[i : i + chunk_size]
                    content_range = f"bytes {i}-{i + len(chunk) - 1}/{file_size}"
                    chunk_headers = {"Content-Length": str(len(chunk)), "Content-Range": content_range}
                    chunk_res = await client.put(upload_url, headers=chunk_headers, content=chunk)
                    if chunk_res.status_code not in [200, 201, 202]:
                        raise RuntimeError(
                            f"Chunk upload failed (Session): {chunk_res.status_code} - {chunk_res.text}"
                        )
        except Exception as e:
            if not isinstance(e, RuntimeError):
                raise RuntimeError(f"Error uploading file '{local_file_path.name}': {e}")
            raise

    async def cleanup_sharepoint_folder_files(
        self, client: httpx.AsyncClient, site_id: str, drive_id: str, folder_url: str
    ):
        """Cleans up existing files in a SharePoint target folder prior to uploading."""
        try:
            token = await self.auth.get_access_token()
            headers = {"Authorization": f"Bearer {token}"}
            encoded_path = SharePointUtils.encode_path_segments(folder_url)

            url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}:/children"
            response = await client.get(url, headers=headers)

            if response.status_code == 200:
                children = response.json().get("value", [])
                files = [item for item in children if "file" in item]
                deleted_count = 0
                for f in files:
                    item_id = f.get("id")
                    file_name = f.get("name", "unknown")
                    delete_url = (
                        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{item_id}"
                    )
                    del_res = await client.delete(delete_url, headers=headers)
                    if del_res.status_code in [200, 204]:
                        deleted_count += 1
                        print(f"[{SharePointUtils.ts()}]   [CLEANUP] Deleted existing SharePoint file: {file_name}")
                    else:
                        print(
                            f"[{SharePointUtils.ts()}]   [CLEANUP] Could not delete SharePoint file {file_name}: {del_res.text}"
                        )
                if deleted_count > 0:
                    print(
                        f"[{SharePointUtils.ts()}]   [CLEANUP] Cleared {deleted_count} existing file(s) from folder: {folder_url}"
                    )
            elif response.status_code == 404:
                pass
            else:
                print(
                    f"[{SharePointUtils.ts()}]   [CLEANUP] Notice on folder query ({response.status_code}): {response.text}"
                )
        except Exception as e:
            print(f"[{SharePointUtils.ts()}]   [CLEANUP] Error clearing SharePoint folder {folder_url}: {e}")

    async def upload_ndc_reports(
        self, client: httpx.AsyncClient, site_id: str, drive_id: str, target_base: str
    ):
        """Uploads the latest NDC Excel report to SharePoint."""
        try:
            if not self.config.ndc_reports_dir.exists():
                return

            try:
                latest_file = SharePointUtils.get_latest_report(self.config.ndc_reports_dir)
            except FileNotFoundError:
                print(f"[{SharePointUtils.ts()}] No NDC report files found.")
                return

            SharePointUtils.cleanup_old_ndc_reports(latest_file, self.config.ndc_reports_dir)

            target_folder_url = f"{target_base}/NDC_Reports"
            await self.cleanup_sharepoint_folder_files(client, site_id, drive_id, target_folder_url)

            print(f"[{SharePointUtils.ts()}] NDC report: {latest_file.name}")
            try:
                await self.upload_file(client, site_id, drive_id, target_folder_url, latest_file)
                print(f"[{SharePointUtils.ts()}] Uploaded successfully: {latest_file.name}")
            except Exception as e:
                print(f"[{SharePointUtils.ts()}] [ERROR] Upload failed: {e}\n{traceback.format_exc()}")
        except Exception as e:
            print(f"[{SharePointUtils.ts()}] [ERROR] NDC Upload Workflow Error: {e}")

    async def upload_ff_reports(
        self, client: httpx.AsyncClient, site_id: str, drive_id: str, target_base: str
    ):
        """Uploads all downloaded F&F documents to SharePoint."""
        try:
            if not self.config.ff_reports_dir.exists():
                return

            subdirs = [d for d in self.config.ff_reports_dir.iterdir() if d.is_dir()]
            if not subdirs:
                print(f"[{SharePointUtils.ts()}] No F&F employee folders found.")
                return

            target_base_url = f"{target_base}/F&F_Documents"
            print(f"[{SharePointUtils.ts()}] F&F documents: {len(subdirs)} employee(s)")

            for subdir in subdirs:
                emp_id = subdir.name
                files = [f for f in subdir.iterdir() if f.is_file()]
                if not files:
                    continue

                emp_target_url = f"{target_base_url}/{emp_id}"

                await self.cleanup_sharepoint_folder_files(client, site_id, drive_id, emp_target_url)

                for f in files:
                    try:
                        await self.upload_file(client, site_id, drive_id, emp_target_url, f)
                        print(f"[{SharePointUtils.ts()}] Uploaded successfully: {f.name} for employee {emp_id}")
                        try:
                            if f.exists():
                                f.unlink()
                                print(f"[{SharePointUtils.ts()}]   [CLEANUP] Deleted local file: {f.name}")
                        except Exception as del_err:
                            print(f"[{SharePointUtils.ts()}]   [WARN] Could not delete local file {f.name}: {del_err}")
                    except Exception as e:
                        print(
                            f"[{SharePointUtils.ts()}] [ERROR] Failed to upload {f.name} for employee {emp_id}: {e}\n{traceback.format_exc()}"
                        )

                # Clean up empty employee directory
                try:
                    if subdir.exists():
                        remaining = list(subdir.iterdir())
                        if not remaining:
                            subdir.rmdir()
                            print(f"[{SharePointUtils.ts()}]   [CLEANUP] Removed empty folder: {subdir.name}")
                except Exception:
                    pass
        except Exception as e:
            print(f"[{SharePointUtils.ts()}] [ERROR] F&F Upload Workflow Error: {e}")


# Main CLI Runner
class SharePointCLI:
    """Manages command-line argument parsing and execution for manual SharePoint uploads."""

    @staticmethod
    async def run():
        """Parses CLI arguments and coordinates the requested upload job."""
        parser = argparse.ArgumentParser(description="Upload report files to SharePoint.")
        parser.add_argument(
            "--type",
            choices=["ndc", "ff", "all"],
            default="all",
            help="Type of reports to upload: 'ndc' (Excel files), 'ff' (Employee F&F folders), or 'all' (default).",
        )
        parser.add_argument(
            "--file",
            type=str,
            help="Path to a specific file to upload. If provided, uploads directly to the main target folder.",
        )
        args = parser.parse_args()

        service = SharePointService()

        try:
            async with httpx.AsyncClient(verify=False) as client:
                print(f"[{SharePointUtils.ts()}] Connecting to SharePoint...")
                site_id = await service.get_site_id(client)
                drive_id, folder_path = await service.get_drive_details(client, site_id)
                print(f"[{SharePointUtils.ts()}] Connected successfully.")
                print(f"[{SharePointUtils.ts()}] Target folder: {service.config.target_folder}")

                if args.file:
                    local_file = Path(args.file)
                    if not local_file.exists():
                        raise FileNotFoundError(f"Specified file not found: {args.file}")
                    print(f"[{SharePointUtils.ts()}] Uploading single file: {local_file.name} to {folder_path}")
                    await service.upload_file(client, site_id, drive_id, folder_path, local_file)
                    print(f"[{SharePointUtils.ts()}] Upload Completed successfully.")
                else:
                    if args.type in ["ndc", "all"]:
                        print(f"[{SharePointUtils.ts()}] Starting NDC reports upload...")
                        await service.upload_ndc_reports(client, site_id, drive_id, folder_path)

                    if args.type in ["ff", "all"]:
                        print(f"[{SharePointUtils.ts()}] Starting F&F reports upload...")
                        await service.upload_ff_reports(client, site_id, drive_id, folder_path)

        except ValueError as e:
            print(f"\n[{SharePointUtils.ts()}] [ERROR] Configuration Error: {traceback.format_exc()}")
        except FileNotFoundError as e:
            print(f"\n[{SharePointUtils.ts()}] [ERROR] File Not Found: {traceback.format_exc()}")
        except PermissionError as e:
            print(f"\n[{SharePointUtils.ts()}] [ERROR] Local Permission Error: {traceback.format_exc()}")
        except Exception as e:
            print(f"\n[{SharePointUtils.ts()}] [ERROR] Unexpected Error: {traceback.format_exc()}")


if __name__ == "__main__":
    asyncio.run(SharePointCLI.run())
