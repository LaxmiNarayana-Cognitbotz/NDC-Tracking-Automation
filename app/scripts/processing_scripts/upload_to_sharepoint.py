"""
SharePoint File Upload Script using Microsoft Graph API
"""

import argparse
import asyncio
import os
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

from app.core.config import (
    FF_REPORTS_DIR,
    NDC_REPORTS_DIR,
    PREJOINING_REPORTS_DIR,
    UPLOADS_DIR,
)
from app.core.logger import setup_logger

load_dotenv(verbose=True)

# SharePoint Configurations
TENANT_ID = os.getenv("SHAREPOINT_TENANT_ID")
SITE_URL = os.getenv("SHAREPOINT_SITE_URL")
CLIENT_ID = os.getenv("SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = os.getenv("SHAREPOINT_CLIENT_SECRET")
TARGET_FOLDER = os.getenv("SHAREPOINT_TARGET_FOLDER")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}" if TENANT_ID else ""
SCOPES = ["https://graph.microsoft.com/.default"]

# Cache to avoid repeating site/drive lookups
_SITE_ID: Optional[str] = None
_DRIVE_ID: Optional[str] = None


def ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


async def get_access_token() -> str:
    """Obtain a Microsoft Graph API access token using client credentials flow."""
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    
    session = requests.Session()
    session.verify = False

    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
        http_client=session
    )
    
    # Try getting token silently from cache
    result = app.acquire_token_silent(SCOPES, account=None)
    if not result:
        result = await asyncio.to_thread(
            app.acquire_token_for_client,
            scopes=SCOPES
        )
        
    if "access_token" in result:
        return result["access_token"]
    else:
        error_msg = result.get("error_description") or result.get("error") or "Unknown error"
        raise Exception(f"SharePoint authentication failed: {error_msg}")


async def get_site_id(client: httpx.AsyncClient) -> str:
    """Resolve the SharePoint Site ID from the SITE_URL."""
    global _SITE_ID
    if _SITE_ID:
        return _SITE_ID

    parsed_url = urllib.parse.urlparse(SITE_URL)
    hostname = parsed_url.netloc
    relative_path = parsed_url.path.rstrip("/")
    
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    
    url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{relative_path}"
    response = await client.get(url, headers=headers)
    if response.status_code != 200:
        raise Exception(f"Failed to resolve site ID: {response.status_code} - {response.text}")
        
    site_data = response.json()
    _SITE_ID = site_data.get("id")
    if not _SITE_ID:
        raise Exception("Site details resolved but 'id' field is missing.")
    return _SITE_ID


async def get_drive_details(client: httpx.AsyncClient, site_id: str, target_folder: Optional[str] = None) -> Tuple[str, str]:
    """Resolves the Drive ID and the base folder path inside that drive."""
    global _DRIVE_ID

    # Get site path (e.g. /sites/AGEL-Automation)
    parsed_site = urllib.parse.urlparse(SITE_URL)
    site_path = parsed_site.path.rstrip("/")
    
    # Get normalized target folder relative to site path
    folder = target_folder or TARGET_FOLDER or os.getenv("SHAREPOINT_TARGET_FOLDER") or ""
    if folder.startswith(site_path):
        folder = folder[len(site_path):]
    folder = folder.strip("/")
    
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    
    # Fetch all drives in the site
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    response = await client.get(url, headers=headers)
    if response.status_code != 200:
        raise Exception(f"Failed to list drives: {response.status_code}")
        
    drives = response.json().get("value", [])
    
    drive_id = None
    folder_path_in_drive = ""
    
    # Match target folder against each drive
    for d in drives:
        # Get the relative path of the drive under the site
        drive_web_url = d.get("webUrl", "")
        parsed_drive = urllib.parse.urlparse(drive_web_url)
        drive_path = parsed_drive.path.rstrip("/")
        
        # Relative path of the drive from the site root (e.g. "Shared Documents")
        drive_rel_path = ""
        if drive_path.startswith(site_path):
            drive_rel_path = drive_path[len(site_path):].strip("/")
            
        drive_name = d.get("name", "")
        
        # Candidates for the drive matching
        candidates = []
        if drive_rel_path:
            candidates.append(drive_rel_path)
            candidates.append(urllib.parse.unquote(drive_rel_path))
            candidates.append(urllib.parse.quote(drive_rel_path))
        if drive_name:
            candidates.append(drive_name)
            
        # Check if the folder path starts with any of the candidates
        for candidate in candidates:
            candidate_clean = candidate.strip("/")
            if not candidate_clean:
                continue
                
            # Exact match or prefix match followed by /
            if folder == candidate_clean:
                drive_id = d.get("id")
                folder_path_in_drive = ""
                break
            elif folder.startswith(candidate_clean + "/"):
                drive_id = d.get("id")
                folder_path_in_drive = folder[len(candidate_clean):].strip("/")
                break
                
        if drive_id:
            break

    # If no matching drive was found, fallback to the site's default drive
    if not drive_id:
        default_drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
        res = await client.get(default_drive_url, headers=headers)
        if res.status_code == 200:
            drive_id = res.json().get("id")
            folder_path_in_drive = folder
        else:
            raise Exception(f"Failed to resolve site default drive: {res.text}")
            
    _DRIVE_ID = drive_id
    return drive_id, folder_path_in_drive


def encode_path_segments(path: str) -> str:
    """Safely URL-encodes each path segment individually, preserving slashes."""
    segments = [urllib.parse.quote(s) for s in path.split("/") if s]
    return "/".join(segments)


async def upload_file(client: httpx.AsyncClient, site_id: str, drive_id: str, target_folder_url: str, local_file_path: Path):
    """Upload a file to a specific folder on SharePoint via Graph API."""
    file_name = local_file_path.name
    file_size = local_file_path.stat().st_size
    
    # Determine the target path for the file
    clean_folder = "/".join([p for p in target_folder_url.split("/") if p])
    target_file_path = f"{clean_folder}/{file_name}"
    encoded_path = encode_path_segments(target_file_path)
    
    token = await get_access_token()
    
    # Read the file content
    with open(local_file_path, "rb") as f:
        file_data = f.read()

    if file_size <= 4 * 1024 * 1024:  # 4MB
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}:/content"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream"
        }
        response = await client.put(url, headers=headers, content=file_data)
        if response.status_code not in [200, 201]:
            raise Exception(f"Upload failed (PUT): {response.status_code} - {response.text}")
    else:
        # Create upload session for larger files (explicitly replace to avoid _1 duplicates)
        session_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}:/createUploadSession"
        session_payload = {
            "item": {
                "@microsoft.graph.conflictBehavior": "replace"
            }
        }
        session_response = await client.post(
            session_url,
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=session_payload,
        )
        if session_response.status_code != 200:
            raise Exception(f"Failed to create upload session: {session_response.status_code} - {session_response.text}")
        
        upload_url = session_response.json().get("uploadUrl")
        chunk_size = 3276800  # Must be multiple of 327,680 bytes
        for i in range(0, file_size, chunk_size):
            chunk = file_data[i:i+chunk_size]
            content_range = f"bytes {i}-{i+len(chunk)-1}/{file_size}"
            chunk_headers = {
                "Content-Length": str(len(chunk)),
                "Content-Range": content_range
            }
            chunk_res = await client.put(upload_url, headers=chunk_headers, content=chunk)
            if chunk_res.status_code not in [200, 201, 202]:
                raise Exception(f"Chunk upload failed (Session): {chunk_res.status_code} - {chunk_res.text}")


async def cleanup_sharepoint_folder_files(client: httpx.AsyncClient, site_id: str, drive_id: str, folder_url: str):
    """
    Delete all existing files inside a specific SharePoint folder before uploading new ones.
    Prevents duplicate files (_1, _2 suffixes) or retaining outdated documents.
    """
    try:
        token = await get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        encoded_path = encode_path_segments(folder_url)

        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}:/children"
        response = await client.get(url, headers=headers)

        if response.status_code == 200:
            children = response.json().get("value", [])
            files = [item for item in children if "file" in item]
            deleted_count = 0
            for f in files:
                item_id = f.get("id")
                file_name = f.get("name", "unknown")
                delete_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{item_id}"
                del_res = await client.delete(delete_url, headers=headers)
                if del_res.status_code in [200, 204]:
                    deleted_count += 1
                    print(f"[{ts()}]   [CLEANUP] Deleted existing SharePoint file: {file_name}")
                else:
                    print(f"[{ts()}]   [CLEANUP] Could not delete SharePoint file {file_name}: {del_res.text}")
            if deleted_count > 0:
                print(f"[{ts()}]   [CLEANUP] Cleared {deleted_count} existing file(s) from SharePoint folder: {folder_url}")
        elif response.status_code == 404:
            # Folder does not exist yet on SharePoint -- will be created on upload
            pass
        else:
            print(f"[{ts()}]   [CLEANUP] Notice on folder query ({response.status_code}): {response.text}")
    except Exception as e:
        print(f"[{ts()}]   [CLEANUP] Error clearing SharePoint folder {folder_url}: {e}")


def get_latest_report(directory: Path, pattern: str = "*.xls") -> Path:
    """Find the most recently modified file matching the pattern in the directory."""
    if not directory.exists():
        raise FileNotFoundError(f"Local reports directory '{directory}' does not exist.")

    files = list(directory.glob(pattern))
    if not files:
        files = list(directory.glob("*.xlsx"))
        if not files:
            raise FileNotFoundError(f"No files matching '{pattern}' or '*.xlsx' found in '{directory}'.")

    files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    return files[0]


def cleanup_old_ndc_reports(latest_file: Path):
    """Delete all files in uploads/NDC_Reports except for the latest_file."""
    try:
        latest_str = str(latest_file.resolve()).lower()
        for f in NDC_REPORTS_DIR.iterdir():
            if f.is_file() and str(f.resolve()).lower() != latest_str:
                try:
                    f.unlink()
                    print(f"[{ts()}] [CLEANUP] Deleted local file: {f.name}")
                except Exception as e:
                    print(f"[{ts()}] [CLEANUP] Warning: could not delete local file {f.name}: {e}")
    except Exception as e:
        print(f"[{ts()}] [CLEANUP] Local cleanup error: {traceback.format_exc()}")


async def upload_ndc_reports(client: httpx.AsyncClient, site_id: str, drive_id: str, target_base: str):
    """Upload only the latest Excel file from local uploads/NDC_Reports to SharePoint."""
    if not NDC_REPORTS_DIR.exists():
        return

    try:
        latest_file = get_latest_report(NDC_REPORTS_DIR)
    except FileNotFoundError:
        print(f"[{ts()}] No NDC report files found.")
        return

    cleanup_old_ndc_reports(latest_file)

    target_folder_url = f"{target_base}/NDC_Reports"
    await cleanup_sharepoint_folder_files(client, site_id, drive_id, target_folder_url)

    print(f"[{ts()}] NDC report: {latest_file.name}")
    try:
        await upload_file(client, site_id, drive_id, target_folder_url, latest_file)
        print(f"[{ts()}] Uploaded successfully: {latest_file.name}")
    except Exception as e:
        print(f"[{ts()}] [ERROR] Upload failed: {e}\n{traceback.format_exc()}")


async def upload_ff_reports(client: httpx.AsyncClient, site_id: str, drive_id: str, target_base: str):
    """Upload all F&F documents from local uploads/FF_Reports to SharePoint."""
    if not FF_REPORTS_DIR.exists():
        return

    subdirs = [d for d in FF_REPORTS_DIR.iterdir() if d.is_dir()]
    if not subdirs:
        print(f"[{ts()}] No F&F employee folders found.")
        return

    target_base_url = f"{target_base}/F&F_Documents"
    print(f"[{ts()}] F&F documents: {len(subdirs)} employee(s)")

    for subdir in subdirs:
        emp_id = subdir.name
        files = [f for f in subdir.iterdir() if f.is_file()]
        if not files:
            continue

        emp_target_url = f"{target_base_url}/{emp_id}"

        # ==== Delete all existing files in this employee SharePoint folder first ====
        await cleanup_sharepoint_folder_files(client, site_id, drive_id, emp_target_url)

        # ==== Upload all documents for this employee from local system as-is ====
        for f in files:
            try:
                await upload_file(client, site_id, drive_id, emp_target_url, f)
                print(f"[{ts()}] Uploaded successfully: {f.name} for employee {emp_id}")
            except Exception as e:
                print(f"[{ts()}] [ERROR] Failed to upload {f.name} for employee {emp_id}: {e}\n{traceback.format_exc()}")


async def upload_prejoining_reports(client: httpx.AsyncClient, site_id: str, drive_id: str, target_base: str):
    """Upload Pre-Joining summary and employee document folders to SharePoint."""
    if not PREJOINING_REPORTS_DIR.exists():
        return

    target_base_url = f"{target_base}/Pre_Joining_Reports"

    # Upload Master Summary Excel if present
    master_file = PREJOINING_REPORTS_DIR / "Master_Pre_Joining_Summary.xlsx"
    if master_file.exists():
        print(f"[{ts()}] Uploading Master Pre-Joining Summary to SharePoint...")
        try:
            await upload_file(client, site_id, drive_id, target_base_url, master_file)
            print(f"[{ts()}] Uploaded successfully: {master_file.name}")
        except Exception as e:
            print(f"[{ts()}] [ERROR] Failed to upload {master_file.name}: {e}")

    # Upload employee document folders
    subdirs = [d for d in PREJOINING_REPORTS_DIR.iterdir() if d.is_dir()]
    for subdir in subdirs:
        emp_id = subdir.name
        files = [f for f in subdir.iterdir() if f.is_file()]
        if not files:
            continue
        emp_target_url = f"{target_base_url}/{emp_id}"
        await cleanup_sharepoint_folder_files(client, site_id, drive_id, emp_target_url)
        for f in files:
            try:
                await upload_file(client, site_id, drive_id, emp_target_url, f)
                print(f"[{ts()}] Uploaded successfully: {f.name} for employee {emp_id}")
            except Exception as e:
                print(f"[{ts()}] [ERROR] Failed to upload {f.name} for employee {emp_id}: {e}")


async def main():
    parser = argparse.ArgumentParser(description="Upload report files to SharePoint.")
    parser.add_argument(
        "--type",
        choices=["ndc", "ff", "prejoining", "all"],
        default="all",
        help="Type of reports to upload: 'ndc' (Excel files), 'ff' (Employee F&F folders), 'prejoining' (Pre-Joining reports), or 'all' (default).",
    )
    parser.add_argument(
        "--file",
        type=str,
        help="Path to a specific file to upload. If provided, uploads directly to the main target folder.",
    )
    args = parser.parse_args()

    try:
        async with httpx.AsyncClient(verify=False) as client:
            print(f"[{ts()}] Connecting to SharePoint...")
            site_id = await get_site_id(client)
            drive_id, folder_path = await get_drive_details(client, site_id)
            print(f"[{ts()}] Connected successfully.")
            print(f"[{ts()}] Target folder: {TARGET_FOLDER}")

            if args.file:
                local_file = Path(args.file)
                if not local_file.exists():
                    raise FileNotFoundError(f"Specified file not found: {args.file}")
                print(f"[{ts()}] Uploading single file: {local_file.name} to {folder_path}")
                await upload_file(client, site_id, drive_id, folder_path, local_file)
                print(f"[{ts()}] Upload Completed successfully.")
            else:
                if args.type in ["ndc", "all"]:
                    print(f"[{ts()}] Starting NDC reports upload...")
                    await upload_ndc_reports(client, site_id, drive_id, folder_path)

                if args.type in ["ff", "all"]:
                    print(f"[{ts()}] Starting F&F reports upload...")
                    await upload_ff_reports(client, site_id, drive_id, folder_path)

                if args.type in ["prejoining", "all"]:
                    print(f"[{ts()}] Starting Pre-Joining reports upload...")
                    await upload_prejoining_reports(client, site_id, drive_id, folder_path)

    except ValueError as e:
        print(f"\n[{ts()}] [ERROR] Configuration Error: {traceback.format_exc()}")
    except FileNotFoundError as e:
        print(f"\n[{ts()}] [ERROR] File Not Found: {traceback.format_exc()}")
    except PermissionError as e:
        print(f"\n[{ts()}] [ERROR] Local Permission Error: {traceback.format_exc()}")
    except Exception as e:
        print(f"\n[{ts()}] [ERROR] Unexpected Error: {traceback.format_exc()}")


if __name__ == "__main__":
    setup_logger("upload_to_sharepoint.log")
    asyncio.run(main())
