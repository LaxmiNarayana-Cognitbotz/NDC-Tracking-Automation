"""
Downloads F&F report from Sharepoint, runs F&F extraction, and uploads results.
"""

import sys
import os
from pathlib import Path

# Add project root to sys.path so the script can be run directly or via VS Code Run
ROOT_DIR = Path(__file__).resolve().parents[3]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import asyncio
import importlib
import traceback
from datetime import datetime

import httpx
import pandas as pd
from dotenv import load_dotenv

from app.core.config import UPLOADS_DIR
from app.core.logger import setup_logger
from app.scripts.processing_scripts.upload_to_sharepoint import (
    FF_REPORTS_DIR,
    encode_path_segments,
    get_access_token,
    get_drive_details,
    get_site_id,
    upload_file,
)

# Dynamically import the module since it contains an ampersand
download_module = importlib.import_module("app.scripts.automation_scripts.download_f&f_report")
download_ff_reports = download_module.download_ff_reports

# Load environment variables
load_dotenv(verbose=True)

# ==== Logging Setup ====
LOG_FILE = setup_logger("process_fnf_closed_reports.log")

LOCAL_EXCEL_DIR = UPLOADS_DIR / "FNF_Active_Reports"

def ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

async def process_fnf_active_reports():
    setup_logger("process_fnf_closed_reports.log")
    LOCAL_EXCEL_DIR.mkdir(parents=True, exist_ok=True)
    print(f"[{ts()}] Starting F&F Active Reports processing...")
    

    
    async with httpx.AsyncClient(timeout=60.0, verify=False) as client:
        try:
            print(f"[{ts()}] Connecting to SharePoint...")
            site_id = await get_site_id(client)
            target_sp = os.getenv("SHAREPOINT_TARGET_FOLDER")
            drive_id, folder_path = await get_drive_details(client, site_id, target_sp)
            token = await get_access_token()
            headers = {"Authorization": f"Bearer {token}"}
            
            # ==== List files in the F&F_Active_Report folder ====
            # The folder_path returned by get_drive_details is relative to the Drive root.
            fnf_folder_in_drive = f"{folder_path}/F&F_Active_Report".strip("/")
            encoded_path = encode_path_segments(fnf_folder_in_drive)
            
            url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}:/children"
            print(f"[{ts()}] Checking SharePoint directory: {fnf_folder_in_drive}")
            
            response = await client.get(url, headers=headers)
            if response.status_code != 200:
                print(f"[{ts()}] [ERROR] Failed to list SharePoint folder: {response.status_code} - {response.text}")
                return

            children = response.json().get("value", [])
            excel_files = []
            for item in children:
                if "file" in item:
                    name = item.get("name", "")
                    if name.lower().endswith(".xlsx"):
                        excel_files.append(item)
            
            print(f"[{ts()}] Found {len(excel_files)} Excel file(s) in '{fnf_folder_in_drive}'.")
            
            if not excel_files:
                return

            for file_item in excel_files:
                file_name = file_item.get("name")
                download_url = file_item.get("@microsoft.graph.downloadUrl")
                if not download_url:
                    print(f"[{ts()}] [ERROR] File '{file_name}' does not have a download URL.")
                    continue
                
                local_file_path = LOCAL_EXCEL_DIR / file_name
                print(f"[{ts()}] Downloading '{file_name}'...")
                
                dl_response = await client.get(download_url)
                if dl_response.status_code != 200:
                    print(f"[{ts()}] [ERROR] Failed to download '{file_name}' (status {dl_response.status_code}).")
                    continue
                
                with open(local_file_path, "wb") as f:
                    f.write(dl_response.content)
                
                print(f"[{ts()}] Successfully downloaded '{file_name}'. Extracting Person Numbers...")
                
                # ==== Read the Excel file and extract IDs ====
                try:
                    df = pd.read_excel(local_file_path)
                    # Support 'Person Number' or 'Employee ID' variations
                    person_col = None
                    for col in df.columns:
                        col_str = str(col).strip().lower()
                        if "person number" in col_str or "employee id" in col_str:
                            person_col = col
                            break
                    
                    if not person_col:
                        print(f"[{ts()}] [ERROR] Could not find 'Person Number' column in {file_name}. Columns found: {list(df.columns)}")
                        continue
                        
                    # Extract IDs, ensuring they are strings and removing NAs
                    employee_numbers = df[person_col].dropna().astype(str).str.replace(r'\.0$', '', regex=True).str.strip().tolist()
                    employee_numbers = [e for e in employee_numbers if e and e.lower() != "nan"]
                    
                    # Deduplicate
                    unique_employees = list(dict.fromkeys(employee_numbers))
                    print(f"[{ts()}] Found {len(unique_employees)} unique employee(s) in {file_name}: {unique_employees}")
                    
                    if not unique_employees:
                        print(f"[{ts()}] No valid employee IDs found in {file_name}.")
                        continue
                        
                    # ==== Define the upload callback for parallel upload ====
                    target_base_url = f"{folder_path}/F&F_Documents"
                    
                    async def on_employee_downloaded(emp_id: str, downloaded_paths: list):
                        emp_target_url = f"{target_base_url}/{emp_id}"
                        for file_path_str in downloaded_paths:
                            f_path = Path(file_path_str)
                            if f_path.exists():
                                try:
                                    await upload_file(client, site_id, drive_id, emp_target_url, f_path)
                                    print(f"[{ts()}] Uploaded successfully: {f_path.name} for employee {emp_id}")
                                    f_path.unlink()  # delete after successful upload
                                except Exception as e:
                                    print(f"[{ts()}] [ERROR] Failed to upload {f_path.name} for {emp_id}: {e}")
                                    
                        # Try to remove the employee folder if empty
                        emp_dir = FF_REPORTS_DIR / emp_id
                        if emp_dir.exists() and emp_dir.is_dir():
                            try:
                                emp_dir.rmdir()
                            except:
                                pass

                    # ==== Call the F&F download script with the callback ====
                    print(f"[{ts()}] Initiating F&F document download for {len(unique_employees)} employee(s)...")
                    await download_ff_reports(unique_employees, upload_callback=on_employee_downloaded)
                    
                    # ==== Clean up local Excel file ====
                    if local_file_path.exists():
                        local_file_path.unlink()
                        print(f"[{ts()}] [CLEANUP] Deleted local Excel file: {file_name}")
                                

                except Exception as e:
                    print(f"[{ts()}] [ERROR] Error processing '{file_name}': {e}\n{traceback.format_exc()}")
                
        except Exception as e:
            print(f"[{ts()}] [ERROR] SharePoint operation failed: {e}\n{traceback.format_exc()}")
            
    print(f"[{ts()}] Finished F&F Active Reports processing.")

if __name__ == "__main__":
    asyncio.run(process_fnf_active_reports())
