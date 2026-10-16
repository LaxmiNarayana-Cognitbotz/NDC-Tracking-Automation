"""
Business Service Layer orchestrating automation pipelines and job tracking.
"""

import importlib
from datetime import datetime
from typing import Dict, Optional

from app.core.runner import get_active_job_count, get_job, list_jobs, submit_automation_job
from app.core.config import ROOT_DIR
from app.dto.schemas import (
    AutomationJobResponse,
    AutomationJobStatus,
    EmployeeDownloadRequest,
    HealthCheckResponse,
    SharePointUploadRequest,
)

# Dynamic import for module containing ampersand
ff_module = importlib.import_module("app.scripts.automation_scripts.download_f&f_report")
download_ff_reports = ff_module.download_ff_reports

import app.scripts.automation_scripts.download_ndc_report as ndc_script
import app.scripts.automation_scripts.download_prejoining_report as prejoining_script
import app.scripts.automation_scripts.elcm_report_downloader as elcm_script
import app.scripts.processing_scripts.process_fnf_closed_reports as fnf_process_script
import app.scripts.processing_scripts.upload_to_sharepoint as sp_uploader


# ==== Health Service ====

class HealthService:
    """Provides operational health metrics."""

    @staticmethod
    def get_health_status() -> HealthCheckResponse:
        """Returns simple health status and current timestamp."""
        try:
            return HealthCheckResponse(
                status="healthy",
                timestamp=datetime.now().isoformat(),
            )
        except Exception as exc:
            raise RuntimeError(f"Health check failed: {exc}") from exc


# ==== Automation Pipeline Service ====

class AutomationService:
    """Encapsulates execution workflows for all RPA automation pipelines."""

    @staticmethod
    def run_ndc_download() -> AutomationJobResponse:
        """Triggers Oracle Fusion NDC report download and SharePoint upload."""
        try:
            job = submit_automation_job("ndc_report_download", ndc_script.download_ndc_report)
            return AutomationJobResponse(**job)
        except Exception as exc:
            raise RuntimeError(f"Failed to submit NDC download: {exc}") from exc

    @staticmethod
    def run_fnf_process() -> AutomationJobResponse:
        """Processes active F&F records from SharePoint and downloads matching files."""
        try:
            job = submit_automation_job("fnf_process_closed_reports", fnf_process_script.process_fnf_active_reports)
            return AutomationJobResponse(**job)
        except Exception as exc:
            raise RuntimeError(f"Failed to submit F&F process: {exc}") from exc

    @staticmethod
    def run_fnf_download_by_ids(request: EmployeeDownloadRequest) -> AutomationJobResponse:
        """Downloads F&F documents for specific employee numbers."""
        clean_ids = [str(e).strip() for e in request.employee_numbers if str(e).strip()]
        if not clean_ids:
            raise ValueError("No valid employee numbers provided.")

        try:
            job = submit_automation_job("fnf_download_by_ids", lambda: download_ff_reports(clean_ids))
            return AutomationJobResponse(**job)
        except Exception as exc:
            raise RuntimeError(f"Failed to submit F&F download: {exc}") from exc

    @staticmethod
    def run_elcm_download() -> AutomationJobResponse:
        """Triggers ELCM New Joinees report download and SharePoint sync."""
        try:
            job = submit_automation_job("elcm_report_download", elcm_script.download_elcm_report)
            return AutomationJobResponse(**job)
        except Exception as exc:
            raise RuntimeError(f"Failed to submit ELCM download: {exc}") from exc

    @staticmethod
    def run_prejoining_download(request: EmployeeDownloadRequest) -> AutomationJobResponse:
        """Downloads Pre-Joining documents (Offer, BGV, Appointment) for employee numbers."""
        clean_ids = [str(e).strip() for e in request.employee_numbers if str(e).strip()]
        if not clean_ids:
            raise ValueError("No valid employee numbers provided.")

        try:
            job = submit_automation_job("prejoining_download", lambda: prejoining_script.download_prejoining_reports(clean_ids))
            return AutomationJobResponse(**job)
        except Exception as exc:
            raise RuntimeError(f"Failed to submit Pre-Joining download: {exc}") from exc

    @staticmethod
    def run_sharepoint_upload(request: SharePointUploadRequest) -> AutomationJobResponse:
        """Uploads local reports directly to SharePoint."""
        raw_target = request.target_pipeline.value if hasattr(request.target_pipeline, "value") else (request.target_pipeline or "all")
        target = str(raw_target).lower().strip()

        async def _upload():
            async with sp_uploader.httpx.AsyncClient(verify=False) as client:
                site_id = await sp_uploader.get_site_id(client)
                drive_id, folder_path = await sp_uploader.get_drive_details(client, site_id)

                if target in ["all", "ndc"]:
                    await sp_uploader.upload_ndc_reports(client, site_id, drive_id, folder_path)
                if target in ["all", "ff"]:
                    await sp_uploader.upload_ff_reports(client, site_id, drive_id, folder_path)
                if target in ["all", "prejoining"]:
                    await sp_uploader.upload_prejoining_reports(client, site_id, drive_id, folder_path)

        try:
            job = submit_automation_job(f"sharepoint_upload_{target}", _upload)
            return AutomationJobResponse(**job)
        except Exception as exc:
            raise RuntimeError(f"Failed to submit SharePoint upload: {exc}") from exc


# ==== Job Tracking Service ====

class JobService:
    """Manages querying and reporting background automation job states."""

    @staticmethod
    def list_all_jobs() -> Dict[str, AutomationJobStatus]:
        """Returns all tracked automation jobs."""
        try:
            return {jid: AutomationJobStatus(**info) for jid, info in list_jobs().items()}
        except Exception as exc:
            raise RuntimeError(f"Failed to retrieve jobs: {exc}") from exc

    @staticmethod
    def get_job_by_id(job_id: str) -> Optional[AutomationJobStatus]:
        """Finds a job by its unique UUID."""
        if not job_id or not job_id.strip():
            raise ValueError("Job ID cannot be empty.")

        try:
            info = get_job(job_id.strip())
            return AutomationJobStatus(**info) if info else None
        except Exception as exc:
            raise RuntimeError(f"Failed to query job '{job_id}': {exc}") from exc
