from typing import Dict, Optional
from fastapi import APIRouter, HTTPException, Query, status

from app.dto.schemas import (
    AutomationJobResponse,
    AutomationJobStatus,
    EmployeeDownloadRequest,
    PipelineTarget,
    SharePointUploadRequest,
)
from app.services.automation_service import AutomationService, JobService

router = APIRouter(prefix="/api/v1/automation", tags=["Automation Pipelines"])


@router.post("/ndc/download", response_model=AutomationJobResponse, status_code=status.HTTP_202_ACCEPTED)
async def trigger_ndc_download():
    """Trigger NDC Report download in the background."""
    try:
        return AutomationService.run_ndc_download()
    except Exception as exc:
        raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(exc)) from exc


@router.post("/fnf/process-closed", response_model=AutomationJobResponse, status_code=status.HTTP_202_ACCEPTED)
async def trigger_fnf_process():
    """Process active F&F reports directly from SharePoint."""
    try:
        return AutomationService.run_fnf_process()
    except Exception as exc:
        raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(exc)) from exc


@router.post("/fnf/download-by-ids", response_model=AutomationJobResponse, status_code=status.HTTP_202_ACCEPTED)
async def trigger_fnf_download_by_ids(request: EmployeeDownloadRequest):
    """Download F&F documents for provided employee IDs."""
    try:
        return AutomationService.run_fnf_download_by_ids(request)
    except ValueError as val_err:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(val_err)) from val_err
    except Exception as exc:
        raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(exc)) from exc


@router.post("/elcm/download", response_model=AutomationJobResponse, status_code=status.HTTP_202_ACCEPTED)
async def trigger_elcm_download():
    """Trigger ELCM New Joinees report download."""
    try:
        return AutomationService.run_elcm_download()
    except Exception as exc:
        raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(exc)) from exc


@router.post("/prejoining/download", response_model=AutomationJobResponse, status_code=status.HTTP_202_ACCEPTED)
async def trigger_prejoining_download(request: EmployeeDownloadRequest):
    """Download Pre-Joining documents (Offer, BGV, Appointment) for provided employee IDs."""
    try:
        return AutomationService.run_prejoining_download(request)
    except ValueError as val_err:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(val_err)) from val_err
    except Exception as exc:
        raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(exc)) from exc


@router.post("/sharepoint/upload", response_model=AutomationJobResponse, status_code=status.HTTP_202_ACCEPTED)
async def trigger_sharepoint_upload(
    target_pipeline: Optional[PipelineTarget] = Query(
        default=PipelineTarget.ALL,
        description="Target pipeline to upload to SharePoint ('all', 'ndc', 'ff', 'prejoining')",
    ),
    request: Optional[SharePointUploadRequest] = None,
):
    """Trigger SharePoint upload for downloaded reports."""
    try:
        # If request body specifies a pipeline other than ALL, use it; otherwise use the query parameter
        selected_pipeline = target_pipeline or PipelineTarget.ALL
        if request and request.target_pipeline:
            selected_pipeline = request.target_pipeline

        return AutomationService.run_sharepoint_upload(SharePointUploadRequest(target_pipeline=selected_pipeline))
    except Exception as exc:
        raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(exc)) from exc


@router.get("/jobs", response_model=Dict[str, AutomationJobStatus])
async def get_all_jobs():
    """Retrieve all tracked background automation jobs."""
    try:
        return JobService.list_all_jobs()
    except Exception as exc:
        raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(exc)) from exc


@router.get("/jobs/{job_id}", response_model=AutomationJobStatus)
async def get_job_status(job_id: str):
    """Retrieve execution status for a specific background job ID."""
    try:
        job = JobService.get_job_by_id(job_id)
        if not job:
            raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail=f"Job '{job_id}' not found.")
        return job
    except HTTPException:
        raise
    except ValueError as val_err:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(val_err)) from val_err
    except Exception as exc:
        raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(exc)) from exc
