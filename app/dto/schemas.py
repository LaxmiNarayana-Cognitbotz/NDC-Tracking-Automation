from enum import Enum
from typing import Any, Dict, List, Optional
from pydantic import BaseModel, Field


class PipelineTarget(str, Enum):
    ALL = "all"
    NDC = "ndc"
    FF = "ff"
    PREJOINING = "prejoining"


# ==== Request DTOs ====
class EmployeeDownloadRequest(BaseModel):
    employee_numbers: List[str] = Field(
        ...,
        description="List of employee IDs to download documents for",
        examples=[["10002440", "10002394"]],
    )


class SharePointUploadRequest(BaseModel):
    target_pipeline: PipelineTarget = Field(
        default=PipelineTarget.ALL,
        description="Target pipeline to upload ('all', 'ndc', 'ff', 'prejoining')",
        examples=[PipelineTarget.ALL],
    )


# ==== Response DTOs ====
class HealthCheckResponse(BaseModel):
    status: str = "healthy"
    timestamp: str


class AutomationJobResponse(BaseModel):
    job_id: str
    pipeline: str
    status: str = "queued"
    message: str
    created_at: str


class AutomationJobStatus(BaseModel):
    job_id: str
    pipeline: str
    status: str
    started_at: Optional[str] = None
    completed_at: Optional[str] = None
    duration_seconds: Optional[float] = None
    result: Optional[Any] = None
    error: Optional[str] = None
