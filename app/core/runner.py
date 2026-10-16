"""
Background Automation Task Runner & State Registry.
"""

import asyncio
import traceback
import uuid
from datetime import datetime
from typing import Any, Callable, Coroutine, Dict, Optional

# In-memory registry for tracking background job states
JOB_REGISTRY: Dict[str, Dict[str, Any]] = {}


def get_job(job_id: str) -> Optional[Dict[str, Any]]:
    """Retrieve details of a job by ID."""
    return JOB_REGISTRY.get(job_id)


def list_jobs() -> Dict[str, Dict[str, Any]]:
    """Retrieve all tracked automation jobs."""
    return JOB_REGISTRY


def get_active_job_count() -> int:
    """Return count of currently running background jobs."""
    return sum(1 for j in JOB_REGISTRY.values() if j.get("status") == "running")


async def _execute_job_wrapper(job_id: str, coroutine_func: Callable[[], Coroutine[Any, Any, Any]]):
    """Executes a coroutine in background and records duration, result, or error."""
    job = JOB_REGISTRY.get(job_id)
    if not job:
        return

    job["status"] = "running"
    start_time = datetime.now()
    job["started_at"] = start_time.isoformat()

    try:
        result = await coroutine_func()
        job["status"] = "completed"
        job["result"] = str(result) if result is not None else "Execution completed successfully."
    except Exception as exc:
        job["status"] = "failed"
        job["error"] = f"{type(exc).__name__}: {exc}"
        job["traceback"] = traceback.format_exc()
    finally:
        end_time = datetime.now()
        job["completed_at"] = end_time.isoformat()
        job["duration_seconds"] = round((end_time - start_time).total_seconds(), 2)


def submit_automation_job(pipeline_name: str, coroutine_func: Callable[[], Coroutine[Any, Any, Any]]) -> Dict[str, Any]:
    """Submits an async automation task to run in the background."""
    try:
        job_id = str(uuid.uuid4())
        now_str = datetime.now().isoformat()

        job_record = {
            "job_id": job_id,
            "pipeline": pipeline_name,
            "status": "queued",
            "created_at": now_str,
            "started_at": None,
            "completed_at": None,
            "duration_seconds": None,
            "result": None,
            "error": None,
            "traceback": None,
        }
        JOB_REGISTRY[job_id] = job_record

        asyncio.create_task(_execute_job_wrapper(job_id, coroutine_func))

        return {
            "job_id": job_id,
            "pipeline": pipeline_name,
            "status": "queued",
            "message": f"Pipeline '{pipeline_name}' started successfully in background.",
            "created_at": now_str,
        }
    except Exception as exc:
        raise RuntimeError(f"Failed to submit job '{pipeline_name}': {exc}") from exc
