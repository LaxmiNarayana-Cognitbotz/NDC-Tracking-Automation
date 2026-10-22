"""Background Automation Task Runner and State Registry."""

import asyncio
import logging
import sys
import traceback
import uuid
from datetime import datetime
from typing import Any, Callable, Coroutine, Dict, Optional

logger = logging.getLogger(__name__)

# In-memory registry for tracking background job states
JOB_REGISTRY: Dict[str, Dict[str, Any]] = {}


def get_job(job_id: str) -> Optional[Dict[str, Any]]:
    """Retrieve details of a job by ID."""
    try:
        return JOB_REGISTRY.get(job_id)
    except Exception as e:
        logger.error(f"Error getting job {job_id}: {e}")
        return None


def list_jobs() -> Dict[str, Dict[str, Any]]:
    """Retrieve all tracked automation jobs."""
    try:
        return JOB_REGISTRY
    except Exception as e:
        logger.error(f"Error listing jobs: {e}")
        return {}


def get_active_job_count() -> int:
    """Return count of currently running background jobs."""
    try:
        return sum(1 for j in JOB_REGISTRY.values() if j.get("status") == "running")
    except Exception as e:
        logger.error(f"Error counting active jobs: {e}")
        return 0


def _run_in_worker_proactor(func_or_coro: Any) -> Any:
    """Run async task in worker thread with Proactor loop for Playwright on Windows."""
    try:
        if sys.platform == "win32":
            loop = asyncio.WindowsProactorEventLoopPolicy().new_event_loop()
        else:
            loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            if asyncio.iscoroutine(func_or_coro):
                coro = func_or_coro
            elif callable(func_or_coro):
                coro = func_or_coro()
            else:
                raise TypeError(f"Expected callable or coroutine, got {type(func_or_coro)}")
            return loop.run_until_complete(coro)
        finally:
            try:
                loop.close()
            except Exception:
                pass
    except Exception as e:
        logger.error(f"Worker proactor error: {e}")
        raise


async def _execute_job_wrapper(job_id: str, coroutine_func: Callable[[], Coroutine[Any, Any, Any]]):
    """Execute background automation task and record execution telemetry."""
    job = JOB_REGISTRY.get(job_id)
    if not job:
        return

    job["status"] = "running"
    start_time = datetime.now()
    job["started_at"] = start_time.isoformat()

    try:
        if sys.platform == "win32":
            result = await asyncio.to_thread(_run_in_worker_proactor, coroutine_func)
        else:
            result = await coroutine_func()
        job["status"] = "completed"
        job["result"] = str(result) if result is not None else "Execution completed successfully."
    except Exception as exc:
        job["status"] = "failed"
        job["error"] = f"{type(exc).__name__}: {exc}"
        job["traceback"] = traceback.format_exc()
        logger.error(f"Job {job_id} failed: {job['error']}")
    finally:
        end_time = datetime.now()
        job["completed_at"] = end_time.isoformat()
        job["duration_seconds"] = round((end_time - start_time).total_seconds(), 2)


def submit_automation_job(pipeline_name: str, coroutine_func: Callable[[], Coroutine[Any, Any, Any]]) -> Dict[str, Any]:
    """Submit an async automation task to run in the background."""
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
        logger.error(f"Failed to submit job '{pipeline_name}': {exc}")
        raise RuntimeError(f"Failed to submit job '{pipeline_name}': {exc}") from exc
