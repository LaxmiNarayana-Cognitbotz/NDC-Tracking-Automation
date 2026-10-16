"""
FastAPI Application Entry Point & Uvicorn Server Launcher.
"""

import asyncio
import sys
import uvicorn
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from datetime import datetime
from app.core.logger import setup_server_logger, UVICORN_LOG_CONFIG

# Configure centralized logging for FastAPI server
setup_server_logger("server_logs.log")

from app.api.routes import router as automation_router
from app.dto.schemas import HealthCheckResponse
from app.services.automation_service import HealthService

load_dotenv(verbose=True)

# Define explicit Proactor loop factory for Windows (required by Playwright when running under Uvicorn)
def proactor_loop_factory(use_subprocess: bool = False):
    """Returns a ProactorEventLoop instance to support Playwright subprocesses on Windows."""
    if sys.platform == "win32":
        return asyncio.ProactorEventLoop()
    return asyncio.new_event_loop()


app = FastAPI(
    title="NDC & HR Operations Tracking Automation API",
    description="Enterprise Automation REST API for NDC, F&F, ELCM, and Pre-Joining Report Workflows.",
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(automation_router)


# ==== Health & Root Endpoints ====

@app.get("/health", response_model=HealthCheckResponse, tags=["Health"])
async def health_check():
    """Verify application health and return current timestamp."""
    try:
        return HealthService.get_health_status()
    except Exception as exc:
        raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(exc)) from exc


@app.get("/", tags=["Root"])
async def root_index():
    """Root welcome endpoint providing links to documentation and health check."""
    return JSONResponse(
        content={
            "message": "NDC & HR Operations Tracking Automation API is online.",
            "health": "/health",
            "docs": "/docs",
            "version": "1.0.0",
        }
    )


if __name__ == "__main__":
    print("Starting server on http://0.0.0.0:8000 (docs: http://localhost:8000/docs)")
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
        loop="main:proactor_loop_factory",
        log_config=UVICORN_LOG_CONFIG,
        use_colors=True,
    )
