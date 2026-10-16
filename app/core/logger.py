"""
Logger writer utility redirecting automation outputs to dedicated log files inside logs/,
while managing server-mode consolidation and terminal output filtering.
"""

import io
import re
import sys
from pathlib import Path

from app.core.config import LOGS_DIR

from datetime import datetime

ANSI_ESCAPE_RE = re.compile(r"\x1b\[[0-9;?]*[a-zA-Z]|\x1b\([a-zA-Z]")


def enable_windows_vt_support():
    """Enables ANSI/VT100 escape code processing on Windows console handles."""
    if sys.platform == "win32":
        try:
            import ctypes

            kernel32 = ctypes.windll.kernel32
            for std_id in (-11, -12):  # STD_OUTPUT_HANDLE, STD_ERROR_HANDLE
                handle = kernel32.GetStdHandle(std_id)
                if handle and handle != -1:
                    mode = ctypes.c_ulong()
                    if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
                        kernel32.SetConsoleMode(handle, mode.value | 0x0004)  # ENABLE_VIRTUAL_TERMINAL_PROCESSING
        except Exception:
            pass


# Standard Uvicorn logging configuration with terminal colors enabled
UVICORN_LOG_CONFIG = {
    "version": 1,
    "disable_existing_loggers": False,
    "formatters": {
        "default": {
            "()": "uvicorn.logging.DefaultFormatter",
            "fmt": "%(levelprefix)s %(message)s",
            "use_colors": True,
        },
        "access": {
            "()": "uvicorn.logging.AccessFormatter",
            "fmt": '%(levelprefix)s %(client_addr)s - "%(request_line)s" %(status_code)s',
            "use_colors": True,
        },
    },
    "handlers": {
        "default": {
            "class": "logging.StreamHandler",
            "formatter": "default",
            "stream": "ext://sys.stderr",
        },
        "access": {
            "class": "logging.StreamHandler",
            "formatter": "access",
            "stream": "ext://sys.stdout",
        },
    },
    "loggers": {
        "uvicorn": {"handlers": ["default"], "level": "INFO", "propagate": False},
        "uvicorn.error": {"handlers": ["default"], "level": "INFO", "propagate": False},
        "uvicorn.access": {"handlers": ["access"], "level": "INFO", "propagate": False},
    },
}


class LoggerWriter:
    """Redirects stream to target log files while managing console output and timestamps."""

    _active_file: Path = LOGS_DIR / "server_logs.log"
    _is_server_mode: bool = False
    _last_was_allowed: bool = False

    def __init__(self, file_path: Path, stream=None):
        self.terminal = stream if stream is not None else getattr(sys, "__stdout__", sys.stdout)
        self.file_path = file_path
        self._ensure_file(file_path)

    @classmethod
    def _ensure_file(cls, path: Path):
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            if not path.exists():
                path.touch(exist_ok=True)
        except Exception:
            pass

    @classmethod
    def set_active_file(cls, file_path: Path, is_server: bool = False):
        """Sets the active target log file for subsequent writes."""
        if is_server:
            cls._is_server_mode = True
        elif cls._is_server_mode:
            # If server mode is active, keep server_logs.log as the consolidated log target
            return

        cls._ensure_file(file_path)
        cls._active_file = file_path

    @classmethod
    def is_server_mode(cls) -> bool:
        """Returns True if the process is running as a FastAPI server."""
        return cls._is_server_mode

    def write(self, message: str):
        if not message:
            return

        # Clean raw ANSI color codes for file logging
        cleaned = ANSI_ESCAPE_RE.sub("", message)
        stripped = cleaned.strip()

        # Format message for the log file: prepend timestamp if not already present
        file_msg = cleaned
        if stripped and not stripped.startswith("[20"):
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            file_msg = f"[{now_str}] {cleaned}"

        # 1. ALWAYS write timestamped clean text to the active log file
        target_file = self._active_file or self.file_path
        try:
            with open(target_file, "a", encoding="utf-8") as f:
                f.write(file_msg)
                f.flush()
        except Exception:
            pass

        # 2. Terminal echo behavior:
        # In standalone script mode, echo all logs to terminal with original colors
        if not self._is_server_mode:
            self._echo_raw(message)
            return

        if not stripped:
            if self._last_was_allowed and self.terminal is not None:
                self._echo_raw(message)
            return

        should_echo = self._is_server_message(stripped)
        self.__class__._last_was_allowed = should_echo

        if should_echo:
            # Echo normal print to terminal preserving original colors (without timestamp prefix)
            self._echo_raw(message)

    def _echo_raw(self, message: str):
        if self.terminal is not None:
            try:
                self.terminal.write(message)
                self.terminal.flush()
            except UnicodeEncodeError:
                try:
                    self.terminal.write(
                        message.encode("cp1252", errors="replace").decode("cp1252")
                    )
                    self.terminal.flush()
                except Exception:
                    pass
            except Exception:
                pass

    @staticmethod
    def _is_server_message(stripped: str) -> bool:
        """Determines if a message is an essential server message (Uvicorn, FastAPI, Errors)."""
        # Uvicorn structured level tags: "[INFO]", "[WARNING]", "[ERROR]", "[CRITICAL]", "[DEBUG]"
        if any(tag in stripped for tag in ("[INFO]", "[WARNING]", "[ERROR]", "[CRITICAL]", "[DEBUG]")):
            # Suppress granular automation steps from cluttering the terminal
            automation_indicators = (
                "Starting NDC", "Launching browser", "Navigating to", "Waiting for",
                "Home page ready", "Opening NDC", "Report parameters", "From Date set",
                "Selecting Business", "Applying report", "Report Completed", "Saved:",
                "Checking if report", "Could not locate", "SUCCESS! Report",
                "Pushing NDC", "Deleted existing", "Cleared 1", "NDC report:",
                "Uploaded successfully", "SharePoint sync", "[CLEANUP]", "Concurrency:",
                "tab(s) for", "Starting Pre-Joining", "Starting F&F"
            )
            if any(ind in stripped for ind in automation_indicators):
                return False
            return True

        if stripped.startswith(("INFO:", "WARNING:", "ERROR:", "CRITICAL:", "DEBUG:")):
            return True
        if "Starting server on" in stripped or "Uvicorn running on" in stripped:
            return True
        if "HTTP/1.1" in stripped or "HTTP/2" in stripped:
            return True
        if stripped.startswith("Traceback") or "Traceback (most recent call last)" in stripped:
            return True
        if stripped.startswith('File "') and ", line " in stripped:
            return True
        return False

    def flush(self):
        if self.terminal is not None:
            try:
                self.terminal.flush()
            except Exception:
                pass

    def isatty(self) -> bool:
        if self.terminal is not None and hasattr(self.terminal, "isatty"):
            try:
                return self.terminal.isatty()
            except Exception:
                return False
        return False

    @property
    def encoding(self) -> str:
        return getattr(self.terminal, "encoding", "utf-8")

    def fileno(self):
        if self.terminal is not None and hasattr(self.terminal, "fileno"):
            return self.terminal.fileno()
        raise io.UnsupportedOperation("fileno")


def setup_logger(log_filename: str) -> Path:
    """
    Configures stream redirection.
    - In standalone mode: redirects output to the specified individual log file (e.g. download_ndc_report.log).
    - In FastAPI server mode: routes all logs into server_logs.log.
    """
    enable_windows_vt_support()
    if LoggerWriter.is_server_mode():
        return LOGS_DIR / "server_logs.log"

    log_file = LOGS_DIR / log_filename
    LoggerWriter.set_active_file(log_file, is_server=False)
    if not isinstance(sys.stdout, LoggerWriter):
        sys.stdout = LoggerWriter(log_file, stream=getattr(sys, "__stdout__", sys.stdout))
    else:
        sys.stdout.file_path = log_file
    return log_file


def setup_server_logger(log_filename: str = "server_logs.log") -> Path:
    """Configures centralized server-level logging for the FastAPI application."""
    enable_windows_vt_support()
    log_file = LOGS_DIR / log_filename
    LoggerWriter.set_active_file(log_file, is_server=True)
    sys.stdout = LoggerWriter(log_file, stream=getattr(sys, "__stdout__", sys.stdout))
    sys.stderr = LoggerWriter(log_file, stream=getattr(sys, "__stderr__", sys.stderr))
    return log_file

