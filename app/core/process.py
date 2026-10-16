import subprocess
import sys


def suppress_subprocess_window():
    """Suppresses console window creation for background Chrome/Chromium processes on Windows."""
    if sys.platform == "win32":
        _original_popen_init = subprocess.Popen.__init__

        def _patched_popen_init(self, *args, **kwargs):
            if "creationflags" not in kwargs:
                kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
            else:
                kwargs["creationflags"] |= subprocess.CREATE_NO_WINDOW
            _original_popen_init(self, *args, **kwargs)

        subprocess.Popen.__init__ = _patched_popen_init

