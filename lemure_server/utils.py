from __future__ import annotations

import os
import datetime as dt
import threading
import time as time_mod
import webbrowser


def log_exception_to_file(prefix: str, exc: Exception, *, project_root: str) -> str:
    """Write traceback to a local file and return that file path."""
    import traceback as _tb

    try:
        log_path = os.path.join(project_root, "export_template_error.log")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write("\n===== " + prefix + " " + dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + " =====\n")
            f.write(_tb.format_exc())
            f.write("\n")
        try:
            print("[TEMPLATE] ERROR logged to", log_path)
        except Exception:
            pass
        return log_path
    except Exception:
        return ""


def open_browser_later(host: str, port: int, delay_s: float = 0.8) -> None:
    def _open():
        try:
            webbrowser.open(f"http://{host}:{port}/", new=1)
        except Exception:
            pass

    threading.Timer(delay_s, _open).start()


def now_asset_version_fallback() -> str:
    return str(int(time_mod.time()))
