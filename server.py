"""LeMuRe Viewer server entrypoint.

Run:
  python server.py

The implementation is split into functional modules under `lemure_server/`.
"""

from __future__ import annotations

from lemure_server.app_factory import create_app
from lemure_server.config import APP_HOST, APP_PORT
from lemure_server.persistence import ensure_dirs
from lemure_server.utils import open_browser_later


app = create_app()


if __name__ == '__main__':
    ensure_dirs()
    open_browser_later(APP_HOST, APP_PORT)
    # threaded=True keeps UI responsive; use_reloader=False to avoid double-run
    app.run(host=APP_HOST, port=APP_PORT, debug=False, threaded=True, use_reloader=False)
