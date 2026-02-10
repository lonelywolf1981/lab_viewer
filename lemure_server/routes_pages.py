from __future__ import annotations

from pathlib import Path

from flask import Blueprint, render_template

from .config import APP_PORT, PROJECT_ROOT
from .utils import now_asset_version_fallback

pages_bp = Blueprint('pages', __name__)


def _compute_asset_version() -> str:
    """Compute asset version once from max mtime of static/templates."""
    try:
        base = Path(PROJECT_ROOT)
        mt = 0.0
        for p in (base / 'static').rglob('*'):
            try:
                if p.is_file():
                    mt = max(mt, p.stat().st_mtime)
            except Exception:
                pass
        for p in (base / 'templates').rglob('*.html'):
            try:
                mt = max(mt, p.stat().st_mtime)
            except Exception:
                pass
        return str(int(mt)) if mt > 0 else now_asset_version_fallback()
    except Exception:
        return now_asset_version_fallback()


_ASSET_V: str = _compute_asset_version()


@pages_bp.after_app_request
def add_no_cache_headers(resp):
    try:
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        resp.headers["Pragma"] = "no-cache"
        resp.headers["Expires"] = "0"
    except Exception:
        pass
    return resp


@pages_bp.route('/')
def index():
    return render_template('index.html', port=APP_PORT, asset_v=_ASSET_V)


@pages_bp.route('/favicon.ico')
def favicon():
    return ('', 204)
