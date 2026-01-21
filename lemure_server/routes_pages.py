from __future__ import annotations

from pathlib import Path

from flask import Blueprint, render_template

from .config import APP_PORT, PROJECT_ROOT
from .utils import now_asset_version_fallback

pages_bp = Blueprint('pages', __name__)


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
    try:
        base = Path(PROJECT_ROOT)
        mt = 0.0
        # static files
        for p in (base / 'static').rglob('*'):
            try:
                if p.is_file():
                    mt = max(mt, p.stat().st_mtime)
            except Exception:
                pass
        # templates
        for p in (base / 'templates').rglob('*.html'):
            try:
                mt = max(mt, p.stat().st_mtime)
            except Exception:
                pass
        asset_v = str(int(mt)) if mt > 0 else now_asset_version_fallback()
    except Exception:
        asset_v = now_asset_version_fallback()

    return render_template('index.html', port=APP_PORT, asset_v=asset_v)


@pages_bp.route('/favicon.ico')
def favicon():
    return ('', 204)
