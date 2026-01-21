from __future__ import annotations

import os
from flask import Flask

from .config import PROJECT_ROOT
from .routes_pages import pages_bp
from .routes_api import api_bp


def create_app() -> Flask:
    app = Flask(
        __name__,
        static_folder=os.path.join(PROJECT_ROOT, 'static'),
        template_folder=os.path.join(PROJECT_ROOT, 'templates'),
    )

    app.register_blueprint(pages_bp)
    app.register_blueprint(api_bp)

    return app
