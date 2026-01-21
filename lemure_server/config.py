from __future__ import annotations

import os

# Project root (one level above this package)
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

APP_HOST = "127.0.0.1"
APP_PORT = 8787

# Files in the project root
ORDER_FILE = os.path.join(PROJECT_ROOT, "channel_order.json")
TEMPLATE_FILE = os.path.join(PROJECT_ROOT, "template.xlsx")

ORDERS_DIR = os.path.join(PROJECT_ROOT, "saved_orders")
PRESETS_DIR = os.path.join(PROJECT_ROOT, "saved_presets")

SETTINGS_FILE = os.path.join(PROJECT_ROOT, "viewer_settings.json")


def send_file_compat(send_file_fn, fp, mimetype: str, filename: str):
    """send_file compat for different Flask versions (download_name vs attachment_filename)."""
    try:
        return send_file_fn(fp, mimetype=mimetype, as_attachment=True, download_name=filename)
    except TypeError:
        return send_file_fn(fp, mimetype=mimetype, as_attachment=True, attachment_filename=filename)
