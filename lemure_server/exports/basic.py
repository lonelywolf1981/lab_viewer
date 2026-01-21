from __future__ import annotations

import io
import datetime as dt
from typing import Any, Dict, List


def export_csv(rows: List[Dict[str, Any]], channels: List[str]) -> bytes:
    import csv

    out = io.StringIO()
    writer = csv.writer(out, delimiter=";")
    writer.writerow(["timestamp"] + channels)
    for r in rows:
        ts = dt.datetime.fromtimestamp(r["t_ms"] / 1000).strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        writer.writerow([ts] + [r.get(c) if r.get(c) is not None else "" for c in channels])
    return out.getvalue().encode("utf-8-sig")


def export_xlsx(rows: List[Dict[str, Any]], channels: List[str]) -> bytes:
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "data"
    ws.append(["timestamp"] + channels)
    for r in rows:
        ts = dt.datetime.fromtimestamp(r["t_ms"] / 1000).strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        ws.append([ts] + [r.get(c) for c in channels])
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()
