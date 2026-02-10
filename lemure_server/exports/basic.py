from __future__ import annotations

import io
import datetime as dt
from typing import Any, Dict, List


def export_csv(t_ms_list: List[int], columns: Dict[str, list], channels: List[str]) -> bytes:
    import csv

    out = io.StringIO()
    writer = csv.writer(out, delimiter=";")
    writer.writerow(["timestamp"] + channels)
    for i, t_ms in enumerate(t_ms_list):
        ts = dt.datetime.fromtimestamp(t_ms / 1000).strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        row = [ts]
        for c in channels:
            col = columns.get(c)
            v = col[i] if col is not None and i < len(col) else None
            row.append(v if v is not None else "")
        writer.writerow(row)
    return out.getvalue().encode("utf-8-sig")


def export_xlsx(t_ms_list: List[int], columns: Dict[str, list], channels: List[str]) -> bytes:
    from openpyxl import Workbook

    wb = Workbook(write_only=True)
    ws = wb.create_sheet(title="data")
    ws.append(["timestamp"] + channels)
    for i, t_ms in enumerate(t_ms_list):
        ts = dt.datetime.fromtimestamp(t_ms / 1000).strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        row = [ts]
        for c in channels:
            col = columns.get(c)
            row.append(col[i] if col is not None and i < len(col) else None)
        ws.append(row)
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()
