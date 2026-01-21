from __future__ import annotations

import os
import datetime as dt
from bisect import bisect_left, bisect_right
from pathlib import Path
from typing import Any, Dict, List, Tuple

from lemure_reader import load_test, ChannelInfo

STATE: Dict[str, Any] = {
    "loaded": False,
    "folder": "",
    "data": None,
    "t_list": [],
}


def build_state(folder: str) -> Dict[str, Any]:
    data = load_test(folder)
    t_list = [r["t_ms"] for r in data["rows"]]
    return {"loaded": True, "folder": folder, "data": data, "t_list": t_list}


def channel_to_dict(ch: ChannelInfo) -> Dict[str, str]:
    return {"code": ch.code, "name": ch.name, "unit": ch.unit, "label": ch.label}


def summary(data: Dict[str, Any]) -> Dict[str, Any]:
    rows = data["rows"]
    if not rows:
        return {"points": 0}
    t0 = rows[0]["t_ms"]
    t1 = rows[-1]["t_ms"]
    return {
        "points": len(rows),
        "start_ms": t0,
        "end_ms": t1,
        "start": dt.datetime.fromtimestamp(t0 / 1000).strftime("%Y-%m-%d %H:%M:%S"),
        "end": dt.datetime.fromtimestamp(t1 / 1000).strftime("%Y-%m-%d %H:%M:%S"),
    }


def slice_by_time(t_list: List[int], start_ms: int, end_ms: int) -> Tuple[int, int]:
    if start_ms > end_ms:
        start_ms, end_ms = end_ms, start_ms
    i0 = bisect_left(t_list, start_ms)
    i1 = bisect_right(t_list, end_ms)
    return i0, i1


def validate_folder_path(folder: str) -> bool:
    try:
        real_path = Path(folder).resolve()
        return real_path.exists() and real_path.is_dir()
    except Exception:
        return False


def nearest_index(t_list: List[int], target_ms: int) -> int:
    if not t_list:
        return -1
    i = bisect_left(t_list, target_ms)
    if i <= 0:
        return 0
    if i >= len(t_list):
        return len(t_list) - 1
    a = t_list[i - 1]
    b = t_list[i]
    if abs(target_ms - a) <= abs(b - target_ms):
        return i - 1
    return i


class ChannelResolver:
    """Fast O(1) channel lookup (kept for future use)."""

    def __init__(self, cols: List[str]):
        self.cols = set(cols)
        self.by_suffix: Dict[str, List[str]] = {}
        for col in cols:
            if '-' in col:
                suffix = col.split('-', 1)[1]
                self.by_suffix.setdefault(suffix, []).append(col)

    def resolve(self, key: str, prefer: List[str] | None = None) -> str:
        if key in self.cols:
            return key
        candidates = self.by_suffix.get(key, [])
        if not candidates:
            return ""
        prefer = prefer or ["A-", "C-"]
        for prefix in prefer:
            for c in candidates:
                if c.startswith(prefix):
                    return c
        return candidates[0]
