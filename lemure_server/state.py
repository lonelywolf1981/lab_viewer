from __future__ import annotations

import os
import datetime as dt
import threading
from bisect import bisect_left, bisect_right
from pathlib import Path
from typing import Any, Dict, List, Tuple

from lemure_reader import load_test, ChannelInfo

STATE_LOCK = threading.Lock()

STATE: Dict[str, Any] = {
    "loaded": False,
    "folder": "",
    "data": None,
    "t_list": [],
}

DATA_VERSION: int = 0


def get_data_version() -> int:
    return DATA_VERSION


def build_state(folder: str) -> Dict[str, Any]:
    global DATA_VERSION
    data = load_test(folder)
    t_list = data["t_ms"]
    with STATE_LOCK:
        DATA_VERSION += 1
    return {"loaded": True, "folder": folder, "data": data, "t_list": t_list}


def channel_to_dict(ch: ChannelInfo) -> Dict[str, str]:
    return {"code": ch.code, "name": ch.name, "unit": ch.unit, "label": ch.label}


def summary(data: Dict[str, Any]) -> Dict[str, Any]:
    t_ms = data["t_ms"]
    if not t_ms:
        return {"points": 0}
    t0 = t_ms[0]
    t1 = t_ms[-1]
    return {
        "points": len(t_ms),
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
