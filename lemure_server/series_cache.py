from __future__ import annotations

from functools import lru_cache
from typing import Any, Dict, List, Tuple

from .state import STATE, slice_by_time


@lru_cache(maxsize=8)
def cached_series_slice(data_version: int, channels_key: str, start_ms: int, end_ms: int, step_i: int) -> Tuple[List[int], Dict[str, List[float | None]]]:
    """Return (t_ms, series_dict) for the requested channels and time range.

    Cached to speed up repeated redraws with the same parameters.
    data_version is a monotonic counter that changes when data is reloaded.
    """
    data = STATE["data"]
    t_list = STATE["t_list"]
    columns = data["columns"]
    channels = [c.strip() for c in channels_key.split(",") if c.strip()]

    i0, i1 = slice_by_time(t_list, start_ms, end_ms)
    t = t_list[i0:i1:step_i]

    series: Dict[str, List[float | None]] = {}
    n = len(t)
    for code in channels:
        col = columns.get(code)
        if col is not None:
            series[code] = col[i0:i1:step_i]
        else:
            series[code] = [None] * n

    return t, series


def clear_cache() -> None:
    try:
        cached_series_slice.cache_clear()
    except Exception:
        pass
