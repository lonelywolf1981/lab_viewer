from __future__ import annotations

from functools import lru_cache
from typing import Any, Dict, List, Tuple

from .state import STATE, slice_by_time


@lru_cache(maxsize=8)
def cached_series_slice(data_id: int, channels_key: str, start_ms: int, end_ms: int, step_i: int) -> Tuple[List[int], Dict[str, List[float | None]]]:
    """Return (t_ms, series_dict) for the requested channels and time range.

    Cached to speed up repeated redraws with the same parameters.
    """
    rows = STATE["data"]["rows"]
    t_list = STATE["t_list"]
    channels = [c.strip() for c in channels_key.split(",") if c.strip()]

    i0, i1 = slice_by_time(t_list, start_ms, end_ms)
    sliced = rows[i0:i1:step_i]

    t = [r["t_ms"] for r in sliced]

    def _to_float(v: Any):
        if v is None:
            return None
        if isinstance(v, (int, float)):
            return float(v)
        try:
            return float(str(v).replace(",", "."))
        except Exception:
            return None

    series: Dict[str, List[float | None]] = {}
    for code in channels:
        series[code] = [_to_float(r.get(code)) for r in sliced]

    return t, series


def clear_cache() -> None:
    try:
        cached_series_slice.cache_clear()
    except Exception:
        pass
