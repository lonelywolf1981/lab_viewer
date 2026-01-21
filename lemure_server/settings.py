from __future__ import annotations

import json
import os
import re
from typing import Any, Dict

from .config import SETTINGS_FILE

DEFAULT_VIEWER_SETTINGS: Dict[str, Any] = {
    'row_mark': {
        'threshold_T': 150,
        'color': '#EAD706',  # мягкий жёлтый
        'intensity': 100,
    },
    'discharge_mark': {'threshold': None, 'color': '#FFC000'},
    'suction_mark': {'threshold': None, 'color': '#00B0F0'},
    'scales': {
        'W': {'min': -1, 'opt': 1, 'max': 2, 'colors': {'min': '#1CBCF2', 'opt': '#00FF00', 'max': '#F3919B'}},
        'X': {'min': -1, 'opt': 1, 'max': 2, 'colors': {'min': '#1CBCF2', 'opt': '#00FF00', 'max': '#F3919B'}},
        'Y': {'min': -1, 'opt': 1, 'max': 2, 'colors': {'min': '#1CBCF2', 'opt': '#00FF00', 'max': '#F3919B'}},
    },
}


def _deep_merge(dst: Dict[str, Any], src: Dict[str, Any]) -> Dict[str, Any]:
    for k, v in (src or {}).items():
        if isinstance(v, dict) and isinstance(dst.get(k), dict):
            _deep_merge(dst[k], v)
        else:
            dst[k] = v
    return dst


def _normalize_hex_color(s: str, default: str = '#FFF2CC') -> str:
    s = (s or '').strip()
    if not s:
        return default
    if not s.startswith('#'):
        s = '#' + s
    if re.fullmatch(r'#[0-9A-Fa-f]{6}', s):
        return s.upper()
    return default


def _argb_from_hex_and_intensity(hex_color: str, intensity_0_100: int) -> str:
    """Convert CSS #RRGGBB + intensity to Excel ARGB (FFRRGGBB), blending with white.
    intensity=100 -> original color; intensity=0 -> white.
    """
    hex_color = _normalize_hex_color(hex_color)
    i = max(0, min(100, int(intensity_0_100)))
    w = 1.0 - i / 100.0
    r = int(hex_color[1:3], 16)
    g = int(hex_color[3:5], 16)
    b = int(hex_color[5:7], 16)
    r2 = int(round(r * (1.0 - w) + 255 * w))
    g2 = int(round(g * (1.0 - w) + 255 * w))
    b2 = int(round(b * (1.0 - w) + 255 * w))
    return f'FF{r2:02X}{g2:02X}{b2:02X}'


def load_viewer_settings() -> Dict[str, Any]:
    # Start with defaults
    s = json.loads(json.dumps(DEFAULT_VIEWER_SETTINGS))

    user_s: Dict[str, Any] = {}
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                user_s = json.load(f)
            if not isinstance(user_s, dict):
                user_s = {}
    except Exception:
        user_s = {}

    # Backward compatibility (older keys)
    if isinstance(user_s, dict):
        if 'yellow_threshold_T' in user_s and 'row_mark' not in user_s:
            try:
                s['row_mark']['threshold_T'] = float(user_s.get('yellow_threshold_T'))
            except Exception:
                pass
        if 'yellow_intensity' in user_s and 'row_mark' not in user_s:
            try:
                s['row_mark']['intensity'] = int(user_s.get('yellow_intensity'))
            except Exception:
                pass

    _deep_merge(s, user_s)

    # Normalize
    rm = s.get('row_mark') if isinstance(s.get('row_mark'), dict) else {}
    try:
        rm['threshold_T'] = float(rm.get('threshold_T', 150))
    except Exception:
        rm['threshold_T'] = 150.0
    try:
        rm['intensity'] = max(0, min(100, int(rm.get('intensity', 100))))
    except Exception:
        rm['intensity'] = 100
    rm['color'] = _normalize_hex_color(str(rm.get('color') or ''), default='#FFF2CC')
    s['row_mark'] = rm

    # discharge_mark
    dm = s.get('discharge_mark') if isinstance(s.get('discharge_mark'), dict) else {}
    thr = dm.get('threshold', None)
    if thr is None or str(thr).strip() == '':
        dm['threshold'] = None
    else:
        try:
            dm['threshold'] = float(thr)
        except Exception:
            dm['threshold'] = None
    dm['color'] = _normalize_hex_color(str(dm.get('color') or ''), default='#FFC000')
    s['discharge_mark'] = dm

    # suction_mark
    sm = s.get('suction_mark') if isinstance(s.get('suction_mark'), dict) else {}
    thr = sm.get('threshold', None)
    if thr is None or str(thr).strip() == '':
        sm['threshold'] = None
    else:
        try:
            sm['threshold'] = float(thr)
        except Exception:
            sm['threshold'] = None
    sm['color'] = _normalize_hex_color(str(sm.get('color') or ''), default='#00B0F0')
    s['suction_mark'] = sm

    # scales
    scales = s.get('scales') if isinstance(s.get('scales'), dict) else {}
    out_scales = {}
    for k, defaults in DEFAULT_VIEWER_SETTINGS['scales'].items():
        d = scales.get(k, {}) if isinstance(scales.get(k), dict) else {}
        merged = dict(defaults)
        merged.update(d)
        for kk in ('min', 'opt', 'max'):
            try:
                merged[kk] = float(merged.get(kk))
            except Exception:
                merged[kk] = float(defaults[kk])
        if merged['min'] >= merged['opt']:
            merged['min'] = merged['opt'] - 1
        if merged['opt'] >= merged['max']:
            merged['max'] = merged['opt'] + 1

        colors = merged.get('colors') if isinstance(merged.get('colors'), dict) else {}

        def _col(key, default):
            return _normalize_hex_color(str(colors.get(key) or ''), default=default)

        merged['colors'] = {
            'min': _col('min', '#0000FF'),
            'opt': _col('opt', '#00FF00'),
            'max': _col('max', '#FF0000'),
        }
        out_scales[k] = merged

    s['scales'] = out_scales
    return s


def normalize_viewer_settings(user_s: Dict[str, Any]) -> Dict[str, Any]:
    s = json.loads(json.dumps(DEFAULT_VIEWER_SETTINGS))
    if not isinstance(user_s, dict):
        user_s = {}
    _deep_merge(s, user_s)

    # reuse same normalization as load
    tmp = s
    s2 = load_viewer_settings()  # start from defaults+file to keep same structure
    # but overwrite with tmp
    s2 = json.loads(json.dumps(DEFAULT_VIEWER_SETTINGS))
    _deep_merge(s2, tmp)

    # normalize by saving through load-like rules
    # (cheap way: call load_viewer_settings logic on s2 by manual normalization)
    # row_mark
    rm = s2.get('row_mark') if isinstance(s2.get('row_mark'), dict) else {}
    try:
        rm['threshold_T'] = float(rm.get('threshold_T', 150))
    except Exception:
        rm['threshold_T'] = 150.0
    try:
        rm['intensity'] = max(0, min(100, int(rm.get('intensity', 100))))
    except Exception:
        rm['intensity'] = 100
    rm['color'] = _normalize_hex_color(str(rm.get('color') or ''), default='#FFF2CC')
    s2['row_mark'] = rm

    dm = s2.get('discharge_mark') if isinstance(s2.get('discharge_mark'), dict) else {}
    thr = dm.get('threshold', None)
    if thr is None or str(thr).strip() == '':
        dm['threshold'] = None
    else:
        try:
            dm['threshold'] = float(thr)
        except Exception:
            dm['threshold'] = None
    dm['color'] = _normalize_hex_color(str(dm.get('color') or ''), default='#FFC000')
    s2['discharge_mark'] = dm

    sm = s2.get('suction_mark') if isinstance(s2.get('suction_mark'), dict) else {}
    thr = sm.get('threshold', None)
    if thr is None or str(thr).strip() == '':
        sm['threshold'] = None
    else:
        try:
            sm['threshold'] = float(thr)
        except Exception:
            sm['threshold'] = None
    sm['color'] = _normalize_hex_color(str(sm.get('color') or ''), default='#00B0F0')
    s2['suction_mark'] = sm

    scales = s2.get('scales') if isinstance(s2.get('scales'), dict) else {}
    out_scales = {}
    for k, defaults in DEFAULT_VIEWER_SETTINGS['scales'].items():
        d = scales.get(k, {}) if isinstance(scales.get(k), dict) else {}
        merged = dict(defaults)
        merged.update(d)
        for kk in ('min', 'opt', 'max'):
            try:
                merged[kk] = float(merged.get(kk))
            except Exception:
                merged[kk] = float(defaults[kk])
        if merged['min'] >= merged['opt']:
            merged['min'] = merged['opt'] - 1
        if merged['opt'] >= merged['max']:
            merged['max'] = merged['opt'] + 1

        colors = merged.get('colors') if isinstance(merged.get('colors'), dict) else {}

        def _col(key, default):
            return _normalize_hex_color(str(colors.get(key) or ''), default=default)

        merged['colors'] = {
            'min': _col('min', '#0000FF'),
            'opt': _col('opt', '#00FF00'),
            'max': _col('max', '#FF0000'),
        }
        out_scales[k] = merged
    s2['scales'] = out_scales

    return s2


def save_viewer_settings(settings: Dict[str, Any]) -> None:
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


VIEWER_SETTINGS: Dict[str, Any] = load_viewer_settings()


def get_viewer_settings() -> Dict[str, Any]:
    return VIEWER_SETTINGS


def set_viewer_settings(settings: Dict[str, Any]) -> Dict[str, Any]:
    global VIEWER_SETTINGS
    VIEWER_SETTINGS = settings
    return VIEWER_SETTINGS
