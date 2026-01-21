from __future__ import annotations

import json
import os
import re
import threading
import datetime as dt
from typing import Any, Dict, List

from .config import ORDER_FILE, ORDERS_DIR, PRESETS_DIR

PRESETS_LOCK = threading.Lock()


def ensure_orders_dir() -> None:
    try:
        os.makedirs(ORDERS_DIR, exist_ok=True)
    except Exception:
        pass


def ensure_presets_dir() -> None:
    try:
        os.makedirs(PRESETS_DIR, exist_ok=True)
    except Exception:
        pass


def ensure_dirs() -> None:
    ensure_orders_dir()
    ensure_presets_dir()


def sanitize_key(name: str) -> str:
    name = (name or '').strip()
    name = re.sub(r'[^0-9A-Za-zА-Яа-я _.-]+', '_', name)
    name = re.sub(r'\s+', ' ', name).strip()
    name = name.replace('..', '.')
    if len(name) > 64:
        name = name[:64].strip()
    return name


def order_path_by_key(key: str) -> str:
    ensure_orders_dir()
    return os.path.join(ORDERS_DIR, f'{key}.json')


def preset_path_by_key(key: str) -> str:
    ensure_presets_dir()
    return os.path.join(PRESETS_DIR, f'{key}.json')


def load_saved_order() -> List[str]:
    try:
        if os.path.isfile(ORDER_FILE):
            with open(ORDER_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict) and isinstance(data.get("order"), list):
                return [str(x) for x in data["order"]]
            if isinstance(data, list):
                return [str(x) for x in data]
    except Exception:
        pass
    return []


def save_order(order: List[str]) -> None:
    try:
        with open(ORDER_FILE, "w", encoding="utf-8") as f:
            json.dump({"order": order}, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def list_orders() -> List[Dict[str, Any]]:
    ensure_orders_dir()
    items = []
    for fn in os.listdir(ORDERS_DIR):
        if not fn.lower().endswith('.json'):
            continue
        path = os.path.join(ORDERS_DIR, fn)
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            key = str(data.get('key') or os.path.splitext(fn)[0])
            name = str(data.get('name') or key)
            order = data.get('order')
            count = len(order) if isinstance(order, list) else 0
            saved_at = str(data.get('saved_at') or '')
            items.append({"key": key, "name": name, "count": count, "saved_at": saved_at})
        except Exception:
            continue
    items.sort(key=lambda x: x.get('name', '').lower())
    return items


def save_named_order(name: str, order: List[str]) -> Dict[str, Any]:
    key = sanitize_key(name)
    payload = {
        "name": name,
        "key": key,
        "order": order,
        "saved_at": dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    }
    path = order_path_by_key(key)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    return payload


def load_named_order(key_or_name: str) -> Dict[str, Any] | None:
    key = sanitize_key(key_or_name)
    path = order_path_by_key(key)
    if not os.path.isfile(path):
        return None
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data if isinstance(data, dict) else None


def list_presets() -> List[Dict[str, Any]]:
    ensure_presets_dir()
    items = []
    for fn in os.listdir(PRESETS_DIR):
        if not fn.lower().endswith('.json'):
            continue
        path = os.path.join(PRESETS_DIR, fn)
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            key = str(data.get('key') or os.path.splitext(fn)[0])
            name = str(data.get('name') or key)
            channels = data.get('channels')
            if not isinstance(channels, list):
                channels = (data.get('preset') or {}).get('channels') if isinstance(data.get('preset'), dict) else []
            count = len(channels) if isinstance(channels, list) else 0
            saved_at = str(data.get('saved_at') or '')
            items.append({"key": key, "name": name, "count": count, "saved_at": saved_at})
        except Exception:
            continue
    items.sort(key=lambda x: x.get('name', '').lower())
    return items


def save_preset(payload: Dict[str, Any]) -> None:
    key = str(payload.get('key') or '')
    if not key:
        raise ValueError('preset key missing')
    path = preset_path_by_key(key)
    with PRESETS_LOCK:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)


def load_preset(key_or_name: str) -> Dict[str, Any] | None:
    key = sanitize_key(key_or_name)
    path = preset_path_by_key(key)
    if not os.path.isfile(path):
        return None
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data if isinstance(data, dict) else None


def delete_preset(key_or_name: str) -> bool:
    key = sanitize_key(key_or_name)
    path = preset_path_by_key(key)
    if not os.path.isfile(path):
        return False
    with PRESETS_LOCK:
        os.remove(path)
    return True
