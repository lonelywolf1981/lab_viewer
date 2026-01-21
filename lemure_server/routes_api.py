from __future__ import annotations

import io
import os
import datetime as dt
from typing import Any, Dict, List

from flask import Blueprint, jsonify, request, send_file

from .config import APP_PORT, PROJECT_ROOT, send_file_compat
from .settings import get_viewer_settings, normalize_viewer_settings, save_viewer_settings, set_viewer_settings
from .state import STATE, build_state, channel_to_dict, summary, slice_by_time, validate_folder_path
from .series_cache import cached_series_slice, clear_cache
from .persistence import (
    load_saved_order,
    save_order,
    list_orders,
    sanitize_key,
    save_named_order,
    load_named_order,
    list_presets,
    save_preset,
    load_preset,
    delete_preset,
    ensure_orders_dir,
    ensure_presets_dir,
)
from .exports.basic import export_csv, export_xlsx
from .exports.template import export_template_impl
from .utils import log_exception_to_file

api_bp = Blueprint('api', __name__)


@api_bp.route('/api/settings', methods=['GET', 'POST'])
def api_settings():
    if request.method == 'GET':
        return jsonify({'ok': True, 'settings': get_viewer_settings()})

    body = request.get_json(force=True, silent=True) or {}
    user_s = body.get('settings') if isinstance(body, dict) and isinstance(body.get('settings'), dict) else body
    s = normalize_viewer_settings(user_s if isinstance(user_s, dict) else {})
    set_viewer_settings(s)
    save_viewer_settings(s)
    return jsonify({'ok': True, 'settings': s})


@api_bp.route('/api/pick_folder', methods=['POST'])
def pick_folder():
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.wm_attributes('-topmost', 1)
        folder = filedialog.askdirectory(title='Выберите папку с тестом (где Prova*.dbf)')
        root.destroy()
        return jsonify({'ok': True, 'folder': folder or ''})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e), 'folder': ''})


@api_bp.route('/api/load', methods=['POST'])
def api_load():
    body = request.get_json(force=True, silent=True) or {}
    folder = (body.get('folder') or '').strip()

    if not validate_folder_path(folder):
        return jsonify({'ok': False, 'error': 'Недопустимый путь'})
    if not folder:
        return jsonify({'ok': False, 'error': 'Путь к папке пустой'})
    if not os.path.isdir(folder):
        return jsonify({'ok': False, 'error': 'Папка не существует: ' + folder})

    try:
        st = build_state(folder)
        STATE.update(st)
        clear_cache()

        data = STATE['data']
        channels = data['channels']
        cols = data['cols']

        ch_list = []
        for c in cols:
            if c in channels:
                ch_list.append(channel_to_dict(channels[c]))
            else:
                ch_list.append({'code': c, 'name': '', 'unit': '', 'label': c})

        return jsonify({
            'ok': True,
            'folder': data['root'],
            'meta': data['meta'],
            'summary': summary(data),
            'channels': ch_list,
            'file_order': [c.get('code') for c in ch_list],
            'saved_order': load_saved_order(),
        })
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})


@api_bp.route('/api/series', methods=['GET'])
def api_series():
    if not STATE.get('loaded'):
        return jsonify({'ok': False, 'error': 'Данные не загружены'})

    data = STATE['data']
    rows: List[Dict[str, Any]] = data['rows']
    t_list: List[int] = STATE['t_list']

    channels = request.args.get('channels', '')
    ch = [c for c in channels.split(',') if c.strip()]
    if not ch:
        return jsonify({'ok': False, 'error': 'Не выбраны каналы'})

    try:
        start_ms = int(float(request.args.get('start_ms', rows[0]['t_ms'])))
        end_ms = int(float(request.args.get('end_ms', rows[-1]['t_ms'])))
    except Exception:
        start_ms = rows[0]['t_ms']
        end_ms = rows[-1]['t_ms']

    if start_ms > end_ms:
        start_ms, end_ms = end_ms, start_ms

    try:
        step_i = max(1, int(request.args.get('step', '1')))
    except Exception:
        step_i = 1

    max_points = request.args.get('max_points', '')
    try:
        target = int(float(max_points)) if str(max_points).strip() != '' else 0
    except Exception:
        target = 0

    i0, i1 = slice_by_time(t_list, start_ms, end_ms)
    raw_pts = max(0, i1 - i0)

    if target and target > 0:
        step_i = max(1, (raw_pts + target - 1) // target)

    pts_after = max(0, (raw_pts + step_i - 1) // step_i) if raw_pts > 0 else 0
    cells = pts_after * len(ch)
    use_cache = (pts_after <= 40000) and (cells <= 400000)

    channels_key = ','.join(ch)

    try:
        if use_cache:
            t_ms, series = cached_series_slice(id(rows), channels_key, start_ms, end_ms, step_i)
        else:
            sliced = rows[i0:i1:step_i]
            t_ms = [r['t_ms'] for r in sliced]

            def _to_float(v):
                if v is None:
                    return None
                if isinstance(v, (int, float)):
                    return float(v)
                try:
                    return float(str(v).replace(',', '.'))
                except Exception:
                    return None

            series = {code: [_to_float(r.get(code)) for r in sliced] for code in ch}

        return jsonify({'ok': True, 't_ms': t_ms, 'series': series, 'step': step_i, 'points': len(t_ms)})

    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})


@api_bp.route('/api/range_stats', methods=['GET'])
def api_range_stats():
    if not STATE.get('loaded'):
        return jsonify({'ok': False, 'error': 'Данные не загружены'}), 400

    t_list = STATE.get('t_list') or []
    if not t_list:
        return jsonify({'ok': True, 'points': 0, 'total': 0})

    try:
        start_ms = int(float(request.args.get('start_ms', t_list[0])))
        end_ms = int(float(request.args.get('end_ms', t_list[-1])))
    except Exception:
        start_ms = t_list[0]
        end_ms = t_list[-1]

    if start_ms > end_ms:
        start_ms, end_ms = end_ms, start_ms

    i0, i1 = slice_by_time(t_list, start_ms, end_ms)
    pts = max(0, i1 - i0)
    return jsonify({'ok': True, 'start_ms': start_ms, 'end_ms': end_ms, 'points': pts, 'total': len(t_list)})


@api_bp.route('/api/save_order', methods=['POST'])
def api_save_order():
    body = request.get_json(force=True, silent=True) or {}
    order = body.get('order')
    if not isinstance(order, list):
        return jsonify({'ok': False, 'error': 'order must be a list'}), 400

    seen = set()
    normalized = []
    for x in order:
        s = str(x)
        if s not in seen:
            seen.add(s)
            normalized.append(s)

    save_order(normalized)
    return jsonify({'ok': True, 'saved': len(normalized)})


@api_bp.route('/api/orders_list', methods=['GET'])
def api_orders_list():
    try:
        ensure_orders_dir()
        return jsonify({'ok': True, 'orders': list_orders()})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e), 'orders': []})


@api_bp.route('/api/orders_save', methods=['POST'])
def api_orders_save():
    body = request.get_json(force=True, silent=True) or {}
    name = str(body.get('name') or '').strip()
    order = body.get('order')
    if not name:
        return jsonify({'ok': False, 'error': 'Имя не задано'}), 400
    if not isinstance(order, list):
        return jsonify({'ok': False, 'error': 'order must be a list'}), 400

    key = sanitize_key(name)
    if not key:
        return jsonify({'ok': False, 'error': 'Имя некорректное'}), 400

    seen = set()
    normalized = []
    for x in order:
        s = str(x)
        if s and s not in seen:
            seen.add(s)
            normalized.append(s)

    try:
        save_named_order(name, normalized)
        return jsonify({'ok': True, 'key': key, 'saved': len(normalized)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@api_bp.route('/api/orders_load', methods=['GET'])
def api_orders_load():
    key = str(request.args.get('key') or '').strip()
    name = str(request.args.get('name') or '').strip()
    if not key and name:
        key = sanitize_key(name)
    if not key:
        return jsonify({'ok': False, 'error': 'Не задан ключ (key)'}), 400

    try:
        data = load_named_order(key)
        if not data:
            return jsonify({'ok': False, 'error': 'Сохранённый порядок не найден: ' + key}), 404
        order = data.get('order')
        if not isinstance(order, list):
            order = []
        out = []
        seen = set()
        for x in order:
            s = str(x)
            if s and s not in seen:
                seen.add(s)
                out.append(s)
        return jsonify({'ok': True, 'key': key, 'name': str(data.get('name') or key), 'order': out})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@api_bp.route('/api/presets_list', methods=['GET'])
def api_presets_list():
    try:
        ensure_presets_dir()
        return jsonify({'ok': True, 'presets': list_presets()})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e), 'presets': []})


@api_bp.route('/api/presets_save', methods=['POST'])
def api_presets_save():
    body = request.get_json(force=True, silent=True) or {}
    name = str(body.get('name') or '').strip()
    preset = body.get('preset') if isinstance(body.get('preset'), dict) else {}
    if not name:
        return jsonify({'ok': False, 'error': 'Имя не задано'}), 400

    key = sanitize_key(name)
    if not key:
        return jsonify({'ok': False, 'error': 'Имя некорректное'}), 400

    channels = preset.get('channels')
    if not isinstance(channels, list):
        channels = []
    channels_n = []
    seen = set()
    for x in channels:
        s = str(x)
        if s and s not in seen:
            seen.add(s)
            channels_n.append(s)

    order = preset.get('order')
    if isinstance(order, list):
        o2 = []
        seen2 = set()
        for x in order:
            s = str(x)
            if s and s not in seen2:
                seen2.add(s)
                o2.append(s)
        order = o2
    else:
        order = []

    payload = {
        'name': name,
        'key': key,
        'saved_at': dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'channels': channels_n,
        'sort_mode': str(preset.get('sort_mode') or ''),
        'order': order,
        'step_auto': bool(preset.get('step_auto')) if isinstance(preset.get('step_auto'), bool) else True,
        'step_target': int(preset.get('step_target') or 5000),
        'step': int(preset.get('step') or 1),
        'show_legend': bool(preset.get('show_legend')) if isinstance(preset.get('show_legend'), bool) else True,
    }

    try:
        save_preset(payload)
        return jsonify({'ok': True, 'key': key, 'count': len(channels_n)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@api_bp.route('/api/presets_load', methods=['GET'])
def api_presets_load():
    key = str(request.args.get('key') or '').strip()
    name = str(request.args.get('name') or '').strip()
    if not key and name:
        key = sanitize_key(name)
    if not key:
        return jsonify({'ok': False, 'error': 'Не задан ключ (key)'}), 400

    try:
        data = load_preset(key)
        if not data:
            return jsonify({'ok': False, 'error': 'Сохранённый набор не найден: ' + key}), 404
        preset = dict(data)
        preset.pop('ok', None)
        return jsonify({'ok': True, 'key': key, 'name': str(data.get('name') or key), 'preset': preset})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@api_bp.route('/api/presets_delete', methods=['POST'])
def api_presets_delete():
    body = request.get_json(force=True, silent=True) or {}
    key = str(body.get('key') or '').strip()
    if not key:
        return jsonify({'ok': False, 'error': 'Не задан key'}), 400

    try:
        ok = delete_preset(key)
        if not ok:
            return jsonify({'ok': False, 'error': 'Набор не найден: ' + key}), 404
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@api_bp.route('/api/export', methods=['GET'])
def api_export():
    if not STATE.get('loaded'):
        return jsonify({'ok': False, 'error': 'Данные не загружены'}), 400

    data = STATE['data']
    rows: List[Dict[str, Any]] = data['rows']
    t_list: List[int] = STATE['t_list']

    fmt = (request.args.get('format', 'csv') or 'csv').lower()
    channels = request.args.get('channels', '')
    ch = [c for c in channels.split(',') if c.strip()]
    if not ch:
        return jsonify({'ok': False, 'error': 'Не выбраны каналы'}), 400

    try:
        start_ms = int(float(request.args.get('start_ms', rows[0]['t_ms'])))
        end_ms = int(float(request.args.get('end_ms', rows[-1]['t_ms'])))
    except Exception:
        start_ms = rows[0]['t_ms']
        end_ms = rows[-1]['t_ms']

    try:
        step_i = max(1, int(request.args.get('step', '1')))
    except Exception:
        step_i = 1

    i0, i1 = slice_by_time(t_list, start_ms, end_ms)
    sliced = rows[i0:i1:step_i]

    if fmt == 'xlsx':
        payload = export_xlsx(sliced, ch)
        return send_file_compat(send_file, io.BytesIO(payload),
                                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                'export.xlsx')

    payload = export_csv(sliced, ch)
    return send_file_compat(send_file, io.BytesIO(payload), 'text/csv', 'export.csv')


@api_bp.route('/api/export_template', methods=['GET'])
def api_export_template():
    try:
        result = export_template_impl(request.args)
        # if export_template_impl already returned a flask response/tuple
        if isinstance(result, tuple) and len(result) == 2 and isinstance(result[1], int):
            return result
        # else: (bio, filename, headers)
        bio, fn, headers = result
        resp = send_file_compat(send_file, bio,
                                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                fn)
        try:
            for k, v in (headers or {}).items():
                resp.headers[k] = v
        except Exception:
            pass
        return resp
    except Exception as e:
        lp = log_exception_to_file('api_export_template', e, project_root=PROJECT_ROOT)
        msg = str(e)
        if lp:
            msg = msg + f" (подробности в {os.path.basename(lp)})"
        return jsonify({'ok': False, 'error': msg}), 500
