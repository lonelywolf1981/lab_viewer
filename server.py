from __future__ import annotations

# server.py
# Локальный веб-просмотрщик LeMuRe (Windows 7 friendly)
# Запуск: python server.py

import os
import re
import json
import io
import threading
import webbrowser
from datetime import datetime, timedelta, time
from bisect import bisect_left, bisect_right
from typing import List, Dict, Any, Tuple

from flask import Flask, render_template, request, jsonify, send_file

def _send_file_compat(fp, mimetype: str, filename: str):
    """send_file compat for different Flask versions (download_name vs attachment_filename)."""
    try:
        return send_file(fp, mimetype=mimetype, as_attachment=True, download_name=filename)
    except TypeError:
        return send_file(fp, mimetype=mimetype, as_attachment=True, attachment_filename=filename)





from lemure_reader import load_test, ChannelInfo

APP_HOST = "127.0.0.1"
APP_PORT = 8787

app = Flask(__name__)

STATE: Dict[str, Any] = {
    "loaded": False,
    "folder": "",
    "data": None,
    "t_list": [],
}

# Order persistence (saved in the same folder as server.py)
ORDER_FILE = os.path.join(os.path.dirname(__file__), "channel_order.json")
TEMPLATE_FILE = os.path.join(os.path.dirname(__file__), "template.xlsx")
ORDERS_DIR = os.path.join(os.path.dirname(__file__), 'saved_orders')

def _ensure_orders_dir() -> None:
    try:
        os.makedirs(ORDERS_DIR, exist_ok=True)
    except Exception:
        pass

def _sanitize_order_key(name: str) -> str:
    # Keep it filesystem-safe (Windows 7 friendly)
    name = (name or '').strip()
    name = re.sub(r'[^0-9A-Za-zА-Яа-я _.-]+', '_', name)
    name = re.sub(r'\s+', ' ', name).strip()
    name = name.replace('..', '.')
    if len(name) > 64:
        name = name[:64].strip()
    return name

def _order_path_by_key(key: str) -> str:
    _ensure_orders_dir()
    return os.path.join(ORDERS_DIR, f'{key}.json')


def load_saved_order() -> List[str]:
    try:
        if os.path.isfile(ORDER_FILE):
            import json
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
        import json
        with open(ORDER_FILE, "w", encoding="utf-8") as f:
            json.dump({"order": order}, f, ensure_ascii=False, indent=2)
    except Exception:
        # do not crash server
        pass

def _build_state(folder: str) -> Dict[str, Any]:
    data = load_test(folder)
    t_list = [r["t_ms"] for r in data["rows"]]
    return {"loaded": True, "folder": folder, "data": data, "t_list": t_list}

def _channel_to_dict(ch: ChannelInfo) -> Dict[str, str]:
    return {"code": ch.code, "name": ch.name, "unit": ch.unit, "label": ch.label}

def _summary(data: Dict[str, Any]) -> Dict[str, Any]:
    rows = data["rows"]
    if not rows:
        return {"points": 0}
    t0 = rows[0]["t_ms"]
    t1 = rows[-1]["t_ms"]
    return {
        "points": len(rows),
        "start_ms": t0,
        "end_ms": t1,
        "start": datetime.fromtimestamp(t0/1000).strftime("%Y-%m-%d %H:%M:%S"),
        "end": datetime.fromtimestamp(t1/1000).strftime("%Y-%m-%d %H:%M:%S"),
    }

def _slice_by_time(t_list: List[int], start_ms: int, end_ms: int) -> Tuple[int, int]:
    if start_ms > end_ms:
        start_ms, end_ms = end_ms, start_ms
    i0 = bisect_left(t_list, start_ms)
    i1 = bisect_right(t_list, end_ms)
    return i0, i1


@app.after_request
def add_no_cache_headers(resp):
    # Disable aggressive caching so browser always picks up latest JS/CSS after updates
    try:
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        resp.headers["Pragma"] = "no-cache"
        resp.headers["Expires"] = "0"
    except Exception:
        pass
    return resp

@app.route("/")
def index():
    return render_template("index.html", port=APP_PORT)

@app.route("/api/pick_folder", methods=["POST"])
def pick_folder():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes("-topmost", 1)
        folder = filedialog.askdirectory(title="Выберите папку с тестом (где Prova*.dbf)")
        root.destroy()
        return jsonify({"ok": True, "folder": folder or ""})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e), "folder": ""})

@app.route("/api/load", methods=["POST"])
def api_load():
    body = request.get_json(force=True, silent=True) or {}
    folder = (body.get("folder") or "").strip()
    if not folder:
        return jsonify({"ok": False, "error": "Путь к папке пустой"})

    if not os.path.isdir(folder):
        return jsonify({"ok": False, "error": "Папка не существует: " + folder})

    try:
        st = _build_state(folder)
        STATE.update(st)
        data = STATE["data"]
        channels: Dict[str, ChannelInfo] = data["channels"]
        cols = data["cols"]

        ch_list = []
        for c in cols:
            if c in channels:
                ch_list.append(_channel_to_dict(channels[c]))
            else:
                ch_list.append({"code": c, "name": "", "unit": "", "label": c})

        return jsonify({
            "ok": True,
            "folder": data["root"],
            "meta": data["meta"],
            "summary": _summary(data),
            "channels": ch_list,
            "file_order": [c.get("code") for c in ch_list],
            "saved_order": load_saved_order(),
        })
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})

@app.route("/api/series", methods=["GET"])
def api_series():
    if not STATE.get("loaded"):
        return jsonify({"ok": False, "error": "Данные не загружены"})

    data = STATE["data"]
    rows: List[Dict[str, Any]] = data["rows"]
    t_list: List[int] = STATE["t_list"]

    channels = request.args.get("channels", "")
    ch = [c for c in channels.split(",") if c.strip()]
    if not ch:
        return jsonify({"ok": False, "error": "Не выбраны каналы"})

    try:
        start_ms = int(float(request.args.get("start_ms", rows[0]["t_ms"])))
        end_ms = int(float(request.args.get("end_ms", rows[-1]["t_ms"])))
    except Exception:
        start_ms = rows[0]["t_ms"]
        end_ms = rows[-1]["t_ms"]

    try:
        step_i = max(1, int(request.args.get("step", "1")))
    except Exception:
        step_i = 1

    i0, i1 = _slice_by_time(t_list, start_ms, end_ms)
    sliced = rows[i0:i1:step_i]
    t = [r["t_ms"] for r in sliced]

    series = {}
    for code in ch:
        vals = []
        for r in sliced:
            v = r.get(code)
            if v is None:
                vals.append(None)
            else:
                try:
                    vals.append(float(v))
                except Exception:
                    vals.append(None)
        series[code] = vals

    return jsonify({"ok": True, "t_ms": t, "series": series})


@app.route("/api/save_order", methods=["POST"])
def api_save_order():
    body = request.get_json(force=True, silent=True) or {}
    order = body.get("order")
    if not isinstance(order, list):
        return jsonify({"ok": False, "error": "order must be a list"}), 400
    seen = set()
    normalized = []
    for x in order:
        s = str(x)
        if s not in seen:
            seen.add(s)
            normalized.append(s)
    save_order(normalized)
    return jsonify({"ok": True, "saved": len(normalized)})



@app.route("/api/orders_list", methods=["GET"])
def api_orders_list():
    """Return list of saved named orders."""
    try:
        _ensure_orders_dir()
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
                # ignore broken files
                continue
        items.sort(key=lambda x: x.get('name','').lower())
        return jsonify({"ok": True, "orders": items})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e), "orders": []})


@app.route("/api/orders_save", methods=["POST"])
def api_orders_save():
    """Save a named order."""
    body = request.get_json(force=True, silent=True) or {}
    name = str(body.get('name') or '').strip()
    order = body.get('order')
    if not name:
        return jsonify({"ok": False, "error": "Имя не задано"}), 400
    if not isinstance(order, list):
        return jsonify({"ok": False, "error": "order must be a list"}), 400

    key = _sanitize_order_key(name)
    if not key:
        return jsonify({"ok": False, "error": "Имя некорректное"}), 400

    seen = set()
    normalized = []
    for x in order:
        s = str(x)
        if s and s not in seen:
            seen.add(s)
            normalized.append(s)

    payload = {
        "name": name,
        "key": key,
        "order": normalized,
        "saved_at": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    }
    try:
        path = _order_path_by_key(key)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        return jsonify({"ok": True, "key": key, "saved": len(normalized)})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/orders_load", methods=["GET"])
def api_orders_load():
    """Load a named order by key (or by name)."""
    key = str(request.args.get('key') or '').strip()
    name = str(request.args.get('name') or '').strip()
    if not key and name:
        key = _sanitize_order_key(name)
    if not key:
        return jsonify({"ok": False, "error": "Не задан ключ (key)"}), 400

    try:
        path = _order_path_by_key(key)
        if not os.path.isfile(path):
            return jsonify({"ok": False, "error": "Сохранённый порядок не найден: " + key}), 404
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        order = data.get('order')
        if not isinstance(order, list):
            order = []
        # normalize
        out = []
        seen = set()
        for x in order:
            s = str(x)
            if s and s not in seen:
                seen.add(s)
                out.append(s)
        return jsonify({"ok": True, "key": key, "name": str(data.get('name') or key), "order": out})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
def _export_csv(rows: List[Dict[str, Any]], channels: List[str]) -> bytes:
    import csv
    out = io.StringIO()
    writer = csv.writer(out, delimiter=";")
    writer.writerow(["timestamp"] + channels)
    for r in rows:
        ts = datetime.fromtimestamp(r["t_ms"]/1000).strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        writer.writerow([ts] + [r.get(c) if r.get(c) is not None else "" for c in channels])
    return out.getvalue().encode("utf-8-sig")

def _export_xlsx(rows: List[Dict[str, Any]], channels: List[str]) -> bytes:
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "data"
    ws.append(["timestamp"] + channels)
    for r in rows:
        ts = datetime.fromtimestamp(r["t_ms"]/1000).strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        ws.append([ts] + [r.get(c) for c in channels])
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()

@app.route("/api/export", methods=["GET"])
def api_export():
    if not STATE.get("loaded"):
        return jsonify({"ok": False, "error": "Данные не загружены"}), 400

    data = STATE["data"]
    rows: List[Dict[str, Any]] = data["rows"]
    t_list: List[int] = STATE["t_list"]

    fmt = (request.args.get("format", "csv") or "csv").lower()
    channels = request.args.get("channels", "")
    ch = [c for c in channels.split(",") if c.strip()]
    if not ch:
        return jsonify({"ok": False, "error": "Не выбраны каналы"}), 400

    try:
        start_ms = int(float(request.args.get("start_ms", rows[0]["t_ms"])))
        end_ms = int(float(request.args.get("end_ms", rows[-1]["t_ms"])))
    except Exception:
        start_ms = rows[0]["t_ms"]
        end_ms = rows[-1]["t_ms"]

    try:
        step_i = max(1, int(request.args.get("step", "1")))
    except Exception:
        step_i = 1

    i0, i1 = _slice_by_time(t_list, start_ms, end_ms)
    sliced = rows[i0:i1:step_i]

    if fmt == "xlsx":
        payload = _export_xlsx(sliced, ch)
        return send_file(
            io.BytesIO(payload),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="export.xlsx",
        )
    else:
        payload = _export_csv(sliced, ch)
        return send_file(
            io.BytesIO(payload),
            mimetype="text/csv",
            as_attachment=True,
            download_name="export.csv",
        )



def _nearest_index(t_list: List[int], target_ms: int) -> int:
    """Return nearest index in sorted t_list to target_ms."""
    if not t_list:
        return -1
    i = bisect_left(t_list, target_ms)
    if i <= 0:
        return 0
    if i >= len(t_list):
        return len(t_list) - 1
    a = t_list[i-1]
    b = t_list[i]
    if abs(target_ms - a) <= abs(b - target_ms):
        return i-1
    return i


def _resolve_channel(cols: List[str], key: str) -> str:
    """Resolve sensor code from available columns by semantic key."""
    # special channels without prefix
    if key in ("T-sie", "UR-sie"):
        return key if key in cols else ""

    # common suffix mapping (prefer A- then C-)
    candidates = [f"A-{key}", f"C-{key}", key]
    for c in candidates:
        if c in cols:
            return c

    # fallback: find any ending with -key
    suf = "-" + key
    for c in cols:
        if c.endswith(suf):
            return c
    return ""


# ---------------- Template export (fill template.xlsx preserving formatting) ----------------

@app.route("/api/export_template", methods=["GET"])
def api_export_template():
    """Заполнить template.xlsx, сохранив форматирование/цвета/формулы.

    Правила заполнения (как вы описали):
      - Строки данных начинаются с 4-й строки.
      - Колонка B: время испытания = 00:00:00 + N*20 секунд.
      - Колонка C: текущее время берём из файла (только время суток, без даты).
      - Колонки D..T: сопоставляем с показаниями датчиков.

    Важно: в шаблоне есть формулы (например U..Y). Их сохраняем и протягиваем вниз,
    а значения в остальных «сырых» колонках очищаем, чтобы не оставались примерные данные.
    """

    if not STATE.get("loaded"):
        return jsonify({"ok": False, "error": "Данные не загружены"}), 400

    data = STATE["data"]
    rows: List[Dict[str, Any]] = data["rows"]
    t_list: List[int] = STATE["t_list"]
    cols: List[str] = data.get("cols") or []

    if not rows:
        return jsonify({"ok": False, "error": "Нет строк данных"}), 400

    try:
        start_ms = int(float(request.args.get("start_ms", rows[0]["t_ms"])))
        end_ms = int(float(request.args.get("end_ms", rows[-1]["t_ms"])))
    except Exception:
        start_ms = rows[0]["t_ms"]
        end_ms = rows[-1]["t_ms"]

    if start_ms > end_ms:
        start_ms, end_ms = end_ms, start_ms

    # Find first sample within range
    i0 = bisect_left(t_list, start_ms)
    if i0 >= len(t_list):
        return jsonify({"ok": False, "error": "Диапазон вне данных"}), 400
    t0 = t_list[i0]
    if t0 > end_ms:
        return jsonify({"ok": False, "error": "Пустой диапазон"}), 400

    # Build 20s grid starting from the first real sample in the selected range
    grid_ms: List[int] = []
    g = t0
    step_ms = 20000
    while g <= end_ms:
        grid_ms.append(g)
        g += step_ms

    # Map each grid point to nearest sample
    idxs: List[int] = []
    for g in grid_ms:
        idx = _nearest_index(t_list, g)
        if idx < 0 or abs(t_list[idx] - g) > 30000:
            idxs.append(-1)
        else:
            idxs.append(idx)

    # Respect selected channels (optional query param "channels" like other exports)
    # If provided, only those channels will be written to the template.
    ch_arg = (request.args.get("channels") or "").strip()
    selected = set([c.strip() for c in ch_arg.split(",") if c.strip()])
    # If the browser did not send selected channels, do NOT export everything silently.
    if not selected:
        return jsonify({"ok": False, "error": "Не получены выбранные каналы (похоже, браузер использует старый app.js). Нажми Ctrl+F5 и попробуй снова."}), 400

    # Debug: log what we received (goes to run log)
    try:
        print("[TEMPLATE] start_ms=", start_ms, "end_ms=", end_ms, "channels=", ",".join(sorted(selected)))
    except Exception:
        pass
    def resolve_for_selection(key: str) -> str:
        """Resolve a channel code for a given key.

        If selection is provided, try to pick a matching channel that is selected.
        Otherwise fall back to the standard preference (A-*, C-*, plain).
        """
        matches = []
        # Exact match
        if key in cols:
            matches.append(key)
        # Suffix match (A-T1, C-T1, etc.)
        suf = f"-{key}"
        for c in cols:
            if c.endswith(suf) and c not in matches:
                matches.append(c)

        if not matches:
            return ""

        if selected:
            for c in matches:
                if c in selected:
                    return c
            # Selection is provided but none of the matching channels were selected
            return ""

        # Prefer A- then C- then the first remaining
        for pref in ("A-", "C-"):
            for c in matches:
                if c.startswith(pref):
                    return c
        return matches[0]

    # Resolve channel -> template column mapping (D..T)
    key_to_col = {
        "Pc": 4,
        "Pe": 5,
        "T-sie": 6,
        "UR-sie": 7,
        "Tc": 8,
        "Te": 9,
        "T1": 10,
        "T2": 11,
        "T3": 12,
        "T4": 13,
        "T5": 14,
        "T6": 15,
        "T7": 16,
        "I": 17,
        "F": 18,
        "V": 19,
        "W": 20,
    }
    key_to_code = {k: resolve_for_selection(k) for k in key_to_col.keys()}

    template_path = TEMPLATE_FILE
    if not os.path.isfile(template_path):
        return jsonify({"ok": False, "error": "Не найден template.xlsx рядом с server.py"}), 400

    try:
        from openpyxl import load_workbook
        from openpyxl.utils import get_column_letter
        from openpyxl.formula.translate import Translator
        from copy import copy
        import datetime as _dt

        wb = load_workbook(template_path)
        ws = wb.active

        start_row = 4
        pattern_row = start_row
        max_col = ws.max_column

        # Find formula columns in the pattern row and remember pattern formulas
        formula_cols = set()
        pattern_formulas = {}
        for c in range(1, max_col + 1):
            v = ws.cell(row=pattern_row, column=c).value
            if isinstance(v, str) and v.startswith('='):
                formula_cols.add(c)
                pattern_formulas[c] = v

        def ensure_row_with_style(r: int) -> None:
            # create row cells & copy style from pattern row (for new rows)
            if r <= ws.max_row:
                return
            for c in range(1, max_col + 1):
                ws.cell(row=r, column=c)
            try:
                ws.row_dimensions[r].height = ws.row_dimensions[pattern_row].height
            except Exception:
                pass
            for c in range(1, max_col + 1):
                s = ws.cell(row=pattern_row, column=c)
                d = ws.cell(row=r, column=c)
                d._style = copy(s._style)
                d.number_format = s.number_format
                d.font = copy(s.font)
                d.border = copy(s.border)
                d.fill = copy(s.fill)
                d.alignment = copy(s.alignment)
                d.protection = copy(s.protection)

        needed_last_row = start_row + len(idxs) - 1
        for r in range(ws.max_row + 1, needed_last_row + 1):
            ensure_row_with_style(r)

        # Prepare rows: keep formulas (translated), clear everything else
        for j in range(len(idxs)):
            r = start_row + j
            for c in range(1, max_col + 1):
                cell = ws.cell(row=r, column=c)
                if c in formula_cols:
                    src_formula = pattern_formulas.get(c)
                    if src_formula:
                        src_addr = f"{get_column_letter(c)}{pattern_row}"
                        dst_addr = f"{get_column_letter(c)}{r}"
                        try:
                            cell.value = Translator(src_formula, origin=src_addr).translate_formula(dst_addr)
                        except Exception:
                            cell.value = src_formula
                else:
                    cell.value = None

        # Write data
        for j, idx in enumerate(idxs):
            r = start_row + j
            ws.cell(row=r, column=2).value = _dt.timedelta(seconds=j * 20)

            if idx >= 0:
                tms = rows[idx]["t_ms"]
                dt = datetime.fromtimestamp(tms/1000)
                ws.cell(row=r, column=3).value = dt.time()

                for key, colnum in key_to_col.items():
                    code = key_to_code.get(key) or ""
                    if not code:
                        continue
                    if selected and code not in selected:
                        # Only write selected channels
                        continue
                    v = rows[idx].get(code)
                    if v is None:
                        ws.cell(row=r, column=colnum).value = None
                    else:
                        try:
                            ws.cell(row=r, column=colnum).value = float(v)
                        except Exception:
                            ws.cell(row=r, column=colnum).value = None
            else:
                ws.cell(row=r, column=3).value = None

        # Clear a bit below to avoid leftover sample values
        clear_from = start_row + len(idxs)
        clear_to = min(ws.max_row, clear_from + 300)
        for r in range(clear_from, clear_to + 1):
            for c in range(1, max_col + 1):
                if c in formula_cols:
                    continue
                ws.cell(row=r, column=c).value = None

        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)

        fn = f"template_filled_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return _send_file_compat(
            bio,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            fn,
        )

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


def _open_browser():
    try:
        webbrowser.open(f"http://{APP_HOST}:{APP_PORT}/", new=1)
    except Exception:
        pass

if __name__ == "__main__":
    threading.Timer(0.8, _open_browser).start()
    _ensure_orders_dir()
    # threaded=False — чтобы tkinter "Обзор..." работал предсказуемо
    app.run(host=APP_HOST, port=APP_PORT, debug=False, threaded=False)
