# lemure_reader.py
# Чтение тестов LeMuRe/Project Engineering (Prova*.dbf + Set/Canali.def)
# Реализовано без pandas/numpy — подходит для старых систем (в т.ч. Windows 7 + Python 3.8)

from __future__ import annotations
import os, re
from dataclasses import dataclass
from datetime import datetime, date
from typing import Dict, List, Tuple, Optional, Any


@dataclass
class ChannelInfo:
    code: str
    name: str
    unit: str

    @property
    def label(self) -> str:
        u = f" ({self.unit})" if self.unit else ""
        n = f" — {self.name}" if self.name else ""
        return f"{self.code}{n}{u}"


def _safe_read_text(path: str) -> str:
    for enc in ("utf-8", "cp1251", "cp1252", "latin1"):
        try:
            with open(path, "r", encoding=enc, errors="strict") as f:
                return f.read()
        except Exception:
            continue
    with open(path, "r", encoding="latin1", errors="ignore") as f:
        return f.read()


def parse_canali_def(canali_def_path: str) -> Dict[str, ChannelInfo]:
    if not os.path.isfile(canali_def_path):
        return {}
    text = _safe_read_text(canali_def_path)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    if not lines:
        return {}

    if re.fullmatch(r"\d+", lines[0]):
        lines = lines[1:]

    out: Dict[str, ChannelInfo] = {}
    for ln in lines:
        parts = [p.strip() for p in ln.split(";")]
        if len(parts) < 3:
            continue
        code, name, unit = parts[0], parts[1], parts[2]
        out[code] = ChannelInfo(code=code, name=name, unit=unit)
    return out


def parse_prova_dat(prova_dat_path: str) -> Dict[str, str]:
    if not os.path.isfile(prova_dat_path):
        return {}
    text = _safe_read_text(prova_dat_path)
    meta: Dict[str, str] = {}
    for ln in text.splitlines():
        ln = ln.strip()
        if not ln or ";" not in ln:
            continue
        k, v = ln.split(";", 1)
        meta[k.strip()] = v.strip()
    return meta


# ---------------- DBF reader ----------------

@dataclass
class DbfField:
    name: str
    ftype: str
    length: int
    decimals: int


@dataclass
class DbfHeader:
    version: int
    last_update_ymd: Tuple[int, int, int]
    records: int
    header_len: int
    record_len: int
    fields: List[DbfField]


def _read_dbf_header(buf: bytes) -> DbfHeader:
    if len(buf) < 32:
        raise ValueError("DBF слишком короткий")
    version = buf[0]
    yy, mm, dd = buf[1], buf[2], buf[3]
    records = int.from_bytes(buf[4:8], "little", signed=False)
    header_len = int.from_bytes(buf[8:10], "little", signed=False)
    record_len = int.from_bytes(buf[10:12], "little", signed=False)

    fields: List[DbfField] = []
    off = 32
    while off + 32 <= header_len:
        desc = buf[off:off + 32]
        if desc[0] == 0x0D:
            break
        name = desc[0:11].split(b"\x00", 1)[0].decode("ascii", errors="ignore").strip()
        ftype = chr(desc[11])
        length = desc[16]
        decimals = desc[17]
        fields.append(DbfField(name=name, ftype=ftype, length=length, decimals=decimals))
        off += 32

    return DbfHeader(
        version=version,
        last_update_ymd=(yy, mm, dd),
        records=records,
        header_len=header_len,
        record_len=record_len,
        fields=fields,
    )


def _parse_dbf_value(raw: bytes, f: DbfField):
    s = raw.decode("ascii", errors="ignore").strip()
    if f.ftype == "N":  # numeric
        if s == "" or s == ".":
            return None
        try:
            return float(s)
        except Exception:
            return None
    if f.ftype == "D":  # YYYYMMDD
        if len(s) != 8 or not s.isdigit():
            return None
        try:
            y = int(s[0:4]); m = int(s[4:6]); d = int(s[6:8])
            return date(y, m, d)
        except Exception:
            return None
    return s if s != "" else None


def read_dbf_rows(dbf_path: str) -> Tuple[DbfHeader, List[Dict[str, Any]]]:
    with open(dbf_path, "rb") as f:
        buf = f.read()

    hdr = _read_dbf_header(buf)
    fields = hdr.fields

    start = hdr.header_len
    rec_len = hdr.record_len

    rows: List[Dict[str, Any]] = []
    pos = start
    for _ in range(hdr.records):
        rec = buf[pos:pos + rec_len]
        pos += rec_len
        if not rec or len(rec) < rec_len:
            break
        if rec[0:1] == b"*":
            continue
        off = 1
        row: Dict[str, Any] = {}
        for fdef in fields:
            raw = rec[off:off + fdef.length]
            off += fdef.length
            row[fdef.name] = _parse_dbf_value(raw, fdef)
        rows.append(row)

    return hdr, rows


# ------------- Test loader -------------

_TIME_COLS = {"Data", "Ore", "Minuti", "Secondi", "mSecondi"}


def _find_test_root(folder: str) -> str:
    folder = os.path.abspath(folder)
    if os.path.isdir(folder):
        try:
            if any(re.match(r"Prova\d+\.dbf$", f, re.IGNORECASE) for f in os.listdir(folder)):
                return folder
        except Exception:
            pass

    candidates = []
    for root, _, files in os.walk(folder):
        if any(re.match(r"Prova\d+\.dbf$", f, re.IGNORECASE) for f in files):
            candidates.append(root)
    if not candidates:
        raise FileNotFoundError("Не нашёл Prova*.dbf в выбранной папке")
    candidates.sort(key=lambda p: len(p))
    return candidates[0]


def _dbf_sort_key(path: str) -> int:
    m = re.search(r"Prova(\d+)\.dbf$", os.path.basename(path), re.IGNORECASE)
    return int(m.group(1)) if m else 0


def load_test(folder: str):
    root = _find_test_root(folder)

    set_dir = os.path.join(root, "Set")
    channels = parse_canali_def(os.path.join(set_dir, "Canali.def"))

    meta = {}
    for fname in os.listdir(root):
        if re.match(r"Prova\d+\.dat$", fname, re.IGNORECASE):
            meta = parse_prova_dat(os.path.join(root, fname))
            break

    dbfs = []
    for fname in os.listdir(root):
        if re.match(r"Prova\d+\.dbf$", fname, re.IGNORECASE):
            dbfs.append(os.path.join(root, fname))
    if not dbfs:
        raise FileNotFoundError("В папке теста нет Prova*.dbf")
    dbfs.sort(key=_dbf_sort_key)

    rows_all: List[Dict[str, Any]] = []
    for dbf in dbfs:
        _, rows = read_dbf_rows(dbf)
        rows_all.extend(rows)

    data_rows = []
    for r in rows_all:
        d = r.get("Data")
        if not isinstance(d, date):
            continue
        hh = int(float(r.get("Ore") or 0))
        mm = int(float(r.get("Minuti") or 0))
        ss = int(float(r.get("Secondi") or 0))
        mss = int(float(r.get("mSecondi") or 0))
        ts = datetime(d.year, d.month, d.day, hh, mm, ss, mss * 1000)

        row2 = {"t_ms": int(ts.timestamp() * 1000)}
        for k, v in r.items():
            if k in _TIME_COLS:
                continue
            if v is None:
                row2[k] = None
            elif isinstance(v, (int, float)):
                row2[k] = float(v)
            else:
                try:
                    row2[k] = float(str(v).replace(",", "."))
                except Exception:
                    row2[k] = str(v)
        data_rows.append(row2)

    data_rows.sort(key=lambda x: x["t_ms"])
    cols = sorted({k for r in data_rows for k in r.keys()} - {"t_ms"})

    return {
        "root": root,
        "meta": meta,
        "channels": channels,  # dict[str, ChannelInfo]
        "rows": data_rows,     # list[dict], t_ms + values
        "cols": cols,
    }
