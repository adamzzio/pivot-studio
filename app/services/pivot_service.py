"""Kalkulasi pivot table berbasis pandas.

Definisi agregasi:
- SUM, AVERAGE, MEDIAN : hanya field numerik, nilai kosong diabaikan.
- VAR, STD             : populasi (ddof=0).
- COUNT                : jumlah SEMUA baris pada grup.
- COUNT_NON_NULL       : jumlah nilai yang tidak kosong.
- COUNT_DISTINCT       : jumlah nilai unik (kosong diabaikan).
- MODE                 : nilai paling sering; jika seri dipilih yang terkecil;
                         kosong jika semua nilai unik.
"""
from __future__ import annotations

import json
import re
import threading
from collections import OrderedDict

import numpy as np
import pandas as pd

from app.utils import AppError

from . import excel_service as xs

AGG_LABELS = {
    "SUM": "SUM",
    "AVERAGE": "AVERAGE",
    "COUNT": "COUNT",
    "COUNT_NON_NULL": "COUNT NON-NULL",
    "COUNT_DISTINCT": "COUNT DISTINCT",
    "VAR": "VAR",
    "STD": "STD",
    "MEDIAN": "MEDIAN",
    "MODE": "MODE",
}
NUMERIC_AGGS = {"SUM", "AVERAGE", "VAR", "STD", "MEDIAN"}
TYPE_ID = {"number": "angka", "text": "teks", "date": "tanggal"}


# ---------------------------------------------------------------------------
# Parsing konfigurasi
# ---------------------------------------------------------------------------


def _str_list(v, limit=8, name="field"):
    if v is None:
        return []
    if not isinstance(v, list) or not all(isinstance(x, str) for x in v):
        raise AppError(f"Konfigurasi {name} tidak valid.", 400)
    if len(v) > limit:
        raise AppError(f"Maksimal {limit} field pada {name}.", 400)
    return v


def parse_config(payload) -> dict:
    if not isinstance(payload, dict):
        raise AppError("Konfigurasi pivot tidak valid.", 400)
    src = payload.get("source") or {}
    file_id, sheet = src.get("file_id"), src.get("sheet")
    if not isinstance(file_id, str) or not isinstance(sheet, str) or not sheet:
        raise AppError("Sumber data belum dipilih.", 400)

    rows = _str_list(payload.get("rows"), name="Rows")
    cols = _str_list(payload.get("columns"), name="Columns")

    values = []
    raw_values = payload.get("values") or []
    if not isinstance(raw_values, list) or len(raw_values) > 20:
        raise AppError("Konfigurasi Values tidak valid (maksimal 20 metrik).", 400)
    for v in raw_values:
        if not isinstance(v, dict):
            raise AppError("Konfigurasi Values tidak valid.", 400)
        agg = str(v.get("agg", "")).upper()
        if agg not in AGG_LABELS:
            raise AppError(f"Fungsi agregasi '{agg}' tidak dikenali.", 400)
        values.append(
            {"field": str(v.get("field", "")), "agg": agg, "label": str(v.get("label") or "").strip()[:80]}
        )

    filters = payload.get("filters") or []
    if not isinstance(filters, list) or len(filters) > 30 or not all(isinstance(f, dict) for f in filters):
        raise AppError("Konfigurasi Filters tidak valid.", 400)
    clean_filters = []
    for f in filters:
        clean_filters.append(
            {
                "field": str(f.get("field", "")),
                "kind": f.get("kind"),
                "values": [str(x) for x in f["values"]] if isinstance(f.get("values"), list) else None,
                "min": f.get("min"),
                "max": f.get("max"),
                "start": f.get("start"),
                "end": f.get("end"),
            }
        )

    types = payload.get("types") or {}
    if not isinstance(types, dict):
        types = {}
    types = {str(k): v for k, v in types.items() if v in xs.VALID_TYPES}

    o = payload.get("options") or {}
    rs = o.get("row_sort") or {}
    cs = o.get("col_sort") or {}
    options = {
        "subtotals": bool(o.get("subtotals", True)),
        "grand_total": bool(o.get("grand_total", True)),
        "row_sort": {
            "by": "value" if rs.get("by") == "value" else "label",
            "dir": "desc" if rs.get("dir") == "desc" else "asc",
            "metric": int(rs.get("metric") or 0) if str(rs.get("metric") or 0).isdigit() else 0,
        },
        "col_sort": {"dir": "desc" if cs.get("dir") == "desc" else "asc"},
    }
    return {
        "source": {"file_id": file_id, "sheet": sheet},
        "types": types,
        "rows": rows,
        "columns": cols,
        "values": values,
        "filters": clean_filters,
        "options": options,
    }


def _validate(cfg: dict, columns: list[str]) -> None:
    cset = set(columns)
    used = cfg["rows"] + cfg["columns"] + [v["field"] for v in cfg["values"]] + [f["field"] for f in cfg["filters"]]
    for f in used:
        if f not in cset:
            raise AppError(f"Kolom '{f}' tidak ditemukan pada sumber data.", 422)
    if len(set(cfg["rows"])) != len(cfg["rows"]) or len(set(cfg["columns"])) != len(cfg["columns"]):
        raise AppError("Field yang sama tidak boleh muncul dua kali dalam satu area.", 422)
    both = set(cfg["rows"]) & set(cfg["columns"])
    if both:
        raise AppError(f"Field '{sorted(both)[0]}' tidak boleh berada di Rows dan Columns sekaligus.", 422)


def metric_labels(values: list[dict]) -> list[str]:
    out, seen = [], set()
    for v in values:
        base = v["label"] or f"{AGG_LABELS[v['agg']]} {v['field']}"
        name, k = base, 1
        while name in seen:
            k += 1
            name = f"{base} ({k})"
        seen.add(name)
        out.append(name)
    return out


def describe_filters(cfg: dict) -> list[str]:
    out = []
    for f in cfg["filters"]:
        if f["kind"] == "values" and f["values"]:
            vals = f["values"]
            txt = ", ".join(vals[:5]) + (f" (+{len(vals) - 5} lainnya)" if len(vals) > 5 else "")
            out.append(f"{f['field']} = {txt}")
        elif f["kind"] == "range":
            out.append(f"{f['field']}: {f['min'] if f['min'] not in (None, '') else '...'} s/d {f['max'] if f['max'] not in (None, '') else '...'}")
        elif f["kind"] == "date_range":
            out.append(f"{f['field']}: {f['start'] or '...'} s/d {f['end'] or '...'}")
    return out


# ---------------------------------------------------------------------------
# Filter dan sorting
# ---------------------------------------------------------------------------


def _num(v):
    if v in (None, ""):
        return None
    try:
        x = float(v)
    except (TypeError, ValueError):
        raise AppError("Nilai filter rentang harus berupa angka.", 422)
    return None if x != x else x


def _date(v):
    if v in (None, ""):
        return None
    d = pd.to_datetime(v, errors="coerce")
    if pd.isna(d):
        raise AppError("Format tanggal filter tidak valid.", 422)
    return d


def _filter_mask(typed: pd.DataFrame, types: dict, filters: list[dict]) -> pd.Series:
    mask = pd.Series(True, index=typed.index)
    for f in filters:
        field, kind = f["field"], f["kind"]
        s, t = typed[field], types[field]
        if kind == "values":
            if not f["values"]:
                raise AppError(f"Filter '{field}': pilih minimal satu nilai.", 422)
            mask &= xs.label_series(s, t).isin(set(f["values"]))
        elif kind == "range":
            if t != "number":
                raise AppError(f"Filter rentang hanya untuk kolom angka ('{field}' bertipe {TYPE_ID[t]}).", 422)
            lo, hi = _num(f["min"]), _num(f["max"])
            if lo is None and hi is None:
                raise AppError(f"Filter '{field}': isi batas minimum atau maksimum.", 422)
            if lo is not None:
                mask &= s >= lo
            if hi is not None:
                mask &= s <= hi
        elif kind == "date_range":
            if t != "date":
                raise AppError(f"Filter tanggal hanya untuk kolom tanggal ('{field}' bertipe {TYPE_ID[t]}).", 422)
            start, end = _date(f["start"]), _date(f["end"])
            if start is None and end is None:
                raise AppError(f"Filter '{field}': isi tanggal awal atau akhir.", 422)
            if start is not None:
                mask &= s >= start
            if end is not None:
                mask &= s < end + pd.Timedelta(days=1)
        else:
            raise AppError("Jenis filter tidak dikenali.", 422)
    return mask


def _natural(s: str):
    return [int(p) if i % 2 else p.casefold() for i, p in enumerate(re.split(r"(\d+)", s))]


def _sortkey_map(labels: pd.Series, typed: pd.Series, t: str) -> dict:
    pair = pd.DataFrame({"l": labels.values, "v": typed.values}).drop_duplicates("l")
    out = {}
    for lab, v in zip(pair["l"], pair["v"]):
        if lab == xs.NULL_LABEL:
            continue
        if t == "number":
            out[lab] = float(v)
        elif t == "date":
            out[lab] = pd.Timestamp(v).value
        else:
            out[lab] = _natural(lab)
    return out


def _k(keymap: dict, label: str):
    return (0, keymap[label]) if label in keymap else (1, 0)


def _sorted_labels(labels, keymap, reverse):
    known = sorted([l for l in labels if l in keymap], key=lambda l: keymap[l], reverse=reverse)
    return known + [l for l in labels if l not in keymap]


# ---------------------------------------------------------------------------
# Agregasi
# ---------------------------------------------------------------------------


def _make_mode(flags: dict):
    def mode_fn(s: pd.Series):
        s = s.dropna()
        if s.empty:
            return np.nan
        vc = s.value_counts()
        top = vc.max()
        if top <= 1:
            flags["no_mode"] = True
            return np.nan
        winners = list(vc[vc == top].index)
        if len(winners) > 1:
            flags["multi_mode"] = True
        try:
            return min(winners)
        except TypeError:
            return sorted(map(str, winners))[0]

    return mode_fn


def _run_agg(g, col: str, agg: str, mode_fn):
    if agg == "COUNT":
        return g.size()
    s = g[col]
    if agg == "SUM":
        return s.sum(min_count=1)
    if agg == "AVERAGE":
        return s.mean()
    if agg == "COUNT_NON_NULL":
        return s.count()
    if agg == "COUNT_DISTINCT":
        return s.nunique(dropna=True)
    if agg == "VAR":
        return s.var(ddof=0)
    if agg == "STD":
        return s.std(ddof=0)
    if agg == "MEDIAN":
        return s.median()
    if agg == "MODE":
        return s.agg(mode_fn)
    raise AppError(f"Fungsi agregasi '{agg}' tidak dikenali.", 400)


def _clean(x):
    if x is None:
        return None
    if isinstance(x, (np.integer,)):
        return int(x)
    if isinstance(x, (float, np.floating)):
        return None if (x != x or x in (float("inf"), float("-inf"))) else float(x)
    if isinstance(x, str):
        return x
    try:
        if pd.isna(x):
            return None
    except (TypeError, ValueError):
        pass
    return x if isinstance(x, (int, bool)) else str(x)


def _metric_series(series: pd.Series, t: str, v: dict) -> pd.Series:
    agg, field = v["agg"], v["field"]
    if agg in NUMERIC_AGGS and t != "number":
        raise AppError(
            f"{AGG_LABELS[agg]} memerlukan field numerik, tetapi kolom '{field}' bertipe {TYPE_ID[t]}. "
            "Ubah tipe kolom menjadi angka atau gunakan COUNT, COUNT DISTINCT, atau MODE.",
            422,
        )
    if agg == "MODE" and t == "date":
        return xs.label_series(series, t).where(series.notna())
    return series


# ---------------------------------------------------------------------------
# Perhitungan utama
# ---------------------------------------------------------------------------


def compute_pivot(df: pd.DataFrame, inferred: dict, cfg: dict, max_cells: int | None = None) -> dict:
    rows, cols, values, filters = cfg["rows"], cfg["columns"], cfg["values"], cfg["filters"]
    _validate(cfg, list(df.columns))
    if not values:
        raise AppError("Tambahkan minimal satu field ke area Values.", 422)

    needed = list(dict.fromkeys(rows + cols + [v["field"] for v in values] + [f["field"] for f in filters]))
    typed, types, notes = xs.apply_types(df, inferred, cfg["types"], needed)
    total_records = len(df)
    mask = _filter_mask(typed, types, filters)
    if not bool(mask.all()):
        typed = typed[mask]
    records = len(typed)

    labels = {f: xs.label_series(typed[f], types[f]) for f in rows + cols}
    keymaps = {f: _sortkey_map(labels[f], typed[f], types[f]) for f in rows + cols}

    mlabels = metric_labels(values)
    metrics = [{"label": mlabels[i], "agg": v["agg"], "field": v["field"]} for i, v in enumerate(values)]
    M = len(metrics)

    work = {}
    for i, v in enumerate(values):
        work[f"m{i}"] = _metric_series(typed[v["field"]], types[v["field"]], v)
    row_keys, col_keys = [], []
    for j, f in enumerate(rows):
        work[f"r{j}"] = labels[f]
        row_keys.append(f"r{j}")
    for j, f in enumerate(cols):
        work[f"c{j}"] = labels[f]
        col_keys.append(f"c{j}")
    frame = pd.DataFrame(work, index=typed.index)

    flags = {"multi_mode": False, "no_mode": False}
    mode_fn = _make_mode(flags)

    def grid(row_n: int, with_cols: bool) -> dict:
        rk = row_keys[:row_n]
        keys = rk + (col_keys if with_cols else [])
        if keys:
            g = frame.groupby(keys, sort=False, dropna=False)
        else:
            g = frame.groupby(np.zeros(len(frame), dtype=int), sort=False)
        res = pd.concat(
            [_run_agg(g, f"m{i}", m["agg"], mode_fn).rename(f"m{i}") for i, m in enumerate(metrics)], axis=1
        )
        out, nr = {}, len(rk)
        for idx, vals in zip(res.index, res.itertuples(index=False, name=None)):
            tup = () if not keys else ((idx,) if len(keys) == 1 else tuple(idx))
            out[(tup[:nr], tup[nr:])] = [_clean(x) for x in vals]
        return out

    R = len(rows)
    opts = cfg["options"]
    rs, cs = opts["row_sort"], opts["col_sort"]
    show_sub = opts["subtotals"] and R > 1
    show_grand = opts["grand_total"]
    s_levels = {R} | (set(range(1, R)) if show_sub else set()) | ({0} if (show_grand or R == 0) else set())
    sort_value = rs["by"] == "value" and R > 0
    t_levels = set(s_levels) | (set(range(1, R + 1)) if sort_value else set())

    S = {lv: grid(lv, True) for lv in sorted(s_levels)}
    if cols:
        T = {lv: grid(lv, False) for lv in sorted(t_levels)}
    else:
        T = dict(S)
        for lv in sorted(t_levels - set(S)):
            T[lv] = grid(lv, True)

    row_tuples = {k[0] for k in S[R]}
    col_paths = sorted({k[1] for k in S[R]}, key=lambda cp: tuple(_k(keymaps[cols[i]], cp[i]) for i in range(len(cols)))) if cols else [()]
    if cols and cs["dir"] == "desc":
        col_paths.reverse()
    if not col_paths:
        col_paths = [()]

    est = max(len(row_tuples), 1) * (len(col_paths) + 1) * M
    if max_cells and est > max_cells:
        raise AppError(
            f"Hasil pivot terlalu besar ({est:,} sel). Kurangi jumlah field, tambahkan filter, atau gunakan field dengan nilai lebih sedikit.",
            422,
        )

    tree: dict = {}
    for rt in row_tuples:
        node = tree
        for lab in rt:
            node = node.setdefault(lab, {})

    def sort_children(kids, prefix, level):
        field = rows[level]
        reverse = rs["dir"] == "desc"
        base = _sorted_labels(kids, keymaps[field], False)
        if rs["by"] != "value":
            return _sorted_labels(kids, keymaps[field], reverse)
        mi = min(rs["metric"], M - 1)

        def val(lab):
            cell = T[level + 1].get((prefix + (lab,), ()))
            return None if not cell else cell[mi]

        have = [l for l in base if val(l) is not None]
        none = [l for l in base if val(l) is None]
        try:
            have.sort(key=val, reverse=reverse)
        except TypeError:
            have.sort(key=lambda l: str(val(l)), reverse=reverse)
        return have + none

    entries: list[tuple[str, tuple]] = []

    def walk(node, prefix, level):
        for lab in sort_children(list(node.keys()), prefix, level):
            path = prefix + (lab,)
            if level + 1 == R:
                entries.append(("detail", path))
            else:
                walk(node[lab], path, level + 1)
                if show_sub:
                    entries.append(("subtotal", path))

    if R == 0:
        entries.append(("grand", ()))
    else:
        walk(tree, (), 0)
        if show_grand:
            entries.append(("grand", ()))

    grand_col = bool(cols) and show_grand

    def row_values(path):
        lv = len(path)
        vals = []
        for cp in col_paths:
            cell = S[lv].get((path, cp)) if lv in S else None
            vals.extend(cell if cell else [None] * M)
        if grand_col:
            cell = T[lv].get((path, ())) if lv in T else None
            vals.extend(cell if cell else [None] * M)
        return vals

    body = [{"kind": kind, "labels": list(path), "values": row_values(path)} for kind, path in entries]

    if flags["multi_mode"]:
        notes.append("MODE: beberapa nilai memiliki frekuensi tertinggi yang sama, nilai terkecil yang ditampilkan.")
    if flags["no_mode"]:
        notes.append("MODE: sel dikosongkan jika semua nilai pada grup unik (tidak ada modus).")

    return {
        "header": {
            "row_fields": rows,
            "col_fields": cols,
            "col_paths": [list(cp) for cp in col_paths],
            "metrics": metrics,
            "grand_col": grand_col,
        },
        "body": body,
        "meta": {"records": int(records), "total_records": int(total_records), "notes": notes},
    }


# ---------------------------------------------------------------------------
# Eksekusi dari request (dengan cache hasil kecil)
# ---------------------------------------------------------------------------

_results: OrderedDict = OrderedDict()
_lock = threading.Lock()


def run_pivot(sid: str, upload_dir, payload: dict, max_cells: int | None = None):
    cfg = parse_config(payload)
    path = xs.file_path(upload_dir, sid, cfg["source"]["file_id"])
    df, inferred = xs.load_frame(path, cfg["source"]["sheet"])
    key = (sid, path.stat().st_mtime_ns, json.dumps(cfg, sort_keys=True, default=str))
    with _lock:
        if key in _results:
            _results.move_to_end(key)
            return _results[key], df, cfg
    result = compute_pivot(df, inferred, cfg, max_cells)
    with _lock:
        _results[key] = result
        while len(_results) > 8:
            _results.popitem(last=False)
    return result, df, cfg


def field_values(df: pd.DataFrame, inferred: dict, overrides: dict, field: str, search: str, limit: int) -> dict:
    if field not in df.columns:
        raise AppError(f"Kolom '{field}' tidak ditemukan.", 422)
    typed, types, _ = xs.apply_types(df, inferred, overrides, [field])
    s, t = typed[field], types[field]
    labels = xs.label_series(s, t)
    out = {"type": t}
    if t == "number" and s.notna().any():
        out["min"], out["max"] = float(s.min()), float(s.max())
    if t == "date" and s.notna().any():
        out["min"], out["max"] = s.min().strftime("%Y-%m-%d"), s.max().strftime("%Y-%m-%d")
    counts = labels.value_counts()
    total_unique = int(len(counts))
    if search:
        counts = counts[counts.index.str.casefold().str.contains(search.casefold(), regex=False)]
    km = _sortkey_map(labels, s, t)
    items = sorted(counts.index, key=lambda l: _k(km, l))
    out["total_unique"] = total_unique
    out["matched"] = int(len(items))
    out["values"] = [{"label": l, "count": int(counts[l])} for l in items[:limit]]
    return out
