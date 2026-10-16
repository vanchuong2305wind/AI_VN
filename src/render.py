from __future__ import annotations

import datetime as dt

import pandas as pd

from .parser import Table


def fmt(v) -> str:
    if v is None or (isinstance(v, float) and v != v) or v is pd.NaT:
        return ""
    if isinstance(v, (pd.Timestamp, dt.datetime)):
        return v.strftime("%Y-%m-%d") if (v.hour, v.minute, v.second) == (0, 0, 0) else str(v)
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).replace("|", "/").replace("\n", " ")


def est_tokens(text: str) -> int:
    digits = sum(ch.isdigit() for ch in text)
    non_ascii = sum(ord(ch) > 127 for ch in text)
    return int(1.6 * digits + 0.22 * (len(text) - digits - non_ascii) + 0.95 * non_ascii)


def _id_cols(t: Table) -> list[str]:
    cols = ["row"] + [c for c in ("section", "group") if (t.df[c] != "").any()]
    return cols + (["label"] if "label" in t.df else [])


def _header(t: Table) -> list[str]:
    lines = [f"## Table {t.id} (sheet '{t.sheet}', {len(t.df)} rows, kind={t.kind})"]
    if t.context:
        lines.append("Context: " + " | ".join(t.context))
    if t.label_header:
        lines.append(f"Label column header: {t.label_header}")
    lines.append("Columns:")
    for c, name in t.columns.items():
        dtype = t.df[c].dtype
        kind = "number" if pd.api.types.is_numeric_dtype(dtype) else \
            "date" if pd.api.types.is_datetime64_any_dtype(dtype) else "text"
        lines.append(f"  {c} = {name}  [{kind}]")
    return lines


def _rows(df: pd.DataFrame, cols: list[str]) -> list[str]:
    lines = ["| " + " | ".join(cols) + " |"]
    for rec in df[cols].itertuples(index=False):
        lines.append("| " + " | ".join(fmt(v) for v in rec) + " |")
    return lines


def render_full(t: Table) -> str:
    lines = _header(t) + _rows(t.df, _id_cols(t) + list(t.columns))
    if t.notes:
        lines.append("Notes: " + " ".join(t.notes))
    return "\n".join(lines)


def render_compact(t: Table, sample_rows: int = 5, max_distinct: int = 40) -> str:
    lines = _header(t)
    ids = _id_cols(t)
    if len(ids) > 1:
        lines.append("All row labels (no values):")
        lines += _rows(t.df, ids)
    lines.append(f"First {sample_rows} rows (with values):")
    lines += _rows(t.df.head(sample_rows), ids + list(t.columns))
    for c in t.columns:
        if t.df[c].dtype == object:
            counts = t.df[c].dropna().astype(str).value_counts()
            if len(counts) > 0.5 * len(t.df) and len(counts) > max_distinct:
                lines.append(f"{c}: {len(counts)} unique values, e.g. {counts.index[:3].tolist()}")
            else:
                vals = counts.index[:max_distinct].tolist()
                lines.append(f"Distinct values of {c} ({len(vals)} of {len(counts)}): {vals}")
    if t.notes:
        lines.append("Notes: " + " ".join(t.notes))
    return "\n".join(lines)


def _num(v: float) -> str:
    return fmt(round(float(v), 4))


def render_stats(t: Table, max_distinct: int = 40) -> list[str]:
    df = t.df
    lines = [f"Column statistics over ALL {len(df)} rows:"]
    for c, name in t.columns.items():
        s = df[c].dropna()
        if s.empty:
            lines.append(f"  {c}: all empty")
        elif pd.api.types.is_numeric_dtype(s):
            lines.append(f"  {c}: count={len(s)} sum={_num(s.sum())} min={_num(s.min())} max={_num(s.max())} "
                         f"mean={_num(s.mean())}")
        elif pd.api.types.is_datetime64_any_dtype(s):
            lines.append(f"  {c}: count={len(s)} min={fmt(s.min())} max={fmt(s.max())}")
        else:
            counts = s.astype(str).value_counts()
            if len(counts) <= max_distinct:
                lines.append(f"  {c}: " + ", ".join(f"{k} ({v})" for k, v in counts.items()))
            else:
                lines.append(f"  {c}: {len(counts)} distinct values")
    for c in ("section", "group"):
        vals = [v for v in dict.fromkeys(df[c]) if v]
        if vals:
            lines.append(f"  {c} values ({len(vals)}): " + " | ".join(vals[:200]))
    return lines


def render_partial(t: Table, max_tokens: int) -> str:
    lines = _header(t) + render_stats(t)
    cols = _id_cols(t) + list(t.columns)
    notes = "Notes: " + " ".join(t.notes) if t.notes else ""
    used = est_tokens("\n".join(lines) + notes) + 80
    rows = ["| " + " | ".join(cols) + " |"]
    for rec in t.df[cols].itertuples(index=False):
        line = "| " + " | ".join(fmt(v) for v in rec) + " |"
        cost = est_tokens(line) + 1
        if used + cost > max_tokens:
            break
        rows.append(line)
        used += cost
    shown = len(rows) - 1
    lines.append(f"Rows (showing the first {shown} of {len(t.df)} rows; the code solver gets all rows in "
                 f'tables["{t.id}"]):')
    lines += rows
    if shown < len(t.df):
        lines.append(f"... {len(t.df) - shown} more rows not shown (use the statistics above or code).")
    if notes:
        lines.append(notes)
    return "\n".join(lines)


def cut_lines(text: str, max_tokens: int) -> str:
    out, used = [], 0
    for line in text.split("\n"):
        cost = est_tokens(line) + 1
        if used + cost > max_tokens:
            out.append("... (truncated: table too large for the prompt budget)")
            break
        out.append(line)
        used += cost
    return "\n".join(out)


def render_tables(tables: list[Table], budget: int) -> str:
    texts = {t.id: render_full(t) for t in tables}
    size = {k: est_tokens(v) for k, v in texts.items()}
    if sum(size.values()) > budget:
        left = budget
        order = sorted(tables, key=lambda t: size[t.id])
        for i, t in enumerate(order):
            share = left // (len(order) - i)
            if size[t.id] > share:
                texts[t.id] = render_partial(t, share)
                size[t.id] = est_tokens(texts[t.id])
            left -= size[t.id]
    if sum(size.values()) > budget:
        per = budget // max(len(tables), 1)
        texts = {k: cut_lines(v, per) for k, v in texts.items()}
    return "\n\n".join(texts[t.id] for t in tables)


def describe_frames(tables: list[Table]) -> str:
    lines = []
    for t in tables:
        lines.append(f'- tables["{t.id}"]: columns [{", ".join(_id_cols(t) + list(t.columns))}]')
        if t.long is not None:
            lines.append(f'- tables["{t.id}_long"]: columns [row, section, group, label, col, col_name, value] '
                         f"(one row per numeric cell)")
    return "\n".join(lines)
