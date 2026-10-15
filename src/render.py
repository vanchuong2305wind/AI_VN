"""Bước ②: bảng -> văn bản cho prompt. Bảng lớn chỉ in tên cột, nhãn dòng và vài dòng mẫu."""
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
    return len(text) // 3


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
            if len(counts) > 0.5 * len(t.df) and len(counts) > max_distinct:   # cột mã/ID
                lines.append(f"{c}: {len(counts)} unique values, e.g. {counts.index[:3].tolist()}")
            else:
                vals = counts.index[:max_distinct].tolist()
                lines.append(f"Distinct values of {c} ({len(vals)} of {len(counts)}): {vals}")
    if t.notes:
        lines.append("Notes: " + " ".join(t.notes))
    return "\n".join(lines)


def render_tables(tables: list[Table], budget: int) -> str:
    """In toàn bộ nếu vừa ngân sách token, nếu không thì rút gọn dần bảng lớn nhất."""
    texts = {t.id: render_full(t) for t in tables}
    for t in sorted(tables, key=lambda t: -len(texts[t.id])):
        if sum(est_tokens(x) for x in texts.values()) <= budget:
            break
        texts[t.id] = render_compact(t)
    if sum(est_tokens(x) for x in texts.values()) > budget:
        for t in tables:
            texts[t.id] = texts[t.id][: budget * 3 // max(len(tables), 1)]
    return "\n\n".join(texts[t.id] for t in tables)


def describe_frames(tables: list[Table]) -> str:
    lines = []
    for t in tables:
        lines.append(f'- tables["{t.id}"]: columns [{", ".join(_id_cols(t) + list(t.columns))}]')
        if t.long is not None:
            lines.append(f'- tables["{t.id}_long"]: columns [row, section, group, label, col, col_name, value] '
                         f"(one row per numeric cell)")
    return "\n".join(lines)
