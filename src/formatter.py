"""Bước ⑥: chuẩn hoá đáp án theo định dạng câu hỏi yêu cầu (Exact Match)."""
from __future__ import annotations

import difflib
import re
from dataclasses import dataclass

WORD_NUM = {"one": 1, "two": 2, "three": 3, "four": 4, "zero": 0}
NUMERIC_STR = re.compile(r"^[+-]?\$?\d[\d,]*(\.\d+)?%?$")


@dataclass
class Spec:
    yesno: bool = False
    integer: bool = False
    decimals: int | None = None
    is_list: bool = False


def answer_spec(query: str) -> Spec:
    q = query.lower()
    s = Spec()
    s.yesno = bool(re.search(r"\byes\W+or\W+no\b", q))
    m = re.search(r"\b(\d|one|two|three|four|zero)\s+decimal", q)
    if m:
        s.decimals = int(m.group(1)) if m.group(1).isdigit() else WORD_NUM[m.group(1)]
    s.integer = s.decimals is None and bool(re.search(r"\bintegers?\b|whole number", q))
    s.is_list = bool(re.search(r"separat\w*|comma|\btwo integers\b|\bin the order\b", q))
    return s


def to_number(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str) and NUMERIC_STR.match(v.strip()):
        return float(re.sub(r"[^\d.+-]", "", v.strip()))
    return None


def fmt_number(x: float, spec: Spec) -> str:
    if spec.decimals is not None:
        return f"{x:.{spec.decimals}f}"
    if spec.integer or float(x).is_integer():
        return str(int(round(x)))
    return f"{x:.10f}".rstrip("0").rstrip(".")


def snap_label(s: str, labels: set[str]) -> str:
    """Khớp đáp án chữ về đúng nhãn gốc trong bảng."""
    if not labels or s in labels:
        return s
    lower = {l.lower(): l for l in labels}
    if s.lower() in lower:
        return lower[s.lower()]
    m = difflib.get_close_matches(s.lower(), list(lower), n=1, cutoff=0.92)
    return lower[m[0]] if m else s


def clean_str(s: str) -> str:
    s = s.strip().strip("`").strip()
    if len(s) >= 2 and s[0] == s[-1] and s[0] in "\"'":
        s = s[1:-1].strip()
    s = re.sub(r"\s*\(\d+\)", "", s)
    return s.rstrip(".").strip() if not re.search(r"\b[A-Z][a-z]{0,3}\.$", s) else s   # giữ "Oct."


def fmt_scalar(v, spec: Spec, labels: set[str]) -> str:
    if v is None:
        return ""
    if spec.yesno:
        if isinstance(v, bool):
            return "Yes" if v else "No"
        m = re.match(r"\s*(yes|no|true|false)\b", str(v), re.I)
        if m:
            return "Yes" if m.group(1).lower() in ("yes", "true") else "No"
    num = to_number(v)
    if num is not None:
        return fmt_number(num, spec)
    return snap_label(clean_str(str(v)), labels)


def format_answer(value, query: str, labels: set[str] | None = None) -> str:
    labels = labels or set()
    spec = answer_spec(query)
    if isinstance(value, (list, tuple)):
        if len(value) == 1:
            return fmt_scalar(value[0], spec, labels)
        return ", ".join(fmt_scalar(x, spec, labels) for x in value)
    # "a, b" có khoảng trắng sau dấu phẩy là nhiều phần; "7,668" vẫn là 1 số
    if isinstance(value, str) and re.search(r",\s", value):
        whole = snap_label(clean_str(value), labels)
        if whole in labels:
            return whole
        parts = [p.strip() for p in value.split(",")]
        if all(to_number(p) is not None for p in parts):
            return ", ".join(fmt_scalar(p, spec, labels) for p in parts)
        if spec.is_list:
            items = re.split(r"\s*,\s*(?:and\s+)?", clean_str(value))
            return ", ".join(snap_label(clean_str(p), labels) for p in items if p)
    return fmt_scalar(value, spec, labels)


def same_answer(a: str, b: str) -> bool:
    if a.strip().lower() == b.strip().lower():
        return True
    na, nb = to_number(a), to_number(b)
    return na is not None and nb is not None and abs(na - nb) <= 1e-9 * max(1.0, abs(na))
