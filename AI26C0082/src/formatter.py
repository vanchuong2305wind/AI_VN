from __future__ import annotations

import difflib
import re
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal

WORD_NUM = {"zero": 0, "one": 1, "two": 2, "three": 3, "four": 4,
            "một": 1, "hai": 2, "ba": 3, "bốn": 4,
            "un": 1, "une": 1, "deux": 2, "trois": 3, "quatre": 4,
            "uno": 1, "una": 1, "dos": 2, "tres": 3, "cuatro": 4,
            "eine": 1, "einer": 1, "zwei": 2, "drei": 3, "vier": 4,
            "첫째": 1, "둘째": 2, "셋째": 3, "넷째": 4, "첫": 1, "두": 2, "세": 3,
            "一": 1, "两": 2, "二": 2, "三": 3, "四": 4}
_N = r"(\d+|" + "|".join(sorted(WORD_NUM, key=len, reverse=True)) + ")"
DECIMAL_RES = [re.compile(p) for p in (
    rf"\b{_N}\s+decimal",
    rf"\b{_N}\s+chữ\s+số\s+(?:thập\s+phân|sau\s+dấu\s+phẩy)",
    rf"\b{_N}\s+décimales?|\b{_N}\s+chiffres?\s+après\s+la\s+virgule",
    rf"\b{_N}\s+(?:cifras?\s+)?decimal(?:es)?",
    rf"\b{_N}\s+(?:nachkommastellen|dezimalstellen)",
    rf"소수점\s*(?:이하\s*)?{_N}\s*(?:번째\s*)?자리",
    rf"(?:保留|精确到|取){_N}位小数|小数点后{_N}位",
    rf"小数(?:点以下|第)\s*{_N}\s*(?:位|桁)",
)]
ROUND_AT_RE = re.compile(r"^\s*(?:에서\s*반올림|を四捨五入)")
NEAREST = [(re.compile(r"nearest\s+tenth|hàng\s+phần\s+mười|au\s+dixième"), 1),
           (re.compile(r"nearest\s+hundredth|hàng\s+phần\s+trăm|au\s+centième"), 2),
           (re.compile(r"nearest\s+thousandth|hàng\s+phần\s+nghìn|au\s+millième"), 3)]
INTEGER_RE = re.compile(r"\bintegers?\b|whole\s+number|số\s+nguyên|nombre\s+entier|\bentiers?\b|número\s+entero|"
                        r"ganze\s+zahl|정수|整数")
NO_COMMA_RE = re.compile(r"without\s+(?:thousands\s+separators|commas?)|no\s+commas?|không\s+(?:có\s+|kèm\s+)?dấu\s+phẩy|"
                         r"sans\s+(?:espace\s+ni\s+)?virgule|쉼표\s*없이|不加逗号|不带逗号")
LIST_RE = re.compile(r"separat\w*|comma|\btwo integers\b|\bin the order\b|dấu\s+phẩy|virgule|쉼표|逗号")
SHEET_RE = re.compile(r"\bsheets?\b|worksheet|feuille|trang\s+tính|시트|工作表")
NUMERIC_STR = re.compile(r"^[+-]?\$?\d[\d,]*(\.\d+)?%?$")
LIST_SEP = re.compile(r"\s*,(?!\d{3}(?!\d))\s*")


@dataclass
class Spec:
    yesno: bool = False
    integer: bool = False
    decimals: int | None = None
    is_list: bool = False
    sheet: bool = False


def _decimals(q: str) -> int | None:
    for rx in DECIMAL_RES:
        m = rx.search(q)
        if m:
            word = next(g for g in m.groups() if g)
            n = int(word) if word.isdigit() else WORD_NUM[word]
            return max(n - 1, 0) if ROUND_AT_RE.match(q[m.end():]) else n
    for rx, n in NEAREST:
        if rx.search(q):
            return n
    return None


def answer_spec(query: str) -> Spec:
    q = query.lower()
    s = Spec()
    s.yesno = bool(re.search(r"\byes\W+or\W+no\b", q))
    s.decimals = _decimals(q)
    s.integer = s.decimals is None and bool(INTEGER_RE.search(q))
    s.is_list = bool(LIST_RE.search(NO_COMMA_RE.sub(" ", q)))
    s.sheet = bool(SHEET_RE.search(q))
    return s


def to_number(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str) and NUMERIC_STR.match(v.strip()):
        return float(re.sub(r"[^\d.+-]", "", v.strip()))
    return None


def round_half_up(x: float, d: int) -> str:
    q = Decimal(f"{x:.12g}").quantize(Decimal(1).scaleb(-d), rounding=ROUND_HALF_UP)
    s = f"{q:.{d}f}"
    return s[1:] if s.startswith("-") and q == 0 else s


def fmt_number(x: float, spec: Spec) -> str:
    if spec.decimals is not None:
        return round_half_up(x, spec.decimals)
    if spec.integer or float(x).is_integer():
        return round_half_up(x, 0)
    return f"{x:.10f}".rstrip("0").rstrip(".")


def snap_label(s: str, labels: set[str]) -> str:
    if not labels or s in labels:
        return s
    lower = {l.lower(): l for l in labels}
    if s.lower() in lower:
        return lower[s.lower()]
    m = difflib.get_close_matches(s.lower(), list(lower), n=1, cutoff=0.92)
    return lower[m[0]] if m else s


def strip_path(s: str, labels: set[str]) -> str:
    if " > " not in s:
        return s
    parts = [p for p in s.split(" > ") if re.search(r"[^\W\d_]{2}", p)]
    known = [p for p in parts if snap_label(p, labels) in labels]
    return (known or parts or [s])[-1]


AS_IN_TABLE_RE = re.compile(r"as (?:it )?appears? in the table|exactly as (?:written|shown) in the table", re.I)


def question_name(s: str, query: str) -> str:
    if not s or s in query or AS_IN_TABLE_RE.search(query):
        return s
    words = s.split()
    for k in range(len(words) - 1, 0, -1):
        name = " ".join(words[:k])
        if re.search(rf"(?:^|[,:;`“\"(]|\bor|\band)\s*{re.escape(name)}\s*(?=$|[,?.;`”\")]|\s+or\b|\s+and\b)", query):
            return name
    return s


def clean_str(s: str) -> str:
    s = s.strip().strip("`").strip()
    if len(s) >= 2 and s[0] == s[-1] and s[0] in "\"'":
        s = s[1:-1].strip()
    s = re.sub(r"\s*\(\d+\)", "", s)
    return s.rstrip(".").strip() if not re.search(r"\b[A-Z][a-z]{0,3}\.$", s) else s


def fmt_scalar(v, spec: Spec, labels: set[str], aliases: dict[str, str], query: str = "") -> str:
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
    s = clean_str(str(v))
    if spec.sheet and aliases:
        m = re.fullmatch(r"(?:table\s+)?(T\d+)(?:_long)?", s, re.I)
        if m and m.group(1).upper() in aliases:
            return aliases[m.group(1).upper()]
    return question_name(snap_label(strip_path(s, labels), labels), query)


def format_answer(value, query: str, labels: set[str] | None = None, aliases: dict[str, str] | None = None) -> str:
    labels, aliases = labels or set(), aliases or {}
    spec = answer_spec(query)
    fmt = lambda x: fmt_scalar(x, spec, labels, aliases, query)
    if isinstance(value, (list, tuple)):
        return fmt(value[0]) if len(value) == 1 else ", ".join(fmt(x) for x in value)
    if isinstance(value, str) and LIST_SEP.search(value):
        whole = snap_label(clean_str(value), labels)
        if whole in labels:
            return whole
        parts = [p for p in LIST_SEP.split(value.strip()) if p]
        if all(to_number(p) is not None for p in parts):
            return ", ".join(fmt(p) for p in parts)
        if spec.is_list:
            items = re.split(r"\s*,(?!\d{3}(?!\d))\s*(?:and\s+)?", clean_str(value))
            return ", ".join(fmt(p) for p in items if p)
    return fmt(value)


def same_answer(a: str, b: str) -> bool:
    if a.strip().lower() == b.strip().lower():
        return True
    na, nb = to_number(a), to_number(b)
    return na is not None and nb is not None and abs(na - nb) <= 1e-9 * max(1.0, abs(na))


def near_miss(a: str, b: str, rel: float = 0.005) -> bool:
    pa, pb = LIST_SEP.split(a.strip()), LIST_SEP.split(b.strip())
    if len(pa) != len(pb):
        return False
    for x, y in zip(pa, pb):
        nx, ny = to_number(x), to_number(y)
        if nx is None or ny is None or ("." not in x and "." not in y):
            if x.strip().lower() != y.strip().lower():
                return False
        elif abs(nx - ny) > rel * max(abs(nx), abs(ny), 1e-9):
            return False
    return True
