import sys

sys.path.pop(0)

import datetime as dt
import json
import math
import pickle
import re
import traceback

import numpy as np
import pandas as pd


def to_jsonable(v):
    if isinstance(v, pd.DataFrame):
        if v.size == 1:
            return to_jsonable(v.iloc[0, 0])
        return to_jsonable(v.values.tolist())
    if isinstance(v, (pd.Series, pd.Index, np.ndarray, list, tuple, set)):
        items = list(v.tolist() if hasattr(v, "tolist") else v)
        return to_jsonable(items[0]) if len(items) == 1 and not isinstance(v, (list, tuple)) else [to_jsonable(x) for x in items]
    if isinstance(v, np.generic):
        return v.item()
    if isinstance(v, (pd.Timestamp, dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and math.isnan(v):
        return None
    if isinstance(v, (str, int, float, bool)) or v is None:
        return v
    return str(v)


def main():
    with open(sys.argv[1], "rb") as f:
        tables = pickle.load(f)
    code = open(sys.argv[2], encoding="utf-8").read()
    env = {"pd": pd, "np": np, "math": math, "re": re, "datetime": dt}
    try:
        exec(compile(code, "solution.py", "exec"), env)
        result = env["solve"](tables)
        print(json.dumps({"ok": True, "result": to_jsonable(result)}, ensure_ascii=True))
    except Exception:
        print(json.dumps({"ok": False, "error": traceback.format_exc(limit=3)[-1500:]}, ensure_ascii=True))


if __name__ == "__main__":
    main()
