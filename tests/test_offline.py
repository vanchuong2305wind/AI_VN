"""Kiểm thử offline (không cần API): code solve() viết tay đóng vai LLM,
chạy qua parser -> sandbox -> formatter thật, so với đáp án mẫu.
Mục đích: chứng minh dữ liệu sau parse đủ để trả lời và formatter ra đúng định dạng.
File này KHÔNG được dùng trong pipeline chính.

Chạy:  python tests/test_offline.py
"""
import asyncio
import csv
import json
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from src.formatter import format_answer          # noqa: E402
from src.parser import parse_workbook            # noqa: E402
from src.sandbox import dump_tables, run_code    # noqa: E402

DATA = ROOT / "data_raw"

# Code mẫu như LLM sẽ viết (dùng cột/nhãn đúng như phần render cho LLM thấy)
SOLUTIONS = {
    "p4_q001": """
def solve(tables):
    df = tables["T1"]; df = df.assign(age=df.c6 - df.c5)
    return df.sort_values("age").iloc[1]["c4"]""",
    "p4_q002": """
def solve(tables):
    df = tables["T1"]; df = df.assign(age=df.c6 - df.c5)
    return df.sort_values("age", ascending=False).iloc[1]["c7"]""",
    "p4_q003": """
def solve(tables):
    df = tables["T1"]; df = df.assign(age=df.c6 - df.c5)
    return df.nsmallest(10, "age")["c7"].max()""",
    "p4_q004": """
def solve(tables):
    df = tables["T1"]; r = df[df.label == "Agricultural exports"].iloc[0]
    return "Yes" if r.c6 > r.c5 else "No" """,
    "p4_q005": """
def solve(tables):
    df = tables["T1"]
    return df.loc[df.c7.abs().idxmax(), "label"]""",
    "p4_q006": """
def solve(tables):
    df = tables["T1"]; w = df[df.section == "Women"]
    return abs(w[w.label == "Civilian noninstitutional population"].c4.iloc[0] - w[w.label == "Civilian labor force"].c4.iloc[0])""",
    "p4_q007": """
def solve(tables):
    df = tables["T1"]; r = df[(df.section == "Women") & (df.label == "Unemployment rate")]
    return r[[f"c{i}" for i in range(1, 9)]].iloc[0].mean()""",
    "p4_q008": """
def solve(tables):
    df = tables["T1"]; t = df[(df.section == "TOTAL") & (df.group == "Total, 16 years and over")]
    cols = [f"c{i}" for i in range(1, 13)]
    return t[t.label == "20 to 24 years"][cols].sum(axis=1).iloc[0] - t[t.label == "16 to 19 years"][cols].sum(axis=1).iloc[0]""",
    "p4_q009": """
def solve(tables):
    df = tables["T1"]; r = df[(df.section == "TOTAL") & (df.group == "Total, 16 years and over") & (df.label == "20 to 24 years")].iloc[0]
    return r.c2 + r.c3""",
    "p4_q010": """
def solve(tables):
    L = tables["T1_long"]; r = L[(L.section == "TOTAL") & (L.group == "Total, 16 years and over") & (L.label == "16 to 19 years")]
    return r.loc[r.value.idxmax(), "col_name"]""",
    "p4_q011": """
def solve(tables):
    df = tables["T1"]; r = df[(df.section == "TOTAL") & (df.group == "Total, 16 years and over") & (df.label == "20 years and over")].iloc[0]
    return r.c7 > r.c5""",
    "p4_q012": """
def solve(tables):
    df = tables["T1"]
    return df[df.label == "Total, 16 years and over"].c4.iloc[0]""",
    "p4_q013": """
def solve(tables):
    df = tables["T1"]; t = df[df.section == "TOTAL"].set_index("label")
    v = {"Agriculture": t.loc["Agriculture and related", "c11"], "Manufacturing": t.loc["Manufacturing", "c11"], "Construction": t.loc["Construction", "c11"]}
    return sorted(v, key=v.get, reverse=True)""",
    "p4_q014": """
def solve(tables):
    df = tables["T1"]; r = df[(df.section == "TOTAL") & (df.label == "Manufacturing")].iloc[0]
    return r.c4 + r.c5""",
    "p4_q015": """
def solve(tables):
    df = tables["T1"]
    return df[df.label == "5 to 14 hours"].c2.iloc[0]""",
    "p4_q016": """
def solve(tables):
    df = tables["T1"]; d = df[df.label != "Total, persons at work"]
    return d.loc[d.c5.idxmax(), "label"]""",
    "p4_q017": """
def solve(tables):
    df = tables["T1"]
    return abs(df[df.label == "Manufacturing"].c1.iloc[0] - df[df.label == "Construction"].c1.iloc[0])""",
    "p4_q018": """
def solve(tables):
    df = tables["T1"]; r = df[df.label == "Total, 16 years and over"].iloc[0]
    return [r.c2, r.c6]""",
    "p4_q019": """
def solve(tables):
    df = tables["T1"]; d = df[df.label == "Widowed, divorced, or separated"]
    m = d[d.section.str.contains("Men")].c7.iloc[0]; w = d[d.section.str.contains("Women")].c7.iloc[0]
    return [m, w]""",
    "p4_q020": """
def solve(tables):
    df = tables["T1"]
    return df[df.label == "Total, 16 years and over"].c1.iloc[0]""",
}


async def main():
    queries = json.loads((DATA / "queries.json").read_text(encoding="utf-8"))
    with open(DATA / "sample_submission.csv", encoding="utf-8-sig") as f:
        gold = {r["id"]: r["answer"] for r in csv.DictReader(f)}
    cache, ok = {}, 0
    for q in queries:
        fn = q["file_name"]
        if fn not in cache:
            tables = parse_workbook(DATA / "xlsx_files" / fn)
            labels = set().union(*(t.labels for t in tables))
            cache[fn] = (dump_tables(tables, "test_" + fn), labels)
        pkl, labels = cache[fn]
        res = await run_code(SOLUTIONS[q["id"]], pkl)
        ans = format_answer(res["result"], q["query"], labels) if res.get("ok") else f"ERROR {res.get('error')}"
        good = ans == gold[q["id"]]
        ok += good
        print(f"{'✓' if good else '✗'} {q['id']}: {ans!r}" + ("" if good else f"  (gold {gold[q['id']]!r})"))
    print(f"\n{ok}/{len(queries)} đúng")


if __name__ == "__main__":
    asyncio.run(main())
