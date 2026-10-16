import asyncio
import csv
import json
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from src.formatter import format_answer
from src.parser import parse_workbook
from src.sandbox import dump_tables, run_code

DATA = ROOT / "data_raw"

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


FORMAT_CASES = [
    (50.0, "Chỉ trả lời bằng số với đúng một chữ số thập phân.", "50.0"),
    (14.2, "소수점 둘째 자리까지, % 기호 없이 숫자만 답하십시오.", "14.20"),
    (82, "保留一位小数，只回答数字，不带%符号。", "82.0"),
    (14.8, "Répondez uniquement par le nombre avec deux décimales, sans le symbole %.", "14.80"),
    (3.456, "소수점 둘째 자리에서 반올림하여 답하십시오", "3.5"),
    (27124.0, "Chỉ trả lời bằng số nguyên, không có dấu phẩy.", "27124"),
    (2.675, "using exactly two decimal places", "2.68"),
    (2.5, "Answer with only the integer.", "3"),
    ("1,000+, 2.8", "Answer as `size class, gap`, separated by a comma and one space, with the gap using exactly one "
                    "decimal place.", "1,000+, 2.8"),
    ("7,668", "Answer with only the integer, without commas or units.", "7668"),
    ("Chinese > 2014", "Answer with only the group name exactly as it appears in the table.", "Chinese"),
    ("T3, 2232", "Which sheet reports the highest value? Answer as `sheet name, integer value`, separated by a comma.",
     "table_2, 2232"),
    ("Manufacturing, Construction, Agriculture and related",
     "Rank the following sectors: Agriculture, Manufacturing, and Construction. Answer with the three sector names "
     "in descending order, separated by a comma and one space.", "Manufacturing, Construction, Agriculture"),
    ("Hispanic or Latino ethnicity", "Answer with only the group name: White, Asian, or Hispanic or Latino.",
     "Hispanic or Latino"),
    ("Asian Indian", "Which group? Answer with only the group name exactly as it appears in the table.", "Asian Indian"),
]


def test_formatter() -> int:
    labels, sheets = {"Chinese", "1,000+", "2014", "Chinese > 2014"}, {"T3": "table_2"}
    ok = 0
    for value, query, want in FORMAT_CASES:
        got = format_answer(value, query, labels, sheets)
        ok += got == want
        if got != want:
            print(f"✗ format {value!r}: {got!r}  (muốn {want!r})")
    print(f"Formatter: {ok}/{len(FORMAT_CASES)} đúng\n")
    return ok


async def main():
    test_formatter()
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
