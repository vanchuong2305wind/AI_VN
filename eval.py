"""Chấm Exact Match: python eval.py --pred submission.csv --gold data_raw/sample_submission.csv"""
import argparse
import csv
import json
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def read_csv(path) -> dict:
    with open(path, encoding="utf-8-sig", newline="") as f:
        return {r["id"]: r["answer"] for r in csv.DictReader(f)}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pred", default="submission.csv")
    ap.add_argument("--gold", default="data_raw/sample_submission.csv")
    ap.add_argument("--log", default="", help="file log .jsonl (mặc định: log mới nhất trong logs/)")
    args = ap.parse_args()

    pred, gold = read_csv(args.pred), read_csv(args.gold)
    log_path = args.log or max(Path("logs").glob("*.jsonl"), key=lambda p: p.stat().st_mtime, default="")
    logs = {}
    if log_path:
        for line in Path(log_path).read_text(encoding="utf-8").splitlines():
            r = json.loads(line)
            logs[r["id"]] = r

    ids = [i for i in gold if i in pred]
    correct = [i for i in ids if pred[i].strip() == gold[i].strip()]
    print(f"Exact Match: {len(correct)}/{len(ids)} = {len(correct) / max(len(ids), 1):.1%}\n")
    for i in ids:
        if i in correct:
            continue
        r = logs.get(i, {})
        print(f"✗ {i}: pred={pred[i]!r}  gold={gold[i]!r}  [{r.get('source', '?')}]")
        if r:
            print(f"   Q: {r.get('query')}")
            print(f"   code={r.get('ans_a')!r}  direct={r.get('ans_b')!r}  judge={r.get('ans_judge')!r}")
    if logs:
        rs = [logs[i] for i in ids if i in logs]
        if rs:
            n = len(rs)
            print(f"\nTB/câu: {sum(r['seconds'] for r in rs) / n:.1f}s, "
                  f"{sum(r['prompt_tokens'] + r['completion_tokens'] for r in rs) / n:.0f} token, "
                  f"{sum(r['calls'] for r in rs) / n:.1f} lượt gọi")
            src = {}
            for r in rs:
                src[r["source"]] = src.get(r["source"], 0) + 1
            print("Nguồn đáp án:", src)


if __name__ == "__main__":
    main()
