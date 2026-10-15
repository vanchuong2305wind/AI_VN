"""Chạy pipeline: python main.py --input_dir data_raw --output submission.csv"""
import argparse
import asyncio
import csv
import json
import sys
import time
from pathlib import Path

from src import config
from src.llm import LLM
from src.solver import Solver

sys.stdout.reconfigure(encoding="utf-8", errors="replace")   # console Windows không in được tiếng Việt
sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def load_queries(input_dir: Path) -> list[dict]:
    path = input_dir / "queries.json"
    if not path.exists():
        path = next(input_dir.rglob("queries.json"))
    queries = json.loads(path.read_text(encoding="utf-8-sig"))
    for q in queries:
        q["file_name"] = q.get("file_name") or q.get("file name") or q.get("filename")
    return queries


def write_csv(path: Path, queries: list[dict], answers: dict):
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["id", "answer"])
        for q in queries:
            w.writerow([q["id"], answers.get(str(q["id"]), "")])


async def run(args):
    start = time.time()
    input_dir = Path(args.input_dir)
    xlsx_dir = input_dir / "xlsx_files" if (input_dir / "xlsx_files").is_dir() else input_dir
    queries = load_queries(input_dir)
    if args.ids:
        queries = [q for q in queries if str(q["id"]) in set(args.ids.split(","))]
    out, answers = Path(args.output), {}
    write_csv(out, queries, answers)

    Path("logs").mkdir(exist_ok=True)
    log_path = Path("logs") / f"run_{time.strftime('%Y%m%d_%H%M%S')}.jsonl"

    try:
        llm = LLM()
    except RuntimeError as e:
        sys.exit(f"Lỗi cấu hình: {e}")
    solver = Solver(llm, xlsx_dir, start)
    sem = asyncio.Semaphore(args.concurrency)

    async def one(q):
        async with sem:
            res = await solver.solve(str(q["id"]), q["file_name"], q["query"])
        answers[res["id"]] = res["answer"]
        write_csv(out, queries, answers)          # ghi lại sau mỗi câu để luôn có file hợp lệ
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(json.dumps(res, ensure_ascii=False, default=str) + "\n")
        print(f"[{len(answers)}/{len(queries)}] {res['id']}: {res['answer']!r}  "
              f"({res['source']}, {res['seconds']}s, {res['calls']} calls, "
              f"{res['prompt_tokens'] + res['completion_tokens']} tok)")
        return res

    results = await asyncio.gather(*(one(q) for q in queries))
    await llm.close()

    n = max(len(results), 1)
    print(f"\nXong {len(results)} câu trong {time.time() - start:.1f}s -> {out}")
    print(f"TB/câu: {sum(r['seconds'] for r in results) / n:.1f}s, "
          f"{sum(r['prompt_tokens'] + r['completion_tokens'] for r in results) / n:.0f} token, "
          f"${sum(r['cost'] for r in results) / n:.5f}")
    print(f"Log: {log_path}")


def main():
    ap = argparse.ArgumentParser(description="Excel Table QA")
    ap.add_argument("--input_dir", required=True, help="thư mục chứa xlsx_files/ và queries.json")
    ap.add_argument("--output", default="submission.csv")
    ap.add_argument("--concurrency", type=int, default=config.CONCURRENCY)
    ap.add_argument("--ids", default="", help="chỉ chạy các id này, cách nhau bởi dấu phẩy")
    asyncio.run(run(ap.parse_args()))


if __name__ == "__main__":
    main()
