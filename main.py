import argparse
import asyncio
import csv
import json
import os
import sys
import time
from pathlib import Path

from src import config
from src.llm import LLM
from src.solver import Solver

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def load_queries(path: Path) -> list[dict]:
    queries, seen = [], set()
    for q in json.loads(path.read_text(encoding="utf-8-sig")):
        if str(q["id"]) in seen:
            continue
        seen.add(str(q["id"]))
        q["file_name"] = q.get("file_name") or q.get("file name") or q.get("filename")
        queries.append(q)
    return queries


def write_csv(path: Path, queries: list[dict], answers: dict):
    tmp = path.with_name(path.name + ".tmp")
    with open(tmp, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["id", "answer"])
        for q in queries:
            w.writerow([q["id"], answers.get(str(q["id"]), "")])
    os.replace(tmp, path)


def open_log() -> Path | None:
    if not config.LOG_DIR:
        return None
    try:
        log_dir = Path(config.LOG_DIR)
        log_dir.mkdir(parents=True, exist_ok=True)
        log_path = log_dir / f"run_{time.strftime('%Y%m%d_%H%M%S')}.jsonl"
        try:
            log_path.open("x").close()
        except FileExistsError:
            log_path = log_path.with_name(f"{log_path.stem}_{os.getpid()}.jsonl")
            log_path.touch()
        return log_path
    except OSError:
        return None


def resolve_paths(args) -> tuple[Path, Path]:
    if args.xlsx_dir and args.queries:
        return Path(args.xlsx_dir), Path(args.queries)
    if not args.input_dir:
        sys.exit("Cần --xlsx-dir và --queries (hoặc --input_dir).")
    input_dir = Path(args.input_dir)
    xlsx_dir = Path(args.xlsx_dir) if args.xlsx_dir else (
        input_dir / "xlsx_files" if (input_dir / "xlsx_files").is_dir() else input_dir)
    queries = Path(args.queries) if args.queries else input_dir / "queries.json"
    if not queries.exists():
        queries = next(input_dir.rglob("queries.json"))
    return xlsx_dir, queries


def apply_llm_args(args):
    if args.openrouter_key:
        config.API_KEY = args.openrouter_key
    if args.openrouter_base_url:
        config.BASE_URL = args.openrouter_base_url
        config.IS_OPENROUTER = "openrouter.ai" in config.BASE_URL
    if args.openrouter_model:
        config.MODEL = args.openrouter_model


async def run(args):
    start = time.time()
    apply_llm_args(args)
    xlsx_dir, queries_path = resolve_paths(args)
    queries = load_queries(queries_path)
    if args.ids:
        queries = [q for q in queries if str(q["id"]) in set(args.ids.split(","))]
    out, answers = Path(args.output), {}
    out.parent.mkdir(parents=True, exist_ok=True)
    write_csv(out, queries, answers)
    log_path = open_log()

    try:
        llm = LLM(model=config.MODEL)
    except RuntimeError as e:
        sys.exit(f"Lỗi cấu hình: {e}")
    solver = Solver(llm, xlsx_dir, start)
    sem = asyncio.Semaphore(args.concurrency)

    async def one(q):
        async with sem:
            res = await solver.solve(str(q["id"]), q["file_name"], q["query"])
        answers[res["id"]] = res["answer"]
        write_csv(out, queries, answers)
        if log_path:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(json.dumps(res, ensure_ascii=False, default=str) + "\n")
        print(f"[{len(answers)}/{len(queries)}] {res['id']}: {res['answer']!r}  "
              f"({res['source']}, {res['seconds']}s, {res['calls']} calls, "
              f"{res['prompt_tokens'] + res['completion_tokens']} tok)", flush=True)
        return res

    tasks = [asyncio.create_task(one(q)) for q in queries]
    done, pending = await asyncio.wait(tasks, timeout=max(config.HARD_DEADLINE - (time.time() - start), 1)) \
        if tasks else (set(), set())
    for t in pending:
        t.cancel()
    if pending:
        await asyncio.gather(*pending, return_exceptions=True)
        print(f"⚠ Hết giờ ({config.HARD_DEADLINE:.0f}s): {len(pending)} câu chưa xong, ghi đáp án mặc định.")
        for q in queries:
            answers.setdefault(str(q["id"]), "0")
    write_csv(out, queries, answers)
    await llm.close()

    results = [t.result() for t in done if not t.cancelled() and t.exception() is None]
    n = max(len(results), 1)
    print(f"\nXong {len(results)}/{len(queries)} câu trong {time.time() - start:.1f}s -> {out}")
    print(f"TB/câu: {sum(r['seconds'] for r in results) / n:.1f}s, "
          f"{sum(r['prompt_tokens'] + r['completion_tokens'] for r in results) / n:.0f} token, "
          f"${sum(r['cost'] for r in results) / n:.5f}")
    rs = [r for r in results if "max_prompt" in r]
    if rs:
        print(f"TB/câu: input {sum(r['prompt_tokens'] for r in rs) / n:.0f} + reasoning "
              f"{sum(r['reasoning_tokens'] for r in rs) / n:.0f} = "
              f"{sum(r['prompt_tokens'] + r['reasoning_tokens'] for r in rs) / n:.0f} token (input + reasoning)")
        big = [r["id"] for r in rs if r["max_prompt"] >= 32000]
        print(f"Prompt lớn nhất 1 request: {max(r['max_prompt'] for r in rs)} token"
              + (f"  ⚠ {len(big)} câu có request >= 32K (bậc giá đắt hơn): {big[:10]}" if big else ""))
    if log_path:
        print(f"Log: {log_path}")


def main():
    ap = argparse.ArgumentParser(description="Excel Table QA")
    ap.add_argument("--xlsx-dir", dest="xlsx_dir", default="", help="thư mục chứa các file .xlsx")
    ap.add_argument("--queries", default="", help="đường dẫn queries.json")
    ap.add_argument("--output", default="submission.csv")
    ap.add_argument("--openrouter-key", dest="openrouter_key", default="")
    ap.add_argument("--openrouter-model", dest="openrouter_model", default="")
    ap.add_argument("--openrouter-base-url", dest="openrouter_base_url", default="")
    ap.add_argument("--input_dir", default="", help="thư mục chứa xlsx_files/ và queries.json (chạy local)")
    ap.add_argument("--concurrency", type=int, default=config.CONCURRENCY)
    ap.add_argument("--ids", default="", help="chỉ chạy các id này, cách nhau bởi dấu phẩy")
    asyncio.run(run(ap.parse_args()))


if __name__ == "__main__":
    main()
