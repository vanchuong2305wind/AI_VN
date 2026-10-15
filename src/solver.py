"""Điều phối 1 câu hỏi: parse -> render -> code ‖ đọc trực tiếp -> sandbox -> kiểm tra chéo -> chuẩn hoá."""
from __future__ import annotations

import asyncio
import hashlib
import math
import re
import time
from dataclasses import dataclass
from pathlib import Path

from . import config, prompts
from .formatter import format_answer, same_answer
from .llm import LLM, Usage
from .parser import Table, parse_workbook
from .render import describe_frames, render_tables
from .sandbox import dump_tables, run_code


@dataclass
class Workbook:
    tables: list[Table]
    pkl: Path
    rendered: str
    frames: str
    labels: set[str]


def extract_code(text: str) -> str:
    blocks = re.findall(r"```(?:python|py)?\s*\n(.*?)```", text, flags=re.S)
    for b in reversed(blocks):
        if "def solve" in b:
            return b
    if "def solve" in text:
        return text[text.index("def solve"):].replace("```", "")
    return ""


def extract_answer(text: str) -> str:
    m = re.findall(r"ANSWER:\s*(.+)", text)
    return m[-1].strip() if m else ""


def is_valid(result) -> bool:
    if result is None or result == "" or result == []:
        return False
    if isinstance(result, float) and math.isnan(result):
        return False
    if isinstance(result, list) and all(x is None for x in result):
        return False
    return True


def find_file(xlsx_dir: Path, file_name: str) -> Path:
    p = xlsx_dir / file_name
    if p.exists():
        return p
    target = Path(file_name).name.lower()
    for f in xlsx_dir.rglob("*"):
        if f.name.lower() == target:
            return f
    raise FileNotFoundError(file_name)


class Solver:
    def __init__(self, llm: LLM, xlsx_dir: Path, start: float):
        self.llm = llm
        self.xlsx_dir = xlsx_dir
        self.start = start
        self._cache: dict[str, Workbook] = {}
        self._locks: dict[str, asyncio.Lock] = {}

    async def load(self, file_name: str) -> Workbook:
        """Parse + render mỗi file đúng 1 lần."""
        lock = self._locks.setdefault(file_name, asyncio.Lock())
        async with lock:
            if file_name not in self._cache:
                self._cache[file_name] = await asyncio.to_thread(self._load_sync, file_name)
        return self._cache[file_name]

    def _load_sync(self, file_name: str) -> Workbook:
        path = find_file(self.xlsx_dir, file_name)
        tables = parse_workbook(path)
        key = hashlib.md5(str(path.resolve()).encode()).hexdigest()[:12]
        labels = set().union(*(t.labels for t in tables)) if tables else set()
        return Workbook(tables=tables, pkl=dump_tables(tables, key),
                        rendered=render_tables(tables, config.TABLE_TOKEN_BUDGET),
                        frames=describe_frames(tables), labels=labels)

    async def code_path(self, wb: Workbook, question: str, usage: Usage):
        """LLM viết code -> chạy sandbox -> lỗi thì gửi lỗi cho LLM sửa."""
        messages = [{"role": "system", "content": prompts.CODE_SYSTEM},
                    {"role": "user", "content": prompts.CODE_USER.format(
                        tables=wb.rendered, frames=wb.frames, question=question)}]
        code, error = "", ""
        for _ in range(config.MAX_REPAIRS + 1):
            reply = await self.llm.chat(messages, usage, temperature=0.0)
            code = extract_code(reply)
            if not code:
                error = "No python code block with `def solve(tables):` found."
            else:
                res = await run_code(code, wb.pkl)
                if res.get("ok") and is_valid(res.get("result")):
                    return code, res["result"]
                error = res.get("error") or f"Returned invalid result: {res.get('result')!r}"
            if not reply:
                break
            messages += [{"role": "assistant", "content": reply},
                         {"role": "user", "content": prompts.REPAIR_USER.format(error=error)}]
        return code, None

    async def direct_path(self, wb: Workbook, question: str, usage: Usage):
        reply = await self.llm.chat(
            [{"role": "system", "content": prompts.DIRECT_SYSTEM},
             {"role": "user", "content": prompts.DIRECT_USER.format(tables=wb.rendered, question=question)}],
            usage, temperature=0.0)
        return reply, extract_answer(reply)

    async def judge(self, wb, question, code, ans_a, reasoning_b, ans_b, usage) -> str:
        reply = await self.llm.chat(
            [{"role": "system", "content": prompts.JUDGE_SYSTEM},
             {"role": "user", "content": prompts.JUDGE_USER.format(
                 tables=wb.rendered, question=question, code=code, result_a=ans_a,
                 reasoning_b=reasoning_b[-2500:], result_b=ans_b)}],
            usage, temperature=0.0)
        return extract_answer(reply)

    async def solve(self, qid: str, file_name: str, question: str) -> dict:
        t0, usage = time.time(), Usage()
        log = {"id": qid, "file": file_name, "query": question}
        try:
            wb = await self.load(file_name)
            fmt = lambda v: format_answer(v, question, wb.labels)          # noqa: E731

            if config.MODE == "code":
                code, raw_a = await self.code_path(wb, question, usage)
                reasoning_b, raw_b = ("", "") if raw_a is not None else await self.direct_path(wb, question, usage)
            else:
                (code, raw_a), (reasoning_b, raw_b) = await asyncio.gather(
                    self.code_path(wb, question, usage), self.direct_path(wb, question, usage))
            ans_a = fmt(raw_a) if raw_a is not None else ""
            ans_b = fmt(raw_b) if raw_b else ""
            log.update(code=code, raw_a=raw_a, ans_a=ans_a, ans_b=ans_b)

            if ans_a and ans_b and same_answer(ans_a, ans_b):
                answer, source = ans_a, "agree"
            elif ans_a and not ans_b:
                answer, source = ans_a, "code_only"
            elif ans_b and not ans_a:
                answer, source = ans_b, "direct_only"
            elif not ans_a and not ans_b:
                answer, source = "0", "fallback"
            elif time.time() - self.start > config.SOFT_DEADLINE:
                answer, source = ans_a, "code_deadline"
            else:
                raw_j = await self.judge(wb, question, code, ans_a, reasoning_b, ans_b, usage)
                ans_j = fmt(raw_j) if raw_j else ""
                answer, source = (ans_j, "judge") if ans_j else (ans_a, "code_after_judge_fail")
                log["ans_judge"] = ans_j
        except Exception as e:
            answer, source = "0", "error"
            log["error"] = f"{type(e).__name__}: {e}"
        log.update(answer=answer, source=source, seconds=round(time.time() - t0, 2),
                   calls=usage.calls, prompt_tokens=usage.prompt_tokens,
                   completion_tokens=usage.completion_tokens, cost=round(usage.cost, 6))
        return log
