from __future__ import annotations

import asyncio
import hashlib
import math
import re
import time
from dataclasses import dataclass
from pathlib import Path

from . import config, prompts
from .formatter import format_answer, near_miss, same_answer
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
    sheets: dict[str, str]


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


@dataclass
class Cand:
    answer: str
    kind: str
    code: str = ""
    reasoning: str = ""


@dataclass
class Cluster:
    members: list[Cand]

    @property
    def n_code(self) -> int:
        return sum(c.kind == "code" for c in self.members)

    @property
    def n_direct(self) -> int:
        return sum(c.kind == "direct" for c in self.members)

    @property
    def weight(self) -> int:
        return len(self.members)

    @property
    def answer(self) -> str:
        return next((c.answer for c in self.members if c.kind == "code"), self.members[0].answer)

    @property
    def code(self) -> str:
        return next((c.code for c in self.members if c.code), "")

    @property
    def reasoning(self) -> str:
        return next((c.reasoning for c in self.members if c.reasoning), "")


def vote(cands: list[Cand]) -> list[Cluster]:
    clusters: list[Cluster] = []
    for c in cands:
        cl = next((cl for cl in clusters if same_answer(c.answer, cl.answer) or near_miss(cl.answer, c.answer)), None)
        if cl:
            cl.members.append(c)
        else:
            clusters.append(Cluster([c]))
    return sorted(clusters, key=lambda cl: (-cl.weight, -cl.n_code))


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
                        frames=describe_frames(tables), labels=labels, sheets={t.id: t.sheet for t in tables})

    async def code_path(self, wb: Workbook, question: str, usage: Usage, temperature: float = 0.0):
        messages = [{"role": "system", "content": prompts.CODE_SYSTEM},
                    {"role": "user", "content": prompts.CODE_USER.format(
                        tables=wb.rendered, frames=wb.frames, question=question)}]
        code, error = "", ""
        for _ in range(config.MAX_REPAIRS + 1):
            reply = await self.llm.chat(messages, usage, temperature=temperature)
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

    async def direct_path(self, wb: Workbook, question: str, usage: Usage, temperature: float = 0.0):
        reply = await self.llm.chat(
            [{"role": "system", "content": prompts.DIRECT_SYSTEM},
             {"role": "user", "content": prompts.DIRECT_USER.format(tables=wb.rendered, question=question)}],
            usage, temperature=temperature)
        return reply, extract_answer(reply)

    async def judge(self, wb, question, clusters: list[Cluster], usage) -> str:
        blocks = []
        for i, cl in enumerate(clusters[:4], 1):
            lines = [f"Candidate {i}: {cl.answer}",
                     f"  supported by {cl.n_code} pandas-code solver(s) and {cl.n_direct} table-reading solver(s)"]
            if cl.code:
                lines.append(f"  code of one supporter:\n```python\n{cl.code[:2500]}\n```")
            if cl.reasoning:
                lines.append(f"  reasoning of one supporter (end):\n{cl.reasoning[-1500:]}")
            blocks.append("\n".join(lines))
        reply = await self.llm.chat(
            [{"role": "system", "content": prompts.JUDGE_SYSTEM},
             {"role": "user", "content": prompts.JUDGE_USER.format(
                 tables=wb.rendered, question=question, candidates="\n\n".join(blocks))}],
            usage, temperature=0.0)
        return extract_answer(reply)

    async def solve(self, qid: str, file_name: str, question: str) -> dict:
        t0, usage = time.time(), Usage()
        log = {"id": qid, "file": file_name, "query": question}
        try:
            wb = await self.load(file_name)
            fmt = lambda v: format_answer(v, question, wb.labels, wb.sheets)
            late = time.time() - self.start > config.SOFT_DEADLINE
            n_code, n_direct = (1, 1) if late else (config.CODE_SAMPLES, config.DIRECT_SAMPLES)
            temps = lambda n: [0.0][:n] + [config.SAMPLE_TEMPERATURE] * (n - 1)

            if config.MODE == "code":
                codes = [await self.code_path(wb, question, usage)]
                directs = [] if codes[0][1] is not None else [await self.direct_path(wb, question, usage)]
            else:
                codes, directs = await asyncio.gather(
                    asyncio.gather(*(self.code_path(wb, question, usage, t) for t in temps(n_code))),
                    asyncio.gather(*(self.direct_path(wb, question, usage, t) for t in temps(n_direct))))
            cands = [Cand(fmt(raw), "code", code=code) for code, raw in codes if raw is not None]
            cands += [Cand(fmt(ans), "direct", reasoning=reply) for reply, ans in directs if ans]
            clusters = vote([c for c in cands if c.answer])
            log.update(code=codes[0][0], raw_a=codes[0][1],
                       ans_a=fmt(codes[0][1]) if codes[0][1] is not None else "",
                       ans_b=fmt(directs[0][1]) if directs and directs[0][1] else "",
                       candidates=[[cl.answer, cl.n_code, cl.n_direct] for cl in clusters])

            if not clusters:
                answer, source = "0", "fallback"
            elif len(clusters) == 1:
                cl = clusters[0]
                answer = cl.answer
                source = "agree" if cl.n_code and cl.n_direct else ("code_only" if cl.n_code else "direct_only")
            elif clusters[0].n_code and clusters[0].n_direct and clusters[0].weight > clusters[1].weight:
                answer, source = clusters[0].answer, "vote"
            elif clusters[0].n_code >= 2 and clusters[0].n_code == sum(c.kind == "code" for c in cands):
                answer, source = clusters[0].answer, "code_unanimous"
            elif late:
                answer, source = clusters[0].answer, "deadline"
            else:
                raw_j = await self.judge(wb, question, clusters, usage)
                ans_j = fmt(raw_j) if raw_j else ""
                match = next((cl for cl in clusters if same_answer(ans_j, cl.answer) or near_miss(cl.answer, ans_j)),
                             None) if ans_j else None
                answer = match.answer if match else (ans_j or clusters[0].answer)
                source = "judge" if ans_j else "judge_fail"
                log["ans_judge"] = ans_j
        except Exception as e:
            answer, source = "0", "error"
            log["error"] = f"{type(e).__name__}: {e}"
        log.update(answer=answer, source=source, seconds=round(time.time() - t0, 2),
                   calls=usage.calls, prompt_tokens=usage.prompt_tokens,
                   completion_tokens=usage.completion_tokens, reasoning_tokens=usage.reasoning_tokens,
                   max_prompt=usage.max_prompt, cost=round(usage.cost, 6))
        return log
