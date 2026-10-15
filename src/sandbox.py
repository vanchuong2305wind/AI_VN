"""Bước ④: chạy code của LLM trong tiến trình con, có timeout."""
from __future__ import annotations

import asyncio
import json
import os
import pickle
import sys
import tempfile
import uuid
from pathlib import Path

from . import config
from .parser import Table

RUNNER = Path(__file__).with_name("sandbox_runner.py")
WORKDIR = Path(tempfile.gettempdir()) / "tableqa_sandbox"


def dump_tables(tables: list[Table], key: str) -> Path:
    WORKDIR.mkdir(exist_ok=True)
    path = WORKDIR / f"{key}.pkl"
    frames = {}
    for t in tables:
        frames[t.id] = t.df
        if t.long is not None:
            frames[f"{t.id}_long"] = t.long
    with open(path, "wb") as f:
        pickle.dump(frames, f)
    return path


async def run_code(code: str, tables_pkl: Path) -> dict:
    """-> {"ok": True, "result": ...} hoặc {"ok": False, "error": "..."}"""
    code_path = WORKDIR / f"sol_{uuid.uuid4().hex}.py"
    code_path.write_text(code, encoding="utf-8")
    proc = await asyncio.create_subprocess_exec(
        sys.executable, str(RUNNER), str(tables_pkl), str(code_path),
        cwd=str(WORKDIR), stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE)
    try:
        out, err = await asyncio.wait_for(proc.communicate(), timeout=config.SANDBOX_TIMEOUT)
    except asyncio.TimeoutError:
        proc.kill()
        return {"ok": False, "error": f"Timeout after {config.SANDBOX_TIMEOUT}s"}
    finally:
        try:
            os.remove(code_path)
        except OSError:
            pass
    lines = out.decode("utf-8", "replace").strip().splitlines()
    try:
        return json.loads(lines[-1])
    except (IndexError, json.JSONDecodeError):
        return {"ok": False, "error": (err.decode("utf-8", "replace") or "no output")[-1500:]}
