import os
from pathlib import Path


def _load_dotenv(path: Path) -> None:
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, val = line.split("=", 1)
        os.environ.setdefault(key.strip(), val.strip().strip('"').strip("'"))


_load_dotenv(Path(__file__).resolve().parent.parent / ".env")

API_KEY = os.getenv("LLM_API_KEY") or os.getenv("OPENROUTER_API_KEY", "")
BASE_URL = os.getenv("LLM_BASE_URL") or os.getenv("OPENROUTER_BASE_URL", "https://openrouter.ai/api/v1")
MODEL = os.getenv("MODEL", "qwen/qwen3.7-flash")
IS_OPENROUTER = "openrouter.ai" in BASE_URL
REASONING = os.getenv("REASONING", "low")
REASONING_MAX_TOKENS = int(os.getenv("REASONING_MAX_TOKENS", 2048))
MAX_TOKENS = int(os.getenv("MAX_TOKENS", 16384))
LLM_TIMEOUT = float(os.getenv("LLM_TIMEOUT", 180))
LLM_RETRIES = int(os.getenv("LLM_RETRIES", 2))

CONCURRENCY = int(os.getenv("CONCURRENCY", 8))
MAX_REPAIRS = int(os.getenv("MAX_REPAIRS", 2))
SANDBOX_TIMEOUT = float(os.getenv("SANDBOX_TIMEOUT", 15))
SANDBOX_CONCURRENCY = int(os.getenv("SANDBOX_CONCURRENCY", min(8, os.cpu_count() or 4)))
TABLE_TOKEN_BUDGET = int(os.getenv("TABLE_TOKEN_BUDGET", 20000))
SOFT_DEADLINE = float(os.getenv("SOFT_DEADLINE", 30 * 60))
HARD_DEADLINE = float(os.getenv("HARD_DEADLINE", 36 * 60))
LOG_DIR = os.getenv("LOG_DIR", "logs")
MODE = os.getenv("MODE", "full")

CODE_SAMPLES = int(os.getenv("CODE_SAMPLES", 3))
DIRECT_SAMPLES = int(os.getenv("DIRECT_SAMPLES", 2))
SAMPLE_TEMPERATURE = float(os.getenv("SAMPLE_TEMPERATURE", 0.7))
