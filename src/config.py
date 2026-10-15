"""Cấu hình đọc từ biến môi trường hoặc file .env."""
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

# Mặc định: Qwen qua OpenRouter (bắt buộc khi chấm). LLM_* dùng để thử với nhà cung cấp khác (vd Gemini).
API_KEY = os.getenv("LLM_API_KEY") or os.getenv("OPENROUTER_API_KEY", "")
BASE_URL = os.getenv("LLM_BASE_URL") or os.getenv("OPENROUTER_BASE_URL", "https://openrouter.ai/api/v1")
MODEL = os.getenv("MODEL", "qwen/qwen3.7-flash")
IS_OPENROUTER = "openrouter.ai" in BASE_URL
REASONING = os.getenv("REASONING", "low")          # off | low | medium | high | default
MAX_TOKENS = int(os.getenv("MAX_TOKENS", 4096))
LLM_TIMEOUT = float(os.getenv("LLM_TIMEOUT", 60))
LLM_RETRIES = int(os.getenv("LLM_RETRIES", 2))

CONCURRENCY = int(os.getenv("CONCURRENCY", 8))
MAX_REPAIRS = int(os.getenv("MAX_REPAIRS", 2))
SANDBOX_TIMEOUT = float(os.getenv("SANDBOX_TIMEOUT", 15))
TABLE_TOKEN_BUDGET = int(os.getenv("TABLE_TOKEN_BUDGET", 20000))   # giữ prompt < 32K để được giá rẻ
SOFT_DEADLINE = float(os.getenv("SOFT_DEADLINE", 30 * 60))         # quá mốc này thì bỏ bước giám khảo
MODE = os.getenv("MODE", "full")                                   # full | code
