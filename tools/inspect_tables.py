import argparse
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from src.parser import parse_workbook
from src.render import est_tokens, render_compact, render_full


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("path", help="file .xlsx hoặc thư mục")
    ap.add_argument("--compact", action="store_true", help="in bản rút gọn")
    args = ap.parse_args()
    p = Path(args.path)
    files = sorted(p.glob("*.xlsx")) if p.is_dir() else [p]
    for f in files:
        print("=" * 100, f"\n{f.name}")
        for t in parse_workbook(f):
            text = render_compact(t) if args.compact else render_full(t)
            print(text)
            print(f"-- ~{est_tokens(text)} tokens\n")


if __name__ == "__main__":
    main()
