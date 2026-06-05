from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

try:
    from generate_rule_based_daily_report import main as generate_main
except ModuleNotFoundError:
    from scripts.generate_rule_based_daily_report import main as generate_main


WORKSPACE = Path(__file__).resolve().parents[1]
INPUT_ROOT = WORKSPACE / "input"
OUTPUT_ROOT = WORKSPACE / "output" / "rule_based"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Launcher for rule-based street report generation. "
            "Without arguments it prompts for ledger, street, date, and output directory."
        )
    )
    parser.add_argument("ledger", nargs="?", type=Path, help="Source ledger .xls/.xlsx.")
    parser.add_argument("--street", help="Street name, for example: 月坛街道.")
    parser.add_argument("--date", default="", help="Date prefix for filename, for example: 5.17 or 6.3.")
    parser.add_argument("-o", "--output", type=Path, help="Output .docx path.")
    parser.add_argument("--output-dir", type=Path, help="Output directory when --output is omitted.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    ledger = args.ledger or prompt_path("请输入台账 .xls/.xlsx 路径", default_newest_ledger())
    street = args.street or prompt_text("请输入街道名称（例如：月坛街道）")
    date_text = args.date or prompt_text("请输入日期前缀（例如：5.17，可留空）", required=False)
    output = args.output or default_output_path(
        ledger=ledger,
        street=street,
        output_dir=args.output_dir or OUTPUT_ROOT,
        date_text=date_text or infer_date_prefix(ledger),
    )

    print(f"[run] ledger: {ledger}")
    print(f"[run] street: {street}")
    print(f"[run] output: {output}")

    old_argv = sys.argv[:]
    try:
        sys.argv = [
            "generate_rule_based_daily_report.py",
            str(ledger),
            "--street",
            street,
            "-o",
            str(output),
        ]
        return generate_main()
    finally:
        sys.argv = old_argv


def default_output_path(ledger: Path, street: str, output_dir: Path, date_text: str = "") -> Path:
    date_prefix = date_text.strip() or infer_date_prefix(ledger)
    street_name = short_street_name(street)
    filename = f"{date_prefix}区级{street_name}检查日报.docx" if date_prefix else f"区级{street_name}检查日报.docx"
    return output_dir / filename


def short_street_name(street: str) -> str:
    street = street.strip()
    return street[:-2] if street.endswith("街道") else street


def infer_date_prefix(path: Path) -> str:
    text = path.stem
    match = re.search(r"(\d{4})(\d{2})(\d{2})", text)
    if match:
        return f"{int(match.group(2))}.{int(match.group(3))}"
    match = re.search(r"(\d{1,2})月(\d{1,2})日", text)
    if match:
        return f"{int(match.group(1))}.{int(match.group(2))}"
    match = re.search(r"(\d{1,2})[.．](\d{1,2})", text)
    if match:
        return f"{int(match.group(1))}.{int(match.group(2))}"
    return ""


def default_newest_ledger() -> Path | None:
    candidates = [
        path
        for root in (INPUT_ROOT, WORKSPACE / "analysis_samples")
        if root.exists()
        for suffix in ("*.xls", "*.xlsx")
        for path in root.glob(suffix)
        if not path.name.startswith("~$")
    ]
    return max(candidates, key=lambda path: path.stat().st_mtime) if candidates else None


def prompt_path(label: str, default: Path | None = None) -> Path:
    suffix = f" [{default}]" if default else ""
    while True:
        value = input(f"{label}{suffix}: ").strip().strip('"')
        path = Path(value) if value else default
        if path and path.exists():
            return path
        print("[fail] 文件不存在，请重新输入。")


def prompt_text(label: str, required: bool = True) -> str:
    while True:
        value = input(f"{label}: ").strip()
        if value or not required:
            return value
        print("[fail] 不能为空，请重新输入。")


if __name__ == "__main__":
    raise SystemExit(main())
