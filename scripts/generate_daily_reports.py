from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


WORKSPACE = Path(__file__).resolve().parents[1]
OUTPUT_ROOT = WORKSPACE / "output"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate structured extraction and street daily reports from the newest .xls in output/."
    )
    parser.add_argument(
        "input",
        nargs="?",
        type=Path,
        help="Original .xls file. Defaults to newest .xls in output/.",
    )
    parser.add_argument("--overwrite", action="store_true", help="Overwrite generated files.")
    parser.add_argument("--date", help="Optional report date, such as 5.17 or 2026-05-17.")
    parser.add_argument("--template", type=Path, help="Daily report template docx.")
    parser.add_argument("--output-dir", type=Path, help="Output directory for street daily reports.")
    return parser.parse_args()


def newest_xls() -> Path:
    candidates = [
        path
        for path in OUTPUT_ROOT.glob("*.xls")
        if not path.name.startswith("~$")
    ]
    if not candidates:
        raise FileNotFoundError("No .xls file found in output/.")
    return max(candidates, key=lambda path: path.stat().st_mtime)


def run_command(args: list[str]) -> None:
    print("[run] " + " ".join(args))
    subprocess.run(args, cwd=WORKSPACE, check=True)


def main() -> int:
    args = parse_args()
    try:
        source = args.input.resolve() if args.input else newest_xls()
        if not source.exists():
            raise FileNotFoundError(f"Input file does not exist: {source}")

        xlsx = source.with_suffix(".xlsx")
        structured = source.with_name(f"{source.stem}_\u7ed3\u6784\u5316\u63d0\u53d6_\u5408\u5e76\u7f16\u53f7.docx")

        extract_cmd = [
            sys.executable,
            str(WORKSPACE / "scripts" / "extract_xls_structure.py"),
            str(source),
            "--xlsx",
            str(xlsx),
            "-o",
            str(structured),
        ]
        split_cmd = [
            sys.executable,
            str(WORKSPACE / "scripts" / "split_street_daily_reports.py"),
            str(structured),
        ]
        if args.overwrite:
            extract_cmd.append("--overwrite")
            split_cmd.append("--overwrite")
        if args.date:
            split_cmd.extend(["--date", args.date])
        if args.template:
            split_cmd.extend(["--template", str(args.template.resolve())])
        if args.output_dir:
            split_cmd.extend(["--output-dir", str(args.output_dir.resolve())])

        run_command(extract_cmd)
        run_command(split_cmd)
        return 0
    except Exception as exc:
        print(f"[fail] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
