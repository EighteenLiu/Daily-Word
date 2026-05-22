from __future__ import annotations

import argparse
import sys
import tempfile
from pathlib import Path


WORKSPACE = Path(__file__).resolve().parents[1]
INPUT_ROOT = WORKSPACE / "input"
OUTPUT_ROOT = WORKSPACE / "output"
NEW_REFERENCES_ROOT = INPUT_ROOT
NEW_XLS_ROOT = NEW_REFERENCES_ROOT / "xls\u6587\u4ef6"
XLSX_ROOT = OUTPUT_ROOT / "\u8f6c\u6362xlsx"
EXTRACT_DOCX_ROOT = OUTPUT_ROOT / "\u63d0\u53d6docx"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate structured extraction and street daily reports from the newest .xls in input/."
    )
    parser.add_argument(
        "input",
        nargs="?",
        type=Path,
        help="Original .xls file. Defaults to newest .xls in input/ or input/xls文件/.",
    )
    parser.add_argument("--overwrite", action="store_true", help="Overwrite generated files.")
    parser.add_argument("--date", help="Optional report date, such as 5.17 or 2026-05-17.")
    parser.add_argument("--template", type=Path, help="Daily report template docx.")
    parser.add_argument("--output-dir", type=Path, help="Output directory for street daily reports.")
    parser.add_argument("--transfer-doc", type=Path, help="Optional transfer-station daily report .doc/.docx.")
    return parser.parse_args()


def newest_xls() -> Path:
    candidates = [
        path
        for path in NEW_XLS_ROOT.glob("*.xls")
        if not path.name.startswith("~$")
    ]
    if not candidates:
        candidates = [
            path
            for path in INPUT_ROOT.glob("*.xls")
            if not path.name.startswith("~$")
        ]
    if not candidates:
        candidates = [
            path
            for path in OUTPUT_ROOT.glob("*.xls")
            if not path.name.startswith("~$")
        ]
    if not candidates:
        raise FileNotFoundError("No .xls file found in input/ or input/xls文件/.")
    return max(candidates, key=lambda path: path.stat().st_mtime)


def generate_reports(
    source: Path | None = None,
    overwrite: bool = False,
    date_value: str | None = None,
    template: Path | None = None,
    output_dir: Path | None = None,
    transfer_doc: Path | None = None,
) -> int:
    import extract_xls_structure as extract
    import split_street_daily_reports as split

    source = source.resolve() if source else newest_xls()
    if not source.exists():
        raise FileNotFoundError(f"Input file does not exist: {source}")

    XLSX_ROOT.mkdir(parents=True, exist_ok=True)
    EXTRACT_DOCX_ROOT.mkdir(parents=True, exist_ok=True)
    xlsx = XLSX_ROOT / f"{source.stem}.xlsx"
    structured = EXTRACT_DOCX_ROOT / f"{source.stem}_\u7ed3\u6784\u5316\u63d0\u53d6_\u5408\u5e76\u7f16\u53f7.docx"

    extract_args = argparse.Namespace(
        input=source,
        output=structured,
        xlsx=xlsx,
        sheet=None,
        header_row=2,
        overwrite=overwrite,
        visible=False,
        no_registry_protection=False,
        keep_registry_setting=False,
    )
    print(
        "[run] extract_xls_structure "
        f"{source} --xlsx {xlsx} -o {structured}"
        + (" --overwrite" if overwrite else "")
    )
    xlsx_path = extract.ensure_xlsx(extract_args)
    if structured.exists() and not overwrite:
        raise FileExistsError(f"Output already exists. Use --overwrite: {structured}")
    with tempfile.TemporaryDirectory(prefix="xls_extract_images_") as temp:
        sheet_name, items, image_count = extract.extract_items(xlsx_path, None, 2, Path(temp))
        if not items:
            raise RuntimeError("No problem rows were extracted. Check the header row and column names.")
        extract.write_docx(items, structured, source.name, sheet_name)
    embedded_count = sum(len(item.images) for item in items)
    print(f"[ok] xlsx: {xlsx_path}")
    print(f"[ok] docx: {structured}")
    print(f"[ok] extracted rows: {len(items)}")
    print(f"[ok] extracted images: {image_count}, embedded images: {embedded_count}")

    split_args = argparse.Namespace(
        input=structured,
        output_dir=output_dir,
        template=template or split.REFERENCES_ROOT / "\u65e5\u62a5\u6a21\u7248.docx",
        date=date_value,
        source_xlsx=xlsx,
        transfer_doc=transfer_doc,
        overwrite=overwrite,
        include_non_street_headings=False,
    )
    print(
        "[run] split_street_daily_reports "
        f"{structured}"
        + (" --overwrite" if overwrite else "")
        + (f" --template {split_args.template}" if split_args.template else "")
        + (f" --output-dir {output_dir}" if output_dir else "")
        + (f" --transfer-doc {transfer_doc}" if transfer_doc else "")
    )
    return split.run_split(split_args)


def main() -> int:
    args = parse_args()
    try:
        return generate_reports(
            source=args.input,
            overwrite=args.overwrite,
            date_value=args.date,
            template=args.template,
            output_dir=args.output_dir,
            transfer_doc=args.transfer_doc,
        )
    except Exception as exc:
        print(f"[fail] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
