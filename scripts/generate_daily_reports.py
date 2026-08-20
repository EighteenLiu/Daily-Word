from __future__ import annotations

import argparse
import sys
import shutil
from pathlib import Path


WORKSPACE = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[1]
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
    parser.add_argument("--garbage-summary-template", type=Path, help="Optional garbage classification daily summary template.")
    parser.add_argument("--daily-summary-template", type=Path, help="Optional simple daily summary template.")
    parser.add_argument("--summary-output-dir", type=Path, help="Output directory for generated daily summaries.")
    parser.add_argument("--outside-bucket-file", type=Path, help="Optional outside-bucket Word report to merge into street reports.")
    parser.add_argument(
        "--image-compression",
        choices=("none", "light", "standard", "strong"),
        default="standard",
        help="Compression level for images inserted into generated Word reports.",
    )
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


def resolve_summary_report_date(source: Path, date_value: str | None, split_args, structured: Path):
    import split_street_daily_reports as split

    if date_value and date_value.strip():
        return split.parse_report_date(date_value.strip())
    return split.infer_date_from_filename(source) or split.resolve_report_date(split_args, structured)


def generate_reports(
    source: Path | None = None,
    overwrite: bool = False,
    date_value: str | None = None,
    template: Path | None = None,
    output_dir: Path | None = None,
    transfer_doc: Path | None = None,
    garbage_summary_template: Path | None = None,
    daily_summary_template: Path | None = None,
    summary_output_dir: Path | None = None,
    outside_bucket_path: Path | None = None,
    image_compression: str = "standard",
) -> int:
    try:
        import extract_xls_structure as extract
        import generate_daily_summaries as summaries
        import split_street_daily_reports as split
        from workspace_temp import temporary_directory
        from daily_report_builder import build_street_report, extract_outside_bucket_issues, load_ledger_rows
        from daily_report_renderer import (
            render_street_report_docx,
            render_street_report_docx_from_docxtpl,
        )
        from generate_rule_based_daily_report import has_jinja_tags
        from outside_bucket_parser import ensure_outside_bucket_docx, filter_outside_bucket_items, parse_outside_bucket_docx
    except ModuleNotFoundError:
        from scripts import extract_xls_structure as extract
        from scripts import generate_daily_summaries as summaries
        from scripts import split_street_daily_reports as split
        from scripts.workspace_temp import temporary_directory
        from scripts.daily_report_builder import build_street_report, extract_outside_bucket_issues, load_ledger_rows
        from scripts.daily_report_renderer import (
            render_street_report_docx,
            render_street_report_docx_from_docxtpl,
        )
        from scripts.generate_rule_based_daily_report import has_jinja_tags
        from scripts.outside_bucket_parser import ensure_outside_bucket_docx, filter_outside_bucket_items, parse_outside_bucket_docx

    source = source.resolve() if source else newest_xls()
    if not source.exists():
        raise FileNotFoundError(f"Input file does not exist: {source}")

    XLSX_ROOT.mkdir(parents=True, exist_ok=True)
    EXTRACT_DOCX_ROOT.mkdir(parents=True, exist_ok=True)
    xlsx = XLSX_ROOT / f"{source.stem}.xlsx"
    structured = EXTRACT_DOCX_ROOT / f"{source.stem}_\u7ed3\u6784\u5316\u63d0\u53d6_\u5408\u5e76\u7f16\u53f7.docx"
    cleanup_files: list[Path] = []
    cleanup_dirs: list[Path] = []

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
    image_source_path = source if source.suffix.lower() == ".xls" else None
    with temporary_directory(prefix="xls_extract_images_") as temp:
        sheet_name, items, image_count = extract.extract_items(
            xlsx_path,
            None,
            2,
            Path(temp),
            image_source_path=image_source_path,
        )
        if not items:
            raise RuntimeError("No problem rows were extracted. Check the header row and column names.")
        extract.write_docx(items, structured, source.name, sheet_name)
    if source.suffix.lower() == ".xls" and _is_relative_to(xlsx_path.resolve(), XLSX_ROOT.resolve()):
        cleanup_files.append(xlsx_path)
        cleanup_dirs.append(xlsx_path.parent / "extracted_images" / source.stem)
    elif source.suffix.lower() == ".xlsx":
        cleanup_dirs.append(xlsx_path.parent / "extracted_images" / xlsx_path.stem)
    if _is_relative_to(structured.resolve(), EXTRACT_DOCX_ROOT.resolve()):
        cleanup_files.append(structured)
    embedded_count = sum(len(item.images) for item in items)
    print(f"[ok] xlsx: {xlsx_path}")
    print(f"[ok] docx: {structured}")
    print(f"[ok] extracted rows: {len(items)}")
    print(f"[ok] extracted images: {image_count}, embedded images: {embedded_count}")
    print(f"[ok] report image compression: {image_compression}")

    split_args = argparse.Namespace(
        input=structured,
        output_dir=output_dir,
        template=template or split.REFERENCES_ROOT / "\u65e5\u62a5\u6a21\u7248.docx",
        date=date_value,
        source_xlsx=xlsx_path,
        transfer_doc=transfer_doc,
        overwrite=overwrite,
        include_non_street_headings=False,
    )
    effective_template = split_args.template
    report_date = split.resolve_report_date(split_args, structured)
    street_output_dir = (
        output_dir.resolve()
        if output_dir
        else OUTPUT_ROOT / f"{split.CN_REPORT_DIR}\uff08{report_date.short}\uff09"
    )
    print(
        "[run] rule_based_street_daily_reports "
        f"{xlsx_path}"
        + (" --overwrite" if overwrite else "")
        + (f" --template {effective_template}" if effective_template else "")
        + (f" --output-dir {output_dir}" if output_dir else "")
    )
    rows = load_ledger_rows(xlsx_path, include_images=True, image_source_path=image_source_path)
    streets = _ordered_streets(rows)
    outside_bucket_items = []
    ledger_outside_bucket_items = extract_outside_bucket_issues(rows)
    if ledger_outside_bucket_items:
        image_total = sum(len(item.get("images", [])) for item in ledger_outside_bucket_items)
        print(f"[ok] 台账桶外摆提取完成: 共识别 {len(ledger_outside_bucket_items)} 条问题, 图片 {image_total} 张")
    outside_bucket_temp_dir = None
    if outside_bucket_path and not ledger_outside_bucket_items:
        outside_bucket_path = outside_bucket_path.resolve()
        if not outside_bucket_path.exists():
            raise FileNotFoundError(f"Outside-bucket file does not exist: {outside_bucket_path}")
        outside_bucket_temp_dir = temporary_directory(prefix="outside_bucket_images_")
        print(f"[run] 桶外摆文件: {outside_bucket_path}")
        normalized_outside_bucket_path = ensure_outside_bucket_docx(
            outside_bucket_path,
            Path(outside_bucket_temp_dir.name) / "converted",
        )
        outside_bucket_items = parse_outside_bucket_docx(
            normalized_outside_bucket_path,
            Path(outside_bucket_temp_dir.name) / "images",
        )
        image_total = sum(len(item.image_paths) for item in outside_bucket_items)
        print(f"[ok] 桶外摆解析完成: 共识别 {len(outside_bucket_items)} 条问题, 图片 {image_total} 张")
    elif outside_bucket_path and ledger_outside_bucket_items:
        print("[ok] 已从台账识别桶外摆专项记录，跳过旧桶外摆 Word 回退以避免重复输出")
    effective_template = (
        split.convert_template_to_docx(effective_template.resolve())
        if effective_template and effective_template.exists()
        else effective_template
    )
    use_jinja_template = bool(effective_template and effective_template.exists() and has_jinja_tags(effective_template))
    written = 0
    street_report_paths: list[tuple[str, Path]] = []
    try:
        for street in streets:
            report = build_street_report(rows, street)
            if ledger_outside_bucket_items:
                outside_bucket_matches = [
                    item for item in ledger_outside_bucket_items if str(item.get("street_name") or item.get("street") or "").strip() == street
                ]
            else:
                outside_bucket_matches = filter_outside_bucket_items(outside_bucket_items, [street]) if outside_bucket_items else []
            if not any((report.communities, report.restaurants, report.social_units, outside_bucket_matches)):
                continue
            if outside_bucket_path or ledger_outside_bucket_items:
                matched_images = sum(
                    len(item.get("images", [])) if isinstance(item, dict) else len(item.image_paths)
                    for item in outside_bucket_matches
                )
                print(f"[ok] 桶外摆匹配完成: {street} {len(outside_bucket_matches)} 条, 图片 {matched_images} 张")
                if not outside_bucket_matches:
                    print(f"[warn] 未匹配到当前街道桶外摆问题: {street}")
            filename = f"{report_date.short}{split.CN_DISTRICT}{split.short_street_name(street)}{split.CN_CHECK_REPORT}.docx"
            output_path = street_output_dir / split.safe_filename(filename)
            if output_path.exists() and not overwrite:
                street_report_paths.append((street, output_path))
                print(f"[skip] exists: {output_path}")
                continue
            if use_jinja_template:
                output_path = render_street_report_docx_from_docxtpl(
                    report=report,
                    template_path=effective_template,
                    output_path=output_path,
                    report_title=f"{report_date.short}区级{split.short_street_name(street)}检查日报",
                    report_date_text=report_date.chinese,
                    outside_bucket_issues=outside_bucket_matches,
                    image_compression=image_compression,
                )
            else:
                output_path = render_street_report_docx(
                    report,
                    output_path,
                    outside_bucket_issues=outside_bucket_matches,
                    image_compression=image_compression,
                )
            written += 1
            street_report_paths.append((street, output_path))
            print(f"[ok] {street} -> {output_path}")
    finally:
        if outside_bucket_temp_dir is not None:
            outside_bucket_temp_dir.cleanup()

    print(f"[ok] date: {report_date.short}")
    print(f"[ok] output_dir: {street_output_dir}")
    print(f"[ok] streets: {len(streets)}, written: {written}")

    if garbage_summary_template or daily_summary_template:
        reports = split.parse_structured_docx(
            structured,
            include_non_street_headings=False,
        )
        print(
            "[run] generate_daily_summaries"
            + (f" --garbage-summary-template {garbage_summary_template}" if garbage_summary_template else "")
            + (f" --daily-summary-template {daily_summary_template}" if daily_summary_template else "")
            + (f" --summary-output-dir {summary_output_dir}" if summary_output_dir else "")
        )
        written_summaries = summaries.write_summaries(
            reports=reports,
            report_date=report_date,
            garbage_template=garbage_summary_template,
            daily_template=daily_summary_template,
            output_dir=summary_output_dir,
            street_report_paths=street_report_paths,
            ledger_rows=rows,
        )
        for path in written_summaries:
            print(f"[ok] summary: {path}")

    cleanup_generated_intermediates(cleanup_files, cleanup_dirs)

    return 0


def _is_relative_to(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
        return True
    except ValueError:
        return False


def cleanup_generated_intermediates(files: list[Path], dirs: list[Path]) -> None:
    for path in files:
        try:
            if path.exists():
                path.unlink()
                print(f"[clean] removed intermediate: {path}")
        except PermissionError:
            print(f"[warn] intermediate is occupied, cannot remove: {path}", file=sys.stderr)
    for directory in dirs:
        try:
            if directory.exists():
                shutil.rmtree(directory)
                print(f"[clean] removed intermediate dir: {directory}")
        except PermissionError:
            print(f"[warn] intermediate dir is occupied, cannot remove: {directory}", file=sys.stderr)


def _ordered_streets(rows) -> list[str]:
    streets: list[str] = []
    seen: set[str] = set()
    for row in rows:
        street = row.street.strip()
        if not street or street in seen:
            continue
        seen.add(street)
        streets.append(street)
    try:
        from garbage_daily_report import CANONICAL_STREET_ORDER
    except ModuleNotFoundError:
        from scripts.garbage_daily_report import CANONICAL_STREET_ORDER

    order = {street: index for index, street in enumerate(CANONICAL_STREET_ORDER)}
    return sorted(streets, key=lambda street: (order.get(street, len(order)), street))


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
            garbage_summary_template=args.garbage_summary_template,
            daily_summary_template=args.daily_summary_template,
            summary_output_dir=args.summary_output_dir,
            outside_bucket_path=args.outside_bucket_file,
            image_compression=args.image_compression,
        )
    except Exception as exc:
        print(f"[fail] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
