from __future__ import annotations

import argparse
import hashlib
import shutil
import sys
from pathlib import Path

try:
    from daily_report_builder import build_street_report, load_ledger_rows
    from daily_report_renderer import render_street_report_docx, render_street_report_docx_from_docxtpl
except ModuleNotFoundError:
    from scripts.daily_report_builder import build_street_report, load_ledger_rows
    from scripts.daily_report_renderer import render_street_report_docx, render_street_report_docx_from_docxtpl


def has_jinja_tags(path: Path) -> bool:
    """Check if a .docx file has Jinja2 template tags in its body."""
    try:
        from docx import Document
        doc = Document(str(path))
        for p in doc.paragraphs:
            text = p.text
            if "{{" in text or "{%" in text:
                return True
    except Exception:
        pass
    return False


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate a street daily report from the rule-based ledger mapping.")
    parser.add_argument("ledger", type=Path, help="Source ledger .xls/.xlsx file.")
    parser.add_argument("--street", required=True, help="Street name, for example: \u6708\u575b\u8857\u9053.")
    parser.add_argument("-o", "--output", type=Path, required=True, help="Output .docx path.")
    parser.add_argument("--template", type=Path, default=None, help="Optional Jinja2 .docx template.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        source = args.ledger.resolve()
        ledger = prepare_ledger(args.ledger, args.output.parent)
        image_source_path = source if source.suffix.lower() == ".xls" else None
        rows = load_ledger_rows(ledger, include_images=True, image_source_path=image_source_path)
        report = build_street_report(rows, args.street)
        if not any((report.communities, report.restaurants, report.social_units)):
            raise RuntimeError(f"No report content found for street: {args.street}")

        template = normalize_template(args.template) if args.template else None
        if template and has_jinja_tags(template):
            output = render_street_report_docx_from_docxtpl(
                report=report,
                template_path=template,
                output_path=args.output,
            )
        else:
            output = render_street_report_docx(report, args.output)
        
        print(f"[ok] output: {output}")
        return 0
    except Exception as exc:
        print(f"[fail] {exc}", file=sys.stderr)
        return 1


def prepare_ledger(source: Path, output_dir: Path) -> Path:
    source = source.resolve()
    if source.suffix.lower() == ".xlsx":
        return source
    if source.suffix.lower() != ".xls":
        raise ValueError(f"Ledger must be .xls or .xlsx: {source}")

    ascii_xls, ascii_xlsx = intermediate_paths_for_source(source, output_dir)
    ascii_xls.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, ascii_xls)
    if ascii_xlsx.exists() and ascii_xlsx.stat().st_mtime >= ascii_xls.stat().st_mtime:
        return ascii_xlsx

    import win32com.client

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        workbook = excel.Workbooks.Open(str(ascii_xls), ReadOnly=True)
        workbook.SaveAs(str(ascii_xlsx), FileFormat=51)
        workbook.Close(False)
    finally:
        excel.Quit()
    return ascii_xlsx


def intermediate_paths_for_source(source: Path, output_dir: Path) -> tuple[Path, Path]:
    source = source.resolve()
    stat = source.stat()
    key = f"{source}|{stat.st_size}|{stat.st_mtime_ns}".encode("utf-8", errors="surrogatepass")
    digest = hashlib.sha1(key).hexdigest()[:12]
    intermediate_dir = output_dir.resolve() / "_intermediate"
    return intermediate_dir / f"ledger_{digest}.xls", intermediate_dir / f"ledger_{digest}.xlsx"


def normalize_template(template: Path) -> Path:
    template = template.resolve()
    if not template.exists():
        raise FileNotFoundError(f"Template does not exist: {template}")
    if template.suffix.lower() == ".docx":
        return template
    if template.suffix.lower() != ".doc":
        raise ValueError(f"Template must be .doc or .docx: {template}")

    import win32com.client

    destination = template.with_suffix(".docx")
    if destination.exists() and destination.stat().st_mtime >= template.stat().st_mtime:
        return destination

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        document = word.Documents.Open(str(template))
        document.SaveAs(str(destination), FileFormat=16)
        document.Close(False)
    finally:
        word.Quit()
    return destination


if __name__ == "__main__":
    raise SystemExit(main())
