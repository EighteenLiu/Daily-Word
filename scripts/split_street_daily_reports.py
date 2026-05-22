from __future__ import annotations

import argparse
import hashlib
import re
import sys
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path


WORKSPACE = Path(__file__).resolve().parents[1]
OUTPUT_ROOT = WORKSPACE / "output"
REFERENCES_ROOT = WORKSPACE / "references"

CN_SONG = "\u5b8b\u4f53"
CN_STREET = "\u8857\u9053"
CN_REPORT_DIR = "\u8857\u9053\u65e5\u62a5"
CN_DISTRICT = "\u533a\u7ea7"
CN_CHECK_REPORT = "\u68c0\u67e5\u65e5\u62a5"
CN_NO_CONTENT = "\u65e0"
STYLE_HEADING_1 = "\u8bba\u65871"
STYLE_HEADING_2 = "\u8bba\u65871.1"
STYLE_HEADING_3 = "\u8bba\u65871.1.1"
STYLE_BODY = "\u5b8b5\u6b63\u6587"
TARGET_CATEGORIES = (
    "\u5c45\u4f4f\u5c0f\u533a\u3001\u5e73\u623f\u80e1\u540c",
    "\u9910\u996e\u5355\u4f4d",
    "\u793e\u4f1a\u5355\u4f4d",
)


@dataclass
class ImageBlob:
    blob: bytes
    content_type: str
    suffix: str
    name: str


@dataclass
class ContentBlock:
    text: str = ""
    images: list[ImageBlob] = field(default_factory=list)


@dataclass
class PlaceBlock:
    name: str
    blocks: list[ContentBlock] = field(default_factory=list)


@dataclass
class ReportDate:
    value: date

    @property
    def short(self) -> str:
        return f"{self.value.month}.{self.value.day}"

    @property
    def chinese(self) -> str:
        return f"{self.value.year}\u5e74{self.value.month}\u6708{self.value.day}\u65e5"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Split the generated structured docx into one street daily report per street. "
            "Defaults are normalized for this project: template from references, source from output, "
            "and reports written to output/street-daily-report(date)."
        )
    )
    parser.add_argument(
        "input",
        nargs="?",
        type=Path,
        help="Structured .docx file. Defaults to the newest structured .docx in output/.",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        help="Output directory. Defaults to output/街道日报（日期）.",
    )
    parser.add_argument(
        "--template",
        type=Path,
        default=REFERENCES_ROOT / "\u65e5\u62a5\u6a21\u7248.docx",
        help="Daily report template docx. Defaults to references/日报模版.docx.",
    )
    parser.add_argument(
        "--date",
        help="Report date such as 5.17, 2026-05-17, or 20260517. Auto-detected when omitted.",
    )
    parser.add_argument(
        "--source-xlsx",
        type=Path,
        help="Converted xlsx used only for date inference. Defaults to the xlsx matching the input docx prefix.",
    )
    parser.add_argument("--overwrite", action="store_true", help="Overwrite existing report files.")
    parser.add_argument(
        "--include-non-street-headings",
        action="store_true",
        help="Also generate reports for Heading 2 names that do not contain '街道'.",
    )
    return parser.parse_args()


def strip_heading_number(text: str) -> str:
    text = text.strip()
    patterns = (
        r"^[\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e\u96f6]+[\u3001.．]\s*",
        r"^[（(][\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e\u96f6]+[）)]\s*",
        r"^\d+[.．、]\s*",
        r"^[（(]\d+[）)]\s*",
    )
    for pattern in patterns:
        text = re.sub(pattern, "", text)
    return text.strip()


def chinese_numeral(number: int) -> str:
    digits = "\u96f6\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d"
    units = ["", "\u5341", "\u767e", "\u5343"]
    if number <= 0:
        return str(number)
    if number < 10:
        return digits[number]
    if number < 20:
        return "\u5341" + (digits[number % 10] if number % 10 else "")

    chars: list[str] = []
    zero_pending = False
    text = str(number)
    length = len(text)
    for index, char in enumerate(text):
        digit = int(char)
        unit_index = length - index - 1
        if digit == 0:
            zero_pending = bool(chars)
            continue
        if zero_pending:
            chars.append(digits[0])
            zero_pending = False
        chars.append(digits[digit] + units[unit_index])
    return "".join(chars)


def safe_filename(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*\r\n\t]+', "_", name.strip())
    name = re.sub(r"\s+", " ", name).strip(" .")
    return name or "report"


def short_street_name(street: str) -> str:
    return street[:-2] if street.endswith(CN_STREET) else street


def find_default_input() -> Path:
    candidates = [
        path
        for path in OUTPUT_ROOT.glob("*.docx")
        if not path.name.startswith("~$")
        and "\u7ed3\u6784\u5316\u63d0\u53d6" in path.name
    ]
    if not candidates:
        raise FileNotFoundError("No structured extraction .docx found in output/.")
    return max(candidates, key=lambda path: path.stat().st_mtime)


def parse_report_date(value: str) -> ReportDate:
    value = value.strip()
    if re.fullmatch(r"\d{8}", value):
        return ReportDate(datetime.strptime(value, "%Y%m%d").date())
    if re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}", value):
        return ReportDate(datetime.strptime(value, "%Y-%m-%d").date())
    if re.fullmatch(r"\d{1,2}\.\d{1,2}", value):
        month, day = value.split(".")
        return ReportDate(date(datetime.now().year, int(month), int(day)))
    raise ValueError(f"Unsupported date format: {value}")


def matching_xlsx_for_docx(input_path: Path) -> Path | None:
    stem = input_path.stem.split("_\u7ed3\u6784\u5316\u63d0\u53d6", 1)[0]
    candidate = input_path.with_name(f"{stem}.xlsx")
    if candidate.exists():
        return candidate
    matches = sorted(OUTPUT_ROOT.glob("*.xlsx"), key=lambda path: path.stat().st_mtime, reverse=True)
    return matches[0] if matches else None


def infer_date_from_xlsx(xlsx_path: Path) -> ReportDate | None:
    try:
        import openpyxl
    except ImportError:
        return None

    workbook = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        sheet = workbook[workbook.sheetnames[0]]
        for row in range(1, min(sheet.max_row, 30) + 1):
            for col in range(1, min(sheet.max_column, 10) + 1):
                value = sheet.cell(row, col).value
                match = re.search(r"(20\d{6})", str(value or ""))
                if match:
                    return parse_report_date(match.group(1))
    finally:
        workbook.close()
    return None


def infer_date_from_filename(path: Path) -> ReportDate | None:
    match = re.search(r"(20\d{2})[-_]?(\d{2})[-_]?(\d{2})", path.name)
    if not match:
        return None
    detected = date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    return ReportDate(detected - timedelta(days=1))


def resolve_report_date(args: argparse.Namespace, input_path: Path) -> ReportDate:
    if args.date:
        return parse_report_date(args.date)

    source_xlsx = args.source_xlsx.resolve() if args.source_xlsx else matching_xlsx_for_docx(input_path)
    if source_xlsx and source_xlsx.exists():
        inferred = infer_date_from_xlsx(source_xlsx)
        if inferred:
            return inferred

    inferred = infer_date_from_filename(input_path)
    if inferred:
        return inferred

    raise RuntimeError("Could not infer report date. Pass --date 5.17 or --date 2026-05-17.")


def extract_paragraph_images(paragraph) -> list[ImageBlob]:
    from docx.oxml.ns import qn

    images: list[ImageBlob] = []
    for blip in paragraph._p.xpath('.//*[local-name()="blip"]'):
        r_id = blip.get(qn("r:embed"))
        if not r_id:
            continue
        image_part = paragraph.part.related_parts.get(r_id)
        if image_part is None:
            continue
        suffix = "." + image_part.partname.ext
        images.append(
            ImageBlob(
                blob=image_part.blob,
                content_type=image_part.content_type,
                suffix=suffix,
                name=Path(str(image_part.partname)).name,
            )
        )
    return images


def parse_structured_docx(
    input_path: Path,
    include_non_street_headings: bool,
) -> dict[str, dict[str, dict[str, PlaceBlock]]]:
    from docx import Document

    document = Document(input_path)
    reports: dict[str, dict[str, dict[str, PlaceBlock]]] = {}
    current_category = ""
    current_street = ""
    current_place = ""

    for paragraph in document.paragraphs:
        style = paragraph.style.name
        text = paragraph.text.strip()

        if style == "Heading 1":
            current_category = strip_heading_number(text)
            current_street = ""
            current_place = ""
            continue

        if style == "Heading 2":
            current_street = strip_heading_number(text)
            current_place = ""
            if not include_non_street_headings and CN_STREET not in current_street:
                current_street = ""
                continue
            if current_category in TARGET_CATEGORIES and current_street:
                reports.setdefault(current_street, {}).setdefault(current_category, {})
            continue

        if style == "Heading 3":
            current_place = strip_heading_number(text)
            if current_category in TARGET_CATEGORIES and current_street and current_place:
                category = reports.setdefault(current_street, {}).setdefault(current_category, {})
                category.setdefault(current_place, PlaceBlock(current_place))
            continue

        if current_category not in TARGET_CATEGORIES or not current_street or not current_place:
            continue

        images = extract_paragraph_images(paragraph)
        if not text and not images:
            continue

        category = reports.setdefault(current_street, {}).setdefault(current_category, {})
        place = category.setdefault(current_place, PlaceBlock(current_place))
        place.blocks.append(ContentBlock(text=text, images=images))

    return reports


def remove_template_body_content(document) -> None:
    body = document._body._element
    for element in list(body):
        if element.tag.endswith("sectPr"):
            continue
        if element.tag.endswith("tbl"):
            continue
        body.remove(element)


def update_template_date(document, report_date: ReportDate) -> None:
    date_pattern = re.compile(r"\d{4}\u5e74\d{1,2}\u6708\d{1,2}\u65e5")
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.text = date_pattern.sub(report_date.chinese, run.text)


def set_document_styles(document) -> None:
    from docx.oxml.ns import qn
    from docx.shared import Pt

    normal = document.styles["Normal"]
    normal.font.name = CN_SONG
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), CN_SONG)
    normal.font.size = Pt(10.5)


def new_document(template_path: Path, report_date: ReportDate):
    from docx import Document

    if template_path.exists():
        document = Document(template_path)
        update_template_date(document, report_date)
        remove_template_body_content(document)
    else:
        document = Document()
    set_document_styles(document)
    return document


def add_raw_picture_from_blob(run, image: ImageBlob, width: int) -> None:
    from PIL import Image
    from docx.oxml.shape import CT_Inline
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.parts.image import ImagePart
    from docx.shared import Emu

    try:
        with Image.open(BytesIO(image.blob)) as pil_image:
            pixel_width, pixel_height = pil_image.size
    except Exception as exc:
        raise RuntimeError(f"Cannot read image dimensions without recompression: {image.name}") from exc

    image_parts = run.part.package.image_parts
    digest = hashlib.sha1(image.blob).hexdigest()
    image_part = image_parts._get_by_sha1(digest)
    if image_part is None:
        partname = image_parts._next_image_partname(image.suffix.lstrip("."))
        image_part = ImagePart(partname, image.content_type, image.blob)
        image_parts.append(image_part)

    r_id = run.part.relate_to(image_part, RT.IMAGE)
    height = Emu(int(width * pixel_height / pixel_width))
    inline = CT_Inline.new_pic_inline(run.part.next_id, r_id, image.name, width, height)
    run._r.add_drawing(inline)


def add_images(document, images: list[ImageBlob]) -> None:
    if not images:
        return

    from docx.enum.text import WD_ALIGN_PARAGRAPH

    section = document.sections[0]
    usable_width = section.page_width - section.left_margin - section.right_margin
    image_width = int(usable_width / 3 * 0.92)

    for start in range(0, len(images), 3):
        paragraph = document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for image in images[start:start + 3]:
            run = paragraph.add_run()
            add_raw_picture_from_blob(run, image, image_width)
            run.add_text(" ")


def add_report_title(document, street: str, report_date: ReportDate) -> None:
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt

    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(f"{report_date.short}{CN_DISTRICT}{short_street_name(street)}{CN_CHECK_REPORT}")
    run.bold = True
    run.font.size = Pt(18)


def add_section_heading(document, text: str, level: int) -> None:
    from docx.oxml.ns import qn
    from docx.shared import Pt

    font_sizes = {1: 16, 2: 14, 3: 12}
    style_names = {1: STYLE_HEADING_1, 2: STYLE_HEADING_2, 3: STYLE_HEADING_3}
    paragraph = document.add_paragraph()
    style_name = style_names.get(level)
    if style_name and style_name in [style.name for style in document.styles]:
        paragraph.style = style_name
    paragraph.paragraph_format.space_before = Pt(6 if level == 1 else 3)
    paragraph.paragraph_format.space_after = Pt(3)
    if level == 2:
        paragraph.paragraph_format.left_indent = Pt(10.5)
    elif level >= 3:
        paragraph.paragraph_format.left_indent = Pt(21)
    run = paragraph.add_run(text)
    run.bold = True
    run.font.name = CN_SONG
    run._element.rPr.rFonts.set(qn("w:eastAsia"), CN_SONG)
    run.font.size = Pt(font_sizes.get(level, 12))


def write_report(
    output_path: Path,
    template_path: Path,
    report_date: ReportDate,
    street: str,
    categories: dict[str, dict[str, PlaceBlock]],
) -> None:
    document = new_document(template_path, report_date)
    add_report_title(document, street, report_date)

    visible_category_index = 0
    for category_name in TARGET_CATEGORIES:
        places = categories.get(category_name, {})
        if not places:
            continue

        visible_category_index += 1
        add_section_heading(document, f"{chinese_numeral(visible_category_index)}\u3001{category_name}", level=1)

        for place_index, place in enumerate(places.values(), start=1):
            add_section_heading(document, f"\uff08{chinese_numeral(place_index)}\uff09{place.name}", level=2)
            for block in place.blocks:
                if block.text:
                    paragraph = document.add_paragraph(block.text)
                    if STYLE_BODY in [style.name for style in document.styles]:
                        paragraph.style = STYLE_BODY
                add_images(document, block.images)

    if visible_category_index == 0:
        document.add_paragraph(CN_NO_CONTENT)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)


def main() -> int:
    args = parse_args()
    try:
        input_path = args.input.resolve() if args.input else find_default_input()
        if not input_path.exists():
            raise FileNotFoundError(f"Input file does not exist: {input_path}")

        report_date = resolve_report_date(args, input_path)
        output_dir = (
            args.output_dir.resolve()
            if args.output_dir
            else OUTPUT_ROOT / f"{CN_REPORT_DIR}\uff08{report_date.short}\uff09"
        )
        template_path = args.template.resolve()
        reports = parse_structured_docx(input_path, include_non_street_headings=args.include_non_street_headings)
        if not reports:
            raise RuntimeError("No street content found. Check Heading 1/2/3 structure.")

        written = 0
        for street, categories in reports.items():
            filename = f"{report_date.short}{CN_DISTRICT}{short_street_name(street)}{CN_CHECK_REPORT}.docx"
            output_path = output_dir / safe_filename(filename)
            if output_path.exists() and not args.overwrite:
                print(f"[skip] exists: {output_path}")
                continue
            write_report(output_path, template_path, report_date, street, categories)
            written += 1
            print(f"[ok] {street} -> {output_path}")

        print(f"[ok] date: {report_date.short}")
        print(f"[ok] output_dir: {output_dir}")
        print(f"[ok] streets: {len(reports)}, written: {written}")
        return 0
    except Exception as exc:
        print(f"[fail] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
