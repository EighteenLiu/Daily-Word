from __future__ import annotations

import re
from dataclasses import fields, is_dataclass
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


IMAGE_WIDTH_CM = 10.2
IMAGE_HEIGHT_CM = 5.74


def render_street_report_docx_from_docxtpl(
    report: Any,
    template_path: Path,
    output_path: Path,
    report_title: str = "",
    report_date_text: str = "",
    image_width_cm: float = IMAGE_WIDTH_CM,
    image_height_cm: float = IMAGE_HEIGHT_CM,
) -> Path:
    """
    使用 docxtpl 直接渲染 Word 模板。
    这个函数会保留模板的页眉、页脚、表格、样式、段落格式等。
    """
    from docxtpl import DocxTemplate, InlineImage
    from docx.shared import Cm

    with TemporaryDirectory(prefix="daily_report_docxtpl_") as temp_dir:
        prepared_template = _prepare_docxtpl_template(
            template_path,
            Path(temp_dir) / "template.docx",
        )
        doc = DocxTemplate(str(prepared_template))
        image_cache: dict[Path, Path] = {}

        def convert(obj):
            if isinstance(obj, Path):
                image_path = _docxtpl_image_path(obj, Path(temp_dir), image_cache)
                if image_path:
                    return InlineImage(doc, str(image_path), width=Cm(image_width_cm), height=Cm(image_height_cm))
                return ""

            if isinstance(obj, list):
                return [convert(item) for item in obj]

            if isinstance(obj, tuple):
                return [convert(item) for item in obj]

            if isinstance(obj, dict):
                return {key: convert(value) for key, value in obj.items()}

            if is_dataclass(obj):
                return {
                    field.name: convert(getattr(obj, field.name))
                    for field in fields(obj)
                }

            return obj

        context = convert(report)
        context["report_title"] = report_title or "区级检查日报"
        context["report_date_text"] = report_date_text

        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.render(context)
        doc.save(str(output_path))
    return output_path


def _docxtpl_image_path(image_path: Path, temp_dir: Path, image_cache: dict[Path, Path]) -> Path | None:
    if image_path in image_cache:
        return image_cache[image_path]
    if not image_path.exists():
        return None

    from PIL import Image

    output = temp_dir / f"image_{len(image_cache) + 1}.png"
    try:
        with Image.open(image_path) as image:
            image.convert("RGB").save(output, "PNG")
    except Exception:
        return None
    image_cache[image_path] = output
    return output


def _prepare_docxtpl_template(template_path: Path, output_path: Path) -> Path:
    document = Document(str(template_path))
    _split_soft_break_paragraphs(document)
    _remove_template_instruction_tail(document)
    _drop_unmatched_jinja_control_paragraphs(document)
    document.save(str(output_path))
    return output_path


def _split_soft_break_paragraphs(document: Document) -> None:
    for paragraph in list(document.paragraphs):
        lines = paragraph.text.splitlines()
        if len(lines) <= 1:
            continue

        parent = paragraph._element.getparent()
        if parent is None:
            continue

        insert_at = parent.index(paragraph._element)
        for line in lines:
            new_paragraph = _clone_paragraph_with_text(paragraph, line)
            insert_at += 1
            parent.insert(insert_at, new_paragraph._element)
        parent.remove(paragraph._element)


def _clone_paragraph_with_text(paragraph, text: str):
    from copy import deepcopy

    new_paragraph = deepcopy(paragraph)
    for run in new_paragraph.runs:
        run.text = ""
    if new_paragraph.runs:
        new_paragraph.runs[0].text = text
    else:
        new_paragraph.add_run(text)
    return new_paragraph


def _remove_template_instruction_tail(document: Document) -> None:
    marker = "模板上下文字段说明"
    paragraphs = list(document.paragraphs)
    start = next((index for index, para in enumerate(paragraphs) if marker in para.text), None)
    if start is None:
        return

    if start > 0 and _paragraph_has_page_break_only(paragraphs[start - 1]):
        start -= 1

    for paragraph in paragraphs[start:]:
        _remove_paragraph(paragraph)


def _drop_unmatched_jinja_control_paragraphs(document: Document) -> None:
    stack: list[str] = []
    for paragraph in list(document.paragraphs):
        tag = _jinja_control_tag(paragraph.text)
        if not tag:
            continue

        if tag in {"if", "for"}:
            stack.append(tag)
            continue
        if tag == "endif":
            if stack and stack[-1] == "if":
                stack.pop()
            else:
                _remove_paragraph(paragraph)
            continue
        if tag == "endfor":
            if stack and stack[-1] == "for":
                stack.pop()
            else:
                _remove_paragraph(paragraph)



def _paragraph_has_page_break_only(paragraph) -> bool:
    return not paragraph.text.strip() and bool(paragraph._p.xpath('.//*[local-name()="br"]'))


def _remove_paragraph(paragraph) -> None:
    parent = paragraph._element.getparent()
    if parent is not None:
        parent.remove(paragraph._element)


try:
    from scripts.daily_report_builder import StreetReport, UnitSection
except ModuleNotFoundError:
    from daily_report_builder import StreetReport, UnitSection




def render_street_report_docx_from_jinja(
    report: "StreetReport",
    template_path: Path,
    output_path: Path,
    report_title: str = "",
    report_date_text: str = "",
    max_template_paras: int | None = 73,
) -> Path:
    """Render a StreetReport using a Jinja2-based .docx template."""
    import jinja2
    
    from docx import Document as DocxDocument
    from docx.shared import Cm as DX_Cm, Pt as DX_Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH as DX_ALIGN
    from docx.oxml.ns import qn as dx_qn

    # Step 1: Extract Jinja2 template text from the docx
    document = DocxDocument(str(template_path))
    template_lines = [p.text for p in document.paragraphs]
    if max_template_paras:
        template_lines = template_lines[:max_template_paras]
    template_lines = _drop_unmatched_jinja_end_tags(template_lines)
    template_text = "\n".join(template_lines)

    # Step 2: Replace all Path objects with unique marker strings
    marker_to_image: dict[str, Path] = {}
    marker_counter: list[int] = [0]

    def _replace_paths(obj: Any) -> Any:
        if isinstance(obj, Path):
            marker = f"__IMG_{marker_counter[0]}__"
            marker_to_image[marker] = obj
            marker_counter[0] += 1
            return marker
        if isinstance(obj, list):
            return [_replace_paths(item) for item in obj]
        if isinstance(obj, dict):
            return {k: _replace_paths(v) for k, v in obj.items()}
        if hasattr(obj, "__dataclass_fields__"):
            return {f: _replace_paths(getattr(obj, f)) for f in obj.__dataclass_fields__}
        return obj

    context = _replace_paths(report)
    context["report_title"] = report_title
    context["report_date_text"] = report_date_text

    # Step 3: Render with Jinja2
    env = jinja2.Environment()
    rendered = env.from_string(template_text).render(**context)

    # Step 4: Build output docx from rendered text
    clean_doc = DocxDocument()
    _set_document_styles(clean_doc)

    for line in rendered.split("\n"):
        text = line.strip()
        if not text:
            continue

        if text.startswith("__IMG_") and text.endswith("__"):
            img_path = marker_to_image.get(text)
            if img_path and img_path.exists():
                para = clean_doc.add_paragraph()
                para.alignment = DX_ALIGN.CENTER
                para.paragraph_format.space_before = DX_Pt(0)
                para.paragraph_format.space_after = DX_Pt(0)
                run = para.add_run()
                try:
                    from PIL import Image as PILImg
                    from io import BytesIO
                    with PILImg.open(img_path) as pil_img:
                        buf = BytesIO()
                        pil_img.convert('RGB').save(buf, 'PNG')
                        buf.seek(0)
                        run.add_picture(buf, width=DX_Cm(IMAGE_WIDTH_CM), height=DX_Cm(IMAGE_HEIGHT_CM))
                except Exception:
                    fb = para.add_run("[图片无法插入：" + img_path.name + "]")
                    fb.font.size = DX_Pt(9)
            else:
                fb = clean_doc.add_paragraph().add_run("[图片无法插入图像]")
            continue

        para = clean_doc.add_paragraph()
        para.paragraph_format.space_before = DX_Pt(0)
        para.paragraph_format.space_after = DX_Pt(0)
        run = para.add_run(text)
        run.font.name = FONT_FANGSONG
        run._element.rPr.rFonts.set(dx_qn("w:eastAsia"), FONT_FANGSONG)
        run.font.size = DX_Pt(14)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    clean_doc.save(str(output_path))
    return output_path


def _drop_unmatched_jinja_end_tags(lines: list[str]) -> list[str]:
    stack: list[str] = []
    filtered: list[str] = []
    for line in lines:
        tag = _jinja_control_tag(line)
        if not tag:
            filtered.append(line)
            continue

        if tag in {"if", "for"}:
            stack.append(tag)
            filtered.append(line)
            continue
        if tag == "endif":
            if stack and stack[-1] == "if":
                stack.pop()
                filtered.append(line)
            continue
        if tag == "endfor":
            if stack and stack[-1] == "for":
                stack.pop()
                filtered.append(line)
            continue
        filtered.append(line)
    return filtered


def _jinja_control_tag(text: str) -> str | None:
    stripped = text.strip()
    tag_match = re.fullmatch(r"\{%-?\s*(?:(?:p|tr|tc|r)\s+)?(\w+)\b.*?-?%\}", stripped)
    return tag_match.group(1) if tag_match else None


FONT_SONG = "\u5b8b\u4f53"
FONT_FANGSONG = "\u4eff\u5b8b"
FONT_HEITI = "\u9ed1\u4f53"


def render_street_report_docx(report: StreetReport, output_path: Path) -> Path:
    document = Document()
    _set_document_styles(document)

    with TemporaryDirectory(prefix="daily_report_images_") as temp_dir:
        image_cache = ImageCache(Path(temp_dir))
        if report.communities:
            _add_heading(document, "一、居住小区", level=1)
            section_index = 1
            for community in report.communities:
                _add_heading(document, f"（{community.index_cn}）{community.name}", level=2)
                _add_heading(document, "1.小区整体情况", level=3)
                _add_body(document, f"{community.overall_intro}存在的问题是：{community.overall_problem_summary}")
                _add_body(document, "小区宣传氛围：")
                _add_images(document, community.promo_images, image_cache)
                _add_body(document, "小区公示牌：")
                _add_images(document, community.notice_board_images, image_cache)

                if community.is_pure_box_room:
                    _add_body(document, "装修垃圾投放点设置：")
                    _add_body(document, "预约收集，集中密闭运输。")
                    _add_body(document, "大件垃圾投放点设置：")
                    _add_body(document, "预约收集，集中密闭运输。")

                _add_heading(document, "2.桶站设置情况", level=3)
                if community.stations:
                    for station in community.stations:
                        _add_heading(document, f"{station.title}：{station.problem_summary}", level=3)
                        _add_images(document, station.images, image_cache)
                else:
                    _add_heading(document, "1号桶站设置情况：无问题", level=3)

                if community.resident_delivery:
                    _add_heading(document, "3.居民投放情况", level=3)
                    _add_body(document, community.resident_delivery.summary)
                    _add_images(document, community.resident_delivery.error_images, image_cache)

        section_number = 1
        if report.communities:
            section_number += 1

        if report.restaurants:
            _render_units(document, f"{_chinese_section_number(section_number)}、餐饮单位", report.restaurants, image_cache)
            section_number += 1

        if report.social_units:
            _render_units(document, f"{_chinese_section_number(section_number)}、社会单位", report.social_units, image_cache)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        document.save(output_path)
    return output_path


class ImageCache:
    def __init__(self, temp_dir: Path) -> None:
        self.temp_dir = temp_dir
        self.converted: dict[Path, Path] = {}

    def word_image_path(self, image_path: Path) -> Path:
        if image_path in self.converted:
            return self.converted[image_path]
        from PIL import Image

        output = self.temp_dir / f"{len(self.converted) + 1}.png"
        with Image.open(image_path) as image:
            image.convert("RGB").save(output, "PNG")
        self.converted[image_path] = output
        return output


def _render_units(document: Document, heading: str, units: list[UnitSection], image_cache: ImageCache) -> None:
    _add_heading(document, heading, level=1)
    for unit in units:
        _add_heading(document, f"（{unit.index_cn}）{unit.name}", level=2)
        _add_heading(document, "1.整体情况", level=3)
        _add_body(document, f"存在的问题是：{unit.overall_problem_summary}")
        suffix = f"{unit.promo_text}" if unit.promo_text and unit.promo_text != "无问题" else ""
        _add_body(document, f"（1）宣传氛围：{suffix}")
        _add_images(document, unit.promo_images, image_cache)
        _add_heading(document, f"2.桶站设置情况：{unit.container_problem_summary}", level=3)
        _add_images(document, unit.container_images, image_cache)


def _set_document_styles(document: Document) -> None:
    normal = document.styles["Normal"]
    normal.font.name = FONT_SONG
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_SONG)
    normal.font.size = Pt(10.5)

    for section in document.sections:
        section.top_margin = Pt(72)
        section.bottom_margin = Pt(72)
        section.left_margin = Pt(72)
        section.right_margin = Pt(72)


def _add_heading(document: Document, text: str, level: int) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(text)
    run.font.name = FONT_HEITI if level == 1 else FONT_FANGSONG
    run._element.rPr.rFonts.set(qn("w:eastAsia"), run.font.name)
    run.font.size = Pt(14)
    run.bold = level in (1, 2)


def _add_body(document: Document, text: str, bold: bool = False) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(text)
    run.font.name = FONT_FANGSONG
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_FANGSONG)
    run.font.size = Pt(14)
    run.bold = bold

def _add_images(document: Document, images: list[Path], image_cache: ImageCache) -> None:
    if not images:
        return

    for image in images:
        paragraph = document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        run = paragraph.add_run()
        try:
            run.add_picture(str(image_cache.word_image_path(image)), width=Cm(IMAGE_WIDTH_CM), height=Cm(IMAGE_HEIGHT_CM))
        except Exception:
            run2 = paragraph.add_run(f"[图片无法插入：{image.name}]")
            run2.font.size = Pt(9)

def _chinese_section_number(number: int) -> str:
    values = {1: "一", 2: "二", 3: "三", 4: "四", 5: "五"}
    return values.get(number, str(number))



