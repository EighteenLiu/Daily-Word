from __future__ import annotations

import re
import sys
from pathlib import Path


WORKSPACE = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[1]
OUTPUT_ROOT = WORKSPACE / "output"
SUMMARY_ROOT = OUTPUT_ROOT / "\u6bcf\u65e5\u603b\u7ed3"

CN_COMMUNITY_CATEGORY = "\u5c45\u4f4f\u5c0f\u533a\u3001\u5e73\u623f\u80e1\u540c"
CN_PROBLEM_PREFIX = "\u5b58\u5728\u7684\u95ee\u9898\u662f\uff1a"
CN_DISTRICT = "\u533a\u7ea7"
CN_GARBAGE_DAILY = "\u5783\u573e\u5206\u7c7b\u5de5\u4f5c\u65e5\u62a5"
CN_DAILY_SUMMARY = "\u6bcf\u65e5\u6c47\u603b\u60c5\u51b5"
CN_STREET = "\u8857\u9053"
FONT_FANGSONG_GB2312 = "\u4eff\u5b8b_GB2312"
FONT_SONG = "\u5b8b\u4f53"


def chinese_numeral(number: int) -> str:
    import split_street_daily_reports as split

    return split.chinese_numeral(number)


def short_street_name(street: str) -> str:
    return street[:-2] if street.endswith(CN_STREET) else street


def clean_problem_number(text: str) -> str:
    text = text.strip()
    text = re.sub(r"^[\uff08(]\d+[\uff09)]\s*", "", text)
    return text.strip()


def split_problem_segments(text: str) -> list[str]:
    parts = re.split(r"(?=[\uff08(]\d+[\uff09)])", text.strip())
    return [part.strip().strip("\uff1b;") for part in parts if part.strip().strip("\uff1b;")]


def is_real_problem(text: str) -> bool:
    cleaned = clean_problem_number(text)
    if not cleaned:
        return False
    no_problem_prefixes = (
        "\u65e0\u95ee\u9898",
        "\u672a\u53d1\u73b0\u95ee\u9898",
        "\u672a\u89c1\u95ee\u9898",
        "\u65e0\u660e\u663e\u95ee\u9898",
    )
    return not cleaned.startswith(no_problem_prefixes)


def place_problem_texts(place) -> list[str]:
    texts: list[str] = []
    for block in place.blocks:
        if block.text and block.text.strip():
            texts.append(block.text.strip())
    return texts


def merge_problem_texts(texts: list[str]) -> str:
    return "\uff1b".join(text.strip().strip("\uff1b;") for text in texts if text.strip())


def count_real_problems(texts: list[str]) -> int:
    count = 0
    for text in texts:
        for segment in split_problem_segments(text):
            if is_real_problem(segment):
                count += 1
    return count


def body_elements(document):
    body = document._body._element
    return list(body)


def remove_elements_between(document, start_index: int, end_index: int) -> None:
    body = document._body._element
    for element in body_elements(document)[start_index:end_index]:
        body.remove(element)


def paragraph_body_index(document, paragraph) -> int:
    for index, element in enumerate(body_elements(document)):
        if element is paragraph._p:
            return index
    raise ValueError("Paragraph is not in document body.")


def set_run_font(run, font_name: str, font_size_pt: int, bold: bool = False) -> None:
    from docx.oxml.ns import qn
    from docx.shared import Pt

    run.font.name = font_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    run.font.size = Pt(font_size_pt)
    run.bold = bold


def add_formatted_paragraph_at(
    document,
    index: int,
    runs: list[tuple[str, bool]],
    font_name: str,
    font_size_pt: int,
    style=None,
):
    paragraph = document.add_paragraph()
    if style is not None:
        paragraph.style = style
    for text, bold in runs:
        run = paragraph.add_run(text)
        set_run_font(run, font_name, font_size_pt, bold=bold)

    body = document._body._element
    body.remove(paragraph._p)
    body.insert(index, paragraph._p)
    return paragraph


def add_text_paragraph_at(document, index: int, text: str, font_name: str, font_size_pt: int, style=None):
    return add_formatted_paragraph_at(
        document=document,
        index=index,
        runs=[(text, False)],
        font_name=font_name,
        font_size_pt=font_size_pt,
        style=style,
    )


def add_street_heading_at(document, index: int, text: str):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    style = document.styles["Heading 2"] if "Heading 2" in [item.name for item in document.styles] else document.styles["Normal"]
    paragraph = add_text_paragraph_at(document, index, text, FONT_FANGSONG_GB2312, 16, style)
    p_pr = paragraph._p.get_or_add_pPr()
    outline = p_pr.find(qn("w:outlineLvl"))
    if outline is None:
        outline = OxmlElement("w:outlineLvl")
        p_pr.append(outline)
    outline.set(qn("w:val"), "1")
    for run in paragraph.runs:
        set_run_font(run, FONT_FANGSONG_GB2312, 16)
    return paragraph


def add_place_problem_at(document, index: int, place_name: str, problem_text: str):
    return add_formatted_paragraph_at(
        document=document,
        index=index,
        runs=[(place_name, True), (f"{CN_PROBLEM_PREFIX}{problem_text}", False)],
        font_name=FONT_FANGSONG_GB2312,
        font_size_pt=16,
        style=document.styles["Normal"],
    )


def find_paragraph_index(document, predicate) -> int | None:
    for index, paragraph in enumerate(document.paragraphs):
        if predicate(paragraph.text.strip()):
            return index
    return None


def body_index_for_paragraph_number(document, paragraph_number: int) -> int:
    return paragraph_body_index(document, document.paragraphs[paragraph_number])


def default_summary_output_dir(report_date) -> Path:
    return SUMMARY_ROOT / f"\u6bcf\u65e5\u603b\u7ed3\uff08{report_date.short}\uff09"


def chinese_day_prefix(report_date) -> str:
    return f"{report_date.value.month}\u6708{report_date.value.day}\u65e5"


def write_garbage_summary(template_path: Path, reports: dict, report_date, output_dir: Path) -> Path:
    from docx import Document
    import split_street_daily_reports as split

    template_path = split.convert_template_to_docx(template_path.resolve())
    document = Document(template_path)
    split.update_template_date(document, report_date)

    start_para = find_paragraph_index(
        document,
        lambda text: text.startswith("\u4e8c\u3001\u533a\u7ea7\u68c0\u67e5\u5c45\u4f4f\u5c0f\u533a/\u5e73\u623f\u80e1\u540c\u60c5\u51b5"),
    )
    if start_para is None:
        raise RuntimeError("Garbage summary template is missing section: 二、区级检查居住小区/平房胡同情况")

    end_para = find_paragraph_index(
        document,
        lambda text: text.startswith("\u4e09\u3001") and "\u533a\u7ea7\u68c0\u67e5\u5c45\u4f4f" not in text,
    )
    if end_para is None or end_para <= start_para:
        end_body_index = len(body_elements(document)) - 1
    else:
        end_body_index = body_index_for_paragraph_number(document, end_para)

    insert_index = body_index_for_paragraph_number(document, start_para) + 1
    remove_elements_between(document, insert_index, end_body_index)

    inserted = 0
    intro = (
        f"{report_date.chinese}\uff0c\u533a\u5783\u573e\u5206\u7c7b\u63a8\u8fdb\u5de5\u4f5c\u6307\u6325\u90e8\u529e\u516c\u5ba4"
        f"\u5bf9\u897f\u57ce\u533a{len(reports)}\u4e2a\u8857\u9053\u751f\u6d3b\u5783\u573e\u5206\u7c7b\u65e5\u5e38\u8fd0\u884c\u60c5\u51b5"
        f"\u8fdb\u884c\u91cd\u70b9\u62bd\u67e5\uff0c\u68c0\u67e5\u5b58\u5728\u7684\u95ee\u9898\u5982\u4e0b\uff1a"
        f"\uff08\u8be6\u89c1\u300a\u897f\u57ce\u533a\u751f\u6d3b\u5783\u573e\u5206\u7c7b\u65e5\u5e38\u8fd0\u884c\u68c0\u67e5\u62a5\u544a-{report_date.chinese}\u300b\uff0c"
        f"\u8bf7\u76f8\u5173\u8857\u9053\u8ba4\u771f\u5206\u6790\uff0c\u5e76\u4e8e3\u65e5\u5185\u5b8c\u6210\u6574\u6539\uff09\u3002"
    )
    add_text_paragraph_at(document, insert_index + inserted, intro, FONT_FANGSONG_GB2312, 16, document.styles["Normal"])
    inserted += 1
    for street_index, (street, categories) in enumerate(reports.items(), start=1):
        places = categories.get(CN_COMMUNITY_CATEGORY, {})
        if not places:
            continue
        add_street_heading_at(document, insert_index + inserted, f"\uff08{chinese_numeral(street_index)}\uff09{street}")
        inserted += 1
        for place in places.values():
            problem_text = merge_problem_texts(place_problem_texts(place))
            if problem_text:
                add_place_problem_at(document, insert_index + inserted, place.name, problem_text)
            else:
                add_place_problem_at(document, insert_index + inserted, place.name, "\u65e0\u95ee\u9898")
            inserted += 1

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{chinese_day_prefix(report_date)}{CN_GARBAGE_DAILY}.docx"
    document.save(output_path)
    return output_path


def write_daily_summary(template_path: Path, reports: dict, report_date, output_dir: Path) -> Path:
    from docx import Document
    import split_street_daily_reports as split

    template_path = split.convert_template_to_docx(template_path.resolve())
    document = Document(template_path)
    split.update_template_date(document, report_date)

    marker_para = find_paragraph_index(
        document,
        lambda text: text.startswith("\u4eca\u65e5\u5e02\u7ea7\u68c0\u67e5\u60c5\u51b5\u66f4\u65b0\uff1a"),
    )
    if marker_para is None:
        raise RuntimeError("Daily summary template is missing section: 今日市级检查情况更新：")

    marker_body_index = body_index_for_paragraph_number(document, marker_para)
    remove_elements_between(document, 0, marker_body_index)

    lines = [f"{report_date.chinese}\uff0c\u533a\u5783\u573e\u5206\u7c7b\u6307\u6325\u90e8\u68c0\u67e5{len(reports)}\u4e2a\u8857\u9053\u529e\u4e8b\u5904\uff0c\u5404\u8857\u9053\u95ee\u9898\u6570\u5206\u522b\u4e3a"]
    street_items = list(reports.items())
    for index, (street, categories) in enumerate(street_items):
        places = categories.get(CN_COMMUNITY_CATEGORY, {})
        problem_count = sum(count_real_problems(place_problem_texts(place)) for place in places.values())
        suffix = "\u3002" if index == len(street_items) - 1 else "\uff0c"
        lines.append(f"{short_street_name(street)}{problem_count}\u4e2a{suffix}")

    for offset, line in enumerate(lines):
        add_text_paragraph_at(document, offset, line, FONT_SONG, 12, document.styles["Normal"])

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{CN_DAILY_SUMMARY}({report_date.value.day}).docx"
    document.save(output_path)
    return output_path


def write_summaries(
    reports: dict,
    report_date,
    garbage_template: Path | None = None,
    daily_template: Path | None = None,
    output_dir: Path | None = None,
) -> list[Path]:
    output_dir = output_dir or default_summary_output_dir(report_date)
    written: list[Path] = []
    if garbage_template:
        written.append(write_garbage_summary(garbage_template, reports, report_date, output_dir))
    if daily_template:
        written.append(write_daily_summary(daily_template, reports, report_date, output_dir))
    return written
