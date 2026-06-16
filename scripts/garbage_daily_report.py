from __future__ import annotations

import re
import posixpath
import zipfile
from collections import OrderedDict, defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Iterable, Any
import xml.etree.ElementTree as ET


COMMUNITY_CATEGORY = "居住小区、平房胡同"
RESTAURANT_CATEGORY = "餐饮单位"
SOCIAL_UNIT_CATEGORY = "社会单位"
SPECIAL_CHECK_CATEGORY = "专项检查"
SPECIAL_CHECK_NO_PROBLEM_TYPES = {"良好，未发现问题", "良好,未发现问题", "无问题", "未发现问题"}

CANONICAL_STREET_ORDER = [
    "德胜街道",
    "什刹海街道",
    "西长安街街道",
    "大栅栏街道",
    "天桥街道",
    "新街口街道",
    "金融街街道",
    "椿树街道",
    "陶然亭街道",
    "展览路街道",
    "月坛街道",
    "广内街道",
    "牛街街道",
    "白纸坊街道",
    "广外街道",
]

JINJA_TOKENS = ("{{", "}}", "{%", "%}")
NON_ISSUE_FRAGMENTS = (
    "设施齐全",
    "除臭排风灭蝇均有设置",
    "均有设置",
    "有宣传氛围",
    "居民知晓垃圾分类",
    "没有智能可回收垃圾箱",
    "无智能可回收垃圾箱",
    "该小区是纯箱房小区",
    "该小区是纯厢房小区",
    "不是箱房小区",
    "正常运行",
    "良好，未发现问题",
    "良好,未发现问题",
    "已核实",
    "投放正确",
)
NON_ISSUE_PATTERNS = (
    r"^\s*(无|无问题|没问题|正常|合格|未发现问题|未见问题|无明显问题)[。；;，,\s]*$",
    r"本小区(?:今天)?(?:一共)?(?:检查)?容器(?:检查)?数量\d+个",
    r"本小区(?:今天)?(?:一共)?检查了?\d+个容器",
    r"小区(?:今天)?(?:一共)?(?:检查了?|有|共)?\d+个(?:垃圾桶|容器)",
    r"容器检查[一二三四五六七八九十百\d]+个",
    r"访问\d+人合格\d+人",
)
COUNT_CN = {
    "一": 1,
    "二": 2,
    "两": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
    "七": 7,
    "八": 8,
    "九": 9,
    "十": 10,
}


@dataclass(frozen=True)
class NormalizedIssue:
    label: str
    count: int = 1

    @property
    def text(self) -> str:
        return f"{self.label}{self.count}处"


class AttachmentManager:
    def __init__(self) -> None:
        self.next_no = 1

    def assign(self, obj: dict[str, Any], enabled: bool) -> None:
        if enabled:
            obj["attachment_no"] = self.next_no
            self.next_no += 1
        else:
            obj["attachment_no"] = None


def normalize_punctuation(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("，", ",").replace("；", ";").replace("：", ":")
    return re.sub(r"\s+", "", text).strip()


def is_non_issue(text: object) -> bool:
    normalized = normalize_punctuation(text)
    if not normalized:
        return True
    if normalized.isdigit():
        return True
    if any(re.search(pattern, normalized) for pattern in NON_ISSUE_PATTERNS):
        residue = normalized
        for pattern in NON_ISSUE_PATTERNS:
            residue = re.sub(pattern, "", residue)
        for fragment in NON_ISSUE_FRAGMENTS:
            residue = residue.replace(fragment, "")
        residue = residue.strip("。.;,、:（）() ")
        return residue in {"", "无", "无问题", "没问题", "未发现问题", "未见问题", "合格", "正常"}
    return any(fragment in normalized for fragment in NON_ISSUE_FRAGMENTS) and not _has_problem_keyword(normalized)


def normalize_issue(row: Any, unit_kind: str = "residential") -> list[NormalizedIssue]:
    raw_problem = _clean_issue_source(getattr(row, "problem", ""))
    if is_non_issue(raw_problem):
        raw_problem = ""
    indicator2 = _clean_issue_source(getattr(row, "indicator2", ""))
    indicator3 = _clean_issue_source(getattr(row, "indicator3", ""))
    text = "".join(part for part in (indicator2, indicator3, raw_problem) if part)
    if not text or is_non_issue(text):
        return []

    mappings = _unit_issue_mappings() if unit_kind in {"social", "catering"} else _residential_issue_mappings()
    issues: list[NormalizedIssue] = []
    seen: set[str] = set()
    for label, patterns in mappings:
        if label in seen:
            continue
        if any(re.search(pattern, text) for pattern in patterns):
            issues.append(NormalizedIssue(label=label, count=extract_count(raw_problem or text)))
            seen.add(label)
    return issues


def extract_count(text: object) -> int:
    normalized = normalize_punctuation(text)
    explicit = [int(match) for match in re.findall(r"(\d+)处", normalized)]
    explicit.extend(COUNT_CN[value] for value in re.findall(r"([一二两三四五六七八九十])处", normalized))
    return max(explicit) if explicit else 1


def format_issue_list(issues: Iterable[dict[str, Any] | NormalizedIssue]) -> str:
    texts = []
    for issue in issues:
        if isinstance(issue, NormalizedIssue):
            texts.append(issue.text)
        else:
            texts.append(str(issue.get("text") or f"{issue['label']}{issue.get('count', 1)}处"))
    if not texts:
        return "无问题。"
    return "；".join(f"（{index}）{text}" for index, text in enumerate(texts, start=1)) + "。"


def build_garbage_daily_context(rows: list[Any], report_date: Any) -> dict[str, Any]:
    context: dict[str, Any] = {
        "report_date": getattr(report_date, "value", report_date),
        "report_date_text": report_date.chinese,
        "street_count": len(_ordered_streets(rows)),
        "city_check": {
            "summary": "今日市级检查情况未更新。",
            "note": "",
            "has_attachment": False,
            "attachment_no": None,
            "table_title": "市级重点抽查情况",
        },
        "city_check_rows": [],
        "residential_streets": _build_category_streets(rows, COMMUNITY_CATEGORY, "residential"),
        "social_streets": _build_category_streets(rows, SOCIAL_UNIT_CATEGORY, "social"),
        "catering_streets": _build_category_streets(rows, RESTAURANT_CATEGORY, "catering"),
        "special_check": build_special_check_context(rows, report_date),
        "enforcement": _default_enforcement(report_date),
        "waste_data": {
            "date_text": _previous_day_text(report_date),
            "summary": "生活垃圾清运量数据暂未更新。",
        },
    }
    manager = AttachmentManager()
    manager.assign(context["city_check"], context["city_check"]["has_attachment"])
    manager.assign(context["enforcement"], True)
    context["enforcement"]["attachment_title"] = (
        f"{_previous_day_short(report_date)}西城区《北京市生活垃圾管理条例》专线执法统计表"
    )
    return context


def build_special_check_context(rows: list[Any], report_date: Any) -> dict[str, Any]:
    report_value = _report_date_value(report_date)
    previous_value = report_value - timedelta(days=1)
    return {
        "summary": (
            f"{_format_cn_date(previous_value)}-{_format_cn_date(report_value)}，"
            "区垃圾分类指挥部针对各街道的桶站满冒脏污问题开展专项检查，各街道的问题如下表所示："
        ),
        "rows": extract_special_check_rows(rows, report_date),
    }


def extract_special_check_rows(rows: Iterable[Any], report_date: Any | None = None) -> list[dict[str, str]]:
    special_rows: list[dict[str, str]] = []
    for row in rows:
        if _clean_special_text(getattr(row, "category", "")) != SPECIAL_CHECK_CATEGORY:
            continue

        problem_type = _clean_special_text(getattr(row, "indicator3", ""))
        issue_text = _clean_special_text(getattr(row, "problem", ""))
        if problem_type in SPECIAL_CHECK_NO_PROBLEM_TYPES or "无问题" in issue_text:
            continue

        raw_time = _clean_time_source(getattr(row, "created_time", "")) or _clean_time_source(getattr(row, "report_time", ""))
        normalized_time = normalize_datetime_text(raw_time)
        special_rows.append(
            {
                "time": normalized_time,
                "street_name": _clean_special_text(getattr(row, "street", "")),
                "point_name": _clean_special_text(getattr(row, "place", "")),
                "problem_type": problem_type or infer_problem_type(issue_text),
            }
        )

    special_rows.sort(key=lambda item: (_datetime_sort_key(item["time"]), item["time"], item["street_name"], item["point_name"]))
    return special_rows


def normalize_datetime_text(value: object) -> str:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d 00:00:00")

    text = _clean_time_source(value)
    if not text:
        return ""
    parsed = _parse_datetime_text(text)
    if parsed is not None:
        return parsed.strftime("%Y-%m-%d %H:%M:%S")
    return text


def infer_problem_type(issue_text: object) -> str:
    text = _clean_special_text(issue_text)
    if not text:
        return ""
    for label, patterns in _residential_issue_mappings():
        if any(re.search(pattern, text) for pattern in patterns):
            return label
    return text


def render_garbage_daily_report(template_path: Path, context: dict[str, Any], output_path: Path) -> Path:
    from docxtpl import DocxTemplate

    with TemporaryDirectory(prefix="garbage_daily_docxtpl_") as temp_dir:
        sanitized = _sanitize_missing_media_relationships(template_path, Path(temp_dir) / "sanitized.docx")
        _validate_garbage_daily_template(sanitized, template_path)
        prepared = _prepare_template(sanitized, Path(temp_dir) / "template.docx", context)
        doc = DocxTemplate(str(prepared))
        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.render(context)
        doc.save(str(output_path))
    validate_rendered_docx(output_path)
    return output_path


def validate_rendered_docx(docx_path: Path) -> None:
    text = extract_all_text_from_docx(docx_path)
    for token in JINJA_TOKENS:
        if token in text:
            raise ValueError(f"模板占位符未清理：{token}")
    forbidden = (
        "详情请见附件2",
        "本小区检查容器数量",
        "本小区容器数量",
        "容器检查",
        "设施齐全",
        "除臭排风灭蝇均有设置",
        "无问题；",
    )
    for value in forbidden:
        if value in text:
            raise ValueError(f"日报包含非标准文本：{value}")
    first = text.find("（一）德胜街道")
    second = text.find("（二）什刹海街道")
    if first != -1 and second != -1 and first > second:
        raise ValueError("居住小区街道顺序错误：德胜街道应在什刹海街道之前")


def extract_all_text_from_docx(docx_path: Path) -> str:
    from docx import Document

    document = Document(docx_path)
    parts: list[str] = []
    parts.extend(paragraph.text for paragraph in document.paragraphs)
    for table in document.tables:
        for row in table.rows:
            parts.extend(cell.text for cell in row.cells)
    for section in document.sections:
        for paragraph in section.header.paragraphs + section.footer.paragraphs:
            parts.append(paragraph.text)
        for table in section.header.tables + section.footer.tables:
            for row in table.rows:
                parts.extend(cell.text for cell in row.cells)
    return "\n".join(parts)


def _validate_garbage_daily_template(template_path: Path, original_path: Path) -> None:
    text = extract_all_text_from_docx(template_path)
    if "{{" in text or "{%" in text:
        return
    if "桶外摆" in text:
        raise RuntimeError(
            f"垃圾分类工作日报模板选择错误：{original_path} 看起来是桶外摆日报，"
            "请在“桶外摆日报”栏选择该文件，在“垃圾分类日报模板”栏选择垃圾分类工作日报模板。"
        )
    raise RuntimeError(f"垃圾分类工作日报模板缺少 Jinja 占位符，请重新选择正确模板：{original_path}")


def _sanitize_missing_media_relationships(template_path: Path, output_path: Path) -> Path:
    relationship_namespace = "http://schemas.openxmlformats.org/package/2006/relationships"
    ET.register_namespace("", relationship_namespace)

    with zipfile.ZipFile(template_path) as archive:
        names = set(archive.namelist())
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as fixed:
            for item in archive.infolist():
                data = archive.read(item.filename)
                if item.filename.endswith(".rels"):
                    data = _drop_missing_media_relationships(
                        item.filename,
                        data,
                        names,
                        relationship_namespace,
                    )
                fixed.writestr(item, data)
    return output_path


def _drop_missing_media_relationships(
    rels_member: str,
    data: bytes,
    package_members: set[str],
    relationship_namespace: str,
) -> bytes:
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return data

    changed = False
    for relationship in list(root):
        rel_type = relationship.attrib.get("Type", "")
        target = relationship.attrib.get("Target", "")
        if "image" not in rel_type or not target or "://" in target:
            continue
        member = _relationship_target_member(rels_member, target)
        if member not in package_members:
            root.remove(relationship)
            changed = True

    if not changed:
        return data
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _relationship_target_member(rels_member: str, target: str) -> str:
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    if "/_rels/" not in rels_member:
        base_dir = posixpath.dirname(rels_member)
    else:
        base_dir, rels_name = rels_member.split("/_rels/", 1)
        source_name = rels_name.removesuffix(".rels")
        base_dir = posixpath.dirname(posixpath.join(base_dir, source_name))
    return posixpath.normpath(posixpath.join(base_dir, target))


def _build_category_streets(rows: list[Any], category: str, unit_kind: str) -> list[dict[str, Any]]:
    by_street: OrderedDict[str, OrderedDict[str, list[Any]]] = OrderedDict()
    for row in rows:
        if getattr(row, "category", "") != category:
            continue
        street = getattr(row, "street", "").strip()
        place = getattr(row, "place", "").strip()
        if not street or not place:
            continue
        by_street.setdefault(street, OrderedDict()).setdefault(place, []).append(row)

    streets: list[dict[str, Any]] = []
    for street in _sort_streets(by_street):
        points = []
        outside_bucket_issues = []
        for place, place_rows in by_street[street].items():
            issue_counts: OrderedDict[str, int] = OrderedDict()
            for row in place_rows:
                issue_rows = normalize_issue(row, unit_kind)
                if unit_kind == "residential" and _is_street_outside_bucket(row, issue_rows):
                    outside_bucket_issues.append({"text": _format_outside_bucket_issue(row)})
                    continue
                for issue in issue_rows:
                    issue_counts[issue.label] = issue_counts.get(issue.label, 0) + issue.count
            points.append(
                {
                    "name": place,
                    "issues": [
                        {"label": label, "count": count, "text": f"{label}{count}处"}
                        for label, count in _sort_issue_counts(issue_counts, unit_kind)
                    ],
                }
            )
        streets.append(
            {
                "index_cn": _chinese_numeral(len(streets) + 1),
                "name": street,
                "points": points,
                "outside_bucket_issues": _dedupe_issue_dicts(outside_bucket_issues),
            }
        )
    return streets


def _residential_issue_mappings() -> list[tuple[str, tuple[str, ...]]]:
    return [
        ("桶站周边不洁", (r"周边.*不洁", r"站外不洁", r"环境不洁")),
        ("桶站满冒", (r"满冒",)),
        ("高峰时段桶站未开盖", (r"(高峰|晚高峰).*未开盖", r"未开盖.*(高峰|晚高峰)")),
        ("居民自主投放不准确", (r"混投", r"投放不规范", r"投放不准确", r"投放错误")),
        ("站外摆桶", (r"站外摆桶", r"桶外摆", r"垃圾桶外摆")),
        ("无宣传氛围", (r"无宣传氛围", r"没有看到.*宣传")),
        ("无小区公示牌", (r"无小区公示牌", r"未见公示牌")),
        ("无桶站", (r"无桶站", r"未见垃圾桶站")),
        ("散桶", (r"散桶",)),
        ("垃圾车混装混运", (r"垃圾车混装混运",)),
    ]


def _unit_issue_mappings() -> list[tuple[str, tuple[str, ...]]]:
    return [
        ("分类投放不准确", (r"分类投放不准确", r"投放不准确", r"投放错误", r"混投", r"厨余.*其他", r"其他.*厨余")),
        ("桶站周边不洁", (r"周边.*不洁", r"环境不洁")),
        ("桶站满冒", (r"满冒",)),
        ("站外摆桶", (r"站外摆桶", r"桶外摆", r"垃圾桶外摆")),
        ("无宣传氛围", (r"无宣传氛围", r"没有看到.*宣传")),
    ]


def _sort_issue_counts(issue_counts: OrderedDict[str, int], unit_kind: str) -> list[tuple[str, int]]:
    mappings = _unit_issue_mappings() if unit_kind in {"social", "catering"} else _residential_issue_mappings()
    order = {label: index for index, (label, _patterns) in enumerate(mappings)}
    return sorted(issue_counts.items(), key=lambda item: (order.get(item[0], len(order)), item[0]))


def _clean_issue_source(value: object) -> str:
    text = normalize_punctuation(value)
    text = re.sub(r"^[（(]?\d+[）)]?", "", text)
    for pattern in NON_ISSUE_PATTERNS:
        text = re.sub(pattern, "", text)
    for fragment in NON_ISSUE_FRAGMENTS:
        text = text.replace(fragment, "")
    text = re.sub(r"\d+个(?:垃圾桶|容器)", "", text)
    text = re.sub(r"([0-9A-Za-z]+|[一二三四五六七八九十]+)(?:号)?桶站", "", text)
    return text.strip("。.;,、:（）() ")


def _has_problem_keyword(text: str) -> bool:
    return any(
        keyword in text
        for keyword in (
            "不洁",
            "满冒",
            "未开盖",
            "混投",
            "不规范",
            "不准确",
            "桶外摆",
            "站外摆桶",
            "无宣传",
            "无小区公示牌",
            "无桶站",
            "散桶",
            "混装混运",
        )
    )


def _is_street_outside_bucket(row: Any, issues: list[NormalizedIssue]) -> bool:
    if not any(issue.label == "站外摆桶" for issue in issues):
        return False
    text = normalize_punctuation(getattr(row, "problem", ""))
    if "桶站" in text or "小区" in text:
        return False
    return any(keyword in text for keyword in ("门口", "道路", "路", "街", "胡同", "大街", "北口", "南口", "东口", "西口", "商户"))


def _format_outside_bucket_issue(row: Any) -> str:
    text = str(getattr(row, "problem", "")).strip("。；; ")
    text = re.sub(r"一处$", "1处", text)
    if not re.search(r"\d+处$", text):
        text += "1处"
    return text


def _dedupe_issue_dicts(items: list[dict[str, str]]) -> list[dict[str, str]]:
    result: list[dict[str, str]] = []
    seen: set[str] = set()
    for item in items:
        text = item["text"]
        if text in seen:
            continue
        seen.add(text)
        result.append(item)
    return result


def _ordered_streets(rows: list[Any]) -> list[str]:
    seen: set[str] = set()
    streets: list[str] = []
    for row in rows:
        street = getattr(row, "street", "").strip()
        if street and street not in seen:
            seen.add(street)
            streets.append(street)
    return _sort_streets({street: None for street in streets})


def _sort_streets(streets: dict[str, Any] | OrderedDict[str, Any]) -> list[str]:
    order = {street: index for index, street in enumerate(CANONICAL_STREET_ORDER)}
    return sorted(streets, key=lambda street: (order.get(street, len(order)), street))


def _default_enforcement(report_date: Any) -> dict[str, Any]:
    return {
        "attachment_no": None,
        "date_text": _previous_day_text(report_date),
        "attachment_title": "",
        "rows": [
            {"street_name": street, "check_count": 0, "case_count": 0, "fine_amount": 0}
            for street in CANONICAL_STREET_ORDER
        ],
        "total_checks": 0,
        "case_count": 0,
        "fine_amount": 0,
    }


def _prepare_template(template_path: Path, output_path: Path, context: dict[str, Any]) -> Path:
    from docx import Document

    document = Document(str(template_path))
    if not context.get("city_check", {}).get("has_attachment"):
        for table in list(document.tables):
            table_text = "\n".join(cell.text for row in table.rows for cell in row.cells)
            if "city_check" in table_text:
                table_element = table._tbl
                table_element.getparent().remove(table_element)
    if not context.get("special_check", {}).get("rows"):
        for table in list(document.tables):
            table_text = "\n".join(cell.text for row in table.rows for cell in row.cells)
            if "special_check.rows" in table_text:
                table_element = table._tbl
                table_element.getparent().remove(table_element)
    for paragraph in document.paragraphs:
        _replace_paragraph_text(paragraph, "附件2", "附件{{ enforcement.attachment_no }}")
    document.save(str(output_path))
    return output_path


def _replace_paragraph_text(paragraph: Any, old: str, new: str) -> None:
    if old not in paragraph.text:
        return
    replaced = paragraph.text.replace(old, new)
    for index, run in enumerate(paragraph.runs):
        run.text = replaced if index == 0 else ""


def _previous_day_text(report_date: Any) -> str:
    return _format_cn_date(_report_date_value(report_date) - timedelta(days=1))


def _previous_day_short(report_date: Any) -> str:
    value = _report_date_value(report_date) - timedelta(days=1)
    return f"{value.month}.{value.day}"


def _report_date_value(report_date: Any) -> date:
    value = getattr(report_date, "value", report_date)
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    raise TypeError(f"Unsupported report date: {report_date!r}")


def _format_cn_date(value: date) -> str:
    return f"{value.year}年{value.month}月{value.day}日"


def _clean_special_text(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return re.sub(r"[ \t\u3000]+", "", text).strip()


def _clean_time_source(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return re.sub(r"[ \t\u3000]+", " ", text).strip()


def _parse_datetime_text(text: str) -> datetime | None:
    normalized = text.strip()
    normalized = normalized.replace("年", "-").replace("月", "-").replace("日", " ")
    normalized = normalized.replace("/", "-").replace(".", "-")
    normalized = re.sub(r"\s+", " ", normalized).strip()
    normalized = re.sub(r"^(\d{4})-(\d{1,2})-(\d{1,2})(\d{1,2}:\d{2})", r"\1-\2-\3 \4", normalized)
    formats = (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S.%f",
    )
    for fmt in formats:
        try:
            return datetime.strptime(normalized, fmt)
        except ValueError:
            continue
    return None


def _datetime_sort_key(text: str) -> datetime:
    parsed = _parse_datetime_text(text)
    return parsed or datetime.max


def _chinese_numeral(number: int) -> str:
    digits = "零一二三四五六七八九"
    if number <= 0:
        return str(number)
    if number < 10:
        return digits[number]
    if number < 20:
        return "十" + (digits[number % 10] if number % 10 else "")
    tens, ones = divmod(number, 10)
    return digits[tens] + "十" + (digits[ones] if ones else "")
