from __future__ import annotations

import re
import uuid
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from lxml import etree


ISSUE_PATTERN = re.compile(
    r"^\s*(?:\d+\s*[\.．、]\s*)?(?P<street>[^：:，,。\s]+街道)\s*[：:]\s*(?P<text>.+?)\s*$"
)

WORD_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


@dataclass
class OutsideBucketIssue:
    street_name: str
    clean_text: str
    image_paths: list[Path] = field(default_factory=list)

    @property
    def images(self) -> list[Path]:
        return self.image_paths

    @property
    def image_rows(self) -> list[list[Path]]:
        return chunk_list(self.image_paths, 2)


def clean_issue_text(text: str) -> str:
    text = (text or "").strip()
    text = re.sub(r"^\s*\d+\s*[\.．、]?\s*", "", text)
    return re.sub(r"\s+", "", text).strip()


def split_issue_line(line: str) -> OutsideBucketIssue | None:
    line = clean_issue_text(line)
    match = ISSUE_PATTERN.match(line)
    if not match:
        return None
    clean_text = match.group("text").strip()
    if clean_text and not clean_text.endswith(("。", "；", ";", "！", "!", "?", "？")):
        clean_text += "。"
    return OutsideBucketIssue(
        street_name=match.group("street").strip(),
        clean_text=clean_text,
    )


def chunk_list(values: list, size: int = 2) -> list[list]:
    return [values[index:index + size] for index in range(0, len(values), size)]


def parse_outside_bucket_docx(docx_path: Path, image_output_dir: Path) -> list[OutsideBucketIssue]:
    docx_path = Path(docx_path)
    image_output_dir = Path(image_output_dir)
    image_map = extract_images_from_docx(docx_path, image_output_dir)

    issues: list[OutsideBucketIssue] = []
    current: OutsideBucketIssue | None = None
    for text, rids in iter_docx_paragraph_blocks(docx_path):
        parsed = split_issue_line(text) if text else None
        if parsed:
            current = parsed
            issues.append(current)
        if current is None:
            continue
        for rid in rids:
            image_path = image_map.get(rid)
            if image_path:
                current.image_paths.append(image_path)
    return issues


def ensure_outside_bucket_docx(source: Path, output_dir: Path) -> Path:
    source = Path(source)
    if source.suffix.lower() == ".docx":
        return source
    if source.suffix.lower() != ".doc":
        raise ValueError("桶外摆文件仅支持 .docx 或 .doc，请上传 Word 文档。")

    try:
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError("转换桶外摆 .doc 文件需要 pywin32，请先另存为 .docx 后再上传。") from exc

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    destination = output_dir / f"{source.stem}.docx"
    if destination.exists() and destination.stat().st_mtime >= source.stat().st_mtime:
        return destination

    word = None
    document = None
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(str(source.resolve()), ReadOnly=True, AddToRecentFiles=False)
        document.SaveAs2(str(destination.resolve()), FileFormat=16)
    finally:
        try:
            if document is not None:
                document.Close(False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
    return destination


def filter_outside_bucket_items(
    items: Iterable[OutsideBucketIssue],
    target_street_names: Iterable[str],
) -> list[OutsideBucketIssue]:
    targets = {name.strip() for name in target_street_names if name and name.strip()}
    if not targets:
        return []
    return [item for item in items if item.street_name in targets]


def extract_images_from_docx(docx_path: Path, output_dir: Path) -> dict[str, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    rel_map: dict[str, Path] = {}
    with zipfile.ZipFile(docx_path) as archive:
        rels = read_document_relationships(archive)
        names = set(archive.namelist())
        for rel_id, target in rels.items():
            member = normalize_word_member(target)
            if not member or member not in names:
                continue
            extension = Path(member).suffix or ".png"
            output_path = output_dir / f"{uuid.uuid4().hex}{extension}"
            output_path.write_bytes(archive.read(member))
            rel_map[rel_id] = output_path
    return rel_map


def read_document_relationships(archive: zipfile.ZipFile) -> dict[str, str]:
    try:
        rels_xml = archive.read("word/_rels/document.xml.rels")
    except KeyError:
        return {}
    root = etree.fromstring(rels_xml)
    rels: dict[str, str] = {}
    for rel in root.xpath(".//rel:Relationship", namespaces=WORD_NS):
        rel_type = rel.get("Type") or ""
        target = rel.get("Target") or ""
        rel_id = rel.get("Id") or ""
        if rel_id and "image" in rel_type:
            rels[rel_id] = target
    return rels


def normalize_word_member(target: str) -> str:
    if not target or "://" in target:
        return ""
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return target.lstrip("/")
    while target.startswith("../"):
        target = target[3:]
    return f"word/{target}" if not target.startswith("word/") else target


def iter_docx_paragraph_blocks(docx_path: Path) -> Iterable[tuple[str, list[str]]]:
    with zipfile.ZipFile(docx_path) as archive:
        document_xml = archive.read("word/document.xml")
    root = etree.fromstring(document_xml)
    body = root.find("w:body", namespaces=WORD_NS)
    if body is None:
        return
    for child in body:
        yield from iter_block_paragraphs(child)


def iter_block_paragraphs(element) -> Iterable[tuple[str, list[str]]]:
    local_name = etree.QName(element).localname
    if local_name == "p":
        yield paragraph_text_and_images(element)
    elif local_name == "tbl":
        for paragraph in element.xpath(".//w:p", namespaces=WORD_NS):
            yield paragraph_text_and_images(paragraph)


def paragraph_text_and_images(paragraph) -> tuple[str, list[str]]:
    texts = paragraph.xpath(".//w:t/text()", namespaces=WORD_NS)
    rids = []
    for blip in paragraph.xpath(".//a:blip", namespaces=WORD_NS):
        rid = blip.get(f"{{{WORD_NS['r']}}}embed") or blip.get(f"{{{WORD_NS['r']}}}link")
        if rid:
            rids.append(rid)
    return "".join(texts).strip(), rids
