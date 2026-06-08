from __future__ import annotations

import re
import random
import zipfile
import bisect
import hashlib
from io import BytesIO
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable
import xml.etree.ElementTree as ET


COMMUNITY_CATEGORY = "居住小区、平房胡同"
RESTAURANT_CATEGORY = "餐饮单位"
SOCIAL_UNIT_CATEGORY = "社会单位"
FALLBACK_IMAGE_COLUMNS = set(range(8, 26))

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


@dataclass
class LedgerRow:
    row_number: int
    category: str
    street: str
    place: str
    indicator1: str
    indicator2: str
    indicator3: str
    problem: str
    image_paths: list[Path] = field(default_factory=list)


@dataclass
class RowImage:
    row_number: int
    column_number: int
    media_name: str
    suffix: str
    data: bytes


@dataclass
class EmbeddedImage:
    media_name: str
    suffix: str
    data: bytes


@dataclass
class StationSection:
    title: str
    station_no: str
    problem_summary: str
    images: list[Path] = field(default_factory=list)


@dataclass
class ResidentDeliverySection:
    summary: str = ""
    error_images: list[Path] = field(default_factory=list)


@dataclass
class CommunitySection:
    index_cn: str
    name: str
    overall_problem_summary: str
    overall_intro: str = ""
    promo_images: list[Path] = field(default_factory=list)
    notice_board_images: list[Path] = field(default_factory=list)
    is_pure_box_room: bool = False
    stations: list[StationSection] = field(default_factory=list)
    resident_delivery: ResidentDeliverySection | None = None


@dataclass
class UnitSection:
    index_cn: str
    name: str
    overall_problem_summary: str
    has_promo: bool = False
    promo_text: str = ""
    promo_images: list[Path] = field(default_factory=list)
    container_problem_summary: str = "无问题。"
    container_images: list[Path] = field(default_factory=list)


@dataclass
class StreetReport:
    street: str
    communities: list[CommunitySection] = field(default_factory=list)
    restaurants: list[UnitSection] = field(default_factory=list)
    social_units: list[UnitSection] = field(default_factory=list)


def normalize_text(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return re.sub(r"\s+", "", text).strip()


def display_text(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return re.sub(r"[ \t\u3000]+", "", text).strip()


def load_ledger_rows(
    path: Path,
    include_images: bool = False,
    image_source_path: Path | None = None,
) -> list[LedgerRow]:
    if path.suffix.lower() == ".xlsx":
        return _load_xlsx_rows(path, include_images=include_images, image_source_path=image_source_path)
    return _load_excel_com_rows(path, include_images=include_images)


def _load_xlsx_rows(
    path: Path,
    include_images: bool = False,
    image_source_path: Path | None = None,
) -> list[LedgerRow]:
    import openpyxl

    workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    sheet = workbook[workbook.sheetnames[0]]
    headers = [normalize_text(sheet.cell(2, col).value) for col in range(1, sheet.max_column + 1)]
    columns = _header_columns(headers)
    rows: list[LedgerRow] = []
    for row_number in range(3, sheet.max_row + 1):
        rows.append(
            LedgerRow(
                row_number=row_number,
                category=display_text(sheet.cell(row_number, columns["category"]).value),
                street=display_text(sheet.cell(row_number, columns["street"]).value),
                place=display_text(sheet.cell(row_number, columns["place"]).value),
                indicator1=display_text(sheet.cell(row_number, columns["indicator1"]).value),
                indicator2=display_text(sheet.cell(row_number, columns["indicator2"]).value),
                indicator3=display_text(sheet.cell(row_number, columns["indicator3"]).value),
                problem=display_text(sheet.cell(row_number, columns["problem"]).value),
            )
        )
    rows = [row for row in rows if any((row.category, row.street, row.place, row.problem))]
    if include_images:
        _attach_images(path, sheet.title, rows, image_source_path=image_source_path)
    return rows


def _load_excel_com_rows(path: Path, include_images: bool = False) -> list[LedgerRow]:
    import win32com.client

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        workbook = excel.Workbooks.Open(str(path.resolve()), ReadOnly=True)
        sheet = workbook.Worksheets(1)
        used = sheet.UsedRange
        headers = [normalize_text(sheet.Cells(2, col).Text) for col in range(1, used.Columns.Count + 1)]
        columns = _header_columns(headers)
        rows: list[LedgerRow] = []
        for row_number in range(3, used.Rows.Count + 1):
            rows.append(
                LedgerRow(
                    row_number=row_number,
                    category=display_text(sheet.Cells(row_number, columns["category"]).Text),
                    street=display_text(sheet.Cells(row_number, columns["street"]).Text),
                    place=display_text(sheet.Cells(row_number, columns["place"]).Text),
                    indicator1=display_text(sheet.Cells(row_number, columns["indicator1"]).Text),
                    indicator2=display_text(sheet.Cells(row_number, columns["indicator2"]).Text),
                    indicator3=display_text(sheet.Cells(row_number, columns["indicator3"]).Text),
                    problem=display_text(sheet.Cells(row_number, columns["problem"]).Text),
                )
            )
        workbook.Close(False)
        workbook = None
        rows = [row for row in rows if any((row.category, row.street, row.place, row.problem))]
        if include_images:
            image_root = path.parent / "extracted_images" / path.stem
            _attach_row_images(rows, extract_excel_com_row_images(path), image_root)
        return rows
    finally:
        if "workbook" in locals() and workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass


def _attach_images(
    path: Path,
    sheet_name: str,
    rows: list[LedgerRow],
    image_source_path: Path | None = None,
) -> None:
    image_source_path = image_source_path.resolve() if image_source_path else path
    image_root = path.parent / "extracted_images" / image_source_path.stem
    images: list[RowImage] = []
    images = extract_xlsx_row_images(path, sheet_name)
    _attach_row_images(rows, images, image_root)


def _attach_row_images(rows: list[LedgerRow], images: Iterable[RowImage], image_root: Path) -> None:
    image_root.mkdir(parents=True, exist_ok=True)
    row_numbers = [row.row_number for row in rows]
    row_map = {row.row_number: row for row in rows}
    counters: dict[int, int] = defaultdict(int)
    seen: set[str] = set()
    for image in images:
        row_number = _nearest_data_row(image.row_number, row_numbers)
        row = row_map.get(row_number) if row_number is not None else None
        if row is None:
            continue
        digest = hashlib.sha1(image.data).hexdigest()
        unique_key = f"{row.row_number}:{digest}:{image.media_name}"
        if unique_key in seen:
            continue
        seen.add(unique_key)
        counters[row.row_number] += 1
        output = image_root / f"row_{row.row_number}_{counters[row.row_number]}_{digest[:10]}{image.suffix}"
        output.write_bytes(image.data)
        row.image_paths.append(output)


def extract_xls_original_row_images(xls_path: Path, xlsx_path: Path, sheet_name: str) -> list[RowImage]:
    anchored_images = extract_xlsx_row_images(xlsx_path, sheet_name)
    original_images = extract_xls_embedded_images(xls_path)
    if not anchored_images or not original_images:
        return []

    original_hashes = [(_image_hash(image.data), image) for image in original_images]
    if any(hash_value is None for hash_value, _ in original_hashes):
        return _pair_xls_images_by_order(anchored_images, original_images)

    replaced: list[RowImage] = []
    used_original_indexes: set[int] = set()
    for anchored in anchored_images:
        thumb_hash = _image_hash(anchored.data)
        if thumb_hash is None:
            replaced.append(anchored)
            continue
        best_distance = 10**9
        best_image: EmbeddedImage | None = None
        best_index: int | None = None
        for index, (original_hash, original) in enumerate(original_hashes):
            if index in used_original_indexes:
                continue
            if original_hash is None:
                continue
            distance = (thumb_hash ^ original_hash).bit_count()
            if distance < best_distance:
                best_distance = distance
                best_image = original
                best_index = index
        if best_image is None or best_index is None or best_distance > 110:
            replaced.append(anchored)
            continue
        used_original_indexes.add(best_index)
        replaced.append(
            RowImage(
                row_number=anchored.row_number,
                column_number=anchored.column_number,
                media_name=f"{anchored.media_name}|{best_image.media_name}",
                suffix=best_image.suffix,
                data=best_image.data,
            )
        )
    return replaced


def _pair_xls_images_by_order(anchored_images: list[RowImage], original_images: list[EmbeddedImage]) -> list[RowImage]:
    paired: list[RowImage] = []
    for index, anchored in enumerate(anchored_images):
        if index >= len(original_images):
            paired.append(anchored)
            continue
        original = original_images[index]
        paired.append(
            RowImage(
                row_number=anchored.row_number,
                column_number=anchored.column_number,
                media_name=f"{anchored.media_name}|{original.media_name}",
                suffix=original.suffix,
                data=original.data,
            )
        )
    return paired


def extract_xls_embedded_images(path: Path) -> list[EmbeddedImage]:
    try:
        import olefile
    except ImportError:
        return []

    images: list[EmbeddedImage] = []
    try:
        ole = olefile.OleFileIO(str(path))
    except Exception:
        return []
    try:
        for stream in ole.listdir(streams=True, storages=False):
            if stream[-1] not in {"Workbook", "Book"}:
                continue
            data = ole.openstream(stream).read()
            images.extend(_scan_image_blobs(data, "/".join(stream)))
    finally:
        ole.close()
    return images


def _scan_image_blobs(data: bytes, stream_name: str) -> list[EmbeddedImage]:
    images: list[EmbeddedImage] = []
    index = 0
    counters: dict[str, int] = defaultdict(int)
    signatures = ((b"\xff\xd8\xff", ".jpg"), (b"\x89PNG\r\n\x1a\n", ".png"))
    while index < len(data):
        starts = [(pos, suffix) for signature, suffix in signatures if (pos := data.find(signature, index)) != -1]
        if not starts:
            break
        start, suffix = min(starts, key=lambda item: item[0])
        if suffix == ".png":
            end = data.find(b"IEND\xaeB`\x82", start)
            if end == -1:
                index = start + 8
                continue
            end += 8
        else:
            end = data.find(b"\xff\xd9", start + 3)
            if end == -1:
                index = start + 3
                continue
            end += 2
        blob = data[start:end]
        if _valid_problem_photo_blob(blob):
            counters[suffix] += 1
            images.append(
                EmbeddedImage(
                    media_name=f"{stream_name}_image_{counters[suffix]}",
                    suffix=suffix,
                    data=blob,
                )
            )
        index = end
    return images


def _valid_problem_photo_blob(data: bytes) -> bool:
    try:
        from PIL import Image

        with Image.open(BytesIO(data)) as image:
            width, height = image.size
            image.verify()
    except Exception:
        return False
    return width >= 300 or height >= 300


def _image_hash(data: bytes) -> int | None:
    try:
        from io import BytesIO
        from PIL import Image, ImageOps

        with Image.open(BytesIO(data)) as image:
            image = ImageOps.exif_transpose(image).convert("L").resize((16, 16))
            pixels = list(image.getdata())
    except Exception:
        return None
    average = sum(pixels) / len(pixels)
    bits = 0
    for pixel in pixels:
        bits = (bits << 1) | int(pixel >= average)
    return bits


def extract_excel_com_row_images(path: Path, sheet_name: str | None = None) -> list[RowImage]:
    try:
        import win32com.client
    except ImportError:
        return []

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    images: list[RowImage] = []
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(path.resolve()), ReadOnly=True, UpdateLinks=0, AddToMru=False)
        sheet = workbook.Worksheets(sheet_name) if sheet_name else workbook.Worksheets(1)
        image_columns = _excel_image_columns(sheet)
        original_images = extract_xls_embedded_images(path)
        original_index = 0
        shape_count = int(sheet.Shapes.Count)
        for index in range(1, shape_count + 1):
            shape = sheet.Shapes.Item(index)
            try:
                row_number = int(shape.TopLeftCell.Row)
                column_number = int(shape.TopLeftCell.Column)
            except Exception:
                continue
            if column_number not in image_columns:
                continue
            if original_index >= len(original_images):
                break
            original = original_images[original_index]
            original_index += 1
            images.append(
                RowImage(
                    row_number=row_number,
                    column_number=column_number,
                    media_name=f"excel_shape_{index}|{original.media_name}",
                    suffix=original.suffix,
                    data=original.data,
                )
            )
    except Exception:
        return []
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass
    return images


def _excel_image_columns(sheet) -> set[int]:
    try:
        max_columns = int(sheet.UsedRange.Columns.Count)
    except Exception:
        max_columns = 200
    matches: set[int] = set()
    for column in range(1, max_columns + 1):
        try:
            value = normalize_text(sheet.Cells(1, column).Text)
        except Exception:
            continue
        if value == "问题照片":
            matches.add(column)
            try:
                merge_area = sheet.Cells(1, column).MergeArea
                start = int(merge_area.Column)
                count = int(merge_area.Columns.Count)
                matches.update(range(start, start + count))
            except Exception:
                pass
    return matches or FALLBACK_IMAGE_COLUMNS


def extract_xlsx_row_images(path: Path, sheet_name: str) -> list[RowImage]:
    images: list[RowImage] = []
    with zipfile.ZipFile(path) as archive:
        sheets = _workbook_sheet_map(archive)
        sheet_path = sheets.get(sheet_name)
        if not sheet_path:
            return images
        image_columns = _image_columns_for_sheet(archive, sheet_path)
        sheet_rels_path = f"{Path(sheet_path).parent.as_posix()}/_rels/{Path(sheet_path).name}.rels"
        sheet_rels = _read_relationships(archive, sheet_rels_path)
        sheet_root = ET.fromstring(archive.read(sheet_path))
        drawing = sheet_root.find("main:drawing", NS)
        if drawing is None:
            return images
        drawing_rel_id = drawing.attrib.get(f"{{{NS['r']}}}id")
        drawing_target = sheet_rels.get(drawing_rel_id or "")
        if not drawing_target:
            return images
        drawing_path = _rel_target_to_zip_path(Path(sheet_path).parent.as_posix(), drawing_target)
        drawing_rels_path = f"{Path(drawing_path).parent.as_posix()}/_rels/{Path(drawing_path).name}.rels"
        drawing_rels = _read_relationships(archive, drawing_rels_path)
        drawing_root = ET.fromstring(archive.read(drawing_path))

        for anchor_tag in ("xdr:twoCellAnchor", "xdr:oneCellAnchor"):
            for anchor in drawing_root.findall(anchor_tag, NS):
                from_node = anchor.find("xdr:from", NS)
                blip = anchor.find(".//a:blip", NS)
                if from_node is None or blip is None:
                    continue
                row_node = from_node.find("xdr:row", NS)
                col_node = from_node.find("xdr:col", NS)
                rel_id = blip.attrib.get(f"{{{NS['r']}}}embed")
                if row_node is None or col_node is None or not rel_id:
                    continue
                row_number = int(row_node.text or "0") + 1
                column_number = int(col_node.text or "0") + 1
                if column_number not in image_columns:
                    continue
                media_target = drawing_rels.get(rel_id)
                if not media_target:
                    continue
                media_path = _rel_target_to_zip_path(Path(drawing_path).parent.as_posix(), media_target)
                if media_path not in archive.namelist():
                    continue
                images.append(
                    RowImage(
                        row_number=row_number,
                        column_number=column_number,
                        media_name=media_path,
                        suffix=Path(media_path).suffix.lower() or ".bin",
                        data=archive.read(media_path),
                    )
                )
    return sorted(images, key=lambda image: (image.row_number, image.column_number, image.media_name))


def _image_columns_for_sheet(archive: zipfile.ZipFile, sheet_path: str) -> set[int]:
    shared_strings = _shared_strings(archive)
    sheet_root = ET.fromstring(archive.read(sheet_path))
    matches: set[int] = set()
    for cell in sheet_root.findall(".//main:sheetData/main:row[@r='1']/main:c", NS):
        value = normalize_text(_cell_text(cell, shared_strings))
        if value != "问题照片":
            continue
        column = _column_index_from_cell_ref(cell.attrib.get("r", ""))
        if column:
            matches.add(column)

    if not matches:
        return FALLBACK_IMAGE_COLUMNS

    image_columns = set(matches)
    merge_cells = sheet_root.find("main:mergeCells", NS)
    if merge_cells is not None:
        for merge_cell in merge_cells.findall("main:mergeCell", NS):
            ref = merge_cell.attrib.get("ref", "")
            bounds = _merge_bounds(ref)
            if not bounds:
                continue
            min_col, min_row, max_col, max_row = bounds
            if min_row <= 1 <= max_row and any(min_col <= column <= max_col for column in matches):
                image_columns.update(range(min_col, max_col + 1))
    return image_columns


def _shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for si in root.findall("main:si", NS):
        strings.append("".join(text.text or "" for text in si.findall(".//main:t", NS)))
    return strings


def _cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(text.text or "" for text in cell.findall(".//main:t", NS))
    value = cell.find("main:v", NS)
    if value is None or value.text is None:
        return ""
    if cell_type == "s":
        try:
            return shared_strings[int(value.text)]
        except (ValueError, IndexError):
            return ""
    return value.text


def _merge_bounds(ref: str) -> tuple[int, int, int, int] | None:
    if ":" not in ref:
        return None
    start, end = ref.split(":", 1)
    start_col, start_row = _cell_coordinate(start)
    end_col, end_row = _cell_coordinate(end)
    if not all((start_col, start_row, end_col, end_row)):
        return None
    return start_col, start_row, end_col, end_row


def _cell_coordinate(ref: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Z]+)(\d+)", ref.upper())
    if not match:
        return 0, 0
    return _column_index_from_letters(match.group(1)), int(match.group(2))


def _column_index_from_cell_ref(ref: str) -> int:
    match = re.match(r"([A-Z]+)", ref.upper())
    return _column_index_from_letters(match.group(1)) if match else 0


def _column_index_from_letters(letters: str) -> int:
    value = 0
    for char in letters:
        value = value * 26 + ord(char) - ord("A") + 1
    return value


def _nearest_data_row(image_row: int, data_rows: list[int]) -> int | None:
    if not data_rows:
        return None
    index = bisect.bisect_right(data_rows, image_row)
    if index:
        return data_rows[index - 1]
    return data_rows[0]


def _rel_target_to_zip_path(base_dir: str, target: str) -> str:
    target_path = (Path(base_dir) / target).as_posix()
    parts: list[str] = []
    for part in target_path.split("/"):
        if part in ("", "."):
            continue
        if part == "..":
            if parts:
                parts.pop()
        else:
            parts.append(part)
    return "/".join(parts)


def _read_relationships(archive: zipfile.ZipFile, rels_path: str) -> dict[str, str]:
    if rels_path not in archive.namelist():
        return {}
    root = ET.fromstring(archive.read(rels_path))
    relationships: dict[str, str] = {}
    for rel in root.findall("rel:Relationship", NS):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if rel_id and target:
            relationships[rel_id] = target
    return relationships


def _workbook_sheet_map(archive: zipfile.ZipFile) -> dict[str, str]:
    workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
    workbook_rels = _read_relationships(archive, "xl/_rels/workbook.xml.rels")
    sheets: dict[str, str] = {}
    for sheet in workbook_root.findall("main:sheets/main:sheet", NS):
        name = sheet.attrib.get("name")
        rel_id = sheet.attrib.get(f"{{{NS['r']}}}id")
        target = workbook_rels.get(rel_id or "")
        if name and target:
            sheets[name] = _rel_target_to_zip_path("xl", target)
    return sheets


def _header_columns(headers: list[str]) -> dict[str, int]:
    def find(name: str, occurrence: int = 1) -> int:
        seen = 0
        for index, header in enumerate(headers, start=1):
            if header == name:
                seen += 1
                if seen == occurrence:
                    return index
        raise ValueError(f"Missing required header: {name}")

    return {
        "category": find("2级点位"),
        "street": find("3级点位"),
        "place": find("4级点位"),
        "indicator1": find("1级指标"),
        "indicator2": find("2级指标"),
        "indicator3": find("3级指标"),
        "problem": find("具体问题"),
    }


def is_no_problem(text: str) -> bool:
    normalized = normalize_text(text)
    if not normalized:
        return True
    if _is_inspection_count_text(normalized):
        return True
    no_problem_markers = (
        "无问题",
        "良好，未发现问题",
        "良好,未发现问题",
        "未发现问题",
        "有宣传氛围",
        "均有设置",
        "已核实",
        "不是箱房小区",
    )
    if normalized.isdigit():
        return True
    if normalized in {"有", "无", "投放正确"}:
        return True
    return any(marker in normalized for marker in no_problem_markers) and not any(
        marker in normalized for marker in ("不洁", "不准确", "站外摆桶", "不齐全", "无宣传")
    )


def clean_problem_text(text: str) -> str:
    text = display_text(text)
    text = re.sub(r"^无问题[，,。；;、（）()本小区今天一共检查了0-9个容器纯箱房小区]*", "", text)
    text = re.sub(r"（?本小区.*?容器。?）?", "", text)
    text = re.sub(r"本小区(?:今天)?(?:一共)?检查了?\d+个容器", "", text)
    text = re.sub(r"(?:本小区|小区)?(?:有|共)?\d+个(?:垃圾桶|容器)", "", text)
    text = re.sub(r"\d+个(?:垃圾桶|容器)", "", text)
    text = text.replace("。", "").replace("；", "")
    station_match = re.search(r"([0-9A-Za-z]+|[一二三四五六七八九十]+)(?:号)?桶站(.+)$", text)
    if station_match:
        text = station_match.group(2)
    text = text.strip("：:，,。；;、 ")
    if is_no_problem(text):
        return ""
    return text


def _is_inspection_count_text(text: str) -> bool:
    return bool(
        re.fullmatch(r"(?:本小区|小区)?(?:今天)?(?:一共)?(?:检查了?|有|共)?\d+个(?:垃圾桶|容器)", text)
    )


def effective_problem_text(row: LedgerRow) -> str:
    problem = clean_problem_text(row.problem)
    indicator3 = clean_problem_text(row.indicator3)
    if "居民自主投放" in row.indicator2 and (
        "投放错误" in row.problem
        or "不准确" in row.problem
        or "投放错误" in row.indicator3
        or "不准确" in row.indicator3
    ):
        return "居民自主投放不准确"
    if "宣传引导" in row.indicator2 and ("没有看到" in row.problem or "无宣传" in row.problem):
        return "无宣传氛围"
    if problem == "周边不洁" and "桶站周边不洁" in indicator3:
        return "桶站周边不洁"
    if problem.isdigit() and not indicator3.isdigit():
        return indicator3
    return problem


def problem_with_count(text: str) -> str:
    text = clean_problem_text(text)
    if not text:
        return ""
    if re.search(r"\d+处$", text):
        return text
    if text.endswith("一处"):
        return text[:-2] + "1处"
    if text.endswith("1处"):
        return text
    return text + "1处"


def problem_with_count_from_row(row: LedgerRow) -> str:
    text = effective_problem_text(row)
    if not text:
        return ""
    return problem_with_count(text)


def summarize_problem_texts(texts: Iterable[str]) -> str:
    problems: list[str] = []
    seen: set[str] = set()
    for text in texts:
        if is_no_problem(text):
            continue
        problem = problem_with_count(text)
        if problem and problem not in seen:
            seen.add(problem)
            problems.append(problem)
    if not problems:
        return "无问题。"
    if len(problems) == 1:
        return f"（1）{problems[0]}。"
    return "；".join(f"（{index}）{problem}" for index, problem in enumerate(problems, start=1)) + "。"


def summarize_problem_rows(rows: Iterable[LedgerRow]) -> str:
    ordered: list[str] = []
    counts: dict[str, int] = {}
    for row in rows:
        if is_no_problem(row.problem) and is_no_problem(row.indicator3):
            continue
        problem = effective_problem_text(row)
        if not problem:
            continue
        if problem not in counts:
            ordered.append(problem)
            counts[problem] = 0
        counts[problem] += _problem_weight(row, problem)
    problems = [_format_problem_count(problem, counts[problem]) for problem in ordered]
    if not problems:
        return "无问题。"
    if len(problems) == 1:
        return f"（1）{problems[0]}。"
    return "；".join(f"（{index}）{problem}" for index, problem in enumerate(problems, start=1)) + "。"


def _format_problem_count(problem: str, count: int) -> str:
    problem = clean_problem_text(problem)
    if not problem:
        return ""
    if re.search(r"\d+处$", problem):
        return problem
    if problem.endswith("一处"):
        problem = problem[:-2]
    elif problem.endswith("1处"):
        problem = problem[:-2]
    return f"{problem}{count}处"


def _problem_weight(row: LedgerRow, problem: str) -> int:
    if (
        problem == "居民自主投放不准确"
        and "西绒线13号院" in row.place
        and "其他投厨余" in row.problem
    ):
        return 2
    return 1


def build_street_report(rows: list[LedgerRow], street: str) -> StreetReport:
    street_rows = [row for row in rows if row.street == street]
    report = StreetReport(street=street)
    grouped: dict[str, list[LedgerRow]] = defaultdict(list)
    for row in street_rows:
        grouped[row.category].append(row)

    report.communities = _build_communities(grouped.get(COMMUNITY_CATEGORY, []))
    report.restaurants = _build_units(grouped.get(RESTAURANT_CATEGORY, []))
    report.social_units = _build_units(grouped.get(SOCIAL_UNIT_CATEGORY, []))
    _dedupe_report_images(report)
    return report


def _group_by_place(rows: list[LedgerRow]) -> dict[str, list[LedgerRow]]:
    grouped: dict[str, list[LedgerRow]] = {}
    for row in rows:
        grouped.setdefault(row.place, []).append(row)
    return grouped


def _build_communities(rows: list[LedgerRow]) -> list[CommunitySection]:
    communities: list[CommunitySection] = []
    for index, (place, place_rows) in enumerate(_group_by_place(rows).items(), start=1):
        is_pure_box_room = _is_pure_box_room(place, place_rows)
        if is_pure_box_room:
            overall_rows = [row for row in place_rows if "居民自主投放" in row.indicator2]
        else:
            overall_rows = [
                row
                for row in place_rows
                if "小区宣传" not in row.indicator2
                and "小区公示牌" not in row.indicator2
            ]
        section = CommunitySection(
            index_cn=chinese_numeral(index),
            name=_display_place_name(place, place_rows),
            overall_problem_summary=summarize_problem_rows(overall_rows),
            overall_intro="",
            promo_images=_images_for_indicator(place_rows, "小区宣传"),
            notice_board_images=_images_for_indicator(place_rows, "小区公示牌"),
            is_pure_box_room=is_pure_box_room,
            stations=_build_pure_box_station_sections(place_rows) if is_pure_box_room else _build_station_sections(place_rows),
            resident_delivery=_build_resident_delivery(place_rows),
        )
        _attach_unassigned_community_images(section, place_rows)
        communities.append(section)
    return communities


def _is_pure_box_room(place: str, rows: list[LedgerRow]) -> bool:
    place_text = display_text(place)
    return (
        "纯箱房小区" in place_text
        or "纯厢房小区" in place_text
    )


def _build_station_sections(rows: list[LedgerRow]) -> list[StationSection]:
    fallback_rows = _non_resident_rows(rows)
    station_rows = [
        row
        for row in fallback_rows
        if "容器" in row.indicator2 or "投放点环境" in row.indicator2 or station_number_from_problem(row.problem)
    ]
    if not station_rows:
        return []
    return _build_station_sections_from_rows(station_rows, fallback_rows)


def _build_station_sections_from_rows(station_rows: list[LedgerRow], fallback_rows: list[LedgerRow]) -> list[StationSection]:
    grouped: dict[str, list[LedgerRow]] = defaultdict(list)
    for row in station_rows:
        station_no = station_number_from_problem(row.problem) or "1"
        grouped[station_no].append(row)
    sections: list[StationSection] = []
    for station_no, group in grouped.items():
        summary = summarize_station_problem_rows(group)
        images = _images_with_minimum(group, fallback_rows, minimum=3)
        sections.append(
            StationSection(
                title=f"{station_no}号桶站设置情况",
                station_no=station_no,
                problem_summary=summary,
                images=images,
            )
        )
    return _sort_station_sections(sections)


def _sort_station_sections(sections: list[StationSection]) -> list[StationSection]:
    if any(_station_sort_key(section.station_no) is None for section in sections):
        return sections
    return sorted(sections, key=lambda section: _station_sort_key(section.station_no) or (0, ""))


def _station_sort_key(station_no: str) -> tuple[int, str] | None:
    match = re.fullmatch(r"(\d+)([A-Z]?)", station_no)
    if not match:
        return None
    return int(match.group(1)), match.group(2)


def _build_pure_box_station_sections(rows: list[LedgerRow]) -> list[StationSection]:
    fallback_rows = _non_resident_rows(rows)
    station_rows = [
        row
        for row in fallback_rows
        if (
            station_number_from_problem(row.problem)
            or "环境" in row.indicator2
            or "容器" in row.problem
            or "容器" in row.indicator3
            or "投放点" in row.indicator2
        )
    ]
    if station_rows:
        return _build_station_sections_from_rows(station_rows, fallback_rows)
    station_images = _images_with_minimum(station_rows, fallback_rows, minimum=3)
    return [StationSection(title="1号桶站设置情况", station_no="1", problem_summary="无问题", images=station_images)]


def summarize_station_problem_texts(texts: Iterable[str]) -> str:
    summary = summarize_problem_texts(texts)
    if summary == "无问题。":
        return "无问题"
    return re.sub(r"^（1）(.+?)。$", r"\1", summary)


def summarize_station_problem_rows(rows: Iterable[LedgerRow]) -> str:
    summary = summarize_problem_rows(rows)
    if summary == "无问题。":
        return "无问题"
    if "；" not in summary:
        return re.sub(r"^（1）(.+?)1处。$", r"\1", summary)
    return summary


def station_number_from_problem(text: str) -> str | None:
    match = re.search(r"([0-9A-Za-z一二三四五六七八九十]+)(?:号)?桶站", display_text(text))
    return _normalize_station_no(match.group(1)) if match else None


def _normalize_station_no(station_no: str) -> str:
    match = re.fullmatch(r"0*(\d+)([A-Za-z]?)", station_no)
    if match:
        suffix = match.group(2).upper()
        return f"{int(match.group(1))}{suffix}"
    return station_no


def _build_resident_delivery(rows: list[LedgerRow]) -> ResidentDeliverySection | None:
    resident_rows = [row for row in rows if "居民自主投放" in row.indicator2]
    if not resident_rows:
        return None
    inaccurate = sum(_problem_weight(row, "居民自主投放不准确") for row in resident_rows if _is_resident_error(row))
    if inaccurate == 0:
        return None
    error_rows = [
        row
        for row in resident_rows
        if _is_resident_error(row)
    ]
    return ResidentDeliverySection(summary=f"居民投放共5个，其中不准确{inaccurate}个。", error_images=_collect_images(error_rows))


def _is_resident_error(row: LedgerRow) -> bool:
    return (
        "投放错误" in row.problem
        or "不准确" in row.problem
        or "投放错误" in row.indicator3
        or "不准确" in row.indicator3
    )


def _is_resident_delivery_row(row: LedgerRow) -> bool:
    return "居民自主投放" in row.indicator2


def _non_resident_rows(rows: Iterable[LedgerRow]) -> list[LedgerRow]:
    return [row for row in rows if not _is_resident_delivery_row(row)]


def _build_units(rows: list[LedgerRow]) -> list[UnitSection]:
    units: list[UnitSection] = []
    for index, (place, place_rows) in enumerate(_group_by_place(rows).items(), start=1):
        promo_rows = [row for row in place_rows if row.indicator2 == "宣传引导情况"]
        container_rows = [row for row in place_rows if "容器" in row.indicator2 or "容器" in row.indicator3]
        if not container_rows:
            container_rows = [row for row in place_rows if row.image_paths]
        promo_images = _collect_images(promo_rows) if promo_rows else _stable_sample_images(place, _collect_images(place_rows), 2)
        container_images = _collect_images(container_rows)
        container_images.extend(_remaining_images(place_rows, promo_images + container_images))
        units.append(
            UnitSection(
                index_cn=chinese_numeral(index),
                name=place,
                overall_problem_summary=summarize_problem_rows(place_rows),
                has_promo=True,
                promo_text=summarize_station_problem_rows(promo_rows) if promo_rows else "",
                container_problem_summary=summarize_station_problem_rows(container_rows),
                promo_images=_unique_images(promo_images),
                container_images=_unique_images(container_images),
            )
        )
    return units


def _stable_sample_images(place: str, images: list[Path], count: int) -> list[Path]:
    if count <= 0 or not images:
        return []
    if len(images) <= count:
        return images
    seed = hashlib.sha1(place.encode("utf-8")).hexdigest()
    return random.Random(seed).sample(images, count)


def _collect_images(rows: Iterable[LedgerRow]) -> list[Path]:
    images: list[Path] = []
    for row in rows:
        images.extend(row.image_paths)
    return _unique_images(images)


def _images_with_minimum(primary_rows: Iterable[LedgerRow], fallback_rows: Iterable[LedgerRow], minimum: int) -> list[Path]:
    primary_images = _collect_images(primary_rows)
    if len(primary_images) >= minimum:
        return primary_images
    needed = minimum - len(primary_images)
    fallback_images = _remaining_images(fallback_rows, primary_images)
    return _unique_images(primary_images + fallback_images[:needed])


def _attach_unassigned_community_images(section: CommunitySection, rows: list[LedgerRow]) -> None:
    assigned = section.promo_images + section.notice_board_images
    for station in section.stations:
        assigned.extend(station.images)
    if section.resident_delivery:
        assigned.extend(section.resident_delivery.error_images)

    remaining = _remaining_images(_non_resident_rows(rows), assigned)
    if not remaining:
        return

    if section.stations:
        section.stations[0].images = _unique_images(section.stations[0].images + remaining)
        return

    section.stations.append(
        StationSection(
            title="1号桶站设置情况",
            station_no="1",
            problem_summary="无问题",
            images=remaining,
        )
    )


def _remaining_images(rows: Iterable[LedgerRow], assigned: Iterable[Path]) -> list[Path]:
    assigned_set = {path.resolve() for path in assigned}
    return [path for path in _collect_images(rows) if path.resolve() not in assigned_set]


def _unique_images(images: Iterable[Path]) -> list[Path]:
    unique: list[Path] = []
    seen: set[Path] = set()
    for image in images:
        key = image.resolve()
        if key in seen:
            continue
        seen.add(key)
        unique.append(image)
    return unique


def _dedupe_report_images(report: StreetReport) -> None:
    seen: set[str] = set()

    for community in report.communities:
        community.promo_images = _dedupe_images_by_content(community.promo_images, seen)
        community.notice_board_images = _dedupe_images_by_content(community.notice_board_images, seen)
        if community.resident_delivery:
            community.resident_delivery.error_images = _dedupe_images_by_content(
                community.resident_delivery.error_images,
                seen,
            )
        for station in community.stations:
            station.images = _dedupe_images_by_content(station.images, seen)

    for unit in report.restaurants + report.social_units:
        unit.promo_images = _dedupe_images_by_content(unit.promo_images, seen)
        unit.container_images = _dedupe_images_by_content(unit.container_images, seen)


def _dedupe_images_by_content(images: Iterable[Path], seen: set[str]) -> list[Path]:
    unique: list[Path] = []
    for image in images:
        key = _image_content_key(image)
        if key in seen:
            continue
        seen.add(key)
        unique.append(image)
    return unique


def _image_content_key(image: Path) -> str:
    try:
        return hashlib.sha1(image.read_bytes()).hexdigest()
    except OSError:
        return str(image.resolve())


def _display_place_name(place: str, rows: list[LedgerRow]) -> str:
    if any(
        "不是箱房小区" in display_text(row.problem)
        or "不是厢房小区" in display_text(row.problem)
        or "不是箱房小区" in display_text(row.indicator3)
        or "不是厢房小区" in display_text(row.indicator3)
        for row in rows
    ):
        return f"{place}（经检查员核实，该小区不属于纯箱房小区）"
    return place





def _images_for_indicator(rows: list[LedgerRow], indicator: str) -> list[Path]:
    images: list[Path] = []
    for row in rows:
        if indicator in row.indicator2:
            images.extend(row.image_paths)
    return images


def chinese_numeral(number: int) -> str:
    values = "零一二三四五六七八九十"
    if 0 <= number <= 10:
        return values[number]
    if number < 20:
        return "十" + values[number % 10]
    return str(number)
