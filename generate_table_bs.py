from __future__ import annotations

import argparse
import io
import posixpath
import re
import sys
import zipfile
from dataclasses import dataclass
from datetime import date, timedelta
from decimal import Decimal, ROUND_CEILING, ROUND_HALF_UP
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
import xml.etree.ElementTree as ET

from excel_check_tool import (
    MAIN_NS,
    WorkbookError,
    SpreadsheetZip,
    SheetXml,
    WorkbookSheet,
    qn,
    workbook_sheets,
)


XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

ET.register_namespace("xdr", XDR_NS)
ET.register_namespace("a", A_NS)
ET.register_namespace("r", R_NS)


DEFAULT_MORNING_START = "06:00"
DEFAULT_MORNING_END = "12:00"
DEFAULT_AFTERNOON_START = "14:00"
DEFAULT_AFTERNOON_END = "18:00"
DEFAULT_NORMAL_HOURS = Decimal("10")
WINDOW_PAYABLE_CODES = {"V", "S"}
PAYABLE_CODES = {"V", "S"}
LEAVE_LABELS = {
    "A": "Absent",
    "E": "Emergency Leave",
    "S": "Sick Leave",
    "V": "Vacation",
}
SIGNATURE_FONT_PATH = Path("Nothing_You_Could_Do") / "NothingYouCouldDo-Regular.ttf"
SIGNATURE_MEDIA_WIDTH = 900
SIGNATURE_MEDIA_HEIGHT = 260
MIN_SIGNATURE_SCALE = 30
MAX_SIGNATURE_SCALE = 200


@dataclass(frozen=True)
class WorkSchedule:
    morning_start: Decimal
    morning_end: Decimal
    afternoon_start: Decimal
    afternoon_end: Decimal
    normal_hours: Decimal

    @property
    def morning_hours(self) -> Decimal:
        return (self.morning_end - self.morning_start) * Decimal("24")


def parse_time_to_day_fraction(value: str) -> Decimal:
    text = (value or "").strip()
    match = re.fullmatch(r"(\d{1,2}):(\d{2})", text)
    if not match:
        raise WorkbookError(f"时间格式无效: {value!r}，请使用 HH:MM，例如 06:00")
    hour = int(match.group(1))
    minute = int(match.group(2))
    if hour < 0 or hour > 23 or minute < 0 or minute > 59:
        raise WorkbookError(f"时间超出范围: {value!r}")
    return (Decimal(hour) + Decimal(minute) / Decimal("60")) / Decimal("24")


def build_work_schedule(
    morning_start: str = DEFAULT_MORNING_START,
    morning_end: str = DEFAULT_MORNING_END,
    afternoon_start: str = DEFAULT_AFTERNOON_START,
    afternoon_end: str = DEFAULT_AFTERNOON_END,
    normal_hours: str | int | Decimal = DEFAULT_NORMAL_HOURS,
) -> WorkSchedule:
    try:
        normal = Decimal(str(normal_hours).strip())
    except Exception as exc:
        raise WorkbookError(f"常规工作小时数无效: {normal_hours!r}") from exc
    if normal <= 0:
        raise WorkbookError("常规工作小时数必须大于 0")
    schedule = WorkSchedule(
        morning_start=parse_time_to_day_fraction(morning_start),
        morning_end=parse_time_to_day_fraction(morning_end),
        afternoon_start=parse_time_to_day_fraction(afternoon_start),
        afternoon_end=parse_time_to_day_fraction(afternoon_end),
        normal_hours=normal,
    )
    if not (schedule.morning_start < schedule.morning_end <= schedule.afternoon_start < schedule.afternoon_end):
        raise WorkbookError("上下班时间顺序必须为：上午上班 < 上午下班 <= 下午上班 < 下午下班")
    return schedule


def col_to_num(col: str) -> int:
    value = 0
    for char in col:
        value = value * 26 + ord(char) - 64
    return value


def num_to_col(num: int) -> str:
    chars: List[str] = []
    while num > 0:
        num, remainder = divmod(num - 1, 26)
        chars.append(chr(65 + remainder))
    return "".join(reversed(chars))


def decimal_to_text(value: Decimal) -> str:
    normalized = value.quantize(Decimal("0.000000000000001"), rounding=ROUND_HALF_UP).normalize()
    text = format(normalized, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


def excel_serial_to_date(serial_value: str) -> date:
    serial = int(Decimal(serial_value))
    origin = date(1899, 12, 30)
    return origin + timedelta(days=serial)


def date_to_excel_serial(day: date) -> int:
    origin = date(1899, 12, 30)
    return (day - origin).days


def sanitize_filename_part(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\\\|?*]+', "-", value).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned = cleaned.rstrip(". ")
    return cleaned or "blank"


def cell_style(sheet: SheetXml, cell_ref: str) -> int:
    cell = sheet.get_cell(cell_ref)
    if cell is None:
        return 0
    return int(cell.attrib.get("s", "0"))


def remove_children(cell: ET.Element, tags: Iterable[str]) -> None:
    for tag in tags:
        for node in list(cell.findall(qn(tag))):
            cell.remove(node)


def set_cell_number(sheet: SheetXml, cell_ref: str, value: Optional[str | int | Decimal]) -> None:
    cell = sheet.ensure_cell(cell_ref)
    remove_children(cell, ("f", "is", "v"))
    cell.attrib.pop("t", None)
    if value in (None, ""):
        return
    node = ET.SubElement(cell, qn("v"))
    if isinstance(value, Decimal):
        node.text = decimal_to_text(value)
    else:
        node.text = str(value)


def set_cell_text(sheet: SheetXml, cell_ref: str, value: Optional[str]) -> None:
    cell = sheet.ensure_cell(cell_ref)
    remove_children(cell, ("f", "is", "v"))
    if value in (None, ""):
        cell.attrib.pop("t", None)
        return
    cell.attrib["t"] = "inlineStr"
    is_node = ET.SubElement(cell, qn("is"))
    text_node = ET.SubElement(is_node, qn("t"))
    text_node.text = value


def signature_text(employee_name: str) -> str:
    parts = [part for part in re.split(r"\s+", employee_name.strip()) if part]
    return " ".join(parts[:2]) if len(parts) >= 2 else " ".join(parts)


def app_resource_path(relative_path: Path) -> Path:
    base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base_dir / relative_path


def validate_signature_scale(signature_scale: int | str) -> int:
    try:
        scale = int(str(signature_scale).strip().rstrip("%"))
    except ValueError as exc:
        raise WorkbookError("签名大小必须是数字，例如 100") from exc
    if scale < MIN_SIGNATURE_SCALE or scale > MAX_SIGNATURE_SCALE:
        raise WorkbookError(f"签名大小必须在 {MIN_SIGNATURE_SCALE}% 到 {MAX_SIGNATURE_SCALE}% 之间")
    return scale


def render_signature_png(signature: str, signature_scale: int = 100) -> bytes:
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError as exc:
        raise WorkbookError("生成签名图片需要 Pillow，请先安装: pip install pillow") from exc

    font_path = app_resource_path(SIGNATURE_FONT_PATH)
    if not font_path.exists():
        raise WorkbookError(f"缺少签名字体文件: {font_path}")

    scale = validate_signature_scale(signature_scale)
    width_ratio = min(0.98, 0.9 * scale / 100)
    height_ratio = min(0.9, 0.72 * scale / 100)
    font = ImageFont.truetype(str(font_path), int(150 * scale / 100))
    measure = Image.new("RGBA", (1, 1), (255, 255, 255, 0))
    measure_draw = ImageDraw.Draw(measure)
    bbox = measure_draw.textbbox((0, 0), signature, font=font)
    text_width = max(1, bbox[2] - bbox[0])
    text_height = max(1, bbox[3] - bbox[1])
    padding = max(24, int(40 * scale / 100))

    text_image = Image.new("RGBA", (text_width + padding * 2, text_height + padding * 2), (255, 255, 255, 0))
    text_draw = ImageDraw.Draw(text_image)
    text_draw.text((padding - bbox[0], padding - bbox[1]), signature, fill=(20, 20, 20, 245), font=font)

    max_width = int(SIGNATURE_MEDIA_WIDTH * width_ratio)
    max_height = int(SIGNATURE_MEDIA_HEIGHT * height_ratio)
    resize_ratio = min(max_width / text_image.width, max_height / text_image.height, 1)
    resized = text_image.resize(
        (max(1, int(text_image.width * resize_ratio)), max(1, int(text_image.height * resize_ratio))),
        Image.Resampling.LANCZOS,
    )

    image = Image.new("RGBA", (SIGNATURE_MEDIA_WIDTH, SIGNATURE_MEDIA_HEIGHT), (255, 255, 255, 0))
    x = (SIGNATURE_MEDIA_WIDTH - resized.width) // 2
    y = (SIGNATURE_MEDIA_HEIGHT - resized.height) // 2 + int(SIGNATURE_MEDIA_HEIGHT * 0.04)
    image.alpha_composite(resized, (x, y))
    output = io.BytesIO()
    image.save(output, format="PNG")
    return output.getvalue()


def xdr(tag: str) -> str:
    return f"{{{XDR_NS}}}{tag}"


def a_tag(tag: str) -> str:
    return f"{{{A_NS}}}{tag}"


def rel_tag(tag: str) -> str:
    return f"{{{REL_NS}}}{tag}"


def content_type_tag(tag: str) -> str:
    return f"{{{CONTENT_TYPES_NS}}}{tag}"


def relationship_id(value: str) -> str:
    return f"{{{R_NS}}}{value}"


def anchor_bounds(anchor: ET.Element) -> Tuple[int, int, int, int]:
    from_node = anchor.find(xdr("from"))
    to_node = anchor.find(xdr("to"))
    if from_node is None or to_node is None:
        return 0, 0, 0, 0
    from_col = int((from_node.findtext(xdr("col")) or "0"))
    from_row = int((from_node.findtext(xdr("row")) or "0"))
    to_col = int((to_node.findtext(xdr("col")) or "0"))
    to_row = int((to_node.findtext(xdr("row")) or "0"))
    return from_col, from_row, to_col, to_row


def overlaps_range(anchor: ET.Element, min_col: int, min_row: int, max_col: int, max_row: int) -> bool:
    from_col, from_row, to_col, to_row = anchor_bounds(anchor)
    return from_col <= max_col and to_col >= min_col and from_row <= max_row and to_row >= min_row


def next_relationship_id(rels_root: ET.Element) -> str:
    max_id = 0
    for rel in rels_root.findall(rel_tag("Relationship")):
        rel_id = rel.attrib.get("Id", "")
        match = re.fullmatch(r"rId(\d+)", rel_id)
        if match:
            max_id = max(max_id, int(match.group(1)))
    return f"rId{max_id + 1}"


def next_picture_id(drawing_root: ET.Element) -> int:
    max_id = 0
    for node in drawing_root.iter(xdr("cNvPr")):
        raw_id = node.attrib.get("id")
        if raw_id and raw_id.isdigit():
            max_id = max(max_id, int(raw_id))
    return max_id + 1


def create_fallback_anchor(
    from_col: int,
    from_row: int,
    to_col: int,
    to_row: int,
    rel_id: str,
    picture_id: int,
    picture_name: str,
) -> ET.Element:
    anchor = ET.Element(xdr("twoCellAnchor"))
    from_node = ET.SubElement(anchor, xdr("from"))
    ET.SubElement(from_node, xdr("col")).text = str(from_col)
    ET.SubElement(from_node, xdr("colOff")).text = "0"
    ET.SubElement(from_node, xdr("row")).text = str(from_row)
    ET.SubElement(from_node, xdr("rowOff")).text = "0"
    to_node = ET.SubElement(anchor, xdr("to"))
    ET.SubElement(to_node, xdr("col")).text = str(to_col)
    ET.SubElement(to_node, xdr("colOff")).text = "0"
    ET.SubElement(to_node, xdr("row")).text = str(to_row)
    ET.SubElement(to_node, xdr("rowOff")).text = "0"
    pic = ET.SubElement(anchor, xdr("pic"))
    nv_pic_pr = ET.SubElement(pic, xdr("nvPicPr"))
    ET.SubElement(nv_pic_pr, xdr("cNvPr"), {"id": str(picture_id), "name": picture_name})
    c_nv_pic_pr = ET.SubElement(nv_pic_pr, xdr("cNvPicPr"))
    ET.SubElement(c_nv_pic_pr, a_tag("picLocks"), {"noChangeAspect": "1"})
    blip_fill = ET.SubElement(pic, xdr("blipFill"))
    ET.SubElement(blip_fill, a_tag("blip"), {relationship_id("embed"): rel_id})
    stretch = ET.SubElement(blip_fill, a_tag("stretch"))
    ET.SubElement(stretch, a_tag("fillRect"))
    sp_pr = ET.SubElement(pic, xdr("spPr"))
    ET.SubElement(sp_pr, a_tag("prstGeom"), {"prst": "rect"})
    ET.SubElement(anchor, xdr("clientData"))
    return anchor


def apply_signature_to_drawing(
    book: SpreadsheetZip,
    replacements: Dict[str, bytes],
    drawing_part: str,
    signature_media_path: str,
    target_range: Tuple[int, int, int, int],
    fallback_anchor: Tuple[int, int, int, int],
) -> None:
    rels_part = f"{posixpath.dirname(drawing_part)}/_rels/{posixpath.basename(drawing_part)}.rels"
    drawing_data = replacements[drawing_part] if drawing_part in replacements else book.raw_entries[drawing_part]
    rels_data = replacements[rels_part] if rels_part in replacements else book.raw_entries[rels_part]
    drawing_root = ET.fromstring(drawing_data)
    rels_root = ET.fromstring(rels_data)

    removed_anchor: Optional[ET.Element] = None
    min_col, min_row, max_col, max_row = target_range
    for anchor in list(drawing_root):
        if anchor.find(xdr("pic")) is None:
            continue
        if overlaps_range(anchor, min_col, min_row, max_col, max_row):
            if removed_anchor is None:
                removed_anchor = anchor
            drawing_root.remove(anchor)

    rel_id = next_relationship_id(rels_root)
    ET.SubElement(
        rels_root,
        rel_tag("Relationship"),
        {
            "Id": rel_id,
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            "Target": f"../media/{posixpath.basename(signature_media_path)}",
        },
    )

    if removed_anchor is not None:
        new_anchor = removed_anchor
        blip = new_anchor.find(f".//{a_tag('blip')}")
        if blip is not None:
            blip.attrib[relationship_id("embed")] = rel_id
        c_nv_pr = new_anchor.find(f".//{xdr('cNvPr')}")
        if c_nv_pr is not None:
            c_nv_pr.attrib["id"] = str(next_picture_id(drawing_root))
            c_nv_pr.attrib["name"] = "Generated Signature"
    else:
        new_anchor = create_fallback_anchor(
            *fallback_anchor,
            rel_id=rel_id,
            picture_id=next_picture_id(drawing_root),
            picture_name="Generated Signature",
        )
    drawing_root.append(new_anchor)
    replacements[drawing_part] = ET.tostring(drawing_root, encoding="utf-8", xml_declaration=True)
    replacements[rels_part] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)


def ensure_png_content_type(book: SpreadsheetZip, replacements: Dict[str, bytes]) -> None:
    root = ET.fromstring(replacements.get("[Content_Types].xml", book.raw_entries["[Content_Types].xml"]))
    for default in root.findall(content_type_tag("Default")):
        if default.attrib.get("Extension", "").lower() == "png":
            return
    ET.SubElement(root, content_type_tag("Default"), {"Extension": "png", "ContentType": "image/png"})
    replacements["[Content_Types].xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)


def ensure_drawing_content_type(book: SpreadsheetZip, replacements: Dict[str, bytes], drawing_part: str) -> None:
    root = ET.fromstring(replacements.get("[Content_Types].xml", book.raw_entries["[Content_Types].xml"]))
    part_name = f"/{drawing_part}"
    for override in root.findall(content_type_tag("Override")):
        if override.attrib.get("PartName") == part_name:
            return
    ET.SubElement(
        root,
        content_type_tag("Override"),
        {
            "PartName": part_name,
            "ContentType": "application/vnd.openxmlformats-officedocument.drawing+xml",
        },
    )
    replacements["[Content_Types].xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)


def set_calc_flags(book_root: ET.Element) -> None:
    calc = book_root.find(qn("calcPr"))
    if calc is None:
        calc = ET.SubElement(book_root, qn("calcPr"))
    calc.attrib["calcMode"] = "auto"
    calc.attrib["fullCalcOnLoad"] = "1"
    calc.attrib["forceFullCalc"] = "1"
    calc.attrib.setdefault("calcId", "191029")


def style_fill_signatures(styles_root: ET.Element) -> Dict[int, str]:
    fills_node = styles_root.find(qn("fills"))
    xfs_node = styles_root.find(qn("cellXfs"))
    fills = list(fills_node) if fills_node is not None else []
    xfs = list(xfs_node) if xfs_node is not None else []
    signatures: Dict[int, str] = {}
    for index, xf in enumerate(xfs):
        fill_id = int(xf.attrib.get("fillId", "0"))
        if fill_id < len(fills):
            signatures[index] = ET.tostring(fills[fill_id], encoding="unicode")
        else:
            signatures[index] = ""
    return signatures


def group_day_types(summary_sheet: SheetXml, styles_root: ET.Element) -> List[Tuple[str, int, str]]:
    fill_signatures = style_fill_signatures(styles_root)
    day_cells: List[Tuple[str, int, int]] = []
    column_index = col_to_num("J")
    while True:
        column = num_to_col(column_index)
        raw_day = summary_sheet.get_value(f"{column}2")
        if raw_day in (None, ""):
            break
        text = str(raw_day).strip()
        if not re.fullmatch(r"\d+", text):
            break
        day_cells.append((column, int(Decimal(text)), cell_style(summary_sheet, f"{column}2")))
        column_index += 1

    grouped: Dict[str, List[int]] = {}
    style_by_signature: Dict[str, int] = {}
    for _, day_number, style_id in day_cells:
        signature = fill_signatures.get(style_id, f"style:{style_id}")
        grouped.setdefault(signature, []).append(day_number)
        style_by_signature[signature] = style_id

    if not grouped:
        raise WorkbookError("表C 未找到日期头(J2开始)")

    work_signature = max(grouped.items(), key=lambda item: len(item[1]))[0]
    other_signatures = [signature for signature in grouped if signature != work_signature]

    def rest_score(signature: str) -> Tuple[int, int]:
        days = grouped[signature]
        diffs = [b - a for a, b in zip(days, days[1:])]
        return (sum(1 for diff in diffs if diff == 7), len(days))

    rest_signature = max(other_signatures, key=rest_score) if other_signatures else None
    holiday_signatures = {signature for signature in other_signatures if signature != rest_signature}

    result: List[Tuple[str, int, str]] = []
    for column, day_number, style_id in day_cells:
        signature = fill_signatures.get(style_id, f"style:{style_id}")
        if signature == work_signature:
            day_type = "work"
        elif signature == rest_signature:
            day_type = "rest"
        elif signature in holiday_signatures:
            day_type = "holiday"
        else:
            day_type = "work"
        result.append((column, day_number, day_type))
    return result


def parse_summary_value(raw_value: Optional[str]) -> Tuple[str, Optional[Decimal]]:
    if raw_value is None:
        return "blank", None
    text = str(raw_value).strip()
    if text == "":
        return "blank", None
    if text == "\\":
        return "blank", None
    upper = text.upper()
    if upper in LEAVE_LABELS:
        return upper, None
    try:
        return "hours", Decimal(text)
    except Exception as exc:  # pragma: no cover - defensive
        raise WorkbookError(f"无法解析表C考勤值: {raw_value!r}") from exc


def parse_optional_decimal(raw_value: Optional[str]) -> Decimal:
    if raw_value in (None, ""):
        return Decimal("0")
    text = str(raw_value).strip()
    if text in {"", "\\"}:
        return Decimal("0")
    try:
        return Decimal(text)
    except Exception as exc:
        raise WorkbookError(f"无法解析数字: {raw_value!r}") from exc


@dataclass
class SummaryEmployee:
    row_num: int
    no: int
    employee_no: str
    name: str
    project: str
    company: str
    passport: str
    crew_group: str
    position: str
    joining_date: Optional[str]
    days: Dict[int, Tuple[str, Optional[Decimal]]]
    correction_nwh: Decimal
    correction_normal_ot: Decimal
    correction_weekend_ot: Decimal
    correction_holiday_ot: Decimal


def read_summary(summary_path: Path) -> Tuple[SpreadsheetZip, WorkbookSheet, List[Tuple[str, int, str]], List[SummaryEmployee]]:
    book = SpreadsheetZip(summary_path)
    sheets = workbook_sheets(book)
    summary_sheet_info = sheets[0]
    summary_sheet = summary_sheet_info.sheet
    styles_root = book.load_xml("xl/styles.xml")
    day_headers = group_day_types(summary_sheet, styles_root)

    employees: List[SummaryEmployee] = []
    row_num = 3
    while True:
        name = (summary_sheet.get_value(f"C{row_num}") or "").strip()
        position = (summary_sheet.get_value(f"H{row_num}") or "").strip()
        if not name and not position:
            break
        no_raw = summary_sheet.get_value(f"A{row_num}")
        if no_raw in (None, ""):
            row_num += 1
            continue
        no_value = int(Decimal(str(no_raw)))
        days: Dict[int, Tuple[str, Optional[Decimal]]] = {}
        for column, day_number, _ in day_headers:
            days[day_number] = parse_summary_value(summary_sheet.get_value(f"{column}{row_num}"))
        employees.append(
            SummaryEmployee(
                row_num=row_num,
                no=no_value,
                employee_no=(summary_sheet.get_value(f"B{row_num}") or "").strip(),
                name=name,
                project=(summary_sheet.get_value(f"D{row_num}") or "").strip(),
                company=(summary_sheet.get_value(f"E{row_num}") or "").strip(),
                passport=(summary_sheet.get_value(f"F{row_num}") or "").strip(),
                crew_group=(summary_sheet.get_value(f"G{row_num}") or "").strip(),
                position=position,
                joining_date=summary_sheet.get_value(f"I{row_num}"),
                days=days,
                correction_nwh=parse_optional_decimal(summary_sheet.get_value(f"BQ{row_num}")),
                correction_normal_ot=parse_optional_decimal(summary_sheet.get_value(f"BR{row_num}")),
                correction_weekend_ot=parse_optional_decimal(summary_sheet.get_value(f"BS{row_num}")),
                correction_holiday_ot=parse_optional_decimal(summary_sheet.get_value(f"BT{row_num}")),
            )
        )
        row_num += 1

    return book, summary_sheet_info, day_headers, employees


def build_output_name(employee: SummaryEmployee, suffix: str) -> str:
    name = sanitize_filename_part(employee.name)
    position = sanitize_filename_part(employee.position)
    return f"{employee.no}.{name}-{position}{suffix}"


def day_time_inputs(
    total_hours: Decimal,
    schedule: WorkSchedule,
) -> Tuple[Optional[Decimal], Optional[Decimal], Optional[Decimal], Optional[Decimal]]:
    if total_hours <= 0:
        return None, None, None, None
    if total_hours <= schedule.morning_hours:
        return schedule.morning_start, None, None, schedule.morning_start + total_hours / Decimal("24")
    finish = schedule.afternoon_start + (total_hours - schedule.morning_hours) / Decimal("24")
    return schedule.morning_start, schedule.morning_end, schedule.afternoon_start, finish


def weekday_name(day: date) -> str:
    return day.strftime("%A")


def fill_row(
    sheet: SheetXml,
    row_num: int,
    current_date: date,
    day_type: str,
    entry_type: str,
    total_hours: Optional[Decimal],
    payable_context: Tuple[Optional[int], Optional[int], bool, bool],
    count_holidays: bool,
    unpaid_adjacent_special_days: set[int],
    schedule: WorkSchedule,
) -> Tuple[bool, Decimal, Decimal, Decimal, Decimal]:
    first_active, last_active, prefix_unpaid, suffix_unpaid = payable_context
    row_refs = [f"{col}{row_num}" for col in "ABCDEFGHIJKNOPQRT"]
    for ref in row_refs:
        if ref[0] in {"A", "B"}:
            continue
        if ref[0] in {"C", "D", "E", "F", "G", "H", "I", "J", "K", "N", "O", "P", "Q", "R", "T"}:
            if ref[0] in {"K"}:
                set_cell_text(sheet, ref, None)
            else:
                set_cell_number(sheet, ref, None)

    set_cell_number(sheet, f"A{row_num}", date_to_excel_serial(current_date))
    set_cell_text(sheet, f"B{row_num}", weekday_name(current_date))

    work_hours = Decimal("0")
    work_ot = Decimal("0")
    rest_hours = Decimal("0")
    holiday_hours = Decimal("0")
    payable = False

    if entry_type == "hours" and total_hours is not None and total_hours > 0:
        c_value, d_value, e_value, f_value = day_time_inputs(total_hours, schedule)
        set_cell_number(sheet, f"C{row_num}", c_value)
        set_cell_number(sheet, f"D{row_num}", d_value)
        set_cell_number(sheet, f"E{row_num}", e_value)
        set_cell_number(sheet, f"F{row_num}", f_value)
        set_cell_number(sheet, f"N{row_num}", f_value)
        if day_type == "work":
            work_hours = min(total_hours, schedule.normal_hours)
            work_ot = max(total_hours - schedule.normal_hours, Decimal("0"))
        elif day_type == "rest":
            rest_hours = total_hours
        else:
            holiday_hours = total_hours
        payable = True
    elif entry_type in LEAVE_LABELS:
        set_cell_text(sheet, f"K{row_num}", LEAVE_LABELS[entry_type])
        payable = entry_type in PAYABLE_CODES
    else:
        is_prefix = first_active is not None and current_date.day < first_active
        is_suffix = last_active is not None and current_date.day > last_active
        if day_type in {"rest", "holiday"}:
            payable = not ((is_prefix and prefix_unpaid) or (is_suffix and suffix_unpaid))
        if day_type in {"rest", "holiday"} and current_date.day in unpaid_adjacent_special_days:
            payable = False
        if day_type == "rest":
            set_cell_text(sheet, f"K{row_num}", "Weekend")
        elif day_type == "holiday":
            set_cell_text(sheet, f"K{row_num}", "Public Holiday")

    if entry_type not in LEAVE_LABELS and day_type == "rest":
        set_cell_text(sheet, f"K{row_num}", "Weekend")
    elif entry_type not in LEAVE_LABELS and day_type == "holiday":
        set_cell_text(sheet, f"K{row_num}", "Public Holiday")

    set_cell_number(sheet, f"G{row_num}", work_hours if work_hours > 0 else None)
    set_cell_number(sheet, f"H{row_num}", work_ot if work_ot > 0 else None)
    set_cell_number(sheet, f"I{row_num}", rest_hours if rest_hours > 0 else None)
    set_cell_number(sheet, f"J{row_num}", holiday_hours if holiday_hours > 0 else None)
    set_cell_number(sheet, f"P{row_num}", work_ot if work_ot > 0 else None)
    set_cell_number(sheet, f"Q{row_num}", rest_hours if rest_hours > 0 else None)
    set_cell_number(sheet, f"R{row_num}", holiday_hours if holiday_hours > 0 else None)
    if entry_type == "hours" and total_hours and total_hours > 0:
        set_cell_number(sheet, f"O{row_num}", Decimal("0"))
    else:
        set_cell_number(sheet, f"O{row_num}", None)
    set_cell_number(sheet, f"T{row_num}", 1 if payable else None)
    return payable, work_hours, work_ot, rest_hours, holiday_hours


@dataclass
class OvertimeEntry:
    day: date
    start: Optional[Decimal]
    end: Optional[Decimal]
    normal_hours: Decimal
    weekend_hours: Decimal
    holiday_hours: Decimal


def update_overtime_sheet(
    overtime_sheet: Optional[SheetXml],
    employee: SummaryEmployee,
    overtime_entries: List[OvertimeEntry],
    schedule: WorkSchedule,
) -> None:
    if overtime_sheet is None:
        return
    set_cell_text(overtime_sheet, "B5", employee.name)
    set_cell_text(overtime_sheet, "B7", employee.project)
    set_cell_text(overtime_sheet, "I7", employee.position)
    correction_row = find_correction_row(overtime_sheet)
    detail_end_row = (correction_row - 1) if correction_row is not None else 42
    for row in range(12, detail_end_row + 1):
        for col in ["A", "B", "E", "H", "I", "J"]:
            set_cell_number(overtime_sheet, f"{col}{row}", None)
    if correction_row is not None:
        for col in ["H", "I", "J"]:
            set_cell_number(overtime_sheet, f"{col}{correction_row}", None)
    for row, entry in enumerate(overtime_entries[: detail_end_row - 11], start=12):
        set_cell_number(overtime_sheet, f"A{row}", date_to_excel_serial(entry.day))
        set_cell_number(overtime_sheet, f"B{row}", entry.start)
        set_cell_number(overtime_sheet, f"E{row}", entry.end)
        set_cell_number(overtime_sheet, f"H{row}", entry.normal_hours if entry.normal_hours > 0 else None)
        set_cell_number(overtime_sheet, f"I{row}", entry.weekend_hours if entry.weekend_hours > 0 else None)
        set_cell_number(overtime_sheet, f"J{row}", entry.holiday_hours if entry.holiday_hours > 0 else None)
    if correction_row is not None:
        set_cell_number(overtime_sheet, f"H{correction_row}", employee.correction_normal_ot if employee.correction_normal_ot != 0 else None)
        set_cell_number(overtime_sheet, f"I{correction_row}", employee.correction_weekend_ot if employee.correction_weekend_ot != 0 else None)
        set_cell_number(overtime_sheet, f"J{correction_row}", employee.correction_holiday_ot if employee.correction_holiday_ot != 0 else None)

    normal_total = sum((entry.normal_hours for entry in overtime_entries), Decimal("0"))
    weekend_total = sum((entry.weekend_hours for entry in overtime_entries), Decimal("0"))
    holiday_total = sum((entry.holiday_hours for entry in overtime_entries), Decimal("0"))
    set_cell_number(overtime_sheet, "H43", normal_total if normal_total > 0 else None)
    set_cell_number(overtime_sheet, "I43", weekend_total if weekend_total > 0 else None)
    set_cell_number(overtime_sheet, "J43", holiday_total if holiday_total > 0 else None)
    set_cell_number(overtime_sheet, "H44", normal_total / schedule.normal_hours if normal_total > 0 else None)
    set_cell_number(overtime_sheet, "I44", weekend_total / schedule.normal_hours if weekend_total > 0 else None)
    set_cell_number(overtime_sheet, "J44", holiday_total / schedule.normal_hours if holiday_total > 0 else None)


def sheet_drawing_part(book: SpreadsheetZip, sheet_part: str, sheet: SheetXml) -> Optional[str]:
    drawing = sheet.root.find(qn("drawing"))
    if drawing is None:
        return None
    rel_id = drawing.attrib.get(relationship_id("id"))
    if not rel_id:
        return None
    rels_part = f"{posixpath.dirname(sheet_part)}/_rels/{posixpath.basename(sheet_part)}.rels"
    if rels_part not in book.raw_entries:
        return None
    rels_root = book.load_xml(rels_part)
    for rel in rels_root.findall(rel_tag("Relationship")):
        if rel.attrib.get("Id") != rel_id:
            continue
        target = rel.attrib.get("Target", "")
        return posixpath.normpath(posixpath.join(posixpath.dirname(sheet_part), target))
    return None


def next_drawing_part(book: SpreadsheetZip, replacements: Dict[str, bytes]) -> str:
    used_numbers = set()
    for filename in set(book.raw_entries) | set(replacements):
        match = re.fullmatch(r"xl/drawings/drawing(\d+)\.xml", filename)
        if match:
            used_numbers.add(int(match.group(1)))
    number = 1
    while number in used_numbers:
        number += 1
    return f"xl/drawings/drawing{number}.xml"


def ensure_relationships_root(book: SpreadsheetZip, replacements: Dict[str, bytes], rels_part: str) -> ET.Element:
    if rels_part in replacements:
        return ET.fromstring(replacements[rels_part])
    if rels_part in book.raw_entries:
        return book.load_xml(rels_part)
    return ET.Element(rel_tag("Relationships"))


def ensure_sheet_drawing_part(
    book: SpreadsheetZip,
    replacements: Dict[str, bytes],
    sheet_info: WorkbookSheet,
) -> str:
    existing = sheet_drawing_part(book, sheet_info.part_name, sheet_info.sheet)
    if existing is not None:
        return existing

    drawing_part = next_drawing_part(book, replacements)
    sheet_rels_part = f"{posixpath.dirname(sheet_info.part_name)}/_rels/{posixpath.basename(sheet_info.part_name)}.rels"
    sheet_rels_root = ensure_relationships_root(book, replacements, sheet_rels_part)
    rel_id = next_relationship_id(sheet_rels_root)
    target = posixpath.relpath(drawing_part, posixpath.dirname(sheet_info.part_name))
    ET.SubElement(
        sheet_rels_root,
        rel_tag("Relationship"),
        {
            "Id": rel_id,
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
            "Target": target,
        },
    )
    ET.SubElement(sheet_info.sheet.root, qn("drawing"), {relationship_id("id"): rel_id})

    drawing_root = ET.Element(xdr("wsDr"))
    drawing_rels_part = f"{posixpath.dirname(drawing_part)}/_rels/{posixpath.basename(drawing_part)}.rels"
    drawing_rels_root = ET.Element(rel_tag("Relationships"))

    replacements[sheet_info.part_name] = ET.tostring(sheet_info.sheet.root, encoding="utf-8", xml_declaration=True)
    replacements[sheet_rels_part] = ET.tostring(sheet_rels_root, encoding="utf-8", xml_declaration=True)
    replacements[drawing_part] = ET.tostring(drawing_root, encoding="utf-8", xml_declaration=True)
    replacements[drawing_rels_part] = ET.tostring(drawing_rels_root, encoding="utf-8", xml_declaration=True)
    ensure_drawing_content_type(book, replacements, drawing_part)
    return drawing_part


def choose_template_sheet(sheets: List[WorkbookSheet], sheet_name: str) -> Optional[WorkbookSheet]:
    for sheet in sheets:
        if sheet.display_name == sheet_name:
            return sheet
    return None


def find_correction_row(sheet: SheetXml, start_row: int = 35, end_row: int = 45) -> Optional[int]:
    for row in range(start_row, end_row + 1):
        text = sheet.get_value(f"A{row}") or ""
        if "修正上月加班" in str(text) or "FIX OT" in str(text).upper():
            return row
    return None


def write_employee_workbook(
    template_path: Path,
    employee: SummaryEmployee,
    day_headers: List[Tuple[str, int, str]],
    output_dir: Path,
    count_holidays: bool = False,
    signature_scale: int = 100,
    schedule: Optional[WorkSchedule] = None,
) -> Path:
    schedule = schedule or build_work_schedule()
    book = SpreadsheetZip(template_path)
    workbook_root = book.load_xml("xl/workbook.xml")
    sheets = workbook_sheets(book)
    main_sheet_info = choose_template_sheet(sheets, "New timesheet") or sheets[0]
    overtime_sheet_info = choose_template_sheet(sheets, "Overtime")
    main_sheet = main_sheet_info.sheet
    overtime_sheet = overtime_sheet_info.sheet if overtime_sheet_info else None

    month_serial = main_sheet.get_value("M3")
    if month_serial in (None, ""):
        raise WorkbookError("模板表B缺少 M3 月份信息")
    month_start = excel_serial_to_date(month_serial)
    month_days = (date(month_start.year + (month_start.month // 12), ((month_start.month % 12) + 1), 1) - timedelta(days=1)).day

    if month_days > len(day_headers):
        raise WorkbookError(
            f"模板月份天数({month_days})与表C日期列数量({len(day_headers)})不一致，请确认模板月份是否正确"
        )
    day_headers = day_headers[:month_days]

    holiday_days = [day for _, day, kind in day_headers if kind == "holiday"]
    rest_days = [day for _, day, kind in day_headers if kind == "rest"]
    rest_weekdays = {weekday_name(date(month_start.year, month_start.month, day)) for day in rest_days}
    if len(rest_weekdays) > 1:
        raise WorkbookError(f"表C中检测到多个休息日星期: {sorted(rest_weekdays)}")
    rest_weekday = next(iter(rest_weekdays), "Friday")

    set_cell_text(main_sheet, "A3", employee.name)
    set_cell_text(main_sheet, "C3", employee.position)
    set_cell_text(main_sheet, "E3", employee.passport)
    set_cell_text(main_sheet, "G3", employee.project)
    set_cell_text(main_sheet, "I3", employee.crew_group)
    set_cell_text(main_sheet, "J3", employee.employee_no)
    set_cell_text(main_sheet, "N7", rest_weekday)
    set_cell_number(main_sheet, "N8", schedule.normal_hours)
    signature = signature_text(employee.name)
    signature_png = render_signature_png(signature, signature_scale)
    main_correction_row = find_correction_row(main_sheet)

    for row in range(50, 71):
        set_cell_number(main_sheet, f"J{row}", None)
    for index, day_number in enumerate(holiday_days, start=50):
        holiday_date = date(month_start.year, month_start.month, day_number)
        set_cell_number(main_sheet, f"J{index}", date_to_excel_serial(holiday_date))

    anchor_days = [
        day
        for _, day, _ in day_headers
        if (employee.days.get(day, ("blank", None))[0] == "hours" and (employee.days.get(day, ("blank", None))[1] or Decimal("0")) > 0)
        or employee.days.get(day, ("blank", None))[0] in WINDOW_PAYABLE_CODES
    ]
    first_active = min(anchor_days) if anchor_days else None
    last_active = max(anchor_days) if anchor_days else None

    def has_unpaid_marker(days: Iterable[int]) -> bool:
        type_by_day = {day: kind for _, day, kind in day_headers}
        for day in days:
            entry_type, _ = employee.days.get(day, ("blank", None))
            if entry_type in {"A", "E"}:
                return True
            if entry_type == "blank" and type_by_day.get(day) == "work":
                return True
        return False

    prefix_unpaid = first_active is not None and has_unpaid_marker(range(1, first_active))
    suffix_unpaid = last_active is not None and has_unpaid_marker(range(last_active + 1, len(day_headers) + 1))
    unpaid_leave_days = {day for day, (entry_type, _) in employee.days.items() if entry_type in {"A", "E"}}
    unpaid_adjacent_special_days = set()
    if not count_holidays:
        special_day_numbers = [day for _, day, day_type in day_headers if day_type in {"rest", "holiday"}]
        special_day_groups: List[List[int]] = []
        for day_number in special_day_numbers:
            if not special_day_groups or day_number != special_day_groups[-1][-1] + 1:
                special_day_groups.append([day_number])
            else:
                special_day_groups[-1].append(day_number)
        for group in special_day_groups:
            if not ({group[0] - 1, group[-1] + 1} & unpaid_leave_days):
                continue
            for day_number in group:
                entry_type, total_hours = employee.days.get(day_number, ("blank", None))
                attended = entry_type == "hours" and total_hours is not None and total_hours > 0
                if not attended and entry_type not in {"S", "V"}:
                    unpaid_adjacent_special_days.add(day_number)

    payable_total = 0
    work_sum = Decimal("0")
    work_ot_sum = Decimal("0")
    rest_ot_sum = Decimal("0")
    holiday_ot_sum = Decimal("0")
    vacation_days = 0
    sick_days = 0
    emergency_days = 0
    overtime_entries: List[OvertimeEntry] = []

    for _, day_number, day_type in day_headers:
        current_date = date(month_start.year, month_start.month, day_number)
        entry_type, total_hours = employee.days.get(day_number, ("blank", None))
        row = 9 + day_number
        payable, work_hours, work_ot, rest_hours, holiday_hours = fill_row(
            main_sheet,
            row,
            current_date,
            day_type,
            entry_type,
            total_hours,
            (first_active, last_active, prefix_unpaid, suffix_unpaid),
            count_holidays,
            unpaid_adjacent_special_days,
            schedule,
        )
        payable_total += 1 if payable else 0
        work_sum += work_hours
        work_ot_sum += work_ot
        rest_ot_sum += rest_hours
        holiday_ot_sum += holiday_hours
        if work_ot > 0 or rest_hours > 0 or holiday_hours > 0:
            if day_type == "work":
                start_time = schedule.afternoon_start + schedule.normal_hours / Decimal("24")
                end_time = start_time + work_ot / Decimal("24")
            else:
                start_time, _, _, end_time = day_time_inputs(total_hours or Decimal("0"), schedule)
            overtime_entries.append(
                OvertimeEntry(
                    day=current_date,
                    start=start_time,
                    end=end_time,
                    normal_hours=work_ot,
                    weekend_hours=rest_hours,
                    holiday_hours=holiday_hours,
                )
            )
        if entry_type == "V":
            vacation_days += 1
        elif entry_type == "S":
            sick_days += 1
        elif entry_type == "E":
            emergency_days += 1

    for day_number in range(len(day_headers) + 1, 32):
        row = 9 + day_number
        if row == main_correction_row:
            continue
        for col in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "N", "O", "P", "Q", "R", "T"]:
            ref = f"{col}{row}"
            if col == "K":
                set_cell_text(main_sheet, ref, None)
            else:
                set_cell_number(main_sheet, ref, None)

    if main_correction_row is not None:
        set_cell_number(main_sheet, f"G{main_correction_row}", employee.correction_nwh if employee.correction_nwh != 0 else None)
        set_cell_number(main_sheet, f"H{main_correction_row}", employee.correction_normal_ot if employee.correction_normal_ot != 0 else None)
        set_cell_number(main_sheet, f"I{main_correction_row}", employee.correction_weekend_ot if employee.correction_weekend_ot != 0 else None)
        set_cell_number(main_sheet, f"J{main_correction_row}", employee.correction_holiday_ot if employee.correction_holiday_ot != 0 else None)

    public_payable = 0
    rest_payable = 0
    public_attendance = 0
    rest_attendance = 0
    for _, day_number, kind in day_headers:
        entry_type, total_hours = employee.days.get(day_number, ("blank", None))
        attended = entry_type == "hours" and total_hours is not None and total_hours > 0
        if kind == "holiday" and attended:
            public_attendance += 1
        if kind == "rest" and attended:
            rest_attendance += 1
        if not count_holidays and entry_type == "V":
            continue
        if kind == "holiday" and main_sheet.get_value(f"T{9 + day_number}") not in (None, ""):
            public_payable += 1
        if kind == "rest" and main_sheet.get_value(f"T{9 + day_number}") not in (None, ""):
            rest_payable += 1

    work_day_count = (work_sum / schedule.normal_hours).to_integral_value(rounding=ROUND_CEILING) if work_sum > 0 else Decimal("0")
    public_day_count = Decimal(public_attendance if count_holidays else public_payable)
    vacation_day_count = Decimal(vacation_days)
    sick_day_count = Decimal(0 if count_holidays else sick_days)
    rest_day_count = Decimal(rest_attendance if count_holidays else rest_payable)
    payable_day_count = work_day_count + public_day_count + sick_day_count + rest_day_count
    if not count_holidays:
        payable_day_count += vacation_day_count

    set_cell_number(main_sheet, "A6", len(day_headers))
    set_cell_number(main_sheet, "B6", payable_day_count if payable_day_count > 0 else None)
    set_cell_number(main_sheet, "E6", work_day_count if work_day_count > 0 else None)
    set_cell_number(main_sheet, "I6", public_day_count if public_day_count > 0 else None)
    set_cell_number(main_sheet, "J6", vacation_day_count if vacation_day_count > 0 else None)
    set_cell_number(main_sheet, "K6", sick_day_count if sick_day_count > 0 else None)
    set_cell_number(main_sheet, "L6", rest_day_count if rest_day_count > 0 else None)
    set_cell_number(main_sheet, "M6", None if count_holidays else (emergency_days if emergency_days > 0 else None))
    set_cell_number(main_sheet, "G9", work_sum if work_sum > 0 else None)
    set_cell_number(main_sheet, "H9", work_ot_sum if work_ot_sum > 0 else None)
    set_cell_number(main_sheet, "I9", rest_ot_sum if rest_ot_sum > 0 else None)
    set_cell_number(main_sheet, "J9", holiday_ot_sum if holiday_ot_sum > 0 else None)

    update_overtime_sheet(overtime_sheet, employee, overtime_entries, schedule)
    set_calc_flags(workbook_root)

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / build_output_name(employee, template_path.suffix)
    signature_media_path = "xl/media/generated_signature.png"
    replacements = {
        "xl/workbook.xml": ET.tostring(workbook_root, encoding="utf-8", xml_declaration=True),
        main_sheet_info.part_name: ET.tostring(main_sheet.root, encoding="utf-8", xml_declaration=True),
        signature_media_path: signature_png,
    }
    ensure_png_content_type(book, replacements)
    if overtime_sheet_info is not None:
        replacements[overtime_sheet_info.part_name] = ET.tostring(overtime_sheet.root, encoding="utf-8", xml_declaration=True)
    main_drawing_part = ensure_sheet_drawing_part(book, replacements, main_sheet_info)
    apply_signature_to_drawing(
        book,
        replacements,
        main_drawing_part,
        signature_media_path,
        target_range=(0, 41, 6, 43),
        fallback_anchor=(1, 41, 3, 43),
    )
    if overtime_sheet_info is not None and overtime_sheet is not None:
        overtime_drawing_part = ensure_sheet_drawing_part(book, replacements, overtime_sheet_info)
        apply_signature_to_drawing(
            book,
            replacements,
            overtime_drawing_part,
            signature_media_path,
            target_range=(4, 52, 9, 52),
            fallback_anchor=(7, 51, 8, 53),
        )
    book.save(output_path, replacements)
    return output_path


def choose_template_file(template_arg: Optional[str], table_bs_dir: Optional[str]) -> Path:
    if template_arg:
        path = Path(template_arg)
        if not path.exists():
            raise WorkbookError("模板表B不存在")
        return path

    if table_bs_dir:
        all_candidates = sorted(Path(table_bs_dir).iterdir())
        candidates = [
            path
            for path in all_candidates
            if path.is_file()
            and path.suffix.lower() in {".xlsm", ".xlsx"}
            and zipfile.is_zipfile(path)
        ]
        candidates.sort(key=lambda item: (item.suffix.lower() != ".xlsm", item.name.casefold()))
        if candidates:
            return candidates[0]

    raise WorkbookError("请提供 --template-b，或提供含有现成表B模板的 --table-bs-dir")


def next_output_dir(summary_path: Path) -> Path:
    return summary_path.with_name(f"{summary_path.stem}_生成表Bs")


def create_generation_report(
    report_path: Path,
    generated_files: List[Path],
    template_path: Path,
    count_holidays: bool,
    signature_scale: int,
    schedule: WorkSchedule,
) -> None:
    lines = [
        f"模板表B: {template_path}",
        f"生成数量: {len(generated_files)}",
        f"是否统计假期: {'是' if count_holidays else '否'}",
        f"签名大小: {signature_scale}%",
        f"常规工作小时数: {decimal_to_text(schedule.normal_hours)}",
        "默认规则:",
        "- V(年休假) 计入可支付天数",
        "- S(病假) 计入可支付天数",
        "- E(紧急休假) 不计入可支付天数",
        "- A(缺勤) 不计入可支付天数",
        "",
        "文件列表:",
    ]
    lines.extend(f"- {path.name}" for path in generated_files)
    report_path.write_text("\n".join(lines), encoding="utf-8")


def run_generate(
    table_c_path: Path,
    template_b_path: Path,
    output_dir: Optional[Path] = None,
    count_holidays: bool = False,
    signature_scale: int = 100,
    morning_start: str = DEFAULT_MORNING_START,
    morning_end: str = DEFAULT_MORNING_END,
    afternoon_start: str = DEFAULT_AFTERNOON_START,
    afternoon_end: str = DEFAULT_AFTERNOON_END,
    normal_hours: str | int | Decimal = DEFAULT_NORMAL_HOURS,
) -> Tuple[Path, List[Path], Path]:
    signature_scale = validate_signature_scale(signature_scale)
    schedule = build_work_schedule(morning_start, morning_end, afternoon_start, afternoon_end, normal_hours)
    _, _, day_headers, employees = read_summary(table_c_path)
    output = output_dir or next_output_dir(table_c_path)
    generated_files: List[Path] = []
    for employee in employees:
        generated_files.append(
            write_employee_workbook(
                template_b_path,
                employee,
                day_headers,
                output,
                count_holidays=count_holidays,
                signature_scale=signature_scale,
                schedule=schedule,
            )
        )
    report_path = output / "生成说明.txt"
    create_generation_report(report_path, generated_files, template_b_path, count_holidays, signature_scale, schedule)
    return output, generated_files, report_path


def cli() -> int:
    parser = argparse.ArgumentParser(description="根据表C生成表B文件")
    parser.add_argument("--table-c", required=True, help="表C路径，例如 考勤表汇总.xlsx")
    parser.add_argument("--template-b", help="单个表B模板路径(.xlsm/.xlsx)")
    parser.add_argument("--table-bs-dir", help="现有表B目录；未指定模板时会自动取第一个文件做模板")
    parser.add_argument("--output-dir", help="输出目录")
    parser.add_argument("--count-holidays", action="store_true", help="按实际出勤统计 I6 法定假和 L6 周末假，并将 K6/M6 写为 0")
    parser.add_argument("--signature-scale", type=int, default=100, help="签名大小百分比，默认 100，可选 30-200")
    parser.add_argument("--morning-start", default=DEFAULT_MORNING_START, help="上午上班时间，默认 06:00")
    parser.add_argument("--morning-end", default=DEFAULT_MORNING_END, help="上午下班时间，默认 12:00")
    parser.add_argument("--afternoon-start", default=DEFAULT_AFTERNOON_START, help="下午上班时间，默认 14:00")
    parser.add_argument("--afternoon-end", default=DEFAULT_AFTERNOON_END, help="下午下班时间，默认 18:00")
    parser.add_argument("--normal-hours", default=str(DEFAULT_NORMAL_HOURS), help="常规工作小时数，默认 10")
    args = parser.parse_args()

    template_path = choose_template_file(args.template_b, args.table_bs_dir)
    output_dir, generated_files, report_path = run_generate(
        table_c_path=Path(args.table_c),
        template_b_path=template_path,
        output_dir=Path(args.output_dir) if args.output_dir else None,
        count_holidays=args.count_holidays,
        signature_scale=args.signature_scale,
        morning_start=args.morning_start,
        morning_end=args.morning_end,
        afternoon_start=args.afternoon_start,
        afternoon_end=args.afternoon_end,
        normal_hours=args.normal_hours,
    )
    print(f"输出目录: {output_dir}")
    print(f"生成数量: {len(generated_files)}")
    print(f"报告文件: {report_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(cli())
