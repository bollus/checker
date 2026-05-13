from __future__ import annotations

import argparse
import copy
import re
import zipfile
from dataclasses import dataclass
from datetime import date, timedelta
from decimal import Decimal, ROUND_CEILING, ROUND_HALF_UP
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Tuple
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


DAY_START = Decimal("0.25")
LUNCH_START = Decimal("0.5")
LUNCH_END = Decimal("0.583333333333333")
HOURS_PER_DAY = Decimal("10")
WINDOW_PAYABLE_CODES = {"V", "S"}
PAYABLE_CODES = {"V", "S"}
LEAVE_LABELS = {
    "A": "Absent",
    "E": "Emergency Leave",
    "S": "Sick Leave",
    "V": "Vacation",
}
SIGNATURE_FONT_NAME = "Nothing You Could Do"
SIGNATURE_FONT_SIZE = "22"


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


def ensure_signature_style(styles_root: ET.Element) -> Callable[[int], int]:
    fonts = styles_root.find(qn("fonts"))
    cell_xfs = styles_root.find(qn("cellXfs"))
    if fonts is None or cell_xfs is None:
        raise WorkbookError("styles.xml is missing fonts or cellXfs")

    def font_name(font: ET.Element) -> str:
        name = font.find(qn("name"))
        return name.attrib.get("val", "") if name is not None else ""

    signature_font_id: Optional[int] = None
    for index, font in enumerate(list(fonts)):
        if font_name(font).casefold() == SIGNATURE_FONT_NAME.casefold():
            signature_font_id = index
            break

    style_cache: Dict[int, int] = {}

    def style_for(base_style: int) -> int:
        nonlocal signature_font_id
        if base_style in style_cache:
            return style_cache[base_style]

        all_xfs = list(cell_xfs)
        if base_style >= len(all_xfs):
            base_style = 0
        base_xf = all_xfs[base_style]

        if signature_font_id is None:
            all_fonts = list(fonts)
            base_font_id = int(base_xf.attrib.get("fontId", "0"))
            if base_font_id >= len(all_fonts):
                base_font_id = 0
            new_font = copy.deepcopy(all_fonts[base_font_id])
            for old_name in list(new_font.findall(qn("name"))):
                new_font.remove(old_name)
            new_font.insert(0, ET.Element(qn("name"), {"val": SIGNATURE_FONT_NAME}))
            size = new_font.find(qn("sz"))
            if size is None:
                size = ET.Element(qn("sz"))
                new_font.insert(0, size)
            size.attrib["val"] = SIGNATURE_FONT_SIZE
            fonts.append(new_font)
            signature_font_id = len(list(fonts)) - 1
            fonts.attrib["count"] = str(len(list(fonts)))

        new_xf = copy.deepcopy(base_xf)
        new_xf.attrib["fontId"] = str(signature_font_id)
        new_xf.attrib["applyFont"] = "1"
        cell_xfs.append(new_xf)
        new_index = len(list(cell_xfs)) - 1
        cell_xfs.attrib["count"] = str(len(list(cell_xfs)))
        style_cache[base_style] = new_index
        return new_index

    return style_for


def set_signature_cell(sheet: SheetXml, cell_ref: str, value: str, signature_style_for: Callable[[int], int]) -> None:
    set_cell_text(sheet, cell_ref, value)
    cell = sheet.ensure_cell(cell_ref)
    base_style = int(cell.attrib.get("s", "0"))
    cell.attrib["s"] = str(signature_style_for(base_style))


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
            )
        )
        row_num += 1

    return book, summary_sheet_info, day_headers, employees


def build_output_name(employee: SummaryEmployee, suffix: str) -> str:
    name = sanitize_filename_part(employee.name)
    position = sanitize_filename_part(employee.position)
    return f"{employee.no}.{name}-{position}{suffix}"


def day_time_inputs(total_hours: Decimal) -> Tuple[Optional[Decimal], Optional[Decimal], Optional[Decimal], Optional[Decimal]]:
    if total_hours <= 0:
        return None, None, None, None
    if total_hours <= 6:
        return DAY_START, None, None, DAY_START + total_hours / Decimal("24")
    finish = LUNCH_END + (total_hours - Decimal("6")) / Decimal("24")
    return DAY_START, LUNCH_START, LUNCH_END, finish


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
        c_value, d_value, e_value, f_value = day_time_inputs(total_hours)
        set_cell_number(sheet, f"C{row_num}", c_value)
        set_cell_number(sheet, f"D{row_num}", d_value)
        set_cell_number(sheet, f"E{row_num}", e_value)
        set_cell_number(sheet, f"F{row_num}", f_value)
        set_cell_number(sheet, f"N{row_num}", f_value)
        if day_type == "work":
            work_hours = min(total_hours, HOURS_PER_DAY)
            work_ot = max(total_hours - HOURS_PER_DAY, Decimal("0"))
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
        if day_type == "rest":
            set_cell_text(sheet, f"K{row_num}", "Weekend")
        elif day_type == "holiday":
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


def update_overtime_sheet(
    overtime_sheet: Optional[SheetXml],
    employee: SummaryEmployee,
    signature: str,
    signature_style_for: Callable[[int], int],
) -> None:
    if overtime_sheet is None:
        return
    set_cell_text(overtime_sheet, "B5", employee.name)
    set_cell_text(overtime_sheet, "B7", employee.project)
    set_cell_text(overtime_sheet, "I7", employee.position)
    set_signature_cell(overtime_sheet, "E53", signature, signature_style_for)


def choose_template_sheet(sheets: List[WorkbookSheet], sheet_name: str) -> Optional[WorkbookSheet]:
    for sheet in sheets:
        if sheet.display_name == sheet_name:
            return sheet
    return None


def write_employee_workbook(
    template_path: Path,
    employee: SummaryEmployee,
    day_headers: List[Tuple[str, int, str]],
    output_dir: Path,
    count_holidays: bool = False,
) -> Path:
    book = SpreadsheetZip(template_path)
    workbook_root = book.load_xml("xl/workbook.xml")
    styles_root = book.load_xml("xl/styles.xml")
    signature_style_for = ensure_signature_style(styles_root)
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

    if month_days != len(day_headers):
        raise WorkbookError(
            f"模板月份天数({month_days})与表C日期列数量({len(day_headers)})不一致，请确认模板月份是否正确"
        )

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
    set_cell_number(main_sheet, "N8", HOURS_PER_DAY)
    signature = signature_text(employee.name)
    set_signature_cell(main_sheet, "A42", signature, signature_style_for)

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

    payable_total = 0
    work_sum = Decimal("0")
    work_ot_sum = Decimal("0")
    rest_ot_sum = Decimal("0")
    holiday_ot_sum = Decimal("0")
    vacation_days = 0
    sick_days = 0
    emergency_days = 0

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
        )
        payable_total += 1 if payable else 0
        work_sum += work_hours
        work_ot_sum += work_ot
        rest_ot_sum += rest_hours
        holiday_ot_sum += holiday_hours
        if entry_type == "V":
            vacation_days += 1
        elif entry_type == "S":
            sick_days += 1
        elif entry_type == "E":
            emergency_days += 1

    for day_number in range(len(day_headers) + 1, 32):
        row = 9 + day_number
        for col in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "N", "O", "P", "Q", "R", "T"]:
            ref = f"{col}{row}"
            if col == "K":
                set_cell_text(main_sheet, ref, None)
            else:
                set_cell_number(main_sheet, ref, None)

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
        if kind == "holiday" and main_sheet.get_value(f"T{9 + day_number}") not in (None, ""):
            public_payable += 1
        if kind == "rest" and main_sheet.get_value(f"T{9 + day_number}") not in (None, ""):
            rest_payable += 1

    set_cell_number(main_sheet, "A6", len(day_headers))
    set_cell_number(main_sheet, "B6", payable_total if payable_total > 0 else None)
    set_cell_number(
        main_sheet,
        "E6",
        (work_sum / HOURS_PER_DAY).to_integral_value(rounding=ROUND_CEILING) if work_sum > 0 else None,
    )
    set_cell_number(main_sheet, "I6", public_attendance if count_holidays else (public_payable if public_payable > 0 else None))
    set_cell_number(main_sheet, "J6", vacation_days if vacation_days > 0 else None)
    set_cell_number(main_sheet, "K6", 0 if count_holidays else (sick_days if sick_days > 0 else None))
    set_cell_number(main_sheet, "L6", rest_attendance if count_holidays else (rest_payable if rest_payable > 0 else None))
    set_cell_number(main_sheet, "M6", 0 if count_holidays else (emergency_days if emergency_days > 0 else None))
    set_cell_number(main_sheet, "H9", work_ot_sum if work_ot_sum > 0 else None)
    set_cell_number(main_sheet, "I9", rest_ot_sum if rest_ot_sum > 0 else None)
    set_cell_number(main_sheet, "J9", holiday_ot_sum if holiday_ot_sum > 0 else None)

    update_overtime_sheet(overtime_sheet, employee, signature, signature_style_for)
    set_calc_flags(workbook_root)

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / build_output_name(employee, template_path.suffix)
    replacements = {
        "xl/workbook.xml": ET.tostring(workbook_root, encoding="utf-8", xml_declaration=True),
        "xl/styles.xml": ET.tostring(styles_root, encoding="utf-8", xml_declaration=True),
        main_sheet_info.part_name: ET.tostring(main_sheet.root, encoding="utf-8", xml_declaration=True),
    }
    if overtime_sheet_info is not None:
        replacements[overtime_sheet_info.part_name] = ET.tostring(overtime_sheet.root, encoding="utf-8", xml_declaration=True)
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


def create_generation_report(report_path: Path, generated_files: List[Path], template_path: Path, count_holidays: bool) -> None:
    lines = [
        f"模板表B: {template_path}",
        f"生成数量: {len(generated_files)}",
        f"是否统计假期: {'是' if count_holidays else '否'}",
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
) -> Tuple[Path, List[Path], Path]:
    _, _, day_headers, employees = read_summary(table_c_path)
    output = output_dir or next_output_dir(table_c_path)
    generated_files: List[Path] = []
    for employee in employees:
        generated_files.append(write_employee_workbook(template_b_path, employee, day_headers, output, count_holidays=count_holidays))
    report_path = output / "生成说明.txt"
    create_generation_report(report_path, generated_files, template_b_path, count_holidays)
    return output, generated_files, report_path


def cli() -> int:
    parser = argparse.ArgumentParser(description="根据表C生成表B文件")
    parser.add_argument("--table-c", required=True, help="表C路径，例如 考勤表汇总.xlsx")
    parser.add_argument("--template-b", help="单个表B模板路径(.xlsm/.xlsx)")
    parser.add_argument("--table-bs-dir", help="现有表B目录；未指定模板时会自动取第一个文件做模板")
    parser.add_argument("--output-dir", help="输出目录")
    parser.add_argument("--count-holidays", action="store_true", help="按实际出勤统计 I6 法定假和 L6 周末假，并将 K6/M6 写为 0")
    args = parser.parse_args()

    template_path = choose_template_file(args.template_b, args.table_bs_dir)
    output_dir, generated_files, report_path = run_generate(
        table_c_path=Path(args.table_c),
        template_b_path=template_path,
        output_dir=Path(args.output_dir) if args.output_dir else None,
        count_holidays=args.count_holidays,
    )
    print(f"输出目录: {output_dir}")
    print(f"生成数量: {len(generated_files)}")
    print(f"报告文件: {report_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(cli())
