from __future__ import annotations

import argparse
import copy
import json
import re
import sys
import zipfile
from collections.abc import Callable
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
import xml.etree.ElementTree as ET

sys.modules.setdefault("excel_check_tool", sys.modules[__name__])


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)

HIGHLIGHT_RGB = "FFFFFF00"
NUMERIC_TOLERANCE = Decimal("0.000001")
TEXT_FIELDS = {
    "F": ("A3", "姓名"),
    "J": ("C3", "岗位"),
}
NUMERIC_FIELDS = {
    "N": ("B6", "应支付天数"),
    "W": ("H9", "正常工作日加班"),
    "X": ("I9", "周末加班"),
    "Y": ("J9", "法定假日加班"),
}
POSITION_OPTIONAL_TOKENS = {
    "construction",
    "mechanical",
    "civil",
    "ei",
    "piping",
    "inspector",
    "supervisor",
    "officer",
}
POSITION_ALIASES_PATH = Path(__file__).with_name("position_aliases.json")
CHECK_TEMPLATES_PATH = Path(__file__).with_name("check_templates.json")
COMPARE_TYPE_TEXT = "text"
COMPARE_TYPE_NUMBER = "number"
COMPARE_TYPE_POSITION = "position"
DEFAULT_CHECK_TEMPLATE_NAME = "默认模板"


def qn(name: str, ns: str = MAIN_NS) -> str:
    return f"{{{ns}}}{name}"


def split_cell_ref(cell_ref: str) -> Tuple[str, int]:
    match = re.fullmatch(r"([A-Z]+)(\d+)", cell_ref)
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    return match.group(1), int(match.group(2))


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


def normalize_text(value: Optional[str]) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").strip()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"\s+\(", "(", text)
    text = re.sub(r"\(\s+", "(", text)
    text = re.sub(r"\s+\)", ")", text)
    text = re.sub(r"\s*-\s*", "-", text)
    text = re.sub(r"\s*/\s*", "/", text)
    return text.casefold()


def to_display(value: Optional[str]) -> str:
    if value is None:
        return ""
    text = str(value).replace("\n", " ").replace("\r", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_number(value: Optional[str]) -> Decimal:
    if value is None:
        return Decimal("0")
    text = str(value).strip().replace(",", "")
    if text == "":
        return Decimal("0")
    try:
        return Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(f"Cannot parse numeric value: {value!r}") from exc


def decimal_to_text(value: Decimal) -> str:
    normalized = value.normalize()
    text = format(normalized, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


def default_check_template() -> CheckTemplate:
    return CheckTemplate(
        name=DEFAULT_CHECK_TEMPLATE_NAME,
        number_column="A",
        start_row=7,
        rules=[
            CheckRule("姓名", "F7-Fn", "A3", COMPARE_TYPE_TEXT),
            CheckRule("岗位", "J7-Jn", "C3", COMPARE_TYPE_POSITION),
            CheckRule("应支付天数", "N7-Nn", "B6", COMPARE_TYPE_NUMBER),
            CheckRule("正常工作日加班", "W7-Wn", "H9", COMPARE_TYPE_NUMBER),
            CheckRule("周末加班", "X7-Xn", "I9", COMPARE_TYPE_NUMBER),
            CheckRule("法定假日加班", "Y7-Yn", "J9", COMPARE_TYPE_NUMBER),
        ],
    )


def template_to_dict(template: CheckTemplate) -> Dict[str, object]:
    return {
        "name": template.name,
        "number_column": template.number_column,
        "start_row": template.start_row,
        "rules": [
            {
                "field_name": rule.field_name,
                "main_range": rule.main_range,
                "table_b_cell": rule.table_b_cell,
                "compare_type": rule.compare_type,
            }
            for rule in template.rules
        ],
    }


def check_rule_from_dict(data: Dict[str, object]) -> CheckRule:
    field_name = str(data.get("field_name", "")).strip()
    main_range = str(data.get("main_range", "")).strip().upper()
    table_b_cell = str(data.get("table_b_cell", "")).strip().upper()
    compare_type = str(data.get("compare_type", COMPARE_TYPE_TEXT)).strip().lower()
    if compare_type not in {COMPARE_TYPE_TEXT, COMPARE_TYPE_NUMBER, COMPARE_TYPE_POSITION}:
        compare_type = COMPARE_TYPE_TEXT
    if not field_name or not main_range or not table_b_cell:
        raise WorkbookError("核对模板规则缺少字段名、主表范围或考勤表坐标")
    return CheckRule(
        field_name=field_name,
        main_range=main_range,
        table_b_cell=table_b_cell,
        compare_type=compare_type,
    )


def check_template_from_dict(data: Dict[str, object]) -> CheckTemplate:
    name = str(data.get("name", "")).strip() or DEFAULT_CHECK_TEMPLATE_NAME
    number_column = str(data.get("number_column", "A")).strip().upper() or "A"
    start_row = int(data.get("start_row", 7))
    rules_data = data.get("rules", [])
    if not isinstance(rules_data, list) or not rules_data:
        raise WorkbookError(f"核对模板 {name} 没有规则")
    rules = [check_rule_from_dict(item) for item in rules_data if isinstance(item, dict)]
    return CheckTemplate(name=name, number_column=number_column, start_row=start_row, rules=rules)


def load_check_templates() -> List[CheckTemplate]:
    templates: List[CheckTemplate] = []
    if CHECK_TEMPLATES_PATH.exists():
        try:
            raw = json.loads(CHECK_TEMPLATES_PATH.read_text(encoding="utf-8"))
        except Exception:
            raw = []
        if isinstance(raw, list):
            for item in raw:
                if not isinstance(item, dict):
                    continue
                try:
                    templates.append(check_template_from_dict(item))
                except Exception:
                    continue
    if not any(template.name == DEFAULT_CHECK_TEMPLATE_NAME for template in templates):
        templates.insert(0, default_check_template())
    return templates


def save_check_templates(templates: List[CheckTemplate]) -> None:
    ordered = sorted(templates, key=lambda item: (item.name != DEFAULT_CHECK_TEMPLATE_NAME, item.name.casefold()))
    CHECK_TEMPLATES_PATH.write_text(
        json.dumps([template_to_dict(template) for template in ordered], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def get_check_template(template_name: Optional[str]) -> CheckTemplate:
    templates = load_check_templates()
    if template_name:
        for template in templates:
            if template.name == template_name:
                return template
        raise WorkbookError(f"未找到核对模板: {template_name}")
    return templates[0]


def parse_main_range(main_range: str) -> Tuple[str, int]:
    match = re.fullmatch(r"([A-Z]+)(\d+)-([A-Z]+)N", main_range.strip().upper())
    if not match:
        raise WorkbookError(f"主表范围格式不正确: {main_range}，示例应为 F7-Fn")
    start_column = match.group(1)
    end_column = match.group(3)
    if start_column != end_column:
        raise WorkbookError(f"当前仅支持同一列的范围模板: {main_range}")
    return start_column, int(match.group(2))


def infer_data_end_row(sheet: SheetXml, start_row: int) -> int:
    current = start_row
    last_seen = start_row
    max_row = max(sheet.rows) if sheet.rows else start_row
    while current <= max_row:
        anchor_value = sheet.get_value(f"A{current}")
        if anchor_value in (None, ""):
            break
        try:
            normalize_number(anchor_value)
        except Exception:
            break
        last_seen = current
        current += 1
    return last_seen


def expand_n_row(row_text: str, sheet: SheetXml, start_row: int) -> int:
    return infer_data_end_row(sheet, start_row) if row_text == "N" else int(row_text)


def parse_table_b_expression_tokens(expression: str) -> List[str]:
    text = expression.strip().upper()
    if not text.startswith("SUM(") or not text.endswith(")"):
        raise WorkbookError(f"考勤表表达式格式不正确: {expression}")
    inner = text[4:-1].strip()
    if not inner:
        raise WorkbookError(f"SUM 表达式不能为空: {expression}")
    return [part.strip() for part in inner.split(",") if part.strip()]


def validate_table_b_expression(expression: str) -> None:
    text = expression.strip().upper()
    if text.startswith("SUM("):
        for token in parse_table_b_expression_tokens(text):
            if ":" in token:
                start_ref, end_ref = token.split(":", 1)
                start_match = re.fullmatch(r"([A-Z]+)(\d+)", start_ref)
                end_match = re.fullmatch(r"([A-Z]+)(\d+|N)", end_ref)
                if not start_match or not end_match:
                    raise WorkbookError(f"SUM 范围格式不正确: {token}")
                if col_to_num(start_match.group(1)) > col_to_num(end_match.group(1)):
                    raise WorkbookError(f"SUM 范围起始列不能大于结束列: {token}")
                if int(start_match.group(2)) > (999999 if end_match.group(2) == 'N' else int(end_match.group(2))):
                    raise WorkbookError(f"SUM 范围起始行不能大于结束行: {token}")
            else:
                split_cell_ref(token)
        return
    split_cell_ref(text)


def iter_range_cells(start_ref: str, end_ref: str, sheet: SheetXml) -> List[str]:
    start_col, start_row = split_cell_ref(start_ref)
    end_match = re.fullmatch(r"([A-Z]+)(\d+|N)", end_ref)
    if end_match is None:
        raise WorkbookError(f"范围终点格式不正确: {end_ref}")
    end_col = end_match.group(1)
    end_row = expand_n_row(end_match.group(2), sheet, start_row)
    refs: List[str] = []
    for col_num in range(col_to_num(start_col), col_to_num(end_col) + 1):
        col = num_to_col(col_num)
        for row_num in range(start_row, end_row + 1):
            refs.append(f"{col}{row_num}")
    return refs


def resolve_table_b_value(sheet: SheetXml, expression: str) -> Optional[str]:
    text = expression.strip().upper()
    if not text.startswith("SUM("):
        return sheet.get_value(text)

    total = Decimal("0")
    for token in parse_table_b_expression_tokens(text):
        if ":" in token:
            start_ref, end_ref = token.split(":", 1)
            refs = iter_range_cells(start_ref, end_ref, sheet)
        else:
            refs = [token]
        for ref in refs:
            total += normalize_number(sheet.get_value(ref))
    return decimal_to_text(total)


def parse_check_template(template: CheckTemplate) -> List[ParsedCheckRule]:
    parsed_rules: List[ParsedCheckRule] = []
    for rule in template.rules:
        main_column, main_start_row = parse_main_range(rule.main_range)
        validate_table_b_expression(rule.table_b_cell)
        parsed_rules.append(
            ParsedCheckRule(
                field_name=rule.field_name,
                main_range=rule.main_range,
                table_b_cell=rule.table_b_cell.strip().upper(),
                compare_type=rule.compare_type,
                main_column=main_column,
                main_start_row=main_start_row,
            )
        )
    return parsed_rules


_POSITION_ALIASES_CACHE: Optional[Dict[str, str]] = None


def load_position_aliases() -> Dict[str, str]:
    global _POSITION_ALIASES_CACHE
    if _POSITION_ALIASES_CACHE is not None:
        return _POSITION_ALIASES_CACHE

    aliases: Dict[str, str] = {}
    if POSITION_ALIASES_PATH.exists():
        try:
            data = json.loads(POSITION_ALIASES_PATH.read_text(encoding="utf-8"))
        except Exception:
            data = {}
        if isinstance(data, dict):
            for key, value in data.items():
                normalized_key = normalize_text(str(key))
                normalized_value = normalize_text(str(value))
                if normalized_key and normalized_value:
                    aliases[normalized_key] = normalized_value
    _POSITION_ALIASES_CACHE = aliases
    return aliases


def canonical_position_tokens(value: Optional[str]) -> Tuple[str, ...]:
    text = normalize_text(value)
    if not text:
        return ()
    text = load_position_aliases().get(text, text)

    replacements = [
        (r"\be\s*&\s*i\b", " ei "),
        (r"\be\s+i\b", " ei "),
        (r"\bqci\b", " qc inspector "),
        (r"\bmec\b", " mechanical "),
        (r"\badmin\b", " administrator "),
        (r"\bconstrucion\b", " construction "),
        (r"\bcontroler\b", " controller "),
        (r"\bkeepr\b", " keeper "),
        (r"\bscaffolder\b", " scaffolding "),
    ]
    for pattern, replacement in replacements:
        text = re.sub(pattern, replacement, text)

    text = re.sub(r"[().,/]", " ", text)
    text = re.sub(r"\s*-\s*", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return ()

    tokens = sorted(set(text.split()))
    return tuple(tokens)


def positions_mean_same(left: Optional[str], right: Optional[str]) -> bool:
    left_tokens = canonical_position_tokens(left)
    right_tokens = canonical_position_tokens(right)
    if left_tokens == right_tokens:
        return True
    if not left_tokens or not right_tokens:
        return left_tokens == right_tokens

    left_set = set(left_tokens)
    right_set = set(right_tokens)
    if left_set == right_set:
        return True

    smaller, larger = sorted((left_set, right_set), key=len)
    extra_tokens = larger - smaller
    return len(smaller) >= 2 and smaller.issubset(larger) and extra_tokens.issubset(POSITION_OPTIONAL_TOKENS)


class WorkbookError(Exception):
    pass


@dataclass
class CheckRule:
    field_name: str
    main_range: str
    table_b_cell: str
    compare_type: str


@dataclass
class ParsedCheckRule:
    field_name: str
    main_range: str
    table_b_cell: str
    compare_type: str
    main_column: str
    main_start_row: int


@dataclass
class CheckTemplate:
    name: str
    number_column: str
    start_row: int
    rules: List[CheckRule]


class SheetXml:
    def __init__(self, root: ET.Element, shared_strings: List[str]):
        self.root = root
        self.shared_strings = shared_strings
        self.sheet_data = root.find(qn("sheetData"))
        if self.sheet_data is None:
            raise WorkbookError("Worksheet is missing sheetData")
        self.rows: Dict[int, ET.Element] = {}
        self.cells: Dict[str, ET.Element] = {}
        for row in self.sheet_data.findall(qn("row")):
            row_num = int(row.attrib["r"])
            self.rows[row_num] = row
            for cell in row.findall(qn("c")):
                self.cells[cell.attrib["r"]] = cell

    def get_cell(self, cell_ref: str) -> Optional[ET.Element]:
        return self.cells.get(cell_ref)

    def get_value(self, cell_ref: str) -> Optional[str]:
        cell = self.cells.get(cell_ref)
        if cell is None:
            return None
        cell_type = cell.attrib.get("t")
        value_node = cell.find(qn("v"))
        if cell_type == "inlineStr":
            inline_node = cell.find(qn("is"))
            if inline_node is None:
                return None
            return "".join(text_node.text or "" for text_node in inline_node.iter(qn("t")))
        if value_node is None:
            return None
        value = value_node.text
        if value is None:
            return None
        if cell_type == "s":
            return self.shared_strings[int(value)]
        if cell_type == "b":
            return "TRUE" if value == "1" else "FALSE"
        return value

    def ensure_cell(self, cell_ref: str) -> ET.Element:
        existing = self.cells.get(cell_ref)
        if existing is not None:
            return existing

        col, row_num = split_cell_ref(cell_ref)
        row = self.rows.get(row_num)
        if row is None:
            row = ET.Element(qn("row"), {"r": str(row_num)})
            inserted = False
            for index, current_row in enumerate(list(self.sheet_data)):
                current_num = int(current_row.attrib["r"])
                if current_num > row_num:
                    self.sheet_data.insert(index, row)
                    inserted = True
                    break
            if not inserted:
                self.sheet_data.append(row)
            self.rows[row_num] = row

        cell = ET.Element(qn("c"), {"r": cell_ref})
        base_style = self.guess_style(cell_ref)
        if base_style is not None:
            cell.attrib["s"] = str(base_style)

        row_cells = list(row.findall(qn("c")))
        inserted = False
        for index, current_cell in enumerate(row_cells):
            current_col, _ = split_cell_ref(current_cell.attrib["r"])
            if current_col > col:
                row.insert(index, cell)
                inserted = True
                break
        if not inserted:
            row.append(cell)

        self.cells[cell_ref] = cell
        return cell

    def guess_style(self, cell_ref: str) -> Optional[int]:
        col, row_num = split_cell_ref(cell_ref)
        candidates = [
            f"{col}{row_num - 1}",
            f"{col}{row_num + 1}",
        ]
        for candidate in candidates:
            cell = self.cells.get(candidate)
            if cell is not None and "s" in cell.attrib:
                return int(cell.attrib["s"])

        same_row = self.rows.get(row_num)
        if same_row is not None:
            for cell in same_row.findall(qn("c")):
                if "s" in cell.attrib:
                    return int(cell.attrib["s"])
        return 0


class SpreadsheetZip:
    def __init__(self, path: Path):
        self.path = path
        self.raw_entries: Dict[str, bytes] = {}
        self.zip_infos: Dict[str, zipfile.ZipInfo] = {}
        with zipfile.ZipFile(path) as zf:
            for info in zf.infolist():
                self.zip_infos[info.filename] = info
                if not info.is_dir():
                    self.raw_entries[info.filename] = zf.read(info.filename)

    def load_xml(self, filename: str) -> ET.Element:
        try:
            return ET.fromstring(self.raw_entries[filename])
        except KeyError as exc:
            raise WorkbookError(f"Missing XML part: {filename}") from exc

    def save(self, output_path: Path, replacements: Dict[str, bytes]) -> None:
        with zipfile.ZipFile(output_path, "w") as out_zip:
            written = set()
            for filename, info in self.zip_infos.items():
                if info.is_dir():
                    continue
                data = replacements.get(filename, self.raw_entries[filename])
                new_info = zipfile.ZipInfo(filename)
                new_info.date_time = info.date_time
                new_info.compress_type = info.compress_type
                new_info.comment = info.comment
                new_info.create_system = info.create_system
                new_info.external_attr = info.external_attr
                new_info.extra = info.extra
                new_info.flag_bits = info.flag_bits
                new_info.internal_attr = info.internal_attr
                out_zip.writestr(new_info, data)
                written.add(filename)
            for filename, data in replacements.items():
                if filename not in written:
                    out_zip.writestr(filename, data)


@dataclass
class WorkbookSheet:
    part_name: str
    display_name: str
    sheet: SheetXml


def load_shared_strings(book: SpreadsheetZip) -> List[str]:
    if "xl/sharedStrings.xml" not in book.raw_entries:
        return []
    root = book.load_xml("xl/sharedStrings.xml")
    items: List[str] = []
    for entry in root.findall(qn("si")):
        text = "".join(node.text or "" for node in entry.iter(qn("t")))
        items.append(text)
    return items


def workbook_sheets(book: SpreadsheetZip) -> List[WorkbookSheet]:
    workbook = book.load_xml("xl/workbook.xml")
    rels = book.load_xml("xl/_rels/workbook.xml.rels")
    shared_strings = load_shared_strings(book)

    rel_targets: Dict[str, str] = {}
    for relation in rels.findall(qn("Relationship", PKG_REL_NS)):
        target = relation.attrib["Target"]
        if not target.startswith("xl/"):
            target = f"xl/{target}"
        rel_targets[relation.attrib["Id"]] = target

    sheets: List[WorkbookSheet] = []
    sheets_node = workbook.find(qn("sheets"))
    if sheets_node is None:
        raise WorkbookError("Workbook does not contain any sheets")
    for sheet_node in sheets_node.findall(qn("sheet")):
        rel_id = sheet_node.attrib[f"{{{REL_NS}}}id"]
        target = rel_targets.get(rel_id)
        if target is None or not target.startswith("xl/worksheets/"):
            continue
        sheet_root = book.load_xml(target)
        sheets.append(
            WorkbookSheet(
                part_name=target,
                display_name=sheet_node.attrib["name"],
                sheet=SheetXml(sheet_root, shared_strings),
            )
        )
    if not sheets:
        raise WorkbookError("Workbook does not contain any worksheet parts")
    return sheets


def choose_timesheet_sheet(sheets: List[WorkbookSheet]) -> WorkbookSheet:
    required_refs = ("A3", "B6", "C3", "H9", "I9", "J9")
    ranked = sorted(
        sheets,
        key=lambda item: sum(1 for ref in required_refs if to_display(item.sheet.get_value(ref))),
        reverse=True,
    )
    return ranked[0]


def next_output_path(table_a_path: Path) -> Path:
    return table_a_path.with_name(f"{table_a_path.stem}_核对结果{table_a_path.suffix}")


def next_report_path(output_path: Path) -> Path:
    return output_path.with_name(f"{output_path.stem}_核对报告.txt")


def build_table_b_index(folder: Path) -> Tuple[Dict[int, Path], List[str]]:
    files: Dict[int, List[Path]] = {}
    warnings: List[str] = []
    for path in sorted(folder.iterdir()):
        if path.is_dir():
            continue
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            continue
        match = re.match(r"^\s*(\d+)\.(.+)$", path.stem)
        if not match:
            warnings.append(f"忽略未匹配编号规则的文件: {path.name}")
            continue
        files.setdefault(int(match.group(1)), []).append(path)

    selected: Dict[int, Path] = {}
    for number, candidates in files.items():
        ordered = sorted(
            candidates,
            key=lambda item: (
                "副本" in item.stem or "copy" in item.stem.casefold(),
                item.name.casefold(),
            ),
        )
        selected[number] = ordered[0]
        if len(candidates) > 1:
            joined = ", ".join(candidate.name for candidate in candidates)
            warnings.append(f"No.{number} 存在多个候选文件，已使用 {ordered[0].name}: {joined}")
    return selected, warnings


def add_highlight_style(styles_root: ET.Element) -> Callable[[int], int]:
    fills = styles_root.find(qn("fills"))
    cell_xfs = styles_root.find(qn("cellXfs"))
    if fills is None or cell_xfs is None:
        raise WorkbookError("styles.xml is missing fills or cellXfs")

    highlight_fill_id: Optional[int] = None
    for index, fill in enumerate(list(fills)):
        pattern_fill = fill.find(qn("patternFill"))
        if pattern_fill is None:
            continue
        fg_color = pattern_fill.find(qn("fgColor"))
        if fg_color is not None and fg_color.attrib.get("rgb") == HIGHLIGHT_RGB:
            highlight_fill_id = index
            break

    if highlight_fill_id is None:
        fill = ET.Element(qn("fill"))
        pattern_fill = ET.SubElement(fill, qn("patternFill"), {"patternType": "solid"})
        ET.SubElement(pattern_fill, qn("fgColor"), {"rgb": HIGHLIGHT_RGB})
        ET.SubElement(pattern_fill, qn("bgColor"), {"indexed": "64"})
        fills.append(fill)
        highlight_fill_id = len(list(fills)) - 1
        fills.attrib["count"] = str(len(list(fills)))

    style_cache: Dict[int, int] = {}

    def style_for(base_style: int) -> int:
        if base_style in style_cache:
            return style_cache[base_style]

        all_xfs = list(cell_xfs)
        if base_style >= len(all_xfs):
            base_style = 0
        new_xf = copy.deepcopy(all_xfs[base_style])
        new_xf.attrib["fillId"] = str(highlight_fill_id)
        new_xf.attrib["applyFill"] = "1"
        cell_xfs.append(new_xf)
        new_index = len(list(cell_xfs)) - 1
        cell_xfs.attrib["count"] = str(len(list(cell_xfs)))
        style_cache[base_style] = new_index
        return new_index

    return_value = style_for
    return return_value


def highlight_cell(sheet: SheetXml, cell_ref: str, highlight_style_for) -> None:
    cell = sheet.ensure_cell(cell_ref)
    base_style = int(cell.attrib.get("s", "0"))
    cell.attrib["s"] = str(highlight_style_for(base_style))


@dataclass
class Mismatch:
    row_num: int
    table_a_cell: str
    field_name: str
    table_a_value: str
    table_b_value: str
    table_b_file: str


def compare_row(
    offset: int,
    table_a_sheet: SheetXml,
    table_b_sheet: Optional[SheetXml],
    table_b_file: Optional[Path],
    parsed_rules: List[ParsedCheckRule],
    display_row_num: int,
) -> List[Mismatch]:
    mismatches: List[Mismatch] = []
    file_name = table_b_file.name if table_b_file else "未找到匹配考勤表"

    for rule in parsed_rules:
        table_a_cell = f"{rule.main_column}{rule.main_start_row + offset}"
        left = table_a_sheet.get_value(table_a_cell)
        right = None if table_b_sheet is None else resolve_table_b_value(table_b_sheet, rule.table_b_cell)
        if table_b_sheet is None:
            matches = False
        else:
            if rule.compare_type == COMPARE_TYPE_NUMBER:
                matches = abs(normalize_number(left) - normalize_number(right)) <= NUMERIC_TOLERANCE
            elif rule.compare_type == COMPARE_TYPE_POSITION:
                matches = positions_mean_same(left, right)
            else:
                matches = normalize_text(left) == normalize_text(right)
        if not matches:
            mismatches.append(
                Mismatch(
                    row_num=display_row_num,
                    table_a_cell=table_a_cell,
                    field_name=rule.field_name,
                    table_a_value=to_display(left),
                    table_b_value=to_display(right),
                    table_b_file=file_name,
                )
            )
    return mismatches


def locate_data_offsets(sheet: SheetXml, template: CheckTemplate, parsed_rules: List[ParsedCheckRule]) -> List[int]:
    offsets: List[int] = []
    current_offset = 0
    while True:
        refs = [f"{template.number_column}{template.start_row + current_offset}"]
        refs.extend(f"{rule.main_column}{rule.main_start_row + current_offset}" for rule in parsed_rules)
        if all(sheet.get_value(ref) in (None, "") for ref in refs):
            break
        offsets.append(current_offset)
        current_offset += 1
    return offsets


def create_report(
    report_path: Path,
    table_a_path: Path,
    output_path: Path,
    warnings: Iterable[str],
    mismatches: List[Mismatch],
) -> None:
    lines: List[str] = []
    lines.append(f"主表: {table_a_path}")
    lines.append(f"结果文件: {output_path}")
    lines.append(f"不一致数量: {len(mismatches)}")
    lines.append("")
    warning_items = list(warnings)
    if warning_items:
        lines.append("提示:")
        for warning in warning_items:
            lines.append(f"- {warning}")
        lines.append("")
    if mismatches:
        lines.append("明细:")
        for mismatch in mismatches:
            lines.append(
                f"- 第 {mismatch.row_num} 行 {mismatch.table_a_cell}({mismatch.field_name}) | "
                f"主表='{mismatch.table_a_value}' | 考勤表='{mismatch.table_b_value}' | 文件={mismatch.table_b_file}"
            )
    else:
        lines.append("未发现不一致。")
    report_path.write_text("\n".join(lines), encoding="utf-8")


def run_check(
    table_a_path: Path,
    table_bs_folder: Path,
    output_path: Optional[Path] = None,
    template_name: Optional[str] = None,
) -> Tuple[Path, Path, List[Mismatch], List[str]]:
    if table_a_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise WorkbookError("主表只支持 .xlsx 或 .xlsm")
    if not table_bs_folder.is_dir():
        raise WorkbookError("考勤表目录不存在")

    output = output_path or next_output_path(table_a_path)
    report = next_report_path(output)

    table_a_book = SpreadsheetZip(table_a_path)
    table_a_sheets = workbook_sheets(table_a_book)
    table_a_sheet_info = table_a_sheets[0]
    table_a_sheet = table_a_sheet_info.sheet

    check_template = get_check_template(template_name)
    parsed_rules = parse_check_template(check_template)
    table_b_files, warnings = build_table_b_index(table_bs_folder)
    data_offsets = locate_data_offsets(table_a_sheet, check_template, parsed_rules)

    styles_root = table_a_book.load_xml("xl/styles.xml")
    highlight_style_for = add_highlight_style(styles_root)

    mismatches: List[Mismatch] = []
    workbook_cache: Dict[Path, WorkbookSheet] = {}

    for offset in data_offsets:
        display_row_num = check_template.start_row + offset
        row_number_raw = table_a_sheet.get_value(f"{check_template.number_column}{display_row_num}")
        if row_number_raw is None or str(row_number_raw).strip() == "":
            continue
        try:
            file_number = int(Decimal(str(row_number_raw)))
        except InvalidOperation as exc:
            raise WorkbookError(
                f"主表 {check_template.number_column}{display_row_num} 不是有效编号: {row_number_raw!r}"
            ) from exc

        table_b_path = table_b_files.get(file_number)
        table_b_sheet: Optional[SheetXml] = None
        if table_b_path is None:
            warnings.append(f"No.{file_number} 未找到匹配考勤表文件")
        else:
            if table_b_path not in workbook_cache:
                table_b_book = SpreadsheetZip(table_b_path)
                table_b_sheet_info = choose_timesheet_sheet(workbook_sheets(table_b_book))
                workbook_cache[table_b_path] = table_b_sheet_info
            table_b_sheet = workbook_cache[table_b_path].sheet

        row_mismatches = compare_row(offset, table_a_sheet, table_b_sheet, table_b_path, parsed_rules, display_row_num)
        mismatches.extend(row_mismatches)
        for mismatch in row_mismatches:
            highlight_cell(table_a_sheet, mismatch.table_a_cell, highlight_style_for)

    replacements = {
        "xl/styles.xml": ET.tostring(styles_root, encoding="utf-8", xml_declaration=True),
        table_a_sheet_info.part_name: ET.tostring(table_a_sheet.root, encoding="utf-8", xml_declaration=True),
    }
    table_a_book.save(output, replacements)
    create_report(report, table_a_path, output, warnings, mismatches)
    return output, report, mismatches, warnings


def cli(argv: List[str]) -> int:
    parser = argparse.ArgumentParser(description="工资表与考勤表核对工具")
    parser.add_argument("--table-a", help="主表路径 (.xlsx/.xlsm)")
    parser.add_argument("--table-bs", help="考勤表目录路径")
    parser.add_argument("--output", help="输出文件路径")
    parser.add_argument("--template-name", help="核对模板名称")
    args = parser.parse_args(argv)

    if not args.table_a and not args.table_bs and not args.output:
        launch_gui()
        return 0

    if not args.table_a or not args.table_bs:
        parser.error("--table-a 和 --table-bs 必须同时提供，或直接无参数启动图形界面")

    output, report, mismatches, warnings = run_check(
        table_a_path=Path(args.table_a),
        table_bs_folder=Path(args.table_bs),
        output_path=Path(args.output) if args.output else None,
        template_name=args.template_name,
    )
    print(f"结果文件: {output}")
    print(f"报告文件: {report}")
    print(f"不一致数量: {len(mismatches)}")
    if warnings:
        print("提示:")
        for warning in warnings:
            print(f"- {warning}")
    return 0


def launch_gui() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("表格工具")
    root.geometry("980x700")

    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True)

    check_tab = ttk.Frame(notebook, padding=16)
    generate_tab = ttk.Frame(notebook, padding=16)
    template_tab = ttk.Frame(notebook, padding=16)
    notebook.add(check_tab, text="工资表核对")
    notebook.add(generate_tab, text="生成考勤表")
    notebook.add(template_tab, text="核对模板")

    for tab in (check_tab, generate_tab, template_tab):
        tab.columnconfigure(1, weight=1)
    check_tab.rowconfigure(7, weight=1)
    generate_tab.rowconfigure(8, weight=1)
    template_tab.columnconfigure(1, weight=1)
    template_tab.rowconfigure(0, weight=1)

    def clear_log(widget: tk.Text) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.configure(state="disabled")

    def append_log(widget: tk.Text, text: str) -> None:
        widget.configure(state="normal")
        widget.insert("end", text + "\n")
        widget.see("end")
        widget.configure(state="disabled")

    templates_state = load_check_templates()

    def template_names() -> List[str]:
        return [template.name for template in templates_state]

    def refresh_template_choices(selected_name: Optional[str] = None) -> None:
        names = template_names()
        check_template_combo["values"] = names
        if selected_name in names:
            check_template_var.set(selected_name)
        elif check_template_var.get() not in names:
            check_template_var.set(names[0] if names else "")

        template_listbox.delete(0, "end")
        for name in names:
            template_listbox.insert("end", name)
        if selected_name in names:
            index = names.index(selected_name)
            template_listbox.selection_clear(0, "end")
            template_listbox.selection_set(index)
            template_listbox.activate(index)
        elif names:
            template_listbox.selection_clear(0, "end")
            template_listbox.selection_set(0)
            template_listbox.activate(0)

    table_a_var = tk.StringVar()
    table_bs_var = tk.StringVar()
    output_var = tk.StringVar()
    check_template_var = tk.StringVar(value=template_names()[0] if templates_state else "")

    def choose_table_a() -> None:
        path = filedialog.askopenfilename(
            title="选择主表",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm")],
        )
        if not path:
            return
        table_a_var.set(path)
        if not output_var.get():
            output_var.set(str(next_output_path(Path(path))))

    def choose_table_bs() -> None:
        path = filedialog.askdirectory(title="选择考勤表目录")
        if path:
            table_bs_var.set(path)

    def choose_output() -> None:
        path = filedialog.asksaveasfilename(
            title="选择输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm")],
        )
        if path:
            output_var.set(path)

    def start_check() -> None:
        table_a_text = table_a_var.get().strip()
        table_bs_text = table_bs_var.get().strip()
        output_text = output_var.get().strip()
        template_name = check_template_var.get().strip() or None
        if not table_a_text:
            messagebox.showerror("缺少主表", "请选择主表文件。")
            return
        if not table_bs_text:
            messagebox.showerror("缺少目录", "请选择考勤表目录。")
            return
        check_button.configure(state="disabled")
        clear_log(check_log_text)
        append_log(check_log_text, f"开始核对，模板: {template_name or DEFAULT_CHECK_TEMPLATE_NAME}")
        try:
            output, report, mismatches, warnings = run_check(
                table_a_path=Path(table_a_text),
                table_bs_folder=Path(table_bs_text),
                output_path=Path(output_text) if output_text else None,
                template_name=template_name,
            )
        except Exception as exc:
            append_log(check_log_text, f"失败: {exc}")
            messagebox.showerror("执行失败", str(exc))
        else:
            append_log(check_log_text, f"结果文件: {output}")
            append_log(check_log_text, f"报告文件: {report}")
            append_log(check_log_text, f"不一致数量: {len(mismatches)}")
            for warning in warnings:
                append_log(check_log_text, f"提示: {warning}")
            messagebox.showinfo(
                "核对完成",
                f"已完成核对。\n\n结果文件:\n{output}\n\n报告文件:\n{report}\n\n不一致数量: {len(mismatches)}",
            )
        finally:
            check_button.configure(state="normal")

    ttk.Label(check_tab, text="主表").grid(row=0, column=0, sticky="w", pady=(0, 8))
    ttk.Entry(check_tab, textvariable=table_a_var).grid(row=0, column=1, sticky="ew", padx=(8, 8), pady=(0, 8))
    ttk.Button(check_tab, text="选择文件", command=choose_table_a).grid(row=0, column=2, sticky="ew", pady=(0, 8))

    ttk.Label(check_tab, text="考勤表目录").grid(row=1, column=0, sticky="w", pady=(0, 8))
    ttk.Entry(check_tab, textvariable=table_bs_var).grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=(0, 8))
    ttk.Button(check_tab, text="选择目录", command=choose_table_bs).grid(row=1, column=2, sticky="ew", pady=(0, 8))

    ttk.Label(check_tab, text="输出文件").grid(row=2, column=0, sticky="w", pady=(0, 8))
    ttk.Entry(check_tab, textvariable=output_var).grid(row=2, column=1, sticky="ew", padx=(8, 8), pady=(0, 8))
    ttk.Button(check_tab, text="选择位置", command=choose_output).grid(row=2, column=2, sticky="ew", pady=(0, 8))

    ttk.Label(check_tab, text="核对模板").grid(row=3, column=0, sticky="w", pady=(0, 8))
    check_template_combo = ttk.Combobox(check_tab, textvariable=check_template_var, state="readonly")
    check_template_combo.grid(row=3, column=1, sticky="ew", padx=(8, 8), pady=(0, 8))
    ttk.Button(check_tab, text="去模板页编辑", command=lambda: notebook.select(template_tab)).grid(
        row=3, column=2, sticky="ew", pady=(0, 8)
    )

    check_button = ttk.Button(check_tab, text="开始核对", command=start_check)
    check_button.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(4, 8))

    ttk.Label(check_tab, text="日志").grid(row=6, column=0, columnspan=3, sticky="w", pady=(0, 6))
    check_log_text = tk.Text(check_tab, height=16, wrap="word", state="disabled")
    check_log_text.grid(row=7, column=0, columnspan=3, sticky="nsew")

    table_c_var = tk.StringVar()
    template_b_var = tk.StringVar()
    generate_table_bs_var = tk.StringVar()
    generate_output_var = tk.StringVar()
    count_holidays_var = tk.BooleanVar(value=False)

    def choose_table_c() -> None:
        path = filedialog.askopenfilename(
            title="选择汇总表",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm")],
        )
        if not path:
            return
        table_c_var.set(path)
        if not generate_output_var.get():
            generate_output_var.set(str(Path(path).with_name(f"{Path(path).stem}_生成考勤表目录")))

    def choose_template_b() -> None:
        path = filedialog.askopenfilename(
            title="选择考勤表模板",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm")],
        )
        if path:
            template_b_var.set(path)

    def choose_generate_bs_dir() -> None:
        path = filedialog.askdirectory(title="选择现有考勤表目录")
        if path:
            generate_table_bs_var.set(path)

    def choose_generate_output_dir() -> None:
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            generate_output_var.set(path)

    def start_generate() -> None:
        from generate_table_bs import choose_template_file, run_generate

        table_c_text = table_c_var.get().strip()
        template_b_text = template_b_var.get().strip()
        table_bs_dir_text = generate_table_bs_var.get().strip()
        output_dir_text = generate_output_var.get().strip()

        if not table_c_text:
            messagebox.showerror("缺少汇总表", "请选择汇总表文件。")
            return
        if not template_b_text and not table_bs_dir_text:
            messagebox.showerror("缺少模板", "请选择考勤表模板，或选择现有考勤表目录。")
            return

        generate_button.configure(state="disabled")
        clear_log(generate_log_text)
        append_log(generate_log_text, "开始生成考勤表...")
        try:
            template_path = choose_template_file(template_b_text or None, table_bs_dir_text or None)
            append_log(generate_log_text, f"使用模板: {template_path}")
            output_dir, generated_files, report_path = run_generate(
                table_c_path=Path(table_c_text),
                template_b_path=template_path,
                output_dir=Path(output_dir_text) if output_dir_text else None,
                count_holidays=count_holidays_var.get(),
            )
        except Exception as exc:
            append_log(generate_log_text, f"失败: {exc}")
            messagebox.showerror("执行失败", str(exc))
        else:
            append_log(generate_log_text, f"输出目录: {output_dir}")
            append_log(generate_log_text, f"生成数量: {len(generated_files)}")
            append_log(generate_log_text, f"说明文件: {report_path}")
            for generated_file in generated_files[:10]:
                append_log(generate_log_text, f"已生成: {generated_file.name}")
            if len(generated_files) > 10:
                append_log(generate_log_text, f"...其余 {len(generated_files) - 10} 个文件已省略显示")
            messagebox.showinfo(
                "生成完成",
                f"已完成生成。\n\n输出目录:\n{output_dir}\n\n生成数量: {len(generated_files)}\n\n说明文件:\n{report_path}",
            )
        finally:
            generate_button.configure(state="normal")

    ttk.Label(generate_tab, text="汇总表").grid(row=0, column=0, sticky="w", pady=(0, 8))
    ttk.Entry(generate_tab, textvariable=table_c_var).grid(row=0, column=1, sticky="ew", padx=(8, 8), pady=(0, 8))
    ttk.Button(generate_tab, text="选择文件", command=choose_table_c).grid(row=0, column=2, sticky="ew", pady=(0, 8))

    ttk.Label(generate_tab, text="考勤表模板").grid(row=1, column=0, sticky="w", pady=(0, 8))
    ttk.Entry(generate_tab, textvariable=template_b_var).grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=(0, 8))
    ttk.Button(generate_tab, text="选择模板", command=choose_template_b).grid(row=1, column=2, sticky="ew", pady=(0, 8))

    ttk.Label(generate_tab, text="现有考勤表目录").grid(row=2, column=0, sticky="w", pady=(0, 8))
    ttk.Entry(generate_tab, textvariable=generate_table_bs_var).grid(row=2, column=1, sticky="ew", padx=(8, 8), pady=(0, 8))
    ttk.Button(generate_tab, text="选择目录", command=choose_generate_bs_dir).grid(row=2, column=2, sticky="ew", pady=(0, 8))

    ttk.Label(generate_tab, text="输出目录").grid(row=3, column=0, sticky="w", pady=(0, 8))
    ttk.Entry(generate_tab, textvariable=generate_output_var).grid(row=3, column=1, sticky="ew", padx=(8, 8), pady=(0, 8))
    ttk.Button(generate_tab, text="选择目录", command=choose_generate_output_dir).grid(row=3, column=2, sticky="ew", pady=(0, 8))

    ttk.Label(
        generate_tab,
        text="说明：优先使用“考勤表模板”；如果不填模板，就会从“现有考勤表目录”里自动挑一个可用文件做模板。",
    ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(0, 8))

    ttk.Checkbutton(
        generate_tab,
        text="统计假期",
        variable=count_holidays_var,
    ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(0, 8))

    generate_button = ttk.Button(generate_tab, text="开始生成考勤表", command=start_generate)
    generate_button.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(0, 8))

    ttk.Label(generate_tab, text="日志").grid(row=7, column=0, columnspan=3, sticky="w", pady=(0, 6))
    generate_log_text = tk.Text(generate_tab, height=14, wrap="word", state="disabled")
    generate_log_text.grid(row=8, column=0, columnspan=3, sticky="nsew")

    template_left = ttk.Frame(template_tab)
    template_right = ttk.Frame(template_tab)
    template_left.grid(row=0, column=0, sticky="nsw", padx=(0, 12))
    template_right.grid(row=0, column=1, sticky="nsew")
    template_right.columnconfigure(1, weight=1)
    template_right.rowconfigure(4, weight=1)

    ttk.Label(template_left, text="模板列表").pack(anchor="w")
    template_listbox = tk.Listbox(template_left, height=18, exportselection=False)
    template_listbox.pack(fill="y", expand=False, pady=(6, 8))

    template_name_var = tk.StringVar()
    template_number_column_var = tk.StringVar()
    template_start_row_var = tk.StringVar()
    rules_tree = ttk.Treeview(
        template_right,
        columns=("field_name", "main_range", "table_b_cell", "compare_type"),
        show="headings",
        height=12,
    )
    for column, title, width in [
        ("field_name", "字段名", 140),
        ("main_range", "主表范围", 140),
        ("table_b_cell", "考勤表坐标", 120),
        ("compare_type", "比较类型", 100),
    ]:
        rules_tree.heading(column, text=title)
        rules_tree.column(column, width=width, anchor="w")

    current_template_name: Optional[str] = None

    def selected_rule_values() -> Optional[Tuple[str, str, str, str]]:
        selected = rules_tree.selection()
        if not selected:
            return None
        values = rules_tree.item(selected[0], "values")
        return tuple(str(value) for value in values)

    def open_rule_dialog(initial: Optional[Tuple[str, str, str, str]] = None) -> Optional[Tuple[str, str, str, str]]:
        dialog = tk.Toplevel(root)
        dialog.title("规则")
        dialog.transient(root)
        dialog.grab_set()
        dialog.columnconfigure(1, weight=1)

        field_var = tk.StringVar(value=initial[0] if initial else "")
        range_var = tk.StringVar(value=initial[1] if initial else "")
        b_cell_var = tk.StringVar(value=initial[2] if initial else "")
        compare_var = tk.StringVar(value=initial[3] if initial else COMPARE_TYPE_TEXT)
        result: Dict[str, Tuple[str, str, str, str]] = {}

        ttk.Label(dialog, text="字段名").grid(row=0, column=0, sticky="w", padx=12, pady=(12, 8))
        ttk.Entry(dialog, textvariable=field_var).grid(row=0, column=1, sticky="ew", padx=(0, 12), pady=(12, 8))
        ttk.Label(dialog, text="主表范围").grid(row=1, column=0, sticky="w", padx=12, pady=(0, 8))
        ttk.Entry(dialog, textvariable=range_var).grid(row=1, column=1, sticky="ew", padx=(0, 12), pady=(0, 8))
        ttk.Label(dialog, text="考勤表坐标").grid(row=2, column=0, sticky="w", padx=12, pady=(0, 8))
        ttk.Entry(dialog, textvariable=b_cell_var).grid(row=2, column=1, sticky="ew", padx=(0, 12), pady=(0, 8))
        ttk.Label(dialog, text="比较类型").grid(row=3, column=0, sticky="w", padx=12, pady=(0, 8))
        ttk.Combobox(
            dialog,
            textvariable=compare_var,
            state="readonly",
            values=(COMPARE_TYPE_TEXT, COMPARE_TYPE_NUMBER, COMPARE_TYPE_POSITION),
        ).grid(row=3, column=1, sticky="ew", padx=(0, 12), pady=(0, 8))

        ttk.Label(dialog, text="示例：主表范围填 F7-Fn；考勤表填 H4 或 SUM(G10:Gn) 或 SUM(H10:Hn,I10:In,J10:Jn)").grid(
            row=4, column=0, columnspan=2, sticky="w", padx=12, pady=(0, 8)
        )

        def confirm() -> None:
            try:
                parse_main_range(range_var.get().strip().upper())
                validate_table_b_expression(b_cell_var.get().strip().upper())
            except Exception as exc:
                messagebox.showerror("规则无效", str(exc), parent=dialog)
                return
            result["rule"] = (
                field_var.get().strip(),
                range_var.get().strip().upper(),
                b_cell_var.get().strip().upper(),
                compare_var.get().strip().lower(),
            )
            dialog.destroy()

        button_row = ttk.Frame(dialog)
        button_row.grid(row=5, column=0, columnspan=2, sticky="e", padx=12, pady=(4, 12))
        ttk.Button(button_row, text="取消", command=dialog.destroy).pack(side="right")
        ttk.Button(button_row, text="确定", command=confirm).pack(side="right", padx=(0, 8))

        dialog.wait_window()
        return result.get("rule")

    def fill_rules_tree(template: CheckTemplate) -> None:
        for item in rules_tree.get_children():
            rules_tree.delete(item)
        for rule in template.rules:
            rules_tree.insert("", "end", values=(rule.field_name, rule.main_range, rule.table_b_cell, rule.compare_type))

    def template_by_name(name: str) -> CheckTemplate:
        for template in templates_state:
            if template.name == name:
                return copy.deepcopy(template)
        raise WorkbookError(f"未找到模板: {name}")

    def load_template_to_form(name: str) -> None:
        nonlocal current_template_name
        template = template_by_name(name)
        current_template_name = template.name
        template_name_var.set(template.name)
        template_number_column_var.set(template.number_column)
        template_start_row_var.set(str(template.start_row))
        fill_rules_tree(template)

    def current_form_template() -> CheckTemplate:
        name = template_name_var.get().strip()
        number_column = template_number_column_var.get().strip().upper()
        start_row_text = template_start_row_var.get().strip()
        if not name:
            raise WorkbookError("模板名称不能为空")
        if not re.fullmatch(r"[A-Z]+", number_column):
            raise WorkbookError("编号列必须是列字母，例如 A")
        try:
            start_row = int(start_row_text)
        except Exception as exc:
            raise WorkbookError("起始行必须是整数") from exc
        rules = []
        for item in rules_tree.get_children():
            field_name, main_range, table_b_cell, compare_type = [str(value) for value in rules_tree.item(item, "values")]
            rules.append(
                CheckRule(
                    field_name=field_name,
                    main_range=main_range,
                    table_b_cell=table_b_cell,
                    compare_type=compare_type,
                )
            )
        template = CheckTemplate(name=name, number_column=number_column, start_row=start_row, rules=rules)
        parse_check_template(template)
        if not template.rules:
            raise WorkbookError("模板至少需要一条规则")
        return template

    def save_template_action() -> None:
        nonlocal templates_state, current_template_name
        try:
            template = current_form_template()
        except Exception as exc:
            messagebox.showerror("模板无效", str(exc))
            return
        templates_state = [item for item in templates_state if item.name not in {current_template_name, template.name}]
        templates_state.append(template)
        save_check_templates(templates_state)
        current_template_name = template.name
        refresh_template_choices(template.name)
        load_template_to_form(template.name)
        messagebox.showinfo("保存成功", f"已保存模板：{template.name}")

    def delete_template_action() -> None:
        nonlocal templates_state, current_template_name
        selected = template_name_var.get().strip()
        if not selected:
            return
        if selected == DEFAULT_CHECK_TEMPLATE_NAME:
            messagebox.showerror("不能删除", "默认模板不能删除。")
            return
        if not messagebox.askyesno("确认删除", f"确定删除模板“{selected}”吗？"):
            return
        templates_state = [item for item in templates_state if item.name != selected]
        save_check_templates(templates_state)
        refresh_template_choices(DEFAULT_CHECK_TEMPLATE_NAME)
        load_template_to_form(DEFAULT_CHECK_TEMPLATE_NAME)

    def new_template_action() -> None:
        nonlocal current_template_name
        current_template_name = None
        template_name_var.set("")
        template_number_column_var.set("A")
        template_start_row_var.set("7")
        for item in rules_tree.get_children():
            rules_tree.delete(item)

    def add_rule_action() -> None:
        result = open_rule_dialog()
        if result:
            rules_tree.insert("", "end", values=result)

    def edit_rule_action() -> None:
        values = selected_rule_values()
        if values is None:
            messagebox.showerror("未选择规则", "请先选择一条规则。")
            return
        result = open_rule_dialog(values)
        if result:
            selected = rules_tree.selection()[0]
            rules_tree.item(selected, values=result)

    def delete_rule_action() -> None:
        selected = rules_tree.selection()
        if not selected:
            messagebox.showerror("未选择规则", "请先选择一条规则。")
            return
        rules_tree.delete(selected[0])

    def move_rule_action(direction: int) -> None:
        selected = rules_tree.selection()
        if not selected:
            messagebox.showerror("未选择规则", "请先选择一条规则。")
            return
        item = selected[0]
        siblings = list(rules_tree.get_children())
        index = siblings.index(item)
        new_index = index + direction
        if new_index < 0 or new_index >= len(siblings):
            return
        rules_tree.move(item, "", new_index)
        rules_tree.selection_set(item)
        rules_tree.focus(item)

    def duplicate_template_action() -> None:
        nonlocal current_template_name
        try:
            template = current_form_template()
        except Exception as exc:
            messagebox.showerror("模板无效", str(exc))
            return
        existing_names = set(template_names())
        base_name = f"{template.name} - 副本"
        new_name = base_name
        index = 2
        while new_name in existing_names:
            new_name = f"{base_name} {index}"
            index += 1
        current_template_name = None
        template_name_var.set(new_name)
        messagebox.showinfo("已复制到当前表单", f"已创建模板副本名称：{new_name}\n点击“保存模板”后生效。")

    def export_template_action() -> None:
        try:
            template = current_form_template()
        except Exception as exc:
            messagebox.showerror("模板无效", str(exc))
            return
        path = filedialog.asksaveasfilename(
            title="导出模板",
            defaultextension=".json",
            initialfile=f"{template.name}.json",
            filetypes=[("JSON 文件", "*.json")],
        )
        if not path:
            return
        Path(path).write_text(json.dumps(template_to_dict(template), ensure_ascii=False, indent=2), encoding="utf-8")
        messagebox.showinfo("导出成功", f"模板已导出到：\n{path}")

    def import_template_action() -> None:
        nonlocal templates_state, current_template_name
        path = filedialog.askopenfilename(
            title="导入模板",
            filetypes=[("JSON 文件", "*.json")],
        )
        if not path:
            return
        try:
            raw = json.loads(Path(path).read_text(encoding="utf-8"))
            imported: List[CheckTemplate] = []
            if isinstance(raw, dict):
                imported.append(check_template_from_dict(raw))
            elif isinstance(raw, list):
                for item in raw:
                    if isinstance(item, dict):
                        imported.append(check_template_from_dict(item))
            if not imported:
                raise WorkbookError("模板文件中没有可导入的模板")
        except Exception as exc:
            messagebox.showerror("导入失败", str(exc))
            return

        for template in imported:
            templates_state = [item for item in templates_state if item.name != template.name]
            templates_state.append(template)
        save_check_templates(templates_state)
        current_template_name = imported[-1].name
        refresh_template_choices(current_template_name)
        load_template_to_form(current_template_name)
        messagebox.showinfo("导入成功", f"已导入 {len(imported)} 个模板。")

    def on_template_select(_event=None) -> None:
        selection = template_listbox.curselection()
        if not selection:
            return
        load_template_to_form(template_listbox.get(selection[0]))

    template_listbox.bind("<<ListboxSelect>>", on_template_select)

    ttk.Label(template_right, text="模板名称").grid(row=0, column=0, sticky="w", pady=(0, 8))
    ttk.Entry(template_right, textvariable=template_name_var).grid(row=0, column=1, sticky="ew", pady=(0, 8))

    ttk.Label(template_right, text="编号列").grid(row=1, column=0, sticky="w", pady=(0, 8))
    ttk.Entry(template_right, textvariable=template_number_column_var).grid(row=1, column=1, sticky="ew", pady=(0, 8))

    ttk.Label(template_right, text="数据起始行").grid(row=2, column=0, sticky="w", pady=(0, 8))
    ttk.Entry(template_right, textvariable=template_start_row_var).grid(row=2, column=1, sticky="ew", pady=(0, 8))

    ttk.Label(template_right, text="规则").grid(row=3, column=0, columnspan=2, sticky="w", pady=(4, 6))
    rules_tree.grid(row=4, column=0, columnspan=2, sticky="nsew")

    rules_button_bar = ttk.Frame(template_right)
    rules_button_bar.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(8, 8))
    ttk.Button(rules_button_bar, text="新增规则", command=add_rule_action).pack(side="left")
    ttk.Button(rules_button_bar, text="编辑规则", command=edit_rule_action).pack(side="left", padx=(8, 0))
    ttk.Button(rules_button_bar, text="删除规则", command=delete_rule_action).pack(side="left", padx=(8, 0))
    ttk.Button(rules_button_bar, text="上移规则", command=lambda: move_rule_action(-1)).pack(side="left", padx=(8, 0))
    ttk.Button(rules_button_bar, text="下移规则", command=lambda: move_rule_action(1)).pack(side="left", padx=(8, 0))

    ttk.Label(
        template_right,
        text="规则示例：主表 F3-Fn 对应 H4；或主表 N7-Nn 对应 SUM(G10:Gn)；比较类型可选 text / number / position。",
    ).grid(row=6, column=0, columnspan=2, sticky="w", pady=(0, 8))

    template_action_bar = ttk.Frame(template_right)
    template_action_bar.grid(row=7, column=0, columnspan=2, sticky="ew")
    ttk.Button(template_action_bar, text="新建模板", command=new_template_action).pack(side="left")
    ttk.Button(template_action_bar, text="复制模板", command=duplicate_template_action).pack(side="left", padx=(8, 0))
    ttk.Button(template_action_bar, text="保存模板", command=save_template_action).pack(side="left", padx=(8, 0))
    ttk.Button(template_action_bar, text="删除模板", command=delete_template_action).pack(side="left", padx=(8, 0))
    ttk.Button(template_action_bar, text="导入模板", command=import_template_action).pack(side="left", padx=(8, 0))
    ttk.Button(template_action_bar, text="导出模板", command=export_template_action).pack(side="left", padx=(8, 0))

    refresh_template_choices(DEFAULT_CHECK_TEMPLATE_NAME)
    load_template_to_form(check_template_var.get() or DEFAULT_CHECK_TEMPLATE_NAME)

    root.mainloop()


if __name__ == "__main__":
    try:
        raise SystemExit(cli(sys.argv[1:]))
    except WorkbookError as exc:
        print(f"错误: {exc}", file=sys.stderr)
        raise SystemExit(1)
