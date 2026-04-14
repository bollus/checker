from __future__ import annotations

import argparse
import copy
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


def qn(name: str, ns: str = MAIN_NS) -> str:
    return f"{{{ns}}}{name}"


def split_cell_ref(cell_ref: str) -> Tuple[str, int]:
    match = re.fullmatch(r"([A-Z]+)(\d+)", cell_ref)
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    return match.group(1), int(match.group(2))


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


class WorkbookError(Exception):
    pass


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
    row_num: int,
    table_a_sheet: SheetXml,
    table_b_sheet: Optional[SheetXml],
    table_b_file: Optional[Path],
) -> List[Mismatch]:
    mismatches: List[Mismatch] = []
    file_name = table_b_file.name if table_b_file else "未找到匹配考勤表"

    for table_a_col, (table_b_ref, field_name) in TEXT_FIELDS.items():
        table_a_cell = f"{table_a_col}{row_num}"
        left = table_a_sheet.get_value(table_a_cell)
        right = None if table_b_sheet is None else table_b_sheet.get_value(table_b_ref)
        if normalize_text(left) != normalize_text(right):
            mismatches.append(
                Mismatch(
                    row_num=row_num,
                    table_a_cell=table_a_cell,
                    field_name=field_name,
                    table_a_value=to_display(left),
                    table_b_value=to_display(right),
                    table_b_file=file_name,
                )
            )

    for table_a_col, (table_b_ref, field_name) in NUMERIC_FIELDS.items():
        table_a_cell = f"{table_a_col}{row_num}"
        left = table_a_sheet.get_value(table_a_cell)
        right = None if table_b_sheet is None else table_b_sheet.get_value(table_b_ref)
        if table_b_sheet is None:
            mismatch = True
        else:
            mismatch = abs(normalize_number(left) - normalize_number(right)) > NUMERIC_TOLERANCE
        if mismatch:
            mismatches.append(
                Mismatch(
                    row_num=row_num,
                    table_a_cell=table_a_cell,
                    field_name=field_name,
                    table_a_value=to_display(left),
                    table_b_value=to_display(right),
                    table_b_file=file_name,
                )
            )
    return mismatches


def locate_data_rows(sheet: SheetXml, start_row: int = 7) -> List[int]:
    rows: List[int] = []
    current = start_row
    while True:
        refs = [f"A{current}", f"F{current}", f"J{current}", f"N{current}", f"W{current}", f"X{current}", f"Y{current}"]
        if all(sheet.get_value(ref) in (None, "") for ref in refs):
            break
        rows.append(current)
        current += 1
    return rows


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


def run_check(table_a_path: Path, table_bs_folder: Path, output_path: Optional[Path] = None) -> Tuple[Path, Path, List[Mismatch], List[str]]:
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

    table_b_files, warnings = build_table_b_index(table_bs_folder)
    data_rows = locate_data_rows(table_a_sheet)

    styles_root = table_a_book.load_xml("xl/styles.xml")
    highlight_style_for = add_highlight_style(styles_root)

    mismatches: List[Mismatch] = []
    workbook_cache: Dict[Path, WorkbookSheet] = {}

    for row_num in data_rows:
        row_number_raw = table_a_sheet.get_value(f"A{row_num}")
        if row_number_raw is None or str(row_number_raw).strip() == "":
            continue
        try:
            file_number = int(Decimal(str(row_number_raw)))
        except InvalidOperation as exc:
            raise WorkbookError(f"主表 A{row_num} 不是有效编号: {row_number_raw!r}") from exc

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

        row_mismatches = compare_row(row_num, table_a_sheet, table_b_sheet, table_b_path)
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
    root.geometry("860x620")

    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True)

    check_tab = ttk.Frame(notebook, padding=16)
    generate_tab = ttk.Frame(notebook, padding=16)
    notebook.add(check_tab, text="工资表核对")
    notebook.add(generate_tab, text="根据表C生成表B")

    for tab in (check_tab, generate_tab):
        tab.columnconfigure(1, weight=1)
    check_tab.rowconfigure(5, weight=1)
    generate_tab.rowconfigure(7, weight=1)

    def clear_log(widget: tk.Text) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.configure(state="disabled")

    def append_log(widget: tk.Text, text: str) -> None:
        widget.configure(state="normal")
        widget.insert("end", text + "\n")
        widget.see("end")
        widget.configure(state="disabled")

    table_a_var = tk.StringVar()
    table_bs_var = tk.StringVar()
    output_var = tk.StringVar()

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
        if not table_a_text:
            messagebox.showerror("缺少主表", "请选择主表文件。")
            return
        if not table_bs_text:
            messagebox.showerror("缺少目录", "请选择考勤表目录。")
            return
        check_button.configure(state="disabled")
        clear_log(check_log_text)
        append_log(check_log_text, "开始核对...")
        try:
            output, report, mismatches, warnings = run_check(
                table_a_path=Path(table_a_text),
                table_bs_folder=Path(table_bs_text),
                output_path=Path(output_text) if output_text else None,
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

    check_button = ttk.Button(check_tab, text="开始核对", command=start_check)
    check_button.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(4, 8))

    ttk.Label(check_tab, text="日志").grid(row=4, column=0, columnspan=3, sticky="w", pady=(0, 6))
    check_log_text = tk.Text(check_tab, height=16, wrap="word", state="disabled")
    check_log_text.grid(row=5, column=0, columnspan=3, sticky="nsew")

    table_c_var = tk.StringVar()
    template_b_var = tk.StringVar()
    generate_table_bs_var = tk.StringVar()
    generate_output_var = tk.StringVar()

    def choose_table_c() -> None:
        path = filedialog.askopenfilename(
            title="选择表C",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm")],
        )
        if not path:
            return
        table_c_var.set(path)
        if not generate_output_var.get():
            generate_output_var.set(str(Path(path).with_name(f"{Path(path).stem}_生成表Bs")))

    def choose_template_b() -> None:
        path = filedialog.askopenfilename(
            title="选择表B模板",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm")],
        )
        if path:
            template_b_var.set(path)

    def choose_generate_bs_dir() -> None:
        path = filedialog.askdirectory(title="选择现有表B目录")
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
            messagebox.showerror("缺少表C", "请选择表C文件。")
            return
        if not template_b_text and not table_bs_dir_text:
            messagebox.showerror("缺少模板", "请选择表B模板，或选择现有表B目录。")
            return

        generate_button.configure(state="disabled")
        clear_log(generate_log_text)
        append_log(generate_log_text, "开始生成表B...")
        try:
            template_path = choose_template_file(template_b_text or None, table_bs_dir_text or None)
            append_log(generate_log_text, f"使用模板: {template_path}")
            output_dir, generated_files, report_path = run_generate(
                table_c_path=Path(table_c_text),
                template_b_path=template_path,
                output_dir=Path(output_dir_text) if output_dir_text else None,
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

    ttk.Label(generate_tab, text="表C").grid(row=0, column=0, sticky="w", pady=(0, 8))
    ttk.Entry(generate_tab, textvariable=table_c_var).grid(row=0, column=1, sticky="ew", padx=(8, 8), pady=(0, 8))
    ttk.Button(generate_tab, text="选择文件", command=choose_table_c).grid(row=0, column=2, sticky="ew", pady=(0, 8))

    ttk.Label(generate_tab, text="表B模板").grid(row=1, column=0, sticky="w", pady=(0, 8))
    ttk.Entry(generate_tab, textvariable=template_b_var).grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=(0, 8))
    ttk.Button(generate_tab, text="选择模板", command=choose_template_b).grid(row=1, column=2, sticky="ew", pady=(0, 8))

    ttk.Label(generate_tab, text="现有表B目录").grid(row=2, column=0, sticky="w", pady=(0, 8))
    ttk.Entry(generate_tab, textvariable=generate_table_bs_var).grid(row=2, column=1, sticky="ew", padx=(8, 8), pady=(0, 8))
    ttk.Button(generate_tab, text="选择目录", command=choose_generate_bs_dir).grid(row=2, column=2, sticky="ew", pady=(0, 8))

    ttk.Label(generate_tab, text="输出目录").grid(row=3, column=0, sticky="w", pady=(0, 8))
    ttk.Entry(generate_tab, textvariable=generate_output_var).grid(row=3, column=1, sticky="ew", padx=(8, 8), pady=(0, 8))
    ttk.Button(generate_tab, text="选择目录", command=choose_generate_output_dir).grid(row=3, column=2, sticky="ew", pady=(0, 8))

    ttk.Label(
        generate_tab,
        text="说明：优先使用“表B模板”；如果不填模板，就会从“现有表B目录”里自动挑一个可用文件做模板。",
    ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(0, 8))

    generate_button = ttk.Button(generate_tab, text="开始生成表B", command=start_generate)
    generate_button.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(0, 8))

    ttk.Label(generate_tab, text="日志").grid(row=6, column=0, columnspan=3, sticky="w", pady=(0, 6))
    generate_log_text = tk.Text(generate_tab, height=14, wrap="word", state="disabled")
    generate_log_text.grid(row=7, column=0, columnspan=3, sticky="nsew")

    root.mainloop()


if __name__ == "__main__":
    try:
        raise SystemExit(cli(sys.argv[1:]))
    except WorkbookError as exc:
        print(f"错误: {exc}", file=sys.stderr)
        raise SystemExit(1)
