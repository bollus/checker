from __future__ import annotations

import json
import os
import sys
import traceback
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Dict, Iterable, Iterator, List, Optional

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import excel_check_tool as check_tool
import generate_table_bs as generator


def json_default(value: Any) -> Any:
    if isinstance(value, Path):
        return str(value)
    return str(value)


def read_request() -> Dict[str, Any]:
    raw = sys.stdin.read()
    if not raw.strip():
        raise check_tool.WorkbookError("没有收到任务参数")
    data = json.loads(raw)
    if not isinstance(data, dict):
        raise check_tool.WorkbookError("任务参数必须是 JSON 对象")
    return data


def ok(data: Dict[str, Any]) -> Dict[str, Any]:
    return {"ok": True, "data": data, "warnings": [], "errors": []}


def fail(exc: BaseException) -> Dict[str, Any]:
    message = str(exc) or exc.__class__.__name__
    return {
        "ok": False,
        "data": None,
        "warnings": [],
        "errors": [message],
        "traceback": traceback.format_exc(),
    }


def load_template_from_payload(payload: Dict[str, Any]) -> Optional[check_tool.CheckTemplate]:
    raw = payload.get("template")
    if isinstance(raw, dict):
        return check_tool.check_template_from_dict(raw)

    template_path = payload.get("template_path")
    if template_path:
        data = json.loads(Path(template_path).read_text(encoding="utf-8"))
        return check_tool.check_template_from_dict(data)

    return None


@contextmanager
def using_template(template: Optional[check_tool.CheckTemplate]) -> Iterator[None]:
    if template is None:
        yield
        return

    original = check_tool.get_check_template

    def get_check_template(_template_name: Optional[str] = None) -> check_tool.CheckTemplate:
        return template

    check_tool.get_check_template = get_check_template
    try:
        yield
    finally:
        check_tool.get_check_template = original


def mismatch_to_dict(item: check_tool.Mismatch) -> Dict[str, Any]:
    return {
        "row_num": item.row_num,
        "table_a_cell": item.table_a_cell,
        "field_name": item.field_name,
        "table_a_value": item.table_a_value,
        "table_b_value": item.table_b_value,
        "table_b_file": item.table_b_file,
    }


def action_check(payload: Dict[str, Any]) -> Dict[str, Any]:
    template = load_template_from_payload(payload)
    with using_template(template):
        output, report, mismatches, warnings = check_tool.run_check(
            table_a_path=Path(payload["table_a_path"]),
            table_bs_folder=Path(payload["table_bs_folder"]),
            output_path=Path(payload["output_path"]) if payload.get("output_path") else None,
            template_name=payload.get("template_name"),
        )
    return ok(
        {
            "output_path": str(output),
            "report_path": str(report),
            "mismatch_count": len(mismatches),
            "mismatches": [mismatch_to_dict(item) for item in mismatches],
            "warnings": list(warnings),
        }
    )


def action_generate(payload: Dict[str, Any]) -> Dict[str, Any]:
    output_dir, files, report = generator.run_generate(
        table_c_path=Path(payload["table_c_path"]),
        template_b_path=Path(payload["template_b_path"]),
        output_dir=Path(payload["output_dir"]) if payload.get("output_dir") else None,
        count_holidays=bool(payload.get("count_holidays", False)),
        signature_scale=int(payload.get("signature_scale", 100)),
        morning_start=str(payload.get("morning_start", generator.DEFAULT_MORNING_START)),
        morning_end=str(payload.get("morning_end", generator.DEFAULT_MORNING_END)),
        afternoon_start=str(payload.get("afternoon_start", generator.DEFAULT_AFTERNOON_START)),
        afternoon_end=str(payload.get("afternoon_end", generator.DEFAULT_AFTERNOON_END)),
        normal_hours=str(payload.get("normal_hours", generator.DEFAULT_NORMAL_HOURS)),
    )
    return ok(
        {
            "output_dir": str(output_dir),
            "report_path": str(report),
            "generated_count": len(files),
            "generated_files": [str(path) for path in files],
        }
    )


def action_list_templates(_payload: Dict[str, Any]) -> Dict[str, Any]:
    templates = [check_tool.template_to_dict(template) for template in check_tool.load_check_templates()]
    return ok({"templates": templates})


def action_save_template(payload: Dict[str, Any]) -> Dict[str, Any]:
    template = check_tool.check_template_from_dict(payload["template"])
    existing = [item for item in check_tool.load_check_templates() if item.name != template.name]
    existing.append(template)
    check_tool.save_check_templates(existing)
    return ok({"template": check_tool.template_to_dict(template)})


def action_delete_template(payload: Dict[str, Any]) -> Dict[str, Any]:
    name = str(payload["name"]).strip()
    templates = [item for item in check_tool.load_check_templates() if item.name != name]
    check_tool.save_check_templates(templates)
    return ok({"deleted": name})


def preview_cells(sheet: check_tool.SheetXml, max_rows: int, max_cols: int) -> List[Dict[str, Any]]:
    cells: List[Dict[str, Any]] = []
    for row in range(1, max_rows + 1):
        for col_index in range(1, max_cols + 1):
            col = check_tool.num_to_col(col_index)
            ref = f"{col}{row}"
            value = check_tool.to_display(sheet.get_value(ref))
            if value:
                cells.append({"ref": ref, "row": row, "col": col, "value": value})
    return cells


def used_bounds(sheet: check_tool.SheetXml) -> Dict[str, int]:
    max_row = max(sheet.rows) if sheet.rows else 1
    max_col = 1
    for ref in sheet.cells:
        col, _row = check_tool.split_cell_ref(ref)
        max_col = max(max_col, check_tool.col_to_num(col))
    return {"max_row": max_row, "max_col": max_col}


def action_inspect_workbook(payload: Dict[str, Any]) -> Dict[str, Any]:
    path = Path(payload["path"])
    max_rows = int(payload.get("max_rows", 80))
    max_cols = int(payload.get("max_cols", 30))
    book = check_tool.SpreadsheetZip(path)
    result = []
    for item in check_tool.workbook_sheets(book):
        bounds = used_bounds(item.sheet)
        result.append(
            {
                "name": item.display_name,
                "part_name": item.part_name,
                "bounds": bounds,
                "cells": preview_cells(item.sheet, min(bounds["max_row"], max_rows), min(bounds["max_col"], max_cols)),
            }
        )
    return ok({"path": str(path), "sheets": result})


def action_validate_template(payload: Dict[str, Any]) -> Dict[str, Any]:
    template = check_tool.check_template_from_dict(payload["template"])
    parsed = check_tool.parse_check_template(template)
    return ok({"rule_count": len(parsed), "template": check_tool.template_to_dict(template)})


ACTIONS = {
    "check": action_check,
    "generate": action_generate,
    "list_templates": action_list_templates,
    "save_template": action_save_template,
    "delete_template": action_delete_template,
    "inspect_workbook": action_inspect_workbook,
    "validate_template": action_validate_template,
}


def main() -> int:
    try:
        request = read_request()
        action = str(request.get("action", "")).strip()
        payload = request.get("payload") or {}
        if action not in ACTIONS:
            raise check_tool.WorkbookError(f"未知任务: {action}")
        response = ACTIONS[action](payload)
    except BaseException as exc:
        response = fail(exc)
    sys.stdout.write(json.dumps(response, ensure_ascii=False, default=json_default))
    return 0 if response.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
