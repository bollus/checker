import { invoke } from "@tauri-apps/api/core";

export type CompareType = "text" | "number" | "position";

export interface CheckRule {
  field_name: string;
  main_range: string;
  table_b_cell: string;
  compare_type: CompareType;
}

export interface CheckTemplate {
  name: string;
  number_column: string;
  start_row: number;
  rules: CheckRule[];
}

export interface BackendEnvelope<T> {
  ok: boolean;
  data: T | null;
  warnings: string[];
  errors: string[];
  traceback?: string;
}

export interface Mismatch {
  row_num: number;
  table_a_cell: string;
  field_name: string;
  table_a_value: string;
  table_b_value: string;
  table_b_file: string;
}

export interface CheckResult {
  output_path: string;
  report_path: string;
  mismatch_count: number;
  mismatches: Mismatch[];
  warnings: string[];
}

export interface GenerateResult {
  output_dir: string;
  report_path: string;
  generated_count: number;
  generated_files: string[];
}

export interface WorkbookCell {
  ref: string;
  row: number;
  col: string;
  value: string;
}

export interface WorkbookSheetPreview {
  name: string;
  part_name: string;
  bounds: { max_row: number; max_col: number };
  cells: WorkbookCell[];
}

export interface WorkbookPreview {
  path: string;
  sheets: WorkbookSheetPreview[];
}

export async function backend<T>(action: string, payload: unknown): Promise<T> {
  const result = await invoke<BackendEnvelope<T>>("run_backend", { action, payload });
  if (!result.ok) {
    throw new Error(result.errors.join("\n") || "任务失败");
  }
  return result.data as T;
}

export function openPath(path: string): Promise<void> {
  return invoke("open_path", { path });
}

export function revealPath(path: string): Promise<void> {
  return invoke("reveal_path", { path });
}
