use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, HashMap};
use std::collections::HashSet;
use std::fs::{self, File};
use std::io::{Cursor, Read, Write};
use std::path::{Path, PathBuf};
use zip::{write::SimpleFileOptions, CompressionMethod, ZipArchive, ZipWriter};

const NUMERIC_TOLERANCE: f64 = 0.000001;

#[derive(Debug, Deserialize)]
pub struct CheckPayload {
    pub table_a_path: String,
    pub table_bs_folder: String,
    pub output_path: Option<String>,
    pub template: CheckTemplate,
    #[serde(default)]
    pub position_aliases_path: Option<String>,
    #[serde(default)]
    pub position_rules_path: Option<String>,
}

#[derive(Debug, Deserialize)]
pub struct CheckTemplate {
    pub name: String,
    pub number_column: String,
    pub start_row: u32,
    pub rules: Vec<CheckRule>,
}

#[derive(Debug, Deserialize)]
pub struct CheckRule {
    pub field_name: String,
    pub main_range: String,
    pub table_b_cell: String,
    pub compare_type: String,
}

#[derive(Debug, Clone)]
struct ParsedRule {
    field_name: String,
    table_b_cell: String,
    compare_type: String,
    main_column: String,
    main_start_row: u32,
}

struct CompareContext {
    position_aliases: HashMap<String, String>,
    position_token_aliases: HashMap<String, Vec<String>>,
    position_optional_tokens: HashSet<String>,
}

#[derive(Debug, Serialize)]
pub struct CheckResult {
    pub output_path: String,
    pub report_path: String,
    pub mismatch_count: usize,
    pub mismatches: Vec<Mismatch>,
    pub warnings: Vec<String>,
    pub progress: ProgressSnapshot,
}

#[derive(Debug, Serialize)]
pub struct Mismatch {
    pub row_num: u32,
    pub table_a_cell: String,
    pub field_name: String,
    pub table_a_value: String,
    pub table_b_value: String,
    pub table_b_file: String,
}

#[derive(Debug, Serialize)]
pub struct ProgressSnapshot {
    pub current: usize,
    pub total: usize,
    pub message: String,
}

#[derive(Debug, Clone)]
struct Sheet {
    cells: HashMap<String, String>,
    styles: HashMap<String, u32>,
}

#[derive(Debug)]
struct Workbook {
    entries: HashMap<String, Vec<u8>>,
    sheets: Vec<WorkbookSheet>,
}

#[derive(Debug)]
struct WorkbookSheet {
    part_name: String,
    sheet: Sheet,
}

pub fn run_check(payload: CheckPayload) -> Result<CheckResult, String> {
    let table_a_path = PathBuf::from(&payload.table_a_path);
    let table_bs_folder = PathBuf::from(&payload.table_bs_folder);
    if !matches!(table_a_path.extension().and_then(|item| item.to_str()).map(|item| item.to_ascii_lowercase()).as_deref(), Some("xlsx" | "xlsm")) {
        return Err("Rust 核对引擎只支持 .xlsx 或 .xlsm".to_string());
    }
    if !table_a_path.is_file() {
        return Err(format!("Rust 核对引擎未找到主工资表: {}", payload.table_a_path));
    }
    if !table_bs_folder.is_dir() {
        return Err(format!("Rust 核对引擎未找到考勤表目录: {}", payload.table_bs_folder));
    }

    let output_path = payload
        .output_path
        .as_ref()
        .filter(|item| !item.trim().is_empty())
        .map(PathBuf::from)
        .unwrap_or_else(|| next_output_path(&table_a_path));
    let report_path = next_report_path(&output_path);

    let parsed_rules = parse_rules(&payload.template)?;
    let position_rules = load_position_rules(payload.position_rules_path.as_deref());
    let compare_context = CompareContext {
        position_aliases: load_position_aliases(payload.position_aliases_path.as_deref()),
        position_token_aliases: position_rules.token_aliases,
        position_optional_tokens: position_rules.optional_tokens,
    };
    let table_b_index = build_table_b_index(&table_bs_folder)?;
    let mut warnings = table_b_index.warnings;
    warnings.push(format!("Rust 实验版核对模板: {}", payload.template.name));

    let table_a_book = Workbook::open(&table_a_path)?;
    let table_a_sheet_info = table_a_book
        .sheets
        .first()
        .ok_or_else(|| "Rust 核对引擎未读取到主表工作表".to_string())?;
    let data_offsets = locate_data_offsets(&table_a_sheet_info.sheet, &payload.template, &parsed_rules);
    let total = data_offsets.len();
    let mut progress = ProgressSnapshot {
        current: 0,
        total,
        message: String::new(),
    };
    let mut workbook_cache: HashMap<PathBuf, Sheet> = HashMap::new();
    let mut mismatches = Vec::new();

    for (index, offset) in data_offsets.into_iter().enumerate() {
        let display_row = payload.template.start_row + offset;
        progress.current = index + 1;
        progress.message = format!("核对第 {display_row} 行");

        let number_ref = format!("{}{}", payload.template.number_column.to_ascii_uppercase(), display_row);
        let number_value = table_a_sheet_info.sheet.get_value(&number_ref);
        if number_value.trim().is_empty() {
            continue;
        }
        let file_number = parse_number(number_value)
            .map(|value| value as i32)
            .map_err(|_| format!("主表 {number_ref} 不是有效编号: {number_value:?}"))?;
        let table_b_path = table_b_index.files.get(&file_number).cloned();
        let table_b_sheet = if let Some(path) = table_b_path.as_ref() {
            if !workbook_cache.contains_key(path) {
                let book = Workbook::open(path)?;
                let sheet = choose_timesheet_sheet(book.sheets, &parsed_rules)?;
                workbook_cache.insert(path.clone(), sheet.sheet);
            }
            workbook_cache.get(path)
        } else {
            warnings.push(format!("No.{file_number} 未找到匹配考勤表文件"));
            None
        };

        for mismatch in compare_row(
            offset,
            display_row,
            &table_a_sheet_info.sheet,
            table_b_sheet,
            table_b_path.as_deref(),
            &parsed_rules,
            &compare_context,
        )? {
            mismatches.push(mismatch);
        }
    }

    write_highlighted_workbook(&table_a_book, table_a_sheet_info, &mismatches, &output_path)?;
    create_report(&report_path, &table_a_path, &output_path, &warnings, &mismatches)?;

    Ok(CheckResult {
        output_path: output_path.to_string_lossy().to_string(),
        report_path: report_path.to_string_lossy().to_string(),
        mismatch_count: mismatches.len(),
        mismatches,
        warnings,
        progress,
    })
}

struct TableBIndex {
    files: HashMap<i32, PathBuf>,
    warnings: Vec<String>,
}

impl Workbook {
    fn open(path: &Path) -> Result<Self, String> {
        let data = fs::read(path).map_err(|err| format!("读取 Excel 文件失败 {}: {err}", path.display()))?;
        let mut archive = ZipArchive::new(Cursor::new(data)).map_err(|err| format!("打开 Excel ZIP 失败 {}: {err}", path.display()))?;
        let mut entries = HashMap::new();
        for index in 0..archive.len() {
            let mut file = archive.by_index(index).map_err(|err| format!("读取 ZIP 项失败: {err}"))?;
            if file.is_dir() {
                continue;
            }
            let mut bytes = Vec::new();
            file.read_to_end(&mut bytes).map_err(|err| format!("读取 ZIP 内容失败: {err}"))?;
            entries.insert(file.name().replace('\\', "/"), bytes);
        }
        let shared_strings = load_shared_strings(&entries);
        let sheets = load_workbook_sheets(&entries, &shared_strings)?;
        Ok(Self { entries, sheets })
    }
}

impl Sheet {
    fn get_value(&self, ref_name: &str) -> &str {
        self.cells.get(&ref_name.to_ascii_uppercase()).map(String::as_str).unwrap_or("")
    }
}

fn load_shared_strings(entries: &HashMap<String, Vec<u8>>) -> Vec<String> {
    let Some(raw) = entries.get("xl/sharedStrings.xml") else {
        return Vec::new();
    };
    let mut reader = Reader::from_reader(Cursor::new(raw.as_slice()));
    let mut buf = Vec::new();
    let mut items = Vec::new();
    let mut in_si = false;
    let mut current = String::new();

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"si" => {
                in_si = true;
                current.clear();
            }
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"si" => {
                items.push(current.clone());
                in_si = false;
            }
            Ok(Event::Text(event)) if in_si => {
                if let Ok(text) = event.decode() {
                    current.push_str(&text);
                }
            }
            Ok(Event::CData(event)) if in_si => {
                if let Ok(text) = event.decode() {
                    current.push_str(&text);
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(_) => break,
        }
    }
    items
}

fn load_workbook_sheets(entries: &HashMap<String, Vec<u8>>, shared_strings: &[String]) -> Result<Vec<WorkbookSheet>, String> {
    let mut rel_targets = HashMap::new();
    let rels_raw = entries
        .get("xl/_rels/workbook.xml.rels")
        .ok_or_else(|| "Excel 缺少必要文件: xl/_rels/workbook.xml.rels".to_string())?;
    let mut rel_reader = Reader::from_reader(Cursor::new(rels_raw.as_slice()));
    let mut rel_buf = Vec::new();
    loop {
        rel_buf.clear();
        match rel_reader.read_event_into(&mut rel_buf) {
            Ok(Event::Empty(event)) | Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"Relationship" => {
                if let (Some(id), Some(target)) = (xml_attr(&rel_reader, &event, b"Id"), xml_attr(&rel_reader, &event, b"Target")) {
                    let normalized = if target.starts_with("xl/") {
                        target
                    } else {
                        format!("xl/{target}")
                    };
                    rel_targets.insert(id, normalized);
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(err) => return Err(format!("解析 workbook.xml.rels 失败: {err}")),
        }
    }

    let mut sheets = Vec::new();
    let workbook_raw = entries
        .get("xl/workbook.xml")
        .ok_or_else(|| "Excel 缺少必要文件: xl/workbook.xml".to_string())?;
    let mut workbook_reader = Reader::from_reader(Cursor::new(workbook_raw.as_slice()));
    let mut workbook_buf = Vec::new();
    loop {
        workbook_buf.clear();
        match workbook_reader.read_event_into(&mut workbook_buf) {
            Ok(Event::Empty(event)) | Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"sheet" => {
                let rel_id = xml_attr(&workbook_reader, &event, b"r:id").or_else(|| xml_attr(&workbook_reader, &event, b"id"));
                let Some(rel_id) = rel_id else {
                    continue;
                };
                let Some(target) = rel_targets.get(&rel_id) else {
                    continue;
                };
                if !target.starts_with("xl/worksheets/") {
                    continue;
                }
                let raw = entries
                    .get(target)
                    .ok_or_else(|| format!("Excel 缺少工作表文件: {target}"))?;
                sheets.push(WorkbookSheet {
                    part_name: target.clone(),
                    sheet: parse_sheet(raw, shared_strings)?,
                });
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(err) => return Err(format!("解析 workbook.xml 失败: {err}")),
        }
    }
    if sheets.is_empty() {
        return Err("Rust 核对引擎未读取到工作表".to_string());
    }
    Ok(sheets)
}

#[derive(Default)]
struct CellBuild {
    ref_name: String,
    cell_type: String,
    style_id: u32,
    value_text: String,
    inline_text: String,
    in_value: bool,
    in_text: bool,
}

fn parse_sheet(raw: &[u8], shared_strings: &[String]) -> Result<Sheet, String> {
    let mut cells = HashMap::new();
    let mut styles = HashMap::new();
    let mut reader = Reader::from_reader(Cursor::new(raw));
    let mut buf = Vec::new();
    let mut current: Option<CellBuild> = None;

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"c" => {
                current = Some(CellBuild {
                    ref_name: xml_attr(&reader, &event, b"r").unwrap_or_default().to_ascii_uppercase(),
                    cell_type: xml_attr(&reader, &event, b"t").unwrap_or_default(),
                    style_id: xml_attr(&reader, &event, b"s").and_then(|value| value.parse::<u32>().ok()).unwrap_or(0),
                    ..CellBuild::default()
                });
            }
            Ok(Event::Empty(event)) if local_name(event.name().as_ref()) == b"c" => {
                if let Some(ref_name) = xml_attr(&reader, &event, b"r") {
                    let ref_name = ref_name.to_ascii_uppercase();
                    let style_id = xml_attr(&reader, &event, b"s").and_then(|value| value.parse::<u32>().ok()).unwrap_or(0);
                    styles.insert(ref_name.clone(), style_id);
                    cells.entry(ref_name).or_insert_with(String::new);
                }
            }
            Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"v" => {
                if let Some(cell) = current.as_mut() {
                    cell.in_value = true;
                }
            }
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"v" => {
                if let Some(cell) = current.as_mut() {
                    cell.in_value = false;
                }
            }
            Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"t" => {
                if let Some(cell) = current.as_mut() {
                    cell.in_text = true;
                }
            }
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"t" => {
                if let Some(cell) = current.as_mut() {
                    cell.in_text = false;
                }
            }
            Ok(Event::Text(event)) => {
                if let Some(cell) = current.as_mut() {
                    if let Ok(text) = event.decode() {
                        if cell.in_value {
                            cell.value_text.push_str(&text);
                        }
                        if cell.in_text {
                            cell.inline_text.push_str(&text);
                        }
                    }
                }
            }
            Ok(Event::CData(event)) => {
                if let Some(cell) = current.as_mut() {
                    if let Ok(text) = event.decode() {
                        if cell.in_value {
                            cell.value_text.push_str(&text);
                        }
                        if cell.in_text {
                            cell.inline_text.push_str(&text);
                        }
                    }
                }
            }
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"c" => {
                if let Some(cell) = current.take() {
                    if !cell.ref_name.is_empty() {
                        let value = finalize_cell(cell, shared_strings);
                        styles.insert(value.0.clone(), value.2);
                        cells.insert(value.0, value.1);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(err) => return Err(format!("解析 worksheet XML 失败: {err}")),
        }
    }
    Ok(Sheet { cells, styles })
}

fn finalize_cell(cell: CellBuild, shared_strings: &[String]) -> (String, String, u32) {
    let value = if cell.cell_type == "inlineStr" {
        cell.inline_text
    } else if cell.cell_type == "s" {
        cell.value_text
            .parse::<usize>()
            .ok()
            .and_then(|index| shared_strings.get(index).cloned())
            .unwrap_or_default()
    } else if cell.cell_type == "b" {
        if cell.value_text == "1" { "TRUE".to_string() } else { "FALSE".to_string() }
    } else {
        cell.value_text
    };
    (cell.ref_name, value, cell.style_id)
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn xml_attr(reader: &Reader<Cursor<&[u8]>>, event: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    for attr in event.attributes().with_checks(false).flatten() {
        if local_name(attr.key.as_ref()) == local_name(name) {
            if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                return Some(value.into_owned());
            }
        }
    }
    None
}

#[derive(Clone)]
struct XfTemplate {
    attrs: Vec<(String, String)>,
}

fn write_highlighted_workbook(
    book: &Workbook,
    main_sheet: &WorkbookSheet,
    mismatches: &[Mismatch],
    output_path: &Path,
) -> Result<(), String> {
    let mismatch_cells: HashSet<String> = mismatches.iter().map(|item| item.table_a_cell.to_ascii_uppercase()).collect();
    let base_styles: HashSet<u32> = mismatch_cells
        .iter()
        .map(|cell| main_sheet.sheet.styles.get(cell).copied().unwrap_or(0))
        .collect();
    let (styles_xml, style_map) = build_highlight_styles(
        book.entries
            .get("xl/styles.xml")
            .ok_or_else(|| "Excel 缺少必要文件: xl/styles.xml".to_string())?,
        &base_styles,
    )?;
    let sheet_xml = rewrite_sheet_styles(
        book.entries
            .get(&main_sheet.part_name)
            .ok_or_else(|| format!("Excel 缺少主工作表文件: {}", main_sheet.part_name))?,
        &mismatch_cells,
        &main_sheet.sheet.styles,
        &style_map,
    )?;

    if let Some(parent) = output_path.parent() {
        fs::create_dir_all(parent).map_err(|err| format!("创建输出目录失败: {err}"))?;
    }
    let file = File::create(output_path).map_err(|err| format!("创建 Rust 核对结果文件失败: {err}"))?;
    let mut writer = ZipWriter::new(file);
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    let mut names = book.entries.keys().cloned().collect::<Vec<_>>();
    names.sort();
    for name in names {
        writer.start_file(&name, options).map_err(|err| format!("写入 ZIP 项失败 {name}: {err}"))?;
        if name == "xl/styles.xml" {
            writer.write_all(&styles_xml).map_err(|err| format!("写入 styles.xml 失败: {err}"))?;
        } else if name == main_sheet.part_name {
            writer.write_all(&sheet_xml).map_err(|err| format!("写入主工作表失败: {err}"))?;
        } else if let Some(data) = book.entries.get(&name) {
            writer.write_all(data).map_err(|err| format!("写入 ZIP 内容失败 {name}: {err}"))?;
        }
    }
    writer.finish().map_err(|err| format!("完成 Rust 核对结果文件失败: {err}"))?;
    Ok(())
}

fn build_highlight_styles(styles_raw: &[u8], base_styles: &HashSet<u32>) -> Result<(Vec<u8>, HashMap<u32, u32>), String> {
    let mut reader = Reader::from_reader(Cursor::new(styles_raw));
    let mut buf = Vec::new();
    let mut in_cell_xfs = false;
    let mut xfs: Vec<XfTemplate> = Vec::new();

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"cellXfs" => in_cell_xfs = true,
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"cellXfs" => in_cell_xfs = false,
            Ok(Event::Empty(event)) if in_cell_xfs && local_name(event.name().as_ref()) == b"xf" => {
                xfs.push(XfTemplate { attrs: collect_attrs(&reader, &event) });
            }
            Ok(Event::Start(event)) if in_cell_xfs && local_name(event.name().as_ref()) == b"xf" => {
                xfs.push(XfTemplate { attrs: collect_attrs(&reader, &event) });
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(err) => return Err(format!("解析 styles.xml 失败: {err}")),
        }
    }
    if xfs.is_empty() {
        return Err("styles.xml 中未找到 cellXfs/xf".to_string());
    }

    let mut bases = base_styles.iter().copied().collect::<Vec<_>>();
    bases.sort_unstable();
    let original_xf_count = xfs.len() as u32;
    let style_map = bases
        .iter()
        .enumerate()
        .map(|(index, base)| (*base, original_xf_count + index as u32))
        .collect::<HashMap<_, _>>();
    let highlight_fill_id = count_fills(styles_raw)? as u32;

    let mut reader = Reader::from_reader(Cursor::new(styles_raw));
    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"fills" => {
                let elem = rewrite_count_attr(&reader, &event, count_fills(styles_raw)? + 1);
                writer.write_event(Event::Start(elem.borrow())).map_err(|err| format!("写入 fills 失败: {err}"))?;
            }
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"fills" => {
                write_highlight_fill(&mut writer)?;
                writer.write_event(Event::End(event.into_owned())).map_err(|err| format!("写入 fills 结束失败: {err}"))?;
            }
            Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"cellXfs" => {
                let elem = rewrite_count_attr(&reader, &event, xfs.len() + bases.len());
                writer.write_event(Event::Start(elem.borrow())).map_err(|err| format!("写入 cellXfs 失败: {err}"))?;
            }
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"cellXfs" => {
                for base in &bases {
                    let template = xfs.get(*base as usize).unwrap_or(&xfs[0]);
                    write_highlight_xf(&mut writer, template, highlight_fill_id)?;
                }
                writer.write_event(Event::End(event.into_owned())).map_err(|err| format!("写入 cellXfs 结束失败: {err}"))?;
            }
            Ok(Event::Eof) => break,
            Ok(event) => writer.write_event(event.into_owned()).map_err(|err| format!("重写 styles.xml 失败: {err}"))?,
            Err(err) => return Err(format!("重写 styles.xml 读取失败: {err}")),
        }
    }
    Ok((writer.into_inner(), style_map))
}

fn rewrite_sheet_styles(
    sheet_raw: &[u8],
    mismatch_cells: &HashSet<String>,
    original_styles: &HashMap<String, u32>,
    style_map: &HashMap<u32, u32>,
) -> Result<Vec<u8>, String> {
    let mut reader = Reader::from_reader(Cursor::new(sheet_raw));
    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"c" => {
                let event = rewrite_cell_style(&reader, &event, mismatch_cells, original_styles, style_map);
                writer.write_event(Event::Start(event.borrow())).map_err(|err| format!("写入 worksheet cell 失败: {err}"))?;
            }
            Ok(Event::Empty(event)) if local_name(event.name().as_ref()) == b"c" => {
                let event = rewrite_cell_style(&reader, &event, mismatch_cells, original_styles, style_map);
                writer.write_event(Event::Empty(event.borrow())).map_err(|err| format!("写入 worksheet empty cell 失败: {err}"))?;
            }
            Ok(Event::Eof) => break,
            Ok(event) => writer.write_event(event.into_owned()).map_err(|err| format!("重写 worksheet 失败: {err}"))?,
            Err(err) => return Err(format!("读取 worksheet 失败: {err}")),
        }
    }
    Ok(writer.into_inner())
}

fn count_fills(styles_raw: &[u8]) -> Result<usize, String> {
    let mut reader = Reader::from_reader(Cursor::new(styles_raw));
    let mut buf = Vec::new();
    let mut in_fills = false;
    let mut count = 0;
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"fills" => in_fills = true,
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"fills" => break,
            Ok(Event::Start(event)) | Ok(Event::Empty(event)) if in_fills && local_name(event.name().as_ref()) == b"fill" => count += 1,
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(err) => return Err(format!("统计 fills 失败: {err}")),
        }
    }
    Ok(count)
}

fn collect_attrs(reader: &Reader<Cursor<&[u8]>>, event: &BytesStart<'_>) -> Vec<(String, String)> {
    event
        .attributes()
        .with_checks(false)
        .flatten()
        .filter_map(|attr| {
            let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
            attr.decode_and_unescape_value(reader.decoder()).ok().map(|value| (key, value.into_owned()))
        })
        .collect()
}

fn rewrite_count_attr(reader: &Reader<Cursor<&[u8]>>, event: &BytesStart<'_>, count: usize) -> BytesStart<'static> {
    let name = String::from_utf8_lossy(event.name().as_ref()).to_string();
    let mut elem = BytesStart::new(name);
    let count_text = count.to_string();
    let mut has_count = false;
    for (key, value) in collect_attrs(reader, event) {
        if local_name(key.as_bytes()) == b"count" {
            elem.push_attribute((key.as_str(), count_text.as_str()));
            has_count = true;
        } else {
            elem.push_attribute((key.as_str(), value.as_str()));
        }
    }
    if !has_count {
        elem.push_attribute(("count", count_text.as_str()));
    }
    elem
}

fn rewrite_cell_style(
    reader: &Reader<Cursor<&[u8]>>,
    event: &BytesStart<'_>,
    mismatch_cells: &HashSet<String>,
    original_styles: &HashMap<String, u32>,
    style_map: &HashMap<u32, u32>,
) -> BytesStart<'static> {
    let ref_name = xml_attr(reader, event, b"r").unwrap_or_default().to_ascii_uppercase();
    if !mismatch_cells.contains(&ref_name) {
        return event.to_owned();
    }
    let base_style = original_styles.get(&ref_name).copied().unwrap_or(0);
    let new_style = style_map.get(&base_style).copied().unwrap_or(base_style);
    let name = String::from_utf8_lossy(event.name().as_ref()).to_string();
    let mut elem = BytesStart::new(name);
    let style_text = new_style.to_string();
    let mut has_style = false;
    for (key, value) in collect_attrs(reader, event) {
        if local_name(key.as_bytes()) == b"s" {
            elem.push_attribute((key.as_str(), style_text.as_str()));
            has_style = true;
        } else {
            elem.push_attribute((key.as_str(), value.as_str()));
        }
    }
    if !has_style {
        elem.push_attribute(("s", style_text.as_str()));
    }
    elem
}

fn write_highlight_fill(writer: &mut Writer<Vec<u8>>) -> Result<(), String> {
    writer.write_event(Event::Start(BytesStart::new("fill"))).map_err(|err| format!("写入高亮 fill 失败: {err}"))?;
    let mut pattern = BytesStart::new("patternFill");
    pattern.push_attribute(("patternType", "solid"));
    writer.write_event(Event::Start(pattern)).map_err(|err| format!("写入高亮 patternFill 失败: {err}"))?;
    let mut fg = BytesStart::new("fgColor");
    fg.push_attribute(("rgb", "FFFFFF00"));
    writer.write_event(Event::Empty(fg)).map_err(|err| format!("写入高亮 fgColor 失败: {err}"))?;
    let mut bg = BytesStart::new("bgColor");
    bg.push_attribute(("indexed", "64"));
    writer.write_event(Event::Empty(bg)).map_err(|err| format!("写入高亮 bgColor 失败: {err}"))?;
    writer.write_event(Event::End(BytesStart::new("patternFill").to_end())).map_err(|err| format!("结束高亮 patternFill 失败: {err}"))?;
    writer.write_event(Event::End(BytesStart::new("fill").to_end())).map_err(|err| format!("结束高亮 fill 失败: {err}"))?;
    Ok(())
}

fn write_highlight_xf(writer: &mut Writer<Vec<u8>>, template: &XfTemplate, fill_id: u32) -> Result<(), String> {
    let mut elem = BytesStart::new("xf");
    let fill_text = fill_id.to_string();
    let mut has_fill = false;
    let mut has_apply = false;
    for (key, value) in &template.attrs {
        if local_name(key.as_bytes()) == b"fillId" {
            elem.push_attribute((key.as_str(), fill_text.as_str()));
            has_fill = true;
        } else if local_name(key.as_bytes()) == b"applyFill" {
            elem.push_attribute((key.as_str(), "1"));
            has_apply = true;
        } else {
            elem.push_attribute((key.as_str(), value.as_str()));
        }
    }
    if !has_fill {
        elem.push_attribute(("fillId", fill_text.as_str()));
    }
    if !has_apply {
        elem.push_attribute(("applyFill", "1"));
    }
    writer.write_event(Event::Empty(elem)).map_err(|err| format!("写入高亮 xf 失败: {err}"))?;
    Ok(())
}

fn parse_rules(template: &CheckTemplate) -> Result<Vec<ParsedRule>, String> {
    template
        .rules
        .iter()
        .map(|rule| {
            let (main_column, main_start_row) = parse_main_range(&rule.main_range)?;
            validate_table_b_expression(&rule.table_b_cell)?;
            Ok(ParsedRule {
                field_name: rule.field_name.clone(),
                table_b_cell: rule.table_b_cell.trim().to_ascii_uppercase(),
                compare_type: normalize_compare_type(&rule.compare_type),
                main_column,
                main_start_row,
            })
        })
        .collect()
}

fn normalize_compare_type(compare_type: &str) -> String {
    match compare_type.trim().to_ascii_lowercase().as_str() {
        "number" => "number".to_string(),
        "position" => "position".to_string(),
        _ => "text".to_string(),
    }
}

fn parse_main_range(range: &str) -> Result<(String, u32), String> {
    let text = range.trim().to_ascii_uppercase();
    let (start_ref, end_ref) = text
        .split_once('-')
        .ok_or_else(|| format!("主表范围格式不正确: {range}，示例应为 F7-Fn"))?;
    let (start_col, row) = split_cell_ref(start_ref).map_err(|_| format!("主表范围格式不正确: {range}，示例应为 F7-Fn"))?;
    let (end_col, end_row) = split_range_end(end_ref).map_err(|_| format!("主表范围格式不正确: {range}，示例应为 F7-Fn"))?;
    if end_row != "N" {
        return Err(format!("主表范围格式不正确: {range}，示例应为 F7-Fn"));
    }
    if start_col != end_col {
        return Err(format!("当前仅支持同一列的范围模板: {range}"));
    }
    Ok((start_col, row))
}

fn validate_table_b_expression(expression: &str) -> Result<(), String> {
    let text = expression.trim().to_ascii_uppercase();
    if text.starts_with("SUM(") {
        if !text.ends_with(')') {
            return Err(format!("考勤表表达式格式不正确: {expression}"));
        }
        for token in parse_sum_tokens(&text)? {
            if let Some((start, end)) = token.split_once(':') {
                let (start_col, start_row) = split_cell_ref(start)?;
                let (end_col, end_row_text) = split_range_end(end)?;
                if col_to_num(&start_col) > col_to_num(&end_col) {
                    return Err(format!("SUM 范围起始列不能大于结束列: {token}"));
                }
                let end_row = if end_row_text == "N" {
                    u32::MAX
                } else {
                    end_row_text.parse::<u32>().map_err(|_| format!("范围终点行号无效: {end}"))?
                };
                if start_row > end_row {
                    return Err(format!("SUM 范围起始行不能大于结束行: {token}"));
                }
            } else {
                split_cell_ref(&token)?;
            }
        }
        return Ok(());
    }
    split_cell_ref(&text)?;
    Ok(())
}

fn parse_sum_tokens(expression: &str) -> Result<Vec<String>, String> {
    let text = expression.trim().to_ascii_uppercase();
    if !text.starts_with("SUM(") || !text.ends_with(')') {
        return Err(format!("考勤表表达式格式不正确: {expression}"));
    }
    let inner = text[4..text.len() - 1].trim();
    if inner.is_empty() {
        return Err(format!("SUM 表达式不能为空: {expression}"));
    }
    Ok(inner.split(',').map(|item| item.trim().to_string()).filter(|item| !item.is_empty()).collect())
}

fn resolve_table_b_value(sheet: &Sheet, expression: &str) -> Result<String, String> {
    let text = expression.trim().to_ascii_uppercase();
    if !text.starts_with("SUM(") {
        return Ok(sheet.get_value(&text).to_string());
    }
    let mut total = 0.0;
    for token in parse_sum_tokens(&text)? {
        let refs = if let Some((start, end)) = token.split_once(':') {
            iter_range_cells(start, end, sheet)?
        } else {
            vec![token]
        };
        for ref_name in refs {
            total += parse_number(sheet.get_value(&ref_name)).unwrap_or(0.0);
        }
    }
    Ok(number_to_text(total))
}

fn locate_data_offsets(sheet: &Sheet, template: &CheckTemplate, rules: &[ParsedRule]) -> Vec<u32> {
    let mut offsets = Vec::new();
    let mut current = 0;
    loop {
        let mut refs = vec![format!("{}{}", template.number_column.to_ascii_uppercase(), template.start_row + current)];
        refs.extend(rules.iter().map(|rule| format!("{}{}", rule.main_column, rule.main_start_row + current)));
        if refs.iter().all(|ref_name| sheet.get_value(ref_name).is_empty()) {
            break;
        }
        offsets.push(current);
        current += 1;
    }
    offsets
}

fn compare_row(
    offset: u32,
    display_row: u32,
    table_a_sheet: &Sheet,
    table_b_sheet: Option<&Sheet>,
    table_b_path: Option<&Path>,
    rules: &[ParsedRule],
    compare_context: &CompareContext,
) -> Result<Vec<Mismatch>, String> {
    let file_name = table_b_path
        .and_then(|path| path.file_name())
        .and_then(|name| name.to_str())
        .unwrap_or("未找到匹配考勤表")
        .to_string();
    let mut mismatches = Vec::new();
    for rule in rules {
        let table_a_cell = format!("{}{}", rule.main_column, rule.main_start_row + offset);
        let left = table_a_sheet.get_value(&table_a_cell).to_string();
        let right = match table_b_sheet {
            Some(sheet) => resolve_table_b_value(sheet, &rule.table_b_cell)?,
            None => String::new(),
        };
        let matches = if table_b_sheet.is_none() {
            false
        } else if rule.compare_type == "number" {
            (parse_number(&left).unwrap_or(0.0) - parse_number(&right).unwrap_or(0.0)).abs() <= NUMERIC_TOLERANCE
        } else if rule.compare_type == "position" {
            positions_mean_same(&left, &right, compare_context)
        } else {
            normalize_text(&left) == normalize_text(&right)
        };
        if !matches {
            mismatches.push(Mismatch {
                row_num: display_row,
                table_a_cell,
                field_name: rule.field_name.clone(),
                table_a_value: to_display(&left),
                table_b_value: to_display(&right),
                table_b_file: file_name.clone(),
            });
        }
    }
    Ok(mismatches)
}

fn build_table_b_index(folder: &Path) -> Result<TableBIndex, String> {
    let mut grouped: BTreeMap<i32, Vec<PathBuf>> = BTreeMap::new();
    let mut warnings = Vec::new();
    for entry in fs::read_dir(folder).map_err(|err| format!("读取考勤表目录失败: {err}"))? {
        let path = entry.map_err(|err| format!("读取目录项失败: {err}"))?.path();
        if path.is_dir() {
            continue;
        }
        let ext = path.extension().and_then(|item| item.to_str()).unwrap_or("").to_ascii_lowercase();
        if ext != "xlsx" && ext != "xlsm" {
            continue;
        }
        let stem = path.file_stem().and_then(|item| item.to_str()).unwrap_or("");
        if let Some(number) = parse_file_index(stem) {
            grouped.entry(number).or_default().push(path);
        } else {
            let name = path.file_name().and_then(|item| item.to_str()).unwrap_or("").to_string();
            warnings.push(format!("忽略未匹配编号规则的文件: {name}"));
        }
    }
    let mut files = HashMap::new();
    for (number, mut candidates) in grouped {
        candidates.sort_by_key(|path| {
            let stem = path.file_stem().and_then(|item| item.to_str()).unwrap_or("").to_ascii_lowercase();
            let name = path.file_name().and_then(|item| item.to_str()).unwrap_or("").to_ascii_lowercase();
            (stem.contains("副本") || stem.contains("copy"), name)
        });
        if candidates.len() > 1 {
            let joined = candidates
                .iter()
                .filter_map(|path| path.file_name().and_then(|item| item.to_str()))
                .collect::<Vec<_>>()
                .join(", ");
            let selected = candidates[0].file_name().and_then(|item| item.to_str()).unwrap_or("");
            warnings.push(format!("No.{number} 存在多个候选文件，已使用 {selected}: {joined}"));
        }
        files.insert(number, candidates[0].clone());
    }
    Ok(TableBIndex { files, warnings })
}

fn parse_file_index(stem: &str) -> Option<i32> {
    let text = stem.trim_start();
    let dot_index = text.find('.')?;
    let number_text = &text[..dot_index];
    if number_text.is_empty() || !number_text.chars().all(|ch| ch.is_ascii_digit()) {
        return None;
    }
    number_text.parse::<i32>().ok()
}

fn choose_timesheet_sheet(sheets: Vec<WorkbookSheet>, rules: &[ParsedRule]) -> Result<WorkbookSheet, String> {
    let score_refs = score_refs_from_rules(rules);
    sheets
        .into_iter()
        .max_by_key(|item| {
            let template_score = score_refs.iter().filter(|ref_name| !item.sheet.get_value(ref_name).is_empty()).count();
            let fallback_score = ["A3", "B6", "C3", "H9", "I9", "J9"]
                .iter()
                .filter(|ref_name| !item.sheet.get_value(ref_name).is_empty())
                .count();
            template_score * 10 + fallback_score
        })
        .ok_or_else(|| "Rust 核对引擎未读取到考勤表工作表".to_string())
}

fn score_refs_from_rules(rules: &[ParsedRule]) -> Vec<String> {
    let mut refs = Vec::new();
    for rule in rules {
        let expression = rule.table_b_cell.trim().to_ascii_uppercase();
        if expression.starts_with("SUM(") {
            if let Ok(tokens) = parse_sum_tokens(&expression) {
                for token in tokens {
                    if let Some((start, _)) = token.split_once(':') {
                        refs.push(start.to_string());
                    } else {
                        refs.push(token);
                    }
                }
            }
        } else {
            refs.push(expression);
        }
    }
    refs.sort();
    refs.dedup();
    refs
}

fn iter_range_cells(start_ref: &str, end_ref: &str, sheet: &Sheet) -> Result<Vec<String>, String> {
    let (start_col, start_row) = split_cell_ref(start_ref)?;
    let (end_col, end_row_text) = split_range_end(end_ref)?;
    let end_row = if end_row_text == "N" {
        infer_data_end_row(sheet, start_row)
    } else {
        end_row_text.parse::<u32>().map_err(|_| format!("范围终点行号无效: {end_ref}"))?
    };
    let mut refs = Vec::new();
    for col_num in col_to_num(&start_col)..=col_to_num(&end_col) {
        let col = num_to_col(col_num);
        for row in start_row..=end_row {
            refs.push(format!("{col}{row}"));
        }
    }
    Ok(refs)
}

fn infer_data_end_row(sheet: &Sheet, start_row: u32) -> u32 {
    let mut current = start_row;
    let mut last_seen = start_row;
    loop {
        let anchor = sheet.get_value(&format!("A{current}"));
        if anchor.is_empty() {
            break;
        }
        let display = to_display(anchor);
        if display.contains("修正上月加班") || display.to_ascii_uppercase().contains("FIX OT") {
            last_seen = current;
            break;
        }
        if !is_data_anchor(&display) {
            break;
        }
        last_seen = current;
        current += 1;
        if current > 2000 {
            break;
        }
    }
    last_seen
}

fn is_data_anchor(value: &str) -> bool {
    let text = value.trim();
    if text.is_empty() {
        return false;
    }
    if parse_number(text).is_ok() {
        return true;
    }
    is_month_day_anchor(text)
}

fn split_cell_ref(ref_name: &str) -> Result<(String, u32), String> {
    let text = ref_name.trim().to_ascii_uppercase();
    let split_at = text.find(|ch: char| ch.is_ascii_digit()).ok_or_else(|| format!("单元格坐标格式不正确: {ref_name}"))?;
    let (col, row_text) = text.split_at(split_at);
    if col.is_empty()
        || row_text.is_empty()
        || !col.chars().all(|ch| ch.is_ascii_uppercase())
        || !row_text.chars().all(|ch| ch.is_ascii_digit())
    {
        return Err(format!("单元格坐标格式不正确: {ref_name}"));
    }
    Ok((col.to_string(), row_text.parse::<u32>().map_err(|_| format!("单元格行号无效: {ref_name}"))?))
}

fn split_range_end(ref_name: &str) -> Result<(String, String), String> {
    let text = ref_name.trim().to_ascii_uppercase();
    let split_at = text
        .find(|ch: char| ch.is_ascii_digit() || ch == 'N')
        .ok_or_else(|| format!("范围终点格式不正确: {ref_name}"))?;
    let (col, row_text) = text.split_at(split_at);
    if col.is_empty()
        || row_text.is_empty()
        || !col.chars().all(|ch| ch.is_ascii_uppercase())
        || !(row_text == "N" || row_text.chars().all(|ch| ch.is_ascii_digit()))
    {
        return Err(format!("范围终点格式不正确: {ref_name}"));
    }
    Ok((col.to_string(), row_text.to_string()))
}

fn col_to_num(col: &str) -> u32 {
    col.bytes().fold(0, |acc, item| acc * 26 + (item as u32 - b'A' as u32 + 1))
}

fn num_to_col(mut num: u32) -> String {
    let mut chars = Vec::new();
    while num > 0 {
        num -= 1;
        chars.push((b'A' + (num % 26) as u8) as char);
        num /= 26;
    }
    chars.iter().rev().collect()
}

fn parse_number(value: &str) -> Result<f64, String> {
    let text = value.trim().replace(',', "");
    if text.is_empty() {
        return Ok(0.0);
    }
    text.parse::<f64>().map_err(|_| format!("无法解析数字: {value:?}"))
}

fn number_to_text(value: f64) -> String {
    if value.fract().abs() <= NUMERIC_TOLERANCE {
        format!("{}", value.round() as i64)
    } else {
        let mut text = format!("{value:.10}");
        while text.contains('.') && text.ends_with('0') {
            text.pop();
        }
        if text.ends_with('.') {
            text.pop();
        }
        text
    }
}

fn normalize_text(value: &str) -> String {
    compact_comparison_spacing(&collapse_whitespace(&value.replace('\u{a0}', " "))).to_lowercase()
}

fn to_display(value: &str) -> String {
    collapse_whitespace(&value.replace(['\n', '\r'], " "))
}

fn collapse_whitespace(value: &str) -> String {
    let mut output = String::new();
    let mut previous_was_space = false;
    for ch in value.trim().chars() {
        if ch.is_whitespace() {
            if !previous_was_space && !output.is_empty() {
                output.push(' ');
            }
            previous_was_space = true;
        } else {
            output.push(ch);
            previous_was_space = false;
        }
    }
    output
}

fn compact_comparison_spacing(value: &str) -> String {
    let chars = value.chars().collect::<Vec<_>>();
    let mut output = String::new();
    for (index, ch) in chars.iter().enumerate() {
        if ch.is_whitespace() {
            let previous = output.chars().next_back();
            let next = chars[index + 1..].iter().find(|item| !item.is_whitespace()).copied();
            if previous.is_some_and(is_tight_punctuation) || next.is_some_and(is_tight_punctuation) {
                continue;
            }
        }
        output.push(*ch);
    }
    output
}

fn is_tight_punctuation(ch: char) -> bool {
    matches!(ch, '(' | ')' | '-' | '/')
}

fn is_month_day_anchor(value: &str) -> bool {
    let mut parts = value.split(|ch: char| matches!(ch, '-' | '/' | '.' | ' ')).filter(|part| !part.is_empty());
    let Some(day) = parts.next() else {
        return false;
    };
    let Some(month) = parts.next() else {
        return false;
    };
    if parts.next().is_some() {
        return false;
    }
    (1..=2).contains(&day.len())
        && day.chars().all(|ch| ch.is_ascii_digit())
        && (3..=9).contains(&month.len())
        && month.chars().all(|ch| ch.is_ascii_alphabetic())
}

fn load_position_aliases(path: Option<&str>) -> HashMap<String, String> {
    let mut aliases = HashMap::new();
    let Some(path) = path else {
        return aliases;
    };
    let Ok(raw) = fs::read_to_string(path) else {
        return aliases;
    };
    let Ok(data) = serde_json::from_str::<HashMap<String, String>>(&raw) else {
        return aliases;
    };
    for (key, value) in data {
        let normalized_key = normalize_text(&key);
        let normalized_value = normalize_text(&value);
        if !normalized_key.is_empty() && !normalized_value.is_empty() {
            aliases.insert(normalized_key, normalized_value);
        }
    }
    aliases
}

#[derive(Default, Deserialize)]
struct PositionRulesFile {
    #[serde(default)]
    token_aliases: HashMap<String, Vec<String>>,
    #[serde(default)]
    optional_tokens: Vec<String>,
}

struct PositionRules {
    token_aliases: HashMap<String, Vec<String>>,
    optional_tokens: HashSet<String>,
}

fn load_position_rules(path: Option<&str>) -> PositionRules {
    let mut rules = PositionRules {
        token_aliases: HashMap::new(),
        optional_tokens: HashSet::new(),
    };
    let Some(path) = path else {
        return rules;
    };
    let Ok(raw) = fs::read_to_string(path) else {
        return rules;
    };
    let Ok(data) = serde_json::from_str::<PositionRulesFile>(&raw) else {
        return rules;
    };
    for (key, value) in data.token_aliases {
        let key = normalize_position_token(&key);
        let tokens = value
            .into_iter()
            .map(|item| normalize_position_token(&item))
            .filter(|item| !item.is_empty())
            .collect::<Vec<_>>();
        if !key.is_empty() && !tokens.is_empty() {
            rules.token_aliases.insert(key, tokens);
        }
    }
    rules.optional_tokens = data
        .optional_tokens
        .into_iter()
        .map(|item| normalize_position_token(&item))
        .filter(|item| !item.is_empty())
        .collect();
    rules
}

fn positions_mean_same(left: &str, right: &str, context: &CompareContext) -> bool {
    let left_tokens = canonical_position_tokens(left, context);
    let right_tokens = canonical_position_tokens(right, context);
    if left_tokens == right_tokens {
        return true;
    }
    if left_tokens.is_empty() || right_tokens.is_empty() {
        return left_tokens == right_tokens;
    }
    let left_set = left_tokens.iter().cloned().collect::<HashSet<_>>();
    let right_set = right_tokens.iter().cloned().collect::<HashSet<_>>();
    if left_set == right_set {
        return true;
    }
    let (smaller, larger) = if left_set.len() <= right_set.len() {
        (&left_set, &right_set)
    } else {
        (&right_set, &left_set)
    };
    let extra_tokens = larger.difference(smaller).cloned().collect::<HashSet<_>>();
    smaller.len() >= 2 && smaller.is_subset(larger) && extra_tokens.iter().all(|item| context.position_optional_tokens.contains(item))
}

fn canonical_position_tokens(value: &str, context: &CompareContext) -> Vec<String> {
    let mut text = normalize_text(value);
    if let Some(alias) = context.position_aliases.get(&text) {
        text = alias.clone();
    }
    let raw_tokens = position_raw_tokens(&text);
    let mut expanded = Vec::new();
    let mut index = 0;
    while index < raw_tokens.len() {
        if index + 1 < raw_tokens.len() && raw_tokens[index] == "e" && raw_tokens[index + 1] == "i" {
            expanded.push("ei".to_string());
            index += 2;
            continue;
        }
        let token = &raw_tokens[index];
        if let Some(alias_tokens) = context.position_token_aliases.get(token) {
            expanded.extend(alias_tokens.iter().cloned());
        } else {
            expanded.push(token.clone());
        }
        index += 1;
    }
    let mut tokens = expanded;
    tokens.sort();
    tokens.dedup();
    tokens
}

fn position_raw_tokens(value: &str) -> Vec<String> {
    value
        .split(|ch: char| !ch.is_ascii_alphanumeric())
        .map(normalize_position_token)
        .filter(|item| !item.is_empty())
        .collect()
}

fn normalize_position_token(value: &str) -> String {
    value
        .chars()
        .filter(|ch| ch.is_ascii_alphanumeric())
        .flat_map(|ch| ch.to_lowercase())
        .collect()
}

fn next_output_path(table_a_path: &Path) -> PathBuf {
    let stem = table_a_path.file_stem().and_then(|item| item.to_str()).unwrap_or("result");
    let ext = table_a_path.extension().and_then(|item| item.to_str()).unwrap_or("xlsx");
    table_a_path.with_file_name(format!("{stem}_Rust核对结果.{ext}"))
}

fn next_report_path(output_path: &Path) -> PathBuf {
    let stem = output_path.file_stem().and_then(|item| item.to_str()).unwrap_or("result");
    output_path.with_file_name(format!("{stem}_核对报告.txt"))
}

fn create_report(report_path: &Path, table_a_path: &Path, output_path: &Path, warnings: &[String], mismatches: &[Mismatch]) -> Result<(), String> {
    let mut lines = vec![
        format!("主表: {}", table_a_path.display()),
        format!("结果文件: {}", output_path.display()),
        format!("不一致数量: {}", mismatches.len()),
        String::new(),
    ];
    if !warnings.is_empty() {
        lines.push("提示:".to_string());
        for warning in warnings {
            lines.push(format!("- {warning}"));
        }
        lines.push(String::new());
    }
    if mismatches.is_empty() {
        lines.push("未发现不一致。".to_string());
    } else {
        lines.push("明细:".to_string());
        for item in mismatches {
            lines.push(format!(
                "- 第 {} 行 {}({}) | 主表='{}' | 考勤表='{}' | 文件={}",
                item.row_num, item.table_a_cell, item.field_name, item.table_a_value, item.table_b_value, item.table_b_file
            ));
        }
    }
    fs::write(report_path, lines.join("\n")).map_err(|err| format!("写入 Rust 核对报告失败: {err}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn position_compare_uses_alias_file_and_optional_tokens() {
        let alias_path = std::env::temp_dir().join("rust-check-position-aliases.json");
        let rules_path = std::env::temp_dir().join("rust-check-position-rules.json");
        fs::write(
            &alias_path,
            r#"{
                "lead welder": "welder",
                "job performer mec.": "job performer mechanical"
            }"#,
        )
        .unwrap();
        fs::write(
            &rules_path,
            r#"{
                "token_aliases": {
                    "qci": ["qc", "inspector"],
                    "mec": ["mechanical"]
                },
                "optional_tokens": ["mechanical", "inspector", "ei"]
            }"#,
        )
        .unwrap();
        let position_rules = load_position_rules(Some(&rules_path.to_string_lossy()));
        let context = CompareContext {
            position_aliases: load_position_aliases(Some(&alias_path.to_string_lossy())),
            position_token_aliases: position_rules.token_aliases,
            position_optional_tokens: position_rules.optional_tokens,
        };
        assert!(positions_mean_same("Lead Welder", "Welder", &context));
        assert!(positions_mean_same("Job Performer Mec.", "Job Performer (Mechanical)", &context));
        assert!(positions_mean_same("QC E&I Inspector", "Qc Inspector E&I", &context));
        assert!(!positions_mean_same("Welder", "Rigger", &context));
        let _ = fs::remove_file(alias_path);
        let _ = fs::remove_file(rules_path);
    }

    #[test]
    fn sum_expression_supports_dynamic_n_and_fix_ot_row() {
        let mut cells = HashMap::new();
        cells.insert("A10".to_string(), "1-Jun".to_string());
        cells.insert("A11".to_string(), "2-Jun".to_string());
        cells.insert("A12".to_string(), "修正上月加班时长".to_string());
        cells.insert("G10".to_string(), "8".to_string());
        cells.insert("G11".to_string(), "7.5".to_string());
        cells.insert("G12".to_string(), "2".to_string());
        let sheet = Sheet {
            cells,
            styles: HashMap::new(),
        };
        assert_eq!(resolve_table_b_value(&sheet, "SUM(G10:GN)").unwrap(), "17.5");
    }

    #[test]
    fn template_expression_validation_rejects_reversed_ranges() {
        assert!(validate_table_b_expression("SUM(G10:HN,I10:IN)").is_ok());
        assert!(validate_table_b_expression("SUM(H10:GN)").is_err());
        assert!(validate_table_b_expression("SUM(G20:G10)").is_err());
    }

    #[test]
    fn parser_handles_formula_cached_values_and_inline_strings() {
        let raw = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row r="1">
                    <c r="A1"><f>SUM(A2:A3)</f><v>12.5</v></c>
                    <c r="B1" t="inlineStr"><is><t>Hello</t></is></c>
                    <c r="C1" t="s"><v>0</v></c>
                </row>
            </sheetData>
        </worksheet>"#;
        let sheet = parse_sheet(raw, &["Shared".to_string()]).unwrap();
        assert_eq!(sheet.get_value("A1"), "12.5");
        assert_eq!(sheet.get_value("B1"), "Hello");
        assert_eq!(sheet.get_value("C1"), "Shared");
    }

    #[test]
    fn text_normalization_matches_python_spacing_rules() {
        assert_eq!(normalize_text(" Job Performer ( Mechanical ) "), "job performer(mechanical)");
        assert_eq!(normalize_text("A / B - C"), "a/b-c");
        assert_eq!(to_display("  A\n  B\tC  "), "A B C");
    }

    #[test]
    fn sheet_selection_scores_template_refs_first() {
        let rules = vec![ParsedRule {
            field_name: "字段".to_string(),
            table_b_cell: "Z9".to_string(),
            compare_type: "text".to_string(),
            main_column: "A".to_string(),
            main_start_row: 1,
        }];
        let fallback_sheet = WorkbookSheet {
            part_name: "xl/worksheets/sheet1.xml".to_string(),
            sheet: Sheet {
                cells: HashMap::from([
                    ("A3".to_string(), "x".to_string()),
                    ("B6".to_string(), "x".to_string()),
                    ("C3".to_string(), "x".to_string()),
                ]),
                styles: HashMap::new(),
            },
        };
        let template_sheet = WorkbookSheet {
            part_name: "xl/worksheets/sheet2.xml".to_string(),
            sheet: Sheet {
                cells: HashMap::from([("Z9".to_string(), "x".to_string())]),
                styles: HashMap::new(),
            },
        };
        let selected = choose_timesheet_sheet(vec![fallback_sheet, template_sheet], &rules).unwrap();
        assert_eq!(selected.part_name, "xl/worksheets/sheet2.xml");
    }

    #[test]
    fn rust_check_real_files_when_available() {
        let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).parent().unwrap().to_path_buf();
        let table_a = root.join("6月岗位外包工资表-26.7.3.xlsx");
        let table_bs = root.join("考勤表-编辑版-307人");
        let template_path = root.join("外包模板-WEP-2026.5.15.json");
        if !(table_a.exists() && table_bs.exists() && template_path.exists()) {
            return;
        }
        let template: CheckTemplate = serde_json::from_str(&fs::read_to_string(&template_path).unwrap()).unwrap();
        let output = std::env::temp_dir().join("rust-check-real-result.xlsx");
        let report = std::env::temp_dir().join("rust-check-real-result_核对报告.txt");
        let _ = fs::remove_file(&output);
        let _ = fs::remove_file(&report);
        let result = run_check(CheckPayload {
            table_a_path: table_a.to_string_lossy().to_string(),
            table_bs_folder: table_bs.to_string_lossy().to_string(),
            output_path: Some(output.to_string_lossy().to_string()),
            position_aliases_path: Some(root.join("position_aliases.json").to_string_lossy().to_string()),
            position_rules_path: Some(root.join("position_rules.json").to_string_lossy().to_string()),
            template,
        })
        .unwrap();
        if result.mismatch_count != 1 {
            eprintln!("mismatch_count={}", result.mismatch_count);
            for item in result.mismatches.iter().take(12) {
                eprintln!(
                    "{} {} {} | left='{}' right='{}' file={}",
                    item.row_num, item.table_a_cell, item.field_name, item.table_a_value, item.table_b_value, item.table_b_file
                );
            }
        }
        assert_eq!(result.mismatch_count, 1);
        assert_eq!(result.progress.current, result.progress.total);
        assert!(PathBuf::from(result.output_path).exists());
        assert!(PathBuf::from(result.report_path).exists());
        let mismatch_cell = result.mismatches[0].table_a_cell.clone();
        let original = Workbook::open(&table_a).unwrap();
        let highlighted = Workbook::open(&output).unwrap();
        let original_style = original.sheets[0].sheet.styles.get(&mismatch_cell).copied().unwrap_or(0);
        let highlighted_style = highlighted.sheets[0].sheet.styles.get(&mismatch_cell).copied().unwrap_or(0);
        assert_ne!(original_style, highlighted_style);
        let _ = fs::remove_file(output);
        let _ = fs::remove_file(report);
    }
}
