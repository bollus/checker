use quick_xml::events::{BytesStart, BytesText, Event};
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;
use rusttype::{point, Font, Scale};
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet};
use std::fs::{self, File};
use std::io::{Cursor, Read, Write};
use std::path::{Path, PathBuf};
use zip::{write::SimpleFileOptions, CompressionMethod, ZipArchive, ZipWriter};

const DEFAULT_MORNING_START: &str = "06:00";
const DEFAULT_MORNING_END: &str = "12:00";
const DEFAULT_AFTERNOON_START: &str = "14:00";
const DEFAULT_AFTERNOON_END: &str = "18:00";
const DEFAULT_NORMAL_HOURS: f64 = 10.0;
const SIGNATURE_MEDIA_WIDTH: usize = 900;
const SIGNATURE_MEDIA_HEIGHT: usize = 260;
const EMU_PER_PIXEL: i32 = 9_525;
const REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DRAWING_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const OFFICE_REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

#[derive(Debug, Deserialize)]
pub struct GeneratePayload {
    pub table_c_path: String,
    pub template_b_path: String,
    pub output_dir: Option<String>,
    #[serde(default)]
    pub count_holidays: bool,
    #[serde(default = "default_signature_scale")]
    pub signature_scale: i32,
    #[serde(default = "default_morning_start")]
    pub morning_start: String,
    #[serde(default = "default_morning_end")]
    pub morning_end: String,
    #[serde(default = "default_afternoon_start")]
    pub afternoon_start: String,
    #[serde(default = "default_afternoon_end")]
    pub afternoon_end: String,
    #[serde(default = "default_normal_hours_text")]
    pub normal_hours: String,
    #[serde(default)]
    pub signature_font_path: Option<String>,
    #[serde(default)]
    pub insert_manager_signature: bool,
    #[serde(default)]
    pub manager_signature_dir: Option<String>,
}

fn default_signature_scale() -> i32 { 100 }
fn default_morning_start() -> String { DEFAULT_MORNING_START.to_string() }
fn default_morning_end() -> String { DEFAULT_MORNING_END.to_string() }
fn default_afternoon_start() -> String { DEFAULT_AFTERNOON_START.to_string() }
fn default_afternoon_end() -> String { DEFAULT_AFTERNOON_END.to_string() }
fn default_normal_hours_text() -> String { DEFAULT_NORMAL_HOURS.to_string() }

#[derive(Debug, Serialize)]
pub struct GenerateResult {
    pub output_dir: String,
    pub report_path: String,
    pub generated_count: usize,
    pub generated_files: Vec<String>,
    pub warnings: Vec<String>,
    pub progress: ProgressSnapshot,
}

#[derive(Debug, Serialize)]
pub struct ProgressSnapshot {
    pub current: usize,
    pub total: usize,
    pub message: String,
}

#[derive(Clone)]
struct Workbook {
    entries: HashMap<String, Vec<u8>>,
    infos: HashMap<String, ZipMeta>,
    sheets: Vec<WorkbookSheet>,
}

#[derive(Clone)]
struct ZipMeta {
    compression: CompressionMethod,
    unix_mode: Option<u32>,
}

#[derive(Clone)]
struct WorkbookSheet {
    name: String,
    part_name: String,
    sheet: Sheet,
}

#[derive(Clone, Default)]
struct Sheet {
    cells: HashMap<String, CellData>,
}

#[derive(Clone, Default)]
struct CellData {
    value: String,
    style: Option<u32>,
}

#[derive(Clone)]
enum CellValue {
    Blank,
    Number(f64),
    Text(String),
}

#[derive(Clone)]
struct Schedule {
    morning_start: f64,
    morning_end: f64,
    afternoon_start: f64,
    afternoon_end: f64,
    normal_hours: f64,
}

#[derive(Clone)]
struct Employee {
    no: i32,
    employee_no: String,
    project: String,
    passport: String,
    crew_group: String,
    name: String,
    position: String,
    days: HashMap<i32, DayEntry>,
    correction_nwh: f64,
    correction_normal_ot: f64,
    correction_weekend_ot: f64,
    correction_holiday_ot: f64,
}

#[derive(Clone, Debug, PartialEq)]
enum DayEntry {
    Blank,
    Hours(f64),
    Leave(String),
}

#[derive(Clone)]
struct OvertimeEntry {
    day: i32,
    start: Option<f64>,
    end: Option<f64>,
    normal_hours: f64,
    weekend_hours: f64,
    holiday_hours: f64,
}

struct ManagerSignatureIndex {
    files: HashMap<String, PathBuf>,
}

#[derive(Clone, Copy)]
enum SignaturePlacement {
    Stretch,
    Contain,
}

#[derive(Clone, Copy)]
struct AnchorMarker {
    col: i32,
    row: i32,
    col_off: i32,
    row_off: i32,
}

#[derive(Clone, Copy)]
struct AnchorBounds {
    from: AnchorMarker,
    to: AnchorMarker,
}

struct SheetMetrics {
    default_col_width: f64,
    default_row_height: f64,
    col_widths: Vec<(i32, i32, f64)>,
    row_heights: HashMap<i32, f64>,
}

#[derive(Default)]
struct DayFillResult {
    payable: bool,
    work_hours: f64,
    work_ot: f64,
    rest_hours: f64,
    holiday_hours: f64,
}

pub fn run_generate(payload: GeneratePayload) -> Result<GenerateResult, String> {
    let table_c_path = PathBuf::from(&payload.table_c_path);
    let template_path = PathBuf::from(&payload.template_b_path);
    if !table_c_path.is_file() {
        return Err(format!("Rust 生成引擎未找到汇总表: {}", payload.table_c_path));
    }
    if !template_path.is_file() {
        return Err(format!("Rust 生成引擎未找到考勤表模板: {}", payload.template_b_path));
    }
    if payload.signature_scale < 30 || payload.signature_scale > 200 {
        return Err("签名大小必须在 30% 到 200% 之间".to_string());
    }
    let signature_font_path = resolve_signature_font(payload.signature_font_path.as_deref())?;
    let manager_signatures = if payload.insert_manager_signature {
        Some(build_manager_signature_index(payload.manager_signature_dir.as_deref())?)
    } else {
        None
    };
    let schedule = Schedule {
        morning_start: parse_time(&payload.morning_start)?,
        morning_end: parse_time(&payload.morning_end)?,
        afternoon_start: parse_time(&payload.afternoon_start)?,
        afternoon_end: parse_time(&payload.afternoon_end)?,
        normal_hours: payload.normal_hours.trim().parse::<f64>().map_err(|_| "常规工作小时数无效".to_string())?,
    };
    if !(schedule.morning_start < schedule.morning_end
        && schedule.morning_end <= schedule.afternoon_start
        && schedule.afternoon_start < schedule.afternoon_end
        && schedule.normal_hours > 0.0)
    {
        return Err("上下班时间或常规小时数无效".to_string());
    }

    let summary_book = Workbook::open(&table_c_path)?;
    let summary_sheet = summary_book.sheets.first().ok_or_else(|| "汇总表没有工作表".to_string())?;
    let template_book = Workbook::open(&template_path)?;
    let main_template = choose_sheet_or_first(&template_book.sheets, "New timesheet").ok_or_else(|| "模板没有考勤主表".to_string())?;
    let month_start = excel_serial_to_date(parse_number(main_template.sheet.value("M3"))? as i32)?;
    let month_days = days_in_month(month_start.0, month_start.1);
    let (day_headers, employees) = match read_summary_table(&summary_book, summary_sheet) {
        Ok(result) => result,
        Err(summary_error) => {
            let day_headers = template_month_headers(main_template, month_start, month_days);
            let employees = read_payroll_summary(summary_sheet, &day_headers, &schedule)
                .map_err(|payroll_error| format!("{summary_error}; {payroll_error}"))?;
            (day_headers, employees)
        }
    };
    if employees.is_empty() {
        return Err("Rust 生成引擎未读取到员工数据，请确认 A/E/F/G/K/L/M 列结构".to_string());
    }

    let output_dir = payload
        .output_dir
        .as_ref()
        .filter(|item| !item.trim().is_empty())
        .map(PathBuf::from)
        .unwrap_or_else(|| table_c_path.with_file_name(format!("{}_Rust生成表Bs", table_c_path.file_stem().and_then(|item| item.to_str()).unwrap_or("汇总表"))));
    fs::create_dir_all(&output_dir).map_err(|err| format!("创建输出目录失败: {err}"))?;
    let mut generated_files = Vec::new();
    let mut warnings = Vec::new();
    let total = employees.len();
    let mut progress = ProgressSnapshot { current: 0, total, message: String::new() };
    for (index, employee) in employees.iter().enumerate() {
        progress.current = index + 1;
        progress.message = format!("生成 {}.{}", employee.no, employee.name);
        let output = write_employee(
            &template_book,
            main_template,
            employee,
            &day_headers,
            month_start,
            &output_dir,
            &template_path,
            &schedule,
            payload.count_holidays,
            &signature_font_path,
            payload.signature_scale,
            manager_signatures.as_ref(),
            &mut warnings,
        )?;
        generated_files.push(output);
    }
    let report_path = output_dir.join("生成说明.txt");
    write_report(&report_path, &generated_files, &template_path, payload.count_holidays, &schedule, &warnings)?;
    if generated_files.is_empty() {
        warnings.push("未生成任何文件".to_string());
    }
    Ok(GenerateResult {
        output_dir: output_dir.to_string_lossy().to_string(),
        report_path: report_path.to_string_lossy().to_string(),
        generated_count: generated_files.len(),
        generated_files: generated_files.iter().map(|item| item.to_string_lossy().to_string()).collect(),
        warnings,
        progress,
    })
}

fn read_summary_table(book: &Workbook, sheet: &WorkbookSheet) -> Result<(Vec<(i32, String)>, Vec<Employee>), String> {
    let day_headers = group_day_types(book, sheet)?;
    let mut employees = Vec::new();
    let mut row = 3;
    loop {
        let name = sheet.sheet.value(&format!("C{row}")).trim().to_string();
        let position = sheet.sheet.value(&format!("H{row}")).trim().to_string();
        if name.is_empty() && position.is_empty() {
            break;
        }
        let no_raw = sheet.sheet.value(&format!("A{row}")).trim();
        if no_raw.is_empty() {
            row += 1;
            continue;
        }
        let no = parse_number(no_raw)? as i32;
        let mut days = HashMap::new();
        for (column_index, (day, _kind)) in day_headers.iter().enumerate() {
            let column = num_to_col(col_to_num("J") + column_index as i32);
            days.insert(*day, parse_summary_value(sheet.sheet.value(&format!("{column}{row}")))?);
        }
        employees.push(Employee {
            no,
            employee_no: sheet.sheet.value(&format!("B{row}")).trim().to_string(),
            name,
            project: sheet.sheet.value(&format!("D{row}")).trim().to_string(),
            passport: sheet.sheet.value(&format!("F{row}")).trim().to_string(),
            crew_group: sheet.sheet.value(&format!("G{row}")).trim().to_string(),
            position,
            days,
            correction_nwh: parse_optional_number(sheet.sheet.value(&format!("BQ{row}"))),
            correction_normal_ot: parse_optional_number(sheet.sheet.value(&format!("BR{row}"))),
            correction_weekend_ot: parse_optional_number(sheet.sheet.value(&format!("BS{row}"))),
            correction_holiday_ot: parse_optional_number(sheet.sheet.value(&format!("BT{row}"))),
        });
        row += 1;
    }
    if employees.is_empty() {
        return Err("考勤汇总表未读取到员工数据".to_string());
    }
    Ok((day_headers, employees))
}

fn read_payroll_summary(sheet: &WorkbookSheet, day_headers: &[(i32, String)], schedule: &Schedule) -> Result<Vec<Employee>, String> {
    let mut employees = Vec::new();
    let mut row = 3;
    loop {
        let no_raw = sheet.sheet.value(&format!("A{row}")).trim().to_string();
        let name = sheet.sheet.value(&format!("E{row}")).trim().to_string();
        let position = sheet.sheet.value(&format!("F{row}")).trim().to_string();
        if no_raw.is_empty() && name.is_empty() && position.is_empty() {
            break;
        }
        let Ok(no) = no_raw.replace(',', "").parse::<f64>() else {
            row += 1;
            continue;
        };
        let normal_hours = parse_optional_number(sheet.sheet.value(&format!("G{row}")));
        let (days, correction_nwh) = distribute_normal_hours(normal_hours, day_headers, schedule);
        employees.push(Employee {
            no: no as i32,
            employee_no: String::new(),
            project: sheet.sheet.value(&format!("C{row}")).trim().to_string(),
            passport: sheet.sheet.value(&format!("D{row}")).trim().to_string(),
            crew_group: String::new(),
            name,
            position,
            days,
            correction_nwh,
            correction_normal_ot: parse_optional_number(sheet.sheet.value(&format!("K{row}"))),
            correction_weekend_ot: parse_optional_number(sheet.sheet.value(&format!("L{row}"))),
            correction_holiday_ot: parse_optional_number(sheet.sheet.value(&format!("M{row}"))),
        });
        row += 1;
    }
    if employees.is_empty() {
        return Err("工资汇总表未读取到员工数据，请确认 A/E/F/G/K/L/M 列结构".to_string());
    }
    Ok(employees)
}

fn template_month_headers(main_template: &WorkbookSheet, month_start: (i32, i32, i32), month_days: i32) -> Vec<(i32, String)> {
    let rest_weekday = main_template.sheet.value("N7").trim();
    let rest_weekday = if rest_weekday.is_empty() { "Friday" } else { rest_weekday };
    (1..=month_days)
        .map(|day| {
            let weekday = weekday_name(month_start.0, month_start.1, day);
            let kind = if weekday.eq_ignore_ascii_case(rest_weekday) { "rest" } else { "work" };
            (day, kind.to_string())
        })
        .collect()
}

fn distribute_normal_hours(total_hours: f64, day_headers: &[(i32, String)], schedule: &Schedule) -> (HashMap<i32, DayEntry>, f64) {
    let mut remaining = total_hours;
    let mut days = HashMap::new();
    for (day, kind) in day_headers {
        if kind != "work" || remaining <= 0.0 {
            days.insert(*day, DayEntry::Blank);
            continue;
        }
        let value = remaining.min(schedule.normal_hours);
        days.insert(*day, DayEntry::Hours(value));
        remaining -= value;
    }
    (days, remaining.max(0.0))
}

fn group_day_types(book: &Workbook, sheet: &WorkbookSheet) -> Result<Vec<(i32, String)>, String> {
    let mut day_cells = Vec::new();
    let mut column_index = col_to_num("J");
    loop {
        let column = num_to_col(column_index);
        let ref_name = format!("{column}2");
        let raw_day = sheet.sheet.value(&ref_name).trim();
        if raw_day.is_empty() {
            break;
        }
        let Ok(day) = raw_day.parse::<f64>() else {
            break;
        };
        let style = sheet.sheet.cells.get(&ref_name).and_then(|cell| cell.style).unwrap_or(0);
        day_cells.push((day as i32, style));
        column_index += 1;
    }
    if day_cells.is_empty() {
        return Err("汇总表未找到日期头(J2开始)".to_string());
    }
    let style_signatures = style_fill_signatures(book).unwrap_or_default();
    let mut grouped: HashMap<String, Vec<i32>> = HashMap::new();
    for (day, style) in &day_cells {
        let signature = style_signatures.get(style).cloned().unwrap_or_else(|| format!("style:{style}"));
        grouped.entry(signature).or_default().push(*day);
    }
    let work_signature = grouped
        .iter()
        .max_by_key(|(_, days)| days.len())
        .map(|(signature, _)| signature.clone())
        .ok_or_else(|| "汇总表日期样式无效".to_string())?;
    let mut other_signatures = grouped.keys().filter(|item| **item != work_signature).cloned().collect::<Vec<_>>();
    let rest_signature = other_signatures
        .iter()
        .max_by_key(|signature| {
            let days = grouped.get(*signature).cloned().unwrap_or_default();
            let weekly = days.windows(2).filter(|pair| pair[1] - pair[0] == 7).count();
            (weekly, days.len())
        })
        .cloned();
    let holiday_signatures = other_signatures
        .drain(..)
        .filter(|signature| Some(signature) != rest_signature.as_ref())
        .collect::<HashSet<_>>();
    Ok(day_cells
        .into_iter()
        .map(|(day, style)| {
            let signature = style_signatures.get(&style).cloned().unwrap_or_else(|| format!("style:{style}"));
            let kind = if signature == work_signature {
                "work"
            } else if Some(&signature) == rest_signature.as_ref() {
                "rest"
            } else if holiday_signatures.contains(&signature) {
                "holiday"
            } else {
                "work"
            };
            (day, kind.to_string())
        })
        .collect())
}

fn style_fill_signatures(book: &Workbook) -> Result<HashMap<u32, String>, String> {
    let Some(raw) = book.entries.get("xl/styles.xml") else {
        return Ok(HashMap::new());
    };
    let fills = collect_child_xml(raw, b"fills", b"fill")?;
    let xfs = collect_xf_fill_ids(raw)?;
    let mut result = HashMap::new();
    for (index, fill_id) in xfs.into_iter().enumerate() {
        let signature = fills.get(fill_id as usize).cloned().unwrap_or_default();
        result.insert(index as u32, signature);
    }
    Ok(result)
}

fn collect_child_xml(raw: &[u8], parent: &[u8], child: &[u8]) -> Result<Vec<String>, String> {
    let mut reader = Reader::from_reader(Cursor::new(raw));
    let mut buf = Vec::new();
    let mut in_parent = false;
    let mut capture_depth = 0usize;
    let mut current = Vec::new();
    let mut items = Vec::new();
    loop {
        buf.clear();
        let event = reader.read_event_into(&mut buf).map_err(|err| format!("解析 styles.xml 失败: {err}"))?;
        if capture_depth > 0 {
            let mut writer = Writer::new(Vec::new());
            writer.write_event(event.to_owned()).map_err(|err| format!("缓存 styles.xml 失败: {err}"))?;
            current.extend(writer.into_inner());
        }
        match &event {
            Event::Start(event) if local_name(event.name().as_ref()) == parent => in_parent = true,
            Event::End(event) if local_name(event.name().as_ref()) == parent => in_parent = false,
            Event::Start(event) if in_parent && local_name(event.name().as_ref()) == child => {
                capture_depth = 1;
                current.clear();
                let mut writer = Writer::new(Vec::new());
                writer.write_event(Event::Start(event.to_owned())).map_err(|err| format!("缓存 styles.xml 失败: {err}"))?;
                current.extend(writer.into_inner());
            }
            Event::Empty(event) if in_parent && local_name(event.name().as_ref()) == child => {
                let mut writer = Writer::new(Vec::new());
                writer.write_event(Event::Empty(event.to_owned())).map_err(|err| format!("缓存 styles.xml 失败: {err}"))?;
                items.push(String::from_utf8_lossy(&writer.into_inner()).to_string());
            }
            Event::Start(_) if capture_depth > 0 => capture_depth += 1,
            Event::End(event) if capture_depth > 0 => {
                capture_depth -= 1;
                if capture_depth == 0 && local_name(event.name().as_ref()) == child {
                    items.push(String::from_utf8_lossy(&current).to_string());
                    current.clear();
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(items)
}

fn collect_xf_fill_ids(raw: &[u8]) -> Result<Vec<u32>, String> {
    let mut reader = Reader::from_reader(Cursor::new(raw));
    let mut buf = Vec::new();
    let mut in_cell_xfs = false;
    let mut result = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"cellXfs" => in_cell_xfs = true,
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"cellXfs" => in_cell_xfs = false,
            Ok(Event::Empty(event)) | Ok(Event::Start(event)) if in_cell_xfs && local_name(event.name().as_ref()) == b"xf" => {
                result.push(xml_attr(&reader, &event, b"fillId").and_then(|value| value.parse::<u32>().ok()).unwrap_or(0));
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(err) => return Err(format!("解析 cellXfs 失败: {err}")),
        }
    }
    Ok(result)
}

fn parse_summary_value(raw: &str) -> Result<DayEntry, String> {
    let text = raw.trim();
    if text.is_empty() || text == "\\" {
        return Ok(DayEntry::Blank);
    }
    let upper = text.to_ascii_uppercase();
    if matches!(upper.as_str(), "A" | "E" | "S" | "V") {
        return Ok(DayEntry::Leave(upper));
    }
    Ok(DayEntry::Hours(parse_number(text)?))
}

fn write_employee(
    template_book: &Workbook,
    main_template: &WorkbookSheet,
    employee: &Employee,
    day_headers: &[(i32, String)],
    month_start: (i32, i32, i32),
    output_dir: &Path,
    template_path: &Path,
    schedule: &Schedule,
    count_holidays: bool,
    signature_font_path: &Path,
    signature_scale: i32,
    manager_signatures: Option<&ManagerSignatureIndex>,
    warnings: &mut Vec<String>,
) -> Result<PathBuf, String> {
    let mut replacements: HashMap<String, Vec<u8>> = HashMap::new();
    let overtime = choose_sheet(&template_book.sheets, "Overtime");
    let mut main_updates = HashMap::new();
    main_updates.insert("A3".to_string(), CellValue::Text(employee.name.clone()));
    main_updates.insert("C3".to_string(), CellValue::Text(employee.position.clone()));
    main_updates.insert("E3".to_string(), CellValue::Text(employee.passport.clone()));
    main_updates.insert("G3".to_string(), CellValue::Text(employee.project.clone()));
    main_updates.insert("I3".to_string(), CellValue::Text(employee.crew_group.clone()));
    main_updates.insert("J3".to_string(), CellValue::Text(employee.employee_no.clone()));
    main_updates.insert("N8".to_string(), CellValue::Number(schedule.normal_hours));
    let holiday_days = day_headers
        .iter()
        .filter(|(_, kind)| kind == "holiday")
        .map(|(day, _)| *day)
        .collect::<Vec<_>>();
    let rest_days = day_headers
        .iter()
        .filter(|(_, kind)| kind == "rest")
        .map(|(day, _)| *day)
        .collect::<Vec<_>>();
    let rest_weekday = rest_days
        .first()
        .map(|day| weekday_name(month_start.0, month_start.1, *day))
        .unwrap_or_else(|| "Friday".to_string());
    main_updates.insert("N7".to_string(), CellValue::Text(rest_weekday));
    for row in 50..=70 {
        main_updates.insert(format!("J{row}"), CellValue::Blank);
    }
    for (index, day) in holiday_days.iter().enumerate() {
        main_updates.insert(format!("J{}", 50 + index as i32), CellValue::Number(date_to_excel_serial(month_start.0, month_start.1, *day) as f64));
    }

    let anchor_days = day_headers
        .iter()
        .filter_map(|(day, _)| {
            let entry = employee.days.get(day).unwrap_or(&DayEntry::Blank);
            if entry_is_attended(entry) || matches!(entry, DayEntry::Leave(code) if code == "V" || code == "S") {
                Some(*day)
            } else {
                None
            }
        })
        .collect::<Vec<_>>();
    let first_active = anchor_days.iter().min().copied();
    let last_active = anchor_days.iter().max().copied();
    let type_by_day = day_headers.iter().map(|(day, kind)| (*day, kind.clone())).collect::<HashMap<_, _>>();
    let prefix_unpaid = first_active.is_some_and(|first| has_unpaid_marker(employee, &type_by_day, 1, first - 1));
    let suffix_unpaid = last_active.is_some_and(|last| has_unpaid_marker(employee, &type_by_day, last + 1, day_headers.len() as i32));
    let unpaid_adjacent_special_days = if count_holidays {
        HashSet::new()
    } else {
        unpaid_adjacent_special_days(employee, day_headers)
    };

    let mut work_sum = 0.0;
    let mut work_ot_sum = 0.0;
    let mut rest_ot_sum = 0.0;
    let mut holiday_ot_sum = 0.0;
    let mut vacation_days = 0;
    let mut sick_days = 0;
    let mut emergency_days = 0;
    let mut overtime_entries = Vec::new();
    for (day, kind) in day_headers {
        let row = 9 + day;
        let entry = employee.days.get(day).unwrap_or(&DayEntry::Blank);
        let result = fill_day_updates(
            &mut main_updates,
            row,
            month_start,
            *day,
            kind,
            entry,
            (first_active, last_active, prefix_unpaid, suffix_unpaid),
            &unpaid_adjacent_special_days,
            schedule,
        );
        work_sum += result.work_hours;
        work_ot_sum += result.work_ot;
        rest_ot_sum += result.rest_hours;
        holiday_ot_sum += result.holiday_hours;
        if result.work_ot > 0.0 || result.rest_hours > 0.0 || result.holiday_hours > 0.0 {
            let (start, end) = if kind == "work" {
                let start = schedule.afternoon_start + schedule.normal_hours / 24.0;
                (Some(start), Some(start + result.work_ot / 24.0))
            } else if let DayEntry::Hours(hours) = entry {
                let (start, _, _, end) = day_time_inputs(*hours, schedule);
                (start, end)
            } else {
                (None, None)
            };
            overtime_entries.push(OvertimeEntry {
                day: *day,
                start,
                end,
                normal_hours: result.work_ot,
                weekend_hours: result.rest_hours,
                holiday_hours: result.holiday_hours,
            });
        }
        if let DayEntry::Leave(code) = entry {
            if code == "V" {
                vacation_days += 1;
            } else if code == "S" {
                sick_days += 1;
            } else if code == "E" {
                emergency_days += 1;
            }
        }
    }
    for day in (day_headers.len() as i32 + 1)..32 {
        let row = 9 + day;
        for col in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "N", "O", "P", "Q", "R", "T"] {
            main_updates.insert(format!("{col}{row}"), CellValue::Blank);
        }
    }
    let correction_row = find_correction_row(&main_template.sheet).unwrap_or(40);
    main_updates.insert(format!("G{correction_row}"), optional_cell(employee.correction_nwh));
    main_updates.insert(format!("H{correction_row}"), optional_cell(employee.correction_normal_ot));
    main_updates.insert(format!("I{correction_row}"), optional_cell(employee.correction_weekend_ot));
    main_updates.insert(format!("J{correction_row}"), optional_cell(employee.correction_holiday_ot));

    let mut public_payable = 0;
    let mut rest_payable = 0;
    let mut public_attendance = 0;
    let mut rest_attendance = 0;
    for (day, kind) in day_headers {
        let entry = employee.days.get(day).unwrap_or(&DayEntry::Blank);
        if kind == "holiday" && entry_is_attended(entry) {
            public_attendance += 1;
        }
        if kind == "rest" && entry_is_attended(entry) {
            rest_attendance += 1;
        }
        if !count_holidays && matches!(entry, DayEntry::Leave(code) if code == "V") {
            continue;
        }
        if kind == "holiday" && update_has_value(&main_updates, &format!("T{}", 9 + day)) {
            public_payable += 1;
        }
        if kind == "rest" && update_has_value(&main_updates, &format!("T{}", 9 + day)) {
            rest_payable += 1;
        }
    }

    let work_day_count = if work_sum > 0.0 { (work_sum / schedule.normal_hours).ceil() } else { 0.0 };
    let public_day_count = if count_holidays { public_attendance as f64 } else { public_payable as f64 };
    let vacation_day_count = vacation_days as f64;
    let sick_day_count = if count_holidays { 0.0 } else { sick_days as f64 };
    let rest_day_count = if count_holidays { rest_attendance as f64 } else { rest_payable as f64 };
    let mut payable_day_count = work_day_count + public_day_count + sick_day_count + rest_day_count;
    if !count_holidays {
        payable_day_count += vacation_day_count;
    }
    let normal_total = work_sum + employee.correction_nwh;
    main_updates.insert("A6".to_string(), CellValue::Number(day_headers.len() as f64));
    main_updates.insert("B6".to_string(), optional_cell(payable_day_count));
    main_updates.insert("E6".to_string(), optional_cell(work_day_count));
    main_updates.insert("I6".to_string(), optional_cell(public_day_count));
    main_updates.insert("J6".to_string(), optional_cell(vacation_day_count));
    main_updates.insert("K6".to_string(), optional_cell(sick_day_count));
    main_updates.insert("L6".to_string(), optional_cell(rest_day_count));
    main_updates.insert("M6".to_string(), if count_holidays { CellValue::Blank } else { optional_cell(emergency_days as f64) });
    main_updates.insert("G9".to_string(), optional_cell(normal_total));
    main_updates.insert("H9".to_string(), optional_cell(work_ot_sum + employee.correction_normal_ot));
    main_updates.insert("I9".to_string(), optional_cell(rest_ot_sum + employee.correction_weekend_ot));
    main_updates.insert("J9".to_string(), optional_cell(holiday_ot_sum + employee.correction_holiday_ot));
    replacements.insert(main_template.part_name.clone(), rewrite_sheet(&template_book.entries[&main_template.part_name], &main_updates)?);

    if let Some(overtime_sheet) = overtime {
        let mut ot_updates = HashMap::new();
        ot_updates.insert("B5".to_string(), CellValue::Text(employee.name.clone()));
        ot_updates.insert("B7".to_string(), CellValue::Text(employee.project.clone()));
        ot_updates.insert("I7".to_string(), CellValue::Text(employee.position.clone()));
        let correction = find_correction_row(&overtime_sheet.sheet).unwrap_or(42);
        for row in 12..=correction {
            for col in ["A", "B", "E", "H", "I", "J"] {
                ot_updates.insert(format!("{col}{row}"), CellValue::Blank);
            }
        }
        let detail_end = correction - 1;
        for (idx, entry) in overtime_entries.iter().take((detail_end - 11).max(0) as usize).enumerate() {
            let row = 12 + idx as i32;
            ot_updates.insert(format!("A{row}"), CellValue::Number(date_to_excel_serial(month_start.0, month_start.1, entry.day) as f64));
            set_optional_number(&mut ot_updates, &format!("B{row}"), entry.start);
            set_optional_number(&mut ot_updates, &format!("E{row}"), entry.end);
            ot_updates.insert(format!("H{row}"), optional_cell(entry.normal_hours));
            ot_updates.insert(format!("I{row}"), optional_cell(entry.weekend_hours));
            ot_updates.insert(format!("J{row}"), optional_cell(entry.holiday_hours));
        }
        ot_updates.insert(format!("H{correction}"), optional_cell(employee.correction_normal_ot));
        ot_updates.insert(format!("I{correction}"), optional_cell(employee.correction_weekend_ot));
        ot_updates.insert(format!("J{correction}"), optional_cell(employee.correction_holiday_ot));
        let normal_detail_total = overtime_entries.iter().map(|entry| entry.normal_hours).sum::<f64>();
        let weekend_detail_total = overtime_entries.iter().map(|entry| entry.weekend_hours).sum::<f64>();
        let holiday_detail_total = overtime_entries.iter().map(|entry| entry.holiday_hours).sum::<f64>();
        let normal_total = normal_detail_total + employee.correction_normal_ot;
        let weekend_total = weekend_detail_total + employee.correction_weekend_ot;
        let holiday_total = holiday_detail_total + employee.correction_holiday_ot;
        let normal_days = overtime_entries.iter().filter(|entry| entry.normal_hours > 0.0).count() as f64 + if employee.correction_normal_ot > 0.0 { 1.0 } else { 0.0 };
        let weekend_days = overtime_entries.iter().filter(|entry| entry.weekend_hours > 0.0).count() as f64 + if employee.correction_weekend_ot > 0.0 { 1.0 } else { 0.0 };
        let holiday_days_count = overtime_entries.iter().filter(|entry| entry.holiday_hours > 0.0).count() as f64 + if employee.correction_holiday_ot > 0.0 { 1.0 } else { 0.0 };
        ot_updates.insert("H43".to_string(), optional_cell(normal_total));
        ot_updates.insert("I43".to_string(), optional_cell(weekend_total));
        ot_updates.insert("J43".to_string(), optional_cell(holiday_total));
        ot_updates.insert("H44".to_string(), optional_cell(normal_days));
        ot_updates.insert("I44".to_string(), optional_cell(weekend_days));
        ot_updates.insert("J44".to_string(), optional_cell(holiday_days_count));
        replacements.insert(overtime_sheet.part_name.clone(), rewrite_sheet(&template_book.entries[&overtime_sheet.part_name], &ot_updates)?);
    }

    let signature_png = render_signature_png(&signature_text(&employee.name), signature_font_path, signature_scale)?;
    apply_signature(
        template_book,
        &mut replacements,
        main_template,
        &signature_png,
        "xl/media/generated_signature.png",
        "Generated Employee Signature",
        SignaturePlacement::Contain,
        (0, 41, 6, 43),
        (1, 41, 3, 43),
    )?;
    if let Some(manager_signatures) = manager_signatures {
        if let Some(signature_path) = manager_signatures.find(&employee.name) {
            let extension = image_extension(signature_path).ok_or_else(|| format!("管理层签名图片格式不支持: {}", signature_path.display()))?;
            let image_bytes = fs::read(signature_path).map_err(|err| format!("读取管理层签名图片失败 {}: {err}", signature_path.display()))?;
            let media_path = format!("xl/media/generated_manager_signature.{extension}");
            apply_signature(
                template_book,
                &mut replacements,
                main_template,
                &image_bytes,
                &media_path,
                "Generated Manager Signature",
                SignaturePlacement::Contain,
                (7, 41, 12, 43),
                (7, 41, 12, 43),
            )?;
            if let Some(overtime_sheet) = overtime {
                apply_signature(
                    template_book,
                    &mut replacements,
                    overtime_sheet,
                    &image_bytes,
                    &media_path,
                    "Generated Manager Signature",
                    SignaturePlacement::Contain,
                    (5, 56, 9, 58),
                    (5, 56, 9, 58),
                )?;
            }
        } else {
            warnings.push(format!("未找到管理层签名图片: {}", employee.name));
        }
    }
    if let Some(overtime_sheet) = overtime {
        apply_signature(
            template_book,
            &mut replacements,
            overtime_sheet,
            &signature_png,
            "xl/media/generated_signature.png",
            "Generated Employee Signature",
            SignaturePlacement::Stretch,
            (4, 52, 9, 52),
            (7, 51, 8, 53),
        )?;
    }

    let output_path = output_dir.join(format!(
        "{}.{}-{}{}",
        employee.no,
        sanitize_filename(&employee.name),
        sanitize_filename(&employee.position),
        template_path.extension().and_then(|item| item.to_str()).map(|item| format!(".{item}")).unwrap_or_else(|| ".xlsm".to_string())
    ));
    save_workbook(template_book, &output_path, &replacements)?;
    Ok(output_path)
}

fn rewrite_sheet(raw: &[u8], updates: &HashMap<String, CellValue>) -> Result<Vec<u8>, String> {
    let mut reader = Reader::from_reader(Cursor::new(raw));
    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();
    let mut skip_cell_depth = 0usize;
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(_event)) if skip_cell_depth > 0 => skip_cell_depth += 1,
            Ok(Event::End(_)) if skip_cell_depth > 1 => skip_cell_depth -= 1,
            Ok(Event::End(_)) if skip_cell_depth == 1 => skip_cell_depth = 0,
            Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"c" => {
                let ref_name = xml_attr(&reader, &event, b"r").unwrap_or_default().to_ascii_uppercase();
                if let Some(value) = updates.get(&ref_name) {
                    write_cell(&mut writer, &reader, &event, value)?;
                    skip_cell_depth = 1;
                } else {
                    writer.write_event(Event::Start(event.into_owned())).map_err(|err| format!("写入 sheet 失败: {err}"))?;
                }
            }
            Ok(Event::Empty(event)) if local_name(event.name().as_ref()) == b"c" => {
                let ref_name = xml_attr(&reader, &event, b"r").unwrap_or_default().to_ascii_uppercase();
                if let Some(value) = updates.get(&ref_name) {
                    write_cell(&mut writer, &reader, &event, value)?;
                } else {
                    writer.write_event(Event::Empty(event.into_owned())).map_err(|err| format!("写入 sheet 失败: {err}"))?;
                }
            }
            Ok(Event::Eof) => break,
            Ok(event) if skip_cell_depth == 0 => writer.write_event(event.into_owned()).map_err(|err| format!("重写 sheet 失败: {err}"))?,
            Ok(_) => {}
            Err(err) => return Err(format!("读取 sheet 失败: {err}")),
        }
    }
    Ok(writer.into_inner())
}

fn write_cell(writer: &mut Writer<Vec<u8>>, reader: &Reader<Cursor<&[u8]>>, event: &BytesStart<'_>, value: &CellValue) -> Result<(), String> {
    let ref_name = xml_attr(reader, event, b"r").unwrap_or_default();
    let style = xml_attr(reader, event, b"s");
    let mut cell = BytesStart::new("c");
    cell.push_attribute(("r", ref_name.as_str()));
    if let Some(style) = style.as_ref() {
        cell.push_attribute(("s", style.as_str()));
    }
    match value {
        CellValue::Blank => writer.write_event(Event::Empty(cell)).map_err(|err| format!("写入空单元格失败: {err}"))?,
        CellValue::Number(number) => {
            writer.write_event(Event::Start(cell)).map_err(|err| format!("写入数字单元格失败: {err}"))?;
            writer.write_event(Event::Start(BytesStart::new("v"))).map_err(|err| format!("写入 v 失败: {err}"))?;
            let text = number_to_text(*number);
            writer.write_event(Event::Text(BytesText::new(&text))).map_err(|err| format!("写入数字失败: {err}"))?;
            writer.write_event(Event::End(BytesStart::new("v").to_end())).map_err(|err| format!("结束 v 失败: {err}"))?;
            writer.write_event(Event::End(BytesStart::new("c").to_end())).map_err(|err| format!("结束 c 失败: {err}"))?;
        }
        CellValue::Text(text) if text.is_empty() => writer.write_event(Event::Empty(cell)).map_err(|err| format!("写入空文本失败: {err}"))?,
        CellValue::Text(text) => {
            cell.push_attribute(("t", "inlineStr"));
            writer.write_event(Event::Start(cell)).map_err(|err| format!("写入文本单元格失败: {err}"))?;
            writer.write_event(Event::Start(BytesStart::new("is"))).map_err(|err| format!("写入 is 失败: {err}"))?;
            writer.write_event(Event::Start(BytesStart::new("t"))).map_err(|err| format!("写入 t 失败: {err}"))?;
            writer.write_event(Event::Text(BytesText::new(text))).map_err(|err| format!("写入文本失败: {err}"))?;
            writer.write_event(Event::End(BytesStart::new("t").to_end())).map_err(|err| format!("结束 t 失败: {err}"))?;
            writer.write_event(Event::End(BytesStart::new("is").to_end())).map_err(|err| format!("结束 is 失败: {err}"))?;
            writer.write_event(Event::End(BytesStart::new("c").to_end())).map_err(|err| format!("结束 c 失败: {err}"))?;
        }
    }
    Ok(())
}

fn resolve_signature_font(path: Option<&str>) -> Result<PathBuf, String> {
    if let Some(path) = path {
        let path = PathBuf::from(path);
        if path.is_file() {
            return Ok(path);
        }
    }
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).parent().unwrap().to_path_buf();
    let candidate = root.join("Nothing_You_Could_Do").join("NothingYouCouldDo-Regular.ttf");
    if candidate.is_file() {
        return Ok(candidate);
    }
    Err("缺少签名字体文件: Nothing_You_Could_Do/NothingYouCouldDo-Regular.ttf".to_string())
}

fn build_manager_signature_index(path: Option<&str>) -> Result<ManagerSignatureIndex, String> {
    let path = path
        .map(str::trim)
        .filter(|item| !item.is_empty())
        .ok_or_else(|| "已勾选插入管理层签名，请选择管理层签名图片目录".to_string())?;
    let dir = PathBuf::from(path);
    if !dir.is_dir() {
        return Err(format!("管理层签名图片目录不存在: {}", dir.display()));
    }
    let mut files = HashMap::new();
    for entry in fs::read_dir(&dir).map_err(|err| format!("读取管理层签名图片目录失败: {err}"))? {
        let path = entry.map_err(|err| format!("读取管理层签名图片失败: {err}"))?.path();
        if !path.is_file() || image_extension(&path).is_none() {
            continue;
        }
        if let Some(stem) = path.file_stem().and_then(|item| item.to_str()) {
            files.entry(normalize_name_key(stem)).or_insert(path);
        }
    }
    if files.is_empty() {
        return Err(format!("管理层签名图片目录中没有 PNG/JPG 图片: {}", dir.display()));
    }
    Ok(ManagerSignatureIndex { files })
}

impl ManagerSignatureIndex {
    fn find(&self, employee_name: &str) -> Option<&PathBuf> {
        self.files.get(&normalize_name_key(employee_name))
    }
}

fn normalize_name_key(value: &str) -> String {
    value.split_whitespace().collect::<Vec<_>>().join(" ").to_ascii_lowercase()
}

fn image_extension(path: &Path) -> Option<String> {
    let extension = path.extension()?.to_str()?.to_ascii_lowercase();
    match extension.as_str() {
        "png" | "jpg" | "jpeg" => Some(extension),
        _ => None,
    }
}

fn image_content_type(extension: &str) -> &'static str {
    match extension.to_ascii_lowercase().as_str() {
        "jpg" | "jpeg" => "image/jpeg",
        _ => "image/png",
    }
}

fn signature_text(employee_name: &str) -> String {
    employee_name
        .split_whitespace()
        .take(2)
        .collect::<Vec<_>>()
        .join(" ")
}

fn render_signature_png(signature: &str, font_path: &Path, signature_scale: i32) -> Result<Vec<u8>, String> {
    let font_data = fs::read(font_path).map_err(|err| format!("读取签名字体失败 {}: {err}", font_path.display()))?;
    let font = Font::try_from_vec(font_data).ok_or_else(|| "签名字体文件无效".to_string())?;
    let mut rgba = vec![0u8; SIGNATURE_MEDIA_WIDTH * SIGNATURE_MEDIA_HEIGHT * 4];
    let scale_value = 150.0 * signature_scale as f32 / 100.0;
    let scale = Scale::uniform(scale_value);
    let v_metrics = font.v_metrics(scale);
    let glyphs = font.layout(signature, scale, point(0.0, v_metrics.ascent)).collect::<Vec<_>>();
    let bounds = glyphs
        .iter()
        .filter_map(|glyph| glyph.pixel_bounding_box())
        .fold(None, |acc, bb| match acc {
            None => Some(bb),
            Some(mut current) => {
                current.min.x = current.min.x.min(bb.min.x);
                current.min.y = current.min.y.min(bb.min.y);
                current.max.x = current.max.x.max(bb.max.x);
                current.max.y = current.max.y.max(bb.max.y);
                Some(current)
            }
        });
    let Some(bounds) = bounds else {
        return encode_png_rgba(&rgba, SIGNATURE_MEDIA_WIDTH as u32, SIGNATURE_MEDIA_HEIGHT as u32);
    };
    let width_ratio = (0.9 * signature_scale as f32 / 100.0).min(0.98);
    let height_ratio = (0.72 * signature_scale as f32 / 100.0).min(0.9);
    let max_width = SIGNATURE_MEDIA_WIDTH as f32 * width_ratio;
    let max_height = SIGNATURE_MEDIA_HEIGHT as f32 * height_ratio;
    let text_width = (bounds.max.x - bounds.min.x).max(1) as f32;
    let text_height = (bounds.max.y - bounds.min.y).max(1) as f32;
    let fit = (max_width / text_width).min(max_height / text_height).min(1.0);
    let scale = Scale::uniform(scale_value * fit);
    let v_metrics = font.v_metrics(scale);
    let glyphs = font.layout(signature, scale, point(0.0, v_metrics.ascent)).collect::<Vec<_>>();
    let bounds = glyphs
        .iter()
        .filter_map(|glyph| glyph.pixel_bounding_box())
        .fold(None, |acc, bb| match acc {
            None => Some(bb),
            Some(mut current) => {
                current.min.x = current.min.x.min(bb.min.x);
                current.min.y = current.min.y.min(bb.min.y);
                current.max.x = current.max.x.max(bb.max.x);
                current.max.y = current.max.y.max(bb.max.y);
                Some(current)
            }
        })
        .ok_or_else(|| "签名文字无法渲染".to_string())?;
    let text_width = bounds.max.x - bounds.min.x;
    let text_height = bounds.max.y - bounds.min.y;
    let offset_x = (SIGNATURE_MEDIA_WIDTH as i32 - text_width) / 2 - bounds.min.x;
    let offset_y = (SIGNATURE_MEDIA_HEIGHT as i32 - text_height) / 2 - bounds.min.y + (SIGNATURE_MEDIA_HEIGHT as f32 * 0.04) as i32;
    for glyph in font.layout(signature, scale, point(offset_x as f32, offset_y as f32 + v_metrics.ascent)) {
        if let Some(bb) = glyph.pixel_bounding_box() {
            glyph.draw(|x, y, coverage| {
                let px = x as i32 + bb.min.x;
                let py = y as i32 + bb.min.y;
                if px < 0 || py < 0 || px >= SIGNATURE_MEDIA_WIDTH as i32 || py >= SIGNATURE_MEDIA_HEIGHT as i32 {
                    return;
                }
                let index = (py as usize * SIGNATURE_MEDIA_WIDTH + px as usize) * 4;
                rgba[index] = 20;
                rgba[index + 1] = 20;
                rgba[index + 2] = 20;
                rgba[index + 3] = (coverage * 245.0).round() as u8;
            });
        }
    }
    encode_png_rgba(&rgba, SIGNATURE_MEDIA_WIDTH as u32, SIGNATURE_MEDIA_HEIGHT as u32)
}

fn encode_png_rgba(rgba: &[u8], width: u32, height: u32) -> Result<Vec<u8>, String> {
    let mut output = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut output, width, height);
        encoder.set_color(png::ColorType::Rgba);
        encoder.set_depth(png::BitDepth::Eight);
        let mut writer = encoder.write_header().map_err(|err| format!("生成签名 PNG 失败: {err}"))?;
        writer.write_image_data(rgba).map_err(|err| format!("写入签名 PNG 失败: {err}"))?;
    }
    Ok(output)
}

fn apply_signature(
    book: &Workbook,
    replacements: &mut HashMap<String, Vec<u8>>,
    sheet_info: &WorkbookSheet,
    signature_png: &[u8],
    media_path: &str,
    picture_name: &str,
    placement: SignaturePlacement,
    target_range: (i32, i32, i32, i32),
    fallback_anchor: (i32, i32, i32, i32),
) -> Result<(), String> {
    replacements.insert(media_path.to_string(), signature_png.to_vec());
    let extension = Path::new(media_path)
        .extension()
        .and_then(|item| item.to_str())
        .unwrap_or("png")
        .to_ascii_lowercase();
    ensure_image_content_type(book, replacements, &extension, image_content_type(&extension))?;
    let drawing_part = ensure_sheet_drawing_part(book, replacements, sheet_info)?;
    let rels_part = format!("{}/_rels/{}.rels", parent_path(&drawing_part), file_name(&drawing_part));
    let rels_raw = replacements
        .get(&rels_part)
        .or_else(|| book.entries.get(&rels_part))
        .cloned()
        .unwrap_or_else(empty_relationships);
    let rel_id = next_relationship_id(&rels_raw);
    replacements.insert(rels_part, append_relationship(&rels_raw, &rel_id, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", &format!("../media/{}", file_name(media_path)))?);
    let drawing_raw = replacements
        .get(&drawing_part)
        .or_else(|| book.entries.get(&drawing_part))
        .cloned()
        .unwrap_or_else(empty_drawing);
    let anchor = match placement {
        SignaturePlacement::Stretch => AnchorBounds::from_cells(fallback_anchor),
        SignaturePlacement::Contain => fit_image_anchor(book, sheet_info, signature_png, &extension, fallback_anchor),
    };
    replacements.insert(drawing_part, rewrite_drawing_with_signature(&drawing_raw, &rel_id, picture_name, target_range, anchor)?);
    Ok(())
}

impl AnchorBounds {
    fn from_cells(bounds: (i32, i32, i32, i32)) -> Self {
        let (from_col, from_row, to_col, to_row) = bounds;
        Self {
            from: AnchorMarker { col: from_col, row: from_row, col_off: 0, row_off: 0 },
            to: AnchorMarker { col: to_col, row: to_row, col_off: 0, row_off: 0 },
        }
    }
}

fn fit_image_anchor(book: &Workbook, sheet_info: &WorkbookSheet, image_bytes: &[u8], extension: &str, bounds: (i32, i32, i32, i32)) -> AnchorBounds {
    let Some((image_width, image_height)) = image_dimensions(image_bytes, extension) else {
        return AnchorBounds::from_cells(bounds);
    };
    if image_width == 0 || image_height == 0 {
        return AnchorBounds::from_cells(bounds);
    }
    let Some(sheet_raw) = book.entries.get(&sheet_info.part_name) else {
        return AnchorBounds::from_cells(bounds);
    };
    let metrics = parse_sheet_metrics(sheet_raw);
    let (from_col, from_row, to_col, to_row) = bounds;
    let target_width = metrics.range_width_pixels(from_col, to_col + 1);
    let target_height = metrics.range_height_pixels(from_row, to_row + 1);
    if target_width <= 0.0 || target_height <= 0.0 {
        return AnchorBounds::from_cells(bounds);
    }
    let image_ratio = image_width as f64 / image_height as f64;
    let target_ratio = target_width / target_height;
    let (fit_width, fit_height) = if image_ratio > target_ratio {
        (target_width * 0.92, target_width * 0.92 / image_ratio)
    } else {
        (target_height * 0.82 * image_ratio, target_height * 0.82)
    };
    let offset_x = ((target_width - fit_width) / 2.0).max(0.0);
    let offset_y = ((target_height - fit_height) / 2.0).max(0.0);
    AnchorBounds {
        from: metrics.marker_from_offset(from_col, from_row, offset_x, offset_y),
        to: metrics.marker_from_offset(from_col, from_row, offset_x + fit_width, offset_y + fit_height),
    }
}

impl SheetMetrics {
    fn col_width_pixels(&self, col: i32) -> f64 {
        self.col_widths
            .iter()
            .find(|(min, max, _)| col >= *min && col <= *max)
            .map(|(_, _, width)| excel_col_width_to_pixels(*width))
            .unwrap_or_else(|| excel_col_width_to_pixels(self.default_col_width))
    }

    fn row_height_pixels(&self, row: i32) -> f64 {
        self.row_heights.get(&row).copied().unwrap_or(self.default_row_height) * 4.0 / 3.0
    }

    fn range_width_pixels(&self, from_col: i32, to_col_exclusive: i32) -> f64 {
        (from_col..to_col_exclusive).map(|col| self.col_width_pixels(col)).sum()
    }

    fn range_height_pixels(&self, from_row: i32, to_row_exclusive: i32) -> f64 {
        (from_row..to_row_exclusive).map(|row| self.row_height_pixels(row)).sum()
    }

    fn marker_from_offset(&self, start_col: i32, start_row: i32, x_pixels: f64, y_pixels: f64) -> AnchorMarker {
        let (col, col_off) = marker_axis_from_offset(start_col, x_pixels, |index| self.col_width_pixels(index));
        let (row, row_off) = marker_axis_from_offset(start_row, y_pixels, |index| self.row_height_pixels(index));
        AnchorMarker { col, row, col_off, row_off }
    }
}

fn marker_axis_from_offset(mut index: i32, mut offset_pixels: f64, size_at: impl Fn(i32) -> f64) -> (i32, i32) {
    while offset_pixels > 0.0 {
        let size = size_at(index).max(1.0);
        if offset_pixels <= size {
            return (index, (offset_pixels * EMU_PER_PIXEL as f64).round() as i32);
        }
        offset_pixels -= size;
        index += 1;
    }
    (index, 0)
}

fn excel_col_width_to_pixels(width: f64) -> f64 {
    (width * 7.0 + 5.0).floor().max(1.0)
}

fn parse_sheet_metrics(raw: &[u8]) -> SheetMetrics {
    let mut metrics = SheetMetrics {
        default_col_width: 8.43,
        default_row_height: 15.0,
        col_widths: Vec::new(),
        row_heights: HashMap::new(),
    };
    let mut reader = Reader::from_reader(Cursor::new(raw));
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(event)) | Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"sheetFormatPr" => {
                if let Some(value) = xml_attr(&reader, &event, b"defaultColWidth").and_then(|item| item.parse::<f64>().ok()) {
                    metrics.default_col_width = value;
                }
                if let Some(value) = xml_attr(&reader, &event, b"defaultRowHeight").and_then(|item| item.parse::<f64>().ok()) {
                    metrics.default_row_height = value;
                }
            }
            Ok(Event::Empty(event)) | Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"col" => {
                let min = xml_attr(&reader, &event, b"min").and_then(|item| item.parse::<i32>().ok()).unwrap_or(1) - 1;
                let max = xml_attr(&reader, &event, b"max").and_then(|item| item.parse::<i32>().ok()).unwrap_or(min + 1) - 1;
                if let Some(width) = xml_attr(&reader, &event, b"width").and_then(|item| item.parse::<f64>().ok()) {
                    metrics.col_widths.push((min, max, width));
                }
            }
            Ok(Event::Empty(event)) | Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"row" => {
                if let (Some(row), Some(height)) = (
                    xml_attr(&reader, &event, b"r").and_then(|item| item.parse::<i32>().ok()),
                    xml_attr(&reader, &event, b"ht").and_then(|item| item.parse::<f64>().ok()),
                ) {
                    metrics.row_heights.insert(row - 1, height);
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(_) => break,
        }
    }
    metrics
}

fn image_dimensions(bytes: &[u8], extension: &str) -> Option<(u32, u32)> {
    match extension.to_ascii_lowercase().as_str() {
        "png" => png_dimensions(bytes),
        "jpg" | "jpeg" => jpeg_dimensions(bytes),
        _ => None,
    }
}

fn png_dimensions(bytes: &[u8]) -> Option<(u32, u32)> {
    if bytes.len() < 24 || &bytes[0..8] != b"\x89PNG\r\n\x1a\n" {
        return None;
    }
    Some((
        u32::from_be_bytes(bytes[16..20].try_into().ok()?),
        u32::from_be_bytes(bytes[20..24].try_into().ok()?),
    ))
}

fn jpeg_dimensions(bytes: &[u8]) -> Option<(u32, u32)> {
    if bytes.len() < 4 || bytes[0] != 0xFF || bytes[1] != 0xD8 {
        return None;
    }
    let mut index = 2usize;
    while index + 9 < bytes.len() {
        while index < bytes.len() && bytes[index] == 0xFF {
            index += 1;
        }
        if index >= bytes.len() {
            return None;
        }
        let marker = bytes[index];
        index += 1;
        if marker == 0xD9 || marker == 0xDA {
            return None;
        }
        if index + 2 > bytes.len() {
            return None;
        }
        let length = u16::from_be_bytes([bytes[index], bytes[index + 1]]) as usize;
        if length < 2 || index + length > bytes.len() {
            return None;
        }
        if matches!(marker, 0xC0 | 0xC1 | 0xC2 | 0xC3 | 0xC5 | 0xC6 | 0xC7 | 0xC9 | 0xCA | 0xCB | 0xCD | 0xCE | 0xCF) && length >= 7 {
            let height = u16::from_be_bytes([bytes[index + 3], bytes[index + 4]]) as u32;
            let width = u16::from_be_bytes([bytes[index + 5], bytes[index + 6]]) as u32;
            return Some((width, height));
        }
        index += length;
    }
    None
}

fn ensure_image_content_type(book: &Workbook, replacements: &mut HashMap<String, Vec<u8>>, extension: &str, content_type: &str) -> Result<(), String> {
    let raw = replacements
        .get("[Content_Types].xml")
        .or_else(|| book.entries.get("[Content_Types].xml"))
        .ok_or_else(|| "Excel 缺少 [Content_Types].xml".to_string())?;
    if String::from_utf8_lossy(raw).to_ascii_lowercase().contains(&format!("extension=\"{}\"", extension.to_ascii_lowercase())) {
        return Ok(());
    }
    replacements.insert("[Content_Types].xml".to_string(), append_content_type_default(raw, extension, content_type)?);
    Ok(())
}

fn ensure_drawing_content_type(book: &Workbook, replacements: &mut HashMap<String, Vec<u8>>, drawing_part: &str) -> Result<(), String> {
    let raw = replacements
        .get("[Content_Types].xml")
        .or_else(|| book.entries.get("[Content_Types].xml"))
        .ok_or_else(|| "Excel 缺少 [Content_Types].xml".to_string())?;
    let part_name = format!("/{drawing_part}");
    if String::from_utf8_lossy(raw).contains(&part_name) {
        return Ok(());
    }
    replacements.insert(
        "[Content_Types].xml".to_string(),
        append_content_type_override(raw, &part_name, "application/vnd.openxmlformats-officedocument.drawing+xml")?,
    );
    Ok(())
}

fn ensure_sheet_drawing_part(book: &Workbook, replacements: &mut HashMap<String, Vec<u8>>, sheet_info: &WorkbookSheet) -> Result<String, String> {
    if let Some(existing) = sheet_drawing_part(book, replacements, sheet_info)? {
        return Ok(existing);
    }
    let drawing_part = next_drawing_part(book, replacements);
    let sheet_rels_part = format!("{}/_rels/{}.rels", parent_path(&sheet_info.part_name), file_name(&sheet_info.part_name));
    let sheet_rels_raw = replacements
        .get(&sheet_rels_part)
        .or_else(|| book.entries.get(&sheet_rels_part))
        .cloned()
        .unwrap_or_else(empty_relationships);
    let rel_id = next_relationship_id(&sheet_rels_raw);
    let target = relative_path(parent_path(&sheet_info.part_name), &drawing_part);
    replacements.insert(sheet_rels_part, append_relationship(&sheet_rels_raw, &rel_id, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", &target)?);
    let sheet_raw = replacements
        .get(&sheet_info.part_name)
        .or_else(|| book.entries.get(&sheet_info.part_name))
        .ok_or_else(|| format!("缺少工作表 XML: {}", sheet_info.part_name))?;
    replacements.insert(sheet_info.part_name.clone(), append_sheet_drawing(sheet_raw, &rel_id)?);
    replacements.insert(drawing_part.clone(), empty_drawing());
    replacements.insert(format!("{}/_rels/{}.rels", parent_path(&drawing_part), file_name(&drawing_part)), empty_relationships());
    ensure_drawing_content_type(book, replacements, &drawing_part)?;
    Ok(drawing_part)
}

fn sheet_drawing_part(book: &Workbook, replacements: &HashMap<String, Vec<u8>>, sheet_info: &WorkbookSheet) -> Result<Option<String>, String> {
    let sheet_raw = replacements
        .get(&sheet_info.part_name)
        .or_else(|| book.entries.get(&sheet_info.part_name))
        .ok_or_else(|| format!("缺少工作表 XML: {}", sheet_info.part_name))?;
    let Some(rel_id) = find_sheet_drawing_rel_id(sheet_raw)? else {
        return Ok(None);
    };
    let rels_part = format!("{}/_rels/{}.rels", parent_path(&sheet_info.part_name), file_name(&sheet_info.part_name));
    let Some(rels_raw) = replacements.get(&rels_part).or_else(|| book.entries.get(&rels_part)) else {
        return Ok(None);
    };
    let Some(target) = relationship_target(rels_raw, &rel_id)? else {
        return Ok(None);
    };
    Ok(Some(normalize_part_path(&parent_path(&sheet_info.part_name), &target)))
}

fn find_sheet_drawing_rel_id(raw: &[u8]) -> Result<Option<String>, String> {
    let mut reader = Reader::from_reader(Cursor::new(raw));
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(event)) | Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"drawing" => {
                return Ok(xml_attr(&reader, &event, b"r:id").or_else(|| xml_attr(&reader, &event, b"id")));
            }
            Ok(Event::Eof) => return Ok(None),
            Ok(_) => {}
            Err(err) => return Err(format!("读取工作表 drawing 失败: {err}")),
        }
    }
}

fn relationship_target(raw: &[u8], id: &str) -> Result<Option<String>, String> {
    let mut reader = Reader::from_reader(Cursor::new(raw));
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(event)) | Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"Relationship" => {
                if xml_attr(&reader, &event, b"Id").as_deref() == Some(id) {
                    return Ok(xml_attr(&reader, &event, b"Target"));
                }
            }
            Ok(Event::Eof) => return Ok(None),
            Ok(_) => {}
            Err(err) => return Err(format!("读取 relationship 失败: {err}")),
        }
    }
}

fn append_relationship(raw: &[u8], id: &str, rel_type: &str, target: &str) -> Result<Vec<u8>, String> {
    let mut reader = Reader::from_reader(Cursor::new(raw));
    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"Relationships" => {
                let mut rel = BytesStart::new("Relationship");
                rel.push_attribute(("Id", id));
                rel.push_attribute(("Type", rel_type));
                rel.push_attribute(("Target", target));
                writer.write_event(Event::Empty(rel)).map_err(|err| format!("写入 relationship 失败: {err}"))?;
                writer.write_event(Event::End(event.into_owned())).map_err(|err| format!("结束 relationship 失败: {err}"))?;
            }
            Ok(Event::Eof) => break,
            Ok(event) => writer.write_event(event.into_owned()).map_err(|err| format!("重写 relationship 失败: {err}"))?,
            Err(err) => return Err(format!("读取 relationship 失败: {err}")),
        }
    }
    Ok(writer.into_inner())
}

fn next_relationship_id(raw: &[u8]) -> String {
    let mut reader = Reader::from_reader(Cursor::new(raw));
    let mut buf = Vec::new();
    let mut max_id = 0;
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(event)) | Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"Relationship" => {
                if let Some(id) = xml_attr(&reader, &event, b"Id") {
                    if let Some(number) = id.strip_prefix("rId").and_then(|item| item.parse::<i32>().ok()) {
                        max_id = max_id.max(number);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(_) => break,
        }
    }
    format!("rId{}", max_id + 1)
}

fn append_sheet_drawing(raw: &[u8], rel_id: &str) -> Result<Vec<u8>, String> {
    let mut reader = Reader::from_reader(Cursor::new(raw));
    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"worksheet" => {
                let mut drawing = BytesStart::new("drawing");
                drawing.push_attribute(("r:id", rel_id));
                writer.write_event(Event::Empty(drawing)).map_err(|err| format!("写入 sheet drawing 失败: {err}"))?;
                writer.write_event(Event::End(event.into_owned())).map_err(|err| format!("结束 sheet 失败: {err}"))?;
            }
            Ok(Event::Eof) => break,
            Ok(event) => writer.write_event(event.into_owned()).map_err(|err| format!("重写 sheet drawing 失败: {err}"))?,
            Err(err) => return Err(format!("读取 sheet drawing 失败: {err}")),
        }
    }
    Ok(writer.into_inner())
}

fn append_content_type_default(raw: &[u8], extension: &str, content_type: &str) -> Result<Vec<u8>, String> {
    let mut reader = Reader::from_reader(Cursor::new(raw));
    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"Types" => {
                let mut default = BytesStart::new("Default");
                default.push_attribute(("Extension", extension));
                default.push_attribute(("ContentType", content_type));
                writer.write_event(Event::Empty(default)).map_err(|err| format!("写入 content type 失败: {err}"))?;
                writer.write_event(Event::End(event.into_owned())).map_err(|err| format!("结束 content type 失败: {err}"))?;
            }
            Ok(Event::Eof) => break,
            Ok(event) => writer.write_event(event.into_owned()).map_err(|err| format!("重写 content type 失败: {err}"))?,
            Err(err) => return Err(format!("读取 content type 失败: {err}")),
        }
    }
    Ok(writer.into_inner())
}

fn append_content_type_override(raw: &[u8], part_name: &str, content_type: &str) -> Result<Vec<u8>, String> {
    let mut reader = Reader::from_reader(Cursor::new(raw));
    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"Types" => {
                let mut override_node = BytesStart::new("Override");
                override_node.push_attribute(("PartName", part_name));
                override_node.push_attribute(("ContentType", content_type));
                writer.write_event(Event::Empty(override_node)).map_err(|err| format!("写入 content type override 失败: {err}"))?;
                writer.write_event(Event::End(event.into_owned())).map_err(|err| format!("结束 content type 失败: {err}"))?;
            }
            Ok(Event::Eof) => break,
            Ok(event) => writer.write_event(event.into_owned()).map_err(|err| format!("重写 content type 失败: {err}"))?,
            Err(err) => return Err(format!("读取 content type 失败: {err}")),
        }
    }
    Ok(writer.into_inner())
}

#[derive(Default)]
struct AnchorState {
    depth: usize,
    has_pic: bool,
    in_from: bool,
    in_to: bool,
    current_tag: Vec<u8>,
    from_col: Option<i32>,
    from_row: Option<i32>,
    to_col: Option<i32>,
    to_row: Option<i32>,
    bytes: Vec<u8>,
}

fn rewrite_drawing_with_signature(raw: &[u8], rel_id: &str, picture_name: &str, target_range: (i32, i32, i32, i32), fallback_anchor: AnchorBounds) -> Result<Vec<u8>, String> {
    let mut reader = Reader::from_reader(Cursor::new(raw));
    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();
    let mut anchor: Option<AnchorState> = None;
    loop {
        buf.clear();
        let event = reader.read_event_into(&mut buf).map_err(|err| format!("读取 drawing 失败: {err}"))?;
        if let Some(state) = anchor.as_mut() {
            write_anchor_event(state, &event)?;
            update_anchor_state(&reader, state, &event);
            if matches!(&event, Event::End(end) if local_name(end.name().as_ref()) == b"twoCellAnchor") {
                let state = anchor.take().unwrap();
                if !(state.has_pic && anchor_overlaps(&state, target_range)) {
                    writer.get_mut().write_all(&state.bytes).map_err(|err| format!("写入 drawing anchor 失败: {err}"))?;
                }
            }
            continue;
        }
        match event {
            Event::Start(event) if local_name(event.name().as_ref()) == b"twoCellAnchor" => {
                let mut state = AnchorState { depth: 1, ..AnchorState::default() };
                write_anchor_event(&mut state, &Event::Start(event.into_owned()))?;
                anchor = Some(state);
            }
            Event::End(event) if local_name(event.name().as_ref()) == b"wsDr" => {
                writer.get_mut().write_all(&signature_anchor_xml(rel_id, picture_name, fallback_anchor)).map_err(|err| format!("写入签名 anchor 失败: {err}"))?;
                writer.write_event(Event::End(event.into_owned())).map_err(|err| format!("结束 drawing 失败: {err}"))?;
            }
            Event::Eof => break,
            event => writer.write_event(event.into_owned()).map_err(|err| format!("重写 drawing 失败: {err}"))?,
        }
    }
    Ok(writer.into_inner())
}

fn write_anchor_event(state: &mut AnchorState, event: &Event<'_>) -> Result<(), String> {
    let mut writer = Writer::new(Vec::new());
    writer.write_event(event.to_owned()).map_err(|err| format!("缓存 drawing anchor 失败: {err}"))?;
    state.bytes.extend(writer.into_inner());
    Ok(())
}

fn update_anchor_state(reader: &Reader<Cursor<&[u8]>>, state: &mut AnchorState, event: &Event<'_>) {
    match event {
        Event::Start(event) => {
            state.depth += 1;
            let name = local_name(event.name().as_ref()).to_vec();
            if name == b"pic" {
                state.has_pic = true;
            } else if name == b"from" {
                state.in_from = true;
            } else if name == b"to" {
                state.in_to = true;
            } else if name == b"col" || name == b"row" {
                state.current_tag = name;
            }
            if local_name(event.name().as_ref()) == b"blip" && xml_attr(reader, event, b"r:embed").is_some() {
                state.has_pic = true;
            }
        }
        Event::Empty(event) => {
            if local_name(event.name().as_ref()) == b"pic" || local_name(event.name().as_ref()) == b"blip" {
                state.has_pic = true;
            }
        }
        Event::Text(text) => {
            if let Ok(value) = text.decode() {
                if let Ok(number) = value.parse::<i32>() {
                    if state.in_from && state.current_tag == b"col" {
                        state.from_col = Some(number);
                    } else if state.in_from && state.current_tag == b"row" {
                        state.from_row = Some(number);
                    } else if state.in_to && state.current_tag == b"col" {
                        state.to_col = Some(number);
                    } else if state.in_to && state.current_tag == b"row" {
                        state.to_row = Some(number);
                    }
                }
            }
        }
        Event::End(event) => {
            let event_name = event.name();
            let name = local_name(event_name.as_ref());
            if name == b"from" {
                state.in_from = false;
            } else if name == b"to" {
                state.in_to = false;
            } else if name == b"col" || name == b"row" {
                state.current_tag.clear();
            }
            state.depth = state.depth.saturating_sub(1);
        }
        _ => {}
    }
}

fn anchor_overlaps(state: &AnchorState, target: (i32, i32, i32, i32)) -> bool {
    let (min_col, min_row, max_col, max_row) = target;
    let from_col = state.from_col.unwrap_or(0);
    let from_row = state.from_row.unwrap_or(0);
    let to_col = state.to_col.unwrap_or(from_col);
    let to_row = state.to_row.unwrap_or(from_row);
    from_col <= max_col && to_col >= min_col && from_row <= max_row && to_row >= min_row
}

fn signature_anchor_xml(rel_id: &str, picture_name: &str, bounds: AnchorBounds) -> Vec<u8> {
    let from_col = bounds.from.col;
    let from_row = bounds.from.row;
    let from_col_off = bounds.from.col_off;
    let from_row_off = bounds.from.row_off;
    let to_col = bounds.to.col;
    let to_row = bounds.to.row;
    let to_col_off = bounds.to.col_off;
    let to_row_off = bounds.to.row_off;
    let picture_id = rel_id
        .strip_prefix("rId")
        .and_then(|item| item.parse::<i32>().ok())
        .map(|number| 9000 + number)
        .unwrap_or(9001);
    format!(
        r#"<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>{from_col}</xdr:col><xdr:colOff>{from_col_off}</xdr:colOff><xdr:row>{from_row}</xdr:row><xdr:rowOff>{from_row_off}</xdr:rowOff></xdr:from><xdr:to><xdr:col>{to_col}</xdr:col><xdr:colOff>{to_col_off}</xdr:colOff><xdr:row>{to_row}</xdr:row><xdr:rowOff>{to_row_off}</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{picture_id}" name="{picture_name}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="{rel_id}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>"#
    )
    .into_bytes()
}

fn empty_relationships() -> Vec<u8> {
    format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{REL_NS}"></Relationships>"#).into_bytes()
}

fn empty_drawing() -> Vec<u8> {
    format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="{DRAWING_NS}" xmlns:r="{OFFICE_REL_NS}" xmlns:a="{A_NS}"></xdr:wsDr>"#).into_bytes()
}

fn next_drawing_part(book: &Workbook, replacements: &HashMap<String, Vec<u8>>) -> String {
    let mut number = 1;
    loop {
        let candidate = format!("xl/drawings/drawing{number}.xml");
        if !book.entries.contains_key(&candidate) && !replacements.contains_key(&candidate) {
            return candidate;
        }
        number += 1;
    }
}

fn parent_path(path: &str) -> String {
    path.rsplit_once('/').map(|item| item.0.to_string()).unwrap_or_default()
}

fn file_name(path: &str) -> String {
    path.rsplit('/').next().unwrap_or(path).to_string()
}

fn normalize_part_path(base: &str, target: &str) -> String {
    if target.starts_with("xl/") {
        return target.to_string();
    }
    let mut parts = base.split('/').filter(|item| !item.is_empty()).map(str::to_string).collect::<Vec<_>>();
    for part in target.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            item => parts.push(item.to_string()),
        }
    }
    parts.join("/")
}

fn relative_path(from_dir: String, target: &str) -> String {
    let from = from_dir.split('/').filter(|item| !item.is_empty()).collect::<Vec<_>>();
    let target_parts = target.split('/').filter(|item| !item.is_empty()).collect::<Vec<_>>();
    let mut common = 0;
    while common < from.len() && common < target_parts.len() && from[common] == target_parts[common] {
        common += 1;
    }
    let mut parts = Vec::new();
    for _ in common..from.len() {
        parts.push("..".to_string());
    }
    parts.extend(target_parts[common..].iter().map(|item| item.to_string()));
    parts.join("/")
}

impl Workbook {
    fn open(path: &Path) -> Result<Self, String> {
        let data = fs::read(path).map_err(|err| format!("读取 Excel 文件失败 {}: {err}", path.display()))?;
        let mut archive = ZipArchive::new(Cursor::new(data)).map_err(|err| format!("打开 Excel ZIP 失败 {}: {err}", path.display()))?;
        let mut entries = HashMap::new();
        let mut infos = HashMap::new();
        for index in 0..archive.len() {
            let mut file = archive.by_index(index).map_err(|err| format!("读取 ZIP 项失败: {err}"))?;
            if file.is_dir() {
                continue;
            }
            let mut bytes = Vec::new();
            file.read_to_end(&mut bytes).map_err(|err| format!("读取 ZIP 内容失败: {err}"))?;
            let name = file.name().replace('\\', "/");
            infos.insert(name.clone(), ZipMeta { compression: file.compression(), unix_mode: file.unix_mode() });
            entries.insert(name, bytes);
        }
        let shared_strings = load_shared_strings(&entries);
        let sheets = load_workbook_sheets(&entries, &shared_strings)?;
        Ok(Self { entries, infos, sheets })
    }
}

impl Sheet {
    fn value(&self, ref_name: &str) -> &str {
        self.cells.get(&ref_name.to_ascii_uppercase()).map(|item| item.value.as_str()).unwrap_or("")
    }
}

fn save_workbook(book: &Workbook, output_path: &Path, replacements: &HashMap<String, Vec<u8>>) -> Result<(), String> {
    let file = File::create(output_path).map_err(|err| format!("创建生成文件失败 {}: {err}", output_path.display()))?;
    let mut writer = ZipWriter::new(file);
    let mut names = book.entries.keys().cloned().collect::<Vec<_>>();
    names.sort();
    for name in names {
        let meta = book.infos.get(&name);
        let mut options = SimpleFileOptions::default().compression_method(meta.map(|item| item.compression).unwrap_or(CompressionMethod::Deflated));
        if let Some(mode) = meta.and_then(|item| item.unix_mode) {
            options = options.unix_permissions(mode);
        }
        writer.start_file(&name, options).map_err(|err| format!("写入 ZIP 项失败 {name}: {err}"))?;
        writer
            .write_all(replacements.get(&name).or_else(|| book.entries.get(&name)).ok_or_else(|| format!("缺少 ZIP 项: {name}"))?)
            .map_err(|err| format!("写入 ZIP 内容失败 {name}: {err}"))?;
    }
    let mut extra_names = replacements
        .keys()
        .filter(|name| !book.entries.contains_key(*name))
        .cloned()
        .collect::<Vec<_>>();
    extra_names.sort();
    for name in extra_names {
        writer
            .start_file(&name, SimpleFileOptions::default().compression_method(CompressionMethod::Deflated))
            .map_err(|err| format!("写入新增 ZIP 项失败 {name}: {err}"))?;
        writer
            .write_all(replacements.get(&name).ok_or_else(|| format!("缺少新增 ZIP 项: {name}"))?)
            .map_err(|err| format!("写入新增 ZIP 内容失败 {name}: {err}"))?;
    }
    writer.finish().map_err(|err| format!("完成生成文件失败: {err}"))?;
    Ok(())
}

fn load_shared_strings(entries: &HashMap<String, Vec<u8>>) -> Vec<String> {
    let Some(raw) = entries.get("xl/sharedStrings.xml") else { return Vec::new(); };
    let mut reader = Reader::from_reader(Cursor::new(raw.as_slice()));
    let mut buf = Vec::new();
    let mut items = Vec::new();
    let mut in_si = false;
    let mut current = String::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"si" => { in_si = true; current.clear(); }
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"si" => { items.push(current.clone()); in_si = false; }
            Ok(Event::Text(event)) if in_si => if let Ok(text) = event.decode() { current.push_str(&text); },
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(_) => break,
        }
    }
    items
}

fn load_workbook_sheets(entries: &HashMap<String, Vec<u8>>, shared_strings: &[String]) -> Result<Vec<WorkbookSheet>, String> {
    let mut rel_targets = HashMap::new();
    let rels_raw = entries.get("xl/_rels/workbook.xml.rels").ok_or_else(|| "Excel 缺少 workbook.xml.rels".to_string())?;
    let mut reader = Reader::from_reader(Cursor::new(rels_raw.as_slice()));
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(event)) | Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"Relationship" => {
                if let (Some(id), Some(target)) = (xml_attr(&reader, &event, b"Id"), xml_attr(&reader, &event, b"Target")) {
                    rel_targets.insert(id, if target.starts_with("xl/") { target } else { format!("xl/{target}") });
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(err) => return Err(format!("解析 workbook rels 失败: {err}")),
        }
    }
    let workbook_raw = entries.get("xl/workbook.xml").ok_or_else(|| "Excel 缺少 workbook.xml".to_string())?;
    let mut reader = Reader::from_reader(Cursor::new(workbook_raw.as_slice()));
    let mut sheets = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(event)) | Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"sheet" => {
                let Some(id) = xml_attr(&reader, &event, b"r:id").or_else(|| xml_attr(&reader, &event, b"id")) else { continue; };
                let Some(target) = rel_targets.get(&id) else { continue; };
                let Some(raw) = entries.get(target) else { continue; };
                sheets.push(WorkbookSheet {
                    name: xml_attr(&reader, &event, b"name").unwrap_or_default(),
                    part_name: target.clone(),
                    sheet: parse_sheet(raw, shared_strings)?,
                });
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(err) => return Err(format!("解析 workbook 失败: {err}")),
        }
    }
    Ok(sheets)
}

#[derive(Default)]
struct CellBuild {
    ref_name: String,
    cell_type: String,
    style: Option<u32>,
    value_text: String,
    inline_text: String,
    in_value: bool,
    in_text: bool,
}

fn parse_sheet(raw: &[u8], shared_strings: &[String]) -> Result<Sheet, String> {
    let mut cells = HashMap::new();
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
                    style: xml_attr(&reader, &event, b"s").and_then(|value| value.parse::<u32>().ok()),
                    ..CellBuild::default()
                });
            }
            Ok(Event::Empty(event)) if local_name(event.name().as_ref()) == b"c" => {
                if let Some(ref_name) = xml_attr(&reader, &event, b"r") {
                    cells.insert(
                        ref_name.to_ascii_uppercase(),
                        CellData {
                            value: String::new(),
                            style: xml_attr(&reader, &event, b"s").and_then(|value| value.parse::<u32>().ok()),
                        },
                    );
                }
            }
            Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"v" => if let Some(cell) = current.as_mut() { cell.in_value = true; },
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"v" => if let Some(cell) = current.as_mut() { cell.in_value = false; },
            Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"t" => if let Some(cell) = current.as_mut() { cell.in_text = true; },
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"t" => if let Some(cell) = current.as_mut() { cell.in_text = false; },
            Ok(Event::Text(event)) => if let Some(cell) = current.as_mut() {
                if let Ok(text) = event.decode() {
                    if cell.in_value { cell.value_text.push_str(&text); }
                    if cell.in_text { cell.inline_text.push_str(&text); }
                }
            },
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"c" => {
                if let Some(cell) = current.take() {
                    if !cell.ref_name.is_empty() {
                        let value = if cell.cell_type == "inlineStr" {
                            cell.inline_text
                        } else if cell.cell_type == "s" {
                            cell.value_text.parse::<usize>().ok().and_then(|index| shared_strings.get(index).cloned()).unwrap_or_default()
                        } else {
                            cell.value_text
                        };
                        cells.insert(cell.ref_name, CellData { value, style: cell.style });
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(err) => return Err(format!("解析 sheet 失败: {err}")),
        }
    }
    Ok(Sheet { cells })
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

fn choose_sheet<'a>(sheets: &'a [WorkbookSheet], name: &str) -> Option<&'a WorkbookSheet> {
    sheets.iter().find(|item| item.name == name)
}

fn choose_sheet_or_first<'a>(sheets: &'a [WorkbookSheet], name: &str) -> Option<&'a WorkbookSheet> {
    choose_sheet(sheets, name).or_else(|| sheets.first())
}

fn find_correction_row(sheet: &Sheet) -> Option<i32> {
    for row in 35..=45 {
        let text = sheet.value(&format!("A{row}"));
        if text.contains("修正上月加班") || text.to_ascii_uppercase().contains("FIX OT") {
            return Some(row);
        }
    }
    None
}

fn parse_time(value: &str) -> Result<f64, String> {
    let (hour, minute) = value.trim().split_once(':').ok_or_else(|| format!("时间格式无效: {value}"))?;
    let hour = hour.parse::<i32>().map_err(|_| format!("时间格式无效: {value}"))?;
    let minute = minute.parse::<i32>().map_err(|_| format!("时间格式无效: {value}"))?;
    if !(0..=23).contains(&hour) || !(0..=59).contains(&minute) {
        return Err(format!("时间超出范围: {value}"));
    }
    Ok((hour as f64 + minute as f64 / 60.0) / 24.0)
}

fn col_to_num(col: &str) -> i32 {
    col.bytes().fold(0, |acc, item| acc * 26 + (item as i32 - b'A' as i32 + 1))
}

fn num_to_col(mut num: i32) -> String {
    let mut chars = Vec::new();
    while num > 0 {
        num -= 1;
        chars.push((b'A' + (num % 26) as u8) as char);
        num /= 26;
    }
    chars.iter().rev().collect()
}

fn day_time_inputs(hours: f64, schedule: &Schedule) -> (Option<f64>, Option<f64>, Option<f64>, Option<f64>) {
    let morning_hours = (schedule.morning_end - schedule.morning_start) * 24.0;
    if hours <= 0.0 {
        return (None, None, None, None);
    }
    if hours <= morning_hours {
        return (Some(schedule.morning_start), None, None, Some(schedule.morning_start + hours / 24.0));
    }
    (Some(schedule.morning_start), Some(schedule.morning_end), Some(schedule.afternoon_start), Some(schedule.afternoon_start + (hours - morning_hours) / 24.0))
}

fn fill_day_updates(
    updates: &mut HashMap<String, CellValue>,
    row: i32,
    month_start: (i32, i32, i32),
    day: i32,
    day_type: &str,
    entry: &DayEntry,
    payable_context: (Option<i32>, Option<i32>, bool, bool),
    unpaid_adjacent_special_days: &HashSet<i32>,
    schedule: &Schedule,
) -> DayFillResult {
    for col in ["C", "D", "E", "F", "G", "H", "I", "J", "K", "N", "O", "P", "Q", "R", "T"] {
        updates.insert(format!("{col}{row}"), CellValue::Blank);
    }
    updates.insert(format!("A{row}"), CellValue::Number(date_to_excel_serial(month_start.0, month_start.1, day) as f64));
    updates.insert(format!("B{row}"), CellValue::Text(weekday_name(month_start.0, month_start.1, day)));

    let mut result = DayFillResult::default();
    match entry {
        DayEntry::Hours(hours) if *hours > 0.0 => {
            let (c, d, e, f) = day_time_inputs(*hours, schedule);
            set_optional_number(updates, &format!("C{row}"), c);
            set_optional_number(updates, &format!("D{row}"), d);
            set_optional_number(updates, &format!("E{row}"), e);
            set_optional_number(updates, &format!("F{row}"), f);
            set_optional_number(updates, &format!("N{row}"), f);
            if day_type == "work" {
                result.work_hours = hours.min(schedule.normal_hours);
                result.work_ot = (hours - schedule.normal_hours).max(0.0);
            } else if day_type == "rest" {
                result.rest_hours = *hours;
            } else {
                result.holiday_hours = *hours;
            }
            result.payable = true;
        }
        DayEntry::Leave(code) => {
            updates.insert(format!("K{row}"), CellValue::Text(leave_label(code).to_string()));
            result.payable = code == "V" || code == "S";
        }
        _ => {
            let (first_active, last_active, prefix_unpaid, suffix_unpaid) = payable_context;
            let is_prefix = first_active.is_some_and(|first| day < first);
            let is_suffix = last_active.is_some_and(|last| day > last);
            if day_type == "rest" || day_type == "holiday" {
                result.payable = !((is_prefix && prefix_unpaid) || (is_suffix && suffix_unpaid));
            }
            if (day_type == "rest" || day_type == "holiday") && unpaid_adjacent_special_days.contains(&day) {
                result.payable = false;
            }
        }
    }

    if !matches!(entry, DayEntry::Leave(_)) && day_type == "rest" {
        updates.insert(format!("K{row}"), CellValue::Text("Weekend".to_string()));
    } else if !matches!(entry, DayEntry::Leave(_)) && day_type == "holiday" {
        updates.insert(format!("K{row}"), CellValue::Text("Public Holiday".to_string()));
    }
    updates.insert(format!("G{row}"), optional_cell(result.work_hours));
    updates.insert(format!("H{row}"), optional_cell(result.work_ot));
    updates.insert(format!("I{row}"), optional_cell(result.rest_hours));
    updates.insert(format!("J{row}"), optional_cell(result.holiday_hours));
    updates.insert(format!("P{row}"), optional_cell(result.work_ot));
    updates.insert(format!("Q{row}"), optional_cell(result.rest_hours));
    updates.insert(format!("R{row}"), optional_cell(result.holiday_hours));
    if entry_is_attended(entry) {
        updates.insert(format!("O{row}"), CellValue::Number(0.0));
    }
    if result.payable {
        updates.insert(format!("T{row}"), CellValue::Number(1.0));
    }
    result
}

fn leave_label(code: &str) -> &'static str {
    match code {
        "A" => "Absent",
        "E" => "Emergency Leave",
        "S" => "Sick Leave",
        "V" => "Vacation",
        _ => "",
    }
}

fn entry_is_attended(entry: &DayEntry) -> bool {
    matches!(entry, DayEntry::Hours(hours) if *hours > 0.0)
}

fn has_unpaid_marker(employee: &Employee, type_by_day: &HashMap<i32, String>, start: i32, end: i32) -> bool {
    if start > end {
        return false;
    }
    for day in start..=end {
        match employee.days.get(&day).unwrap_or(&DayEntry::Blank) {
            DayEntry::Leave(code) if code == "A" || code == "E" => return true,
            DayEntry::Blank if type_by_day.get(&day).map(String::as_str) == Some("work") => return true,
            _ => {}
        }
    }
    false
}

fn unpaid_adjacent_special_days(employee: &Employee, day_headers: &[(i32, String)]) -> HashSet<i32> {
    let unpaid_leave_days = employee
        .days
        .iter()
        .filter_map(|(day, entry)| match entry {
            DayEntry::Leave(code) if code == "A" || code == "E" => Some(*day),
            _ => None,
        })
        .collect::<HashSet<_>>();
    let mut groups: Vec<Vec<i32>> = Vec::new();
    for (day, _kind) in day_headers.iter().filter(|(_, kind)| kind == "rest" || kind == "holiday") {
        if groups.last().is_none_or(|group| *day != group[group.len() - 1] + 1) {
            groups.push(vec![*day]);
        } else if let Some(group) = groups.last_mut() {
            group.push(*day);
        }
    }
    let mut result = HashSet::new();
    for group in groups {
        let Some(first) = group.first().copied() else { continue };
        let Some(last) = group.last().copied() else { continue };
        if !unpaid_leave_days.contains(&(first - 1)) && !unpaid_leave_days.contains(&(last + 1)) {
            continue;
        }
        for day in group {
            let entry = employee.days.get(&day).unwrap_or(&DayEntry::Blank);
            if !entry_is_attended(entry) && !matches!(entry, DayEntry::Leave(code) if code == "S" || code == "V") {
                result.insert(day);
            }
        }
    }
    result
}

fn update_has_value(updates: &HashMap<String, CellValue>, cell: &str) -> bool {
    match updates.get(cell) {
        Some(CellValue::Number(_)) => true,
        Some(CellValue::Text(text)) => !text.is_empty(),
        _ => false,
    }
}

fn set_optional_number(updates: &mut HashMap<String, CellValue>, cell: &str, value: Option<f64>) {
    updates.insert(cell.to_string(), value.map(CellValue::Number).unwrap_or(CellValue::Blank));
}

fn optional_cell(value: f64) -> CellValue {
    if value.abs() <= 0.000001 { CellValue::Blank } else { CellValue::Number(value) }
}

fn parse_optional_number(value: &str) -> f64 {
    value.trim().replace(',', "").parse::<f64>().unwrap_or(0.0)
}

fn parse_number(value: &str) -> Result<f64, String> {
    value.trim().replace(',', "").parse::<f64>().map_err(|_| format!("无法解析数字: {value:?}"))
}

fn number_to_text(value: f64) -> String {
    if value.fract().abs() <= 0.000001 {
        format!("{}", value.round() as i64)
    } else {
        let mut text = format!("{value:.10}");
        while text.contains('.') && text.ends_with('0') { text.pop(); }
        if text.ends_with('.') { text.pop(); }
        text
    }
}

fn sanitize_filename(value: &str) -> String {
    let mut output = String::new();
    for ch in value.trim().chars() {
        if matches!(ch, '<' | '>' | ':' | '"' | '/' | '\\' | '|' | '?' | '*') {
            output.push('-');
        } else {
            output.push(ch);
        }
    }
    let output = output.trim_matches(['.', ' ']).trim().to_string();
    if output.is_empty() { "blank".to_string() } else { output }
}

fn excel_serial_to_date(serial: i32) -> Result<(i32, i32, i32), String> {
    civil_from_days(serial - 25569).ok_or_else(|| format!("Excel 日期序列无效: {serial}"))
}

fn date_to_excel_serial(year: i32, month: i32, day: i32) -> i32 {
    days_from_civil(year, month, day) + 25569
}

fn days_in_month(year: i32, month: i32) -> i32 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if is_leap_year(year) => 29,
        2 => 28,
        _ => 30,
    }
}

fn is_leap_year(year: i32) -> bool {
    (year % 4 == 0 && year % 100 != 0) || year % 400 == 0
}

fn weekday_name(year: i32, month: i32, day: i32) -> String {
    match (days_from_civil(year, month, day) + 4).rem_euclid(7) {
        0 => "Sunday",
        1 => "Monday",
        2 => "Tuesday",
        3 => "Wednesday",
        4 => "Thursday",
        5 => "Friday",
        _ => "Saturday",
    }
    .to_string()
}

fn days_from_civil(year: i32, month: i32, day: i32) -> i32 {
    let y = year - if month <= 2 { 1 } else { 0 };
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400;
    let mp = month + if month > 2 { -3 } else { 9 };
    let doy = (153 * mp + 2) / 5 + day - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146097 + doe - 719468
}

fn civil_from_days(days: i32) -> Option<(i32, i32, i32)> {
    let z = days + 719468;
    let era = if z >= 0 { z } else { z - 146096 } / 146097;
    let doe = z - era * 146097;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let day = doy - (153 * mp + 2) / 5 + 1;
    let month = mp + if mp < 10 { 3 } else { -9 };
    let year = y + if month <= 2 { 1 } else { 0 };
    if (1..=12).contains(&month) && (1..=31).contains(&day) {
        Some((year, month, day))
    } else {
        None
    }
}

fn write_report(report_path: &Path, generated_files: &[PathBuf], template_path: &Path, count_holidays: bool, schedule: &Schedule, warnings: &[String]) -> Result<(), String> {
    let mut lines = vec![
        format!("模板表B: {}", template_path.display()),
        format!("生成数量: {}", generated_files.len()),
        format!("是否统计假期: {}", if count_holidays { "是" } else { "否" }),
        format!("常规工作小时数: {}", number_to_text(schedule.normal_hours)),
    ];
    if !warnings.is_empty() {
        lines.push(String::new());
        lines.push("提示:".to_string());
        lines.extend(warnings.iter().map(|item| format!("- {item}")));
    }
    lines.push(String::new());
    lines.push("文件列表:".to_string());
    lines.extend(generated_files.iter().filter_map(|path| path.file_name().and_then(|name| name.to_str()).map(|name| format!("- {name}"))));
    fs::write(report_path, lines.join("\n")).map_err(|err| format!("写入 Rust 生成报告失败: {err}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn date_roundtrip() {
        let serial = date_to_excel_serial(2026, 6, 1);
        assert_eq!(excel_serial_to_date(serial).unwrap(), (2026, 6, 1));
        assert_eq!(weekday_name(2026, 6, 5), "Friday");
    }

    #[test]
    fn manager_signature_index_matches_employee_names() {
        let dir = std::env::temp_dir().join("excel-check-manager-signature-index-test");
        let _ = fs::remove_dir_all(&dir);
        fs::create_dir_all(&dir).unwrap();
        fs::write(dir.join("John  Doe.PNG"), b"not a real png").unwrap();
        fs::write(dir.join("ignored.txt"), b"ignore").unwrap();

        let index = build_manager_signature_index(Some(&dir.to_string_lossy())).unwrap();
        assert!(index.find("John Doe").is_some());
        assert!(index.find("john   doe").is_some());
        assert!(index.find("Jane Doe").is_none());

        let _ = fs::remove_dir_all(dir);
    }

    #[test]
    fn rust_generate_real_files_when_available() {
        let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).parent().unwrap().to_path_buf();
        let fixture_dir = root.join("fixtures").join("manual");
        let summary = fixture_dir.join("6月岗位外包工资表-26.7.3.xlsx");
        let template = fixture_dir.join("考勤表模板.xlsm");
        if !(summary.exists() && template.exists()) {
            return;
        }
        let schedule = Schedule {
            morning_start: parse_time(DEFAULT_MORNING_START).unwrap(),
            morning_end: parse_time(DEFAULT_MORNING_END).unwrap(),
            afternoon_start: parse_time(DEFAULT_AFTERNOON_START).unwrap(),
            afternoon_end: parse_time(DEFAULT_AFTERNOON_END).unwrap(),
            normal_hours: DEFAULT_NORMAL_HOURS,
        };
        let summary_book = Workbook::open(&summary).unwrap();
        let summary_sheet = summary_book.sheets.first().unwrap();
        let template_book = Workbook::open(&template).unwrap();
        let main_template = choose_sheet_or_first(&template_book.sheets, "New timesheet").unwrap();
        let month_start = excel_serial_to_date(parse_number(main_template.sheet.value("M3")).unwrap() as i32).unwrap();
        let month_days = days_in_month(month_start.0, month_start.1);
        let employees = match read_summary_table(&summary_book, summary_sheet) {
            Ok((_, employees)) => employees,
            Err(_) => {
                let day_headers = template_month_headers(main_template, month_start, month_days);
                read_payroll_summary(summary_sheet, &day_headers, &schedule).unwrap()
            }
        };
        let first_employee_name = employees.first().unwrap().name.clone();
        let manager_dir = std::env::temp_dir().join("rust-generate-manager-signatures");
        let _ = fs::remove_dir_all(&manager_dir);
        fs::create_dir_all(&manager_dir).unwrap();
        let font_path = root.join("Nothing_You_Could_Do").join("NothingYouCouldDo-Regular.ttf");
        let manager_png = render_signature_png("Manager", &font_path, 100).unwrap();
        fs::write(manager_dir.join(format!("{first_employee_name}.png")), manager_png).unwrap();
        let output = std::env::temp_dir().join("rust-generate-real-output");
        let _ = fs::remove_dir_all(&output);
        let result = run_generate(GeneratePayload {
            table_c_path: summary.to_string_lossy().to_string(),
            template_b_path: template.to_string_lossy().to_string(),
            output_dir: Some(output.to_string_lossy().to_string()),
            count_holidays: false,
            signature_scale: 100,
            morning_start: DEFAULT_MORNING_START.to_string(),
            morning_end: DEFAULT_MORNING_END.to_string(),
            afternoon_start: DEFAULT_AFTERNOON_START.to_string(),
            afternoon_end: DEFAULT_AFTERNOON_END.to_string(),
            normal_hours: DEFAULT_NORMAL_HOURS.to_string(),
            signature_font_path: Some(font_path.to_string_lossy().to_string()),
            insert_manager_signature: true,
            manager_signature_dir: Some(manager_dir.to_string_lossy().to_string()),
        })
        .unwrap();
        assert!(result.generated_count > 0);
        assert!(PathBuf::from(result.report_path).exists());
        assert!(result.generated_files.iter().all(|item| PathBuf::from(item).exists()));
        let first_file = PathBuf::from(&result.generated_files[0]);
        let book = Workbook::open(&first_file).unwrap();
        assert!(book.entries.contains_key("xl/media/generated_signature.png"));
        let generated_signature_refs = book
            .entries
            .iter()
            .filter(|(name, _)| name.starts_with("xl/drawings/drawing") && name.ends_with(".xml"))
            .filter(|(_, data)| String::from_utf8_lossy(data).contains("Generated Employee Signature"))
            .count();
        assert!(generated_signature_refs >= 2);
        assert!(book.entries.contains_key("xl/media/generated_manager_signature.png"));
        let generated_manager_signature_refs = book
            .entries
            .iter()
            .filter(|(name, _)| name.starts_with("xl/drawings/drawing") && name.ends_with(".xml"))
            .filter(|(_, data)| String::from_utf8_lossy(data).contains("Generated Manager Signature"))
            .count();
        assert!(generated_manager_signature_refs >= 2);
        let _ = fs::remove_dir_all(output);
        let _ = fs::remove_dir_all(manager_dir);
    }
}
