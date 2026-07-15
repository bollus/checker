use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::env;
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;
use tauri::Manager;

pub mod rust_check;
pub mod rust_generate;

#[cfg(target_os = "windows")]
use std::os::windows::process::CommandExt;

#[cfg(target_os = "windows")]
const CREATE_NO_WINDOW: u32 = 0x08000000;

fn hide_console_window(command: &mut Command) -> &mut Command {
    #[cfg(target_os = "windows")]
    {
        command.creation_flags(CREATE_NO_WINDOW);
    }
    command
}

#[derive(Debug, Serialize, Deserialize)]
struct BackendEnvelope {
    ok: bool,
    data: Option<Value>,
    warnings: Vec<String>,
    errors: Vec<String>,
    #[serde(default)]
    traceback: Option<String>,
}

fn find_file_recursively(root: &Path, file_name: &str, max_depth: usize) -> Option<PathBuf> {
    if max_depth == 0 || !root.is_dir() {
        return None;
    }
    let entries = fs::read_dir(root).ok()?;
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_file() && path.file_name().and_then(|name| name.to_str()) == Some(file_name) {
            return Some(path);
        }
        if path.is_dir() {
            if let Some(found) = find_file_recursively(&path, file_name, max_depth - 1) {
                return Some(found);
            }
        }
    }
    None
}

fn find_resource_file(app: &tauri::AppHandle, file_name: &str) -> Option<PathBuf> {
    let mut roots: Vec<PathBuf> = Vec::new();
    if let Ok(resource_dir) = app.path().resource_dir() {
        roots.push(resource_dir);
    }
    if let Ok(exe_path) = env::current_exe() {
        if let Some(exe_dir) = exe_path.parent() {
            roots.push(exe_dir.to_path_buf());
        }
    }
    if let Ok(current_dir) = env::current_dir() {
        roots.push(current_dir);
    }
    for root in roots {
        for candidate in [
            root.join(file_name),
            root.join("resources").join(file_name),
        ] {
            if candidate.is_file() {
                return Some(candidate);
            }
        }
        if let Some(found) = find_file_recursively(&root, file_name, 4) {
            return Some(found);
        }
    }
    None
}

fn template_store_path(app: &tauri::AppHandle) -> Result<PathBuf, String> {
    let dir = app
        .path()
        .app_data_dir()
        .map_err(|err| format!("获取应用数据目录失败: {err}"))?;
    fs::create_dir_all(&dir).map_err(|err| format!("创建应用数据目录失败: {err}"))?;
    Ok(dir.join("check_templates.json"))
}

fn parse_templates_value(value: Value) -> Result<Vec<rust_check::CheckTemplate>, String> {
    if value.is_array() {
        serde_json::from_value(value).map_err(|err| format!("模板 JSON 格式无效: {err}"))
    } else if value.get("templates").is_some() {
        serde_json::from_value(value.get("templates").cloned().unwrap_or(Value::Array(vec![])))
            .map_err(|err| format!("模板 JSON 格式无效: {err}"))
    } else {
        serde_json::from_value(value)
            .map(|template| vec![template])
            .map_err(|err| format!("模板 JSON 格式无效: {err}"))
    }
}

fn read_templates_from_path(path: &Path) -> Result<Vec<rust_check::CheckTemplate>, String> {
    let raw = fs::read_to_string(path).map_err(|err| format!("读取模板失败: {err}\n路径: {}", path.display()))?;
    let value: Value = serde_json::from_str(raw.trim_start_matches('\u{feff}'))
        .map_err(|err| format!("解析模板 JSON 失败: {err}\n路径: {}", path.display()))?;
    parse_templates_value(value)
}

fn read_templates(app: &tauri::AppHandle) -> Vec<rust_check::CheckTemplate> {
    if let Ok(path) = template_store_path(app) {
        if path.is_file() {
            if let Ok(templates) = read_templates_from_path(&path) {
                return templates;
            }
        }
    }
    if let Some(path) = find_resource_file(app, "check_templates.json") {
        if let Ok(templates) = read_templates_from_path(&path) {
            return templates;
        }
    }
    Vec::new()
}

fn write_templates(app: &tauri::AppHandle, templates: &[rust_check::CheckTemplate]) -> Result<(), String> {
    let path = template_store_path(app)?;
    let raw = serde_json::to_string_pretty(templates).map_err(|err| format!("序列化模板失败: {err}"))?;
    fs::write(&path, raw).map_err(|err| format!("保存模板失败: {err}\n路径: {}", path.display()))
}

#[tauri::command]
async fn reveal_path(path: String) -> Result<(), String> {
    let target = PathBuf::from(path);
    let command_result = if cfg!(target_os = "windows") {
        let mut command = Command::new("explorer");
        command.arg(format!("/select,{}", target.to_string_lossy()));
        hide_console_window(&mut command).status()
    } else if cfg!(target_os = "macos") {
        Command::new("open").arg("-R").arg(target).status()
    } else {
        let folder = if target.is_dir() {
            target
        } else {
            target.parent().unwrap_or(&target).to_path_buf()
        };
        Command::new("xdg-open").arg(folder).status()
    };
    let status = command_result.map_err(|err| format!("打开所在位置失败: {err}"))?;
    if !status.success() {
        return Err(format!("打开所在位置失败: {status}"));
    }
    Ok(())
}

#[tauri::command]
async fn open_path(path: String) -> Result<(), String> {
    let target = PathBuf::from(&path);
    if !target.exists() {
        return Err(format!("路径不存在: {path}"));
    }
    let command_result = if cfg!(target_os = "windows") {
        let mut command = Command::new("powershell");
        command.args(["-NoProfile", "-Command", "Start-Process -LiteralPath $args[0]"]);
        command.arg(&path);
        hide_console_window(&mut command).status()
    } else if cfg!(target_os = "macos") {
        Command::new("open").arg(target).status()
    } else {
        Command::new("xdg-open").arg(target).status()
    };
    let status = command_result.map_err(|err| format!("打开文件失败: {err}"))?;
    if !status.success() {
        return Err(format!("打开文件失败: {status}"));
    }
    Ok(())
}

#[tauri::command]
async fn read_text_file(path: String) -> Result<String, String> {
    fs::read_to_string(&path).map_err(|err| format!("读取文件失败: {err}\n路径: {path}"))
}

#[tauri::command]
async fn list_templates_rust(app: tauri::AppHandle) -> Result<Value, String> {
    Ok(json!({ "templates": read_templates(&app) }))
}

#[tauri::command]
async fn save_template_rust(app: tauri::AppHandle, template: rust_check::CheckTemplate) -> Result<Value, String> {
    rust_check::validate_template(&template)?;
    let mut templates = read_templates(&app);
    templates.retain(|item| item.name != template.name);
    templates.push(template.clone());
    write_templates(&app, &templates)?;
    Ok(json!({ "template": template }))
}

#[tauri::command]
async fn delete_template_rust(app: tauri::AppHandle, name: String) -> Result<Value, String> {
    let mut templates = read_templates(&app);
    templates.retain(|item| item.name != name);
    write_templates(&app, &templates)?;
    Ok(json!({ "deleted": name }))
}

#[tauri::command]
async fn load_template_file_rust(path: String, template_data: Value) -> Result<Value, String> {
    let templates = if template_data.is_null() {
        read_templates_from_path(Path::new(&path))?
    } else {
        parse_templates_value(template_data)?
    };
    for template in &templates {
        rust_check::validate_template(template)?;
    }
    Ok(json!({ "templates": templates }))
}

#[tauri::command]
async fn export_template_file_rust(path: String, template: rust_check::CheckTemplate) -> Result<Value, String> {
    rust_check::validate_template(&template)?;
    let raw = serde_json::to_string_pretty(&template).map_err(|err| format!("序列化模板失败: {err}"))?;
    fs::write(&path, raw).map_err(|err| format!("导出模板失败: {err}\n路径: {path}"))?;
    Ok(json!({ "path": path }))
}

#[tauri::command]
async fn validate_template_rust(template: rust_check::CheckTemplate) -> Result<Value, String> {
    let rule_count = rust_check::validate_template(&template)?;
    Ok(json!({ "rule_count": rule_count, "template": template }))
}

#[tauri::command]
async fn inspect_workbook_rust(path: String, max_rows: Option<u32>, max_cols: Option<u32>) -> Result<Value, String> {
    let data = rust_check::inspect_workbook(&path, max_rows.unwrap_or(80), max_cols.unwrap_or(30))?;
    serde_json::to_value(data).map_err(|err| format!("序列化预览失败: {err}"))
}

#[tauri::command]
async fn run_check_rust(app: tauri::AppHandle, mut payload: Value) -> Result<BackendEnvelope, String> {
    if payload.get("position_aliases_path").is_none() {
        if let Some(path) = find_resource_file(&app, "position_aliases.json") {
            if let Some(object) = payload.as_object_mut() {
                object.insert(
                    "position_aliases_path".to_string(),
                    Value::String(path.to_string_lossy().to_string()),
                );
            }
        }
    }
    if payload.get("position_rules_path").is_none() {
        if let Some(path) = find_resource_file(&app, "position_rules.json") {
            if let Some(object) = payload.as_object_mut() {
                object.insert(
                    "position_rules_path".to_string(),
                    Value::String(path.to_string_lossy().to_string()),
                );
            }
        }
    }
    let payload: rust_check::CheckPayload = serde_json::from_value(payload)
        .map_err(|err| format!("Rust 核对参数无效: {err}"))?;
    match rust_check::run_check(payload) {
        Ok(data) => Ok(BackendEnvelope {
            ok: true,
            data: Some(serde_json::to_value(data).map_err(|err| format!("Rust 核对结果序列化失败: {err}"))?),
            warnings: vec![],
            errors: vec![],
            traceback: None,
        }),
        Err(error) => Ok(BackendEnvelope {
            ok: false,
            data: None,
            warnings: vec![],
            errors: vec![error],
            traceback: None,
        }),
    }
}

#[tauri::command]
async fn run_generate_rust(app: tauri::AppHandle, payload: Value) -> Result<BackendEnvelope, String> {
    let mut payload = payload;
    if payload.get("signature_font_path").is_none() {
        if let Some(path) = find_resource_file(&app, "NothingYouCouldDo-Regular.ttf") {
            if let Some(object) = payload.as_object_mut() {
                object.insert(
                    "signature_font_path".to_string(),
                    Value::String(path.to_string_lossy().to_string()),
                );
            }
        }
    }
    let payload: rust_generate::GeneratePayload = serde_json::from_value(payload)
        .map_err(|err| format!("Rust 生成参数无效: {err}"))?;
    match rust_generate::run_generate(payload) {
        Ok(data) => Ok(BackendEnvelope {
            ok: true,
            data: Some(serde_json::to_value(data).map_err(|err| format!("Rust 生成结果序列化失败: {err}"))?),
            warnings: vec![],
            errors: vec![],
            traceback: None,
        }),
        Err(error) => Ok(BackendEnvelope {
            ok: false,
            data: None,
            warnings: vec![],
            errors: vec![error],
            traceback: None,
        }),
    }
}

pub fn run() {
    let builder = tauri::Builder::default().plugin(tauri_plugin_dialog::init());
    #[cfg(target_os = "windows")]
    let builder = builder
        .plugin(tauri_plugin_updater::Builder::new().build())
        .plugin(tauri_plugin_process::init());

    builder
        .invoke_handler(tauri::generate_handler![
            run_check_rust,
            run_generate_rust,
            list_templates_rust,
            save_template_rust,
            delete_template_rust,
            load_template_file_rust,
            export_template_file_rust,
            validate_template_rust,
            inspect_workbook_rust,
            reveal_path,
            open_path,
            read_text_file
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
