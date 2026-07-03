use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::env;
use std::fs;
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use tauri::Manager;

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

#[derive(Debug)]
struct BackendCommand {
    program: PathBuf,
    args: Vec<String>,
}

fn executable_name(base: &str) -> String {
    if cfg!(target_os = "windows") {
        format!("{base}.exe")
    } else {
        base.to_string()
    }
}

fn find_file_recursively(root: &Path, file_name: &str, max_depth: usize) -> Option<PathBuf> {
    if max_depth == 0 || !root.is_dir() {
        return None;
    }
    let entries = fs::read_dir(root).ok()?;
    for entry in entries.flatten() {
        let path = entry.path();
        if path.file_name().and_then(|name| name.to_str()) == Some(file_name) {
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

fn find_python() -> String {
    if cfg!(target_os = "windows") {
        "python".to_string()
    } else {
        "python3".to_string()
    }
}

fn bundled_backend(app: &tauri::AppHandle) -> Option<PathBuf> {
    let name = executable_name("excel-check-backend");
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
            root.join(&name),
            root.join("dist-sidecar").join(&name),
            root.join("python_backend").join(&name),
            root.join("resources").join(&name),
            root.join("resources").join("dist-sidecar").join(&name),
        ] {
            if candidate.exists() {
                return Some(candidate);
            }
        }
        if let Some(found) = find_file_recursively(&root, &name, 4) {
            return Some(found);
        }
    }
    None
}

fn dev_backend_script() -> Option<PathBuf> {
    let cwd = env::current_dir().ok()?;
    for candidate in [
        cwd.join("python_backend").join("backend_cli.py"),
        cwd.parent()?.join("python_backend").join("backend_cli.py"),
    ] {
        if candidate.exists() {
            return Some(candidate);
        }
    }
    None
}

fn backend_command(app: &tauri::AppHandle) -> Result<BackendCommand, String> {
    if let Ok(path) = env::var("EXCEL_CHECK_BACKEND") {
        let candidate = PathBuf::from(path);
        if candidate.exists() {
            return Ok(BackendCommand {
                program: candidate,
                args: vec![],
            });
        }
    }

    if let Some(path) = bundled_backend(app) {
        return Ok(BackendCommand {
            program: path,
            args: vec![],
        });
    }

    if let Some(script) = dev_backend_script() {
        return Ok(BackendCommand {
            program: PathBuf::from(find_python()),
            args: vec![script.to_string_lossy().to_string()],
        });
    }

    Err("未找到 Python 后端。开发环境请确认 python_backend/backend_cli.py 存在；打包环境请确认 sidecar 已随应用发布。".to_string())
}

#[tauri::command]
async fn run_backend(app: tauri::AppHandle, action: String, payload: Value) -> Result<BackendEnvelope, String> {
    let backend = backend_command(&app)?;
    let request = json!({ "action": action, "payload": payload }).to_string();

    let mut command = Command::new(&backend.program);
    command
        .args(&backend.args)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped());
    let mut child = hide_console_window(&mut command)
        .spawn()
        .map_err(|err| format!("启动后端失败: {err}"))?;

    if let Some(stdin) = child.stdin.as_mut() {
        stdin
            .write_all(request.as_bytes())
            .map_err(|err| format!("写入后端参数失败: {err}"))?;
    }

    let output = child
        .wait_with_output()
        .map_err(|err| format!("读取后端结果失败: {err}"))?;
    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);

    let mut envelope: BackendEnvelope = serde_json::from_str(stdout.trim())
        .map_err(|err| format!("后端返回格式无效: {err}\nSTDOUT:\n{stdout}\nSTDERR:\n{stderr}"))?;
    if !output.status.success() && envelope.ok {
        envelope.ok = false;
        envelope.errors.push(format!("后端退出码异常: {}", output.status));
    }
    if !stderr.trim().is_empty() {
        envelope.warnings.push(stderr.trim().to_string());
    }
    Ok(envelope)
}

#[tauri::command]
async fn reveal_path(path: String) -> Result<(), String> {
    let target = PathBuf::from(path);
    let command_result = if cfg!(target_os = "windows") {
        let mut command = Command::new("explorer");
        command.arg("/select,").arg(target);
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
    command_result.map_err(|err| format!("打开所在位置失败: {err}"))?;
    Ok(())
}

#[tauri::command]
async fn open_path(path: String) -> Result<(), String> {
    let target = PathBuf::from(&path);
    let command_result = if cfg!(target_os = "windows") {
        let mut command = Command::new("cmd");
        command.args(["/C", "start", "", &path]);
        hide_console_window(&mut command).status()
    } else if cfg!(target_os = "macos") {
        Command::new("open").arg(target).status()
    } else {
        Command::new("xdg-open").arg(target).status()
    };
    command_result.map_err(|err| format!("打开文件失败: {err}"))?;
    Ok(())
}

pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![run_backend, reveal_path, open_path])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
