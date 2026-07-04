use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::env;
use std::fs;
use std::io::{BufRead, BufReader, Write};
use std::path::{Path, PathBuf};
use std::process::{Child, ChildStdin, ChildStdout, Command, Stdio};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::Mutex;
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

struct BackendProcess {
    child: Child,
    stdin: ChildStdin,
    stdout: BufReader<ChildStdout>,
}

struct BackendState {
    process: Mutex<Option<BackendProcess>>,
}

static REQUEST_ID: AtomicU64 = AtomicU64::new(1);

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
async fn run_backend(app: tauri::AppHandle, state: tauri::State<'_, BackendState>, action: String, payload: Value) -> Result<BackendEnvelope, String> {
    let mut guard = state
        .process
        .lock()
        .map_err(|_| "后端状态锁定失败".to_string())?;
    match run_backend_persistent(&app, &mut guard, &action, payload.clone()) {
        Ok(envelope) => Ok(envelope),
        Err(first_error) => {
            if let Some(mut process) = guard.take() {
                let _ = process.child.kill();
                let _ = process.child.wait();
            }
            run_backend_persistent(&app, &mut guard, &action, payload)
                .map_err(|second_error| format!("{second_error}\n首次尝试失败: {first_error}"))
        }
    }
}

fn spawn_backend_process(app: &tauri::AppHandle) -> Result<BackendProcess, String> {
    let backend = backend_command(app)?;
    let mut command = Command::new(&backend.program);
    let mut args = backend.args.clone();
    args.push("--serve".to_string());
    command
        .args(&args)
        .env("PYTHONUTF8", "1")
        .env("PYTHONIOENCODING", "utf-8")
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::null());
    let mut child = hide_console_window(&mut command)
        .spawn()
        .map_err(|err| format!("启动后端失败: {err}"))?;
    let stdin = child.stdin.take().ok_or_else(|| "后端 stdin 初始化失败".to_string())?;
    let stdout = child.stdout.take().ok_or_else(|| "后端 stdout 初始化失败".to_string())?;
    Ok(BackendProcess {
        child,
        stdin,
        stdout: BufReader::new(stdout),
    })
}

fn run_backend_persistent(
    app: &tauri::AppHandle,
    process_slot: &mut Option<BackendProcess>,
    action: &str,
    payload: Value,
) -> Result<BackendEnvelope, String> {
    if process_slot.is_none() {
        *process_slot = Some(spawn_backend_process(app)?);
    }
    let process = process_slot.as_mut().ok_or_else(|| "后端进程未启动".to_string())?;
    let id = REQUEST_ID.fetch_add(1, Ordering::Relaxed);
    let request = json!({ "id": id, "action": action, "payload": payload }).to_string();
    process
        .stdin
        .write_all(request.as_bytes())
        .map_err(|err| format!("写入后端参数失败: {err}"))?;
    process
        .stdin
        .write_all(b"\n")
        .map_err(|err| format!("写入后端参数失败: {err}"))?;
    process
        .stdin
        .flush()
        .map_err(|err| format!("刷新后端参数失败: {err}"))?;

    let mut line = String::new();
    process
        .stdout
        .read_line(&mut line)
        .map_err(|err| format!("读取后端结果失败: {err}"))?;
    if line.trim().is_empty() {
        return Err("后端没有返回结果".to_string());
    }
    let envelope: BackendEnvelope = serde_json::from_str(line.trim())
        .map_err(|err| format!("后端返回格式无效: {err}\nSTDOUT:\n{line}"))?;
    Ok(envelope)
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

pub fn run() {
    tauri::Builder::default()
        .manage(BackendState {
            process: Mutex::new(None),
        })
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![run_backend, reveal_path, open_path, read_text_file])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
