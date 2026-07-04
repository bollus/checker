use excel_check_tool_lib::rust_check::{run_check, CheckPayload};
use serde_json::json;
use std::fs;
use std::io::{self, Read};
use std::path::PathBuf;

fn main() {
    let result = run();
    match result {
        Ok(value) => println!("{}", serde_json::to_string(&value).unwrap()),
        Err(error) => {
            println!("{}", serde_json::to_string(&json!({ "ok": false, "errors": [error] })).unwrap());
            std::process::exit(1);
        }
    }
}

fn run() -> Result<serde_json::Value, String> {
    let input = if let Some(path) = std::env::args().nth(1) {
        fs::read_to_string(path).map_err(|err| format!("读取参数文件失败: {err}"))?
    } else {
        let mut input = String::new();
        io::stdin().read_to_string(&mut input).map_err(|err| format!("读取 stdin 失败: {err}"))?;
        input
    };
    let mut payload: CheckPayload = serde_json::from_str(&input).map_err(|err| format!("Rust 核对参数无效: {err}"))?;
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).parent().unwrap().to_path_buf();
    if payload.position_aliases_path.is_none() {
        let path = root.join("position_aliases.json");
        if path.is_file() {
            payload.position_aliases_path = Some(path.to_string_lossy().to_string());
        }
    }
    if payload.position_rules_path.is_none() {
        let path = root.join("position_rules.json");
        if path.is_file() {
            payload.position_rules_path = Some(path.to_string_lossy().to_string());
        }
    }
    let data = run_check(payload)?;
    Ok(json!({ "ok": true, "data": data }))
}
