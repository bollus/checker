use excel_check_tool_lib::rust_generate::{run_generate, GeneratePayload};
use serde_json::json;
use std::fs;
use std::io::{self, Read};
use std::path::PathBuf;

fn main() {
    match run() {
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
    let mut payload: GeneratePayload = serde_json::from_str(&input).map_err(|err| format!("Rust 生成参数无效: {err}"))?;
    if payload.signature_font_path.is_none() {
        let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).parent().unwrap().to_path_buf();
        let path = root.join("Nothing_You_Could_Do").join("NothingYouCouldDo-Regular.ttf");
        if path.is_file() {
            payload.signature_font_path = Some(path.to_string_lossy().to_string());
        }
    }
    let data = run_generate(payload)?;
    Ok(json!({ "ok": true, "data": data }))
}
