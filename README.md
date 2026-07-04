# 表格核对工具

跨平台桌面工具，用于工资表核对、考勤表批量生成和核对模板维护。

当前正式界面是 `Tauri v2 + React + TypeScript + Rust`。Python 旧版已软移除：不再参与打包，也不会随安装包发布；旧代码保存在 `archive/legacy-python/`，只作为回退参考。

## 目录

- `src/`：React 前端界面
- `src-tauri/src/`：Tauri 命令和 Rust Excel 处理逻辑
- `check_templates.json`：默认核对模板
- `position_aliases.json` / `position_rules.json`：岗位语义归一化规则
- `Nothing_You_Could_Do/`：员工签名字体
- `fixtures/manual/`：手工测试用 Excel 样例，不参与应用打包
- `archive/legacy-python/`：旧 Python/Tkinter/PyInstaller 版本归档

## 开发

```bash
npm install
npm run tauri:dev
```

仅检查前端：

```bash
npm run build
```

检查 Rust：

```bash
cargo check --manifest-path src-tauri/Cargo.toml
cargo test --manifest-path src-tauri/Cargo.toml rust_ -- --nocapture
```

打包：

```bash
npm run tauri:build -- --bundles nsis
```

GitHub Actions 中的 `Build Tauri Desktop` 会构建 Windows 和 macOS 包。打包资源只包含模板、岗位规则和签名字体，不包含 Python。

## macOS 打开提示“已损坏”

GitHub Actions 产物不是 Apple Developer ID 签名并公证的正式发行包，macOS Gatekeeper 可能会提示“已损坏，无法打开”。当前工程会对 macOS 包做 ad-hoc 签名，但未公证的包首次打开仍可能需要手动移除隔离标记：

```bash
xattr -cr /Applications/ExcelCheckTool.app
open /Applications/ExcelCheckTool.app
```

如果是从 `.dmg` 里直接拖出来，也可以对拖出的 `.app` 执行同样命令。对外稳定分发需要配置 Apple Developer ID 证书并走 notarization。

## 功能

- 工资表核对：按模板规则对比主工资表和考勤表目录，输出高亮结果和报告。
- 生成考勤表：根据汇总表和考勤表模板批量生成员工考勤表，保留模板格式、图片、宏和固定文本。
- 核对模板：在 GUI 中维护字段、主表范围、考勤表坐标/表达式和比较方式。

模板表达式示例：

```text
F7-Fn
SUM(G10:Gn)
SUM(H10:Hn,I10:In,J10:Jn)
```

比较方式：

- `text`：文本
- `number`：数字
- `position`：岗位语义归一化

## 旧版归档

旧 Python 版本位于 `archive/legacy-python/`，包括：

- `excel_check_tool.py`
- `generate_table_bs.py`
- `python_backend/backend_cli.py`
- PyInstaller spec 和旧 bat 脚本
- 旧版 Windows EXE workflow 归档

这些文件不再被 Tauri 正式应用引用，也不会进入安装包。
