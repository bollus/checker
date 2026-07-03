# 表格核对工具

用于核对主工资表和一个考勤表目录中的 Excel 文件，并把主表里不一致的单元格高亮出来。

## 新版跨平台桌面端

仓库已新增 `Tauri v2 + React + TypeScript + Python sidecar` 版本，用于后续替代旧 Tkinter 界面。

新版目标：

- Windows / macOS 均可打包成独立应用
- 用户电脑不需要安装 Python
- UI 使用左侧导航、主工作区、右侧结果面板的现代桌面布局
- 支持工资表核对、考勤表生成、核对模板管理
- 核对模板支持通过工资表/考勤表预览进行可视化点选创建规则

开发运行：

```bash
npm install
npm run tauri:dev
```

仅检查前端：

```bash
npm run build
```

Python sidecar 本地调试：

```bash
printf '%s' '{"action":"list_templates","payload":{}}' | python3 python_backend/backend_cli.py
```

打包说明：

- GitHub Actions 中的 `Build Tauri Desktop` 只会构建 Windows 和 macOS 桌面应用，不发布 Linux 包
- Windows 产物：`.msi` 和 `setup.exe`
- macOS 产物：`.dmg` 和 `.app.tar.gz`
- Actions 会先用 PyInstaller 把 `python_backend/backend_cli.py` 打包成 `excel-check-backend`
- Tauri 会把 sidecar、字体、模板和 Python 核心逻辑作为资源打进应用
- 本机手动打包需要安装 Node.js、Rust/Cargo、Python 3 和 PyInstaller；Windows/macOS 正式包建议直接使用 GitHub Actions 产物

说明：旧版 `excel_check_tool.py` Tkinter GUI 仍保留，方便回退和调试；新界面通过 sidecar 调用同一套 Excel 核心逻辑。

## 已实现规则

- 主表 `F7-*` 对比考勤表 `A3`
- 主表 `J7-*` 对比考勤表 `C3`
- 主表 `N7-*` 对比考勤表 `B6`
- 主表 `W7-*` 对比考勤表 `H9`
- 主表 `X7-*` 对比考勤表 `I9`
- 主表 `Y7-*` 对比考勤表 `J9`

匹配方式：

- 按主表 `A列 No.` 匹配考勤文件名中的编号前缀，例如 `25.Firoj Alam-Welder.xlsm`
- 文本字段会做空白清洗后比较
- 数字字段按数值比较，空白按 `0` 处理
- `岗位(Position)` 会做语义归一化比较，并读取 `position_aliases.json` 里的别名规则

## Windows 使用

前提：机器上安装了 Python 3。

1. 双击 `启动表格核对工具.bat`
2. GUI 会分成两个标签页：
   `工资表核对`、`根据表C生成表B`、`核对模板`
3. 在对应标签页选择文件和目录
4. 点击相应的开始按钮执行

输出：

- 一份带黄色高亮的新主表
- 一份同目录的 `*_核对报告.txt`

## 没有 Python 的电脑怎么用

不要在目标电脑上装 Python，直接把工具打包成独立 `exe` 发过去。

做法：

1. 在一台有 Python 3 的 Windows 电脑上打开本目录
2. 双击 `build_exe.bat`
3. 打包完成后，把 `dist\表格核对工具.exe` 发给没有 Python 的电脑
4. 对方直接双击 `表格核对工具.exe` 即可使用

说明：

- 这个 `exe` 是独立运行的，不要求目标电脑安装 Python
- 若杀毒软件拦截，需要把程序加入信任
- 首次打包时会自动安装 `PyInstaller`
- 当前仓库里已附带打包配置文件 `excel_check_tool.spec`
- 如果你仍想用中文文件名脚本，也可以继续双击 `打包EXE.bat`，它现在和 `build_exe.bat` 内容相同

## 命令行使用

```bash
python excel_check_tool.py --table-a "3A-ARA 1月人事代理第一次结算工资表.xlsx" --table-bs "考勤表-编辑版"
```

可选输出路径：

```bash
python excel_check_tool.py --table-a "主表.xlsx" --table-bs "考勤目录" --output "主表_核对结果.xlsx"
```

## 说明

- 当前支持 `.xlsx` 和 `.xlsm`
- 不依赖第三方库，Windows 上直接用 Python 标准库运行

## 自定义核对模板

核对模板配置文件：

- `check_templates.json`

你可以在 GUI 的 `核对模板` 标签页中直接维护模板，也可以手动编辑这个 JSON 文件。

GUI 已支持：

- 新建模板
- 复制模板
- 保存模板
- 删除模板
- 导入模板
- 导出模板
- 新增 / 编辑 / 删除规则
- 规则上移 / 下移

模板包含：

- `编号列`：主表中用于匹配表B文件编号的列，默认是 `A`
- `数据起始行`：逻辑上的第一行，默认是 `7`
- `规则列表`

每条规则包含：

- 字段名
- 主表范围，例如 `F7-Fn`
- 考勤表坐标或表达式，例如 `A3`、`SUM(G10:Gn)`、`SUM(H10:Hn,I10:In,J10:Jn)`
- 比较类型：`text` / `number` / `position`

示例：

- 主表 `F3-Fn` 对应 考勤表 `H4`
- 主表 `N7-Nn` 对应 考勤表 `SUM(G10:Gn)`
- 主表 `W7-Wn` 对应 考勤表 `SUM(H10:Hn,I10:In,J10:Jn)`

在 `工资表核对` 标签页选中某个模板后，程序会按这个模板中的规则执行核对。

## 根据表C生成表B

脚本：

- `generate_table_bs.py`

用途：

- 根据 `考勤表汇总.xlsx` 批量生成每个人的 `表B`
- 保留现有表B模板中的格式、固定文本、图片、宏及素材
- 当前做法是：复制一个现有表B作为模板，再把 `New timesheet` 工作表按表C内容静态填充

示例：

```bash
python generate_table_bs.py --table-c "考勤表汇总.xlsx" --table-bs-dir "考勤表-编辑版"
```

也可以显式指定模板：

```bash
python generate_table_bs.py --table-c "考勤表汇总.xlsx" --template-b "考勤表-编辑版/1.Venkateshan Varatharajan Varatharajan-Job Performer.xlsm"
```

可选输出目录：

```bash
python generate_table_bs.py --table-c "考勤表汇总.xlsx" --table-bs-dir "考勤表-编辑版" --output-dir "生成后的表Bs"
```

统计假期：

- GUI 中默认不勾选 `统计假期`
- 不勾选 `统计假期` 时，`B6` 应支付天数按 `E6 + I6 + J6 + K6 + L6` 计算
- 不勾选 `统计假期` 时，空白周末或法定假如果紧挨着 `A` 缺勤或 `E` 事假/紧急休假，则连续挨着的周末/法定假都不计入薪资天数；紧挨着 `S` 病假或 `V` 年休假仍计入
- 不勾选 `统计假期` 时，`V` 年休假计入 `J6` 后，同一天不再重复计入 `I6` 法定假或 `L6` 周末假；扣除时优先扣法定假，再扣周末假
- 勾选 `统计假期` 时，`B6` 应支付天数按 `E6 + I6 + K6 + L6` 计算
- 勾选后，`I6` 按法定假实际出勤天数统计，`L6` 按周末实际出勤天数统计
- 假期计数为 `0` 时留空
- 命令行可使用 `--count-holidays`

员工签名：

- 生成表B时会删除 `New timesheet` 的 `A42:G44` 范围内原模板签名图片，并插入新的员工签名图片
- 如果模板中存在 `Overtime` 工作表，会删除 `E53:J53` 范围内原模板签名图片，并插入新的员工签名图片
- 签名取员工姓名前两个单词，例如 `Salman Raza Raza Khan` 生成 `Salman Raza`
- 签名图片使用 `Nothing_You_Could_Do/NothingYouCouldDo-Regular.ttf` 渲染，不需要覆盖 `Requested by:` 或 `Applicant’s Signature:` 文本
- GUI 可用 `签名大小(%)` 调整签名大小，默认 `100`，可填 `30-200`
- 命令行可使用 `--signature-scale 120` 调整签名大小
- 运行源码需要安装 `Pillow`：`pip install pillow`；打包 EXE 时会自动安装并把字体文件打包进去

工作时间：

- 生成考勤表时可在 GUI 中自定义上午上/下班、下午上/下班和常规工作小时数
- 默认工作时间为 `06:00-12:00`、`14:00-18:00`，常规工作小时数为 `10`
- 如果一天正常工作 `8` 小时，可把常规工作小时数改为 `8`
- 命令行可使用 `--morning-start 08:00 --morning-end 12:00 --afternoon-start 13:00 --afternoon-end 17:00 --normal-hours 8`

新模板修正行：

- 如果表B模板中存在 `修正上月加班时长` 行，生成时会把汇总表 `BQ:BT` 写入主考勤表对应行的 `G:J`
- `BQ` 写入常规工作小时修正，`BR/BS/BT` 分别写入常规工作日/周末/法定假加班修正
- 主表 `G9:J9` 和加班表 `H43:J43` 的小时汇总会包含对应修正值
- 如果存在 `Overtime` 工作表，`BR/BS/BT` 会同步写入修正行的 `H:J`
- 如果汇总表日期列包含 31 号但模板月份只有 30 天，生成时会按模板月份天数自动忽略多余日期列

当前默认假设：

- `V` 年休假：计入可支付天数
- `S` 病假：计入可支付天数
- `E` 紧急休假：不计入可支付天数
- `A` 缺勤：不计入可支付天数
