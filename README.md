# 表格核对工具

用于核对主工资表和一个考勤表目录中的 Excel 文件，并把主表里不一致的单元格高亮出来。

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

## Windows 使用

前提：机器上安装了 Python 3。

1. 双击 `启动表格核对工具.bat`
2. 选择主表文件
3. 选择考勤表目录
4. 选择输出位置
5. 点击“开始核对”

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
