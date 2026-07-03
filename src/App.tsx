import { open, save } from "@tauri-apps/plugin-dialog";
import {
  AlertCircle,
  Archive,
  CheckCircle2,
  ChevronRight,
  Clock3,
  Copy,
  ExternalLink,
  FileCheck2,
  FileCog,
  FileSpreadsheet,
  FolderOpen,
  LayoutTemplate,
  Loader2,
  Play,
  Plus,
  Save,
  Search,
  Settings,
  Sparkles,
  Trash2,
  Wand2,
} from "lucide-react";
import { useEffect, useMemo, useState } from "react";
import {
  backend,
  CheckResult,
  CheckRule,
  CheckTemplate,
  CompareType,
  GenerateResult,
  Mismatch,
  openPath,
  revealPath,
  WorkbookPreview,
  WorkbookSheetPreview,
} from "./api";

type Page = "check" | "generate" | "templates" | "history" | "settings";
type BusyState = "idle" | "checking" | "generating" | "loading";

const DEFAULT_TEMPLATE: CheckTemplate = {
  name: "默认模板",
  number_column: "A",
  start_row: 7,
  rules: [
    { field_name: "姓名", main_range: "F7-Fn", table_b_cell: "A3", compare_type: "text" },
    { field_name: "岗位", main_range: "J7-Jn", table_b_cell: "C3", compare_type: "position" },
    { field_name: "应支付天数", main_range: "N7-Nn", table_b_cell: "B6", compare_type: "number" },
    { field_name: "正常工作日加班", main_range: "W7-Wn", table_b_cell: "H9", compare_type: "number" },
    { field_name: "周末加班", main_range: "X7-Xn", table_b_cell: "I9", compare_type: "number" },
    { field_name: "法定假日加班", main_range: "Y7-Yn", table_b_cell: "J9", compare_type: "number" },
  ],
};

const compareLabels: Record<CompareType, string> = {
  text: "文本",
  number: "数字",
  position: "岗位",
};

function fileName(path: string) {
  return path.split(/[\\/]/).filter(Boolean).pop() || path || "未选择";
}

function colFromRef(ref: string) {
  return ref.replace(/\d+/g, "");
}

function rowFromRef(ref: string) {
  return Number(ref.replace(/[A-Z]+/g, ""));
}

function rangeFromCell(ref: string) {
  const col = colFromRef(ref);
  const row = rowFromRef(ref);
  return `${col}${row}-${col}n`;
}

function shortPath(path: string) {
  if (!path) return "尚未选择";
  const parts = path.split(/[\\/]/).filter(Boolean);
  if (parts.length <= 3) return path;
  return `${parts[0]}/.../${parts.slice(-2).join("/")}`;
}

function useLocalState<T>(key: string, initial: T) {
  const [value, setValue] = useState<T>(() => {
    const raw = localStorage.getItem(key);
    if (!raw) return initial;
    try {
      return JSON.parse(raw) as T;
    } catch {
      return initial;
    }
  });

  useEffect(() => {
    localStorage.setItem(key, JSON.stringify(value));
  }, [key, value]);

  return [value, setValue] as const;
}

function EmptyPath({ title, detail }: { title: string; detail: string }) {
  return (
    <div className="path-empty">
      <FileSpreadsheet size={18} />
      <div>
        <strong>{title}</strong>
        <span>{detail}</span>
      </div>
    </div>
  );
}

function PathPicker({
  label,
  value,
  kind,
  filters,
  onChange,
}: {
  label: string;
  value: string;
  kind: "file" | "folder" | "save";
  filters?: { name: string; extensions: string[] }[];
  onChange: (value: string) => void;
}) {
  async function choose() {
    if (kind === "save") {
      const selected = await save({ filters });
      if (selected) onChange(selected);
      return;
    }
    const selected = await open({
      directory: kind === "folder",
      multiple: false,
      filters,
    });
    if (typeof selected === "string") onChange(selected);
  }

  return (
    <div className="field-card">
      <div>
        <label>{label}</label>
        <p>{shortPath(value)}</p>
      </div>
      <button className="icon-button" type="button" onClick={choose} title="选择">
        <FolderOpen size={18} />
      </button>
    </div>
  );
}

function ResultPanel({
  checkResult,
  generateResult,
  busy,
  log,
}: {
  checkResult: CheckResult | null;
  generateResult: GenerateResult | null;
  busy: BusyState;
  log: string[];
}) {
  const latestPath = checkResult?.output_path || generateResult?.output_dir || "";
  const reportPath = checkResult?.report_path || generateResult?.report_path || "";
  return (
    <aside className="details-pane">
      <div className="pane-title">
        <span>任务详情</span>
        <Clock3 size={17} />
      </div>
      <div className="result-card primary">
        {busy !== "idle" ? (
          <>
            <Loader2 className="spin" size={30} />
            <strong>正在处理</strong>
            <span>Excel 文件会在后台完成读取、核对和写入。</span>
          </>
        ) : checkResult ? (
          <>
            <CheckCircle2 size={30} />
            <strong>{checkResult.mismatch_count} 个不一致</strong>
            <span>{fileName(checkResult.output_path)}</span>
          </>
        ) : generateResult ? (
          <>
            <CheckCircle2 size={30} />
            <strong>生成 {generateResult.generated_count} 份</strong>
            <span>{fileName(generateResult.output_dir)}</span>
          </>
        ) : (
          <>
            <Sparkles size={30} />
            <strong>准备就绪</strong>
            <span>选择文件后开始任务，结果会显示在这里。</span>
          </>
        )}
      </div>

      <div className="action-stack">
        <button disabled={!latestPath} onClick={() => openPath(latestPath)}>
          <ExternalLink size={16} />
          打开结果
        </button>
        <button disabled={!latestPath} onClick={() => revealPath(latestPath)}>
          <FolderOpen size={16} />
          定位文件
        </button>
        <button disabled={!reportPath} onClick={() => openPath(reportPath)}>
          <FileCheck2 size={16} />
          打开报告
        </button>
      </div>

      <div className="pane-title small">
        <span>运行记录</span>
      </div>
      <div className="log-list">
        {log.length === 0 ? <span className="muted">暂无记录</span> : log.slice(-8).map((item, index) => <p key={index}>{item}</p>)}
      </div>
    </aside>
  );
}

function MismatchTable({ items }: { items: Mismatch[] }) {
  if (!items.length) {
    return (
      <div className="success-strip">
        <CheckCircle2 size={18} />
        未发现不一致。
      </div>
    );
  }
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>行</th>
            <th>字段</th>
            <th>主表</th>
            <th>考勤表</th>
            <th>文件</th>
          </tr>
        </thead>
        <tbody>
          {items.slice(0, 120).map((item, index) => (
            <tr key={`${item.table_a_cell}-${index}`}>
              <td>{item.table_a_cell}</td>
              <td>{item.field_name}</td>
              <td>{item.table_a_value || "空"}</td>
              <td>{item.table_b_value || "空"}</td>
              <td>{item.table_b_file}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function CheckPage({
  templates,
  addLog,
  busy,
  setBusy,
  onResult,
}: {
  templates: CheckTemplate[];
  addLog: (text: string) => void;
  busy: BusyState;
  setBusy: (state: BusyState) => void;
  onResult: (result: CheckResult) => void;
}) {
  const [tableA, setTableA] = useLocalState("check.tableA", "");
  const [tableBs, setTableBs] = useLocalState("check.tableBs", "");
  const [output, setOutput] = useLocalState("check.output", "");
  const [templateName, setTemplateName] = useLocalState("check.template", "默认模板");
  const [result, setResult] = useState<CheckResult | null>(null);

  const selectedTemplate = templates.find((item) => item.name === templateName) || templates[0] || DEFAULT_TEMPLATE;

  async function run() {
    if (!tableA || !tableBs) {
      addLog("请先选择工资表和考勤表目录。");
      return;
    }
    setBusy("checking");
    try {
      addLog(`开始核对: ${fileName(tableA)}`);
      const data = await backend<CheckResult>("check", {
        table_a_path: tableA,
        table_bs_folder: tableBs,
        output_path: output || null,
        template: selectedTemplate,
      });
      setResult(data);
      onResult(data);
      addLog(`核对完成: ${data.mismatch_count} 个不一致`);
    } catch (error) {
      addLog(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy("idle");
    }
  }

  return (
    <section className="workspace">
      <div className="hero-row">
        <div>
          <div className="breadcrumb">工资表核对 <ChevronRight size={16} /> 快速任务</div>
          <h1>选择文件，直接核对</h1>
          <p>保留原始 Excel 格式，只在结果工资表中高亮不一致单元格。</p>
        </div>
        <button className="primary-button" disabled={busy !== "idle"} onClick={run}>
          {busy === "checking" ? <Loader2 className="spin" size={18} /> : <Play size={18} />}
          开始核对
        </button>
      </div>

      <div className="input-grid">
        <PathPicker label="主工资表" value={tableA} kind="file" filters={[{ name: "Excel", extensions: ["xlsx", "xlsm"] }]} onChange={setTableA} />
        <PathPicker label="考勤表目录" value={tableBs} kind="folder" onChange={setTableBs} />
        <PathPicker label="结果另存为" value={output} kind="save" filters={[{ name: "Excel", extensions: ["xlsx", "xlsm"] }]} onChange={setOutput} />
        <div className="field-card">
          <div>
            <label>核对模板</label>
            <select value={selectedTemplate.name} onChange={(event) => setTemplateName(event.target.value)}>
              {templates.map((template) => (
                <option key={template.name} value={template.name}>
                  {template.name}
                </option>
              ))}
            </select>
          </div>
          <LayoutTemplate size={18} />
        </div>
      </div>

      <div className="section-head">
        <h2>模板规则</h2>
        <span>{selectedTemplate.rules.length} 条规则</span>
      </div>
      <div className="rule-chips">
        {selectedTemplate.rules.map((rule) => (
          <div className="rule-chip" key={`${rule.field_name}-${rule.main_range}`}>
            <strong>{rule.field_name}</strong>
            <span>{rule.main_range} → {rule.table_b_cell}</span>
          </div>
        ))}
      </div>

      {result ? (
        <>
          <div className="section-head">
            <h2>核对明细</h2>
            <span>{result.mismatch_count} 个不一致</span>
          </div>
          <MismatchTable items={result.mismatches} />
        </>
      ) : (
        <EmptyPath title="等待核对" detail="完成后这里会显示不一致明细和结果路径。" />
      )}
    </section>
  );
}

function GeneratePage({
  busy,
  setBusy,
  addLog,
  onResult,
}: {
  busy: BusyState;
  setBusy: (state: BusyState) => void;
  addLog: (text: string) => void;
  onResult: (result: GenerateResult) => void;
}) {
  const [tableC, setTableC] = useLocalState("generate.tableC", "");
  const [templateB, setTemplateB] = useLocalState("generate.templateB", "");
  const [outputDir, setOutputDir] = useLocalState("generate.outputDir", "");
  const [countHolidays, setCountHolidays] = useLocalState("generate.countHolidays", false);
  const [signatureScale, setSignatureScale] = useLocalState("generate.signatureScale", 100);
  const [normalHours, setNormalHours] = useLocalState("generate.normalHours", "10");
  const [morningStart, setMorningStart] = useLocalState("generate.morningStart", "06:00");
  const [morningEnd, setMorningEnd] = useLocalState("generate.morningEnd", "12:00");
  const [afternoonStart, setAfternoonStart] = useLocalState("generate.afternoonStart", "14:00");
  const [afternoonEnd, setAfternoonEnd] = useLocalState("generate.afternoonEnd", "18:00");

  async function run() {
    if (!tableC || !templateB) {
      addLog("请先选择汇总表和考勤表模板。");
      return;
    }
    setBusy("generating");
    try {
      addLog(`开始生成: ${fileName(tableC)}`);
      const data = await backend<GenerateResult>("generate", {
        table_c_path: tableC,
        template_b_path: templateB,
        output_dir: outputDir || null,
        count_holidays: countHolidays,
        signature_scale: signatureScale,
        morning_start: morningStart,
        morning_end: morningEnd,
        afternoon_start: afternoonStart,
        afternoon_end: afternoonEnd,
        normal_hours: normalHours,
      });
      onResult(data);
      addLog(`生成完成: ${data.generated_count} 份`);
    } catch (error) {
      addLog(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy("idle");
    }
  }

  return (
    <section className="workspace">
      <div className="hero-row">
        <div>
          <div className="breadcrumb">根据表C生成表B <ChevronRight size={16} /> 批量生成</div>
          <h1>从汇总表生成考勤表</h1>
          <p>复制模板格式、固定文本、素材和宏，只填充员工与每日考勤数据。</p>
        </div>
        <button className="primary-button" disabled={busy !== "idle"} onClick={run}>
          {busy === "generating" ? <Loader2 className="spin" size={18} /> : <Wand2 size={18} />}
          开始生成
        </button>
      </div>

      <div className="input-grid">
        <PathPicker label="考勤汇总表 C" value={tableC} kind="file" filters={[{ name: "Excel", extensions: ["xlsx", "xlsm"] }]} onChange={setTableC} />
        <PathPicker label="考勤表模板 B" value={templateB} kind="file" filters={[{ name: "Excel", extensions: ["xlsx", "xlsm"] }]} onChange={setTemplateB} />
        <PathPicker label="输出目录" value={outputDir} kind="folder" onChange={setOutputDir} />
        <div className="field-card compact">
          <label>统计假期</label>
          <button className={countHolidays ? "toggle on" : "toggle"} onClick={() => setCountHolidays(!countHolidays)}>
            {countHolidays ? "已开启" : "默认关闭"}
          </button>
        </div>
      </div>

      <div className="settings-grid">
        <label>
          签名大小
          <input type="number" min={30} max={200} value={signatureScale} onChange={(event) => setSignatureScale(Number(event.target.value))} />
        </label>
        <label>
          常规小时
          <input value={normalHours} onChange={(event) => setNormalHours(event.target.value)} />
        </label>
        <label>
          上午上班
          <input value={morningStart} onChange={(event) => setMorningStart(event.target.value)} />
        </label>
        <label>
          上午下班
          <input value={morningEnd} onChange={(event) => setMorningEnd(event.target.value)} />
        </label>
        <label>
          下午上班
          <input value={afternoonStart} onChange={(event) => setAfternoonStart(event.target.value)} />
        </label>
        <label>
          下午下班
          <input value={afternoonEnd} onChange={(event) => setAfternoonEnd(event.target.value)} />
        </label>
      </div>
    </section>
  );
}

function SheetPreview({
  title,
  preview,
  selected,
  onSelect,
}: {
  title: string;
  preview: WorkbookPreview | null;
  selected: string;
  onSelect: (ref: string) => void;
}) {
  const sheet: WorkbookSheetPreview | undefined = preview?.sheets[0];
  const cells = useMemo(() => new Map(sheet?.cells.map((cell) => [cell.ref, cell.value]) || []), [sheet]);
  const rows = Math.min(sheet?.bounds.max_row || 16, 36);
  const cols = Math.min(sheet?.bounds.max_col || 10, 18);

  return (
    <div className="sheet-panel">
      <div className="sheet-title">
        <strong>{title}</strong>
        <span>{preview ? `${fileName(preview.path)} · ${sheet?.name || ""}` : "未加载"}</span>
      </div>
      <div className="sheet-grid" style={{ gridTemplateColumns: `46px repeat(${cols}, minmax(88px, 1fr))` }}>
        <div className="sheet-corner" />
        {Array.from({ length: cols }, (_, index) => (
          <div className="sheet-header" key={`h-${index}`}>{String.fromCharCode(65 + index)}</div>
        ))}
        {Array.from({ length: rows }, (_, rowIndex) => {
          const row = rowIndex + 1;
          return [
            <div className="sheet-row-head" key={`r-${row}`}>{row}</div>,
            ...Array.from({ length: cols }, (_, colIndex) => {
              const col = String.fromCharCode(65 + colIndex);
              const ref = `${col}${row}`;
              return (
                <button
                  className={selected === ref ? "sheet-cell selected" : "sheet-cell"}
                  key={ref}
                  onClick={() => onSelect(ref)}
                  title={ref}
                >
                  {cells.get(ref) || ""}
                </button>
              );
            }),
          ];
        })}
      </div>
    </div>
  );
}

function TemplatesPage({
  templates,
  setTemplates,
  addLog,
}: {
  templates: CheckTemplate[];
  setTemplates: (templates: CheckTemplate[]) => void;
  addLog: (text: string) => void;
}) {
  const [current, setCurrent] = useState<CheckTemplate>(templates[0] || DEFAULT_TEMPLATE);
  const [mainWorkbook, setMainWorkbook] = useState("");
  const [tableBWorkbook, setTableBWorkbook] = useState("");
  const [mainPreview, setMainPreview] = useState<WorkbookPreview | null>(null);
  const [tableBPreview, setTableBPreview] = useState<WorkbookPreview | null>(null);
  const [mainCell, setMainCell] = useState("");
  const [tableBCell, setTableBCell] = useState("");
  const [fieldName, setFieldName] = useState("新字段");
  const [compareType, setCompareType] = useState<CompareType>("number");
  const [useSum, setUseSum] = useState(false);

  useEffect(() => {
    if (templates.length && !templates.some((item) => item.name === current.name)) {
      setCurrent(templates[0]);
    }
  }, [templates, current.name]);

  async function inspect(path: string, side: "main" | "b") {
    if (!path) return;
    const data = await backend<WorkbookPreview>("inspect_workbook", { path, max_rows: 80, max_cols: 35 });
    if (side === "main") setMainPreview(data);
    else setTableBPreview(data);
  }

  async function chooseWorkbook(side: "main" | "b") {
    const selected = await open({ multiple: false, filters: [{ name: "Excel", extensions: ["xlsx", "xlsm"] }] });
    if (typeof selected !== "string") return;
    if (side === "main") setMainWorkbook(selected);
    else setTableBWorkbook(selected);
    await inspect(selected, side);
  }

  function addRule() {
    if (!mainCell || !tableBCell) {
      addLog("请先在两个预览表中选择单元格。");
      return;
    }
    const target = useSum ? `SUM(${tableBCell}:${colFromRef(tableBCell)}N)` : tableBCell;
    const rule: CheckRule = {
      field_name: fieldName.trim() || "新字段",
      main_range: rangeFromCell(mainCell),
      table_b_cell: target,
      compare_type: compareType,
    };
    setCurrent({ ...current, rules: [...current.rules, rule] });
    addLog(`已添加规则: ${rule.main_range} → ${rule.table_b_cell}`);
  }

  async function saveTemplate() {
    const data = await backend<{ template: CheckTemplate }>("save_template", { template: current });
    const next = templates.filter((item) => item.name !== data.template.name);
    setTemplates([...next, data.template]);
    addLog(`已保存模板: ${data.template.name}`);
  }

  function duplicateTemplate() {
    const name = `${current.name} - 副本`;
    setCurrent({ ...current, name });
  }

  async function deleteTemplate() {
    await backend("delete_template", { name: current.name });
    const next = templates.filter((item) => item.name !== current.name);
    setTemplates(next.length ? next : [DEFAULT_TEMPLATE]);
    setCurrent(next[0] || DEFAULT_TEMPLATE);
    addLog(`已删除模板: ${current.name}`);
  }

  return (
    <section className="workspace template-workspace">
      <div className="hero-row">
        <div>
          <div className="breadcrumb">核对模板 <ChevronRight size={16} /> 可视化点选</div>
          <h1>像选单元格一样创建规则</h1>
          <p>左边点工资表字段，右边点考勤表字段，规则会自动转换成模板表达式。</p>
        </div>
        <div className="toolbar">
          <button onClick={duplicateTemplate}><Copy size={16} />复制</button>
          <button onClick={deleteTemplate}><Trash2 size={16} />删除</button>
          <button className="primary-button slim" onClick={saveTemplate}><Save size={16} />保存</button>
        </div>
      </div>

      <div className="template-top">
        <div className="field-card">
          <div>
            <label>当前模板</label>
            <select value={current.name} onChange={(event) => setCurrent(templates.find((item) => item.name === event.target.value) || current)}>
              {templates.map((template) => (
                <option key={template.name} value={template.name}>{template.name}</option>
              ))}
            </select>
          </div>
          <LayoutTemplate size={18} />
        </div>
        <label className="inline-input">
          模板名称
          <input value={current.name} onChange={(event) => setCurrent({ ...current, name: event.target.value })} />
        </label>
        <label className="inline-input">
          编号列
          <input value={current.number_column} onChange={(event) => setCurrent({ ...current, number_column: event.target.value.toUpperCase() })} />
        </label>
        <label className="inline-input">
          起始行
          <input type="number" value={current.start_row} onChange={(event) => setCurrent({ ...current, start_row: Number(event.target.value) })} />
        </label>
      </div>

      <div className="preview-actions">
        <button onClick={() => chooseWorkbook("main")}><FileSpreadsheet size={16} />选择工资表示例</button>
        <button onClick={() => chooseWorkbook("b")}><FileSpreadsheet size={16} />选择考勤表示例</button>
        <span>{fileName(mainWorkbook)} → {fileName(tableBWorkbook)}</span>
      </div>

      <div className="visual-picker">
        <SheetPreview title="主工资表" preview={mainPreview} selected={mainCell} onSelect={setMainCell} />
        <SheetPreview title="考勤表" preview={tableBPreview} selected={tableBCell} onSelect={setTableBCell} />
      </div>

      <div className="rule-builder">
        <label>
          字段名
          <input value={fieldName} onChange={(event) => setFieldName(event.target.value)} />
        </label>
        <label>
          比较类型
          <select value={compareType} onChange={(event) => setCompareType(event.target.value as CompareType)}>
            <option value="number">数字</option>
            <option value="text">文本</option>
            <option value="position">岗位</option>
          </select>
        </label>
        <button className={useSum ? "toggle on" : "toggle"} onClick={() => setUseSum(!useSum)}>
          {useSum ? "SUM 范围" : "单元格"}
        </button>
        <div className="rule-preview">
          {mainCell ? rangeFromCell(mainCell) : "主表未选"} → {tableBCell ? (useSum ? `SUM(${tableBCell}:${colFromRef(tableBCell)}N)` : tableBCell) : "考勤表未选"}
        </div>
        <button className="primary-button slim" onClick={addRule}><Plus size={16} />添加规则</button>
      </div>

      <div className="section-head">
        <h2>规则列表</h2>
        <span>{current.rules.length} 条</span>
      </div>
      <div className="rule-table">
        {current.rules.map((rule, index) => (
          <div className="rule-row" key={`${rule.field_name}-${index}`}>
            <strong>{rule.field_name}</strong>
            <span>{rule.main_range}</span>
            <span>{rule.table_b_cell}</span>
            <em>{compareLabels[rule.compare_type]}</em>
            <button onClick={() => setCurrent({ ...current, rules: current.rules.filter((_, itemIndex) => itemIndex !== index) })}>
              <Trash2 size={15} />
            </button>
          </div>
        ))}
      </div>
    </section>
  );
}

function StaticPage({ type }: { type: "history" | "settings" }) {
  return (
    <section className="workspace">
      <div className="hero-row">
        <div>
          <div className="breadcrumb">{type === "history" ? "历史记录" : "设置"}</div>
          <h1>{type === "history" ? "最近任务会自动保存在本机" : "默认参数和路径偏好"}</h1>
          <p>{type === "history" ? "当前版本先记录运行日志，后续可扩展为完整任务历史。" : "当前版本使用本机 localStorage 记住常用路径和参数。"}</p>
        </div>
      </div>
      <EmptyPath title="功能已预留" detail="主流程、生成和模板管理已经可用；这里用于后续增强。" />
    </section>
  );
}

export default function App() {
  const [page, setPage] = useState<Page>("check");
  const [busy, setBusy] = useState<BusyState>("idle");
  const [templates, setTemplates] = useState<CheckTemplate[]>([DEFAULT_TEMPLATE]);
  const [log, setLog] = useState<string[]>([]);
  const [checkResult, setCheckResult] = useState<CheckResult | null>(null);
  const [generateResult, setGenerateResult] = useState<GenerateResult | null>(null);

  function addLog(text: string) {
    const time = new Date().toLocaleTimeString();
    setLog((items) => [...items, `${time}  ${text}`]);
  }

  useEffect(() => {
    backend<{ templates: CheckTemplate[] }>("list_templates", {})
      .then((data) => setTemplates(data.templates.length ? data.templates : [DEFAULT_TEMPLATE]))
      .catch((error) => addLog(error instanceof Error ? error.message : String(error)));
  }, []);

  const nav = [
    { id: "check" as Page, label: "工资表核对", icon: FileCheck2 },
    { id: "generate" as Page, label: "生成考勤表", icon: Archive },
    { id: "templates" as Page, label: "核对模板", icon: LayoutTemplate },
    { id: "history" as Page, label: "历史记录", icon: Clock3 },
    { id: "settings" as Page, label: "设置", icon: Settings },
  ];

  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-mark">E</div>
          <span>ExcelCheck</span>
        </div>
        <div className="search-box">
          <Search size={16} />
          <span>Search...</span>
        </div>
        <nav>
          {nav.map((item) => {
            const Icon = item.icon;
            return (
              <button className={page === item.id ? "active" : ""} key={item.id} onClick={() => setPage(item.id)}>
                <Icon size={18} />
                {item.label}
              </button>
            );
          })}
        </nav>
        <div className="tip-card">
          <AlertCircle size={18} />
          <button>本机运行</button>
          <span>文件只在你的电脑上处理，不上传。</span>
        </div>
      </aside>

      <div className="content-area">
        {page === "check" && (
          <CheckPage
            templates={templates}
            busy={busy}
            setBusy={setBusy}
            addLog={addLog}
            onResult={(result) => {
              setCheckResult(result);
              setGenerateResult(null);
            }}
          />
        )}
        {page === "generate" && (
          <GeneratePage
            busy={busy}
            setBusy={setBusy}
            addLog={addLog}
            onResult={(result) => {
              setGenerateResult(result);
              setCheckResult(null);
            }}
          />
        )}
        {page === "templates" && <TemplatesPage templates={templates} setTemplates={setTemplates} addLog={addLog} />}
        {page === "history" && <StaticPage type="history" />}
        {page === "settings" && <StaticPage type="settings" />}
      </div>

      <ResultPanel checkResult={checkResult} generateResult={generateResult} busy={busy} log={log} />
    </main>
  );
}
