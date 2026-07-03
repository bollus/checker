import { open, save } from "@tauri-apps/plugin-dialog";
import {
  AlertCircle,
  Archive,
  CheckCircle2,
  Clock3,
  Copy,
  Download,
  ExternalLink,
  FileCheck2,
  FileCog,
  FileSpreadsheet,
  FolderOpen,
  History,
  Import,
  LayoutTemplate,
  Loader2,
  Play,
  Plus,
  RotateCcw,
  Save,
  Search,
  Settings,
  Trash2,
  Upload,
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
} from "./api";

type Page = "check" | "generate" | "templates" | "history" | "settings";
type BusyState = "idle" | "checking" | "generating";
type TaskStatus = "成功" | "失败" | "有警告";
type TaskType = "工资表核对" | "生成考勤表";

interface HistoryItem {
  id: string;
  time: string;
  type: TaskType;
  source: string;
  template: string;
  outputPath: string;
  reportPath: string;
  status: TaskStatus;
  mismatchCount?: number;
  generatedCount?: number;
  mismatches?: Mismatch[];
}

interface AppSettings {
  defaultPayrollDir: string;
  defaultAttendanceDir: string;
  defaultOutputDir: string;
  countHolidays: boolean;
  signatureScale: number;
  normalHours: string;
  morningStart: string;
  morningEnd: string;
  afternoonStart: string;
  afternoonEnd: string;
  autoOpenResult: boolean;
  keepLogDays: string;
}

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

const DEFAULT_SETTINGS: AppSettings = {
  defaultPayrollDir: "",
  defaultAttendanceDir: "",
  defaultOutputDir: "",
  countHolidays: false,
  signatureScale: 100,
  normalHours: "10",
  morningStart: "06:00",
  morningEnd: "12:00",
  afternoonStart: "14:00",
  afternoonEnd: "18:00",
  autoOpenResult: false,
  keepLogDays: "30",
};

const compareLabels: Record<CompareType, string> = {
  text: "文本",
  number: "数字",
  position: "岗位",
};

const compareOptions: CompareType[] = ["text", "number", "position"];

function fileName(path: string) {
  return path.split(/[\\/]/).filter(Boolean).pop() || path || "未选择";
}

function shortPath(path: string) {
  if (!path) return "尚未选择";
  const parts = path.split(/[\\/]/).filter(Boolean);
  if (parts.length <= 3) return path;
  return `${parts[0]}/.../${parts.slice(-2).join("/")}`;
}

function nowText() {
  return new Date().toLocaleString();
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

async function choosePath(kind: "file" | "folder" | "save", extensions?: string[]) {
  if (kind === "save") {
    return await save({ filters: extensions ? [{ name: "文件", extensions }] : undefined });
  }
  const selected = await open({
    directory: kind === "folder",
    multiple: false,
    filters: extensions ? [{ name: "文件", extensions }] : undefined,
  });
  return typeof selected === "string" ? selected : null;
}

function PageHeader({
  title,
  subtitle,
  actions,
}: {
  title: string;
  subtitle: string;
  actions?: React.ReactNode;
}) {
  return (
    <div className="page-header">
      <div>
        <h1>{title}</h1>
        <p>{subtitle}</p>
      </div>
      {actions ? <div className="header-actions">{actions}</div> : null}
    </div>
  );
}

function PathRow({
  label,
  value,
  kind,
  placeholder,
  extensions,
  onChange,
}: {
  label: string;
  value: string;
  kind: "file" | "folder" | "save";
  placeholder?: string;
  extensions?: string[];
  onChange: (value: string) => void;
}) {
  async function pick() {
    const selected = await choosePath(kind, extensions);
    if (selected) onChange(selected);
  }

  return (
    <div className="form-row">
      <label>{label}</label>
      <input value={value} placeholder={placeholder || "请选择"} onChange={(event) => onChange(event.target.value)} />
      <button className="icon-button" type="button" onClick={pick} title="选择">
        <FolderOpen size={17} />
      </button>
    </div>
  );
}

function Toggle({
  checked,
  onChange,
  label,
}: {
  checked: boolean;
  onChange: (value: boolean) => void;
  label: string;
}) {
  return (
    <label className="check-label">
      <input type="checkbox" checked={checked} onChange={(event) => onChange(event.target.checked)} />
      <span>{label}</span>
    </label>
  );
}

function DetailPanel({
  page,
  busy,
  selectedTemplate,
  checkResult,
  generateResult,
  selectedHistory,
  log,
}: {
  page: Page;
  busy: BusyState;
  selectedTemplate: CheckTemplate;
  checkResult: CheckResult | null;
  generateResult: GenerateResult | null;
  selectedHistory: HistoryItem | null;
  log: string[];
}) {
  const latestPath = selectedHistory?.outputPath || checkResult?.output_path || generateResult?.output_dir || "";
  const reportPath = selectedHistory?.reportPath || checkResult?.report_path || generateResult?.report_path || "";
  const title = page === "templates" ? "模板状态" : page === "settings" ? "环境状态" : "任务详情";

  return (
    <aside className="details-pane">
      <div className="pane-title">
        <span>{title}</span>
        <Clock3 size={16} />
      </div>

      <div className="status-card">
        {busy !== "idle" ? <Loader2 className="spin" size={28} /> : <CheckCircle2 size={28} />}
        <strong>{busy === "checking" ? "正在核对" : busy === "generating" ? "正在生成" : "准备就绪"}</strong>
        <span>{page === "templates" ? `${selectedTemplate.rules.length} 条规则` : "文件只在本机处理"}</span>
      </div>

      {page === "templates" ? (
        <div className="metric-list">
          <div><span>当前模板</span><strong>{selectedTemplate.name}</strong></div>
          <div><span>编号列</span><strong>{selectedTemplate.number_column}</strong></div>
          <div><span>起始行</span><strong>{selectedTemplate.start_row}</strong></div>
          <div><span>模板校验</span><strong>保存时执行</strong></div>
        </div>
      ) : page === "settings" ? (
        <div className="metric-list">
          <div><span>Python sidecar</span><strong>已配置</strong></div>
          <div><span>签名字体</span><strong>已随包发布</strong></div>
          <div><span>Tauri</span><strong>运行正常</strong></div>
          <div><span>数据上传</span><strong>不上传</strong></div>
        </div>
      ) : (
        <div className="metric-list">
          <div><span>模板</span><strong>{selectedTemplate.name}</strong></div>
          <div><span>核对结果</span><strong>{checkResult ? `${checkResult.mismatch_count} 个不一致` : "未开始"}</strong></div>
          <div><span>生成结果</span><strong>{generateResult ? `${generateResult.generated_count} 份` : "未开始"}</strong></div>
        </div>
      )}

      <div className="action-stack">
        <button disabled={!latestPath} onClick={() => openPath(latestPath)}>
          <ExternalLink size={16} />打开结果
        </button>
        <button disabled={!latestPath} onClick={() => revealPath(latestPath)}>
          <FolderOpen size={16} />定位文件
        </button>
        <button disabled={!reportPath} onClick={() => openPath(reportPath)}>
          <FileCheck2 size={16} />打开报告
        </button>
      </div>

      <div className="examples-box">
        <strong>表达式示例</strong>
        <code>F3-Fn</code>
        <code>SUM(H10:Hn,I10:In)</code>
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

function RulePreviewTable({ rules }: { rules: CheckRule[] }) {
  return (
    <div className="data-table">
      <table>
        <thead>
          <tr>
            <th>字段</th>
            <th>主表范围</th>
            <th>考勤表</th>
            <th>类型</th>
          </tr>
        </thead>
        <tbody>
          {rules.map((rule, index) => (
            <tr key={`${rule.field_name}-${index}`}>
              <td>{rule.field_name}</td>
              <td>{rule.main_range}</td>
              <td>{rule.table_b_cell}</td>
              <td>{compareLabels[rule.compare_type]}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function MismatchTable({ items }: { items: Mismatch[] }) {
  if (!items.length) {
    return <div className="empty-state"><CheckCircle2 size={18} />未发现不一致。</div>;
  }
  return (
    <div className="data-table">
      <table>
        <thead>
          <tr>
            <th>行</th>
            <th>字段</th>
            <th>主表值</th>
            <th>考勤表值</th>
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
  settings,
  busy,
  setBusy,
  addLog,
  addHistory,
  onResult,
}: {
  templates: CheckTemplate[];
  settings: AppSettings;
  busy: BusyState;
  setBusy: (state: BusyState) => void;
  addLog: (text: string) => void;
  addHistory: (item: HistoryItem) => void;
  onResult: (result: CheckResult) => void;
}) {
  const [tableA, setTableA] = useLocalState("check.tableA", "");
  const [tableBs, setTableBs] = useLocalState("check.tableBs", "");
  const [output, setOutput] = useLocalState("check.output", "");
  const [templateName, setTemplateName] = useLocalState("check.template", "默认模板");
  const [autoOpen, setAutoOpen] = useLocalState("check.autoOpen", settings.autoOpenResult);
  const [reportOnly, setReportOnly] = useLocalState("check.reportOnly", false);
  const [overwrite, setOverwrite] = useLocalState("check.overwrite", false);
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
      addHistory({
        id: crypto.randomUUID(),
        time: nowText(),
        type: "工资表核对",
        source: tableA,
        template: selectedTemplate.name,
        outputPath: data.output_path,
        reportPath: data.report_path,
        status: data.warnings.length ? "有警告" : "成功",
        mismatchCount: data.mismatch_count,
        mismatches: data.mismatches,
      });
      addLog(`核对完成: ${data.mismatch_count} 个不一致`);
      if (autoOpen) await openPath(data.output_path);
    } catch (error) {
      addLog(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy("idle");
    }
  }

  return (
    <section className="workspace">
      <PageHeader
        title="工资表核对"
        subtitle="选择主表、考勤表目录和模板，一键输出高亮结果。"
        actions={
          <>
            <button className="secondary-button" onClick={() => { setTableA(""); setTableBs(""); setOutput(""); }}>
              <RotateCcw size={16} />清空
            </button>
            <button className="primary-button" disabled={busy !== "idle"} onClick={run}>
              {busy === "checking" ? <Loader2 className="spin" size={18} /> : <Play size={18} />}开始核对
            </button>
          </>
        }
      />

      <div className="panel">
        <PathRow label="主工资表" value={tableA} kind="file" extensions={["xlsx", "xlsm"]} placeholder={settings.defaultPayrollDir || "选择工资表文件"} onChange={setTableA} />
        <PathRow label="考勤表目录" value={tableBs} kind="folder" placeholder={settings.defaultAttendanceDir || "选择考勤表目录"} onChange={setTableBs} />
        <div className="form-row">
          <label>核对模板</label>
          <select value={selectedTemplate.name} onChange={(event) => setTemplateName(event.target.value)}>
            {templates.map((template) => <option key={template.name} value={template.name}>{template.name}</option>)}
          </select>
          <LayoutTemplate size={17} />
        </div>
        <PathRow label="结果另存为" value={output} kind="save" extensions={["xlsx"]} placeholder={settings.defaultOutputDir || "留空则自动生成结果文件"} onChange={setOutput} />
        <div className="option-row">
          <Toggle checked={autoOpen} onChange={setAutoOpen} label="自动打开结果" />
          <Toggle checked={reportOnly} onChange={setReportOnly} label="仅输出报告" />
          <Toggle checked={overwrite} onChange={setOverwrite} label="覆盖同名结果" />
        </div>
      </div>

      <div className="section-head">
        <h2>模板规则预览</h2>
        <span>{selectedTemplate.rules.length} 条规则</span>
      </div>
      <RulePreviewTable rules={selectedTemplate.rules} />

      {result ? (
        <>
          <div className="section-head">
            <h2>核对明细</h2>
            <span>{result.mismatch_count} 个不一致</span>
          </div>
          <MismatchTable items={result.mismatches} />
        </>
      ) : null}
    </section>
  );
}

function GeneratePage({
  settings,
  busy,
  setBusy,
  addLog,
  addHistory,
  onResult,
}: {
  settings: AppSettings;
  busy: BusyState;
  setBusy: (state: BusyState) => void;
  addLog: (text: string) => void;
  addHistory: (item: HistoryItem) => void;
  onResult: (result: GenerateResult) => void;
}) {
  const [tableC, setTableC] = useLocalState("generate.tableC", "");
  const [templateB, setTemplateB] = useLocalState("generate.templateB", "");
  const [outputDir, setOutputDir] = useLocalState("generate.outputDir", "");
  const [countHolidays, setCountHolidays] = useLocalState("generate.countHolidays", settings.countHolidays);
  const [signatureScale, setSignatureScale] = useLocalState("generate.signatureScale", settings.signatureScale);
  const [normalHours, setNormalHours] = useLocalState("generate.normalHours", settings.normalHours);
  const [morningStart, setMorningStart] = useLocalState("generate.morningStart", settings.morningStart);
  const [morningEnd, setMorningEnd] = useLocalState("generate.morningEnd", settings.morningEnd);
  const [afternoonStart, setAfternoonStart] = useLocalState("generate.afternoonStart", settings.afternoonStart);
  const [afternoonEnd, setAfternoonEnd] = useLocalState("generate.afternoonEnd", settings.afternoonEnd);
  const previewRows = [
    ["1", "Kamlesh Kumar", "Rigger", "250", "9", "14", "10"],
    ["2", "Teshan Saminda", "Banksman", "250", "22", "28", "11"],
    ["3", "Imran Khan", "Grinder", "240", "0", "0", "0"],
    ["4", "Precious Nyuykonghe", "Assistant Administrator", "250", "31", "0", "0"],
  ];

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
      addHistory({
        id: crypto.randomUUID(),
        time: nowText(),
        type: "生成考勤表",
        source: tableC,
        template: fileName(templateB),
        outputPath: data.output_dir,
        reportPath: data.report_path,
        status: "成功",
        generatedCount: data.generated_count,
      });
      addLog(`生成完成: ${data.generated_count} 份`);
    } catch (error) {
      addLog(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy("idle");
    }
  }

  return (
    <section className="workspace">
      <PageHeader
        title="生成考勤表"
        subtitle="根据考勤汇总表批量生成员工月度考勤文件。"
        actions={
          <>
            <button className="secondary-button"><FileSpreadsheet size={16} />读取预览</button>
            <button className="primary-button" disabled={busy !== "idle"} onClick={run}>
              {busy === "generating" ? <Loader2 className="spin" size={18} /> : <Wand2 size={18} />}开始生成
            </button>
          </>
        }
      />

      <div className="panel">
        <PathRow label="考勤汇总表 C" value={tableC} kind="file" extensions={["xlsx", "xlsm"]} onChange={setTableC} />
        <PathRow label="考勤表模板 B" value={templateB} kind="file" extensions={["xlsx", "xlsm"]} onChange={setTemplateB} />
        <PathRow label="输出目录" value={outputDir} kind="folder" placeholder={settings.defaultOutputDir || "选择输出目录"} onChange={setOutputDir} />
      </div>

      <div className="settings-strip">
        <Toggle checked={countHolidays} onChange={setCountHolidays} label="统计假期" />
        <label>签名大小<input type="number" min={30} max={200} value={signatureScale} onChange={(event) => setSignatureScale(Number(event.target.value))} /></label>
        <label>常规小时<input value={normalHours} onChange={(event) => setNormalHours(event.target.value)} /></label>
        <label>上午上班<input value={morningStart} onChange={(event) => setMorningStart(event.target.value)} /></label>
        <label>上午下班<input value={morningEnd} onChange={(event) => setMorningEnd(event.target.value)} /></label>
        <label>下午上班<input value={afternoonStart} onChange={(event) => setAfternoonStart(event.target.value)} /></label>
        <label>下午下班<input value={afternoonEnd} onChange={(event) => setAfternoonEnd(event.target.value)} /></label>
      </div>

      <div className="section-head">
        <h2>待生成员工预览</h2>
        <span>{tableC ? "读取预览后显示真实数据" : "示例数据"}</span>
      </div>
      <div className="data-table">
        <table>
          <thead>
            <tr>
              <th>编号</th><th>员工姓名</th><th>岗位</th><th>工作小时</th><th>常规加班</th><th>周末加班</th><th>法定假加班</th><th>文件名</th>
            </tr>
          </thead>
          <tbody>
            {previewRows.map((row) => (
              <tr key={row[0]}>
                {row.map((cell) => <td key={cell}>{cell}</td>)}
                <td>{row[0]}.{row[1]}-{row[2]}.xlsm</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
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

  useEffect(() => {
    if (templates.length && !templates.some((item) => item.name === current.name)) {
      setCurrent(templates[0]);
    }
  }, [templates, current.name]);

  const enabledRules = current.rules;

  function updateRule(index: number, patch: Partial<CheckRule>) {
    setCurrent({ ...current, rules: current.rules.map((rule, itemIndex) => (itemIndex === index ? { ...rule, ...patch } : rule)) });
  }

  function addRule() {
    setCurrent({ ...current, rules: [...current.rules, { field_name: "新字段", main_range: "A3-An", table_b_cell: "A3", compare_type: "text" }] });
  }

  function duplicateTemplate() {
    setCurrent({ ...current, name: `${current.name} - 副本` });
  }

  function newTemplate() {
    setCurrent({
      name: "新模板",
      number_column: "A",
      start_row: 3,
      rules: [
        { field_name: "姓名", main_range: "E3-En", table_b_cell: "A3", compare_type: "text" },
        { field_name: "岗位", main_range: "F3-Fn", table_b_cell: "C3", compare_type: "position" },
        { field_name: "常规工作小时", main_range: "G3-Gn", table_b_cell: "SUM(G10:Gn)", compare_type: "number" },
      ],
    });
  }

  async function saveTemplate() {
    try {
      const data = await backend<{ template: CheckTemplate }>("save_template", { template: current });
      setTemplates([...templates.filter((item) => item.name !== data.template.name), data.template]);
      addLog(`已保存模板: ${data.template.name}`);
    } catch (error) {
      addLog(error instanceof Error ? error.message : String(error));
    }
  }

  async function deleteTemplate() {
    try {
      await backend("delete_template", { name: current.name });
      const next = templates.filter((item) => item.name !== current.name);
      setTemplates(next.length ? next : [DEFAULT_TEMPLATE]);
      setCurrent(next[0] || DEFAULT_TEMPLATE);
      addLog(`已删除模板: ${current.name}`);
    } catch (error) {
      addLog(error instanceof Error ? error.message : String(error));
    }
  }

  async function importTemplate() {
    const path = await choosePath("file", ["json"]);
    if (!path) return;
    try {
      const data = await backend<{ templates: CheckTemplate[] }>("load_template_file", { path });
      const merged = templates.filter((item) => !data.templates.some((incoming) => incoming.name === item.name));
      setTemplates([...merged, ...data.templates]);
      setCurrent(data.templates[0]);
      addLog(`已导入模板: ${data.templates.map((item) => item.name).join(", ")}`);
    } catch (error) {
      addLog(error instanceof Error ? error.message : String(error));
    }
  }

  async function exportTemplate() {
    const path = await choosePath("save", ["json"]);
    if (!path) return;
    try {
      await backend("export_template_file", { path, template: current });
      addLog(`已导出模板: ${path}`);
    } catch (error) {
      addLog(error instanceof Error ? error.message : String(error));
    }
  }

  async function validateTemplate() {
    try {
      await backend("validate_template", { template: current });
      addLog(`模板有效: ${current.name}`);
    } catch (error) {
      addLog(error instanceof Error ? error.message : String(error));
    }
  }

  return (
    <section className="workspace">
      <PageHeader
        title="核对模板"
        subtitle="手动填写规则，直接保存复用。"
        actions={
          <>
            <button className="secondary-button" onClick={newTemplate}><Plus size={16} />新建</button>
            <button className="secondary-button" onClick={duplicateTemplate}><Copy size={16} />复制</button>
            <button className="secondary-button" onClick={importTemplate}><Import size={16} />导入</button>
            <button className="secondary-button" onClick={exportTemplate}><Download size={16} />导出</button>
            <button className="primary-button" onClick={saveTemplate}><Save size={16} />保存</button>
          </>
        }
      />

      <div className="template-meta">
        <label>当前模板<select value={current.name} onChange={(event) => setCurrent(templates.find((item) => item.name === event.target.value) || current)}>{templates.map((template) => <option key={template.name} value={template.name}>{template.name}</option>)}</select></label>
        <label>模板名称<input value={current.name} onChange={(event) => setCurrent({ ...current, name: event.target.value })} /></label>
        <label>编号列<input value={current.number_column} onChange={(event) => setCurrent({ ...current, number_column: event.target.value.toUpperCase() })} /></label>
        <label>数据起始行<input type="number" value={current.start_row} onChange={(event) => setCurrent({ ...current, start_row: Number(event.target.value) })} /></label>
      </div>

      <div className="section-head">
        <h2>规则列表</h2>
        <div className="inline-actions">
          <button className="secondary-button" onClick={validateTemplate}><CheckCircle2 size={16} />校验模板</button>
          <button className="secondary-button" onClick={addRule}><Plus size={16} />添加规则</button>
        </div>
      </div>

      <div className="editable-table">
        <div className="editable-head">
          <span>启用</span><span>字段名</span><span>主表范围</span><span>考勤表坐标/表达式</span><span>比较类型</span><span>备注</span><span>操作</span>
        </div>
        {enabledRules.map((rule, index) => (
          <div className="editable-row" key={`${rule.field_name}-${index}`}>
            <input type="checkbox" checked readOnly />
            <input value={rule.field_name} onChange={(event) => updateRule(index, { field_name: event.target.value })} />
            <input value={rule.main_range} onChange={(event) => updateRule(index, { main_range: event.target.value.toUpperCase() })} />
            <input value={rule.table_b_cell} onChange={(event) => updateRule(index, { table_b_cell: event.target.value.toUpperCase() })} />
            <select value={rule.compare_type} onChange={(event) => updateRule(index, { compare_type: event.target.value as CompareType })}>
              {compareOptions.map((option) => <option key={option} value={option}>{compareLabels[option]}</option>)}
            </select>
            <input placeholder="可选" />
            <div className="row-actions">
              <button onClick={() => setCurrent({ ...current, rules: [...current.rules, { ...rule, field_name: `${rule.field_name} 副本` }] })}><Copy size={14} /></button>
              <button onClick={() => setCurrent({ ...current, rules: current.rules.filter((_, itemIndex) => itemIndex !== index) })}><Trash2 size={14} /></button>
            </div>
          </div>
        ))}
      </div>
    </section>
  );
}

function HistoryPage({
  history,
  selectedId,
  setSelectedId,
  clearHistory,
}: {
  history: HistoryItem[];
  selectedId: string;
  setSelectedId: (id: string) => void;
  clearHistory: () => void;
}) {
  const selected = history.find((item) => item.id === selectedId) || history[0];
  return (
    <section className="workspace">
      <PageHeader title="历史记录" subtitle="查看最近核对与生成任务，快速打开结果文件。" actions={<button className="secondary-button" onClick={clearHistory}><Trash2 size={16} />清理记录</button>} />
      <div className="filters-row">
        <select><option>全部任务</option><option>工资表核对</option><option>生成考勤表</option></select>
        <select><option>本周</option><option>本月</option><option>全部</option></select>
        <div className="search-input"><Search size={15} /><input placeholder="搜索文件名/模板" /></div>
      </div>
      <div className="history-layout">
        <div className="data-table">
          <table>
            <thead>
              <tr><th>时间</th><th>类型</th><th>源文件</th><th>模板</th><th>状态</th><th>操作</th></tr>
            </thead>
            <tbody>
              {history.map((item) => (
                <tr className={selected?.id === item.id ? "selected-row" : ""} key={item.id} onClick={() => setSelectedId(item.id)}>
                  <td>{item.time}</td><td>{item.type}</td><td>{fileName(item.source)}</td><td>{item.template}</td><td><span className="status-pill">{item.status}</span></td>
                  <td><button className="table-icon" onClick={() => openPath(item.outputPath)}><ExternalLink size={14} /></button></td>
                </tr>
              ))}
              {!history.length ? <tr><td colSpan={6}>暂无历史记录</td></tr> : null}
            </tbody>
          </table>
        </div>
        <div className="history-detail">
          <h2>{selected ? selected.type : "未选择任务"}</h2>
          <p>{selected ? shortPath(selected.outputPath) : "完成任务后会在这里显示详情。"}</p>
          {selected ? (
            <>
              <div className="metric-list compact">
                <div><span>结果</span><strong>{selected.mismatchCount !== undefined ? `${selected.mismatchCount} 个不一致` : `${selected.generatedCount || 0} 份`}</strong></div>
                <div><span>模板</span><strong>{selected.template}</strong></div>
                <div><span>状态</span><strong>{selected.status}</strong></div>
              </div>
              <div className="action-stack">
                <button onClick={() => openPath(selected.outputPath)}><ExternalLink size={16} />打开结果</button>
                <button onClick={() => openPath(selected.reportPath)}><FileCheck2 size={16} />打开报告</button>
                <button onClick={() => revealPath(selected.outputPath)}><FolderOpen size={16} />定位文件</button>
              </div>
              <MismatchTable items={selected.mismatches || []} />
            </>
          ) : null}
        </div>
      </div>
    </section>
  );
}

function SettingsPage({
  settings,
  setSettings,
  addLog,
}: {
  settings: AppSettings;
  setSettings: (settings: AppSettings) => void;
  addLog: (text: string) => void;
}) {
  function patch(update: Partial<AppSettings>) {
    setSettings({ ...settings, ...update });
  }
  return (
    <section className="workspace">
      <PageHeader
        title="设置"
        subtitle="配置默认路径、生成参数和应用偏好。"
        actions={
          <>
            <button className="secondary-button" onClick={() => setSettings(DEFAULT_SETTINGS)}><RotateCcw size={16} />重置默认</button>
            <button className="primary-button" onClick={() => addLog("设置已保存")}><Save size={16} />保存设置</button>
          </>
        }
      />
      <div className="settings-layout">
        <div className="settings-main">
          <div className="settings-section">
            <h2>默认路径</h2>
            <PathRow label="默认工资表目录" value={settings.defaultPayrollDir} kind="folder" onChange={(value) => patch({ defaultPayrollDir: value })} />
            <PathRow label="默认考勤表目录" value={settings.defaultAttendanceDir} kind="folder" onChange={(value) => patch({ defaultAttendanceDir: value })} />
            <PathRow label="默认输出目录" value={settings.defaultOutputDir} kind="folder" onChange={(value) => patch({ defaultOutputDir: value })} />
          </div>
          <div className="settings-section">
            <h2>生成默认参数</h2>
            <div className="settings-strip tight">
              <Toggle checked={settings.countHolidays} onChange={(value) => patch({ countHolidays: value })} label="统计假期" />
              <label>签名大小<input type="number" value={settings.signatureScale} onChange={(event) => patch({ signatureScale: Number(event.target.value) })} /></label>
              <label>常规小时<input value={settings.normalHours} onChange={(event) => patch({ normalHours: event.target.value })} /></label>
              <label>上午上班<input value={settings.morningStart} onChange={(event) => patch({ morningStart: event.target.value })} /></label>
              <label>上午下班<input value={settings.morningEnd} onChange={(event) => patch({ morningEnd: event.target.value })} /></label>
              <label>下午上班<input value={settings.afternoonStart} onChange={(event) => patch({ afternoonStart: event.target.value })} /></label>
              <label>下午下班<input value={settings.afternoonEnd} onChange={(event) => patch({ afternoonEnd: event.target.value })} /></label>
            </div>
          </div>
          <div className="settings-section">
            <h2>模板与别名</h2>
            <div className="file-links">
              <button><FileCog size={16} />check_templates.json</button>
              <button><FileCog size={16} />position_aliases.json</button>
              <button><Upload size={16} />导入</button>
            </div>
          </div>
          <div className="settings-section">
            <h2>应用</h2>
            <div className="settings-strip tight">
              <Toggle checked={settings.autoOpenResult} onChange={(value) => patch({ autoOpenResult: value })} label="自动打开结果" />
              <label>日志保留<input value={settings.keepLogDays} onChange={(event) => patch({ keepLogDays: event.target.value })} /></label>
              <label>版本<input readOnly value="0.2.0" /></label>
            </div>
          </div>
        </div>
        <div className="settings-status">
          <h2>环境状态</h2>
          {["Python sidecar 已打包", "签名字体 已找到", "Tauri 运行正常", "本机数据 不上传"].map((item) => (
            <div className="status-line" key={item}><CheckCircle2 size={16} />{item}</div>
          ))}
        </div>
      </div>
    </section>
  );
}

export default function App() {
  const [page, setPage] = useState<Page>("check");
  const [busy, setBusy] = useState<BusyState>("idle");
  const [templates, setTemplates] = useState<CheckTemplate[]>([DEFAULT_TEMPLATE]);
  const [log, setLog] = useState<string[]>([]);
  const [history, setHistory] = useLocalState<HistoryItem[]>("task.history", []);
  const [settings, setSettings] = useLocalState<AppSettings>("app.settings", DEFAULT_SETTINGS);
  const [selectedHistoryId, setSelectedHistoryId] = useState("");
  const [checkResult, setCheckResult] = useState<CheckResult | null>(null);
  const [generateResult, setGenerateResult] = useState<GenerateResult | null>(null);

  useEffect(() => {
    backend<{ templates: CheckTemplate[] }>("list_templates", {})
      .then((data) => setTemplates(data.templates.length ? data.templates : [DEFAULT_TEMPLATE]))
      .catch((error) => addLog(error instanceof Error ? error.message : String(error)));
  }, []);

  const selectedTemplate = templates[0] || DEFAULT_TEMPLATE;
  const selectedHistory = useMemo(() => history.find((item) => item.id === selectedHistoryId) || history[0] || null, [history, selectedHistoryId]);

  function addLog(text: string) {
    const time = new Date().toLocaleTimeString();
    setLog((items) => [...items, `${time}  ${text}`]);
  }

  function addHistory(item: HistoryItem) {
    setHistory([item, ...history].slice(0, 80));
    setSelectedHistoryId(item.id);
  }

  const nav = [
    { id: "check" as Page, label: "工资表核对", icon: FileCheck2 },
    { id: "generate" as Page, label: "生成考勤表", icon: Archive },
    { id: "templates" as Page, label: "核对模板", icon: LayoutTemplate },
    { id: "history" as Page, label: "历史记录", icon: History },
    { id: "settings" as Page, label: "设置", icon: Settings },
  ];

  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-mark">E</div>
          <span>表格核对工具</span>
        </div>
        <div className="search-box"><Search size={15} /><span>搜索功能、模板、记录</span></div>
        <nav>
          {nav.map((item) => {
            const Icon = item.icon;
            return <button className={page === item.id ? "active" : ""} key={item.id} onClick={() => setPage(item.id)}><Icon size={17} />{item.label}</button>;
          })}
        </nav>
        <div className="tip-card">
          <AlertCircle size={17} />
          <strong>本机运行</strong>
          <span>文件只在你的电脑上处理，不上传。</span>
        </div>
      </aside>

      <div className="content-area">
        {page === "check" && <CheckPage templates={templates} settings={settings} busy={busy} setBusy={setBusy} addLog={addLog} addHistory={addHistory} onResult={(result) => { setCheckResult(result); setGenerateResult(null); }} />}
        {page === "generate" && <GeneratePage settings={settings} busy={busy} setBusy={setBusy} addLog={addLog} addHistory={addHistory} onResult={(result) => { setGenerateResult(result); setCheckResult(null); }} />}
        {page === "templates" && <TemplatesPage templates={templates} setTemplates={setTemplates} addLog={addLog} />}
        {page === "history" && <HistoryPage history={history} selectedId={selectedHistoryId} setSelectedId={setSelectedHistoryId} clearHistory={() => { setHistory([]); setSelectedHistoryId(""); }} />}
        {page === "settings" && <SettingsPage settings={settings} setSettings={setSettings} addLog={addLog} />}
      </div>

      <DetailPanel page={page} busy={busy} selectedTemplate={selectedTemplate} checkResult={checkResult} generateResult={generateResult} selectedHistory={selectedHistory} log={log} />
    </main>
  );
}
