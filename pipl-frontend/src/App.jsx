import React, { useMemo, useRef, useState } from "react";
import {
  AlertCircle,
  Check,
  ChevronDown,
  ClipboardCheck,
  Database,
  Download,
  Edit3,
  FileSpreadsheet,
  FileText,
  Plus,
  RefreshCw,
  Search,
  Settings2,
  Sparkles,
  Trash2,
  UploadCloud,
  X,
} from "lucide-react";

const API_BASE = import.meta.env.VITE_API_BASE || "http://127.0.0.1:8787";
const IS_HOSTED_DEMO = !["localhost", "127.0.0.1"].includes(window.location.hostname);

const ORDER_FIELDS = [
  { key: "po_no", label: "PO NO.", hint: "客户每次订单编号不同，生成前必须确认" },
  { key: "buyer", label: "Buyer 买方", hint: "写入 P.I / P.L 的客户主体" },
  { key: "buyer_address", label: "Buyer’s address 买方地址", hint: "建议人工核对，地址最容易出错" },
  { key: "payment_terms", label: "Terms of payment 付款方式", hint: "例如 30% deposit and 70% before shipment" },
  { key: "delivery_time", label: "Delivery Time 交货期", hint: "客户付款后或指定日期" },
  { key: "port_destination", label: "Port of Destination 目的港", hint: "海关和客户收货会用到" },
];

function currency(value = 0) {
  return `$${Number(value || 0).toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}`;
}

function number(value = 0, digits = 2) {
  return Number(value || 0).toLocaleString("en-US", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });
}

function Button({ children, variant = "primary", className = "", ...props }) {
  return (
    <button className={`btn btn-${variant} ${className}`} {...props}>
      {children}
    </button>
  );
}

function Badge({ children, tone = "neutral" }) {
  return <span className={`badge badge-${tone}`}>{children}</span>;
}

function Panel({ title, action, onAction, children, collapsed = false, onToggle }) {
  return (
    <section className={`panel ${collapsed ? "panel-collapsed" : ""}`}>
      <header>
        <h3>{title}</h3>
        <div className="panel-actions">
          {action ? <button type="button" onClick={onAction}>{action}</button> : null}
          {onToggle ? (
            <button type="button" onClick={onToggle}>
              <ChevronDown />
            </button>
          ) : null}
        </div>
      </header>
      {!collapsed ? children : null}
    </section>
  );
}

function emptyOrder() {
  return {
    po_no: "",
    buyer: "",
    buyer_address: "",
    payment_terms: "30% deposit and 70% before shipment",
    delivery_time: "",
    port_destination: "",
    port_loading: "GUANGZHOU,CHINA",
    fabric_code: "",
    pi_no: "",
    order_date: new Date().toISOString().slice(0, 10),
  };
}

function App() {
  const [step, setStep] = useState(1);
  const [analysis, setAnalysis] = useState(null);
  const [selectedFiles, setSelectedFiles] = useState([]);
  const [loading, setLoading] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [error, setError] = useState("");
  const [generated, setGenerated] = useState(null);
  const [draftRows, setDraftRows] = useState([]);
  const [draftOrder, setDraftOrder] = useState(emptyOrder);
  const [selectedFabric, setSelectedFabric] = useState(null);
  const [fabricQuery, setFabricQuery] = useState("");
  const [fabricResults, setFabricResults] = useState([]);
  const [fabricLoading, setFabricLoading] = useState(false);
  const [detailsOpen, setDetailsOpen] = useState(false);
  const [importCollapsed, setImportCollapsed] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const fileInputRef = useRef(null);

  const fileStatus = analysis?.files || [];
  const fabric = selectedFabric?.summary || analysis?.fabric || null;
  const fabricRecord = selectedFabric?.record || analysis?.fabric_record || null;
  const effectiveInvoiceInput = useMemo(
    () => buildInvoiceInput(draftRows, fabricRecord, analysis?.invoice_input || []),
    [draftRows, fabricRecord, analysis?.invoice_input],
  );
  const invoiceRows = useMemo(
    () => buildPreviewInvoiceRows(effectiveInvoiceInput, fabricRecord),
    [effectiveInvoiceInput, fabricRecord],
  );
  const totals = useMemo(() => buildTotals(invoiceRows, analysis?.totals || {}), [invoiceRows, analysis?.totals]);
  const issues = useMemo(() => buildReviewIssues(draftOrder, fabric, analysis, effectiveInvoiceInput), [draftOrder, fabric, analysis, effectiveInvoiceInput]);
  const remainingIssues = issues.filter((item) => !item.passed).length;
  const hasRows = draftRows.length > 0 || effectiveInvoiceInput.length > 0;
  const hasSession = Boolean(analysis?.session_id);
  const hasPiTemplate = fileStatus.some((file) => file.kind === "pi_template" && file.status === "已识别");
  const hasPlTemplate = fileStatus.some((file) => file.kind === "pl_template" && file.status === "已识别");
  const hasPacking = (analysis?.packing_summary || []).length > 0;
  const piBlockingIssues = issues.filter((item) => !item.passed && !["templates"].includes(item.key)).length;
  const canGeneratePi = hasSession && hasPiTemplate && fabricRecord && effectiveInvoiceInput.length > 0 && piBlockingIssues === 0 && !loading && !generating;
  const canGenerate = hasSession && hasPiTemplate && hasPlTemplate && hasPacking && fabricRecord && effectiveInvoiceInput.length > 0 && remainingIssues === 0 && !loading && !generating;
  const completion = Math.round(((issues.length - remainingIssues) / Math.max(issues.length, 1)) * 100);

  async function analyze(files = selectedFiles, options = {}) {
    setLoading(true);
    setError("");
    setGenerated(null);
    try {
      const form = new FormData();
      files.forEach((file) => form.append("files", file));
      const response = await fetch(`${API_BASE}/api/analyze`, { method: "POST", body: form });
      const data = await readJson(response);
      setAnalysis(data);
      setDraftRows((current) => (options.preserveUser && current.length ? current : data.order_image_rows || []));
      setDraftOrder((current) => mergeOrderPreservingUser(current, { ...emptyOrder(), ...(data.order || {}) }, options.preserveUser));
      setSelectedFabric((current) => current || (data.fabric_record ? { summary: data.fabric, record: data.fabric_record, label: data.fabric?.fabric_code } : null));
      setStep((data.order_image_rows || []).length ? 2 : 1);
    } catch (err) {
      setError(err.message || "分析失败，请确认本地 API 已启动。");
    } finally {
      setLoading(false);
    }
  }

  function handleFiles(files) {
    const nextFiles = mergeFiles(selectedFiles, Array.from(files || []));
    setSelectedFiles(nextFiles);
    analyze(nextFiles, { preserveUser: true });
  }

  function clearWorkspace() {
    setStep(1);
    setAnalysis(null);
    setSelectedFiles([]);
    setError("");
    setGenerated(null);
    setDraftRows([]);
    setDraftOrder(emptyOrder());
    setSelectedFabric(null);
    setFabricQuery("");
    setFabricResults([]);
    setDetailsOpen(false);
    if (fileInputRef.current) fileInputRef.current.value = "";
  }

  async function searchFabric(query = fabricQuery) {
    const text = String(query || "").trim();
    setFabricQuery(text);
    if (!text) {
      setFabricResults([]);
      return;
    }
    setFabricLoading(true);
    setError("");
    try {
      const response = await fetch(`${API_BASE}/api/fabrics/search?q=${encodeURIComponent(text)}&limit=12`);
      const data = await readJson(response);
      setFabricResults(data.results || []);
    } catch (err) {
      setError(err.message || "面料数据库检索失败。");
    } finally {
      setFabricLoading(false);
    }
  }

  function selectFabric(item) {
    setSelectedFabric(item);
    setDraftOrder((current) => ({
      ...current,
      fabric_code: item.record?.fabric_code || item.summary?.fabric_code || current.fabric_code,
      art_no: item.record?.fabric_code || item.summary?.fabric_code || current.art_no,
    }));
    setFabricResults([]);
  }

  async function generate() {
    if (!analysis?.session_id) return;
    setGenerating(true);
    setError("");
    try {
      const response = await fetch(`${API_BASE}/api/generate`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          session_id: analysis.session_id,
          order: draftOrder,
          fabric: fabricRecord,
          invoice_input: effectiveInvoiceInput,
          roll_input: analysis.roll_input || [],
        }),
      });
      const data = await readJson(response);
      setGenerated(data.generated);
      setAnalysis((current) => ({ ...current, order: data.order, totals: data.totals }));
      setDraftOrder((current) => ({ ...current, ...(data.order || {}) }));
    } catch (err) {
      setError(err.message || "生成失败，请检查模板、面料和明细。");
    } finally {
      setGenerating(false);
    }
  }

  async function generatePi() {
    if (!analysis?.session_id) return;
    setGenerating(true);
    setError("");
    try {
      const response = await fetch(`${API_BASE}/api/generate/pi`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          session_id: analysis.session_id,
          order: draftOrder,
          fabric: fabricRecord,
          invoice_input: effectiveInvoiceInput,
        }),
      });
      const data = await readJson(response);
      setGenerated((current) => ({ ...(current || {}), ...(data.generated || {}) }));
      setAnalysis((current) => ({ ...current, order: data.order, totals: data.totals }));
      setDraftOrder((current) => ({ ...current, ...(data.order || {}) }));
    } catch (err) {
      setError(err.message || "生成 P.I 失败，请检查模板、面料和明细。");
    } finally {
      setGenerating(false);
    }
  }

  function updateRow(index, key, value) {
    setDraftRows((rows) =>
      rows.map((row, rowIndex) => {
        if (rowIndex !== index) return row;
        const next = { ...row, [key]: value };
        const colorName = String(next.color_name || "").trim();
        const colorCode = String(next.company_color_code || "").trim();
        next.display_color = colorName && colorCode ? `${colorName} ${formatCompanyColorCode(colorCode)}` : colorName || formatCompanyColorCode(colorCode);
        return next;
      }),
    );
  }

  function addRow() {
    setDraftRows((rows) => [
      ...rows,
      {
        source_file: "手动新增",
        style: draftOrder.fabric_code || "",
        color_name: "",
        company_color_code: "",
        display_color: "",
        quantity: "",
        unit: "Meter",
        unit_price_usd: "",
        amount_usd: "",
      },
    ]);
  }

  function removeRow(index) {
    setDraftRows((rows) => rows.filter((_, rowIndex) => rowIndex !== index));
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand">
          <div className="brand-logo"><Sparkles /></div>
          <div>
            <h1>P.I/P.L一点腾</h1>
            <p>外贸单据智能生成工具</p>
          </div>
        </div>
        <div className="topbar-meta">
          <Badge tone={remainingIssues ? "warning" : hasRows ? "success" : "neutral"}>
            {loading ? "正在识别" : hasRows ? `${remainingIssues} 项需确认` : "待上传截图"}
          </Badge>
          <span>北京时间 {new Date().toLocaleString("zh-CN", { hour12: false })}</span>
          <Button variant="ghost" onClick={() => setSettingsOpen(true)}><Settings2 data-icon="inline-start" />设置</Button>
          <Button variant="outline" onClick={clearWorkspace}><X data-icon="inline-start" />清空数据</Button>
        </div>
      </header>

      <section className="step-tabs">
        {[["1", "上传截图识别"], ["2", "填写缺少信息"], ["3", "审查并生成"]].map(([id, label]) => (
          <button key={id} type="button" className={step === Number(id) ? "active" : ""} onClick={() => setStep(Number(id))}>
            <span>{id}</span>
            {label}
          </button>
        ))}
      </section>

      <section className="command-strip command-strip-compact">
        <div>
          <span>当前 PO NO.</span>
          <strong>{draftOrder.po_no || "未填写"}</strong>
        </div>
        <div>
          <span>面料编号</span>
          <strong>{fabric?.fabric_code || draftOrder.fabric_code || "未选择"}</strong>
        </div>
        <div>
          <span>完成度</span>
          <strong>{completion}%</strong>
        </div>
        <div>
          <span>总金额</span>
          <strong>{currency(totals.total_amount)}</strong>
        </div>
        <Button onClick={() => setStep(2)} variant="outline"><Edit3 data-icon="inline-start" />修改</Button>
      </section>

      {IS_HOSTED_DEMO ? (
        <div className="app-alert demo-alert">
          在线入口已打开。真实 OCR、读取本地面料库、生成 Excel 需要同时启动本机后端；否则这里主要用于随时查看界面和演示流程。
        </div>
      ) : null}

      {error ? <div className="app-alert">{error}</div> : null}

      <section className="workflow-layout">
        <aside className="left-rail">
          <Panel title="资料入口" action="选择文件" onAction={() => fileInputRef.current?.click()} collapsed={importCollapsed} onToggle={() => setImportCollapsed((value) => !value)}>
            <input
              ref={fileInputRef}
              className="file-input"
              type="file"
              multiple
              accept=".xlsx,.xlsm,.csv,.png,.jpg,.jpeg,.webp,.bmp"
              onChange={(event) => handleFiles(event.target.files)}
            />
            <button className="drop-zone" type="button" onClick={() => fileInputRef.current?.click()}>
              <UploadCloud />
              <strong>上传客户截图 / 模板</strong>
              <span>截图识别颜色、公司色号、数量和价格；P.I / P.L 模板用于最终生成 Excel。</span>
            </button>
            <FileStatus files={fileStatus} />
          </Panel>

          <Panel title="输出文件" action="清空数据" onAction={clearWorkspace}>
            <OutputRow title="Proforma Invoice" file={generated?.pi} />
            <OutputRow title="Packing List" file={generated?.pl} />
          </Panel>
        </aside>

        <section className="flow-stage">
          {step === 1 ? (
            <StepUpload
              loading={loading}
              draftRows={draftRows}
              onPick={() => fileInputRef.current?.click()}
              onAnalyze={() => analyze(selectedFiles, { preserveUser: false })}
              onUpdate={updateRow}
              onAdd={addRow}
              onRemove={removeRow}
              onNext={() => setStep(2)}
              disabled={!selectedFiles.length}
            />
          ) : null}

          {step === 2 ? (
            <StepFill
              order={draftOrder}
              setOrder={setDraftOrder}
              fabric={fabric}
              fabricQuery={fabricQuery}
              setFabricQuery={setFabricQuery}
              searchFabric={searchFabric}
              fabricResults={fabricResults}
              fabricLoading={fabricLoading}
              selectFabric={selectFabric}
              invoiceRows={invoiceRows}
              onNext={() => setStep(3)}
            />
          ) : null}

          {step === 3 ? (
            <StepReview
              issues={issues}
              completion={completion}
              order={draftOrder}
              setOrder={setDraftOrder}
              invoiceRows={invoiceRows}
              packingSummary={analysis?.packing_summary || []}
              detailsOpen={detailsOpen}
              setDetailsOpen={setDetailsOpen}
              canGeneratePi={canGeneratePi}
              canGenerate={canGenerate}
              generatePi={generatePi}
              generate={generate}
              generating={generating}
              generated={generated}
              missingReason={missingGenerateReason({ hasPiTemplate, hasPlTemplate, hasPacking, fabricRecord, effectiveInvoiceInput, remainingIssues, piBlockingIssues })}
            />
          ) : null}
        </section>

        <aside className="right-preview">
          <Panel title="实时单据预览" action="预览 Excel">
            <DocumentPreview order={draftOrder} invoiceRows={invoiceRows} imageRows={draftRows} totals={totals} />
          </Panel>
        </aside>
      </section>
      {settingsOpen ? <SettingsDialog onClose={() => setSettingsOpen(false)} /> : null}
    </main>
  );
}

function StepUpload({ loading, draftRows, onPick, onAnalyze, onUpdate, onAdd, onRemove, onNext, disabled }) {
  return (
    <section className="step-card">
      <div className="step-head">
        <div>
          <p>Step 1</p>
          <h2>上传客户截图，识别 COLOR CODE 和数量</h2>
          <span>这里识别的是客户截图里的颜色表，不再识别已经做好的 P.I。</span>
        </div>
        <div className="stage-actions">
          <Button variant="outline" onClick={onPick}><UploadCloud data-icon="inline-start" />上传截图</Button>
          <Button variant="outline" onClick={onAnalyze} disabled={disabled || loading}><RefreshCw data-icon="inline-start" />重新识别</Button>
        </div>
      </div>

      {draftRows.length ? (
        <>
          <EditableRows rows={draftRows} onUpdate={onUpdate} onAdd={onAdd} onRemove={onRemove} />
          <div className="step-actions">
            <Button onClick={onNext}>继续填写缺少信息</Button>
          </div>
        </>
      ) : (
        <div className="empty-state compact-empty">
          <UploadCloud />
          <strong>{loading ? "正在识别截图" : "先上传客户颜色/价格截图"}</strong>
          <span>支持类似 STYLE、COLOR NAME、COLOR CODE、FABRIC QTY、PRICE、AMOUNT 的表格截图。</span>
        </div>
      )}
    </section>
  );
}

function StepFill({ order, setOrder, fabric, fabricQuery, setFabricQuery, searchFabric, fabricResults, fabricLoading, selectFabric, invoiceRows, onNext }) {
  return (
    <section className="step-card">
      <div className="step-head">
        <div>
          <p>Step 2</p>
          <h2>填写缺少信息，选择面料编号</h2>
          <span>PO、买方地址、付款方式、目的港这些内容都可以直接修改。</span>
        </div>
        <Button onClick={onNext}>进入审查页面</Button>
      </div>

      <div className="two-column">
        <section className="form-panel">
          <h3>订单信息</h3>
          <div className="form-grid">
            {ORDER_FIELDS.map((field) => (
              <label className={field.key === "buyer_address" ? "field wide" : "field"} key={field.key}>
                <span>{field.label}</span>
                {field.key === "buyer_address" ? (
                  <textarea value={order[field.key] || ""} rows={4} onChange={(event) => setOrder((current) => ({ ...current, [field.key]: event.target.value }))} />
                ) : (
                  <input value={order[field.key] || ""} onChange={(event) => setOrder((current) => ({ ...current, [field.key]: event.target.value }))} />
                )}
                <em>{field.hint}</em>
              </label>
            ))}
          </div>
        </section>

        <section className="form-panel">
          <h3>面料数据库检索</h3>
          <div className="fabric-search">
            <div className="search-box">
              <Search />
              <input
                value={fabricQuery}
                placeholder="输入面料编号，例如 6529"
                onChange={(event) => setFabricQuery(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === "Enter") searchFabric();
                }}
              />
              <Button variant="outline" onClick={() => searchFabric()} disabled={fabricLoading}>{fabricLoading ? "查询中" : "检索"}</Button>
            </div>
            {fabric ? <FabricSummary fabric={fabric} /> : <div className="soft-note">每次检索都会读取本地面料数据库，选中后会用于 KG / Meter / Yard 换算。</div>}
            {fabricResults.length ? (
              <div className="fabric-results">
                {fabricResults.map((item, index) => (
                  <button type="button" key={`${item.record?.fabric_code}-${index}`} onClick={() => selectFabric(item)}>
                    <strong>{item.summary?.fabric_code}{item.pricing_note ? ` · ${item.pricing_note}` : ""}</strong>
                    <span>{item.summary?.composition || "未填写成分"} · {item.summary?.weight || "克重缺失"} · {item.summary?.width || "幅宽缺失"}</span>
                  </button>
                ))}
              </div>
            ) : null}
          </div>

          <h3>客户截图明细预览</h3>
          <MiniTable rows={invoiceRows} />
        </section>
      </div>
    </section>
  );
}

function StepReview({ issues, completion, order, setOrder, invoiceRows, packingSummary, detailsOpen, setDetailsOpen, canGeneratePi, canGenerate, generatePi, generate, generating, generated, missingReason }) {
  return (
    <section className="step-card">
      <div className="step-head">
        <div>
          <p>Step 3</p>
          <h2>审查关键项，然后生成 P.I / P.L</h2>
          <span>系统已经完成 {completion}%，剩下只看会影响付款、交货和报关的字段。</span>
        </div>
        <div className="stage-actions">
          <Button onClick={generatePi} disabled={!canGeneratePi}><ClipboardCheck data-icon="inline-start" />{generating ? "生成中..." : "生成 P.I"}</Button>
          <Button variant="outline" onClick={generate} disabled={!canGenerate}>生成 P.I + P.L</Button>
        </div>
      </div>

      {missingReason ? <div className="app-alert subtle-alert">{missingReason}</div> : null}

      <section className="progress-card slim">
        <div>
          <span>系统已自动完成</span>
          <strong>{completion}%</strong>
        </div>
        <div className="progress-track"><span style={{ width: `${completion}%` }} /></div>
      </section>

      <div className="review-grid">
        {issues.map((issue) => (
          <label className={`review-item ${issue.passed ? "passed" : "attention"}`} key={issue.key}>
            <div>
              <span>{issue.passed ? <Check /> : <AlertCircle />}</span>
              <strong>{issue.label}</strong>
              <Badge tone={issue.passed ? "success" : "warning"}>{issue.passed ? "已确认" : "需要确认"}</Badge>
            </div>
            {issue.editable ? (
              issue.multiline ? (
                <textarea value={order[issue.key] || ""} rows={3} onChange={(event) => setOrder((current) => ({ ...current, [issue.key]: event.target.value }))} />
              ) : (
                <input value={order[issue.key] || ""} onChange={(event) => setOrder((current) => ({ ...current, [issue.key]: event.target.value }))} />
              )
            ) : (
              <p>{issue.value}</p>
            )}
          </label>
        ))}
      </div>

      <div className="stage-actions review-actions">
        <Button variant="outline" onClick={() => setDetailsOpen((value) => !value)}><ChevronDown data-icon="inline-start" />{detailsOpen ? "收起明细" : "展开明细"}</Button>
        {generated?.pi ? <a className="btn btn-outline" href={`${API_BASE}${generated.pi.url}`}><Download data-icon="inline-start" />下载 P.I</a> : null}
        {generated?.pl ? <a className="btn btn-outline" href={`${API_BASE}${generated.pl.url}`}><Download data-icon="inline-start" />下载 P.L</a> : null}
      </div>

      {detailsOpen ? <Details invoiceRows={invoiceRows} packingSummary={packingSummary} /> : null}
    </section>
  );
}

function EditableRows({ rows, onUpdate, onAdd, onRemove }) {
  return (
    <div className="editable-table">
      <table>
        <thead>
          <tr>
            <th>STYLE</th>
            <th>COLOR NAME</th>
            <th>公司色号</th>
            <th>数量</th>
            <th>单位</th>
            <th>单价 USD</th>
            <th>总金额</th>
            <th />
          </tr>
        </thead>
        <tbody>
          {rows.map((row, index) => (
            <tr key={`${row.source_file}-${index}`}>
              <td><input value={row.style || ""} onChange={(event) => onUpdate(index, "style", event.target.value)} /></td>
              <td><input value={row.color_name || ""} onChange={(event) => onUpdate(index, "color_name", event.target.value)} /></td>
              <td><input value={row.company_color_code || ""} onChange={(event) => onUpdate(index, "company_color_code", event.target.value)} /></td>
              <td><input type="number" value={row.quantity ?? ""} onChange={(event) => onUpdate(index, "quantity", event.target.value)} /></td>
              <td>
                <select value={row.unit || "Meter"} onChange={(event) => onUpdate(index, "unit", event.target.value)}>
                  <option>Meter</option>
                  <option>Yard</option>
                  <option>KG</option>
                </select>
              </td>
              <td><input type="number" step="0.0001" value={row.unit_price_usd ?? ""} onChange={(event) => onUpdate(index, "unit_price_usd", event.target.value)} /></td>
              <td><input type="number" step="0.01" value={row.amount_usd ?? ""} onChange={(event) => onUpdate(index, "amount_usd", event.target.value)} /></td>
              <td><button type="button" onClick={() => onRemove(index)}><Trash2 /></button></td>
            </tr>
          ))}
        </tbody>
      </table>
      <Button variant="outline" onClick={onAdd}><Plus data-icon="inline-start" />新增一行</Button>
    </div>
  );
}

function FileStatus({ files }) {
  const defaults = files.length ? files : [
    { kind: "pi_template", label: "P.I 模板", name: "未选择", status: "缺少" },
    { kind: "pl_template", label: "P.L 模板", name: "未选择", status: "缺少" },
    { kind: "fabric_db", label: "面料库", name: "默认库自动读取", status: "可检索" },
  ];
  return (
    <div className="file-list">
      {defaults.map((file) => (
        <div className="file-row" key={file.kind}>
          {file.kind === "fabric_db" ? <Database /> : file.kind === "order_image" ? <FileText /> : <FileSpreadsheet />}
          <div>
            <span>{file.label}</span>
            <strong>{file.name}</strong>
          </div>
          <Badge tone={["已识别", "已解析", "可检索"].includes(file.status) ? "success" : "warning"}>{file.status}</Badge>
        </div>
      ))}
    </div>
  );
}

function OutputRow({ title, file }) {
  return (
    <div className="output-row">
      <FileSpreadsheet />
      <div>
        <span>{title}</span>
        <strong>{file?.name || "等待生成"}</strong>
      </div>
      <Badge tone={file ? "success" : "neutral"}>{file ? "可下载" : "未生成"}</Badge>
    </div>
  );
}

function SettingsDialog({ onClose }) {
  return (
    <div className="drawer-backdrop">
      <aside className="edit-drawer settings-drawer">
        <header>
          <div>
            <span>工具设置</span>
            <h2>当前规则说明</h2>
          </div>
          <button type="button" onClick={onClose}><X /></button>
        </header>
        <div className="settings-list">
          <div>
            <strong>模板用途</strong>
            <p>P.I / P.L 模板只提供 Excel 版式和输出位置，不会覆盖你已填写的 PO、买方、颜色、数量和面料数据。</p>
          </div>
          <div>
            <strong>P.I 数据来源</strong>
            <p>P.I 优先使用客户截图识别结果和人工修改内容；模板里已有的业务数据只在空白时作为补充。</p>
          </div>
          <div>
            <strong>P.L 数据来源</strong>
            <p>P.L 后续单独识别码单细码，包括 LOT、ROLL、Net Weight、Meter、Yard、Tube 和 Gross Weight。</p>
          </div>
          <div>
            <strong>面料库</strong>
            <p>面料编号检索每次读取本地 0428 fabric_prices.xlsx，用量化完成 KG / Meter / Yard 换算。</p>
          </div>
        </div>
        <div className="drawer-actions">
          <Button onClick={onClose}>知道了</Button>
        </div>
      </aside>
    </div>
  );
}

function FabricSummary({ fabric }) {
  return (
    <div className="fabric-summary">
      <div><span>品名</span><strong>{fabric.name_cn || fabric.name_en || "未填写"}</strong></div>
      <div><span>成分</span><strong>{fabric.composition || "-"}</strong></div>
      <div><span>克重</span><strong>{fabric.weight || "-"}</strong></div>
      <div><span>幅宽</span><strong>{fabric.width || "-"}</strong></div>
      <div><span>量化</span><strong>{fabric.quantification_m_per_kg || "-"} m/kg</strong></div>
      <div><span>纸筒空差</span><strong>{fabric.tube_plus_allowance_kg_per_roll ?? "-"} KG</strong></div>
    </div>
  );
}

function MiniTable({ rows }) {
  return (
    <div className="mini-table">
      <table>
        <thead>
          <tr><th>Color</th><th>Qty</th><th>Unit</th><th>USD/KG</th><th>Yard</th><th>USD/Yard</th></tr>
        </thead>
        <tbody>
          {rows.length ? rows.map((row, index) => (
            <tr key={`${row.color}-${index}`}>
              <td>{row.color}</td>
              <td>{number(row.quantity_input, 2)}</td>
              <td>{row.input_unit}</td>
              <td>{number(row.usd_price_per_kg || 0, 4)}</td>
              <td>{number(row.total_yard || 0)}</td>
              <td>{number(row.usd_price_per_yard || 0, 4)}</td>
            </tr>
          )) : (
            <tr><td colSpan="6">选择面料编号后会自动换算。</td></tr>
          )}
        </tbody>
      </table>
    </div>
  );
}

function DocumentPreview({ order, invoiceRows, imageRows = [], totals }) {
  const rows = invoiceRows.length ? invoiceRows : imageRows;
  const rawOnly = !invoiceRows.length && imageRows.length > 0;
  return (
    <div className="document-preview">
      <div className="preview-switch">
        <button type="button" className="active">P.I</button>
        <button type="button">P.L</button>
      </div>
      <div className="document-sheet">
        <div className="sheet-title">
          <div>
            <span>PROFORMA INVOICE</span>
            <strong>{order.pi_no || "生成时自动编号"}</strong>
          </div>
          <Badge tone={rows.length ? "success" : "neutral"}>{rows.length ? "预览中" : "待识别"}</Badge>
        </div>
        <div className="sheet-meta">
          <div><span>Buyer</span><strong>{order.buyer || "-"}</strong></div>
          <div><span>PO</span><strong>{order.po_no || "未填写"}</strong></div>
          <div><span>Fabric</span><strong>{order.fabric_code || "-"}</strong></div>
          <div><span>Date</span><strong>{order.order_date || new Date().toISOString().slice(0, 10)}</strong></div>
        </div>
        <table>
          <thead>
          <tr>
            <th>Color</th>
            <th>{rawOnly ? "Qty" : "KG"}</th>
            {!rawOnly ? <th>Yard</th> : null}
              {!rawOnly ? <th>USD/KG</th> : null}
              <th>{rawOnly ? "Price" : "USD/Yard"}</th>
            <th>Amount</th>
          </tr>
          </thead>
          <tbody>
            {rows.length ? rows.map((row, index) => (
              <tr key={`${row.color || row.display_color}-${index}`}>
                <td>{displayColor(row)}</td>
                <td>{rawOnly ? `${number(row.quantity, 0)} ${row.unit || ""}` : number(row.total_net_weight_kg)}</td>
                {!rawOnly ? <td>{number(row.total_yard)}</td> : null}
                {!rawOnly ? <td>{number(row.usd_price_per_kg, 2)}</td> : null}
                <td>{rawOnly ? currency(row.unit_price_usd) : number(row.usd_price_per_yard, 4)}</td>
                <td>{currency(row.amount_usd)}</td>
              </tr>
            )) : (
              <tr><td colSpan={rawOnly ? 4 : 6}>上传截图后显示明细。</td></tr>
            )}
          </tbody>
        </table>
        <div className="sheet-total">
          <div><span>Total</span><strong>{currency(totals.total_amount)}</strong></div>
          <div><span>Deposit 30%</span><strong>{currency(totals.deposit)}</strong></div>
          <div><span>Balance</span><strong>{currency(totals.balance)}</strong></div>
        </div>
        <div className="pl-summary">
          <div><span>P.L Net KG</span><strong>{number(totals.packing_kg)} KG</strong></div>
          <div><span>Rolls</span><strong>{totals.rolls || 0}</strong></div>
          <div><span>Groups</span><strong>{totals.groups || 0}</strong></div>
        </div>
      </div>
    </div>
  );
}

function Details({ invoiceRows, packingSummary }) {
  return (
    <section className="details-table">
      <header><h3>明细表</h3><span>P.I 金额与 P.L 细码汇总</span></header>
      <table>
        <thead><tr><th>颜色</th><th>KG</th><th>Yard</th><th>USD/KG</th><th>USD/Yard</th><th>Amount</th></tr></thead>
        <tbody>
          {invoiceRows.map((row, index) => (
            <tr key={`${row.color}-${index}`}>
              <td>{row.color}</td><td>{number(row.total_net_weight_kg)}</td><td>{number(row.total_yard)}</td><td>{number(row.usd_price_per_kg, 4)}</td><td>{number(row.usd_price_per_yard, 4)}</td><td>{currency(row.amount_usd)}</td>
            </tr>
          ))}
        </tbody>
      </table>
      <table>
        <thead><tr><th>颜色 / LOT</th><th>Rolls</th><th>Net KG</th><th>Gross KG</th><th>Amount</th></tr></thead>
        <tbody>
          {packingSummary.map((row, index) => (
            <tr key={`${row.color}-${row.lot}-${index}`}>
              <td>{row.color} / {row.lot}</td><td>{row.rolls}</td><td>{number(row.total_net_weight_kg)}</td><td>{number(row.total_gross_weight_kg)}</td><td>{currency(row.amount_usd)}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </section>
  );
}

function buildInvoiceInput(rows, fabricRecord, fallback = []) {
  if (!rows.length) return fallback || [];
  const q = Number(fabricRecord?.quantification_m_per_kg || 0);
  return rows
    .map((row) => {
      const quantity = Number(row.quantity || 0);
      const unit = normalizeUnit(row.unit);
      const unitPrice = Number(row.unit_price_usd || 0);
      const amount = Number(row.amount_usd || 0);
      const effectiveUnitPrice = unitPrice || (quantity ? amount / quantity : 0);
      let usdPerKg = Number(row.usd_price_per_kg || 0);
      if (effectiveUnitPrice && q) {
        if (unit === "Meter") usdPerKg = effectiveUnitPrice * q;
        else if (unit === "Yard") usdPerKg = effectiveUnitPrice * q * 0.9144;
        else usdPerKg = effectiveUnitPrice;
      }
      const colorName = String(row.color_name || "").trim();
      const colorCode = String(row.company_color_code || "").trim();
      const formattedCode = formatCompanyColorCode(colorCode);
      const color = colorName && formattedCode ? `${colorName} ${formattedCode}` : colorName || formattedCode || row.color || "";
      return {
        color,
        art_no: colorName && formattedCode ? `${colorName} - ${formattedCode}` : color,
        ppo_reference_yards: null,
        quantity_input: quantity,
        input_unit: unit,
        usd_price_per_kg: usdPerKg,
        source_fabric_code: row.style || fabricRecord?.fabric_code || "",
        source_note: `客户截图: ${row.source_file || ""}`,
      };
    })
    .filter((row) => row.color && row.quantity_input);
}

function buildPreviewInvoiceRows(invoiceInput, fabricRecord) {
  const q = Number(fabricRecord?.quantification_m_per_kg || 0);
  if (!q) return [];
  return invoiceInput.map((row) => {
    const quantity = Number(row.quantity_input || 0);
    const unit = normalizeUnit(row.input_unit);
    const meter = unit === "KG" ? quantity * q : unit === "Meter" ? quantity : quantity * 0.9144;
    const kg = meter / q;
    const price = Number(row.usd_price_per_kg || 0);
    const yard = meter / 0.9144;
    const amount = kg * price;
    return {
      ...row,
      total_meter: meter,
      total_net_weight_kg: kg,
      total_yard: yard,
      usd_price_per_yard: yard ? amount / yard : 0,
      amount_usd: amount,
    };
  });
}

function buildTotals(invoiceRows, backendTotals) {
  const total = invoiceRows.reduce((sum, row) => sum + Number(row.amount_usd || 0), 0);
  return {
    ...backendTotals,
    total_amount: total || backendTotals.total_amount || 0,
    deposit: (total || backendTotals.total_amount || 0) * 0.3,
    balance: (total || backendTotals.total_amount || 0) * 0.7,
  };
}

function buildReviewIssues(order, fabric, analysis, invoiceInput) {
  return [
    { key: "po_no", label: "PO NO.", value: order.po_no, passed: Boolean(order.po_no), editable: true },
    { key: "buyer_address", label: "买方地址 Buyer’s address", value: order.buyer_address, passed: Boolean(order.buyer_address), editable: true, multiline: true },
    { key: "payment_terms", label: "付款方式 Terms of payment", value: order.payment_terms, passed: Boolean(order.payment_terms), editable: true },
    { key: "delivery_time", label: "交货期 Delivery Time", value: order.delivery_time, passed: Boolean(order.delivery_time), editable: true },
    { key: "port_destination", label: "目的港 Port of Destination", value: order.port_destination, passed: Boolean(order.port_destination), editable: true },
    { key: "fabric", label: "面料编号 / 数据库", value: fabric?.fabric_code || "未选择", passed: Boolean(fabric), editable: false },
    { key: "invoice", label: "P.I 颜色数量明细", value: `${invoiceInput.length} 行`, passed: invoiceInput.length > 0, editable: false },
    { key: "templates", label: "P.L 模板 / 码单细码", value: "P.L 需要单独识别 LOT、ROLL、Net Weight", passed: Boolean(analysis?.can_generate || ((analysis?.packing_summary || []).length > 0)), editable: false },
  ];
}

function missingGenerateReason({ hasPiTemplate, hasPlTemplate, hasPacking, fabricRecord, effectiveInvoiceInput, remainingIssues, piBlockingIssues }) {
  if (!hasPiTemplate) return "还缺 P.I 模板。上传模板后才能生成 Excel。";
  if (!fabricRecord) return "还没有选择面料编号，请在第 2 步检索并选中面料。";
  if (!effectiveInvoiceInput.length) return "还没有可生成的 P.I 明细，请先上传截图或手动新增颜色行。";
  if (piBlockingIssues) return "P.I 还有关键项未填写。填完后可以先单独生成 P.I。";
  if (!hasPlTemplate) return "P.I 已可生成；P.L 还缺 P.L 模板，后续上传后可生成完整 P.I + P.L。";
  if (!hasPacking) return "P.I 已可生成；P.L 还没有码单细码，后续识别 LOT / ROLL / Net Weight 后再生成 P.L。";
  if (remainingIssues) return "还有关键项未填写或未确认。";
  return "";
}

function normalizeUnit(value) {
  const text = String(value || "Meter").trim();
  if (/^kg$/i.test(text)) return "KG";
  if (/^yard|yds?$/i.test(text)) return "Yard";
  return "Meter";
}

function formatCompanyColorCode(value) {
  const text = String(value || "").trim();
  if (!text) return "";
  const cleaned = text.replace(/^#+/, "").replace(/#+$/, "");
  return cleaned ? `${cleaned}#` : "";
}

function displayColor(row) {
  if (row.color_name || row.company_color_code) {
    const colorName = String(row.color_name || "").trim();
    const code = formatCompanyColorCode(row.company_color_code);
    return colorName && code ? `${colorName} ${code}` : colorName || code;
  }
  if (row.color) return String(row.color).replace(/\s+#([^#\s]+)#/g, " $1#");
  if (row.display_color) return String(row.display_color).replace(/\s+#([^#\s]+)#/g, " $1#");
  return "";
}

function mergeFiles(existing, incoming) {
  const byKey = new Map();
  [...existing, ...incoming].forEach((file) => {
    if (!file) return;
    byKey.set(`${file.name}-${file.size}-${file.lastModified}`, file);
  });
  return Array.from(byKey.values());
}

function mergeOrderPreservingUser(current, incoming, preserveUser) {
  if (!preserveUser) return incoming;
  const merged = { ...incoming, ...current };
  Object.keys(incoming || {}).forEach((key) => {
    if (isBlank(current?.[key]) && !isBlank(incoming?.[key])) {
      merged[key] = incoming[key];
    }
  });
  return merged;
}

function isBlank(value) {
  return value === undefined || value === null || String(value).trim() === "";
}

async function readJson(response) {
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.detail || "请求失败");
  return data;
}

export default App;
