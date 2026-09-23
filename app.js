const DB_NAME = "sales-ops-center";
const DB_VERSION = 1;
const PAGE_SIZE = 50;

const state = {
  records: [],
  mapping: new Map(),
  mappingRows: [],
  filtered: [],
  page: 1,
};

const el = (id) => document.getElementById(id);
const money = new Intl.NumberFormat("zh-CN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const integer = new Intl.NumberFormat("zh-CN");
const percent = new Intl.NumberFormat("zh-CN", { style: "percent", minimumFractionDigits: 2, maximumFractionDigits: 2 });

function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => request.result.createObjectStore("state");
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function dbGet(key) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction("state", "readonly");
    const request = tx.objectStore("state").get(key);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function dbSet(key, value) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction("state", "readwrite");
    tx.objectStore("state").put(value, key);
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
  });
}

function normalizeId(value) {
  return String(value ?? "").trim().replace(/\.0$/, "");
}

function normalizeDate(value) {
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "number" && window.XLSX?.SSF) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
  }
  const text = String(value ?? "").trim().replaceAll("/", "-");
  const match = text.match(/(20\d{2})-(\d{1,2})-(\d{1,2})/);
  return match ? `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}` : text;
}

function numeric(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const parsed = Number(String(value ?? "0").replaceAll(",", "").replace("%", ""));
  return Number.isFinite(parsed) ? parsed : 0;
}

function categoryOf(raw, title = "") {
  const category = String(raw || "").trim();
  if (["空净", "配件", "清洁液", "汽车"].includes(category)) return category;
  if (/清洁液|清洁剂|清洗液|洗地液/.test(title)) return "清洁液";
  if (/车载|汽车|车用/.test(title)) return "汽车";
  if (/空气净化|新风|除湿|空调|加湿|净化器/.test(title)) return "空净";
  return "配件";
}

function rebuildMapping(rows) {
  state.mappingRows = rows;
  state.mapping = new Map(rows.map((row) => {
    const sourceOwner = String(row[1] || "未知");
    const owner = sourceOwner === "运营中台" ? "未知" : sourceOwner;
    return [normalizeId(row[0]), { owner, category: categoryOf(row[3], String(row[2] || "")) }];
  }));
}

function enrich(record) {
  const match = state.mapping.get(record.id);
  const included = record.title.includes("1217") && !record.title.includes("打印机");
  return {
    ...record,
    included,
    owner: match?.owner || "未知",
    category: match?.category || categoryOf("", record.title),
    net: record.payment - record.refund,
    expense: numeric(record.expense),
  };
}

function mergeSpendRows(records, spendRows) {
  const merged = new Map(records.map((row) => [`${row.date}|${row.id}`, { ...row, expense: numeric(row.expense) }]));
  for (const row of spendRows) {
    const spend = { date: normalizeDate(row[0]), id: normalizeId(row[1]), title: String(row[2] || "").trim(), expense: numeric(row[3]) };
    const key = `${spend.date}|${spend.id}`;
    const current = merged.get(key) || { date: spend.date, id: spend.id, title: spend.title, payment: 0, refund: 0, expense: 0 };
    merged.set(key, { ...current, title: current.title || spend.title, expense: spend.expense });
  }
  return [...merged.values()];
}

async function initialize() {
  const saved = await dbGet("payload");
  if (saved?.records?.length && saved?.mappingRows?.length) {
    state.records = saved.records;
    rebuildMapping(saved.mappingRows);
    if (saved.version !== 2) {
      const spend = await fetch("./data/spend.json").then((response) => response.json());
      state.records = mergeSpendRows(state.records, spend.rows);
      await persist();
    }
  } else {
    const [seed, spend] = await Promise.all([
      fetch("./data/seed.json").then((response) => response.json()),
      fetch("./data/spend.json").then((response) => response.json()),
    ]);
    rebuildMapping(seed.mapping_rows);
    const salesRows = seed.import_rows.map((row) => ({ date: row[0], id: normalizeId(row[1]), title: String(row[2] || ""), payment: numeric(row[3]), refund: numeric(row[4]), expense: 0 }));
    state.records = mergeSpendRows(salesRows, spend.rows);
    await persist();
  }
  const dates = state.records.map((row) => row.date).filter(Boolean).sort();
  el("startDate").value = dates[0] || "";
  el("endDate").value = dates.at(-1) || "";
  el("detailStartDate").value = dates[0] || "";
  el("detailEndDate").value = dates.at(-1) || "";
  el("syncText").textContent = `${integer.format(state.records.length)} 条数据已就绪`;
  render();
}

async function persist() {
  await dbSet("payload", { version: 2, records: state.records, mappingRows: state.mappingRows });
}

function activeRecords() {
  const start = el("startDate").value;
  const end = el("endDate").value;
  return state.records.map(enrich).filter((row) => row.included && (!start || row.date >= start) && (!end || row.date <= end));
}

function aggregate(rows) {
  const byOwner = new Map();
  const total = { count: 0, payment: 0, refund: 0, net: 0, expense: 0 };
  for (const row of rows) {
    total.count += 1; total.payment += row.payment; total.refund += row.refund; total.net += row.net; total.expense += row.expense;
    if (!byOwner.has(row.owner)) byOwner.set(row.owner, {
      空净: { net: 0, expense: 0 }, 配件: { net: 0, expense: 0 }, 清洁液: { net: 0, expense: 0 }, 汽车: { net: 0, expense: 0 }, totalNet: 0, totalExpense: 0,
    });
    const bucket = byOwner.get(row.owner);
    bucket[row.category].net += row.net;
    bucket[row.category].expense += row.expense;
    bucket.totalNet += row.net;
    bucket.totalExpense += row.expense;
  }
  return { total, byOwner };
}

function aggregateProducts(rows) {
  const byProduct = new Map();
  for (const row of rows) {
    if (!byProduct.has(row.id)) byProduct.set(row.id, { ...row, payment: 0, refund: 0, net: 0, expense: 0 });
    const product = byProduct.get(row.id);
    product.payment += row.payment;
    product.refund += row.refund;
    product.net += row.net;
    product.expense += row.expense;
  }
  return [...byProduct.values()];
}

function promotionRisk(row) {
  if (row.expense <= 0) return "未推广";
  if (row.net <= 0) return "空烧";
  const ratio = row.expense / row.net;
  if (ratio >= 0.5) return "高风险";
  if (ratio >= 0.2) return "需关注";
  return "正常";
}

function render() {
  const rows = activeRecords();
  const { total, byOwner } = aggregate(rows);
  el("recordCount").textContent = integer.format(total.count);
  el("paymentTotal").textContent = money.format(total.payment);
  el("refundTotal").textContent = money.format(total.refund);
  el("netTotal").textContent = money.format(total.net);
  el("expenseTotal").textContent = money.format(total.expense);

  const ownerRows = [...byOwner.entries()].sort((a, b) => b[1].totalNet - a[1].totalNet);
  el("ownerCategoryBody").innerHTML = ownerRows.map(([owner, value]) => {
    const ratio = value.totalNet ? percent.format(value.totalExpense / value.totalNet) : "—";
    return `<tr><td class="owner-name">${escapeHtml(owner)}</td><td>${money.format(value.空净.net)}</td><td class="expense-cell">${money.format(value.空净.expense)}</td><td>${money.format(value.配件.net)}</td><td class="expense-cell">${money.format(value.配件.expense)}</td><td>${money.format(value.清洁液.net)}</td><td class="expense-cell">${money.format(value.清洁液.expense)}</td><td>${money.format(value.汽车.net)}</td><td class="expense-cell">${money.format(value.汽车.expense)}</td><td><strong>${money.format(value.totalNet)}</strong></td><td class="expense-cell"><strong>${money.format(value.totalExpense)}</strong></td><td class="ratio-cell"><strong>${ratio}</strong></td></tr>`;
  }).join("");
  const cats = { 空净: { net: 0, expense: 0 }, 配件: { net: 0, expense: 0 }, 清洁液: { net: 0, expense: 0 }, 汽车: { net: 0, expense: 0 } };
  ownerRows.forEach(([, value]) => Object.keys(cats).forEach((key) => { cats[key].net += value[key].net; cats[key].expense += value[key].expense; }));
  el("ownerCategoryFoot").innerHTML = `<tr><td>合计</td><td>${money.format(cats.空净.net)}</td><td>${money.format(cats.空净.expense)}</td><td>${money.format(cats.配件.net)}</td><td>${money.format(cats.配件.expense)}</td><td>${money.format(cats.清洁液.net)}</td><td>${money.format(cats.清洁液.expense)}</td><td>${money.format(cats.汽车.net)}</td><td>${money.format(cats.汽车.expense)}</td><td>${money.format(total.net)}</td><td>${money.format(total.expense)}</td><td>${total.net ? percent.format(total.expense / total.net) : "—"}</td></tr>`;

  renderPromotionAnalysis(rows);

  const owners = [...new Set(rows.map((row) => row.owner))].sort();
  const currentOwner = el("ownerFilter").value;
  el("ownerFilter").innerHTML = `<option value="">全部负责人</option>${owners.map((owner) => `<option${owner === currentOwner ? " selected" : ""}>${escapeHtml(owner)}</option>`).join("")}`;
  renderDetails(rows);
  registerWebMcp();
}

function renderDetails(baseRows = activeRecords()) {
  const query = el("searchInput").value.trim().toLowerCase();
  const owner = el("ownerFilter").value;
  const category = el("categoryFilter").value;
  const detailStartDate = el("detailStartDate").value;
  const detailEndDate = el("detailEndDate").value;
  const products = aggregateProducts(baseRows.filter((item) => (!detailStartDate || item.date >= detailStartDate) && (!detailEndDate || item.date <= detailEndDate)));
  state.filtered = products.filter((row) => (!query || row.id.includes(query) || row.title.toLowerCase().includes(query)) && (!owner || row.owner === owner) && (!category || row.category === category)).sort((a, b) => b.net - a.net || b.expense - a.expense);
  const pages = Math.max(1, Math.ceil(state.filtered.length / PAGE_SIZE));
  state.page = Math.min(state.page, pages);
  const start = (state.page - 1) * PAGE_SIZE;
  const pageRows = state.filtered.slice(start, start + PAGE_SIZE);
  const chipClass = (category) => ({ 空净: "air", 配件: "parts", 清洁液: "cleaner", 汽车: "car" }[category] || "parts");
  el("detailBody").innerHTML = pageRows.map((row) => {
    const expenseRatio = row.net ? percent.format(row.expense / row.net) : "—";
    return `<tr><td>${escapeHtml(row.owner)}</td><td><span class="category-chip ${chipClass(row.category)}">${row.category}</span></td><td>${escapeHtml(row.id)}</td><td title="${escapeHtml(row.title)}">${escapeHtml(row.title)}</td><td>${money.format(row.payment)}</td><td>${money.format(row.refund)}</td><td><strong>${money.format(row.net)}</strong></td><td class="expense-cell"><strong>${money.format(row.expense)}</strong></td><td class="ratio-cell"><strong>${expenseRatio}</strong></td></tr>`;
  }).join("");
  el("pageInfo").textContent = `第 ${state.page} / ${pages} 页，共 ${integer.format(state.filtered.length)} 条`;
  el("prevPage").disabled = state.page <= 1;
  el("nextPage").disabled = state.page >= pages;
}

function renderPromotionAnalysis(rows = activeRecords()) {
  const promoted = aggregateProducts(rows).filter((row) => row.expense > 0).map((row) => ({ ...row, risk: promotionRisk(row) }));
  const emptyBurn = promoted.filter((row) => row.risk === "空烧");
  const highRisk = promoted.filter((row) => row.risk === "高风险");
  const watch = promoted.filter((row) => row.risk === "需关注");
  const promotedNet = promoted.reduce((sum, row) => sum + row.net, 0);
  const promotedExpense = promoted.reduce((sum, row) => sum + row.expense, 0);
  const emptyBurnExpense = emptyBurn.reduce((sum, row) => sum + row.expense, 0);
  el("promotedProductCount").textContent = integer.format(promoted.length);
  el("promotedNetTotal").textContent = money.format(promotedNet);
  el("emptyBurnCount").textContent = integer.format(emptyBurn.length);
  el("emptyBurnSpend").textContent = `空烧花费 ${money.format(emptyBurnExpense)}`;
  el("highRiskCount").textContent = integer.format(highRisk.length);
  const ratio = promotedNet > 0 ? percent.format(promotedExpense / promotedNet) : "—";
  el("promotionInsight").textContent = `所选期间共有 ${integer.format(promoted.length)} 个商品发生推广花费，合计 ${money.format(promotedExpense)}，对应净销售额 ${money.format(promotedNet)}，整体花费占比 ${ratio}。其中空烧 ${integer.format(emptyBurn.length)} 个，高风险 ${integer.format(highRisk.length)} 个，需关注 ${integer.format(watch.length)} 个。`;

  const selected = el("promotionRiskFilter").value;
  const priority = { 空烧: 0, 高风险: 1, 需关注: 2, 正常: 3 };
  const filtered = promoted.filter((row) => selected === "all" || (selected === "risk" ? row.risk !== "正常" : row.risk === selected)).sort((a, b) => priority[a.risk] - priority[b.risk] || b.expense - a.expense);
  const badgeClass = { 空烧: "burn", 高风险: "high", 需关注: "watch", 正常: "normal" };
  el("promotionBody").innerHTML = filtered.map((row) => {
    const expenseRatio = row.net > 0 ? percent.format(row.expense / row.net) : "—";
    return `<tr><td><span class="risk-badge ${badgeClass[row.risk]}">${row.risk}</span></td><td>${escapeHtml(row.owner)}</td><td>${escapeHtml(row.category)}</td><td>${escapeHtml(row.id)}</td><td title="${escapeHtml(row.title)}">${escapeHtml(row.title)}</td><td>${money.format(row.net)}</td><td class="expense-cell"><strong>${money.format(row.expense)}</strong></td><td class="ratio-cell"><strong>${expenseRatio}</strong></td></tr>`;
  }).join("");
  el("promotionCount").textContent = `共 ${integer.format(filtered.length)} 条`;
}

function escapeHtml(value) {
  return String(value).replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[char]));
}

function locateHeader(rows) {
  return rows.findIndex((row) => {
    const cells = row.map((cell) => String(cell).trim());
    return (cells.includes("商品ID") && cells.includes("商品名称")) || (cells.includes("主体ID") && cells.includes("花费"));
  });
}

async function parseWorkbook(file) {
  if (!window.XLSX) throw new Error("Excel 读取组件尚未载入，请刷新页面后重试。");
  const buffer = await file.arrayBuffer();
  const isCsv = file.name.toLowerCase().endsWith(".csv");
  const workbook = isCsv
    ? XLSX.read(new TextDecoder("gb18030").decode(buffer), { type: "string", cellDates: true })
    : XLSX.read(buffer, { type: "array", cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });
  const headerRow = locateHeader(rows);
  if (headerRow < 0) throw new Error(`${file.name}：未识别到销售或花费字段。`);
  const headers = rows[headerRow].map((value) => String(value).trim());
  if (headers.includes("主体ID") && headers.includes("花费")) {
    const required = ["日期", "主体ID", "主体名称", "花费"];
    const positions = Object.fromEntries(required.map((name) => [name, headers.indexOf(name)]));
    if (required.some((name) => positions[name] < 0)) throw new Error(`${file.name}：缺少 ${required.filter((name) => positions[name] < 0).join("、")}。`);
    return { type: "expense", rows: rows.slice(headerRow + 1).filter((row) => normalizeId(row[positions.主体ID])).map((row) => ({
      date: normalizeDate(row[positions.日期]), id: normalizeId(row[positions.主体ID]), title: String(row[positions.主体名称] || "").trim(), expense: numeric(row[positions.花费]),
    })) };
  }
  const required = ["统计日期", "商品ID", "商品名称", "支付金额", "成功退款金额"];
  const positions = Object.fromEntries(required.map((name) => [name, headers.indexOf(name)]));
  if (required.some((name) => positions[name] < 0)) throw new Error(`${file.name}：缺少 ${required.filter((name) => positions[name] < 0).join("、")}。`);
  return { type: "sales", rows: rows.slice(headerRow + 1).filter((row) => normalizeId(row[positions.商品ID])).map((row) => ({
    date: normalizeDate(row[positions.统计日期]), id: normalizeId(row[positions.商品ID]), title: String(row[positions.商品名称] || "").trim(), payment: numeric(row[positions.支付金额]), refund: numeric(row[positions.成功退款金额]),
  })) };
}

async function importFiles(files) {
  if (!files.length) return;
  try {
    el("syncText").textContent = "正在处理数据文件";
    const batches = await Promise.all([...files].map(parseWorkbook));
    const merged = new Map(state.records.map((row) => [`${row.date}|${row.id}`, row]));
    let imported = 0;
    for (const batch of batches) {
      for (const row of batch.rows) {
        const key = `${row.date}|${row.id}`;
        const current = merged.get(key) || { date: row.date, id: row.id, title: row.title, payment: 0, refund: 0, expense: 0 };
        merged.set(key, batch.type === "sales"
          ? { ...current, ...row, expense: numeric(current.expense) }
          : { ...current, title: current.title || row.title, expense: row.expense });
        imported += 1;
      }
    }
    state.records = [...merged.values()];
    await persist();
    const dates = state.records.map((row) => row.date).filter(Boolean).sort();
    if (!el("startDate").value) el("startDate").value = dates[0] || "";
    el("endDate").value = dates.at(-1) || el("endDate").value;
    el("syncText").textContent = `${integer.format(state.records.length)} 条数据已保存`;
    state.page = 1;
    render();
    showToast(`已处理 ${integer.format(imported)} 条销售/花费记录；重复日期与商品已自动覆盖。`);
  } catch (error) {
    el("syncText").textContent = "导入失败";
    showToast(error.message || "数据导入失败，请检查文件格式。");
  }
}

let toastTimer;
function showToast(message) {
  clearTimeout(toastTimer);
  el("toast").textContent = message;
  el("toast").classList.add("show");
  toastTimer = setTimeout(() => el("toast").classList.remove("show"), 4200);
}

let webMcpRegistered = false;
function registerWebMcp() {
  if (webMcpRegistered || !document.modelContext?.registerTool) return;
  webMcpRegistered = true;
  document.modelContext.registerTool({
    name: "get_sales_summary",
    title: "读取销售汇总",
    description: "读取当前日期区间内，按负责人和空净、配件、清洁液、汽车分类汇总的净销售额。",
    inputSchema: { type: "object", properties: {}, additionalProperties: false },
    annotations: { readOnlyHint: true, untrustedContentHint: false },
    execute() {
      const snapshot = aggregate(activeRecords());
      const rows = [...snapshot.byOwner.entries()].map(([owner, values]) => ({ owner, ...values, expenseRatio: values.totalNet ? values.totalExpense / values.totalNet : null }));
      return { startDate: el("startDate").value, endDate: el("endDate").value, total: snapshot.total, owners: rows };
    },
  });
}

el("fileInput").addEventListener("change", (event) => importFiles(event.target.files));
["startDate", "endDate"].forEach((id) => el(id).addEventListener("change", () => { state.page = 1; render(); }));
["searchInput", "detailStartDate", "detailEndDate", "ownerFilter", "categoryFilter"].forEach((id) => el(id).addEventListener(id === "searchInput" ? "input" : "change", () => { state.page = 1; renderDetails(); }));
el("promotionRiskFilter").addEventListener("change", () => renderPromotionAnalysis());
el("prevPage").addEventListener("click", () => { state.page -= 1; renderDetails(); });
el("nextPage").addEventListener("click", () => { state.page += 1; renderDetails(); });
const dropZone = el("dropZone");
["dragenter", "dragover"].forEach((name) => dropZone.addEventListener(name, (event) => { event.preventDefault(); dropZone.classList.add("drag"); }));
["dragleave", "drop"].forEach((name) => dropZone.addEventListener(name, (event) => { event.preventDefault(); dropZone.classList.remove("drag"); }));
dropZone.addEventListener("drop", (event) => importFiles(event.dataTransfer.files));

initialize().catch((error) => { el("syncText").textContent = "载入失败"; showToast(error.message || "初始数据载入失败"); });
