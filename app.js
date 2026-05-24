const state = {
  sap: null,
  logic: null,
  result: null,
  currentTable: "group1",
};

const fieldLabels = {
  material: "物料/SKU",
  quantity: "数量字段",
  posting_date: "过账日期",
  external_doc: "外部系统单号",
  unloading_point: "卸货点",
  change_datetime: "变动时间",
  doc_type: "单据类型",
  pre_doc: "前置单据编码",
  receipt_doc: "收发货单据",
  warehouse: "逻辑仓编码",
};

const requiredFields = {
  sap: new Set(["material", "quantity", "posting_date", "external_doc", "unloading_point"]),
  logic: new Set(["material", "quantity", "change_datetime", "doc_type", "pre_doc", "receipt_doc", "warehouse"]),
};

const defaultFields = {
  sap: {
    material: ["物料", "SKU编码", "SKU", "商品编码"],
    quantity: ["数量", "变更数量", "移动数量"],
    posting_date: ["过账日期", "输入日期", "变动日期"],
    external_doc: ["外部系统单号", "原始单据号(记录自动产生的单据的源头单据)", "参照"],
    unloading_point: ["卸货点"],
  },
  logic: {
    material: ["SKU编码", "物料", "SKU", "商品编码"],
    quantity: ["变更数量", "数量", "移动数量"],
    change_datetime: ["变动时间"],
    doc_type: ["单据类型"],
    pre_doc: ["前置单据编码"],
    receipt_doc: ["收发货单据"],
    warehouse: ["逻辑仓编码", "库存地点", "仓库", "仓库编码"],
  },
};

const group1LogicTypes = [
  "零售订单",
  "零售退单",
  "JIT唯品退供单",
  "JIT唯品配货单",
  "零售发货通知单",
  "零售退货通知单",
  "平台寄售",
  "平台退供",
  "shein",
  "拼多多TEMU",
];
const group1ExportDocKeys = ["BMS", "ARST", "RO", "RC"];
const group3LogicTypes = ["样衣调出", "样衣领用", "样衣归还"];
const group4LogicTypes = ["销售出库单", "销售退货单", "销售退单", "返修出库", "采购交货单"];
const group4ExportDocKeys = ["XSD", "XSTD"];

const tableTitles = {
  group1: "分组一：零售/JIT/平台按日期数量差异",
  group2: "分组二：其他出入库/库存调整缺失单号",
  group3: "分组三：样衣 RR 单号缺失",
  group4: "分组四：销售/退货/返修/采购按日期数量差异",
  group5: "分组五：采购退货缺失单号",
};

function $(selector) {
  return document.querySelector(selector);
}

function setStatus(side, text, kind = "") {
  const el = $(`#${side}Status`);
  el.textContent = text;
  el.className = `status-pill ${kind}`.trim();
}

function formatNumber(value) {
  return new Intl.NumberFormat("zh-CN", { maximumFractionDigits: 4 }).format(value || 0);
}

function normalizeHeader(value) {
  return String(value ?? "").trim().replace(/\s+/g, "").toLowerCase();
}

function normalizeText(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return formatDateTime(value);
  if (typeof value === "number" && Number.isInteger(value)) return String(value);
  return String(value).trim();
}

function normalizeNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const text = String(value).trim().replaceAll(",", "");
  if (!text) return 0;
  const parsed = Number(text);
  if (Number.isFinite(parsed)) return parsed;
  const match = text.match(/-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : 0;
}

function pad2(value) {
  return String(value).padStart(2, "0");
}

function formatDateTime(value) {
  return `${value.getFullYear()}-${pad2(value.getMonth() + 1)}-${pad2(value.getDate())} ${pad2(
    value.getHours()
  )}:${pad2(value.getMinutes())}:${pad2(value.getSeconds())}`;
}

function parseDay(value) {
  if (value instanceof Date) return formatDateTime(value).slice(0, 10);
  const text = normalizeText(value);
  if (!text) return "";
  const match = text.match(/\d{4}[-/]\d{1,2}[-/]\d{1,2}/);
  if (!match) return text.slice(0, 10);
  const parts = match[0].replaceAll("/", "-").split("-");
  return `${parts[0]}-${pad2(parts[1])}-${pad2(parts[2])}`;
}

function containsAny(value, keywords) {
  const text = normalizeText(value).toLowerCase();
  return keywords.some((keyword) => text.includes(String(keyword).toLowerCase()));
}

function rrOrSrText(value) {
  return /\b(?:RR|SR)[A-Za-z0-9_-]*/i.test(normalizeText(value));
}

function extractRrDoc(value) {
  const match = normalizeText(value).match(/\bRR[A-Za-z0-9_-]*/i);
  return match ? match[0].toUpperCase() : "";
}

function normalizeDoc(value) {
  return normalizeText(value).toUpperCase();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function detectHeader(rows) {
  let bestIndex = 0;
  let bestScore = -1;
  rows.slice(0, 12).forEach((row, index) => {
    const values = row.map(normalizeText);
    const nonEmpty = values.filter(Boolean).length;
    const unique = new Set(values.filter(Boolean)).size;
    let score = nonEmpty + unique;
    if (values.some((value) => ["物料", "SKU编码", "变更数量", "数量"].includes(value))) score += 20;
    if (score > bestScore) {
      bestIndex = index;
      bestScore = score;
    }
  });
  return bestIndex;
}

function makeUniqueHeaders(row) {
  const seen = new Map();
  return row.map((value, index) => {
    const base = normalizeText(value) || `空列${index + 1}`;
    const count = (seen.get(base) || 0) + 1;
    seen.set(base, count);
    return count > 1 ? `${base}_${count}` : base;
  });
}

function guessMapping(headers, side) {
  const normalized = new Map(headers.map((header) => [normalizeHeader(header), header]));
  const mapping = {};
  Object.entries(defaultFields[side]).forEach(([field, candidates]) => {
    let selected = "";
    for (const candidate of candidates) {
      const direct = normalized.get(normalizeHeader(candidate));
      if (direct) {
        selected = direct;
        break;
      }
    }
    if (!selected) {
      selected =
        headers.find((header) =>
          candidates.some((candidate) => normalizeHeader(candidate) && normalizeHeader(header).includes(normalizeHeader(candidate)))
        ) || "";
    }
    mapping[field] = selected;
  });
  return mapping;
}

function readWorkbook(file, side) {
  return new Promise((resolve, reject) => {
    if (!window.XLSX) {
      reject(new Error("Excel 解析库没有加载成功，请刷新页面后重试。"));
      return;
    }
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("读取文件失败"));
    reader.onload = () => {
      try {
        const workbook = XLSX.read(reader.result, { type: "array", cellDates: true });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
        const headerIndex = detectHeader(matrix);
        const headers = makeUniqueHeaders(matrix[headerIndex] || []);
        const rows = matrix.slice(headerIndex + 1).flatMap((row, offset) => {
          if (!row.some((cell) => normalizeText(cell))) return [];
          const record = { _row_number: headerIndex + offset + 2 };
          headers.forEach((header, index) => {
            record[header] = row[index] instanceof Date ? formatDateTime(row[index]) : row[index] ?? "";
          });
          return [record];
        });
        resolve({
          id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random()}`,
          name: file.name,
          rows,
          workbook: {
            sheet: sheetName,
            headers,
            header_row: headerIndex + 1,
            max_row: matrix.length,
            max_column: headers.length,
          },
          mapping_guess: guessMapping(headers, side),
        });
      } catch (error) {
        reject(error);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

function renderMapping(side, upload) {
  const target = $(`#${side}Mapping`);
  const headers = upload.workbook.headers;
  const mapping = upload.mapping_guess || {};
  const fields = Object.keys(fieldLabels).filter((field) => field in mapping);
  target.innerHTML = fields
    .map((field) => {
      const options = ['<option value="">不使用</option>']
        .concat(
          headers.map((header) => {
            const selected = mapping[field] === header ? "selected" : "";
            return `<option value="${escapeHtml(header)}" ${selected}>${escapeHtml(header)}</option>`;
          })
        )
        .join("");
      const required = requiredFields[side]?.has(field) ? "required" : "";
      return `
        <div class="field ${required}">
          <label for="${side}_${field}">${fieldLabels[field]}</label>
          <select id="${side}_${field}" data-side="${side}" data-field="${field}">
            ${options}
          </select>
        </div>
      `;
    })
    .join("");

  target.querySelectorAll("select").forEach((select) => {
    select.addEventListener("change", () => {
      state[side].mapping_guess[select.dataset.field] = select.value;
      refreshCompareButton();
    });
  });
}

async function uploadFile(side, file) {
  setStatus(side, "读取中");
  $(`#${side}Meta`).textContent = "";
  try {
    const payload = await readWorkbook(file, side);
    state[side] = payload;
    setStatus(side, "已上传", "ready");
    $(`#${side}Meta`).textContent = `${payload.name} · ${payload.workbook.sheet} · ${formatNumber(
      payload.rows.length
    )} 行 · ${payload.workbook.max_column} 列`;
    renderMapping(side, payload);
    refreshCompareButton();
  } catch (error) {
    setStatus(side, "读取失败", "error");
    $(`#${side}Meta`).textContent = error.message || "读取失败";
  }
}

function refreshCompareButton() {
  const ready = ["sap", "logic"].every((side) => {
    const upload = state[side];
    if (!upload) return false;
    return [...requiredFields[side]].every((field) => upload.mapping_guess[field]);
  });
  $("#compareButton").disabled = !ready;
}

function value(row, fields, field) {
  return normalizeText(row[fields[field]]);
}

function quantity(row, fields) {
  return normalizeNumber(row[fields.quantity]);
}

function logicDay(row, fields) {
  return parseDay(row[fields.change_datetime]);
}

function sapDay(row, fields) {
  return parseDay(row[fields.posting_date]);
}

function groupSum(rows, fields, dayFn, docField) {
  const grouped = new Map();
  rows.forEach((row) => {
    const material = value(row, fields, "material");
    const day = dayFn(row, fields);
    if (!material || !day) return;
    const key = `${material}\u0000${day}`;
    const bucket = grouped.get(key) || { material, day, quantity: 0, rows: 0, docs: new Set() };
    bucket.quantity += quantity(row, fields);
    bucket.rows += 1;
    if (docField) {
      const doc = value(row, fields, docField);
      if (doc) bucket.docs.add(doc);
    }
    grouped.set(key, bucket);
  });
  return grouped;
}

function compareDateGroup(groupName, logicRows, sapRows, logicFields, sapFields, options = {}) {
  const logicGrouped = groupSum(logicRows, logicFields, logicDay, options.logicDocField);
  const sapGrouped = groupSum(sapRows, sapFields, sapDay, options.sapDocField);
  const keys = [...new Set([...logicGrouped.keys(), ...sapGrouped.keys()])].sort();
  return keys.flatMap((key) => {
    const logicBucket = logicGrouped.get(key) || { quantity: 0, rows: 0, docs: new Set() };
    const sapBucket = sapGrouped.get(key) || { quantity: 0, rows: 0, docs: new Set() };
    const [material, day] = key.split("\u0000");
    const diff = logicBucket.quantity - sapBucket.quantity;
    if (Math.abs(diff) < 0.0000001) return [];
    const row = {
      分组: groupName,
      "SKU/物料": material,
      日期: day,
      逻辑仓数量: Number(logicBucket.quantity.toFixed(6)),
      EXPORT数量: Number(sapBucket.quantity.toFixed(6)),
      差异数量: Number(diff.toFixed(6)),
    };
    if (options.logicDocField) row["逻辑仓前置单据编码"] = [...logicBucket.docs].sort().join("、");
    if (options.sapDocField) row["SAP外部系统单号"] = [...sapBucket.docs].sort().join("、");
    if (options.includeCounts) {
      row["逻辑仓行数"] = logicBucket.rows;
      row["EXPORT行数"] = sapBucket.rows;
    }
    return [row];
  });
}

function missingDocsFromExport(groupName, logicRows, sapRows, logicFields, sapFields, logicDocField, options = {}) {
  const transform = options.docTransform || normalizeDoc;
  const sapDocs = new Set(
    sapRows
      .map((row) => `${value(row, sapFields, "material")}\u0000${transform(value(row, sapFields, options.sapDocField || "external_doc"))}`)
      .filter((key) => !key.endsWith("\u0000"))
  );
  const seen = new Set();
  return logicRows.flatMap((row) => {
    const doc = value(row, logicFields, logicDocField);
    const docKey = transform(doc);
    if (!docKey) return [];
    const material = value(row, logicFields, "material");
    const key = `${material}\u0000${docKey}`;
    if (sapDocs.has(key) || seen.has(key)) return [];
    seen.add(key);
    return [
      {
        分组: groupName,
        "SKU/物料": material,
        逻辑仓单号: docKey,
        单据类型: value(row, logicFields, "doc_type"),
        逻辑仓编码: value(row, logicFields, "warehouse"),
        变动日期: logicDay(row, logicFields),
        变更数量: quantity(row, logicFields),
        结果: "EXPORT不存在",
      },
    ];
  });
}

function compareRecords(sapRows, logicRows, sapFields, logicFields) {
  const logicG1 = logicRows.filter((row) => containsAny(value(row, logicFields, "doc_type"), group1LogicTypes));
  const sapG1 = sapRows.filter((row) => containsAny(value(row, sapFields, "external_doc"), group1ExportDocKeys));
  const logicG2 = logicRows.filter((row) => {
    const docType = value(row, logicFields, "doc_type");
    return (
      (containsAny(docType, ["其他出入库单"]) && value(row, logicFields, "warehouse") !== "L014") ||
      containsAny(docType, ["库存调整单"])
    );
  });
  const logicG3 = logicRows.filter(
    (row) =>
      containsAny(value(row, logicFields, "doc_type"), group3LogicTypes) &&
      extractRrDoc(value(row, logicFields, "receipt_doc")) &&
      value(row, logicFields, "warehouse") !== "L006"
  );
  const sapG3 = sapRows.filter((row) => !value(row, sapFields, "external_doc") && rrOrSrText(value(row, sapFields, "unloading_point")));
  const logicG4 = logicRows.filter((row) => containsAny(value(row, logicFields, "doc_type"), group4LogicTypes));
  const sapG4 = sapRows.filter((row) => {
    const externalDoc = value(row, sapFields, "external_doc");
    const unloadingPoint = value(row, sapFields, "unloading_point");
    return (!externalDoc && !rrOrSrText(unloadingPoint)) || (externalDoc && containsAny(externalDoc, group4ExportDocKeys));
  });
  const logicG5 = logicRows.filter((row) => containsAny(value(row, logicFields, "doc_type"), ["采购退货单"]));

  const group1 = compareDateGroup("分组一", logicG1, sapG1, logicFields, sapFields, {
    logicDocField: "pre_doc",
    sapDocField: "external_doc",
  });
  const group2 = missingDocsFromExport("分组二", logicG2, sapRows, logicFields, sapFields, "pre_doc");
  const group3 = missingDocsFromExport("分组三", logicG3, sapG3, logicFields, sapFields, "receipt_doc", {
    sapDocField: "unloading_point",
    docTransform: extractRrDoc,
  });
  const group4 = compareDateGroup("分组四", logicG4, sapG4, logicFields, sapFields, { logicDocField: "pre_doc" });
  const group5 = missingDocsFromExport("分组五", logicG5, sapRows, logicFields, sapFields, "receipt_doc");

  return {
    counts: {
      sap_rows: sapRows.length,
      logic_rows: logicRows.length,
      group1: group1.length,
      group2: group2.length,
      group3: group3.length,
      group4: group4.length,
      group5: group5.length,
    },
    group1,
    group2,
    group3,
    group4,
    group5,
  };
}

async function compare() {
  $("#compareButton").disabled = true;
  $("#compareButton").textContent = "对比中...";
  try {
    state.result = compareRecords(state.sap.rows, state.logic.rows, state.sap.mapping_guess, state.logic.mapping_guess);
    state.currentTable = "group1";
    renderResults();
  } catch (error) {
    alert(error.message || "对比失败");
  } finally {
    $("#compareButton").disabled = false;
    $("#compareButton").textContent = "开始对比";
  }
}

function renderResults() {
  $("#results").hidden = false;
  const counts = state.result.counts;
  const metricItems = [
    ["SAP 行数", counts.sap_rows],
    ["逻辑仓行数", counts.logic_rows],
    ["分组一", counts.group1],
    ["分组二", counts.group2],
    ["分组三", counts.group3],
    ["分组四", counts.group4],
    ["分组五", counts.group5],
  ];
  $("#metrics").innerHTML = metricItems
    .map(([label, itemValue]) => `<div class="metric"><span>${label}</span><strong>${formatNumber(itemValue)}</strong></div>`)
    .join("");
  document.querySelectorAll(".tab").forEach((tab) => {
    tab.classList.toggle("active", tab.dataset.table === state.currentTable);
  });
  renderTable(state.currentTable);
}

function renderTable(name) {
  const rows = state.result[name] || [];
  $("#tableTitle").textContent = tableTitles[name] || name;
  const table = $("#resultTable");
  if (!rows.length) {
    table.innerHTML = `<tbody><tr><td class="empty-state">当前表没有差异记录</td></tr></tbody>`;
    return;
  }
  const columns = Object.keys(rows[0]);
  table.innerHTML = `
    <thead><tr>${columns.map((column) => `<th>${escapeHtml(column)}</th>`).join("")}</tr></thead>
    <tbody>
      ${rows
        .map((row) => `<tr>${columns.map((column) => `<td>${formatCell(row[column])}</td>`).join("")}</tr>`)
        .join("")}
    </tbody>
  `;
}

function formatCell(value) {
  if (Array.isArray(value) || (value && typeof value === "object")) {
    return `<code>${escapeHtml(JSON.stringify(value, null, 0))}</code>`;
  }
  return escapeHtml(value);
}

function makeCsv(rows) {
  if (!rows.length) return "";
  const columns = Object.keys(rows[0]);
  const escapeCsv = (value) => {
    const text = String(value ?? "");
    return /[",\n\r]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
  };
  return [columns.join(","), ...rows.map((row) => columns.map((column) => escapeCsv(row[column])).join(","))].join("\n");
}

function downloadCurrent() {
  if (!state.result) return;
  const rows = state.result[state.currentTable] || [];
  const csv = `\ufeff${makeCsv(rows)}`;
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${state.currentTable}.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  $("#exportStatus").textContent = `已生成 ${rows.length} 行`;
}

function bindUpload(side) {
  const input = $(`#${side}File`);
  const zone = input.closest(".dropzone");
  input.addEventListener("change", () => {
    const file = input.files?.[0];
    if (file) uploadFile(side, file);
  });
  ["dragenter", "dragover"].forEach((eventName) => {
    zone.addEventListener(eventName, (event) => {
      event.preventDefault();
      zone.classList.add("dragover");
    });
  });
  ["dragleave", "drop"].forEach((eventName) => {
    zone.addEventListener(eventName, (event) => {
      event.preventDefault();
      zone.classList.remove("dragover");
    });
  });
  zone.addEventListener("drop", (event) => {
    const file = event.dataTransfer?.files?.[0];
    if (file) uploadFile(side, file);
  });
}

bindUpload("sap");
bindUpload("logic");

$("#compareButton").addEventListener("click", compare);
$("#downloadButton").addEventListener("click", downloadCurrent);
document.querySelectorAll(".tab").forEach((tab) => {
  tab.addEventListener("click", () => {
    state.currentTable = tab.dataset.table;
    $("#exportStatus").textContent = "";
    renderResults();
  });
});
