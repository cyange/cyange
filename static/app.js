const state = {
  sap: null,
  logic: null,
  result: null,
  currentTable: "group1",
  worker: null,
  taskId: 0,
  pending: new Map(),
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

function ensureWorker() {
  if (state.worker) return state.worker;
  state.worker = new Worker("worker.js");
  state.worker.onmessage = (event) => {
    const { id, status, payload, error, message } = event.data;
    const task = state.pending.get(id);
    if (!task) return;
    if (status === "progress") {
      task.onProgress?.(message);
      return;
    }
    state.pending.delete(id);
    if (status === "done") {
      task.resolve(payload);
    } else {
      task.reject(new Error(error || "任务失败"));
    }
  };
  state.worker.onerror = (event) => {
    const error = new Error(event.message || "后台解析线程异常");
    state.pending.forEach((task) => task.reject(error));
    state.pending.clear();
  };
  return state.worker;
}

function workerTask(action, payload, transfer = [], onProgress) {
  const worker = ensureWorker();
  const id = ++state.taskId;
  return new Promise((resolve, reject) => {
    state.pending.set(id, { resolve, reject, onProgress });
    worker.postMessage({ id, action, ...payload }, transfer);
  });
}

function setStatus(side, text, kind = "") {
  const el = $(`#${side}Status`);
  el.textContent = text;
  el.className = `status-pill ${kind}`.trim();
}

function formatNumber(value) {
  return new Intl.NumberFormat("zh-CN", { maximumFractionDigits: 4 }).format(value || 0);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
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
  $(`#${side}Meta`).textContent = `${file.name} · 正在读取文件...`;
  refreshCompareButton();
  try {
    const buffer = await file.arrayBuffer();
    $(`#${side}Meta`).textContent = `${file.name} · 正在后台解析，请稍等...`;
    const payload = await workerTask(
      "parse",
      { side, name: file.name, buffer },
      [buffer],
      (message) => {
        $(`#${side}Meta`).textContent = `${file.name} · ${message}`;
      }
    );
    state[side] = payload;
    setStatus(side, "已上传", "ready");
    $(`#${side}Meta`).textContent = `${payload.name} · ${payload.workbook.sheet} · ${formatNumber(
      payload.row_count
    )} 行 · ${payload.workbook.max_column} 列`;
    renderMapping(side, payload);
    refreshCompareButton();
  } catch (error) {
    state[side] = null;
    setStatus(side, "读取失败", "error");
    $(`#${side}Meta`).textContent = error.message || "读取失败";
    refreshCompareButton();
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

async function compare() {
  $("#compareButton").disabled = true;
  $("#compareButton").textContent = "对比中...";
  $("#exportStatus").textContent = "后台对比中，请稍等...";
  try {
    state.result = await workerTask(
      "compare",
      {
        sapFields: state.sap.mapping_guess,
        logicFields: state.logic.mapping_guess,
      },
      [],
      (message) => {
        $("#exportStatus").textContent = message;
      }
    );
    state.currentTable = "group1";
    $("#exportStatus").textContent = "";
    renderResults();
  } catch (error) {
    $("#exportStatus").textContent = "";
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
    .map(([label, value]) => `<div class="metric"><span>${label}</span><strong>${formatNumber(value)}</strong></div>`)
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
