importScripts("xlsx.full.min.js");

const datasets = {
  sap: null,
  logic: null,
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
  "京东退货单",
];
const group1ExportDocKeys = ["BMS", "ARST", "RO", "RC"];
const group3LogicTypes = ["样衣调出", "样衣领用", "样衣归还"];
const group4LogicTypes = ["销售出库单", "销售退货单", "销售退单", "返修出库", "采购交货单"];
const group4ExportDocKeys = ["XSD", "XSTD"];

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
  const text = normalizeText(value);
  if (!text) return "";
  const yearFirst = text.match(/\d{4}[-/]\d{1,2}[-/]\d{1,2}/);
  if (yearFirst) {
    const parts = yearFirst[0].replaceAll("/", "-").split("-");
    return `${parts[0]}-${pad2(parts[1])}-${pad2(parts[2])}`;
  }
  const monthFirst = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})(?:\b|\s)/);
  if (monthFirst) {
    const yearNumber = Number(monthFirst[3]);
    const fullYear = yearNumber < 100 ? 2000 + yearNumber : yearNumber;
    return `${fullYear}-${pad2(monthFirst[1])}-${pad2(monthFirst[2])}`;
  }
  return text.slice(0, 10);
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

function detectHeader(rows, maxColumn = 0) {
  let bestIndex = 0;
  let bestScore = -1;
  rows.slice(0, 12).forEach((row, index) => {
    const values = rowCells(row, maxColumn || row.length).map(normalizeText);
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
  return row.map((cell, index) => {
    const base = normalizeText(cell) || `空列${index + 1}`;
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

function parseWorkbook(buffer, side, name) {
  const workbook = XLSX.read(buffer, { type: "array", cellDates: false });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const { matrix, rowNumbers, maxRow, maxColumn } = sheetToMatrix(sheet);
  if (!matrix.length) throw new Error("工作表没有可读取内容");
  const headerIndex = detectHeader(matrix, maxColumn);
  const headers = makeUniqueHeaders(rowCells(matrix[headerIndex], maxColumn));
  const rows = [];
  for (let i = headerIndex + 1; i < matrix.length; i += 1) {
    const row = rowCells(matrix[i], headers.length);
    if (!row.some((cell) => normalizeText(cell))) continue;
    const record = { _row_number: rowNumbers[i] || i + 1 };
    headers.forEach((header, index) => {
      const cell = row[index];
      record[header] = cell instanceof Date ? formatDateTime(cell) : cell ?? "";
    });
    rows.push(record);
  }
  return {
    id: `${Date.now()}-${Math.random()}`,
    name,
    rows,
    workbook: {
      sheet: sheetName,
      headers,
      header_row: rowNumbers[headerIndex] || headerIndex + 1,
      max_row: maxRow,
      max_column: maxColumn || headers.length,
    },
    mapping_guess: guessMapping(headers, side),
  };
}

function cellDisplayValue(cell) {
  if (!cell) return "";
  if (cell.w !== undefined) return cell.w;
  if (cell.v === undefined || cell.v === null) return "";
  return cell.v;
}

function rowCells(row, length = row?.length || 0) {
  return Array.from({ length }, (_, index) => row?.[index] ?? "");
}

function sheetToMatrix(sheet) {
  const rows = new Map();
  let maxRow = 0;
  let maxColumn = 0;

  Object.keys(sheet).forEach((address) => {
    if (address.startsWith("!")) return;
    const displayValue = cellDisplayValue(sheet[address]);
    if (!normalizeText(displayValue)) return;

    const position = XLSX.utils.decode_cell(address);
    const row = rows.get(position.r) || [];
    row[position.c] = displayValue;
    rows.set(position.r, row);
    maxRow = Math.max(maxRow, position.r + 1);
    maxColumn = Math.max(maxColumn, position.c + 1);
  });

  if (!rows.size && sheet["!ref"]) {
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    maxRow = range.e.r + 1;
    maxColumn = range.e.c + 1;
  }

  const rowIndexes = [...rows.keys()].sort((a, b) => a - b);
  return {
    matrix: rowIndexes.map((rowIndex) => rows.get(rowIndex)),
    rowNumbers: rowIndexes.map((rowIndex) => rowIndex + 1),
    maxRow,
    maxColumn,
  };
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

  return {
    counts: {
      sap_rows: sapRows.length,
      logic_rows: logicRows.length,
    },
    group1: compareDateGroup("分组一", logicG1, sapG1, logicFields, sapFields, {
      logicDocField: "pre_doc",
      sapDocField: "external_doc",
    }),
    group2: missingDocsFromExport("分组二", logicG2, sapRows, logicFields, sapFields, "pre_doc"),
    group3: missingDocsFromExport("分组三", logicG3, sapG3, logicFields, sapFields, "receipt_doc", {
      sapDocField: "unloading_point",
      docTransform: extractRrDoc,
    }),
    group4: compareDateGroup("分组四", logicG4, sapG4, logicFields, sapFields, { logicDocField: "pre_doc" }),
    group5: missingDocsFromExport("分组五", logicG5, sapRows, logicFields, sapFields, "receipt_doc"),
  };
}

self.onmessage = (event) => {
  const { id, action } = event.data;
  try {
    if (action === "parse") {
      const { buffer, side, name } = event.data;
      self.postMessage({ id, status: "progress", message: "解析 Excel 中..." });
      const payload = parseWorkbook(buffer, side, name);
      datasets[side] = payload.rows;
      self.postMessage({
        id,
        status: "done",
        payload: {
          id: payload.id,
          name: payload.name,
          row_count: payload.rows.length,
          workbook: payload.workbook,
          mapping_guess: payload.mapping_guess,
        },
      });
      return;
    }
    if (action === "compare") {
      const { sapFields, logicFields } = event.data;
      if (!datasets.sap || !datasets.logic) {
        throw new Error("请先上传两份 Excel 文件");
      }
      self.postMessage({ id, status: "progress", message: "执行五组规则对比中..." });
      const result = compareRecords(datasets.sap, datasets.logic, sapFields, logicFields);
      result.counts.group1 = result.group1.length;
      result.counts.group2 = result.group2.length;
      result.counts.group3 = result.group3.length;
      result.counts.group4 = result.group4.length;
      result.counts.group5 = result.group5.length;
      self.postMessage({ id, status: "done", payload: result });
    }
  } catch (error) {
    self.postMessage({ id, status: "error", error: error.message || String(error) });
  }
};
