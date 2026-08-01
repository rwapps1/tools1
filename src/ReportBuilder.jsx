import { useState, useEffect, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  Plus,
  Trash2,
  Download,
  Upload,
  ChevronRight,
  X,
  Check,
  AlertTriangle,
  FileSpreadsheet,
  ArrowRight,
  ArrowLeft,
  Pencil,
  FileUp,
  RotateCcw,
} from "lucide-react";

const COLORS = {
  bg: "#F0F4F8",
  surface: "#FFFFFF",
  ink: "#161B24",
  muted: "#5B6779",
  border: "#DCE3EB",
  accent: "#2C4F76",
  accentSoft: "#E2EAF3",
  accent2: "#5B84A8",
  accent2Soft: "#E8EFF6",
  danger: "#A23B2E",
  dangerSoft: "#F6E6E3",
  duplicate: "#FDE68A",
};

const FONTS = {
  display: "'Fraunces', Georgia, serif",
  body: "'Public Sans', system-ui, -apple-system, sans-serif",
  mono: "'IBM Plex Mono', ui-monospace, SFMono-Regular, monospace",
};

const LOGO_DATA_URI = "data:image/png;base64,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AAAAAAAAAAAAAAAADBEwAAAAAAAAAAAAAAAAhggYAAAAAAAAAAAAAAAAQwQMAAAAAAAAAAAAAACAIQIGAAAAAAAAAAAAAADAEAEDAAAAAAAAAAAAAABgiIABAAAAAAAAAAAAAAAwRMAAAAAAAAAAAAAAAAAYImAAAAAAAAAAAAAAAAAMETAAAAAAAAAAAAAAAACGCBgAAAAAAAAAAAAAAABDBAwAAAAAAAAAAAAAAIAhAgYAAAAAAAAAAAAAAMAQAQMAAAAAAAAAAAAAAGCIgAEAAAAAAAAAAAAAADBEwAAAAAAAAAAAAAAAABgiYAAAAAAAAAAAAAAAAAwRMAAAAAAAAAAAAAAAAIYIGAAAAAAAAAAAAAAAAEMEDAAAAAAAAAAAAAAAgCECBgAAAAAAAAAAAAAAwBABAwAAAAAAAAAAAAAAYIiAAQAAAAAAAAAAAAAAMETAAAAAAAAAAAAAAAAAGCJgAAAAAAAAAAAAAAAADBEwAAAAAAAAAAAAAAAAhggYAAAAAAAAAAAAAAAAQwQMAAAAAAAAAAAAAACAIQIGAAAAAAAAAAAAAADAEAEDAAAAAAAAAAAAAABgiIABAAAAAAAAAAAAAAAwRMAAAAAAAAAAAAAAAAAYImAAAAAAAAAAAAAAAAAMETAAAAAAAAAAAAAAAACGCBgAAAAAAAAAAAAAAABDBAwAAAAAAAAAAAAAAIAhAgYAAAAAAAAAAAAAAMAQAQMAAAAAAAAAAAAAAGCIgAEAAAAAAAAAAAAAADBEwAAAAAAAAAAAAAAAABgiYAAAAAAAAAAAAAAAAAwRMAAAAAAAAAAAAAAAAIYIGAAAAAAAAAAAAAAAAEMEDAAAAAAAAAAAAAAAgCECBgAAAAAAAAAAAAAAwBABAwAAAAAAAAAAAAAAYIiAAQAAAAAAAAAAAAAAMETAAAAAAAAAAAAAAAAAGCJgAAAAAAAAAAAAAAAADBEwAAAAAAAAAAAAAAAAhggYAAAAAAAAAAAAAAAAQwQMAAAAAAAAAAAAAACAIQIGAAAAAAAAAAAAAADAEAEDAAAAAAAAAAAAAABgiIABAAAAAAAAAAAAAAAwRMAAAAAAAAAAAAAAAAAYImAAAAAAAAAAAAAAAAAMETAAAAAAAAAAAAAAAACGCBgAAAAAAAAAAAAAAABDBAwAAAAAAAAAAAAAAIAhAgYAAAAAAAAAAAAAAMAQAQMAAAAAAAAAAAAAAGCIgAEAAAAAAAAAAAAAADBEwAAAAAAAAAAAAAAAABgiYAAAAAAAAAAAAAAAAAwRMAAAAAAAAAAAAAAAAIYIGAAAAAAAAAAAAAAAAEMEDAAAAAAAAAAAAAAAgCECBgAAAAAAAAAAAAAAwND/D0dqL9V54e4ZAAAAAElFTkSuQmCC";

const OPERATORS = [
  { value: "equals", label: "is" },
  { value: "not_equals", label: "is not" },
  { value: "contains", label: "contains" },
  { value: "greater", label: "is greater than" },
  { value: "less", label: "is less than" },
  { value: "between_dates", label: "is between (dates)" },
  { value: "is_blank", label: "is blank" },
  { value: "is_not_blank", label: "is not blank" },
];

const AGG_OPS = [
  { value: "sum", label: "Sum" },
  { value: "average", label: "Average" },
  { value: "count", label: "Count" },
  { value: "min", label: "Minimum" },
  { value: "max", label: "Maximum" },
];

function uid() {
  return Math.random().toString(36).slice(2, 10);
}

function dateStamp() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

// Older saved presets stored groupBy as a single column name (a plain
// string). Newer ones store an array to support grouping by more than one
// column. This normalizes either shape into an array so nothing breaks when
// loading presets saved before multi-column grouping existed.
function normalizeGroupBy(gb) {
  if (Array.isArray(gb)) return gb.filter(Boolean);
  if (typeof gb === "string" && gb) return [gb];
  return [];
}

function emptyPreset() {
  return {
    id: uid(),
    name: "",
    filters: [],
    filterMatch: "all",
    columns: "all",
    groupBy: [],
    aggregations: [],
    sorts: [],
    extraColumns: [],
    duplicateColumn: "",
    sourceHeaders: [],
  };
}

// Extracts rows/headers from one sheet of an already-parsed workbook, so
// switching sheets doesn't require re-reading the file from disk.
function sheetToData(wb, sheetName) {
  const sheet = wb.Sheets[sheetName];
  if (!sheet) return { error: `Could not find the "${sheetName}" sheet.` };
  const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  if (!json.length) return { error: `The "${sheetName}" sheet doesn't have any rows.` };
  return { rows: json, headers: Object.keys(json[0]) };
}

// Reads a CSV or Excel file. Excel workbooks default to the first (or only)
// sheet; if there's more than one, the caller gets the full sheet name list
// back so it can offer a picker without re-reading the file.
function parseFile(file) {
  return new Promise((resolve, reject) => {
    const isExcel = /\.(xlsx|xls)$/i.test(file.name);
    if (isExcel) {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
          const sheetNames = wb.SheetNames || [];
          if (!sheetNames.length) {
            reject(new Error("This workbook doesn't contain any sheets."));
            return;
          }
          const first = sheetToData(wb, sheetNames[0]);
          if (first.error) {
            reject(new Error(first.error));
            return;
          }
          resolve({
            kind: "excel",
            workbook: wb,
            sheetNames,
            selectedSheet: sheetNames[0],
            rows: first.rows,
            headers: first.headers,
          });
        } catch (err) {
          reject(new Error("Could not read this Excel file."));
        }
      };
      reader.onerror = () => reject(new Error("Could not read this file."));
      reader.readAsArrayBuffer(file);
    } else {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
          const headers = results.meta.fields || [];
          if (!headers.length) {
            reject(new Error("No columns found in this CSV."));
            return;
          }
          resolve({ kind: "csv", workbook: null, sheetNames: [], selectedSheet: "", rows: results.data, headers });
        },
        error: () => reject(new Error("Could not read this CSV file.")),
      });
    }
  });
}

// Parses a cell value (a real Date from Excel, or a string from CSV) into a Date.
// Numeric slash/dash dates are read as DD/MM/YYYY (UK style) since that's the
// common convention for the files this is built to handle. Returns null if the
// value doesn't look like a date at all.
function parseDateValue(value) {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }
  if (typeof value !== "string" || !value.trim()) return null;
  const trimmed = value.trim();

  const dmy = trimmed.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (dmy) {
    let [, d, m, y] = dmy;
    if (y.length === 2) y = (Number(y) < 50 ? "20" : "19") + y;
    const day = Number(d);
    const month = Number(m);
    const year = Number(y);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      const parsed = new Date(year, month - 1, day);
      if (!isNaN(parsed.getTime())) return parsed;
    }
  }

  const native = new Date(trimmed);
  return isNaN(native.getTime()) ? null : native;
}

function toDateOnly(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}

function matchFilter(cellValue, operator, filterValue, filterValue2) {
  if (operator === "is_blank" || operator === "is_not_blank") {
    const isBlank = cellValue === null || cellValue === undefined || String(cellValue).trim() === "";
    return operator === "is_blank" ? isBlank : !isBlank;
  }

  if (operator === "between_dates") {
    const cellDate = parseDateValue(cellValue);
    if (!cellDate) return false;
    const cellDay = toDateOnly(cellDate);
    const start = filterValue ? toDateOnly(new Date(`${filterValue}T00:00:00`)) : null;
    const end = filterValue2 ? toDateOnly(new Date(`${filterValue2}T00:00:00`)) : null;
    if (start !== null && cellDay < start) return false;
    if (end !== null && cellDay > end) return false;
    return true;
  }

  const cell = String(cellValue ?? "").toLowerCase();
  const val = String(filterValue ?? "").toLowerCase();
  const cellNum = parseFloat(cellValue);
  const valNum = parseFloat(filterValue);
  switch (operator) {
    case "equals":
      return cell === val;
    case "not_equals":
      return cell !== val;
    case "contains":
      return cell.includes(val);
    case "greater":
      return !isNaN(cellNum) && !isNaN(valNum) && cellNum > valNum;
    case "less":
      return !isNaN(cellNum) && !isNaN(valNum) && cellNum < valNum;
    default:
      return true;
  }
}

function computeAgg(op, nums) {
  if (!nums.length) return 0;
  switch (op) {
    case "sum":
      return nums.reduce((a, b) => a + b, 0);
    case "average":
      return nums.reduce((a, b) => a + b, 0) / nums.length;
    case "min":
      return Math.min(...nums);
    case "max":
      return Math.max(...nums);
    case "count":
      return nums.length;
    default:
      return 0;
  }
}

function missingColumns(preset, headers) {
  const groupBy = normalizeGroupBy(preset.groupBy);
  const referenced = new Set();
  preset.filters.forEach((f) => f.column && referenced.add(f.column));
  groupBy.forEach((c) => referenced.add(c));
  preset.aggregations.forEach((a) => a.column && referenced.add(a.column));
  // sort levels / explicit column selection only reference raw file headers
  // when there's no grouping — grouped mode sorts on derived output columns.
  if (!groupBy.length) {
    (preset.sorts || []).forEach((s) => s.column && referenced.add(s.column));
  }
  if (!groupBy.length && Array.isArray(preset.columns)) {
    preset.columns.forEach((c) => referenced.add(c));
  }
  if (!groupBy.length && preset.duplicateColumn) referenced.add(preset.duplicateColumn);
  return [...referenced].filter((c) => !headers.includes(c));
}

// Applies Excel-style multi-level sorting: the first level decides order;
// each following level only breaks ties left by the one before it.
function applySort(data, sorts) {
  const levels = (sorts || []).filter((s) => s.column);
  if (!levels.length) return;
  data.sort((a, b) => {
    for (const s of levels) {
      const av = a[s.column];
      const bv = b[s.column];
      const an = parseFloat(av);
      const bn = parseFloat(bv);
      let cmp;
      if (!isNaN(an) && !isNaN(bn)) cmp = an - bn;
      else cmp = String(av ?? "").localeCompare(String(bv ?? ""));
      if (s.direction === "desc") cmp = -cmp;
      if (cmp !== 0) return cmp;
    }
    return 0;
  });
}

function groupKeyFor(value) {
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return "(blank)";
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, "0");
    const d = String(value.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  if (value === null || value === undefined || value === "") return "(blank)";
  return String(value);
}

function detectNumericColumns(data, cols) {
  const numeric = new Set();
  const sample = data.slice(0, 25);
  cols.forEach((c) => {
    if (sample.length && sample.every((r) => r[c] === "" || r[c] == null || !isNaN(parseFloat(r[c])))) {
      numeric.add(c);
    }
  });
  return numeric;
}

// Appends preset-defined blank columns to the end of a report. Skips any name
// that's blank or collides with a column already in the output, so a mistyped
// duplicate never silently overwrites real data.
function withExtraColumns(data, columns, extraColumns) {
  const seen = new Set(columns);
  const extras = [];
  (extraColumns || []).forEach((c) => {
    const name = (c || "").trim();
    if (name && !seen.has(name)) {
      seen.add(name);
      extras.push(name);
    }
  });
  if (!extras.length) return { data, columns };
  const newColumns = [...columns, ...extras];
  const newData = data.map((row) => {
    const r = { ...row };
    extras.forEach((c) => {
      r[c] = "";
    });
    return r;
  });
  return { data: newData, columns: newColumns };
}

// Returns the set of values in `column` that occur more than once. Blank
// values are never treated as duplicates of each other.
function findDuplicateValues(rows, column) {
  const counts = new Map();
  rows.forEach((row) => {
    const key = String(row[column] ?? "").trim();
    if (!key) return;
    counts.set(key, (counts.get(key) || 0) + 1);
  });
  return new Set([...counts.entries()].filter(([, c]) => c > 1).map(([k]) => k));
}

function applyPreset(preset, rows, headers) {
  const missing = missingColumns(preset, headers);
  if (missing.length) {
    return {
      error: `This preset expects ${missing.length > 1 ? "columns" : "a column"} not in this file: ${missing.join(", ")}.`,
    };
  }

  const groupBy = normalizeGroupBy(preset.groupBy);
  const activeFilters = preset.filters.filter((f) => f.column);
  const filtered = rows.filter((row) => {
    if (!activeFilters.length) return true;
    const results = activeFilters.map((f) => matchFilter(row[f.column], f.operator, f.value, f.value2));
    return preset.filterMatch === "any" ? results.some(Boolean) : results.every(Boolean);
  });

  if (groupBy.length) {
    const groups = new Map();
    filtered.forEach((row) => {
      const keyParts = groupBy.map((col) => groupKeyFor(row[col]));
      const compositeKey = keyParts.join("\u0000");
      if (!groups.has(compositeKey)) {
        const keyValues = {};
        groupBy.forEach((col, i) => {
          keyValues[col] = keyParts[i];
        });
        groups.set(compositeKey, { keyValues, rows: [] });
      }
      groups.get(compositeKey).rows.push(row);
    });
    const aggCols = preset.aggregations.filter((a) => a.column);
    const columns = [...groupBy, "Count", ...aggCols.map((a) => `${a.op}_${a.column}`)];
    const data = [...groups.values()].map(({ keyValues, rows: groupRows }) => {
      const record = { ...keyValues, Count: groupRows.length };
      aggCols.forEach((agg) => {
        const nums = groupRows.map((r) => parseFloat(r[agg.column])).filter((n) => !isNaN(n));
        record[`${agg.op}_${agg.column}`] = computeAgg(agg.op, nums);
      });
      return record;
    });
    const validSorts = (preset.sorts || []).filter((s) => columns.includes(s.column));
    applySort(data, validSorts);
    const numericColumns = new Set(["Count", ...aggCols.map((a) => `${a.op}_${a.column}`)]);
    const withExtras = withExtraColumns(data, columns, preset.extraColumns);
    return { data: withExtras.data, columns: withExtras.columns, numericColumns, duplicateColumn: "" };
  }

  const cols = preset.columns === "all" || !preset.columns.length ? headers : preset.columns;
  const dupValues = preset.duplicateColumn ? findDuplicateValues(filtered, preset.duplicateColumn) : null;
  applySort(filtered, preset.sorts);
  const data = filtered.map((row) => {
    const r = {};
    cols.forEach((c) => (r[c] = row[c]));
    if (dupValues) {
      const key = String(row[preset.duplicateColumn] ?? "").trim();
      r.__dup = !!key && dupValues.has(key);
    }
    return r;
  });
  const numericColumns = detectNumericColumns(data, cols);
  const withExtras = withExtraColumns(data, cols, preset.extraColumns);
  return {
    data: withExtras.data,
    columns: withExtras.columns,
    numericColumns,
    duplicateColumn: dupValues ? preset.duplicateColumn : "",
  };
}

function formatCell(value, isNumeric) {
  if (value === null || value === undefined || value === "") return "\u2014";
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return "\u2014";
    return value.toLocaleDateString(undefined, { day: "2-digit", month: "short", year: "numeric" });
  }
  if (isNumeric) {
    const n = parseFloat(value);
    if (!isNaN(n)) return n.toLocaleString(undefined, { maximumFractionDigits: 2 });
  }
  return String(value);
}

function summarize(preset) {
  const parts = [];
  const groupBy = normalizeGroupBy(preset.groupBy);
  const activeFilters = preset.filters.filter((f) => f.column);
  if (activeFilters.length) {
    const joiner = activeFilters.length >= 2 ? (preset.filterMatch === "any" ? " OR " : " AND ") : ", ";
    parts.push({
      label: "Filter",
      value: activeFilters
        .map((f) => {
          if (f.operator === "between_dates") {
            return `${f.column} between ${f.value || "\u2026"} and ${f.value2 || "\u2026"}`;
          }
          if (f.operator === "is_blank" || f.operator === "is_not_blank") {
            const opLabel = OPERATORS.find((o) => o.value === f.operator)?.label || f.operator;
            return `${f.column} ${opLabel}`;
          }
          const opLabel = OPERATORS.find((o) => o.value === f.operator)?.label || f.operator;
          return `${f.column} ${opLabel} "${f.value}"`;
        })
        .join(joiner),
    });
  }
  if (groupBy.length) parts.push({ label: "Group by", value: groupBy.join(", then ") });
  if (preset.aggregations.some((a) => a.column)) {
    parts.push({
      label: "Calculate",
      value: preset.aggregations
        .filter((a) => a.column)
        .map((a) => `${AGG_OPS.find((o) => o.value === a.op)?.label} of ${a.column}`)
        .join(", "),
    });
  }
  if ((preset.sorts || []).some((s) => s.column)) {
    parts.push({
      label: "Sort by",
      value: preset.sorts
        .filter((s) => s.column)
        .map((s) => `${s.column} (${s.direction === "desc" ? "high to low" : "low to high"})`)
        .join(", then "),
    });
  }
  if (!groupBy.length && Array.isArray(preset.columns) && preset.columns.length) {
    parts.push({ label: "Columns", value: preset.columns.join(", ") });
  }
  if (!groupBy.length && preset.duplicateColumn) {
    parts.push({ label: "Highlight duplicates in", value: preset.duplicateColumn });
  }
  if ((preset.extraColumns || []).some((c) => c && c.trim())) {
    parts.push({ label: "Adds columns", value: preset.extraColumns.filter((c) => c && c.trim()).join(", ") });
  }
  return parts;
}

function cellForExport(value) {
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return "";
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, "0");
    const d = String(value.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  return value ?? "";
}

function exportCSV(data, columns, filename) {
  const csv = Papa.unparse({ fields: columns, data: data.map((row) => columns.map((c) => cellForExport(row[c]))) });
  downloadBlob(csv, `${filename}.csv`, "text/csv;charset=utf-8;");
}

function exportExcel(data, columns, filename) {
  const cleaned = data.map((row) => {
    const r = {};
    columns.forEach((c) => {
      r[c] = cellForExport(row[c]);
    });
    return r;
  });
  const ws = XLSX.utils.json_to_sheet(cleaned, { header: columns });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Report");
  XLSX.writeFile(wb, `${filename}.xlsx`);
}

function downloadBlob(content, filename, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ---------------- preset export / import ----------------

function exportPresets(presets) {
  const payload = {
    app: "report-builder-presets",
    version: 1,
    exportedAt: new Date().toISOString(),
    presets,
  };
  downloadBlob(JSON.stringify(payload, null, 2), `report-builder-presets-${dateStamp()}.json`, "application/json");
}

// Accepts either the wrapped export shape ({ presets: [...] }) or a bare
// array, and normalizes/defaults every field so older or hand-edited exports
// don't break. Always assigns fresh ids so an import can never collide with
// an existing preset.
// Fills in defaults for any preset field, normalizing older/partial shapes
// (e.g. legacy string groupBy, or fields that didn't exist yet when a preset
// was first saved) into the current one. Used both when loading from
// storage and when importing a backup file.
function normalizePresetFields(p) {
  return {
    id: typeof p.id === "string" && p.id ? p.id : uid(),
    name: typeof p.name === "string" ? p.name : "",
    filters: Array.isArray(p.filters) ? p.filters : [],
    filterMatch: p.filterMatch === "any" ? "any" : "all",
    columns: p.columns === "all" || Array.isArray(p.columns) ? p.columns : "all",
    groupBy: normalizeGroupBy(p.groupBy),
    aggregations: Array.isArray(p.aggregations) ? p.aggregations : [],
    sorts: Array.isArray(p.sorts) ? p.sorts : [],
    extraColumns: Array.isArray(p.extraColumns) ? p.extraColumns : [],
    duplicateColumn: typeof p.duplicateColumn === "string" ? p.duplicateColumn : "",
    sourceHeaders: Array.isArray(p.sourceHeaders) ? p.sourceHeaders : [],
  };
}

function parsePresetImport(text) {
  let json;
  try {
    json = JSON.parse(text);
  } catch (e) {
    return { error: "That doesn't look like a valid presets file (not valid JSON)." };
  }
  const list = Array.isArray(json) ? json : Array.isArray(json?.presets) ? json.presets : null;
  if (!list) {
    return { error: "That file doesn't contain any presets." };
  }
  const cleaned = [];
  for (const p of list) {
    if (!p || typeof p !== "object" || typeof p.name !== "string" || !p.name.trim()) continue;
    cleaned.push({ ...normalizePresetFields(p), id: uid() });
  }
  if (!cleaned.length) {
    return { error: "No usable presets were found in that file." };
  }
  return { presets: cleaned };
}

// ---------------- file comparison ----------------

function pick(row, cols) {
  const r = {};
  cols.forEach((c) => {
    r[c] = row[c];
  });
  return r;
}

function valuesEqual(a, b) {
  const an = parseFloat(a);
  const bn = parseFloat(b);
  const aIsNum = a !== null && a !== undefined && a !== "" && !isNaN(an);
  const bIsNum = b !== null && b !== undefined && b !== "" && !isNaN(bn);
  if (aIsNum && bIsNum) return an === bn;
  return String(a ?? "").trim() === String(b ?? "").trim();
}

// Matches rows between two datasets by a shared key column and classifies
// each as added, removed, changed (with per-field before/after), or
// unchanged. Refuses to run if the key column isn't unique in either file,
// since a non-unique key would silently produce misleading matches.
function compareDatasets(rowsA, headersA, rowsB, headersB, keyColumn) {
  const commonColumns = headersA.filter((h) => headersB.includes(h));

  const mapA = new Map();
  const dupA = new Set();
  rowsA.forEach((row) => {
    const key = String(row[keyColumn] ?? "").trim();
    if (!key) return;
    if (mapA.has(key)) dupA.add(key);
    mapA.set(key, row);
  });

  const mapB = new Map();
  const dupB = new Set();
  rowsB.forEach((row) => {
    const key = String(row[keyColumn] ?? "").trim();
    if (!key) return;
    if (mapB.has(key)) dupB.add(key);
    mapB.set(key, row);
  });

  if (dupA.size || dupB.size) {
    const which = dupA.size && dupB.size ? "both files" : dupA.size ? "the earlier file" : "the later file";
    return {
      error: `"${keyColumn}" has repeated values in ${which} \u2014 pick a column where every row is unique, like a case reference or ID.`,
    };
  }

  const added = [];
  const removed = [];
  const changed = [];
  let unchangedCount = 0;

  mapB.forEach((rowB, key) => {
    if (!mapA.has(key)) added.push(rowB);
  });

  mapA.forEach((rowA, key) => {
    if (!mapB.has(key)) {
      removed.push(rowA);
      return;
    }
    const rowB = mapB.get(key);
    const diffs = [];
    commonColumns.forEach((col) => {
      if (col === keyColumn) return;
      if (!valuesEqual(rowA[col], rowB[col])) {
        diffs.push({ column: col, from: rowA[col], to: rowB[col] });
      }
    });
    if (diffs.length) {
      changed.push({ key, row: rowB, diffs });
    } else {
      unchangedCount++;
    }
  });

  return { commonColumns, added, removed, changed, unchangedCount };
}

// ---------------- small shared components ----------------

function Stepper({ step }) {
  const steps = ["Upload", "Configure", "Report"];
  return (
    <div className="flex items-center gap-2 mb-6 flex-wrap">
      {steps.map((s, i) => (
        <div key={s} className="flex items-center gap-2">
          <div
            className="flex items-center gap-2 px-2 py-1 rounded"
            style={{
              color: i === step ? COLORS.surface : COLORS.muted,
              backgroundColor: i === step ? COLORS.accent : "transparent",
              fontFamily: FONTS.mono,
            }}
          >
            <span className="text-xs">{String(i + 1).padStart(2, "0")}</span>
            <span className="text-xs uppercase tracking-wide">{s}</span>
          </div>
          {i < steps.length - 1 && <ChevronRight size={14} style={{ color: COLORS.border }} />}
        </div>
      ))}
    </div>
  );
}

function Section({ title, hint, children }) {
  return (
    <div className="rounded-xl p-4" style={{ backgroundColor: COLORS.surface, border: `1px solid ${COLORS.border}` }}>
      <p className={hint ? "text-sm mb-0.5" : "text-sm mb-3"} style={{ fontFamily: FONTS.body, fontWeight: 600 }}>
        {title}
      </p>
      {hint && (
        <p className="text-xs mb-3" style={{ color: COLORS.muted }}>
          {hint}
        </p>
      )}
      {children}
    </div>
  );
}

function ModeTab({ active, onClick, label }) {
  return (
    <button
      onClick={onClick}
      className="px-3 py-2 rounded-lg text-sm"
      style={{
        backgroundColor: active ? COLORS.accent : COLORS.surface,
        color: active ? COLORS.surface : COLORS.muted,
        border: `1px solid ${active ? COLORS.accent : COLORS.border}`,
        fontFamily: FONTS.body,
        fontWeight: 500,
      }}
    >
      {label}
    </button>
  );
}

function SheetPicker({ file, onSwitchSheet }) {
  if (!file.sheetNames || file.sheetNames.length <= 1) return null;
  return (
    <div className="flex items-center gap-2 mt-2">
      <span className="text-xs shrink-0" style={{ color: COLORS.muted }}>
        Sheet
      </span>
      <select
        value={file.selectedSheet}
        onChange={(e) => onSwitchSheet(e.target.value)}
        className="px-2 py-1.5 rounded-lg text-xs flex-1"
        style={{ border: `1px solid ${COLORS.border}`, backgroundColor: COLORS.bg }}
      >
        {file.sheetNames.map((s) => (
          <option key={s} value={s}>
            {s}
          </option>
        ))}
      </select>
    </div>
  );
}

function PresetCard({ preset, hasFile, onRun, onEdit, onDelete }) {
  const parts = summarize(preset);
  const [confirmDelete, setConfirmDelete] = useState(false);

  return (
    <div className="rounded-xl p-4" style={{ backgroundColor: COLORS.surface, border: `1px solid ${COLORS.border}` }}>
      <div className="flex items-start justify-between gap-3 mb-2.5">
        <p style={{ fontFamily: FONTS.display, fontWeight: 600, fontSize: "1.05rem" }}>{preset.name}</p>
        <div className="flex items-center gap-1 shrink-0">
          <button onClick={onEdit} className="p-1.5 rounded" style={{ color: COLORS.muted }} aria-label="Edit preset">
            <Pencil size={14} />
          </button>
          {confirmDelete ? (
            <button
              onClick={onDelete}
              className="text-xs px-2 py-1 rounded"
              style={{ backgroundColor: COLORS.dangerSoft, color: COLORS.danger }}
            >
              Confirm
            </button>
          ) : (
            <button
              onClick={() => setConfirmDelete(true)}
              onBlur={() => setConfirmDelete(false)}
              className="p-1.5 rounded"
              style={{ color: COLORS.muted }}
              aria-label="Delete preset"
            >
              <Trash2 size={14} />
            </button>
          )}
        </div>
      </div>
      <div className="flex flex-wrap items-center gap-x-2 gap-y-1.5 mb-3 text-sm">
        {parts.length === 0 && <span style={{ color: COLORS.muted }}>No steps configured yet</span>}
        {parts.map((part, i) => (
          <span key={i} className="inline-flex items-center gap-2">
            {i > 0 && <ArrowRight size={12} style={{ color: COLORS.border }} />}
            <span>
              <span style={{ color: COLORS.muted }}>{part.label}: </span>
              <span style={{ fontFamily: FONTS.mono, color: COLORS.ink }}>{part.value || "\u2014"}</span>
            </span>
          </span>
        ))}
      </div>
      <button
        onClick={onRun}
        disabled={!hasFile}
        className="text-sm px-3 py-1.5 rounded-lg disabled:opacity-40"
        style={{ backgroundColor: COLORS.accentSoft, color: COLORS.accent, fontFamily: FONTS.body, fontWeight: 500 }}
      >
        {hasFile ? "Apply to this file" : "Upload a file to apply"}
      </button>
    </div>
  );
}

function HomeView({
  file,
  parsing,
  parseError,
  presets,
  presetsReady,
  fileInputRef,
  onFile,
  onSwitchSheet,
  onNewPreset,
  onEditPreset,
  onDeletePreset,
  onRunPreset,
  onClearFile,
  onExportPresets,
  onImportPresets,
}) {
  const [dragOver, setDragOver] = useState(false);
  const importInputRef = useRef(null);

  return (
    <div className="space-y-6">
      {!file && (
        <div
          onDragOver={(e) => {
            e.preventDefault();
            setDragOver(true);
          }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => {
            e.preventDefault();
            setDragOver(false);
            const f = e.dataTransfer.files?.[0];
            if (f) onFile(f);
          }}
          className="rounded-2xl border-2 border-dashed p-10 text-center transition-colors"
          style={{
            borderColor: dragOver ? COLORS.accent : COLORS.border,
            backgroundColor: dragOver ? COLORS.accentSoft : COLORS.surface,
          }}
        >
          <FileUp size={28} style={{ color: COLORS.muted, margin: "0 auto 12px" }} />
          <p className="mb-1" style={{ fontFamily: FONTS.body, fontWeight: 500 }}>
            Drop a CSV or Excel file here
          </p>
          <p className="text-sm mb-4" style={{ color: COLORS.muted }}>
            or choose a file from your device
          </p>
          <button
            onClick={() => fileInputRef.current?.click()}
            className="px-4 py-2 rounded-lg text-sm"
            style={{ backgroundColor: COLORS.accent, color: COLORS.surface, fontFamily: FONTS.body, fontWeight: 500 }}
          >
            Choose file
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept=".csv,.xlsx,.xls"
            className="hidden"
            onChange={(e) => onFile(e.target.files?.[0])}
          />
          {parsing && (
            <p className="text-sm mt-4" style={{ color: COLORS.muted }}>
              Reading file\u2026
            </p>
          )}
          {parseError && (
            <p className="text-sm mt-4 flex items-center justify-center gap-1.5" style={{ color: COLORS.danger }}>
              <AlertTriangle size={14} /> {parseError}
            </p>
          )}
        </div>
      )}

      {file && (
        <div className="rounded-2xl p-5" style={{ backgroundColor: COLORS.surface, border: `1px solid ${COLORS.border}` }}>
          <div className="flex items-start justify-between gap-3 mb-3">
            <div className="flex items-center gap-2.5 min-w-0">
              <FileSpreadsheet size={20} style={{ color: COLORS.accent, flexShrink: 0 }} />
              <div className="min-w-0">
                <p className="text-sm font-medium truncate" style={{ fontFamily: FONTS.body }}>
                  {file.name}
                </p>
                <p className="text-xs" style={{ color: COLORS.muted }}>
                  {file.rows.length.toLocaleString()} rows \u00b7 {file.headers.length} columns
                </p>
              </div>
            </div>
            <button onClick={onClearFile} className="text-xs shrink-0" style={{ color: COLORS.muted }}>
              Change file
            </button>
          </div>
          <SheetPicker file={file} onSwitchSheet={onSwitchSheet} />
          <div className="flex flex-wrap gap-1.5 mt-3">
            {file.headers.map((h) => (
              <span
                key={h}
                className="text-xs px-2 py-1 rounded"
                style={{ backgroundColor: COLORS.bg, color: COLORS.muted, fontFamily: FONTS.mono }}
              >
                {h}
              </span>
            ))}
          </div>
        </div>
      )}

      <div>
        <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
          <h2 className="text-sm uppercase tracking-wide" style={{ color: COLORS.muted, fontFamily: FONTS.mono }}>
            Saved presets
          </h2>
          <div className="flex items-center gap-1">
            <button
              onClick={onExportPresets}
              disabled={!presets.length}
              className="p-1.5 rounded disabled:opacity-30"
              style={{ color: COLORS.muted }}
              aria-label="Export presets"
              title="Export presets to a file"
            >
              <Download size={15} />
            </button>
            <button
              onClick={() => importInputRef.current?.click()}
              className="p-1.5 rounded"
              style={{ color: COLORS.muted }}
              aria-label="Import presets"
              title="Import presets from a file"
            >
              <Upload size={15} />
            </button>
            <input
              ref={importInputRef}
              type="file"
              accept=".json"
              className="hidden"
              onChange={(e) => {
                onImportPresets(e.target.files?.[0]);
                e.target.value = "";
              }}
            />
            <button
              onClick={onNewPreset}
              disabled={!file}
              className="flex items-center gap-1 text-sm px-3 py-1.5 rounded-lg disabled:opacity-40"
              style={{ color: COLORS.accent, fontFamily: FONTS.body, fontWeight: 500 }}
            >
              <Plus size={15} /> New preset
            </button>
          </div>
        </div>

        {!presetsReady ? null : presets.length === 0 ? (
          <div
            className="rounded-xl p-6 text-center text-sm"
            style={{ border: `1px dashed ${COLORS.border}`, color: COLORS.muted }}
          >
            {file
              ? "No presets yet \u2014 create one from this file to reuse it next time."
              : "No presets yet. Upload a file to build your first one, or import a backup above."}
          </div>
        ) : (
          <div className="space-y-3">
            {presets.map((preset) => (
              <PresetCard
                key={preset.id}
                preset={preset}
                hasFile={!!file}
                onRun={() => onRunPreset(preset)}
                onEdit={() => onEditPreset(preset)}
                onDelete={() => onDeletePreset(preset.id)}
              />
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function BuilderView({ preset, setPreset, headers, hasFile, onCancel, onSave }) {
  function update(patch) {
    setPreset({ ...preset, ...patch });
  }

  function addFilter() {
    update({ filters: [...preset.filters, { column: headers[0] || "", operator: "equals", value: "", value2: "" }] });
  }
  function updateFilter(i, patch) {
    const next = preset.filters.slice();
    next[i] = { ...next[i], ...patch };
    update({ filters: next });
  }
  function removeFilter(i) {
    update({ filters: preset.filters.filter((_, idx) => idx !== i) });
  }

  function addAgg() {
    update({ aggregations: [...preset.aggregations, { column: headers[0] || "", op: "sum" }] });
  }
  function updateAgg(i, patch) {
    const next = preset.aggregations.slice();
    next[i] = { ...next[i], ...patch };
    update({ aggregations: next });
  }
  function removeAgg(i) {
    update({ aggregations: preset.aggregations.filter((_, idx) => idx !== i) });
  }

  function toggleColumn(col) {
    const current = preset.columns === "all" ? headers.slice() : preset.columns.slice();
    const next = current.includes(col) ? current.filter((c) => c !== col) : [...current, col];
    update({ columns: next });
  }

  function addSort() {
    update({ sorts: [...(preset.sorts || []), { column: sortOptions[0] || "", direction: "asc" }] });
  }
  function updateSort(i, patch) {
    const next = (preset.sorts || []).slice();
    next[i] = { ...next[i], ...patch };
    update({ sorts: next });
  }
  function removeSort(i) {
    update({ sorts: (preset.sorts || []).filter((_, idx) => idx !== i) });
  }

  function addGroupBy() {
    const used = new Set(groupBy);
    const next = headers.find((h) => !used.has(h)) || headers[0] || "";
    update({ groupBy: [...groupBy, next] });
  }
  function updateGroupByAt(i, value) {
    const next = groupBy.slice();
    next[i] = value;
    update({ groupBy: next });
  }
  function removeGroupBy(i) {
    update({ groupBy: groupBy.filter((_, idx) => idx !== i) });
  }

  function addExtraColumn() {
    update({ extraColumns: [...(preset.extraColumns || []), ""] });
  }
  function updateExtraColumn(i, value) {
    const next = (preset.extraColumns || []).slice();
    next[i] = value;
    update({ extraColumns: next });
  }
  function removeExtraColumn(i) {
    update({ extraColumns: (preset.extraColumns || []).filter((_, idx) => idx !== i) });
  }

  const groupBy = normalizeGroupBy(preset.groupBy);
  const selectedColumns = preset.columns === "all" ? headers : preset.columns;
  const sortOptions = groupBy.length
    ? [...groupBy, "Count", ...preset.aggregations.filter((a) => a.column).map((a) => `${a.op}_${a.column}`)]
    : headers;

  // If a sort level stops being valid — e.g. its calculation was removed or
  // renamed, or grouping was toggled — drop just that level instead of
  // silently pointing at whatever option a native <select> falls back to.
  useEffect(() => {
    const current = preset.sorts || [];
    const valid = current.filter((s) => !s.column || sortOptions.includes(s.column));
    if (valid.length !== current.length) {
      update({ sorts: valid });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [preset.groupBy, preset.aggregations]);

  return (
    <div className="space-y-4 pb-6">
      {headers.length === 0 && (
        <div
          className="rounded-xl p-4 text-sm flex items-start gap-2"
          style={{ backgroundColor: COLORS.dangerSoft, color: COLORS.danger }}
        >
          <AlertTriangle size={16} style={{ marginTop: "2px", flexShrink: 0 }} />
          <span>
            No columns to work with yet \u2014 upload the file this preset was built for once. After that you can
            come back and edit it without a file.
          </span>
        </div>
      )}
      <div>
        <label className="text-xs uppercase tracking-wide block mb-1.5" style={{ color: COLORS.muted, fontFamily: FONTS.mono }}>
          Preset name
        </label>
        <input
          value={preset.name}
          onChange={(e) => update({ name: e.target.value })}
          placeholder='e.g. "Cases by stage"'
          className="w-full px-3 py-2.5 rounded-lg text-base"
          style={{ border: `1px solid ${COLORS.border}`, backgroundColor: COLORS.surface, fontFamily: FONTS.body }}
        />
      </div>

      <Section title="Filter rows" hint="Add one or more conditions, then choose how they combine below">
        {preset.filters.length >= 2 && (
          <div className="flex items-center gap-2 mb-3">
            <span className="text-sm" style={{ color: COLORS.muted }}>
              Match
            </span>
            <button
              onClick={() => update({ filterMatch: "all" })}
              className="text-xs px-2.5 py-1.5 rounded-full"
              style={{
                backgroundColor: preset.filterMatch !== "any" ? COLORS.accentSoft : COLORS.bg,
                color: preset.filterMatch !== "any" ? COLORS.accent : COLORS.muted,
                fontFamily: FONTS.body,
                fontWeight: 500,
              }}
            >
              All filters
            </button>
            <button
              onClick={() => update({ filterMatch: "any" })}
              className="text-xs px-2.5 py-1.5 rounded-full"
              style={{
                backgroundColor: preset.filterMatch === "any" ? COLORS.accentSoft : COLORS.bg,
                color: preset.filterMatch === "any" ? COLORS.accent : COLORS.muted,
                fontFamily: FONTS.body,
                fontWeight: 500,
              }}
            >
              Any filter
            </button>
          </div>
        )}
        {preset.filters.map((f, i) => (
          <div key={i} className="flex flex-wrap items-center gap-2 mb-2">
            <select
              value={f.column}
              onChange={(e) => updateFilter(i, { column: e.target.value })}
              className="px-2 py-2 rounded-lg text-sm flex-1"
              style={{ border: `1px solid ${COLORS.border}`, backgroundColor: COLORS.surface, minWidth: "110px" }}
            >
              {headers.map((h) => (
                <option key={h} value={h}>
                  {h}
                </option>
              ))}
            </select>
            <select
              value={f.operator}
              onChange={(e) => updateFilter(i, { operator: e.target.value })}
              className="px-2 py-2 rounded-lg text-sm"
              style={{ border: `1px solid ${COLORS.border}`, backgroundColor: COLORS.surface }}
            >
              {OPERATORS.map((o) => (
                <option key={o.value} value={o.value}>
                  {o.label}
                </option>
              ))}
            </select>
            {f.operator === "between_dates" ? (
              <div className="flex items-center gap-2 flex-1" style={{ minWidth: "220px" }}>
                <input
                  type="date"
                  value={f.value}
                  onChange={(e) => updateFilter(i, { value: e.target.value })}
                  className="px-2 py-2 rounded-lg text-sm flex-1"
                  style={{ border: `1px solid ${COLORS.border}`, backgroundColor: COLORS.surface, minWidth: "130px" }}
                />
                <span className="text-sm shrink-0" style={{ color: COLORS.muted }}>
                  and
                </span>
                <input
                  type="date"
                  value={f.value2 || ""}
                  onChange={(e) => updateFilter(i, { value2: e.target.value })}
                  className="px-2 py-2 rounded-lg text-sm flex-1"
                  style={{ border: `1px solid ${COLORS.border}`, backgroundColor: COLORS.surface, minWidth: "130px" }}
                />
              </div>
            ) : f.operator === "is_blank" || f.operator === "is_not_blank" ? null : (
              <input
                value={f.value}
                onChange={(e) => updateFilter(i, { value: e.target.value })}
                placeholder="value"
                className="px-2 py-2 rounded-lg text-sm flex-1"
                style={{ border: `1px solid ${COLORS.border}`, backgroundColor: COLORS.surface, minWidth: "90px" }}
              />
            )}
            <button
              onClick={() => removeFilter(i)}
              className="p-2 rounded-lg shrink-0"
              style={{ color: COLORS.danger }}
              aria-label="Remove filter"
            >
              <X size={16} />
            </button>
          </div>
        ))}
        <button
          onClick={addFilter}
          className="flex items-center gap-1.5 text-sm mt-1"
          style={{ color: COLORS.accent, fontFamily: FONTS.body, fontWeight: 500 }}
        >
          <Plus size={14} /> Add filter
        </button>
      </Section>

      <Section title="Group & calculate" hint="Add one or more columns to roll rows up into totals \u2014 leave empty to keep row-by-row detail">
        {groupBy.map((col, i) => (
          <div key={i} className="flex flex-wrap items-center gap-2 mb-2">
            <span className="text-sm shrink-0" style={{ color: COLORS.muted, minWidth: "60px" }}>
              {i === 0 ? "Group by" : "Then by"}
            </span>
            <select
              value={col}
              onChange={(e) => updateGroupByAt(i, e.target.value)}
              className="px-2 py-2 rounded-lg text-sm flex-1"
              style={{ border: `1px solid ${COLORS.border}`, backgroundColor: COLORS.surface, minWidth: "140px" }}
            >
              {headers.map((h) => (
                <option key={h} value={h}>
                  {h}
                </option>
              ))}
            </select>
            <button
              onClick={() => removeGroupBy(i)}
              className="p-2 rounded-lg shrink-0"
              style={{ color: COLORS.danger }}
              aria-label="Remove grouping level"
            >
              <X size={16} />
            </button>
          </div>
        ))}
        <button
          onClick={addGroupBy}
          disabled={groupBy.length >= headers.length}
          className="flex items-center gap-1.5 text-sm mt-1 disabled:opacity-40"
          style={{ color: COLORS.accent, fontFamily: FONTS.body, fontWeight: 500 }}
        >
          <Plus size={14} /> {groupBy.length ? "Add another level" : "Group by a column"}
        </button>

        {groupBy.length > 0 && (
          <>
            <div className="mt-4">
              {preset.aggregations.map((a, i) => (
                <div key={i} className="flex flex-wrap items-center gap-2 mb-2">
                  <select
                    value={a.op}
                    onChange={(e) => updateAgg(i, { op: e.target.value })}
                    className="px-2 py-2 rounded-lg text-sm"
                    style={{ border: `1px solid ${COLORS.border}`, backgroundColor: COLORS.surface }}
                  >
                    {AGG_OPS.map((o) => (
                      <option key={o.value} value={o.value}>
                        {o.label}
                      </option>
                    ))}
                  </select>
                  <span className="text-sm" style={{ color: COLORS.muted }}>
                    of
                  </span>
                  <select
                    value={a.column}
                    onChange={(e) => updateAgg(i, { column: e.target.value })}
                    className="px-2 py-2 rounded-lg text-sm flex-1"
                    style={{ border: `1px solid ${COLORS.border}`, backgroundColor: COLORS.surface, minWidth: "110px" }}
                  >
                    {headers.map((h) => (
                      <option key={h} value={h}>
                        {h}
                      </option>
                    ))}
                  </select>
                  <button
                    onClick={() => removeAgg(i)}
                    className="p-2 rounded-lg shrink-0"
                    style={{ color: COLORS.danger }}
                    aria-label="Remove calculation"
                  >
                    <X size={16} />
                  </button>
                </div>
              ))}
              <button
                onClick={addAgg}
                className="flex items-center gap-1.5 text-sm mt-1"
                style={{ color: COLORS.accent, fontFamily: FONTS.body, fontWeight: 500 }}
              >
                <Plus size={14} /> Add calculation
              </button>
            </div>
          </>
        )}
      </Section>

      {!groupBy.length && (
        <Section title="Columns to include" hint="Leave all selected to keep every column">
          <div className="flex flex-wrap gap-1.5">
            {headers.map((h) => {
              const active = selectedColumns.includes(h);
              return (
                <button
                  key={h}
                  onClick={() => toggleColumn(h)}
                  className="text-xs px-2.5 py-1.5 rounded-full flex items-center gap-1"
                  style={{
                    backgroundColor: active ? COLORS.accentSoft : COLORS.bg,
                    color: active ? COLORS.accent : COLORS.muted,
                    fontFamily: FONTS.mono,
                  }}
                >
                  {active && <Check size={11} />}
                  {h}
                </button>
              );
            })}
          </div>
        </Section>
      )}

      {!groupBy.length && (
        <Section title="Highlight duplicates" hint="Optional \u2014 flags rows where this column repeats a value, in yellow">
          <select
            value={preset.duplicateColumn || ""}
            onChange={(e) => update({ duplicateColumn: e.target.value })}
            className="px-2 py-2 rounded-lg text-sm w-full"
            style={{ border: `1px solid ${COLORS.border}`, backgroundColor: COLORS.surface }}
          >
            <option value="">Don't highlight duplicates</option>
            {headers.map((h) => (
              <option key={h} value={h}>
                {h}
              </option>
            ))}
          </select>
        </Section>
      )}

      <Section title="Add columns" hint='New blank columns added to the end of the report — handy for things like "Notes" to fill in by hand after exporting'>
        {(preset.extraColumns || []).map((col, i) => (
          <div key={i} className="flex items-center gap-2 mb-2">
            <input
              value={col}
              onChange={(e) => updateExtraColumn(i, e.target.value)}
              placeholder='Column heading, e.g. "Notes"'
              className="px-2 py-2 rounded-lg text-sm flex-1"
              style={{ border: `1px solid ${COLORS.border}`, backgroundColor: COLORS.surface }}
            />
            <button
              onClick={() => removeExtraColumn(i)}
              className="p-2 rounded-lg shrink-0"
              style={{ color: COLORS.danger }}
              aria-label="Remove column"
            >
              <X size={16} />
            </button>
          </div>
        ))}
        <button
          onClick={addExtraColumn}
          className="flex items-center gap-1.5 text-sm mt-1"
          style={{ color: COLORS.accent, fontFamily: FONTS.body, fontWeight: 500 }}
        >
          <Plus size={14} /> Add column
        </button>
      </Section>

      <Section title="Sort" hint="Add another level to break ties in the one above it — just like Excel's Sort by / Then by">
        {(preset.sorts || []).map((s, i) => (
          <div key={i} className="flex flex-wrap items-center gap-2 mb-2">
            <span className="text-sm shrink-0" style={{ color: COLORS.muted, minWidth: "60px" }}>
              {i === 0 ? "Sort by" : "Then by"}
            </span>
            <select
              value={s.column}
              onChange={(e) => updateSort(i, { column: e.target.value })}
              className="px-2 py-2 rounded-lg text-sm flex-1"
              style={{ border: `1px solid ${COLORS.border}`, backgroundColor: COLORS.surface, minWidth: "120px" }}
            >
              {sortOptions.map((h) => (
                <option key={h} value={h}>
                  {h}
                </option>
              ))}
            </select>
            <select
              value={s.direction}
              onChange={(e) => updateSort(i, { direction: e.target.value })}
              className="px-2 py-2 rounded-lg text-sm"
              style={{ border: `1px solid ${COLORS.border}`, backgroundColor: COLORS.surface }}
            >
              <option value="asc">Low to high</option>
              <option value="desc">High to low</option>
            </select>
            <button
              onClick={() => removeSort(i)}
              className="p-2 rounded-lg shrink-0"
              style={{ color: COLORS.danger }}
              aria-label="Remove sort level"
            >
              <X size={16} />
            </button>
          </div>
        ))}
        <button
          onClick={addSort}
          disabled={!sortOptions.length}
          className="flex items-center gap-1.5 text-sm mt-1 disabled:opacity-40"
          style={{ color: COLORS.accent, fontFamily: FONTS.body, fontWeight: 500 }}
        >
          <Plus size={14} /> {preset.sorts && preset.sorts.length ? "Add another level" : "Add sort level"}
        </button>
      </Section>

      <div className="flex items-center gap-3 pt-2">
        <button
          onClick={() => onSave(preset)}
          className="px-4 py-2.5 rounded-lg text-sm"
          style={{ backgroundColor: COLORS.accent, color: COLORS.surface, fontFamily: FONTS.body, fontWeight: 500 }}
        >
          {hasFile ? "Save & view report" : "Save preset"}
        </button>
        <button onClick={onCancel} className="px-4 py-2.5 text-sm" style={{ color: COLORS.muted, fontFamily: FONTS.body }}>
          Cancel
        </button>
      </div>
    </div>
  );
}

function ResultsView({ result, presetName, fileName, onBack, onReset }) {
  const { data, columns, numericColumns, duplicateColumn } = result;
  const baseName = (fileName || "report").replace(/\.[^.]+$/, "");

  function exportRows(fmt) {
    if (!duplicateColumn) {
      if (fmt === "csv") exportCSV(data, columns, `${baseName}-${presetName}`);
      else exportExcel(data, columns, `${baseName}-${presetName}`);
      return;
    }
    // Colour doesn't carry over to CSV/Excel, so add a plain "Duplicate"
    // column instead, keeping the information available after export.
    const exportColumns = [...columns, "Duplicate"];
    const exportData = data.map((row) => ({ ...row, Duplicate: row.__dup ? "Yes" : "" }));
    if (fmt === "csv") exportCSV(exportData, exportColumns, `${baseName}-${presetName}`);
    else exportExcel(exportData, exportColumns, `${baseName}-${presetName}`);
  }

  return (
    <div className="space-y-4">
      <div className="flex items-start justify-between gap-3 flex-wrap">
        <div className="min-w-0">
          <p style={{ fontFamily: FONTS.display, fontWeight: 600, fontSize: "1.25rem" }} className="truncate">
            {presetName}
          </p>
          <p className="text-xs" style={{ color: COLORS.muted }}>
            {data.length.toLocaleString()} {data.length === 1 ? "row" : "rows"}
          </p>
        </div>
        <div className="flex items-center gap-3 shrink-0">
          <button onClick={onBack} className="flex items-center gap-1 text-sm" style={{ color: COLORS.muted }}>
            <ArrowLeft size={14} /> Back
          </button>
          <button onClick={onReset} className="flex items-center gap-1 text-sm" style={{ color: COLORS.accent }}>
            <RotateCcw size={14} /> Start over
          </button>
        </div>
      </div>

      {duplicateColumn && (
        <div className="flex items-center gap-2 text-xs" style={{ color: COLORS.muted }}>
          <span className="inline-block rounded" style={{ width: "12px", height: "12px", backgroundColor: COLORS.duplicate }} />
          Rows with a repeated value in <span style={{ fontFamily: FONTS.mono, color: COLORS.ink }}>{duplicateColumn}</span> are
          highlighted
        </div>
      )}

      <div className="rounded-xl overflow-hidden" style={{ border: `1px solid ${COLORS.border}` }}>
        <div className="overflow-x-auto">
          <table className="w-full text-sm" style={{ borderCollapse: "collapse" }}>
            <thead>
              <tr style={{ backgroundColor: COLORS.bg }}>
                {columns.map((c) => (
                  <th
                    key={c}
                    className="px-3 py-2.5 whitespace-nowrap"
                    style={{
                      fontFamily: FONTS.mono,
                      fontWeight: 500,
                      color: COLORS.muted,
                      textAlign: numericColumns?.has(c) ? "right" : "left",
                    }}
                  >
                    {c.replace(/_/g, " ")}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {data.length === 0 ? (
                <tr>
                  <td colSpan={columns.length} className="px-3 py-6 text-center text-sm" style={{ color: COLORS.muted }}>
                    No rows match this preset for this file.
                  </td>
                </tr>
              ) : (
                data.map((row, i) => (
                  <tr
                    key={i}
                    style={{
                      borderTop: `1px solid ${COLORS.border}`,
                      backgroundColor: row.__dup ? COLORS.duplicate : undefined,
                    }}
                  >
                    {columns.map((c) => (
                      <td
                        key={c}
                        className="px-3 py-2 whitespace-nowrap"
                        style={{
                          fontFamily: numericColumns?.has(c) ? FONTS.mono : FONTS.body,
                          textAlign: numericColumns?.has(c) ? "right" : "left",
                        }}
                      >
                        {formatCell(row[c], numericColumns?.has(c))}
                      </td>
                    ))}
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      <div className="flex flex-wrap items-center gap-2">
        <button
          onClick={() => exportRows("csv")}
          className="flex items-center gap-1.5 text-sm px-3 py-2 rounded-lg"
          style={{ backgroundColor: COLORS.accentSoft, color: COLORS.accent, fontFamily: FONTS.body, fontWeight: 500 }}
        >
          <Download size={14} /> Export CSV
        </button>
        <button
          onClick={() => exportRows("excel")}
          className="flex items-center gap-1.5 text-sm px-3 py-2 rounded-lg"
          style={{ backgroundColor: COLORS.accent2Soft, color: COLORS.accent2, fontFamily: FONTS.body, fontWeight: 500 }}
        >
          <Download size={14} /> Export Excel
        </button>
      </div>
    </div>
  );
}

// ---------------- compare mode ----------------

function CompareUploadSlot({ label, file, parsing, error, inputRef, onFile, onSwitchSheet, onClear }) {
  const [dragOver, setDragOver] = useState(false);
  return (
    <div>
      <p className="text-xs uppercase tracking-wide mb-1.5" style={{ color: COLORS.muted, fontFamily: FONTS.mono }}>
        {label}
      </p>
      {!file ? (
        <div
          onDragOver={(e) => {
            e.preventDefault();
            setDragOver(true);
          }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => {
            e.preventDefault();
            setDragOver(false);
            const f = e.dataTransfer.files?.[0];
            if (f) onFile(f);
          }}
          className="rounded-xl border-2 border-dashed p-6 text-center"
          style={{
            borderColor: dragOver ? COLORS.accent : COLORS.border,
            backgroundColor: dragOver ? COLORS.accentSoft : COLORS.surface,
          }}
        >
          <button
            onClick={() => inputRef.current?.click()}
            className="px-3 py-1.5 rounded-lg text-sm"
            style={{ backgroundColor: COLORS.accent, color: COLORS.surface, fontFamily: FONTS.body, fontWeight: 500 }}
          >
            Choose file
          </button>
          <p className="text-xs mt-2" style={{ color: COLORS.muted }}>
            or drop a CSV/Excel file here
          </p>
          <input
            ref={inputRef}
            type="file"
            accept=".csv,.xlsx,.xls"
            className="hidden"
            onChange={(e) => onFile(e.target.files?.[0])}
          />
          {parsing && (
            <p className="text-xs mt-2" style={{ color: COLORS.muted }}>
              Reading file\u2026
            </p>
          )}
          {error && (
            <p className="text-xs mt-2 flex items-center justify-center gap-1" style={{ color: COLORS.danger }}>
              <AlertTriangle size={12} /> {error}
            </p>
          )}
        </div>
      ) : (
        <div className="rounded-xl p-4" style={{ backgroundColor: COLORS.surface, border: `1px solid ${COLORS.border}` }}>
          <div className="flex items-start justify-between gap-2">
            <div className="flex items-center gap-2 min-w-0">
              <FileSpreadsheet size={16} style={{ color: COLORS.accent, flexShrink: 0 }} />
              <div className="min-w-0">
                <p className="text-sm font-medium truncate">{file.name}</p>
                <p className="text-xs" style={{ color: COLORS.muted }}>
                  {file.rows.length.toLocaleString()} rows \u00b7 {file.headers.length} columns
                </p>
              </div>
            </div>
            <button onClick={onClear} className="text-xs shrink-0" style={{ color: COLORS.muted }}>
              Change
            </button>
          </div>
          <SheetPicker file={file} onSwitchSheet={onSwitchSheet} />
        </div>
      )}
    </div>
  );
}

function SummaryPill({ label, count, color, bg }) {
  return (
    <div className="px-3 py-2 rounded-lg" style={{ backgroundColor: bg, minWidth: "84px" }}>
      <p style={{ fontFamily: FONTS.mono, fontWeight: 600, color, fontSize: "1.15rem" }}>{count}</p>
      <p className="text-xs" style={{ color: COLORS.muted }}>
        {label}
      </p>
    </div>
  );
}

function DiffTable({ title, rows, columns, tone }) {
  return (
    <div>
      <p className="text-sm mb-2" style={{ fontWeight: 600, color: tone }}>
        {title} ({rows.length})
      </p>
      <div className="rounded-xl overflow-hidden" style={{ border: `1px solid ${COLORS.border}` }}>
        <div className="overflow-x-auto">
          <table className="w-full text-sm" style={{ borderCollapse: "collapse" }}>
            <thead>
              <tr style={{ backgroundColor: COLORS.bg }}>
                {columns.map((c) => (
                  <th
                    key={c}
                    className="text-left px-3 py-2 whitespace-nowrap"
                    style={{ fontFamily: FONTS.mono, fontWeight: 500, color: COLORS.muted }}
                  >
                    {c}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((row, i) => (
                <tr key={i} style={{ borderTop: `1px solid ${COLORS.border}` }}>
                  {columns.map((c) => (
                    <td key={c} className="px-3 py-2 whitespace-nowrap">
                      {formatCell(row[c], false)}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function CompareResults({ result, keyColumn, fileAName, fileBName, onReset, onBack }) {
  const { added, removed, changed, unchangedCount, commonColumns } = result;
  const totalChanges = added.length + removed.length + changed.length;

  function handleExport() {
    if (!totalChanges) return;
    const cols = ["Change", ...commonColumns, "Details"];
    const rows = [
      ...added.map((r) => ({ Change: "New", ...pick(r, commonColumns), Details: "" })),
      ...removed.map((r) => ({ Change: "Removed", ...pick(r, commonColumns), Details: "" })),
      ...changed.map((c) => ({
        Change: "Changed",
        ...pick(c.row, commonColumns),
        Details: c.diffs.map((d) => `${d.column}: ${formatCell(d.from, false)} -> ${formatCell(d.to, false)}`).join("; "),
      })),
    ];
    exportCSV(rows, cols, `comparison-${dateStamp()}`);
  }

  return (
    <div className="space-y-5">
      <div className="flex items-start justify-between gap-3 flex-wrap">
        <div className="min-w-0">
          <p style={{ fontFamily: FONTS.display, fontWeight: 600, fontSize: "1.25rem" }}>Comparison</p>
          <p className="text-xs truncate" style={{ color: COLORS.muted }}>
            {fileAName} \u2192 {fileBName} \u00b7 matched by {keyColumn}
          </p>
        </div>
        <div className="flex items-center gap-3 shrink-0">
          <button onClick={onBack} className="flex items-center gap-1 text-sm" style={{ color: COLORS.muted }}>
            <ArrowLeft size={14} /> Back
          </button>
          <button onClick={onReset} className="flex items-center gap-1 text-sm" style={{ color: COLORS.accent }}>
            <RotateCcw size={14} /> Start over
          </button>
        </div>
      </div>

      <div className="flex flex-wrap gap-2">
        <SummaryPill label="New" count={added.length} color={COLORS.accent} bg={COLORS.accentSoft} />
        <SummaryPill label="Removed" count={removed.length} color={COLORS.danger} bg={COLORS.dangerSoft} />
        <SummaryPill label="Changed" count={changed.length} color={COLORS.accent2} bg={COLORS.accent2Soft} />
        <SummaryPill label="Unchanged" count={unchangedCount} color={COLORS.muted} bg={COLORS.bg} />
      </div>

      {added.length > 0 && <DiffTable title="New rows" rows={added} columns={commonColumns} tone={COLORS.accent} />}
      {removed.length > 0 && <DiffTable title="Removed rows" rows={removed} columns={commonColumns} tone={COLORS.danger} />}

      {changed.length > 0 && (
        <div>
          <p className="text-sm mb-2" style={{ fontWeight: 600, color: COLORS.accent2 }}>
            Changed rows ({changed.length})
          </p>
          <div className="space-y-2">
            {changed.map((c, i) => (
              <div key={i} className="rounded-xl p-3" style={{ backgroundColor: COLORS.surface, border: `1px solid ${COLORS.border}` }}>
                <p className="text-sm mb-1.5" style={{ fontFamily: FONTS.mono, fontWeight: 500 }}>
                  {c.key}
                </p>
                <div className="flex flex-col gap-1">
                  {c.diffs.map((d, j) => (
                    <p key={j} className="text-xs">
                      <span style={{ color: COLORS.ink }}>{d.column}</span>
                      <span style={{ color: COLORS.muted }}>: {formatCell(d.from, false)} </span>
                      <span style={{ color: COLORS.accent2 }}>\u2192</span>
                      <span style={{ color: COLORS.muted }}> {formatCell(d.to, false)}</span>
                    </p>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {totalChanges === 0 && (
        <div className="rounded-xl p-6 text-center text-sm" style={{ border: `1px dashed ${COLORS.border}`, color: COLORS.muted }}>
          No differences found \u2014 every matched row is identical.
        </div>
      )}

      <button
        onClick={handleExport}
        disabled={!totalChanges}
        className="flex items-center gap-1.5 text-sm px-3 py-2 rounded-lg disabled:opacity-40"
        style={{ backgroundColor: COLORS.accentSoft, color: COLORS.accent, fontFamily: FONTS.body, fontWeight: 500 }}
      >
        <Download size={14} /> Export changes (CSV)
      </button>
    </div>
  );
}

function CompareView() {
  const [fileA, setFileA] = useState(null);
  const [fileB, setFileB] = useState(null);
  const [parsingA, setParsingA] = useState(false);
  const [parsingB, setParsingB] = useState(false);
  const [errorA, setErrorA] = useState("");
  const [errorB, setErrorB] = useState("");
  const [keyColumn, setKeyColumn] = useState("");
  const [result, setResult] = useState(null);
  const [resultError, setResultError] = useState("");
  const inputARef = useRef(null);
  const inputBRef = useRef(null);

  const commonHeaders = fileA && fileB ? fileA.headers.filter((h) => fileB.headers.includes(h)) : [];

  useEffect(() => {
    if (keyColumn && !commonHeaders.includes(keyColumn)) {
      setKeyColumn("");
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fileA, fileB]);

  async function handleFile(which, f) {
    if (!f) return;
    const setFile = which === "a" ? setFileA : setFileB;
    const setParsing = which === "a" ? setParsingA : setParsingB;
    const setError = which === "a" ? setErrorA : setErrorB;
    setParsing(true);
    setError("");
    try {
      const parsed = await parseFile(f);
      setFile({ name: f.name, ...parsed });
      setResult(null);
      setResultError("");
    } catch (err) {
      setError(err.message || "Could not read this file.");
      setFile(null);
    } finally {
      setParsing(false);
    }
  }

  function switchSheet(which, sheetName) {
    const file = which === "a" ? fileA : fileB;
    const setFile = which === "a" ? setFileA : setFileB;
    if (!file || !file.workbook) return;
    const res = sheetToData(file.workbook, sheetName);
    if (res.error) {
      setResultError(res.error);
      return;
    }
    setFile({ ...file, selectedSheet: sheetName, rows: res.rows, headers: res.headers });
    setResult(null);
  }

  function runCompare() {
    if (!fileA || !fileB || !keyColumn) return;
    const res = compareDatasets(fileA.rows, fileA.headers, fileB.rows, fileB.headers, keyColumn);
    if (res.error) {
      setResultError(res.error);
      setResult(null);
      return;
    }
    setResultError("");
    setResult(res);
  }

  function resetAll() {
    setFileA(null);
    setFileB(null);
    setKeyColumn("");
    setResult(null);
    setResultError("");
    if (inputARef.current) inputARef.current.value = "";
    if (inputBRef.current) inputBRef.current.value = "";
  }

  if (result) {
    return (
      <CompareResults
        result={result}
        keyColumn={keyColumn}
        fileAName={fileA.name}
        fileBName={fileB.name}
        onReset={resetAll}
        onBack={() => setResult(null)}
      />
    );
  }

  return (
    <div className="space-y-5">
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        <CompareUploadSlot
          label="Earlier file"
          file={fileA}
          parsing={parsingA}
          error={errorA}
          inputRef={inputARef}
          onFile={(f) => handleFile("a", f)}
          onSwitchSheet={(s) => switchSheet("a", s)}
          onClear={() => {
            setFileA(null);
            setResult(null);
          }}
        />
        <CompareUploadSlot
          label="Later file"
          file={fileB}
          parsing={parsingB}
          error={errorB}
          inputRef={inputBRef}
          onFile={(f) => handleFile("b", f)}
          onSwitchSheet={(s) => switchSheet("b", s)}
          onClear={() => {
            setFileB(null);
            setResult(null);
          }}
        />
      </div>

      {fileA && fileB && (
        <Section title="Match rows by" hint="The column that uniquely identifies a row in both files — like a case reference or ID">
          {commonHeaders.length === 0 ? (
            <p className="text-sm" style={{ color: COLORS.danger }}>
              These files don't share any column names, so there's nothing to match on.
            </p>
          ) : (
            <select
              value={keyColumn}
              onChange={(e) => {
                setKeyColumn(e.target.value);
                setResultError("");
              }}
              className="px-2 py-2 rounded-lg text-sm w-full"
              style={{ border: `1px solid ${COLORS.border}`, backgroundColor: COLORS.surface }}
            >
              <option value="">Choose a column\u2026</option>
              {commonHeaders.map((h) => (
                <option key={h} value={h}>
                  {h}
                </option>
              ))}
            </select>
          )}
        </Section>
      )}

      {resultError && (
        <p className="text-sm flex items-center gap-1.5" style={{ color: COLORS.danger }}>
          <AlertTriangle size={14} /> {resultError}
        </p>
      )}

      <button
        onClick={runCompare}
        disabled={!fileA || !fileB || !keyColumn}
        className="px-4 py-2.5 rounded-lg text-sm disabled:opacity-40"
        style={{ backgroundColor: COLORS.accent, color: COLORS.surface, fontFamily: FONTS.body, fontWeight: 500 }}
      >
        Compare files
      </button>
    </div>
  );
}

// ---------------- main app ----------------

export default function ReportBuilder() {
  const [mode, setMode] = useState("report");
  const [presets, setPresets] = useState([]);
  const [presetsReady, setPresetsReady] = useState(false);
  const [file, setFile] = useState(null);
  const [parseError, setParseError] = useState("");
  const [parsing, setParsing] = useState(false);
  const [view, setView] = useState("home");
  const [editingPreset, setEditingPreset] = useState(null);
  const [activeResult, setActiveResult] = useState(null);
  const [activePresetName, setActivePresetName] = useState("");
  const [toast, setToast] = useState("");
  const fileInputRef = useRef(null);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const raw = localStorage.getItem("presets");
        if (!cancelled && raw) {
          const loaded = JSON.parse(raw);
          setPresets(Array.isArray(loaded) ? loaded.map(normalizePresetFields) : []);
        }
      } catch (e) {
        // no presets saved yet — that's fine
      } finally {
        if (!cancelled) setPresetsReady(true);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(""), 2600);
    return () => clearTimeout(t);
  }, [toast]);

  async function persistPresets(next) {
    setPresets(next);
    try {
      localStorage.setItem("presets", JSON.stringify(next));
    } catch (e) {
      setToast("Could not save that \u2014 try again.");
    }
  }

  async function handleFile(f) {
    if (!f) return;
    setParsing(true);
    setParseError("");
    try {
      const parsed = await parseFile(f);
      setFile({ name: f.name, ...parsed });
      setActiveResult(null);
    } catch (err) {
      setParseError(err.message || "Could not read this file.");
      setFile(null);
    } finally {
      setParsing(false);
    }
  }

  function switchSheet(sheetName) {
    if (!file || !file.workbook) return;
    const res = sheetToData(file.workbook, sheetName);
    if (res.error) {
      setToast(res.error);
      return;
    }
    setFile({ ...file, selectedSheet: sheetName, rows: res.rows, headers: res.headers });
    setActiveResult(null);
  }

  function startNewPreset() {
    setEditingPreset(emptyPreset());
    setView("builder");
  }

  function startEditPreset(preset) {
    setEditingPreset(JSON.parse(JSON.stringify(preset)));
    setView("builder");
  }

  function runPreset(preset) {
    if (!file) return;
    const result = applyPreset(preset, file.rows, file.headers);
    if (result.error) {
      setToast(result.error);
      return;
    }
    setActiveResult(result);
    setActivePresetName(preset.name);
    setView("results");
  }

  async function savePreset(preset) {
    if (!preset.name.trim()) {
      setToast("Give this preset a name first.");
      return;
    }
    const withHeaders = {
      ...preset,
      sourceHeaders: file && file.headers.length ? file.headers : preset.sourceHeaders || [],
    };
    const exists = presets.some((p) => p.id === withHeaders.id);
    const next = exists ? presets.map((p) => (p.id === withHeaders.id ? withHeaders : p)) : [...presets, withHeaders];
    await persistPresets(next);
    setToast(`Saved "${withHeaders.name}"`);
    if (file) {
      runPreset(withHeaders);
    } else {
      setView("home");
    }
  }

  async function deletePreset(id) {
    await persistPresets(presets.filter((p) => p.id !== id));
    setToast("Preset deleted");
  }

  function handleExportPresets() {
    if (!presets.length) return;
    exportPresets(presets);
    setToast("Presets exported");
  }

  async function handleImportPresets(f) {
    if (!f) return;
    try {
      const text = await f.text();
      const result = parsePresetImport(text);
      if (result.error) {
        setToast(result.error);
        return;
      }
      const next = [...presets, ...result.presets];
      await persistPresets(next);
      setToast(`Imported ${result.presets.length} preset${result.presets.length === 1 ? "" : "s"}`);
    } catch (e) {
      setToast("Could not read that file.");
    }
  }

  function reset() {
    setFile(null);
    setActiveResult(null);
    setParseError("");
    setView("home");
    if (fileInputRef.current) fileInputRef.current.value = "";
  }

  const step = view === "home" ? 0 : view === "builder" ? 1 : 2;

  return (
    <div className="min-h-screen w-full" style={{ backgroundColor: COLORS.bg, color: COLORS.ink, fontFamily: FONTS.body }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=Public+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; }
        input:focus, select:focus, button:focus-visible {
          outline: 2px solid ${COLORS.accent};
          outline-offset: 2px;
        }
        @media (prefers-reduced-motion: reduce) {
          * { animation: none !important; transition: none !important; }
        }
      `}</style>

      <div className="max-w-3xl mx-auto px-4 py-8 sm:px-6 sm:py-10">
        <header className="mb-6">
          <img
            src={LOGO_DATA_URI}
            alt="Report Builder \u2014 manipulate spreadsheets quickly and easily"
            style={{ width: "100%", maxWidth: "380px", height: "auto", display: "block", marginBottom: "10px" }}
          />
          <p className="text-sm" style={{ color: COLORS.muted }}>
            {mode === "report" ? "Upload a spreadsheet, apply a saved report, done." : "See what changed between two files."}
          </p>
        </header>

        <div className="flex items-center gap-2 mb-6">
          <ModeTab active={mode === "report"} onClick={() => setMode("report")} label="Build a report" />
          <ModeTab active={mode === "compare"} onClick={() => setMode("compare")} label="Compare files" />
        </div>

        {mode === "report" ? (
          <>
            <Stepper step={step} />

            {view === "home" && (
              <HomeView
                file={file}
                parsing={parsing}
                parseError={parseError}
                presets={presets}
                presetsReady={presetsReady}
                fileInputRef={fileInputRef}
                onFile={handleFile}
                onSwitchSheet={switchSheet}
                onNewPreset={startNewPreset}
                onEditPreset={startEditPreset}
                onDeletePreset={deletePreset}
                onRunPreset={runPreset}
                onClearFile={reset}
                onExportPresets={handleExportPresets}
                onImportPresets={handleImportPresets}
              />
            )}

            {view === "builder" && editingPreset && (
              <BuilderView
                preset={editingPreset}
                setPreset={setEditingPreset}
                headers={file?.headers?.length ? file.headers : editingPreset.sourceHeaders || []}
                hasFile={!!file}
                onCancel={() => setView("home")}
                onSave={savePreset}
              />
            )}

            {view === "results" && activeResult && (
              <ResultsView
                result={activeResult}
                presetName={activePresetName}
                fileName={file?.name}
                onBack={() => setView("home")}
                onReset={reset}
              />
            )}
          </>
        ) : (
          <CompareView />
        )}

        {toast && (
          <div
            className="fixed bottom-6 left-1/2 -translate-x-1/2 px-4 py-2.5 rounded-lg text-sm shadow-lg"
            style={{ backgroundColor: COLORS.ink, color: COLORS.surface, fontFamily: FONTS.body, maxWidth: "90%" }}
          >
            {toast}
          </div>
        )}
      </div>
    </div>
  );
}
