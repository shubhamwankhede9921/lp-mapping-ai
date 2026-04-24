import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { PipelineOutputAnalysis } from "./components/PipelineOutputAnalysis";
import type { PipelineAnalysis } from "./utils/pipelineZipAnalysis";
import { analyzePipelineZipBlob } from "./utils/pipelineZipAnalysis";
import { approveEditSession, createEditSession, updateEditSession } from "./api/mappingApi";

type TabId = "refs" | "det" | "hybrid" | "pipeline" | "nested";

type FilePreview = {
  fileName: string;
  mime: string;
  rows: string[][];
  truncated: boolean;
  error?: string;
};

type MatchRow = {
  partnerField: string;
  excelKey?: string;
  jsonKey: string;
  entity?: string;
  confidence: number; // 0..1
  matchType: string;
  engine: "deterministic" | "fuzzy" | "embeddings" | "llm";
  review: boolean;
  reasoning?: string;
};

const STYLE = `
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  :root {
    --indigo: #4F46E5; --indigo-l: #EEF2FF; --indigo-d: #3730A3;
    --violet: #7C3AED; --violet-l: #F5F3FF;
    --cyan: #0891B2; --cyan-l: #ECFEFF;
    --emerald: #059669; --emerald-l: #ECFDF5;
    --amber: #D97706; --amber-l: #FFFBEB;
    --rose: #E11D48; --rose-l: #FFF1F2;
    --slate: #64748B; --slate-l: #F8FAFC;
    --ink: #0F172A; --ink-2: #334155; --ink-3: #64748B;
    --border: rgba(15,23,42,0.08);
    --radius: 14px; --radius-sm: 10px; --radius-xs: 8px;
  }
  body { font-family: var(--font-sans, system-ui, sans-serif); background: #F1F5F9; color: var(--ink); font-size: 20px; }
  .sr-only{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0,0,0,0);white-space:nowrap;}
  .shell { display: flex; min-height: 100vh; }

  .sidebar {
    width: 240px; flex-shrink: 0;
    background: var(--ink);
    display: flex; flex-direction: column;
    padding: 0 0 1.5rem;
  }
  .sb-logo {
    padding: 1.25rem 1.25rem 1rem;
    border-bottom: 1px solid rgba(255,255,255,0.08);
    margin-bottom: 0.75rem;
  }
  .sb-logo-mark {
    width: 36px; height: 36px; border-radius: 10px;
    background: linear-gradient(135deg, #4F46E5, #7C3AED);
    display: flex; align-items: center; justify-content: center;
    font-size: 18px; margin-bottom: 0.5rem;
  }
  .sb-brand { color: #fff; font-size: 17px; font-weight: 600; line-height: 1.2; }
  .sb-sub { color: rgba(255,255,255,0.45); font-size: 12px; margin-top: 2px; }

  .sb-section { padding: 0 0.75rem 0.25rem; }
  .sb-section-label {
    font-size: 11px; font-weight: 600; letter-spacing: 0.08em;
    text-transform: uppercase; color: rgba(255,255,255,0.35);
    padding: 0.5rem 0.5rem 0.25rem;
  }
  .nav-item {
    display: flex; align-items: center; gap: 10px;
    padding: 0.6rem 0.75rem; border-radius: var(--radius-sm);
    cursor: pointer; color: rgba(255,255,255,0.55);
    font-size: 15px; font-weight: 500;
    transition: all 0.15s; border: none; background: none;
    width: 100%; text-align: left;
    position: relative;
  }
  .nav-item:hover { background: rgba(255,255,255,0.07); color: rgba(255,255,255,0.85); }
  .nav-item.active { background: rgba(79,70,229,0.25); color: #fff; font-weight: 500; }
  .nav-item.active::before {
    content: ''; position: absolute; left: 0; top: 20%; bottom: 20%;
    width: 3px; background: #818CF8; border-radius: 2px;
  }
  .nav-dot {
    width: 28px; height: 28px; border-radius: 8px;
    display: flex; align-items: center; justify-content: center; font-size: 13px;
    flex-shrink: 0;
  }
  .nd-indigo { background: rgba(79,70,229,0.2); }
  .nd-violet { background: rgba(124,58,237,0.2); }
  .nd-cyan { background: rgba(8,145,178,0.2); }
  .nd-emerald { background: rgba(5,150,105,0.2); }
  .nd-amber { background: rgba(217,119,6,0.2); }

  .sb-divider { height: 1px; background: rgba(255,255,255,0.07); margin: 0.75rem 1rem; }

  .sb-fields { padding: 0 0.75rem; display: flex; flex-direction: column; gap: 0.5rem; }
  .sb-field-label { font-size: 11px; color: rgba(255,255,255,0.4); margin-bottom: 3px; display: block; }
  .sb-input {
    width: 100%; background: rgba(255,255,255,0.07);
    border: 1px solid rgba(255,255,255,0.1);
    border-radius: var(--radius-xs); padding: 0.45rem 0.6rem;
    color: rgba(255,255,255,0.85); font-size: 14px; outline: none;
    font-family: inherit;
  }
  .sb-input:focus { border-color: #818CF8; background: rgba(129,140,248,0.1); }

  .sb-toggles { padding: 0 0.75rem; display: flex; flex-direction: column; gap: 2px; }
  .toggle-row {
    display: flex; align-items: center; justify-content: space-between;
    padding: 0.4rem 0.5rem; border-radius: var(--radius-xs);
    color: rgba(255,255,255,0.6); font-size: 12px;
    cursor: pointer;
    user-select: none;
  }
  .toggle-row:hover { background: rgba(255,255,255,0.05); }
  .pill-toggle {
    width: 32px; height: 18px; border-radius: 9px;
    position: relative; cursor: pointer; flex-shrink: 0;
    transition: background 0.2s;
  }
  .pill-toggle.on { background: #4F46E5; }
  .pill-toggle.off { background: rgba(255,255,255,0.15); }
  .pill-toggle::after {
    content: ''; position: absolute; top: 2px; width: 14px; height: 14px;
    background: white; border-radius: 50%; transition: left 0.2s;
  }
  .pill-toggle.on::after { left: 16px; }
  .pill-toggle.off::after { left: 2px; }

  .main { flex: 1; display: flex; flex-direction: column; overflow: auto; }

  .topbar {
    background: white; border-bottom: 1px solid var(--border);
    padding: 1rem 1.75rem;
    display: flex; align-items: center; justify-content: space-between;
    gap: 1rem; flex-shrink: 0;
  }
  .topbar-left { display: flex; flex-direction: column; gap: 2px; }
  .page-title { font-size: 20px; font-weight: 650; color: var(--ink); }
  .page-crumb { font-size: 13px; color: var(--ink-3); }
  .topbar-right { display: flex; align-items: center; gap: 0.75rem; }

  .badge {
    display: inline-flex; align-items: center; gap: 5px;
    padding: 0.3rem 0.7rem; border-radius: 20px; font-size: 11px; font-weight: 500;
  }
  .badge-indigo { background: var(--indigo-l); color: var(--indigo-d); }
  .badge-emerald { background: var(--emerald-l); color: #065F46; }
  .badge-amber { background: var(--amber-l); color: #92400E; }
  .badge-rose { background: var(--rose-l); color: #9F1239; }
  .badge-violet { background: var(--violet-l); color: #5B21B6; }
  .badge-cyan { background: var(--cyan-l); color: #164E63; }
  .badge-dot { width: 6px; height: 6px; border-radius: 50%; background: currentColor; }

  .btn {
    display: inline-flex; align-items: center; gap: 6px;
    padding: 0.5rem 1rem; border-radius: var(--radius-sm);
    font-size: 15px; font-weight: 650; cursor: pointer;
    border: none; font-family: inherit; transition: all 0.15s;
  }
  .btn-primary { background: var(--indigo); color: white; }
  .btn-primary:hover { background: var(--indigo-d); transform: translateY(-1px); }
  .btn-secondary {
    background: white; color: var(--ink-2);
    border: 1px solid var(--border);
  }
  .btn-secondary:hover { background: var(--slate-l); border-color: #CBD5E1; }
  .btn-ghost { background: transparent; color: var(--ink-3); border: 1px solid var(--border); }
  .btn-ghost:hover { background: white; color: var(--ink); }
  .btn:disabled { opacity: 0.4; cursor: not-allowed; transform: none !important; }

  .content { padding: 1.5rem 1.75rem; flex: 1; }

  .metric-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 1rem; margin-bottom: 1.5rem; }
  .metric-card {
    background: white; border-radius: var(--radius); padding: 1.25rem;
    border: 1px solid var(--border); position: relative; overflow: hidden;
  }
  .metric-accent {
    position: absolute; top: 0; left: 0; right: 0; height: 3px;
  }
  .acc-indigo { background: linear-gradient(90deg, #4F46E5, #818CF8); }
  .acc-emerald { background: linear-gradient(90deg, #059669, #34D399); }
  .acc-violet { background: linear-gradient(90deg, #7C3AED, #A78BFA); }
  .acc-amber { background: linear-gradient(90deg, #D97706, #FCD34D); }
  .metric-label { font-size: 11px; color: var(--ink-3); text-transform: uppercase; letter-spacing: 0.06em; margin-bottom: 0.5rem; }
  .metric-value { font-size: 28px; font-weight: 500; line-height: 1; margin-bottom: 0.35rem; }
  .mv-indigo { color: var(--indigo); }
  .mv-emerald { color: var(--emerald); }
  .mv-violet { color: var(--violet); }
  .mv-amber { color: var(--amber); }
  .metric-sub { font-size: 11px; color: var(--ink-3); }

  .panel {
    background: white; border-radius: var(--radius);
    border: 1px solid var(--border); overflow: hidden;
  }
  .panel-header {
    padding: 1.25rem 1.5rem; border-bottom: 1px solid var(--border);
    display: flex; align-items: center; justify-content: space-between;
    gap: 1rem;
  }
  .panel-title { font-size: 16px; font-weight: 650; color: var(--ink); }
  .panel-desc { font-size: 13px; color: var(--ink-3); margin-top: 2px; }
  .panel-body { padding: 1.5rem; }

  .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 1.5rem; margin-bottom: 1.5rem; }

  .field { margin-bottom: 1rem; }
  .field-label { display: block; font-size: 13px; color: var(--ink-3); margin-bottom: 5px; font-weight: 650; }
  .field-input, .field-textarea {
    width: 100%; padding: 0.55rem 0.75rem;
    border: 1px solid #E2E8F0; border-radius: var(--radius-sm);
    font-size: 17px; font-weight: 700; color: var(--ink); background: white; outline: none;
    font-family: inherit; transition: border-color 0.15s;
  }
  .field-input:focus, .field-textarea:focus { border-color: var(--indigo); box-shadow: 0 0 0 3px rgba(79,70,229,0.1); }
  .field-textarea { min-height: 160px; resize: vertical; font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", monospace; font-size: 12px; }

  .upload-zone {
    border: 2px dashed #CBD5E1; border-radius: var(--radius);
    padding: 2rem; text-align: center; cursor: pointer;
    background: #F8FAFC; transition: all 0.2s; position: relative;
    display: flex; flex-direction: column; align-items: center; gap: 0.75rem;
    margin-bottom: 1.25rem;
  }
  .upload-zone:hover { border-color: var(--indigo); background: var(--indigo-l); }
  .upload-zone input { position: absolute; inset: 0; opacity: 0; cursor: pointer; }
  .upload-icon {
    width: 44px; height: 44px; border-radius: 12px;
    background: var(--indigo-l); display: flex; align-items: center; justify-content: center; font-size: 20px;
  }
  .upload-title { font-size: 13px; font-weight: 500; color: var(--ink-2); }
  .upload-hint { font-size: 12px; color: var(--ink-3); }

  .action-bar { display: flex; gap: 0.75rem; flex-wrap: wrap; margin: 1.25rem 0; }

  .table-wrap { overflow: auto; max-height: 380px; border: 1px solid var(--border); border-radius: var(--radius-sm); }
  .data-table { width: 100%; border-collapse: collapse; font-size: 15px; }
  .data-table th {
    background: #F8FAFC; padding: 0.65rem 0.85rem;
    text-align: left; font-size: 14px; font-weight: 800;
    text-transform: uppercase; letter-spacing: 0.07em; color: var(--ink-3);
    border-bottom: 1px solid var(--border); position: sticky; top: 0;
  }
  .data-table td { padding: 0.75rem 0.95rem; border-bottom: 1px solid #F1F5F9; color: var(--ink-2); font-weight: 650; font-size: 15px; font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", monospace; }
  .data-table tr:last-child td { border-bottom: none; }
  .data-table tr:hover td { background: #FAFAFA; }

  .chip {
    display: inline-flex; align-items: center; padding: 2px 8px;
    border-radius: 20px; font-size: 10px; font-weight: 500; font-family: var(--font-sans);
    border: 1px solid;
  }
  .chip-green { background: #ECFDF5; color: #065F46; border-color: #A7F3D0; }
  .chip-blue { background: #EFF6FF; color: #1E40AF; border-color: #BFDBFE; }
  .chip-amber { background: #FFFBEB; color: #92400E; border-color: #FDE68A; }
  .chip-red { background: #FFF1F2; color: #9F1239; border-color: #FECDD3; }

  .engine-grid { display: flex; flex-wrap: wrap; gap: 0.5rem; margin: 1rem 0; }
  .eng-tile {
    background: var(--slate-l); border: 1px solid var(--border);
    border-radius: var(--radius-sm); padding: 0.65rem 1rem;
    display: flex; flex-direction: column; align-items: flex-start;
    min-width: 100px;
  }
  .eng-label { font-size: 10px; color: var(--ink-3); text-transform: uppercase; letter-spacing: 0.06em; }
  .eng-count { font-size: 22px; font-weight: 500; color: var(--indigo); margin-top: 2px; }

  .alert {
    display: flex; align-items: flex-start; gap: 10px;
    padding: 0.85rem 1rem; border-radius: var(--radius-sm);
    font-size: 13px; margin-bottom: 1rem; border: 1px solid;
  }
  .alert-error { background: var(--rose-l); color: #9F1239; border-color: #FECDD3; }
  .alert-success { background: var(--emerald-l); color: #065F46; border-color: #A7F3D0; }
  .alert-dismiss { margin-left: auto; background: none; border: none; cursor: pointer; font-size: 16px; color: inherit; line-height: 1; opacity: 0.6; }
  .alert-dismiss:hover { opacity: 1; }

  .loading-strip {
    height: 3px; background: #EEF2FF; border-radius: 2px; overflow: hidden; margin-bottom: 1rem;
  }
  .loading-fill {
    height: 100%; width: 35%; background: linear-gradient(90deg, transparent, var(--indigo), transparent);
    animation: shimmer 1.1s ease-in-out infinite;
  }
  @keyframes shimmer { 0%{transform:translateX(-200%)} 100%{transform:translateX(500%)} }

  .processing-badge {
    display: inline-flex; align-items: center; gap: 6px;
    padding: 0.35rem 0.85rem; border-radius: 20px;
    background: var(--indigo-l); color: var(--indigo-d);
    font-size: 12px; font-weight: 500; margin-bottom: 0.75rem;
  }
  .proc-dot { width: 7px; height: 7px; border-radius: 50%; background: var(--indigo); animation: blink 0.9s ease-in-out infinite; }
  @keyframes blink { 0%,100%{opacity:0.3} 50%{opacity:1} }

  .refs-status { background: #0F172A; border-radius: var(--radius-sm); padding: 1rem 1.25rem; color: #94A3B8; font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", monospace; font-size: 12px; white-space: pre; margin-top: 1rem; line-height: 1.7; }
  .pre-block { background: #F8FAFC; border: 1px solid var(--border); border-radius: var(--radius-sm); padding: 1rem 1.25rem; font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", monospace; font-size: 12px; white-space: pre; overflow: auto; max-height: 320px; line-height: 1.7; color: var(--ink-2); margin-top: 1rem; }

  .sb-status-row { display: flex; align-items: center; gap: 6px; padding: 0.5rem 0.5rem 0; }
  .status-dot-on { width: 7px; height: 7px; border-radius: 50%; background: #34D399; }
  .status-dot-off { width: 7px; height: 7px; border-radius: 50%; background: rgba(255,255,255,0.2); }
  .status-text { font-size: 11px; color: rgba(255,255,255,0.45); }

  .color-accent-bar { display: flex; height: 4px; border-radius: 2px; overflow: hidden; gap: 2px; margin: 0.75rem 0; }
  .cab-seg { flex: 1; border-radius: 2px; }

  .file-name-tag {
    display: inline-flex; align-items: center; gap: 6px;
    background: var(--indigo-l); color: var(--indigo-d);
    border: 1px solid #C7D2FE; border-radius: var(--radius-xs);
    padding: 3px 10px; font-size: 11px; font-weight: 500; margin-left: 0.75rem;
  }

  .result-head { display: flex; align-items: center; gap: 0.75rem; margin-bottom: 1rem; }
  .result-head-label { font-size: 13px; font-weight: 500; color: var(--ink); }

  .stack-gap { display: flex; flex-direction: column; gap: 1rem; }
`;

const PAGE_META: Record<
  TabId,
  { title: string; crumb: string; badgeClass: string; badgeTxt: string }
> = {
  refs: {
    title: "Reference files",
    crumb: "LP Field Mapping · Dict Management",
    badgeClass: "badge-indigo",
    badgeTxt: "Dict Mgmt",
  },
  det: {
    title: "Deterministic matching",
    crumb: "LP Field Mapping · Phase 1",
    badgeClass: "badge-violet",
    badgeTxt: "Phase 1",
  },
  hybrid: {
    title: "Hybrid + LLM",
    crumb: "LP Field Mapping · Full Stack",
    badgeClass: "badge-cyan",
    badgeTxt: "Full Stack",
  },
  pipeline: {
    title: "Full pipeline",
    crumb: "LP Field Mapping · ZIP Output",
    badgeClass: "badge-emerald",
    badgeTxt: "ZIP Output",
  },
  nested: {
    title: "Nested mapping & schema",
    crumb: "LP Field Mapping · Schema Builder",
    badgeClass: "badge-amber",
    badgeTxt: "Schema",
  },
};

const DEFAULT_NESTED_INPUT = `{
  "client_name": "HDFC Bank",
  "process_name": "COMBINED",
  "mappings": []
}`;

function formatPct01(confidence: number) {
  return `${Math.round(confidence * 100)}%`;
}

function downloadCsv(filename: string, headers: string[], rows: Record<string, string>[]) {
  const esc = (s: string) => {
    const t = String(s ?? "");
    if (/[",\n]/.test(t)) return `"${t.replace(/"/g, '""')}"`;
    return t;
  };
  const lines = [headers.map(esc).join(",")];
  for (const r of rows) {
    lines.push(headers.map((h) => esc(r[h] ?? "")).join(","));
  }
  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(a.href);
}

function confidenceChipClass(confidence: number) {
  if (confidence >= 0.85) return "chip chip-green";
  if (confidence >= 0.65) return "chip chip-blue";
  if (confidence >= 0.4) return "chip chip-amber";
  return "chip chip-red";
}

function parseCsvLine(line: string, delimiter: string) {
  const out: string[] = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (ch === '"') {
      const peek = line[i + 1];
      if (inQuotes && peek === '"') {
        cell += '"';
        i += 1;
        continue;
      }
      inQuotes = !inQuotes;
      continue;
    }
    if (!inQuotes && ch === delimiter) {
      out.push(cell);
      cell = "";
      continue;
    }
    cell += ch;
  }

  out.push(cell);
  return out;
}

function detectDelimiter(firstLine: string) {
  const candidates: Array<{ d: string; count: number }> = [
    { d: ",", count: (firstLine.match(/,/g) ?? []).length },
    { d: "\t", count: (firstLine.match(/\t/g) ?? []).length },
    { d: ";", count: (firstLine.match(/;/g) ?? []).length },
    { d: "|", count: (firstLine.match(/\|/g) ?? []).length },
  ];
  candidates.sort((a, b) => b.count - a.count);
  return candidates[0]?.count ? candidates[0].d : ",";
}

function countEmptyCells(rows: string[][]) {
  let empties = 0;
  for (const row of rows) {
    for (const cell of row) {
      if (cell.trim().length === 0) empties += 1;
    }
  }
  return empties;
}

async function buildPreview(file: File, maxRows: number): Promise<FilePreview> {
  const mime = file.type || "application/octet-stream";
  const lower = file.name.toLowerCase();

  if (lower.endsWith(".csv") || lower.endsWith(".tsv") || lower.endsWith(".txt")) {
    const text = await file.text();
    const lines = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
    const nonEmptyLines = lines.filter((l) => l.trim().length > 0);
    const firstLine = nonEmptyLines[0] ?? "";
    const delimiter = detectDelimiter(firstLine);

    const rows = nonEmptyLines.slice(0, maxRows).map((l) => parseCsvLine(l, delimiter));
    return {
      fileName: file.name,
      mime,
      rows,
      truncated: nonEmptyLines.length > maxRows,
    };
  }

  if (lower.endsWith(".xlsx") || lower.endsWith(".xls")) {
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const firstSheetName = wb.SheetNames[0];
      if (!firstSheetName) {
        return { fileName: file.name, mime, rows: [], truncated: false, error: "No worksheets found in Excel file." };
      }
      const ws = wb.Sheets[firstSheetName];
      const allRows = XLSX.utils.sheet_to_json(ws, {
        header: 1,
        raw: false,
        blankrows: false,
        defval: "",
      }) as unknown as Array<Array<unknown>>;

      const rows = allRows.slice(0, maxRows).map((r) => r.map((v) => String(v ?? "")));
      return {
        fileName: file.name,
        mime,
        rows,
        truncated: allRows.length > maxRows,
      };
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Failed to parse Excel file.";
      return { fileName: file.name, mime, rows: [], truncated: false, error: msg };
    }
  }

  return {
    fileName: file.name,
    mime,
    rows: [],
    truncated: false,
    error: "Unsupported file type for preview. Upload .csv/.tsv for in-app preview.",
  };
}

function PreviewPanel({ preview }: { preview: FilePreview | null }) {
  if (!preview) return null;

  const head = preview.rows[0] ?? [];
  const body = preview.rows.slice(1);
  const shownRows = preview.rows.length > 0 ? preview.rows.length - 1 : 0;
  const columnCount = Math.max(0, ...preview.rows.map((r) => r.length));
  const emptyCells = preview.rows.length ? countEmptyCells(preview.rows) : 0;

  return (
    <div className="panel" style={{ marginTop: "1rem" }}>
      <div className="panel-header">
        <div>
          <div className="panel-title">
            Excel preview
            <span className="file-name-tag">📄 {preview.fileName}</span>
          </div>
          <div className="panel-desc">
            {preview.error ? "Preview unavailable" : "First rows snapshot"} · {preview.mime}
          </div>
        </div>
        <span className={`badge ${preview.error ? "badge-rose" : "badge-emerald"}`}>
          <span className="badge-dot"></span> {preview.error ? "Needs CSV" : "Ready"}
        </span>
      </div>

      <div className="panel-body">
        {preview.error ? (
          <div className="alert alert-error" style={{ marginBottom: 0 }}>
            <span>⚠</span>
            <span style={{ flex: 1 }}>{preview.error}</span>
          </div>
        ) : (
          <>
            <div className="metric-row" style={{ gridTemplateColumns: "repeat(4,1fr)" }}>
              <div className="metric-card">
                <div className="metric-accent acc-indigo"></div>
                <div className="metric-label">Columns</div>
                <div className="metric-value mv-indigo">{columnCount}</div>
                <div className="metric-sub">Detected from preview</div>
              </div>
              <div className="metric-card">
                <div className="metric-accent acc-emerald"></div>
                <div className="metric-label">Rows shown</div>
                <div className="metric-value mv-emerald">{shownRows}</div>
                <div className="metric-sub">
                  {preview.truncated ? "Truncated preview" : "Complete preview"}
                </div>
              </div>
              <div className="metric-card">
                <div className="metric-accent acc-amber"></div>
                <div className="metric-label">Empty cells</div>
                <div className="metric-value mv-amber">{emptyCells}</div>
                <div className="metric-sub">Blank values found</div>
              </div>
              <div className="metric-card">
                <div className="metric-accent acc-violet"></div>
                <div className="metric-label">Header</div>
                <div className="metric-value mv-violet">{head.length ? "Yes" : "No"}</div>
                <div className="metric-sub">Uses first line</div>
              </div>
            </div>

            <div className="table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                    {head.map((h, idx) => (
                      <th key={`${idx}:${h}`}>{h || `Column ${idx + 1}`}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {body.map((r, rIdx) => (
                    <tr key={`r:${rIdx}`}>
                      {head.map((_, cIdx) => (
                        <td key={`c:${rIdx}:${cIdx}`}>{r[cIdx] ?? ""}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

function EngineBreakdown({ rows }: { rows: MatchRow[] }) {
  const counts = useMemo(() => {
    const init = { deterministic: 0, fuzzy: 0, embeddings: 0, llm: 0, unmatched: 0 };
    for (const r of rows) {
      if (!r.jsonKey) init.unmatched += 1;
      else init[r.engine] += 1;
    }
    return init;
  }, [rows]);

  const matched = counts.deterministic + counts.fuzzy + counts.embeddings + counts.llm;
  const total = matched + counts.unmatched;
  const avg = rows.length
    ? rows.reduce((acc, r) => acc + r.confidence, 0) / Math.max(1, rows.length)
    : 0;
  const reviewNeeded = rows.filter((r) => r.review).length;

  return (
    <>
      <div className="result-head">
        <div className="result-head-label">Colorful analysis</div>
        <span className="badge badge-indigo">
          <span className="badge-dot"></span> {matched}/{total} matched
        </span>
        <span className={`badge ${reviewNeeded ? "badge-amber" : "badge-emerald"}`}>
          <span className="badge-dot"></span> {reviewNeeded} review
        </span>
      </div>

      <div className="color-accent-bar">
        <div className="cab-seg" style={{ background: "#4F46E5" }}></div>
        <div className="cab-seg" style={{ background: "#7C3AED" }}></div>
        <div className="cab-seg" style={{ background: "#0891B2" }}></div>
        <div className="cab-seg" style={{ background: "#059669" }}></div>
      </div>

      <div className="engine-grid">
        <div className="eng-tile">
          <div className="eng-label">Deterministic</div>
          <div className="eng-count">{counts.deterministic}</div>
        </div>
        <div className="eng-tile">
          <div className="eng-label">Fuzzy</div>
          <div className="eng-count">{counts.fuzzy}</div>
        </div>
        <div className="eng-tile">
          <div className="eng-label">Embeddings</div>
          <div className="eng-count">{counts.embeddings}</div>
        </div>
        <div className="eng-tile">
          <div className="eng-label">LLM</div>
          <div className="eng-count">{counts.llm}</div>
        </div>
        <div className="eng-tile">
          <div className="eng-label">Unmatched</div>
          <div className="eng-count">{counts.unmatched}</div>
        </div>
      </div>

      <div className="metric-row" style={{ gridTemplateColumns: "repeat(4,1fr)" }}>
        <div className="metric-card">
          <div className="metric-accent acc-indigo"></div>
          <div className="metric-label">Total</div>
          <div className="metric-value mv-indigo">{total}</div>
          <div className="metric-sub">Input fields</div>
        </div>
        <div className="metric-card">
          <div className="metric-accent acc-emerald"></div>
          <div className="metric-label">Matched</div>
          <div className="metric-value mv-emerald">{matched}</div>
          <div className="metric-sub">
            {total ? `${Math.round((matched / total) * 1000) / 10}% coverage` : "—"}
          </div>
        </div>
        <div className="metric-card">
          <div className="metric-accent acc-violet"></div>
          <div className="metric-label">Avg confidence</div>
          <div className="metric-value" style={{ color: "var(--violet)" }}>
            {formatPct01(avg)}
          </div>
          <div className="metric-sub">Across engines</div>
        </div>
        <div className="metric-card">
          <div className="metric-accent acc-amber"></div>
          <div className="metric-label">Review needed</div>
          <div className="metric-value mv-amber">{reviewNeeded}</div>
          <div className="metric-sub">Low confidence</div>
        </div>
      </div>
    </>
  );
}

export default function App() {
  const [tab, setTab] = useState<TabId>("refs");
  const [loading, setLoading] = useState(false);
  const [globalError, setGlobalError] = useState<string | null>(null);

  const [clientName, setClientName] = useState("HDFC Bank");
  const [processName, setProcessName] = useState("COMBINED");
  const [masterId, setMasterId] = useState(1);
  const [apiBaseUrl, setApiBaseUrl] = useState("");

  const [toggles, setToggles] = useState({
    fuzzy: true,
    embeddings: false,
    llm: true,
    loanParamRefine: true,
    entityClassifier: false,
    includeRefsZip: true,
    saveToDb: false,
    skipUnmatched: false,
  });

  const [health, setHealth] = useState<{ online: boolean; text: string }>({
    online: false,
    text: "Offline",
  });

  const [refStatus, setRefStatus] = useState<string | null>(null);
  const [buildOut, setBuildOut] = useState<string | null>(null);

  const [detFile, setDetFile] = useState<File | null>(null);
  const [detPreview, setDetPreview] = useState<FilePreview | null>(null);
  const [detRows, setDetRows] = useState<MatchRow[] | null>(null);

  const [hybFile, setHybFile] = useState<File | null>(null);
  const [hybPreview, setHybPreview] = useState<FilePreview | null>(null);
  const [hybRows, setHybRows] = useState<MatchRow[] | null>(null);
  const [hybDraftRows, setHybDraftRows] = useState<MatchRow[] | null>(null);
  const [editSessionId, setEditSessionId] = useState<string | null>(null);
  const [editSessionStatus, setEditSessionStatus] = useState<string | null>(null);
  const [editSessionMsg, setEditSessionMsg] = useState<string | null>(null);

  const [fpFile, setFpFile] = useState<File | null>(null);
  const [fpPreview, setFpPreview] = useState<FilePreview | null>(null);
  const [fpDone, setFpDone] = useState(false);
  const [fpZipUrl, setFpZipUrl] = useState<string | null>(null);
  const [fpZipName, setFpZipName] = useState<string | null>(null);
  const [fpPipelineAnalysis, setFpPipelineAnalysis] = useState<PipelineAnalysis | null>(null);
  const [fpPipelineAnalysisErr, setFpPipelineAnalysisErr] = useState<string | null>(null);
  const [fpDbMetrics, setFpDbMetrics] = useState<{
    inserted: string;
    skipped: string;
    errors: string;
  } | null>(null);

  const [nestedJson, setNestedJson] = useState(DEFAULT_NESTED_INPUT);
  const [nestedOut, setNestedOut] = useState<string | null>(null);
  const [schemaOut, setSchemaOut] = useState<string | null>(null);

  const detInputRef = useRef<HTMLInputElement | null>(null);
  const hybInputRef = useRef<HTMLInputElement | null>(null);
  const fpInputRef = useRef<HTMLInputElement | null>(null);

  const pageMeta = PAGE_META[tab];

  useEffect(() => {
    return () => {
      if (fpZipUrl) URL.revokeObjectURL(fpZipUrl);
    };
  }, [fpZipUrl]);

  function apiJoin(path: string) {
    const base = apiBaseUrl.trim();
    if (!base) return path; // rely on Vite proxy
    return `${base.replace(/\/+$/, "")}${path.startsWith("/") ? "" : "/"}${path}`;
  }

  async function readJsonOrThrow(res: Response) {
    const text = await res.text();
    if (!res.ok) {
      throw new Error(text || `${res.status} ${res.statusText}`);
    }
    try {
      return text ? JSON.parse(text) : null;
    } catch {
      return text;
    }
  }

  function toMatchRow(m: any): MatchRow {
    const engineRaw = (m?.winning_engine ?? m?.engine ?? "deterministic") as string;
    const engine: MatchRow["engine"] =
      engineRaw === "deterministic" || engineRaw === "fuzzy" || engineRaw === "embeddings" || engineRaw === "llm"
        ? engineRaw
        : "deterministic";
    return {
      partnerField: String(m?.partner_field ?? m?.partnerField ?? ""),
      excelKey: m?.matched_excel_key ?? m?.excel_key ?? m?.excelKey ?? undefined,
      jsonKey: String(m?.json_key ?? m?.lms_column ?? m?.jsonKey ?? ""),
      entity: m?.entity ? String(m.entity) : undefined,
      confidence: Number(m?.confidence ?? 0),
      matchType: String(m?.match_type ?? m?.matchType ?? "unmatched"),
      engine,
      review: Boolean(m?.needs_review ?? m?.review ?? false),
      reasoning: m?.reasoning ? String(m.reasoning) : undefined,
    };
  }

  async function pingHealth() {
    try {
      const res = await fetch(apiJoin("/api/llm_mapping/health"));
      const data = await readJsonOrThrow(res);
      setHealth({ online: true, text: typeof data === "object" && data ? `API online (${data.status ?? "ok"})` : "API online" });
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Offline";
      setHealth({ online: false, text: `Offline (${msg})` });
    }
  }

  async function withLoading<T>(fn: () => Promise<T>) {
    setGlobalError(null);
    setLoading(true);
    try {
      return await fn();
    } catch (err) {
      const msg = err instanceof Error ? err.message : "Unknown error";
      setGlobalError(msg);
      throw err;
    } finally {
      setLoading(false);
    }
  }

  async function runRefStatus() {
    await withLoading(async () => {
      const res = await fetch(apiJoin("/api/llm_mapping/references/status"));
      const data = await readJsonOrThrow(res);
      if (!data || typeof data !== "object") {
        setRefStatus(String(data ?? ""));
        return;
      }
      const ready = Boolean((data as any).ready);
      const files = (data as any).files ?? {};
      const lines = [`Ready: ${ready}`];
      for (const [name, meta] of Object.entries(files)) {
        const exists = Boolean((meta as any)?.exists);
        const sizeKb = (meta as any)?.size_kb;
        lines.push(`${exists ? "OK " : "MISSING"} ${name}${exists && typeof sizeKb === "number" ? ` — ${sizeKb} KB` : ""}`);
      }
      setRefStatus(lines.join("\n"));
    });
  }

  async function runBuild() {
    await withLoading(async () => {
      const q = new URLSearchParams();
      if (refsPutmOverride.trim()) q.set("putm_table_override", refsPutmOverride.trim());
      if (refsMappingOverride.trim()) q.set("mapping_table_override", refsMappingOverride.trim());
      const url = `${apiJoin("/api/llm_mapping/references/build")}${q.toString() ? `?${q.toString()}` : ""}`;
      const res = await fetch(url, { method: "POST" });
      const data = await readJsonOrThrow(res);
      setBuildOut(typeof data === "string" ? data : JSON.stringify(data, null, 2));
    });
  }

  type PutmKeyRow = {
    excel_key: string;
    json_key: string;
    role?: string | null;
    process_name?: string | null;
    process_names?: string[] | null;
    description?: string | null;
    example?: string | null;
  };

  const [refsPutmOverride, setRefsPutmOverride] = useState("");
  const [refsMappingOverride, setRefsMappingOverride] = useState("");
  const [putmKeys, setPutmKeys] = useState<PutmKeyRow[] | null>(null);
  const [putmKeysErr, setPutmKeysErr] = useState<string | null>(null);
  const [putmProcess, setPutmProcess] = useState<string>("ALL");
  const [putmSearch, setPutmSearch] = useState<string>("");
  const [putmKeysMeta, setPutmKeysMeta] = useState<{ matched_total: number; truncated: boolean; limit: number } | null>(
    null
  );
  const [putmFullScreen, setPutmFullScreen] = useState(false);
  const [putmPage, setPutmPage] = useState(1);
  const [putmPageSize, setPutmPageSize] = useState(200);

  const putmProcessOptions = useMemo(() => {
    if (!putmKeys) return ["ALL"];
    const s = new Set<string>();
    for (const r of putmKeys) {
      const p0 = (r.process_name || "").trim().toUpperCase();
      if (p0) {
        s.add(p0);
        continue;
      }
      const names = (r.process_names || [])
        .map((x) => String(x || "").trim().toUpperCase())
        .filter(Boolean);
      if (names.length === 1) s.add(names[0]!);
    }
    return ["ALL", ...Array.from(s).sort((a, b) => a.localeCompare(b))];
  }, [putmKeys]);

  useEffect(() => {
    if (!putmKeys) return;
    const valid = new Set(putmProcessOptions);
    if (putmProcess !== "ALL" && !valid.has(putmProcess)) {
      setPutmProcess("ALL");
    }
  }, [putmKeys, putmProcessOptions, putmProcess]);

  function putmRowMatchesProcess(
    filterPn: string,
    primary: string | null | undefined,
    processNames: string[] | null | undefined
  ): boolean {
    const f = (filterPn || "ALL").trim().toUpperCase();
    if (f === "ALL" || f === "") return true;
    const pr = (primary || "").trim().toUpperCase();
    if (pr === f) return true;
    const names = (processNames || [])
      .map((x) => String(x || "").trim().toUpperCase())
      .filter(Boolean);
    if (!pr && names.length === 1 && names[0] === f) return true;
    return false;
  }

  const filteredPutmKeys = useMemo(() => {
    if (!putmKeys) return [];
    const pn = (putmProcess || "ALL").trim().toUpperCase();
    const q = putmSearch.trim().toLowerCase();
    return putmKeys.filter((r) => {
      if (!putmRowMatchesProcess(pn, r.process_name, r.process_names)) return false;
      if (q) {
        const hay = `${r.excel_key} ${r.json_key} ${r.role ?? ""} ${r.process_name ?? ""} ${(r.process_names || []).join(" ")} ${r.description ?? ""} ${r.example ?? ""}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }, [putmKeys, putmProcess, putmSearch]);

  useEffect(() => {
    setPutmPage(1);
  }, [putmProcess, putmSearch, putmPageSize, putmKeys]);

  const putmTotalPages = useMemo(() => {
    const n = filteredPutmKeys.length;
    const ps = Math.max(1, putmPageSize || 1);
    return Math.max(1, Math.ceil(n / ps));
  }, [filteredPutmKeys.length, putmPageSize]);

  useEffect(() => {
    setPutmPage((p) => Math.min(Math.max(1, p), putmTotalPages));
  }, [putmTotalPages]);

  const pagedPutmKeys = useMemo(() => {
    const ps = Math.max(1, putmPageSize || 1);
    const start = (Math.max(1, putmPage) - 1) * ps;
    return filteredPutmKeys.slice(start, start + ps);
  }, [filteredPutmKeys, putmPage, putmPageSize]);

  useEffect(() => {
    if (!putmFullScreen) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") setPutmFullScreen(false);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [putmFullScreen]);

  async function loadPutmKeys() {
    setPutmKeysErr(null);
    try {
      await withLoading(async () => {
        // Load full catalog once; process + text filters are applied client-side.
        const q = new URLSearchParams();
        q.set("limit", "50000");
        const res = await fetch(`${apiJoin("/api/llm_mapping/references/putm-keys")}?${q.toString()}`);
        const data = await readJsonOrThrow(res);
        const rows = (data as any)?.rows;
        if (!Array.isArray(rows)) {
          throw new Error("PUTM keys response invalid (expected rows[]).");
        }
        setPutmKeys(rows as PutmKeyRow[]);
        setPutmKeysMeta({
          matched_total: Number((data as any)?.matched_total ?? rows.length) || rows.length,
          truncated: Boolean((data as any)?.truncated ?? false),
          limit: Number((data as any)?.limit ?? 0) || 0,
        });
      });
    } catch (e) {
      setPutmKeysErr(e instanceof Error ? e.message : String(e));
    }
  }

  async function onPickFile(kind: "det" | "hyb" | "fp", file: File | null) {
    if (!file) return;

    const preview = await withLoading(async () => buildPreview(file, 26));
    if (kind === "det") {
      setDetFile(file);
      setDetPreview(preview);
      setDetRows(null);
      return;
    }
    if (kind === "hyb") {
      setHybFile(file);
      setHybPreview(preview);
      setHybRows(null);
      return;
    }
    setFpFile(file);
    setFpPreview(preview);
    setFpDone(false);
    setFpPipelineAnalysis(null);
    setFpPipelineAnalysisErr(null);
    setFpDbMetrics(null);
  }

  async function runDeterministic() {
    if (!detFile) {
      setGlobalError("Upload a partner file first.");
      return;
    }
    await withLoading(async () => {
      const fd = new FormData();
      fd.append("file", detFile);
      fd.append("client_name", clientName);
      fd.append("process_name", processName);
      fd.append("use_loanparameter_refinement", String(toggles.loanParamRefine));
      fd.append("use_llm_entity_classifier", String(toggles.entityClassifier));

      const res = await fetch(apiJoin("/api/llm_mapping/mapping/deterministic"), { method: "POST", body: fd });
      const data = await readJsonOrThrow(res);
      const mappings = (data as any)?.mappings ?? [];
      setDetRows(Array.isArray(mappings) ? mappings.map(toMatchRow) : []);
      setSchemaOut(null);
      setNestedOut(null);
    });
  }

  async function runHybrid() {
    if (!hybFile) {
      setGlobalError("Upload a partner file first.");
      return;
    }
    await withLoading(async () => {
      const fd = new FormData();
      fd.append("file", hybFile);
      fd.append("client_name", clientName);
      fd.append("process_name", processName);
      fd.append("use_fuzzy", String(toggles.fuzzy));
      fd.append("use_embeddings", String(toggles.embeddings));
      fd.append("use_llm", String(toggles.llm));
      fd.append("use_loanparameter_refinement", String(toggles.loanParamRefine));
      fd.append("use_llm_entity_classifier", String(toggles.entityClassifier));
      if (toggles.saveToDb) {
        fd.append("save_to_db", "true");
        fd.append("master_id", String(masterId));
        fd.append("skip_unmatched", String(toggles.skipUnmatched));
      } else {
        fd.append("save_to_db", "false");
      }

      const res = await fetch(apiJoin("/api/llm_mapping/mapping/hybrid-llm"), { method: "POST", body: fd });
      const data = await readJsonOrThrow(res);
      const mappings = (data as any)?.mappings ?? [];
      const rows = Array.isArray(mappings) ? mappings.map(toMatchRow) : [];
      setHybRows(rows);
      setHybDraftRows(rows);
      setEditSessionId(null);
      setEditSessionStatus(null);
      setEditSessionMsg(null);
      setSchemaOut(null);
      setNestedOut(null);
    });
  }

  function toBackendMappingPayload(rows: MatchRow[]) {
    return rows.map((r) => ({
      partner_field: r.partnerField,
      matched_excel_key: r.excelKey ?? null,
      json_key: r.jsonKey || null,
      entity: r.entity ?? "OTHER",
      confidence: r.confidence,
      match_type: r.matchType,
      reasoning: r.reasoning ?? "",
      needs_review: r.review,
      winning_engine: r.engine,
    }));
  }

  async function saveDraftSession() {
    if (!hybDraftRows || hybDraftRows.length === 0) {
      setGlobalError("Run Hybrid + LLM first (no rows to save).");
      return;
    }
    await withLoading(async () => {
      setEditSessionMsg(null);
      const mappings = toBackendMappingPayload(hybDraftRows);
      if (!editSessionId) {
        const created = await createEditSession({
          client_name: clientName,
          process_name: processName,
          master_id: masterId,
          created_by: "lp-mapping-ui",
          mappings,
        });
        setEditSessionId(created.session_id);
        setEditSessionStatus(created.status);
        setEditSessionMsg(`Draft saved (session: ${created.session_id}).`);
      } else {
        const updated = await updateEditSession(editSessionId, {
          client_name: clientName,
          process_name: processName,
          master_id: masterId,
          updated_by: "lp-mapping-ui",
          note: "UI edit save",
          mappings,
        });
        setEditSessionStatus(updated.status);
        setEditSessionMsg("Draft updated.");
      }
    });
  }

  async function approveDraftAndWriteDb() {
    if (!editSessionId) {
      setGlobalError("Save a draft first (no session_id yet).");
      return;
    }
    await withLoading(async () => {
      setEditSessionMsg(null);
      const approved = await approveEditSession(editSessionId, {
        approved_by: "lp-mapping-ui",
        master_id: masterId,
        skip_unmatched: toggles.skipUnmatched,
      });
      setEditSessionStatus(approved.status);
      const dbRes = (approved.approval_result as any)?.db_result;
      const inserted = dbRes?.inserted ?? dbRes?.inserted_updated ?? dbRes?.insertedUpdated;
      setEditSessionMsg(
        `Approved and written to DB${inserted != null ? ` (inserted/updated: ${inserted})` : ""}.`
      );
    });
  }

  async function runFullPipeline() {
    if (!fpFile) {
      setGlobalError("Upload a partner file first.");
      return;
    }
    await withLoading(async () => {
      const fd = new FormData();
      fd.append("file", fpFile);
      fd.append("client_name", clientName);
      fd.append("process_name", processName);
      fd.append("use_fuzzy", String(toggles.fuzzy));
      fd.append("use_embeddings", String(toggles.embeddings));
      fd.append("use_llm", String(toggles.llm));
      fd.append("use_loanparameter_refinement", String(toggles.loanParamRefine));
      fd.append("use_llm_entity_classifier", String(toggles.entityClassifier));
      fd.append("master_id", String(masterId));
      fd.append("save_to_db", String(toggles.saveToDb));
      fd.append("skip_unmatched", String(toggles.skipUnmatched));
      fd.append("include_build_references", String(toggles.includeRefsZip));

      const res = await fetch(apiJoin("/api/llm_mapping/mapping/full-pipeline"), { method: "POST", body: fd });
      if (!res.ok) {
        const msg = await res.text();
        throw new Error(msg || `${res.status} ${res.statusText}`);
      }
      setFpDbMetrics({
        inserted: res.headers.get("X-DB-Inserted") ?? "—",
        skipped: res.headers.get("X-DB-Skipped") ?? "—",
        errors: res.headers.get("X-DB-Errors") ?? "—",
      });

      const buf = await res.arrayBuffer();
      const blob = new Blob([buf], { type: "application/zip" });

      const disp = res.headers.get("content-disposition") ?? "";
      const match = disp.match(/filename="?([^"]+)"?/i);
      const fallback = `${clientName}_${processName}_outputs.zip`.replace(/\s+/g, "_");
      const name = (match?.[1] || fallback).trim();

      setFpPipelineAnalysis(null);
      setFpPipelineAnalysisErr(null);
      try {
        const parsed = await analyzePipelineZipBlob(blob);
        setFpPipelineAnalysis(parsed);
      } catch (e) {
        setFpPipelineAnalysisErr(e instanceof Error ? e.message : String(e));
      }

      const nextUrl = URL.createObjectURL(blob);
      setFpZipName(name);
      setFpZipUrl((prev) => {
        if (prev) URL.revokeObjectURL(prev);
        return nextUrl;
      });
      setFpDone(true);
    });
  }

  function downloadFullPipelineZip() {
    if (!fpZipUrl) return;
    const a = document.createElement("a");
    a.href = fpZipUrl;
    a.download = fpZipName ?? "outputs.zip";
    document.body.appendChild(a);
    a.click();
    a.remove();
  }

  async function runNested() {
    await withLoading(async () => {
      let payload: any;
      try {
        payload = JSON.parse(nestedJson);
      } catch {
        throw new Error("Nested JSON input is not valid JSON.");
      }
      const res = await fetch(apiJoin("/api/llm_mapping/generate-nested-mapping"), {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const data = await readJsonOrThrow(res);
      const los = (data as any)?.los_json ?? data;
      setNestedOut(typeof los === "string" ? los : JSON.stringify(los, null, 2));
      setSchemaOut(null);
    });
  }

  async function runSchema() {
    await withLoading(async () => {
      let payload: any;
      try {
        payload = JSON.parse(nestedJson);
      } catch {
        throw new Error("Nested JSON input is not valid JSON.");
      }
      const res = await fetch(apiJoin("/api/llm_mapping/generate-schema"), {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const data = await readJsonOrThrow(res);
      const schema = (data as any)?.los_schema ?? data;
      setSchemaOut(typeof schema === "string" ? schema : JSON.stringify(schema, null, 2));
      setNestedOut(null);
    });
  }

  function sendToNested() {
    const rows = hybRows ?? detRows;
    if (!rows) return;

    const payload = {
      client_name: clientName,
      process_name: processName,
      mappings: rows.map((r) => ({
        partner_field: r.partnerField,
        matched_excel_key: r.excelKey ?? "",
        json_key: r.jsonKey || null,
        confidence: r.confidence,
        match_type: r.matchType,
        reasoning: r.reasoning ?? "",
        needs_review: r.review,
        winning_engine: r.engine,
      })),
    };
    setNestedJson(JSON.stringify(payload, null, 2));
    setTab("nested");
  }

  const visibleTableRows = useMemo(() => {
    const rows = tab === "det" ? detRows : tab === "hybrid" ? hybDraftRows : null;
    if (!rows) return null;

    if (tab === "det") {
      return (
        <div className="table-wrap" style={{ marginTop: 0 }}>
          <table className="data-table">
            <thead>
              <tr>
                <th>Partner field</th>
                <th>Excel key</th>
                <th>JSON key</th>
                <th>Confidence</th>
                <th>Match type</th>
                <th>Engine</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.partnerField}>
                  <td style={{ color: "var(--ink)", fontWeight: 500, fontFamily: "inherit" }}>
                    {r.partnerField}
                  </td>
                  <td>{r.excelKey ?? ""}</td>
                  <td style={{ color: "var(--indigo)" }}>{r.jsonKey}</td>
                  <td>
                    <span className={confidenceChipClass(r.confidence)}>
                      {formatPct01(r.confidence)}
                    </span>
                  </td>
                  <td>{r.matchType}</td>
                  <td>
                    <span className="chip chip-blue">{r.engine}</span>
                  </td>
                  <td>
                    {r.review ? (
                      <span style={{ color: "var(--amber)" }}>⚠ Review</span>
                    ) : (
                      <span style={{ color: "var(--emerald)" }}>✓ OK</span>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      );
    }

    return (
      <div className="table-wrap" style={{ marginTop: 0 }}>
        <table className="data-table">
          <thead>
            <tr>
              <th>Partner field</th>
              <th>Excel key</th>
              <th>JSON key</th>
              <th>Entity</th>
              <th>Confidence</th>
              <th>Match type</th>
              <th>Engine</th>
              <th>Review</th>
              <th>Reasoning</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.partnerField}>
                <td style={{ color: "var(--ink)", fontWeight: 500, fontFamily: "inherit" }}>
                  {r.partnerField}
                </td>
                <td>{r.excelKey ?? ""}</td>
                <td style={{ color: "var(--indigo)" }}>
                  {tab === "hybrid" ? (
                    <input
                      className="field-input"
                      style={{ padding: "0.35rem 0.5rem", fontSize: 12, minWidth: 220 }}
                      value={r.jsonKey}
                      onChange={(e) =>
                        setHybDraftRows((prev) => {
                          if (!prev) return prev;
                          return prev.map((x) =>
                            x.partnerField === r.partnerField ? { ...x, jsonKey: e.target.value } : x
                          );
                        })
                      }
                    />
                  ) : (
                    r.jsonKey
                  )}
                </td>
                <td>
                  {tab === "hybrid" ? (
                    <input
                      className="field-input"
                      style={{ padding: "0.35rem 0.5rem", fontSize: 12, minWidth: 120 }}
                      value={r.entity ?? ""}
                      onChange={(e) =>
                        setHybDraftRows((prev) => {
                          if (!prev) return prev;
                          return prev.map((x) =>
                            x.partnerField === r.partnerField ? { ...x, entity: e.target.value } : x
                          );
                        })
                      }
                    />
                  ) : (
                    r.entity ?? ""
                  )}
                </td>
                <td>
                  <span className={confidenceChipClass(r.confidence)}>{formatPct01(r.confidence)}</span>
                </td>
                <td>{r.matchType}</td>
                <td>
                  <span className="chip chip-blue">{r.engine}</span>
                </td>
                <td>
                  {r.review ? (
                    <span style={{ color: "var(--amber)" }}>⚠ Review</span>
                  ) : (
                    <span style={{ color: "var(--emerald)" }}>✓</span>
                  )}
                </td>
                <td
                  style={{
                    color: "var(--ink-3)",
                    fontSize: 11,
                    maxWidth: 160,
                    whiteSpace: "normal",
                    fontFamily: "inherit",
                  }}
                >
                  {r.reasoning ?? `Matched via ${r.engine} engine with ${formatPct01(r.confidence)} confidence`}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }, [detRows, hybDraftRows, tab]);

  return (
    <>
      <style>{STYLE}</style>

      <h2 className="sr-only">
        LP Field Mapping Test Console — interactive dashboard with five tabs
      </h2>

      <div className="shell">
        <aside className="sidebar">
          <div className="sb-logo">
            <div className="sb-logo-mark">◈</div>
            <div className="sb-brand">LP FIELD MAPPING</div>

          </div>

          <div className="sb-section">
            <div className="sb-section-label">Navigation</div>
            <button
              className={`nav-item ${tab === "refs" ? "active" : ""}`}
              onClick={() => setTab("refs")}
              type="button"
            >
              <div className="nav-dot nd-indigo">📂</div> References
            </button>
            <button
              className={`nav-item ${tab === "det" ? "active" : ""}`}
              onClick={() => setTab("det")}
              type="button"
            >
              <div className="nav-dot nd-violet">🔍</div> Deterministic
            </button>
            <button
              className={`nav-item ${tab === "hybrid" ? "active" : ""}`}
              onClick={() => setTab("hybrid")}
              type="button"
            >
              <div className="nav-dot nd-cyan">⚡</div> Hybrid + LLM
            </button>
            <button
              className={`nav-item ${tab === "pipeline" ? "active" : ""}`}
              onClick={() => setTab("pipeline")}
              type="button"
            >
              <div className="nav-dot nd-emerald">🚀</div> Full Pipeline
            </button>
            <button
              className={`nav-item ${tab === "nested" ? "active" : ""}`}
              onClick={() => setTab("nested")}
              type="button"
            >
              <div className="nav-dot nd-amber">🧩</div> Nested / Schema
            </button>
          </div>

          <div className="sb-divider"></div>

          <div className="sb-section" style={{ marginBottom: "0.5rem" }}>
            <div className="sb-section-label">Config</div>
            <div className="sb-fields">
              <div>
                <span className="sb-field-label">Client name</span>
                <input className="sb-input" value={clientName} onChange={(e) => setClientName(e.target.value)} />
              </div>
              <div>
                <span className="sb-field-label">Process name</span>
                <input className="sb-input" value={processName} onChange={(e) => setProcessName(e.target.value)} />
              </div>
              <div>
                <span className="sb-field-label">Master ID</span>
                <input
                  className="sb-input"
                  type="number"
                  value={masterId}
                  min={1}
                  onChange={(e) => setMasterId(Number(e.target.value || 1))}
                />
              </div>
              <div>
                <span className="sb-field-label">API base URL</span>
                <input
                  className="sb-input"
                  placeholder="Leave empty → proxy"
                  value={apiBaseUrl}
                  onChange={(e) => setApiBaseUrl(e.target.value)}
                />
              </div>
            </div>
          </div>

          <div className="sb-divider"></div>

          <div className="sb-section" style={{ marginBottom: "0.5rem" }}>
            <div className="sb-section-label">Engines</div>
            <div className="sb-toggles">
              <div className="toggle-row" onClick={() => setToggles((p) => ({ ...p, fuzzy: !p.fuzzy }))}>
                <span>Fuzzy matching</span>
                <div className={`pill-toggle ${toggles.fuzzy ? "on" : "off"}`} />
              </div>
              <div className="toggle-row" onClick={() => setToggles((p) => ({ ...p, embeddings: !p.embeddings }))}>
                <span>Embeddings</span>
                <div className={`pill-toggle ${toggles.embeddings ? "on" : "off"}`} />
              </div>
              <div className="toggle-row" onClick={() => setToggles((p) => ({ ...p, llm: !p.llm }))}>
                <span>LLM</span>
                <div className={`pill-toggle ${toggles.llm ? "on" : "off"}`} />
              </div>
              <div className="toggle-row" onClick={() => setToggles((p) => ({ ...p, loanParamRefine: !p.loanParamRefine }))}>
                <span>Loan param refinement</span>
                <div className={`pill-toggle ${toggles.loanParamRefine ? "on" : "off"}`} />
              </div>
              <div className="toggle-row" onClick={() => setToggles((p) => ({ ...p, entityClassifier: !p.entityClassifier }))}>
                <span>Entity classifier</span>
                <div className={`pill-toggle ${toggles.entityClassifier ? "on" : "off"}`} />
              </div>
            </div>
          </div>

          <div className="sb-divider"></div>

          <div className="sb-section">
            <div className="sb-section-label">Options</div>
            <div className="sb-toggles">
              <div className="toggle-row" onClick={() => setToggles((p) => ({ ...p, includeRefsZip: !p.includeRefsZip }))}>
                <span>Include refs in ZIP</span>
                <div className={`pill-toggle ${toggles.includeRefsZip ? "on" : "off"}`} />
              </div>
              <div className="toggle-row" onClick={() => setToggles((p) => ({ ...p, saveToDb: !p.saveToDb }))}>
                <span>Save to DB</span>
                <div className={`pill-toggle ${toggles.saveToDb ? "on" : "off"}`} />
              </div>
              <div className="toggle-row" onClick={() => setToggles((p) => ({ ...p, skipUnmatched: !p.skipUnmatched }))}>
                <span>Skip unmatched</span>
                <div className={`pill-toggle ${toggles.skipUnmatched ? "on" : "off"}`} />
              </div>
            </div>
          </div>

          <div style={{ flex: 1 }}></div>

          <div style={{ padding: "0 1rem" }}>
            <div className="sb-status-row">
              <div className={health.online ? "status-dot-on" : "status-dot-off"}></div>
              <span className="status-text" style={{ color: health.online ? "#34D399" : undefined }}>
                {health.text}
              </span>
              <button
                onClick={() => void pingHealth()}
                type="button"
                style={{
                  marginLeft: "auto",
                  background: "rgba(255,255,255,0.08)",
                  border: "1px solid rgba(255,255,255,0.12)",
                  borderRadius: 6,
                  color: "rgba(255,255,255,0.55)",
                  fontSize: 11,
                  padding: "3px 8px",
                  cursor: "pointer",
                  fontFamily: "inherit",
                }}
              >
                Ping
              </button>
            </div>
          </div>
        </aside>

        <main className="main">
          <div className="topbar">
            <div className="topbar-left">
              <div className="page-title">{pageMeta.title}</div>
              <div className="page-crumb">{pageMeta.crumb}</div>
            </div>
            <div className="topbar-right">
              <span className={`badge ${pageMeta.badgeClass}`}>
                <span className="badge-dot"></span> {pageMeta.badgeTxt}
              </span>
              <div style={{ width: 1, height: 24, background: "var(--border)" }}></div>
              <div
                style={{
                  width: 32,
                  height: 32,
                  borderRadius: "50%",
                  background: "linear-gradient(135deg,#4F46E5,#7C3AED)",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  color: "white",
                  fontSize: 13,
                  fontWeight: 500,
                }}
              >
                DG
              </div>
            </div>
          </div>

          <div className="content">
            {loading && (
              <div>
                <div className="processing-badge">
                  <div className="proc-dot"></div> Processing request…
                </div>
                <div className="loading-strip">
                  <div className="loading-fill"></div>
                </div>
              </div>
            )}

            {globalError && (
              <div className="alert alert-error">
                <span>⚠</span>
                <span style={{ flex: 1 }}>{globalError}</span>
                <button className="alert-dismiss" onClick={() => setGlobalError(null)} type="button">
                  ×
                </button>
              </div>
            )}

            {tab === "refs" && (
              <div>
                <div className="metric-row" style={{ gridTemplateColumns: "repeat(3,1fr)" }}>
                  <div className="metric-card">
                    <div className="metric-accent acc-indigo"></div>
                    <div className="metric-label">Reference files</div>
                    <div className="metric-value mv-indigo">4</div>
                    <div className="metric-sub">On-disk dictionaries</div>
                  </div>
                  <div className="metric-card">
                    <div className="metric-accent acc-emerald"></div>
                    <div className="metric-label">Total keys</div>
                    <div className="metric-value mv-emerald">8.4K</div>
                    <div className="metric-sub">Alias entries indexed</div>
                  </div>
                  <div className="metric-card">
                    <div className="metric-accent acc-amber"></div>
                    <div className="metric-label">Last rebuilt</div>
                    <div className="metric-value" style={{ color: "var(--amber)", fontSize: 20, paddingTop: 4 }}>
                      Today
                    </div>
                    <div className="metric-sub">14:32 IST</div>
                  </div>
                </div>

                <div className="two-col">
                  <div className="panel">
                    <div className="panel-header">
                      <div>
                        <div className="panel-title">Status check</div>
                        <div className="panel-desc">Inspect reference files on disk</div>
                      </div>
                      <button className="btn btn-primary" onClick={() => void runRefStatus()} disabled={loading} type="button">
                        ◉ Check status
                      </button>
                    </div>
                    <div className="panel-body">
                      {refStatus ? (
                        <div className="refs-status">{refStatus}</div>
                      ) : (
                        <div style={{ color: "var(--ink-3)", fontSize: 13 }}>
                          Click "Check status" to inspect reference files.
                        </div>
                      )}
                    </div>
                  </div>

                  <div className="panel">
                    <div className="panel-header">
                      <div>
                        <div className="panel-title">Build references</div>
                        <div className="panel-desc">Rebuild dictionaries from database</div>
                      </div>
                      <button className="btn btn-secondary" onClick={() => void runBuild()} disabled={loading} type="button">
                        ⟳ Build
                      </button>
                    </div>
                    <div className="panel-body">
                      <div className="field">
                        <label className="field-label">PUTM table override</label>
                        <input
                          className="field-input"
                          placeholder="Optional — use default table"
                          value={refsPutmOverride}
                          onChange={(e) => setRefsPutmOverride(e.target.value)}
                        />
                      </div>
                      <div className="field">
                        <label className="field-label">Mapping table override</label>
                        <input
                          className="field-input"
                          placeholder="Optional — use default table"
                          value={refsMappingOverride}
                          onChange={(e) => setRefsMappingOverride(e.target.value)}
                        />
                      </div>
                      {buildOut && <div className="pre-block">{buildOut}</div>}
                    </div>
                  </div>
                </div>

                <div className="panel" style={{ marginTop: "1.25rem" }}>
                  <div className="panel-header">
                    <div>
                      <div className="panel-title">PUTM key search</div>
                      <div className="panel-desc">Search available `excel_key` and its corresponding `json_key`</div>
                    </div>
                    <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                      <button
                        className="btn btn-secondary"
                        onClick={() => setPutmFullScreen(true)}
                        disabled={!putmKeys}
                        type="button"
                        title="Full screen key list"
                      >
                        ⛶ Full screen
                      </button>
                      <button
                        className="btn btn-primary"
                        onClick={() => void loadPutmKeys()}
                        disabled={loading}
                        type="button"
                        title="Loads keys from references/field_dictionary.json"
                      >
                        ⟳ Load keys
                      </button>
                    </div>
                  </div>
                  <div className="panel-body">
                    {putmKeysErr && (
                      <div className="alert alert-error" style={{ marginBottom: "0.75rem" }}>
                        <span>⚠</span>
                        <span style={{ flex: 1 }}>{putmKeysErr}</span>
                        <button className="alert-dismiss" onClick={() => setPutmKeysErr(null)} type="button">
                          ×
                        </button>
                      </div>
                    )}

                    <div style={{ display: "flex", flexWrap: "wrap", gap: "0.75rem", alignItems: "flex-end" }}>
                      <div className="field" style={{ marginBottom: 0, minWidth: 220 }}>
                        <label className="field-label">Process filter</label>
                        <select
                          className="field-input"
                          value={putmProcess}
                          onChange={(e) => setPutmProcess(e.target.value)}
                        >
                          {putmProcessOptions.map((p) => (
                            <option key={p} value={p}>
                              {p}
                            </option>
                          ))}
                        </select>
                      </div>
                      <div className="field" style={{ marginBottom: 0, flex: "1 1 360px" }}>
                        <label className="field-label">Search (excel/json/role/description)</label>
                        <input
                          className="field-input"
                          value={putmSearch}
                          onChange={(e) => setPutmSearch(e.target.value)}
                          placeholder="Type any key…"
                        />
                      </div>
                      <button
                        className="btn btn-secondary"
                        type="button"
                        disabled={!putmKeys || filteredPutmKeys.length === 0}
                        onClick={() =>
                          downloadCsv(
                            `putm_keys_${(putmProcess || "ALL").toLowerCase()}${putmSearch.trim() ? "_filtered" : ""}.csv`,
                            ["process_name", "excel_key", "json_key", "role", "description", "example"],
                            filteredPutmKeys.map((r) => ({
                              process_name: String(r.process_name ?? ""),
                              excel_key: String(r.excel_key ?? ""),
                              json_key: String(r.json_key ?? ""),
                              role: String(r.role ?? ""),
                              description: String(r.description ?? ""),
                              example: String(r.example ?? ""),
                            }))
                          )
                        }
                      >
                        ⬇ Export CSV
                      </button>
                    </div>

                    <div style={{ fontSize: 12, color: "var(--ink-3)", margin: "0.75rem 0 0.5rem" }}>
                      {putmKeys
                        ? `Showing ${pagedPutmKeys.length} key(s) on this page · ${filteredPutmKeys.length} filtered key(s)${putmKeysMeta?.truncated ? ` (server returned ${putmKeys.length}/${putmKeysMeta.matched_total} due to limit=${putmKeysMeta.limit})` : ""}.`
                        : "Click “Load keys” to fetch from the server (requires references built)."}
                    </div>

                    {putmKeys && (
                      <>
                        <div
                          style={{
                            display: "flex",
                            flexWrap: "wrap",
                            gap: "0.75rem",
                            alignItems: "center",
                            justifyContent: "space-between",
                            marginBottom: "0.75rem",
                          }}
                        >
                          <div style={{ display: "flex", flexWrap: "wrap", gap: "0.5rem", alignItems: "center" }}>
                            <button type="button" className="btn btn-secondary" onClick={() => setPutmPage(1)} disabled={putmPage <= 1}>
                              ⏮ First
                            </button>
                            <button
                              type="button"
                              className="btn btn-secondary"
                              onClick={() => setPutmPage((p) => Math.max(1, p - 1))}
                              disabled={putmPage <= 1}
                            >
                              ← Prev
                            </button>
                            <span style={{ fontSize: 12, color: "var(--ink-2)" }}>
                              Page <b>{putmPage}</b> / {putmTotalPages}
                            </span>
                            <button
                              type="button"
                              className="btn btn-secondary"
                              onClick={() => setPutmPage((p) => Math.min(putmTotalPages, p + 1))}
                              disabled={putmPage >= putmTotalPages}
                            >
                              Next →
                            </button>
                            <button
                              type="button"
                              className="btn btn-secondary"
                              onClick={() => setPutmPage(putmTotalPages)}
                              disabled={putmPage >= putmTotalPages}
                            >
                              Last ⏭
                            </button>
                          </div>
                          <div className="field" style={{ marginBottom: 0, minWidth: 170 }}>
                            <label className="field-label">Rows per page</label>
                            <select className="field-input" value={String(putmPageSize)} onChange={(e) => setPutmPageSize(Number(e.target.value))}>
                              <option value="50">50</option>
                              <option value="100">100</option>
                              <option value="200">200</option>
                              <option value="500">500</option>
                              <option value="1000">1000</option>
                            </select>
                          </div>
                        </div>

                        <div className="table-wrap" style={{ maxHeight: 520 }}>
                          <table className="data-table" style={{ fontSize: 13 }}>
                            <thead>
                              <tr>
                                <th style={{ width: 140 }}>Process</th>
                                <th style={{ width: 340 }}>Excel key</th>
                                <th style={{ width: 420 }}>JSON key</th>
                                <th style={{ width: 110 }}>Role</th>
                              </tr>
                            </thead>
                            <tbody>
                              {pagedPutmKeys.map((r, i) => (
                                <tr key={`${r.excel_key}__${r.json_key}__${i}`}>
                                  <td>{r.process_name ?? ""}</td>
                                  <td style={{ color: "var(--ink)" }}>{r.excel_key}</td>
                                  <td style={{ color: "var(--indigo)" }}>{r.json_key}</td>
                                  <td>{r.role ?? ""}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </>
                    )}
                  </div>
                </div>

                {putmFullScreen && (
                  <div
                    role="dialog"
                    aria-modal="true"
                    style={{
                      position: "fixed",
                      inset: 0,
                      background: "rgba(15, 23, 42, 0.55)",
                      zIndex: 9999,
                      padding: 16,
                    }}
                    onMouseDown={() => setPutmFullScreen(false)}
                  >
                    <div
                      style={{
                        height: "100%",
                        maxWidth: "calc(100vw - 32px)",
                        margin: "0 auto",
                        background: "white",
                        borderRadius: 12,
                        border: "1px solid var(--border)",
                        display: "flex",
                        flexDirection: "column",
                        overflow: "hidden",
                      }}
                      onMouseDown={(e) => e.stopPropagation()}
                    >
                      <div
                        style={{
                          padding: "14px 16px",
                          borderBottom: "1px solid var(--border)",
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "space-between",
                          gap: 12,
                        }}
                      >
                        <div style={{ fontSize: 14, fontWeight: 600, color: "var(--ink)" }}>PUTM key search · Full screen</div>
                        <button type="button" className="btn btn-secondary" onClick={() => setPutmFullScreen(false)}>
                          Close (Esc)
                        </button>
                      </div>
                      <div style={{ padding: 16, overflow: "auto", flex: 1 }}>
                        {/* reuse same controls/table */}
                        <div style={{ marginBottom: "0.75rem" }}>
                          <div style={{ display: "flex", flexWrap: "wrap", gap: "0.75rem", alignItems: "flex-end" }}>
                            <div className="field" style={{ marginBottom: 0, minWidth: 220 }}>
                              <label className="field-label">Process filter</label>
                              <select className="field-input" value={putmProcess} onChange={(e) => setPutmProcess(e.target.value)}>
                                {putmProcessOptions.map((p) => (
                                  <option key={p} value={p}>
                                    {p}
                                  </option>
                                ))}
                              </select>
                            </div>
                            <div className="field" style={{ marginBottom: 0, flex: "1 1 360px" }}>
                              <label className="field-label">Search (excel/json/role/description)</label>
                              <input className="field-input" value={putmSearch} onChange={(e) => setPutmSearch(e.target.value)} placeholder="Type any key…" />
                            </div>
                            <button className="btn btn-primary" onClick={() => void loadPutmKeys()} disabled={loading} type="button">
                              ⟳ Reload
                            </button>
                          </div>
                        </div>
                        <div className="table-wrap" style={{ maxHeight: "calc(100vh - 260px)" }}>
                          <table className="data-table" style={{ fontSize: 15 }}>
                            <thead>
                              <tr>
                                <th style={{ width: 160 }}>Process</th>
                                <th style={{ width: 420 }}>Excel key</th>
                                <th style={{ width: 520 }}>JSON key</th>
                                <th style={{ width: 120 }}>Role</th>
                              </tr>
                            </thead>
                            <tbody>
                              {pagedPutmKeys.map((r, i) => (
                                <tr key={`${r.excel_key}__${r.json_key}__fs__${i}`}>
                                  <td>{r.process_name ?? ""}</td>
                                  <td style={{ color: "var(--ink)" }}>{r.excel_key}</td>
                                  <td style={{ color: "var(--indigo)" }}>{r.json_key}</td>
                                  <td>{r.role ?? ""}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            )}

            {tab === "det" && (
              <div className="panel">
                <div className="panel-header">
                  <div>
                    <div className="panel-title">Deterministic matching</div>
                    <div className="panel-desc">Phase 1 — alias and rule-based matching. Fast, zero LLM cost.</div>
                  </div>
                  <span className="badge badge-violet">
                    <span className="badge-dot"></span> Phase 1
                  </span>
                </div>
                <div className="panel-body">
                  <div className="upload-zone" onClick={() => detInputRef.current?.click()}>
                    <input
                      ref={detInputRef}
                      type="file"
                      accept=".xlsx,.xls,.csv"
                      onChange={(e) => void onPickFile("det", e.target.files?.[0] ?? null)}
                    />
                    <div className="upload-icon">📎</div>
                    <div className="upload-title">{detFile ? detFile.name : "Drop partner Excel / CSV here"}</div>
                    <div className="upload-hint">Supports .xlsx, .xls, .csv</div>
                  </div>

                  <div className="action-bar">
                    <button className="btn btn-primary" onClick={() => void runDeterministic()} disabled={loading} type="button">
                      ▶ Run deterministic
                    </button>
                    {detRows && (
                      <button className="btn btn-ghost" onClick={sendToNested} disabled={loading} type="button">
                        → Send to Nested
                      </button>
                    )}
                  </div>

                  <PreviewPanel preview={detPreview} />

                  {detRows && (
                    <div style={{ marginTop: "1rem" }}>
                      <EngineBreakdown rows={detRows} />
                      {visibleTableRows}
                    </div>
                  )}
                </div>
              </div>
            )}

            {tab === "hybrid" && (
              <div className="panel">
                <div className="panel-header">
                  <div>
                    <div className="panel-title">Hybrid + LLM</div>
                    <div className="panel-desc">Full cascade: Deterministic → Fuzzy → Embeddings → LLM</div>
                  </div>
                  <span className="badge badge-cyan">
                    <span className="badge-dot"></span> Full Stack
                  </span>
                </div>
                <div className="panel-body">
                  <div className="upload-zone" onClick={() => hybInputRef.current?.click()}>
                    <input
                      ref={hybInputRef}
                      type="file"
                      accept=".xlsx,.xls,.csv"
                      onChange={(e) => void onPickFile("hyb", e.target.files?.[0] ?? null)}
                    />
                    <div className="upload-icon">⚡</div>
                    <div className="upload-title">{hybFile ? hybFile.name : "Drop partner file here"}</div>
                    <div className="upload-hint">Supports .xlsx, .xls, .csv</div>
                  </div>

                  <div className="action-bar">
                    <button className="btn btn-primary" onClick={() => void runHybrid()} disabled={loading} type="button">
                      ▶ Run hybrid + LLM
                    </button>
                    <button
                      className="btn btn-secondary"
                      onClick={() => void saveDraftSession()}
                      disabled={loading || !hybDraftRows?.length}
                      type="button"
                      title="Saves editable draft (does not write to DB)"
                    >
                      💾 Save draft
                    </button>
                    <button
                      className="btn btn-primary"
                      onClick={() => void approveDraftAndWriteDb()}
                      disabled={loading || !editSessionId}
                      type="button"
                      title="Approves the draft and writes mappings to target DB"
                    >
                      ✓ Approve & write DB
                    </button>
                    {hybRows && (
                      <button className="btn btn-ghost" onClick={sendToNested} disabled={loading} type="button">
                        → Send to Nested
                      </button>
                    )}
                  </div>

                  <PreviewPanel preview={hybPreview} />

                  {editSessionMsg && (
                    <div className="alert alert-success" style={{ marginTop: "0.75rem" }}>
                      <span>✓</span>
                      <span style={{ flex: 1 }}>
                        {editSessionMsg}
                        {editSessionId ? ` (session_id=${editSessionId}${editSessionStatus ? `, status=${editSessionStatus}` : ""})` : ""}
                      </span>
                      <button className="alert-dismiss" onClick={() => setEditSessionMsg(null)} type="button">
                        ×
                      </button>
                    </div>
                  )}

                  {hybDraftRows && (
                    <div style={{ marginTop: "1rem" }}>
                      <EngineBreakdown rows={hybDraftRows} />
                      {visibleTableRows}
                    </div>
                  )}
                </div>
              </div>
            )}

            {tab === "pipeline" && (
              <div className="panel">
                <div className="panel-header">
                  <div>
                    <div className="panel-title">Full pipeline</div>
                    <div className="panel-desc">
                      Runs all phases and outputs a ZIP — Excel, nested JSON, schema, optional reference files
                    </div>
                  </div>
                  <span className="badge badge-emerald">
                    <span className="badge-dot"></span> ZIP Output
                  </span>
                </div>
                <div className="panel-body">
                  <div className="upload-zone" onClick={() => fpInputRef.current?.click()}>
                    <input
                      ref={fpInputRef}
                      type="file"
                      accept=".xlsx,.xls,.csv"
                      onChange={(e) => void onPickFile("fp", e.target.files?.[0] ?? null)}
                    />
                    <div className="upload-icon">🚀</div>
                    <div className="upload-title">{fpFile ? fpFile.name : "Drop partner file here"}</div>
                    <div className="upload-hint">Supports .xlsx, .xls, .csv</div>
                  </div>

                  <div className="field">
                    <label className="field-label">Sheet filter (optional)</label>
                    <input className="field-input" placeholder="e.g. Sheet1" />
                  </div>

                  <div className="action-bar">
                    <button className="btn btn-primary" onClick={() => void runFullPipeline()} disabled={loading} type="button">
                      ▶ Run full pipeline
                    </button>
                    <button
                      className="btn btn-secondary"
                      onClick={downloadFullPipelineZip}
                      disabled={!fpDone || !fpZipUrl || loading}
                      type="button"
                    >
                      ⬇ Download ZIP
                    </button>
                  </div>

                  <PreviewPanel preview={fpPreview} />

                  {fpDone && (
                    <div style={{ marginTop: "1rem" }}>
                      {fpDbMetrics && (
                        <div className="metric-row" style={{ gridTemplateColumns: "repeat(4,1fr)" }}>
                          <div className="metric-card">
                            <div className="metric-accent acc-emerald"></div>
                            <div className="metric-label">DB inserted</div>
                            <div className="metric-value mv-emerald">{fpDbMetrics.inserted}</div>
                            <div className="metric-sub">New mappings</div>
                          </div>
                          <div className="metric-card">
                            <div className="metric-accent acc-indigo"></div>
                            <div className="metric-label">DB skipped</div>
                            <div className="metric-value mv-indigo">{fpDbMetrics.skipped}</div>
                            <div className="metric-sub">Duplicates / rules</div>
                          </div>
                          <div className="metric-card">
                            <div className="metric-accent acc-amber"></div>
                            <div className="metric-label">DB errors</div>
                            <div className="metric-value mv-amber">{fpDbMetrics.errors}</div>
                            <div className="metric-sub">Write failures</div>
                          </div>
                          <div className="metric-card">
                            <div className="metric-accent acc-violet"></div>
                            <div className="metric-label">Master ID</div>
                            <div className="metric-value" style={{ color: "var(--violet)" }}>
                              {masterId}
                            </div>
                            <div className="metric-sub">Target schema</div>
                          </div>
                        </div>
                      )}
                      <div className="alert alert-success" style={{ marginTop: fpDbMetrics ? undefined : 0 }}>
                        <span>✓</span>
                        <span>
                          Pipeline complete. ZIP ready for download. Open the analysis below for Excel breakdown,
                          filters, and charts.
                        </span>
                      </div>
                      <PipelineOutputAnalysis
                        key={fpZipName ?? "pipeline-output"}
                        analysis={fpPipelineAnalysis}
                        error={fpPipelineAnalysisErr}
                        clientName={clientName}
                        processName={processName}
                        masterId={masterId}
                        skipUnmatched={toggles.skipUnmatched}
                      />
                    </div>
                  )}
                </div>
              </div>
            )}

            {tab === "nested" && (
              <div className="two-col" style={{ alignItems: "start" }}>
                <div className="panel">
                  <div className="panel-header">
                    <div>
                      <div className="panel-title">Nested mapping</div>
                      <div className="panel-desc">Paste JSON with a mappings array</div>
                    </div>
                    <span className="badge badge-amber">
                      <span className="badge-dot"></span> Schema
                    </span>
                  </div>
                  <div className="panel-body">
                    <textarea
                      className="field-textarea"
                      style={{ minHeight: 220 }}
                      value={nestedJson}
                      onChange={(e) => setNestedJson(e.target.value)}
                      spellCheck={false}
                    />
                    <div className="action-bar" style={{ marginBottom: 0 }}>
                      <button className="btn btn-primary" onClick={() => void runNested()} disabled={loading} type="button">
                        ⟳ Generate nested JSON
                      </button>
                      <button className="btn btn-secondary" onClick={() => void runSchema()} disabled={loading} type="button">
                        ⟳ Generate schema
                      </button>
                      <button className="btn btn-ghost" onClick={() => setNestedJson(DEFAULT_NESTED_INPUT)} disabled={loading} type="button">
                        Reset
                      </button>
                    </div>
                  </div>
                </div>

                <div className="stack-gap">
                  {nestedOut && (
                    <div className="panel">
                      <div className="panel-header">
                        <div className="panel-title">Nested JSON</div>
                      </div>
                      <div className="panel-body" style={{ paddingTop: 0 }}>
                        <div className="pre-block" style={{ marginTop: "0.75rem" }}>
                          {nestedOut}
                        </div>
                      </div>
                    </div>
                  )}

                  {schemaOut && (
                    <div className="panel">
                      <div className="panel-header">
                        <div className="panel-title">Schema JSON</div>
                      </div>
                      <div className="panel-body" style={{ paddingTop: 0 }}>
                        <div className="pre-block" style={{ marginTop: "0.75rem" }}>
                          {schemaOut}
                        </div>
                      </div>
                    </div>
                  )}

                  {!nestedOut && !schemaOut && (
                    <div
                      style={{
                        background: "white",
                        border: "1px solid var(--border)",
                        borderRadius: "var(--radius)",
                        padding: "2rem",
                        textAlign: "center",
                        color: "var(--ink-3)",
                        fontSize: 13,
                      }}
                    >
                      <div style={{ fontSize: 28, marginBottom: "0.75rem" }}>🧩</div>
                      Run "Generate nested JSON" or "Generate schema" to see output here.
                      {(detRows || hybRows) && (
                        <div style={{ marginTop: "0.75rem" }}>
                          <button className="btn btn-primary" onClick={sendToNested} type="button">
                            ⇢ Load last mapping result
                          </button>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
        </main>
      </div>
    </>
  );
}

