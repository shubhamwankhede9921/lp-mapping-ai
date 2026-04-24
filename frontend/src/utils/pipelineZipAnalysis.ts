import JSZip from "jszip";
import * as XLSX from "xlsx";

export type ExcelKeyKind = "unmatched" | "loan" | "other";

export function classifyExcelKey(mappedKey: string): ExcelKeyKind {
  const t = mappedKey.trim();
  if (!t) return "unmatched";
  const u = t.toUpperCase();
  if (u === "LOANPARAMETER" || /^LOANPARAMETER\d+$/.test(u)) return "loan";
  return "other";
}

function normalizeHeaderKey(h: string): string {
  return h.trim().toLowerCase().replace(/\s+/g, " ");
}

export function pickColumn(row: Record<string, string>, aliases: string[]): string {
  if (!row) return "";
  const keys = Object.keys(row);
  const lowerMap = new Map(keys.map((k) => [k.trim().toLowerCase().replace(/\s+/g, " "), k]));
  for (const a of aliases) {
    const n = normalizeHeaderKey(a);
    const orig = lowerMap.get(n);
    if (orig !== undefined) {
      const v = row[orig];
      if (v !== undefined && v !== null) return String(v).trim();
    }
  }
  const aliasFragments = aliases.map((a) => normalizeHeaderKey(a).replace(/ /g, ""));
  for (const frag of aliasFragments) {
    for (const k of keys) {
      const kn = k.trim().toLowerCase().replace(/\s+/g, "");
      if (kn === frag) {
        const v = row[k];
        if (v !== undefined && v !== null) return String(v).trim();
      }
    }
  }
  return "";
}

/** Normalize to 0..1 when possible (handles 0.92 and 92). */
export function parseConfidence(raw: string): number {
  if (raw === undefined || raw === null || raw === "") return NaN;
  const n = parseFloat(String(raw).replace(/,/g, ""));
  if (!Number.isFinite(n)) return NaN;
  if (n > 1 && n <= 100) return Math.min(1, Math.max(0, n / 100));
  return Math.min(1, Math.max(0, n));
}

export type PipelineAnalysisStats = {
  total: number;
  matched: number;
  unmatched: number;
  loanSlots: number;
  otherKeys: number;
  avgConfidence: number;
  confidenceRowCount: number;
  highConfidence: number;
  reviewCount: number;
  matchTypeCounts: Record<string, number>;
  confidenceBuckets: { label: string; count: number; fill: string }[];
};

export type PipelineAnalysis = {
  excelFileName: string;
  sheetName: string;
  headers: string[];
  rows: Record<string, string>[];
  stats: PipelineAnalysisStats;
  zipEntryNames: string[];
};

function pickSheetName(sheetNames: string[]): string {
  const exact = sheetNames.find((s) => s === "Field Mapping");
  if (exact) return exact;
  const ci = sheetNames.find((s) => s.toLowerCase() === "field mapping");
  if (ci) return ci;
  return sheetNames[0] ?? "";
}

function rowObjectsFromSheet(ws: XLSX.WorkSheet): { headers: string[]; rows: Record<string, string>[] } {
  const raw = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false }) as Record<string, unknown>[];
  if (raw.length === 0) return { headers: [], rows: [] };
  const headers = Object.keys(raw[0] ?? {}).map((h) => String(h));
  const rows = raw.map((r) => {
    const o: Record<string, string> = {};
    for (const h of headers) {
      const v = r[h];
      o[h] = v === undefined || v === null ? "" : String(v);
    }
    return o;
  });
  return { headers, rows };
}

const EXCEL_KEY_ALIASES = ["Mapped Excel Key", "matched_excel_key"];
const CONF_ALIASES = ["Confidence", "confidence"];
const REVIEW_ALIASES = ["Needs Review", "needs_review"];
const MATCH_TYPE_ALIASES = ["Match Type", "match_type"];

function computeStats(rows: Record<string, string>[]): PipelineAnalysisStats {
  let matched = 0;
  let unmatched = 0;
  let loanSlots = 0;
  let otherKeys = 0;
  let confSum = 0;
  let confN = 0;
  let highConfidence = 0;
  let reviewCount = 0;
  const matchTypeCounts: Record<string, number> = {};
  const buckets = [
    { label: "0–50%", max: 0.5, count: 0, fill: "#F8CBAD" },
    { label: "50–70%", max: 0.7, count: 0, fill: "#FED8B1" },
    { label: "70–85%", max: 0.85, count: 0, fill: "#FFEB9C" },
    { label: "85–100%", max: 1.01, count: 0, fill: "#C6EFCE" },
  ];

  for (const row of rows) {
    const ek = pickColumn(row, EXCEL_KEY_ALIASES);
    const kind = classifyExcelKey(ek);
    if (kind === "unmatched") unmatched += 1;
    else {
      matched += 1;
      if (kind === "loan") loanSlots += 1;
      else otherKeys += 1;
    }

    const c = parseConfidence(pickColumn(row, CONF_ALIASES));
    if (!Number.isNaN(c)) {
      confSum += c;
      confN += 1;
      if (c >= 0.85) highConfidence += 1;
      if (c < 0.5) buckets[0].count += 1;
      else if (c < 0.7) buckets[1].count += 1;
      else if (c < 0.85) buckets[2].count += 1;
      else buckets[3].count += 1;
    }

    const rev = pickColumn(row, REVIEW_ALIASES).toUpperCase();
    if (rev === "YES" || rev === "TRUE" || rev === "1") reviewCount += 1;

    const mt = pickColumn(row, MATCH_TYPE_ALIASES) || "(empty)";
    matchTypeCounts[mt] = (matchTypeCounts[mt] ?? 0) + 1;
  }

  const total = rows.length;
  return {
    total,
    matched,
    unmatched,
    loanSlots,
    otherKeys,
    avgConfidence: confN ? confSum / confN : 0,
    confidenceRowCount: confN,
    highConfidence,
    reviewCount,
    matchTypeCounts,
    confidenceBuckets: buckets.map(({ label, count, fill }) => ({ label, count, fill })),
  };
}

export async function analyzePipelineZipBlob(blob: Blob): Promise<PipelineAnalysis> {
  const zip = await JSZip.loadAsync(blob);
  const zipEntryNames = Object.keys(zip.files)
    .filter((n) => !zip.files[n].dir && !n.includes("__MACOSX"))
    .sort();

  const xlsxNames = zipEntryNames.filter(
    (n) =>
      n.toLowerCase().endsWith(".xlsx") && !n.replace(/\\/g, "/").toLowerCase().startsWith("references/")
  );
  if (xlsxNames.length === 0) {
    throw new Error("No Excel (.xlsx) workbook found in the ZIP (skipping references/).");
  }

  const preferred =
    xlsxNames.find((n) => n.toLowerCase().includes("mapping_")) ??
    xlsxNames.find((n) => !n.includes("/")) ??
    [...xlsxNames].sort((a, b) => a.length - b.length)[0];

  const entry = zip.file(preferred);
  if (!entry) throw new Error("Could not open the mapping workbook inside the ZIP.");

  const buf = await entry.async("arraybuffer");
  const wb = XLSX.read(buf, { type: "array" });
  const sheetName = pickSheetName(wb.SheetNames);
  if (!sheetName) throw new Error("Workbook contains no sheets.");

  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet "${sheetName}" is missing.`);

  const { headers, rows } = rowObjectsFromSheet(ws);
  if (rows.length === 0) throw new Error(`Sheet "${sheetName}" has no data rows.`);

  const excelFileName = preferred.replace(/\\/g, "/").split("/").pop() ?? preferred;
  return {
    excelFileName,
    sheetName,
    headers,
    rows,
    stats: computeStats(rows),
    zipEntryNames,
  };
}
