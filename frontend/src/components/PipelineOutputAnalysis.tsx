import { useMemo, useState, useEffect } from "react";
import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import type { PipelineAnalysis } from "../utils/pipelineZipAnalysis";
import { classifyExcelKey, parseConfidence, pickColumn } from "../utils/pipelineZipAnalysis";
import { approveEditSession, createEditSession, updateEditSession } from "../api/mappingApi";

const AL = {
  excelKey: ["Mapped Excel Key", "matched_excel_key"],
  partner: ["Partner Field", "partner_field", "field_name"],
  jsonKey: ["JSON Key", "json_key", "lms_column"],
  entity: ["Entity", "entity"],
  matchType: ["Match Type", "match_type"],
  confidence: ["Confidence", "confidence"],
};

type Segment = "all" | "loan" | "other" | "unmatched";

function formatPct01(x: number) {
  return `${Math.round(x * 1000) / 10}%`;
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

const PIE_COLORS = ["#4F46E5", "#059669", "#94A3B8"];

export function PipelineOutputAnalysis({
  analysis,
  error,
  clientName,
  processName,
  masterId,
  skipUnmatched,
}: {
  analysis: PipelineAnalysis | null;
  error: string | null;
  clientName: string;
  processName: string;
  masterId: number;
  skipUnmatched: boolean;
}) {
  const [segment, setSegment] = useState<Segment>("all");
  const [search, setSearch] = useState("");
  const [entityFilter, setEntityFilter] = useState("");
  const [matchTypeFilter, setMatchTypeFilter] = useState("");
  const [minConf, setMinConf] = useState(0);
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(100);
  const [fullScreen, setFullScreen] = useState(false);
  const [sessionId, setSessionId] = useState<string | null>(null);
  const [sessionStatus, setSessionStatus] = useState<string | null>(null);
  const [sessionMsg, setSessionMsg] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  type DraftRow = {
    id: string;
    partner_field: string;
    matched_excel_key: string | null;
    json_key: string | null;
    entity: string;
    confidence: number | null;
    match_type: string | null;
    needs_review: boolean;
    reasoning: string | null;
    winning_engine: string | null;
  };

  const draftRows: DraftRow[] = useMemo(() => {
    if (!analysis) return [];
    const out: DraftRow[] = [];
    for (let idx = 0; idx < analysis.rows.length; idx += 1) {
      const row = analysis.rows[idx]!;
      const partner = pickColumn(row, AL.partner);
      const excelKey = pickColumn(row, AL.excelKey);
      const jsonKey = pickColumn(row, AL.jsonKey);
      const entity = pickColumn(row, AL.entity) || "OTHER";
      const matchType = pickColumn(row, AL.matchType) || null;
      const confRaw = pickColumn(row, AL.confidence);
      const conf = confRaw ? parseConfidence(confRaw) : Number.NaN;
      const confidence = Number.isNaN(conf) ? null : conf;

      if (!partner) continue;
      out.push({
        id: String(idx),
        partner_field: partner,
        matched_excel_key: excelKey ? excelKey : null,
        json_key: jsonKey ? jsonKey : null,
        entity,
        confidence,
        match_type: matchType,
        needs_review: false,
        reasoning: null,
        winning_engine: null,
      });
    }
    return out;
  }, [analysis]);

  const [edits, setEdits] = useState<
    Record<string, Partial<Pick<DraftRow, "json_key" | "matched_excel_key">>>
  >({});

  const effectiveDraftRows = useMemo(() => {
    if (!draftRows.length) return [];
    return draftRows.map((r) => ({ ...r, ...(edits[r.id] ?? {}) }));
  }, [draftRows, edits]);

  async function saveDraft() {
    if (!analysis) return;
    if (!effectiveDraftRows.length) {
      setSessionMsg("No rows to save.");
      return;
    }
    setBusy(true);
    setSessionMsg(null);
    try {
      const mappings = effectiveDraftRows.map((r) => ({
        partner_field: r.partner_field,
        matched_excel_key: r.matched_excel_key,
        json_key: r.json_key,
        entity: r.entity || "OTHER",
        confidence: r.confidence ?? 0,
        match_type: r.match_type ?? "unmatched",
        reasoning: r.reasoning ?? "",
        needs_review: r.needs_review ?? false,
        winning_engine: r.winning_engine,
      }));

      if (!sessionId) {
        const created = await createEditSession({
          client_name: clientName,
          process_name: processName,
          master_id: masterId,
          created_by: "lp-mapping-ui",
          mappings,
        });
        setSessionId(created.session_id);
        setSessionStatus(created.status);
        setSessionMsg(`Draft saved (session: ${created.session_id}).`);
      } else {
        const updated = await updateEditSession(sessionId, {
          client_name: clientName,
          process_name: processName,
          master_id: masterId,
          updated_by: "lp-mapping-ui",
          note: "Full-pipeline edits save",
          mappings,
        });
        setSessionStatus(updated.status);
        setSessionMsg("Draft updated.");
      }
    } catch (e) {
      setSessionMsg(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }

  async function approveAndWriteDb() {
    if (!sessionId) {
      setSessionMsg("Save a draft first (no session yet).");
      return;
    }
    setBusy(true);
    setSessionMsg(null);
    try {
      const approved = await approveEditSession(sessionId, {
        approved_by: "lp-mapping-ui",
        master_id: masterId,
        skip_unmatched: skipUnmatched,
      });
      setSessionStatus(approved.status);
      const dbRes = (approved.approval_result as any)?.db_result;
      const inserted = dbRes?.inserted ?? dbRes?.inserted_updated ?? dbRes?.insertedUpdated;
      setSessionMsg(
        `Approved and written to DB${inserted != null ? ` (inserted/updated: ${inserted})` : ""}.`
      );
    } catch (e) {
      setSessionMsg(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }

  const entityOptions = useMemo(() => {
    if (!analysis) return [];
    const s = new Set<string>();
    for (const row of analysis.rows) {
      const e = pickColumn(row, AL.entity);
      if (e) s.add(e);
    }
    return [...s].sort((a, b) => a.localeCompare(b));
  }, [analysis]);

  const matchTypeOptions = useMemo(() => {
    if (!analysis) return [];
    return Object.keys(analysis.stats.matchTypeCounts).sort((a, b) => a.localeCompare(b));
  }, [analysis]);

  const filteredRows = useMemo(() => {
    if (!analysis) return [];
    const q = search.trim().toLowerCase();
    const ef = entityFilter.trim().toLowerCase();
    const mt = matchTypeFilter.trim();

    return analysis.rows.flatMap((row, idx) => {
      const ek = pickColumn(row, AL.excelKey);
      const kind = classifyExcelKey(ek);
      if (segment === "loan" && kind !== "loan") return [];
      if (segment === "other" && kind !== "other") return [];
      if (segment === "unmatched" && kind !== "unmatched") return [];

      if (ef && pickColumn(row, AL.entity).toLowerCase() !== ef) return [];
      if (mt && pickColumn(row, AL.matchType) !== mt) return [];

      const c = parseConfidence(pickColumn(row, AL.confidence));
      if (!Number.isNaN(c) && c < minConf) return [];
      if (Number.isNaN(c) && minConf > 0) return [];

      if (q) {
        const hay = [
          pickColumn(row, AL.partner),
          pickColumn(row, AL.jsonKey),
          pickColumn(row, AL.excelKey),
          pickColumn(row, AL.matchType),
        ]
          .join(" ")
          .toLowerCase();
        if (!hay.includes(q)) return [];
      }
      return [{ row, idx }];
    });
  }, [analysis, segment, search, entityFilter, matchTypeFilter, minConf]);

  // Reset paging when filters change (so users don't land on an empty page).
  useEffect(() => {
    setPage(1);
  }, [segment, search, entityFilter, matchTypeFilter, minConf, analysis]);

  const totalPages = useMemo(() => {
    const n = filteredRows.length;
    const ps = Math.max(1, pageSize || 1);
    return Math.max(1, Math.ceil(n / ps));
  }, [filteredRows.length, pageSize]);

  useEffect(() => {
    setPage((p) => Math.min(Math.max(1, p), totalPages));
  }, [totalPages]);

  const pagedRows = useMemo(() => {
    const ps = Math.max(1, pageSize || 1);
    const start = (Math.max(1, page) - 1) * ps;
    return filteredRows.slice(start, start + ps);
  }, [filteredRows, page, pageSize]);

  useEffect(() => {
    if (!fullScreen) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") setFullScreen(false);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [fullScreen]);

  // In full screen we use flex layout so the table can consume all remaining height.
  const tableMaxHeight = fullScreen ? undefined : 480;

  function renderFiltersAndTable(opts: { showFullScreenButton: boolean; isFullScreen: boolean }) {
    const a = analysis!;
    const inputStyle = {
      width: "100%",
      minWidth: opts.isFullScreen ? 280 : 180,
      padding: opts.isFullScreen ? "0.6rem 0.75rem" : "0.7rem 0.8rem",
      fontSize: opts.isFullScreen ? 15 : 15,
      fontWeight: opts.isFullScreen ? 700 : 600,
      lineHeight: 1.25,
      borderWidth: 2,
      borderColor: "#CBD5E1",
      background: "#FFFFFF",
    } as const;
    const editCellStyle = {
      padding: opts.isFullScreen ? "0.5rem 0.6rem" : "0.55rem 0.65rem",
      verticalAlign: "top",
    } as const;
    const tableStyle = {
      fontSize: opts.isFullScreen ? 14 : 13,
    } as const;
    const containerStyle = opts.isFullScreen
      ? ({ display: "flex", flexDirection: "column", height: "100%", minHeight: 0 } as const)
      : undefined;
    const tableWrapStyle = opts.isFullScreen
      ? ({ flex: "1 1 auto", minHeight: 0, maxHeight: "none" } as const)
      : ({ maxHeight: tableMaxHeight } as const);

    return (
      <div style={containerStyle}>
        <div
          style={{
            fontSize: 12,
            fontWeight: 600,
            marginBottom: opts.isFullScreen ? "0.35rem" : "0.75rem",
            color: "var(--ink)",
          }}
        >
          Filters & table
        </div>
        <div
          style={{
            display: "flex",
            flexWrap: "wrap",
            gap: opts.isFullScreen ? "0.5rem" : "0.75rem",
            alignItems: "flex-end",
            marginBottom: opts.isFullScreen ? "0.55rem" : "1rem",
          }}
        >
          <div className="field" style={{ marginBottom: 0, minWidth: 200 }}>
            <label className="field-label">Segment</label>
            <select
              className="field-input"
              value={segment}
              onChange={(e) => setSegment(e.target.value as Segment)}
            >
              <option value="all">All rows</option>
              <option value="loan">LOANPARAMETER* only</option>
              <option value="other">Other matched keys</option>
              <option value="unmatched">Unmatched only</option>
            </select>
          </div>
          <div className="field" style={{ marginBottom: 0, flex: "1 1 200px" }}>
            <label className="field-label">Search (partner / JSON / Excel key / match type)</label>
            <input
              className="field-input"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Type to filter…"
            />
          </div>
          <div className="field" style={{ marginBottom: 0, minWidth: 160 }}>
            <label className="field-label">Entity</label>
            <select
              className="field-input"
              value={entityFilter}
              onChange={(e) => setEntityFilter(e.target.value)}
            >
              <option value="">Any</option>
              {entityOptions.map((e) => (
                <option key={e} value={e}>
                  {e}
                </option>
              ))}
            </select>
          </div>
          <div className="field" style={{ marginBottom: 0, minWidth: 180 }}>
            <label className="field-label">Match type</label>
            <select
              className="field-input"
              value={matchTypeFilter}
              onChange={(e) => setMatchTypeFilter(e.target.value)}
            >
              <option value="">Any</option>
              {matchTypeOptions.map((e) => (
                <option key={e} value={e}>
                  {e}
                </option>
              ))}
            </select>
          </div>
          <div className="field" style={{ marginBottom: 0, minWidth: 200 }}>
            <label className="field-label">Min confidence: {formatPct01(minConf)}</label>
            <input
              type="range"
              min={0}
              max={100}
              value={Math.round(minConf * 100)}
              onChange={(e) => setMinConf(Number(e.target.value) / 100)}
              style={{ width: "100%" }}
            />
          </div>
          <button
            type="button"
            className="btn btn-secondary"
            onClick={() =>
              downloadCsv(
                `filtered_${a.excelFileName.replace(/\.xlsx$/i, "")}_rows.csv`,
                a.headers,
                filteredRows.map((x) => x.row)
              )
            }
            disabled={filteredRows.length === 0}
          >
            ⬇ Export filtered CSV
          </button>
          {opts.showFullScreenButton && (
            <button type="button" className="btn btn-ghost" onClick={() => setFullScreen(true)}>
              ⛶ Full screen
            </button>
          )}
        </div>

        <div style={{ fontSize: 12, color: "var(--ink-3)", marginBottom: opts.isFullScreen ? "0.35rem" : "0.5rem" }}>
          Showing {pagedRows.length} row(s) on this page · {filteredRows.length} filtered row(s)
          {filteredRows.length !== stats.total ? ` (of ${stats.total})` : ""}.
        </div>

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
            <button
              type="button"
              className="btn btn-secondary"
              onClick={() => setPage(1)}
              disabled={page <= 1}
            >
              ⏮ First
            </button>
            <button
              type="button"
              className="btn btn-secondary"
              onClick={() => setPage((p) => Math.max(1, p - 1))}
              disabled={page <= 1}
            >
              ← Prev
            </button>
            <span style={{ fontSize: 12, color: "var(--ink-2)" }}>
              Page <b>{page}</b> / {totalPages}
            </span>
            <button
              type="button"
              className="btn btn-secondary"
              onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
              disabled={page >= totalPages}
            >
              Next →
            </button>
            <button
              type="button"
              className="btn btn-secondary"
              onClick={() => setPage(totalPages)}
              disabled={page >= totalPages}
            >
              Last ⏭
            </button>
          </div>
          <div className="field" style={{ marginBottom: 0, minWidth: 170 }}>
            <label className="field-label">Rows per page</label>
            <select
              className="field-input"
              value={String(pageSize)}
              onChange={(e) => setPageSize(Number(e.target.value))}
            >
              <option value="25">25</option>
              <option value="50">50</option>
              <option value="100">100</option>
              <option value="200">200</option>
              <option value="500">500</option>
            </select>
          </div>
        </div>

        <div className="table-wrap" style={tableWrapStyle}>
          <table className="data-table" style={tableStyle}>
            <thead>
              <tr>
                <th style={{ width: 110 }}>Key kind</th>
                <th style={{ width: opts.isFullScreen ? 320 : 220 }}>Edit Excel key</th>
                <th style={{ width: opts.isFullScreen ? 420 : 260 }}>Edit JSON key</th>
                {a.headers.map((h) => (
                  <th key={h}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {pagedRows.map(({ row, idx }) => {
                const ek = pickColumn(row, AL.excelKey);
                const k = classifyExcelKey(ek);
                const kindLabel = k === "loan" ? "LOANPARAM" : k === "other" ? "Other" : "—";
                const chipClass =
                  k === "loan" ? "chip chip-blue" : k === "other" ? "chip chip-green" : "chip chip-amber";

                const id = String(idx);
                const currentExcel = (edits[id]?.matched_excel_key ?? (ek ?? "")) as string;
                const currentJson = (edits[id]?.json_key ?? pickColumn(row, AL.jsonKey) ?? "") as string;
                return (
                  <tr key={id}>
                    <td>
                      <span className={chipClass}>{kindLabel}</span>
                    </td>
                    <td style={editCellStyle}>
                      <input
                        className="field-input"
                        style={inputStyle}
                        value={currentExcel}
                        onChange={(e) =>
                          setEdits((p) => ({
                            ...p,
                            [id]: { ...(p[id] ?? {}), matched_excel_key: e.target.value },
                          }))
                        }
                      />
                    </td>
                    <td style={editCellStyle}>
                      <input
                        className="field-input"
                        style={inputStyle}
                        value={currentJson}
                        onChange={(e) =>
                          setEdits((p) => ({
                            ...p,
                            [id]: { ...(p[id] ?? {}), json_key: e.target.value },
                          }))
                        }
                      />
                    </td>
                    {a.headers.map((h) => (
                      <td key={h}>{row[h] ?? ""}</td>
                    ))}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    );
  }

  const pieData = useMemo(() => {
    if (!analysis) return [];
    const { loanSlots, otherKeys, unmatched } = analysis.stats;
    return [
      { name: "LOANPARAMETER slots", value: loanSlots },
      { name: "Other Excel keys", value: otherKeys },
      { name: "Unmatched", value: unmatched },
    ].filter((d) => d.value > 0);
  }, [analysis]);

  const matchTypeChartData = useMemo(() => {
    if (!analysis) return [];
    const entries = Object.entries(analysis.stats.matchTypeCounts).sort((a, b) => b[1] - a[1]);
    const top = entries.slice(0, 14);
    const restSum = entries.slice(14).reduce((s, [, c]) => s + c, 0);
    if (restSum > 0) top.push(["Other (aggregated)", restSum]);
    return top.map(([name, count]) => ({ name: name.length > 28 ? `${name.slice(0, 26)}…` : name, count, fullName: name }));
  }, [analysis]);

  if (error) {
    return (
      <div className="panel" style={{ marginTop: "1rem" }}>
        <div className="panel-header">
          <div>
            <div className="panel-title">Output analysis</div>
            <div className="panel-desc">Could not parse the ZIP in the browser</div>
          </div>
        </div>
        <div className="panel-body">
          <div className="alert alert-error" style={{ marginBottom: 0 }}>
            <span>⚠</span>
            <span style={{ flex: 1 }}>{error}</span>
          </div>
        </div>
      </div>
    );
  }

  if (!analysis) return null;

  const { stats } = analysis;
  const coverage = stats.total ? stats.matched / stats.total : 0;
  const loanShare = stats.matched ? stats.loanSlots / stats.matched : 0;
  const otherShare = stats.matched ? stats.otherKeys / stats.matched : 0;

  return (
    <>
      <div className="panel" style={{ marginTop: "1rem" }}>
        <div className="panel-header">
          <div>
            <div className="panel-title">
              Output analysis
              <span className="file-name-tag">📊 {analysis.excelFileName}</span>
            </div>
            <div className="panel-desc">
              Sheet “{analysis.sheetName}” · {stats.total} rows · ZIP has {analysis.zipEntryNames.length} entries
            </div>
          </div>
          <span className="badge badge-indigo">
            <span className="badge-dot"></span> Live from ZIP
          </span>
        </div>

        <div className="panel-body">
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: "0.75rem" }}>
          <button type="button" className="btn btn-secondary" onClick={() => void saveDraft()} disabled={busy || effectiveDraftRows.length === 0}>
            💾 Save draft
          </button>
          <button type="button" className="btn btn-primary" onClick={() => void approveAndWriteDb()} disabled={busy || !sessionId}>
            ✓ Approve & write DB
          </button>
          {sessionId && (
            <span className="chip chip-blue" style={{ fontFamily: "inherit" }}>
              session_id: {sessionId}{sessionStatus ? ` · ${sessionStatus}` : ""}
            </span>
          )}
        </div>
        {sessionMsg && (
          <div className={`alert ${sessionMsg.toLowerCase().includes("approved") || sessionMsg.toLowerCase().includes("draft") ? "alert-success" : "alert-error"}`}>
            <span>{sessionMsg.toLowerCase().includes("approved") || sessionMsg.toLowerCase().includes("draft") ? "✓" : "⚠"}</span>
            <span style={{ flex: 1 }}>{sessionMsg}</span>
            <button className="alert-dismiss" onClick={() => setSessionMsg(null)} type="button">
              ×
            </button>
          </div>
        )}

        {analysis.zipEntryNames.length > 0 && (
          <div style={{ marginBottom: "1rem", display: "flex", flexWrap: "wrap", gap: 6 }}>
            {analysis.zipEntryNames.slice(0, 24).map((n) => (
              <span key={n} className="chip chip-blue" style={{ fontFamily: "inherit" }}>
                {n.replace(/\\/g, "/")}
              </span>
            ))}
            {analysis.zipEntryNames.length > 24 && (
              <span className="chip chip-amber" style={{ fontFamily: "inherit" }}>
                +{analysis.zipEntryNames.length - 24} more
              </span>
            )}
          </div>
        )}

        <div className="metric-row" style={{ gridTemplateColumns: "repeat(4, 1fr)", marginBottom: "1rem" }}>
          <div className="metric-card">
            <div className="metric-accent acc-indigo"></div>
            <div className="metric-label">Match coverage</div>
            <div className="metric-value mv-indigo">{formatPct01(coverage)}</div>
            <div className="metric-sub">
              {stats.matched} matched / {stats.total} total
            </div>
          </div>
          <div className="metric-card">
            <div className="metric-accent acc-violet"></div>
            <div className="metric-label">Avg confidence</div>
            <div className="metric-value" style={{ color: "var(--violet)" }}>
              {stats.confidenceRowCount ? formatPct01(stats.avgConfidence) : "—"}
            </div>
            <div className="metric-sub">
              {stats.confidenceRowCount ? `Across ${stats.confidenceRowCount} rows with scores` : "No scores parsed"}
            </div>
          </div>
          <div className="metric-card">
            <div className="metric-accent acc-emerald"></div>
            <div className="metric-label">High confidence</div>
            <div className="metric-value mv-emerald">{stats.highConfidence}</div>
            <div className="metric-sub">≥ 85% where confidence present</div>
          </div>
          <div className="metric-card">
            <div className="metric-accent acc-amber"></div>
            <div className="metric-label">Needs review</div>
            <div className="metric-value mv-amber">{stats.reviewCount}</div>
            <div className="metric-sub">Flagged in workbook</div>
          </div>
        </div>

        <div
          className="metric-row"
          style={{
            gridTemplateColumns: "repeat(3, 1fr)",
            marginBottom: "1.25rem",
          }}
        >
          <div className="metric-card">
            <div className="metric-accent acc-indigo"></div>
            <div className="metric-label">LOANPARAMETER slots</div>
            <div className="metric-value mv-indigo">{stats.loanSlots}</div>
            <div className="metric-sub">
              {stats.matched ? `${formatPct01(loanShare)} of matched` : "—"}
            </div>
          </div>
          <div className="metric-card">
            <div className="metric-accent acc-emerald"></div>
            <div className="metric-label">Other Excel keys</div>
            <div className="metric-value mv-emerald">{stats.otherKeys}</div>
            <div className="metric-sub">
              {stats.matched ? `${formatPct01(otherShare)} of matched` : "—"}
            </div>
          </div>
          <div className="metric-card">
            <div className="metric-accent acc-amber"></div>
            <div className="metric-label">Unmapped rows</div>
            <div className="metric-value mv-amber">{stats.unmatched}</div>
            <div className="metric-sub">Blank mapped Excel key</div>
          </div>
        </div>

        <div
          style={{
            display: "grid",
            gridTemplateColumns: "1fr 1fr",
            gap: "1.25rem",
            marginBottom: "1.25rem",
          }}
        >
          <div
            style={{
              background: "#F8FAFC",
              border: "1px solid var(--border)",
              borderRadius: "var(--radius-sm)",
              padding: "1rem",
            }}
          >
            <div style={{ fontSize: 12, fontWeight: 600, color: "var(--ink-2)", marginBottom: 8 }}>
              Mapped key mix
            </div>
            <div style={{ width: "100%", height: 260 }}>
              {pieData.length === 0 ? (
                <div
                  style={{
                    height: "100%",
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    color: "var(--ink-3)",
                    fontSize: 13,
                  }}
                >
                  No mapped rows to chart
                </div>
              ) : (
                <ResponsiveContainer>
                  <PieChart>
                    <Pie
                      data={pieData}
                      dataKey="value"
                      nameKey="name"
                      cx="50%"
                      cy="50%"
                      innerRadius={52}
                      outerRadius={88}
                      paddingAngle={2}
                    >
                      {pieData.map((_, i) => (
                        <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />
                      ))}
                    </Pie>
                    <Tooltip formatter={(value) => [`${value ?? ""}`, "Rows"]} />
                  </PieChart>
                </ResponsiveContainer>
              )}
            </div>
            {pieData.length > 0 && (
              <div style={{ display: "grid", gap: 6, marginTop: 10 }}>
                {(() => {
                  const total = pieData.reduce((s, d) => s + (Number(d.value) || 0), 0);
                  return pieData.map((d, i) => {
                    const v = Number(d.value) || 0;
                    const pct = total > 0 ? Math.round((v / total) * 1000) / 10 : 0;
                    const color = PIE_COLORS[i % PIE_COLORS.length];
                    return (
                      <div
                        key={d.name}
                        style={{
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "space-between",
                          gap: 10,
                          fontSize: 13,
                          color: "var(--ink-2)",
                        }}
                      >
                        <div style={{ display: "flex", alignItems: "center", gap: 8, minWidth: 0 }}>
                          <span
                            aria-hidden="true"
                            style={{
                              width: 10,
                              height: 10,
                              borderRadius: 999,
                              background: color,
                              border: "1px solid rgba(15,23,42,0.12)",
                              flex: "0 0 auto",
                            }}
                          />
                          <span style={{ fontWeight: 700, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                            {d.name}
                          </span>
                        </div>
                        <span style={{ fontFamily: "ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, Liberation Mono, monospace" }}>
                          {v} ({pct}%)
                        </span>
                      </div>
                    );
                  });
                })()}
              </div>
            )}
          </div>
          <div
            style={{
              background: "#F8FAFC",
              border: "1px solid var(--border)",
              borderRadius: "var(--radius-sm)",
              padding: "1rem",
            }}
          >
            <div style={{ fontSize: 12, fontWeight: 600, color: "var(--ink-2)", marginBottom: 8 }}>
              Confidence bands
            </div>
            <div style={{ width: "100%", height: 260 }}>
              <ResponsiveContainer>
                <BarChart data={stats.confidenceBuckets} layout="vertical" margin={{ left: 8, right: 16 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
                  <XAxis type="number" allowDecimals={false} />
                  <YAxis type="category" dataKey="label" width={72} tick={{ fontSize: 11 }} />
                  <Tooltip formatter={(value) => [`${value ?? ""}`, "Rows"]} />
                  <Bar dataKey="count" radius={[0, 4, 4, 0]}>
                    {stats.confidenceBuckets.map((e, i) => (
                      <Cell key={i} fill={e.fill} />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>

        <div
          style={{
            background: "#F8FAFC",
            border: "1px solid var(--border)",
            borderRadius: "var(--radius-sm)",
            padding: "1rem",
            marginBottom: "1.25rem",
          }}
        >
          <div style={{ fontSize: 12, fontWeight: 600, color: "var(--ink-2)", marginBottom: 8 }}>
            Match type distribution
          </div>
          <div style={{ width: "100%", height: 280 }}>
            <ResponsiveContainer>
              <BarChart data={matchTypeChartData} margin={{ bottom: 64, left: 8, right: 8 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
                <XAxis dataKey="name" angle={-28} textAnchor="end" interval={0} height={70} tick={{ fontSize: 10 }} />
                <YAxis allowDecimals={false} tick={{ fontSize: 11 }} />
                <Tooltip
                  formatter={(value, _name, item) => {
                    const full = (item?.payload as { fullName?: string } | undefined)?.fullName;
                    return [`${value ?? ""}`, full ?? "Rows"];
                  }}
                />
                <Bar dataKey="count" fill="#4F46E5" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {renderFiltersAndTable({ showFullScreenButton: true, isFullScreen: false })}

        <p style={{ fontSize: 11, color: "var(--ink-3)", marginTop: "0.75rem", lineHeight: 1.5 }}>
          “Accuracy” here means coverage (share of rows with a mapped Excel key) and average model confidence where the
          workbook provides it — not ground-truth correctness against a labeled set.
        </p>
        </div>
      </div>
      {fullScreen && (
        <div
          role="dialog"
          aria-modal="true"
          style={{
            position: "fixed",
            inset: 0,
            background: "rgba(15, 23, 42, 0.55)",
            zIndex: 9999,
            padding: 6,
          }}
          onMouseDown={() => setFullScreen(false)}
        >
          <div
            style={{
              height: "100%",
              maxWidth: "calc(100vw - 10px)",
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
                padding: "8px 10px",
                borderBottom: "1px solid var(--border)",
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                gap: 12,
              }}
            >
              <div style={{ fontSize: 14, fontWeight: 600, color: "var(--ink)" }}>
                Full screen editor · {analysis.excelFileName}
              </div>
              <button type="button" className="btn btn-secondary" onClick={() => setFullScreen(false)}>
                Close (Esc)
              </button>
            </div>
            <div style={{ padding: 8, overflow: "hidden", flex: 1, minHeight: 0 }}>
              {renderFiltersAndTable({ showFullScreenButton: false, isFullScreen: true })}
            </div>
          </div>
        </div>
      )}
    </>
  );
}
