import type { FieldMapping, Stats } from "../api/mappingApi";

const COLS: (keyof FieldMapping | string)[] = [
  "partner_field",
  "matched_excel_key",
  "json_key",
  "confidence",
  "match_type",
  "entity",
  "needs_review",
  "winning_engine",
  "reasoning",
];

export function StatsBar({ stats }: { stats: Stats | null }) {
  if (!stats) return null;
  return (
    <div className="stats-bar">
      <div className="stat">
        <span className="stat-label">Total</span>
        <span className="stat-val">{stats.total_fields}</span>
      </div>
      <div className="stat">
        <span className="stat-label">Matched</span>
        <span className="stat-val ok">{stats.matched}</span>
      </div>
      <div className="stat">
        <span className="stat-label">Unmatched</span>
        <span className="stat-val warn">{stats.unmatched}</span>
      </div>
      <div className="stat">
        <span className="stat-label">Match rate</span>
        <span className="stat-val">{stats.match_rate_pct}%</span>
      </div>
      <div className="stat">
        <span className="stat-label">Avg conf.</span>
        <span className="stat-val">{stats.avg_confidence.toFixed(2)}</span>
      </div>
      <div className="stat">
        <span className="stat-label">Review</span>
        <span className="stat-val">{stats.needs_review}</span>
      </div>
    </div>
  );
}

function cellVal(m: FieldMapping, key: string): string {
  const v = (m as unknown as Record<string, unknown>)[key];
  if (v === null || v === undefined) return "";
  if (typeof v === "boolean") return v ? "yes" : "no";
  return String(v);
}

function confClass(c: number): string {
  if (c >= 0.9) return "conf-high";
  if (c >= 0.8) return "conf-mid";
  if (c >= 0.7) return "conf-low";
  return "conf-bad";
}

export function MappingsTable({ rows }: { rows: FieldMapping[] }) {
  if (!rows.length) {
    return <p className="muted">No mappings returned.</p>;
  }

  const keys = COLS.filter((k) =>
    rows.some((r) => (r as unknown as Record<string, unknown>)[k as string] != null)
  );

  return (
    <div className="table-wrap">
      <table className="data-table">
        <thead>
          <tr>
            {keys.map((k) => (
              <th key={String(k)}>{String(k).replace(/_/g, " ")}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((m, i) => (
            <tr key={i}>
              {keys.map((k) => {
                const s = cellVal(m, k as string);
                const isConf = k === "confidence";
                return (
                  <td
                    key={String(k)}
                    className={isConf ? confClass(Number(s) || 0) : undefined}
                  >
                    {s.length > 120 ? `${s.slice(0, 117)}…` : s}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export function EngineBreakdown({ b }: { b: Record<string, number> | null }) {
  if (!b || !Object.keys(b).length) return null;
  return (
    <div className="engine-row">
      {Object.entries(b).map(([k, v]) => (
        <span key={k} className="chip">
          {k}: {v}
        </span>
      ))}
    </div>
  );
}
