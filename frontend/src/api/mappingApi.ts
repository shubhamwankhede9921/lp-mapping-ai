/**
 * Client for LP Field Mapping FastAPI (`/api/llm_mapping`).
 * Base URL: use empty string in dev with Vite proxy, or set VITE_API_BASE_URL.
 */

const PREFIX = "/api/llm_mapping";

const SESSION_API_BASE = "lp_mapping_api_base";

/** Empty string = same origin (use Vite dev proxy) or VITE_API_BASE_URL at build time. */
function apiRoot(): string {
  try {
    const s = sessionStorage.getItem(SESSION_API_BASE);
    if (s != null && s.trim() !== "") {
      return s.trim().replace(/\/$/, "");
    }
  } catch {
    /* private mode */
  }
  const base = import.meta.env.VITE_API_BASE_URL ?? "";
  return String(base).replace(/\/$/, "");
}

export function mappingUrl(path: string): string {
  const p = path.startsWith("/") ? path : `/${path}`;
  return `${apiRoot()}${PREFIX}${p}`;
}

export interface FieldMapping {
  partner_field: string;
  column_category?: string | null;
  entity: string;
  matched_excel_key?: string | null;
  json_key?: string | null;
  confidence: number;
  match_type: string;
  reasoning: string;
  needs_review: boolean;
  fuzzy_score?: number | null;
  embedding_score?: number | null;
  llm_confidence?: number | null;
  winning_engine?: string | null;
}

export interface Stats {
  total_fields: number;
  matched: number;
  unmatched: number;
  match_rate_pct: number;
  needs_review: number;
  avg_confidence: number;
  by_match_type: Record<string, number>;
  by_entity: Record<string, number>;
  by_confidence_band: Record<string, number>;
}

export interface DeterministicResponse {
  client_name: string;
  process_name: string;
  mappings: FieldMapping[];
  unmatched_fields: Record<string, unknown>[];
  llm_prompts_count: number;
  stats: Stats;
}

export interface HybridLLMResponse {
  client_name: string;
  process_name: string;
  mappings: FieldMapping[];
  stats: Stats;
  engine_breakdown: Record<string, number>;
}

export interface RefStatusResponse {
  ready: boolean;
  files: Record<
    string,
    { exists: boolean; size_kb: number | null }
  >;
}

async function parseJsonError(res: Response): Promise<string> {
  try {
    const j = (await res.json()) as { detail?: unknown };
    if (typeof j.detail === "string") return j.detail;
    return JSON.stringify(j.detail ?? j);
  } catch {
    return await res.text();
  }
}

export async function getHealth(signal?: AbortSignal): Promise<{ status: string }> {
  const res = await fetch(mappingUrl("/health"), { signal });
  if (!res.ok) throw new Error(await parseJsonError(res));
  return res.json();
}

export async function getReferencesStatus(signal?: AbortSignal): Promise<RefStatusResponse> {
  const res = await fetch(mappingUrl("/references/status"), { signal });
  if (!res.ok) throw new Error(await parseJsonError(res));
  return res.json();
}

export async function postBuildReferences(
  params: { putm_table_override?: string; mapping_table_override?: string },
  signal?: AbortSignal
): Promise<unknown> {
  const q = new URLSearchParams();
  if (params.putm_table_override) q.set("putm_table_override", params.putm_table_override);
  if (params.mapping_table_override) q.set("mapping_table_override", params.mapping_table_override);
  const url = `${mappingUrl("/references/build")}${q.toString() ? `?${q}` : ""}`;
  const res = await fetch(url, { method: "POST", signal });
  if (!res.ok) throw new Error(await parseJsonError(res));
  return res.json();
}

function appendBool(fd: FormData, key: string, v: boolean): void {
  fd.append(key, v ? "true" : "false");
}

export async function postDeterministic(
  file: File,
  body: {
    client_name: string;
    process_name: string;
    sheet_filter?: string;
    use_loanparameter_refinement: boolean;
    use_llm_entity_classifier: boolean;
  },
  signal?: AbortSignal
): Promise<DeterministicResponse> {
  const fd = new FormData();
  fd.append("file", file);
  fd.append("client_name", body.client_name);
  fd.append("process_name", body.process_name);
  fd.append("sheet_filter", body.sheet_filter ?? "");
  appendBool(fd, "use_loanparameter_refinement", body.use_loanparameter_refinement);
  appendBool(fd, "use_llm_entity_classifier", body.use_llm_entity_classifier);
  const res = await fetch(mappingUrl("/mapping/deterministic"), {
    method: "POST",
    body: fd,
    signal,
  });
  if (!res.ok) throw new Error(await parseJsonError(res));
  return res.json();
}

export async function postHybridLlm(
  file: File,
  body: {
    client_name: string;
    process_name: string;
    use_fuzzy: boolean;
    use_embeddings: boolean;
    use_llm: boolean;
    use_loanparameter_refinement: boolean;
    use_llm_entity_classifier: boolean;
    master_id: number;
    save_to_db: boolean;
    skip_unmatched: boolean;
  },
  signal?: AbortSignal
): Promise<HybridLLMResponse> {
  const fd = new FormData();
  fd.append("file", file);
  fd.append("client_name", body.client_name);
  fd.append("process_name", body.process_name);
  appendBool(fd, "use_fuzzy", body.use_fuzzy);
  appendBool(fd, "use_embeddings", body.use_embeddings);
  appendBool(fd, "use_llm", body.use_llm);
  appendBool(fd, "use_loanparameter_refinement", body.use_loanparameter_refinement);
  appendBool(fd, "use_llm_entity_classifier", body.use_llm_entity_classifier);
  fd.append("master_id", String(body.master_id));
  appendBool(fd, "save_to_db", body.save_to_db);
  appendBool(fd, "skip_unmatched", body.skip_unmatched);
  const res = await fetch(mappingUrl("/mapping/hybrid-llm"), {
    method: "POST",
    body: fd,
    signal,
  });
  if (!res.ok) throw new Error(await parseJsonError(res));
  return res.json();
}

export interface FullPipelineResult {
  blob: Blob;
  filename: string;
  headers: {
    dbInserted: string;
    dbSkipped: string;
    dbErrors: string;
    masterId: string;
    refPutmRows?: string;
    refMappingRows?: string;
  };
}

export async function postFullPipeline(
  file: File,
  body: {
    client_name: string;
    process_name: string;
    use_fuzzy: boolean;
    use_embeddings: boolean;
    use_llm: boolean;
    use_loanparameter_refinement: boolean;
    use_llm_entity_classifier: boolean;
    sheet_filter?: string;
    master_id: number;
    save_to_db: boolean;
    skip_unmatched: boolean;
    include_build_references: boolean;
  },
  signal?: AbortSignal
): Promise<FullPipelineResult> {
  const fd = new FormData();
  fd.append("file", file);
  fd.append("client_name", body.client_name);
  fd.append("process_name", body.process_name);
  appendBool(fd, "use_fuzzy", body.use_fuzzy);
  appendBool(fd, "use_embeddings", body.use_embeddings);
  appendBool(fd, "use_llm", body.use_llm);
  appendBool(fd, "use_loanparameter_refinement", body.use_loanparameter_refinement);
  appendBool(fd, "use_llm_entity_classifier", body.use_llm_entity_classifier);
  fd.append("sheet_filter", body.sheet_filter ?? "");
  fd.append("master_id", String(body.master_id));
  appendBool(fd, "save_to_db", body.save_to_db);
  appendBool(fd, "skip_unmatched", body.skip_unmatched);
  appendBool(fd, "include_build_references", body.include_build_references);
  const res = await fetch(mappingUrl("/mapping/full-pipeline"), {
    method: "POST",
    body: fd,
    signal,
  });
  if (!res.ok) throw new Error(await parseJsonError(res));
  const cd = res.headers.get("Content-Disposition") ?? "";
  let filename = "outputs.zip";
  const m = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/.exec(cd);
  if (m?.[1]) filename = m[1].replace(/['"]/g, "").trim();
  const blob = await res.blob();
  return {
    blob,
    filename,
    headers: {
      dbInserted: res.headers.get("X-DB-Inserted") ?? "—",
      dbSkipped: res.headers.get("X-DB-Skipped") ?? "—",
      dbErrors: res.headers.get("X-DB-Errors") ?? "—",
      masterId: res.headers.get("X-Master-Id") ?? "—",
      refPutmRows: res.headers.get("X-References-PutM-Rows") ?? undefined,
      refMappingRows: res.headers.get("X-References-Mapping-Rows") ?? undefined,
    },
  };
}

export interface NestedMappingResponse {
  client_name?: string | null;
  los_json: Record<string, unknown>;
  total_input: number;
  mapped_count: number;
  skipped_count: number;
  processing_time_ms: number;
}

export interface SchemaResponse {
  client_name?: string | null;
  los_schema: Record<string, unknown>;
  total_input: number;
  mapped_count: number;
  skipped_count: number;
  processing_time_ms: number;
}

export interface LosJsonRequest {
  client_name?: string | null;
  process_name?: string | null;
  mappings: unknown[];
}

export async function postNestedMapping(
  body: LosJsonRequest,
  signal?: AbortSignal
): Promise<NestedMappingResponse> {
  const res = await fetch(mappingUrl("/generate-nested-mapping"), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
    signal,
  });
  if (!res.ok) throw new Error(await parseJsonError(res));
  return res.json();
}

export async function postGenerateSchema(
  body: LosJsonRequest,
  signal?: AbortSignal
): Promise<SchemaResponse> {
  const res = await fetch(mappingUrl("/generate-schema"), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
    signal,
  });
  if (!res.ok) throw new Error(await parseJsonError(res));
  return res.json();
}

export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

// ── Edit sessions (draft → approve → DB write) ────────────────────────────────

export interface EditSessionSummary {
  session_id: string;
  status: string;
  client_name?: string | null;
  process_name?: string | null;
  master_id?: number | null;
  created_at?: string | null;
  updated_at?: string | null;
  approved_at?: string | null;
}

export interface EditSessionResponse {
  session_id: string;
  status: "draft" | "approved" | string;
  client_name: string;
  process_name: string;
  master_id?: number | null;
  created_at: string;
  updated_at: string;
  created_by?: string | null;
  approved_at?: string | null;
  approved_by?: string | null;
  mappings: Record<string, unknown>[];
  approval_result?: Record<string, unknown> | null;
}

export async function listEditSessions(signal?: AbortSignal): Promise<EditSessionSummary[]> {
  const res = await fetch(mappingUrl("/edit-sessions"), { signal });
  if (!res.ok) throw new Error(await parseJsonError(res));
  return res.json();
}

export async function createEditSession(
  body: {
    client_name: string;
    process_name: string;
    master_id?: number;
    created_by?: string;
    mappings: Record<string, unknown>[];
  },
  signal?: AbortSignal
): Promise<EditSessionResponse> {
  const res = await fetch(mappingUrl("/edit-sessions"), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
    signal,
  });
  if (!res.ok) throw new Error(await parseJsonError(res));
  return res.json();
}

export async function getEditSession(session_id: string, signal?: AbortSignal): Promise<EditSessionResponse> {
  const res = await fetch(mappingUrl(`/edit-sessions/${encodeURIComponent(session_id)}`), { signal });
  if (!res.ok) throw new Error(await parseJsonError(res));
  return res.json();
}

export async function updateEditSession(
  session_id: string,
  body: {
    client_name?: string;
    process_name?: string;
    master_id?: number;
    updated_by?: string;
    note?: string;
    mappings?: Record<string, unknown>[];
  },
  signal?: AbortSignal
): Promise<EditSessionResponse> {
  const res = await fetch(mappingUrl(`/edit-sessions/${encodeURIComponent(session_id)}`), {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
    signal,
  });
  if (!res.ok) throw new Error(await parseJsonError(res));
  return res.json();
}

export async function approveEditSession(
  session_id: string,
  body: { approved_by?: string; master_id?: number; skip_unmatched?: boolean },
  signal?: AbortSignal
): Promise<EditSessionResponse> {
  const res = await fetch(mappingUrl(`/edit-sessions/${encodeURIComponent(session_id)}/approve`), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
    signal,
  });
  if (!res.ok) throw new Error(await parseJsonError(res));
  return res.json();
}
