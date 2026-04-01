"""
LP Field Mapping — Streamlit Test Dashboard
Run with:  streamlit run lp_mapping_dashboard.py
"""

import io
import json
import time
import zipfile

import pandas as pd
import requests
import streamlit as st

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="LP Field Mapping",
    page_icon="🔗",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Syne:wght@400;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Syne', sans-serif;
}
code, pre, .stCode {
    font-family: 'DM Mono', monospace !important;
}

/* Background */
.stApp {
    background: #0d0f14;
    color: #e8eaf0;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: #111318 !important;
    border-right: 1px solid #1e2130;
}
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] .stMarkdown p {
    color: #8b90a8 !important;
    font-size: 0.78rem;
    letter-spacing: 0.05em;
    text-transform: uppercase;
}
[data-testid="stSidebar"] input,
[data-testid="stSidebar"] .stTextInput input {
    background: #1a1d28 !important;
    border: 1px solid #2a2e42 !important;
    color: #e8eaf0 !important;
    border-radius: 6px !important;
}

/* Metric cards */
[data-testid="metric-container"] {
    background: #13161f;
    border: 1px solid #1e2130;
    border-radius: 10px;
    padding: 1rem 1.2rem;
}
[data-testid="metric-container"] label {
    color: #5b6080 !important;
    font-size: 0.72rem !important;
    letter-spacing: 0.08em;
    text-transform: uppercase;
}
[data-testid="metric-container"] [data-testid="stMetricValue"] {
    color: #c8f0d8 !important;
    font-size: 2rem !important;
    font-weight: 700 !important;
}

/* Buttons */
.stButton > button {
    background: linear-gradient(135deg, #3a7bd5, #6c63ff);
    color: #fff;
    border: none;
    border-radius: 8px;
    font-family: 'Syne', sans-serif;
    font-weight: 700;
    letter-spacing: 0.04em;
    padding: 0.55rem 1.6rem;
    transition: opacity 0.2s, transform 0.1s;
}
.stButton > button:hover {
    opacity: 0.88;
    transform: translateY(-1px);
}

/* Tabs */
.stTabs [data-baseweb="tab"] {
    font-family: 'Syne', sans-serif;
    font-weight: 600;
    color: #5b6080;
    font-size: 0.82rem;
    letter-spacing: 0.06em;
    text-transform: uppercase;
}
.stTabs [aria-selected="true"] {
    color: #6c63ff !important;
    border-bottom: 2px solid #6c63ff !important;
}

/* File uploader */
[data-testid="stFileUploadDropzone"] {
    background: #13161f !important;
    border: 1.5px dashed #2a2e42 !important;
    border-radius: 10px !important;
}

/* Expander */
.streamlit-expanderHeader {
    background: #13161f !important;
    border: 1px solid #1e2130 !important;
    border-radius: 8px !important;
    color: #8b90a8 !important;
    font-size: 0.8rem;
    letter-spacing: 0.04em;
}

/* Dataframe */
[data-testid="stDataFrame"] {
    border: 1px solid #1e2130;
    border-radius: 8px;
    overflow: hidden;
}

/* Select boxes & toggles */
[data-testid="stSelectbox"] div[data-baseweb="select"] {
    background: #1a1d28 !important;
    border-color: #2a2e42 !important;
}

/* Badge-style chip */
.chip {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 20px;
    font-size: 0.7rem;
    font-weight: 600;
    letter-spacing: 0.05em;
    font-family: 'DM Mono', monospace;
}
.chip-green  { background: #0f2e1e; color: #4ade80; border: 1px solid #166534; }
.chip-yellow { background: #2e2410; color: #fbbf24; border: 1px solid #78350f; }
.chip-red    { background: #2e1010; color: #f87171; border: 1px solid #7f1d1d; }
.chip-blue   { background: #0f1e2e; color: #60a5fa; border: 1px solid #1e3a5f; }

/* Header rule */
hr { border-color: #1e2130; }

/* JSON viewer */
.stJson { background: #13161f !important; }

div[data-testid="stAlert"] {
    border-radius: 8px;
}
</style>
""", unsafe_allow_html=True)


# ── Sidebar — Config ───────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ API Config")
    base_url = st.text_input("Base URL", value="http://localhost:8000", placeholder="http://host:port")
    prefix   = "/api/llm_mapping"

    st.markdown("---")
    st.markdown("## 🔧 Common Params")
    client_name  = st.text_input("Client Name",  value="HDFC Bank")
    process_name = st.text_input("Process Name", value="COMBINED")
    master_id    = st.number_input("Master ID", min_value=1, value=1, step=1)

    st.markdown("---")
    st.markdown("## 🔬 Engine Flags")
    use_fuzzy      = st.toggle("Fuzzy Matching",      value=True)
    use_embeddings = st.toggle("Embedding Matching",  value=False)
    use_llm        = st.toggle("LLM Matching",        value=True)

    st.markdown("---")
    st.markdown("## 💾 DB Write")
    save_to_db     = st.toggle("Save to DB",       value=False)
    skip_unmatched = st.toggle("Skip Unmatched",   value=False)


# ── Header ─────────────────────────────────────────────────────────────────────
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.markdown("# 🔗 LP Field Mapping")
    st.markdown("<p style='color:#5b6080;margin-top:-0.5rem;font-size:0.9rem;'>Test dashboard · Dvara Gateway</p>", unsafe_allow_html=True)
with col_h2:
    if st.button("🏥 Health Check"):
        try:
            r = requests.get(f"{base_url}{prefix}/health", timeout=5)
            if r.status_code == 200:
                st.success("API is online ✓")
            else:
                st.error(f"HTTP {r.status_code}")
        except Exception as e:
            st.error(f"Unreachable: {e}")

st.markdown("---")


# ── Tabs ───────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📂 References",
    "🎯 Deterministic",
    "⚡ Hybrid + LLM",
    "🚀 Full Pipeline",
    "🌿 Nested / Schema",
])


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — References
# ═══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("Reference Files")

    c1, c2 = st.columns(2)

    with c1:
        st.markdown("#### 📊 Status")
        if st.button("Check Reference Status", key="ref_status"):
            with st.spinner("Checking…"):
                try:
                    r = requests.get(f"{base_url}{prefix}/references/status", timeout=10)
                    data = r.json()
                    ready = data.get("ready", False)
                    if ready:
                        st.success("All reference files present ✓")
                    else:
                        st.warning("Some reference files missing")
                    for fname, info in data.get("files", {}).items():
                        icon = "✅" if info["exists"] else "❌"
                        size = f"{info['size_kb']} KB" if info["size_kb"] else "—"
                        st.markdown(f"`{icon}` **{fname}** — {size}")
                except Exception as e:
                    st.error(str(e))

    with c2:
        st.markdown("#### 🔨 Build References")
        putm_override    = st.text_input("PUTM Table Override (optional)")
        mapping_override = st.text_input("Mapping Table Override (optional)")

        if st.button("Build References", key="build_refs"):
            params = {}
            if putm_override:    params["putm_table_override"]    = putm_override
            if mapping_override: params["mapping_table_override"] = mapping_override

            with st.spinner("Building references from DB… this may take a moment"):
                try:
                    r = requests.post(
                        f"{base_url}{prefix}/references/build",
                        params=params,
                        timeout=120,
                    )
                    if r.status_code == 200:
                        st.success("References built successfully ✓")
                        st.json(r.json().get("data", {}))
                    else:
                        st.error(f"Error {r.status_code}: {r.text}")
                except Exception as e:
                    st.error(str(e))


# ═══════════════════════════════════════════════════════════════════════════════
# Shared result renderer
# ═══════════════════════════════════════════════════════════════════════════════
def _confidence_chip(conf: float) -> str:
    if conf >= 0.90:
        return f'<span class="chip chip-green">{conf:.2f}</span>'
    elif conf >= 0.80:
        return f'<span class="chip chip-blue">{conf:.2f}</span>'
    elif conf >= 0.70:
        return f'<span class="chip chip-yellow">{conf:.2f}</span>'
    return f'<span class="chip chip-red">{conf:.2f}</span>'


def _review_chip(needs_review: bool) -> str:
    if needs_review:
        return '<span class="chip chip-yellow">REVIEW</span>'
    return '<span class="chip chip-green">OK</span>'


def render_mappings(mappings: list, show_stats: bool = True, stats: dict = None):
    if not mappings:
        st.info("No mappings returned.")
        return

    if show_stats and stats:
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Total Fields",   stats.get("total_fields", 0))
        m2.metric("Matched",        stats.get("matched", 0))
        m3.metric("Unmatched",      stats.get("unmatched", 0))
        m4.metric("Match Rate",     f"{stats.get('match_rate_pct', 0)}%")
        m5.metric("Avg Confidence", f"{stats.get('avg_confidence', 0):.2f}")

        st.markdown("##### Match Type Breakdown")
        breakdown = stats.get("by_match_type", {})
        if breakdown:
            bd_df = pd.DataFrame([{"Type": k, "Count": v} for k, v in breakdown.items()])
            st.bar_chart(bd_df.set_index("Type"))

    df = pd.DataFrame(mappings)

    # Reorder for readability
    priority_cols = [
        "partner_field", "matched_excel_key", "json_key",
        "confidence", "match_type", "entity",
        "needs_review", "reasoning", "winning_engine",
    ]
    cols = [c for c in priority_cols if c in df.columns] + \
           [c for c in df.columns if c not in priority_cols]
    df = df[cols]

    # Colour confidence column
    def _style_conf(val):
        if not isinstance(val, (int, float)):
            return ""
        if val >= 0.90: return "background-color:#0f2e1e;color:#4ade80"
        if val >= 0.80: return "background-color:#0f1e2e;color:#60a5fa"
        if val >= 0.70: return "background-color:#2e2410;color:#fbbf24"
        return "background-color:#2e1010;color:#f87171"

    styled = df.style.applymap(_style_conf, subset=["confidence"]) if "confidence" in df.columns else df.style
    st.dataframe(styled, use_container_width=True, height=420)

    # Download
    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    st.download_button(
        "⬇️ Download as Excel",
        data=buf.getvalue(),
        file_name="mappings.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Deterministic
# ═══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("Deterministic Matching")
    st.markdown("<p style='color:#5b6080;font-size:0.85rem;'>Phase 1 only — alias + exact rule-based matching. Fast, zero LLM cost.</p>", unsafe_allow_html=True)

    file_det    = st.file_uploader("Upload partner field file (Excel / CSV)", key="det_upload")
    sheet_filter = st.text_input("Sheet filter (optional)", key="det_sheet")

    if st.button("▶ Run Deterministic", key="run_det") and file_det:
        with st.spinner("Running deterministic engine…"):
            t0 = time.time()
            try:
                resp = requests.post(
                    f"{base_url}{prefix}/mapping/deterministic",
                    data={
                        "client_name":   client_name,
                        "process_name":  process_name,
                        "sheet_filter":  sheet_filter or "",
                    },
                    files={"file": (file_det.name, file_det.getvalue(), file_det.type)},
                    timeout=120,
                )
                elapsed = round(time.time() - t0, 2)

                if resp.status_code == 200:
                    data    = resp.json()
                    st.success(f"Completed in {elapsed}s")

                    mappings = [m for m in data.get("mappings", [])]
                    stats_r  = data.get("stats", {})
                    unmatched = data.get("unmatched_fields", [])
                    llm_cnt   = data.get("llm_prompts_count", 0)

                    col_a, col_b = st.columns([3, 1])
                    with col_a:
                        render_mappings(mappings, stats=stats_r)
                    with col_b:
                        st.markdown("##### Unmatched Fields")
                        st.metric("Count", len(unmatched))
                        st.metric("LLM Prompts Ready", llm_cnt)
                        if unmatched:
                            st.dataframe(
                                pd.DataFrame(unmatched)[["partner_field", "entity"]],
                                use_container_width=True,
                            )
                else:
                    st.error(f"HTTP {resp.status_code}: {resp.text}")

            except Exception as e:
                st.error(str(e))
    elif st.button("▶ Run Deterministic", key="run_det_nofile"):
        st.warning("Please upload a file first.")


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 — Hybrid + LLM
# ═══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("Hybrid + LLM Matching")
    st.markdown("<p style='color:#5b6080;font-size:0.85rem;'>Phase 1 + 2 — deterministic → fuzzy → embeddings → LLM. Engine flags from sidebar.</p>", unsafe_allow_html=True)

    file_hyb = st.file_uploader("Upload partner field file", key="hyb_upload")

    if st.button("▶ Run Hybrid + LLM", key="run_hyb") and file_hyb:
        with st.spinner("Running full hybrid pipeline… (LLM calls may take a while)"):
            t0 = time.time()
            try:
                resp = requests.post(
                    f"{base_url}{prefix}/mapping/hybrid-llm",
                    data={
                        "client_name":    client_name,
                        "process_name":   process_name,
                        "use_fuzzy":      str(use_fuzzy).lower(),
                        "use_embeddings": str(use_embeddings).lower(),
                        "use_llm":        str(use_llm).lower(),
                        "master_id":      str(master_id),
                        "save_to_db":     str(save_to_db).lower(),
                        "skip_unmatched": str(skip_unmatched).lower(),
                    },
                    files={"file": (file_hyb.name, file_hyb.getvalue(), file_hyb.type)},
                    timeout=600,
                )
                elapsed = round(time.time() - t0, 2)

                if resp.status_code == 200:
                    data = resp.json()
                    st.success(f"Completed in {elapsed}s")

                    breakdown = data.get("engine_breakdown", {})
                    b1, b2, b3, b4 = st.columns(4)
                    b1.metric("Deterministic", breakdown.get("deterministic", "—"))
                    b2.metric("Fuzzy",         breakdown.get("fuzzy", 0))
                    b3.metric("Embedding",     breakdown.get("embedding", 0))
                    b4.metric("LLM",           breakdown.get("llm", 0))

                    st.markdown("---")
                    render_mappings(data.get("mappings", []), stats=data.get("stats", {}))
                else:
                    st.error(f"HTTP {resp.status_code}: {resp.text}")

            except Exception as e:
                st.error(str(e))


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 4 — Full Pipeline
# ═══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.subheader("Full Pipeline")
    st.markdown("<p style='color:#5b6080;font-size:0.85rem;'>All phases → ZIP: Excel + nested mapping JSON + schema JSON. Optional DB write.</p>", unsafe_allow_html=True)

    file_fp      = st.file_uploader("Upload partner field file", key="fp_upload")
    sheet_fp     = st.text_input("Sheet filter (optional)", key="fp_sheet")

    if st.button("🚀 Run Full Pipeline", key="run_fp") and file_fp:
        with st.spinner("Running full pipeline… (may take several minutes)"):
            t0 = time.time()
            try:
                resp = requests.post(
                    f"{base_url}{prefix}/mapping/full-pipeline",
                    data={
                        "client_name":    client_name,
                        "process_name":   process_name,
                        "use_fuzzy":      str(use_fuzzy).lower(),
                        "use_embeddings": str(use_embeddings).lower(),
                        "use_llm":        str(use_llm).lower(),
                        "sheet_filter":   sheet_fp or "",
                        "master_id":      str(master_id),
                        "save_to_db":     str(save_to_db).lower(),
                        "skip_unmatched": str(skip_unmatched).lower(),
                    },
                    files={"file": (file_fp.name, file_fp.getvalue(), file_fp.type)},
                    timeout=900,
                )
                elapsed = round(time.time() - t0, 2)

                if resp.status_code == 200:
                    # DB header stats
                    inserted = resp.headers.get("X-DB-Inserted", "—")
                    skipped  = resp.headers.get("X-DB-Skipped",  "—")
                    errors   = resp.headers.get("X-DB-Errors",   "—")

                    st.success(f"Pipeline complete in {elapsed}s")
                    d1, d2, d3 = st.columns(3)
                    d1.metric("DB Inserted", inserted)
                    d2.metric("DB Skipped",  skipped)
                    d3.metric("DB Errors",   errors)

                    # Unpack ZIP in memory and preview
                    zb = zipfile.ZipFile(io.BytesIO(resp.content))
                    st.markdown("##### 📦 ZIP Contents")

                    for name in zb.namelist():
                        st.markdown(f"`{name}`")

                    # Preview excel
                    excel_files = [n for n in zb.namelist() if n.endswith(".xlsx")]
                    if excel_files:
                        xl_bytes = zb.read(excel_files[0])
                        df_prev  = pd.read_excel(io.BytesIO(xl_bytes))
                        with st.expander(f"Preview: {excel_files[0]}", expanded=True):
                            st.dataframe(df_prev, use_container_width=True, height=350)

                    # Preview nested mapping
                    json_files = [n for n in zb.namelist() if "nested" in n]
                    if json_files:
                        with st.expander(f"Preview: {json_files[0]}"):
                            st.json(json.loads(zb.read(json_files[0]).decode()))

                    # Full ZIP download
                    st.download_button(
                        "⬇️ Download Full ZIP",
                        data=resp.content,
                        file_name=f"{client_name}_{process_name}_outputs.zip",
                        mime="application/zip",
                    )
                else:
                    st.error(f"HTTP {resp.status_code}: {resp.text}")

            except Exception as e:
                st.error(str(e))


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 5 — Nested / Schema
# ═══════════════════════════════════════════════════════════════════════════════
with tab5:
    st.subheader("Nested Mapping & Schema Generator")
    st.markdown("<p style='color:#5b6080;font-size:0.85rem;'>Convert flat mappings to a nested JSON tree or generate a blank schema.</p>", unsafe_allow_html=True)

    st.markdown("##### Paste or upload flat mappings JSON")
    sample_payload = json.dumps({
        "client_name": client_name,
        "mappings": [
            {"client_column": "Date_of_Birth", "json_key": "loanAccount.customer.dateOfBirth", "entity": "APPLICANT"},
            {"client_column": "First_Name",    "json_key": "loanAccount.customer.firstName",   "entity": "APPLICANT"},
            {"client_column": "Loan_Amount",   "json_key": "loanAccount.loanAmount",           "entity": "LOAN"},
        ]
    }, indent=2)

    json_input = st.text_area(
        "Mappings JSON (LOSJsonRequest format)",
        value=sample_payload,
        height=220,
        key="nested_json_input",
    )

    c_nest, c_schema = st.columns(2)

    with c_nest:
        if st.button("🌿 Generate Nested Mapping", key="gen_nested"):
            try:
                payload = json.loads(json_input)
                with st.spinner("Building nested mapping…"):
                    resp = requests.post(
                        f"{base_url}{prefix}/generate-nested-mapping",
                        json=payload,
                        timeout=60,
                    )
                if resp.status_code == 200:
                    result = resp.json()
                    st.success(f"Done — {result.get('mapped_count', '?')} mapped fields")
                    st.json(result.get("los_json", {}))
                    st.download_button(
                        "⬇️ Download nested_mapping.json",
                        data=json.dumps(result.get("los_json", {}), indent=2),
                        file_name="nested_mapping.json",
                        mime="application/json",
                    )
                else:
                    st.error(f"HTTP {resp.status_code}: {resp.text}")
            except json.JSONDecodeError:
                st.error("Invalid JSON — check the input above.")
            except Exception as e:
                st.error(str(e))

    with c_schema:
        if st.button("📐 Generate Schema", key="gen_schema"):
            try:
                payload = json.loads(json_input)
                with st.spinner("Building schema…"):
                    resp = requests.post(
                        f"{base_url}{prefix}/generate-schema",
                        json=payload,
                        timeout=60,
                    )
                if resp.status_code == 200:
                    result = resp.json()
                    st.success(f"Done — {result.get('mapped_count', '?')} paths")
                    st.json(result.get("los_schema", {}))
                    st.download_button(
                        "⬇️ Download schema.json",
                        data=json.dumps(result.get("los_schema", {}), indent=2),
                        file_name="schema.json",
                        mime="application/json",
                    )
                else:
                    st.error(f"HTTP {resp.status_code}: {resp.text}")
            except json.JSONDecodeError:
                st.error("Invalid JSON — check the input above.")
            except Exception as e:
                st.error(str(e))


# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<p style='text-align:center;color:#2e3248;font-size:0.72rem;letter-spacing:0.1em;'>"
    "LP FIELD MAPPING · DVARA GATEWAY · TEST DASHBOARD"
    "</p>",
    unsafe_allow_html=True,
)