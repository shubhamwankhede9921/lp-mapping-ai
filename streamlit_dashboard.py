"""
LP Field Mapping — Streamlit Dashboard (Redesigned)
Run with:  streamlit run streamlit_dashboard.py
"""

import io
import json
import time
import zipfile
from typing import Any, Dict, List, Optional, Tuple

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

# ── Session defaults ───────────────────────────────────────────────────────────
if "nested_payload_hint" not in st.session_state:
    st.session_state["nested_payload_hint"] = ""
if "_nested_last_upload_name" not in st.session_state:
    st.session_state["_nested_last_upload_name"] = None


def _default_los_sample(cname: str, pname: str) -> str:
    return json.dumps(
        {
            "client_name": cname,
            "process_name": pname,
            "mappings": [
                {
                    "client_column": "Date_of_Birth",
                    "json_key": "loanAccount.customer.dateOfBirth",
                    "entity": "APPLICANT",
                },
                {
                    "client_column": "First_Name",
                    "json_key": "loanAccount.customer.firstName",
                    "entity": "APPLICANT",
                },
                {
                    "client_column": "Loan_Amount",
                    "json_key": "loanAccount.loanAmount",
                    "entity": "LOAN",
                },
            ],
        },
        indent=2,
    )


def store_mappings_for_nested_tab(
    mappings: List[Dict[str, Any]],
    client: str,
    process: str,
    source: str,
) -> None:
    payload = {
        "client_name": client,
        "process_name": process,
        "mappings": mappings,
    }
    st.session_state["nested_json_input"] = json.dumps(payload, indent=2)
    st.session_state["nested_payload_hint"] = source
    st.session_state["_nested_last_upload_name"] = None
    st.session_state["_nested_flash"] = (
        "Mappings loaded into **Nested / Schema** — open that tab to generate nested JSON or schema."
    )
    st.rerun()


def _style_confidence_cell(val: Any) -> str:
    if not isinstance(val, (int, float)):
        return ""
    if val >= 0.90:
        return "background-color:#0a2318;color:#4ade80;font-weight:600"
    if val >= 0.80:
        return "background-color:#0d1f3c;color:#60a5fa;font-weight:600"
    if val >= 0.70:
        return "background-color:#2d1f05;color:#fbbf24;font-weight:600"
    return "background-color:#2d0a0a;color:#f87171;font-weight:600"


def _apply_confidence_style(styler: Any) -> Any:
    if "confidence" not in styler.data.columns:
        return styler
    fn = getattr(styler, "map", None) or getattr(styler, "applymap", None)
    if fn:
        return fn(_style_confidence_cell, subset=["confidence"])
    return styler


def _excel_rows_to_mapping_payload(xdf: pd.DataFrame) -> List[Dict[str, Any]]:
    colmap: Dict[str, str] = {}
    for c in xdf.columns:
        key = str(c).strip().lower().replace(" ", "_")
        colmap[key] = str(c)

    def pick(*aliases: str) -> Optional[str]:
        for a in aliases:
            k = a.lower().replace(" ", "_")
            if k in colmap:
                return colmap[k]
        return None

    c_pf = pick("partner_field", "partnerfield")
    c_jk = pick("json_key", "jsonkey")
    c_ent = pick("entity")
    if not c_pf or not c_jk:
        return []

    out: List[Dict[str, Any]] = []
    for _, row in xdf.iterrows():
        pf = row.get(c_pf)
        jk = row.get(c_jk)
        ent = row.get(c_ent) if c_ent else None
        if pd.isna(pf) or not str(pf).strip():
            continue
        out.append(
            {
                "partner_field": str(pf).strip(),
                "json_key": str(jk).strip() if pd.notna(jk) else "",
                "entity": str(ent).strip() if ent is not None and pd.notna(ent) else "OTHER",
            }
        )
    return out


def _unpack_zip_artifacts(
    zip_bytes: bytes,
) -> Tuple[Optional[str], Optional[bytes], Optional[bytes], Optional[bytes], List[str]]:
    zb = zipfile.ZipFile(io.BytesIO(zip_bytes))
    names = zb.namelist()
    excel_name = next((n for n in names if n.lower().endswith(".xlsx")), None)
    nested_name = next((n for n in names if "nested" in n.lower() and n.lower().endswith(".json")), None)
    schema_name = next(
        (n for n in names if "schema" in n.lower() and n.lower().endswith(".json") and "nested" not in n.lower()),
        None,
    )
    xl_b = zb.read(excel_name) if excel_name else None
    nest_b = zb.read(nested_name) if nested_name else None
    sch_b = zb.read(schema_name) if schema_name else None
    return excel_name, xl_b, nest_b, sch_b, names


# ══════════════════════════════════════════════════════════════════════════════
# CUSTOM CSS — Bold dark-luxury aesthetic · LARGE, READABLE, IMPACTFUL
# Fonts: Syne (display 800) + JetBrains Mono (code) + Plus Jakarta Sans (body)
# Palette: Deep navy base · electric indigo accent · violet/teal highlights
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@600;700;800&family=JetBrains+Mono:wght@400;500&family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');

/* ── Root variables ───────────────────────────────────────────── */
:root {
    --bg-base:       #07090f;
    --bg-surface:    #0c0f18;
    --bg-raised:     #10141f;
    --bg-hover:      #151a28;
    --accent-indigo: #6366f1;
    --accent-violet: #8b5cf6;
    --accent-amber:  #f59e0b;
    --accent-teal:   #14b8a6;
    --accent-rose:   #f43f5e;
    --text-primary:  #eef0f8;
    --text-secondary:#9aa3bc;
    --text-muted:    #525c78;
    --border-subtle: rgba(99,102,241,0.13);
    --border-mid:    rgba(99,102,241,0.26);
    --border-strong: rgba(99,102,241,0.45);
    --glow-indigo:   rgba(99,102,241,0.18);
}

/* ── Base typography — BIGGER ─────────────────────────────────── */
html, body, [class*="css"] {
    font-family: 'Plus Jakarta Sans', system-ui, sans-serif !important;
    font-size: 18px !important;
    -webkit-font-smoothing: antialiased;
    color: var(--text-primary);
}

p, li, span, div {
    font-size: 17px;
    line-height: 1.65;
}

label, .stMarkdown p {
    font-size: 16px !important;
    line-height: 1.6 !important;
}

code, pre, .stCode {
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 15px !important;
    letter-spacing: 0.01em;
}

/* ── App shell ────────────────────────────────────────────────── */
.stApp {
    background:
        radial-gradient(ellipse 120% 70% at 0% -8%,   rgba(99,102,241,0.18) 0%, transparent 52%),
        radial-gradient(ellipse 90%  55% at 100% 8%,  rgba(139,92,246,0.13) 0%, transparent 48%),
        radial-gradient(ellipse 70%  45% at 55% 100%, rgba(20,184,166,0.08) 0%, transparent 42%),
        var(--bg-base);
    color: var(--text-primary);
    min-height: 100vh;
}

.block-container {
    padding-top: 1rem !important;
    padding-bottom: 3rem !important;
    max-width: 1500px !important;
}

/* ── Sidebar ──────────────────────────────────────────────────── */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #090c16 0%, #0b0f1a 100%) !important;
    border-right: 1px solid var(--border-subtle) !important;
}

[data-testid="stSidebar"] > div:first-child {
    padding-top: 1.5rem;
}

[data-testid="stSidebar"] .stMarkdown h2 {
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-size: 12px !important;
    font-weight: 700 !important;
    letter-spacing: 0.2em !important;
    text-transform: uppercase;
    color: var(--text-muted) !important;
    margin-bottom: 0.8rem;
    padding-bottom: 0.6rem;
    border-bottom: 1px solid var(--border-subtle);
}

[data-testid="stSidebar"] label {
    color: var(--text-secondary) !important;
    font-size: 15px !important;
    font-weight: 500 !important;
}

[data-testid="stSidebar"] .stTextInput input,
[data-testid="stSidebar"] .stNumberInput input {
    background: var(--bg-raised) !important;
    border: 1px solid var(--border-subtle) !important;
    border-radius: 9px !important;
    color: var(--text-primary) !important;
    font-size: 16px !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    padding: 0.5rem 0.75rem !important;
    transition: border-color 0.2s;
}

[data-testid="stSidebar"] .stTextInput input:focus,
[data-testid="stSidebar"] .stNumberInput input:focus {
    border-color: var(--accent-indigo) !important;
    box-shadow: 0 0 0 3px var(--glow-indigo) !important;
}

[data-testid="stSidebar"] hr {
    border-color: var(--border-subtle) !important;
    margin: 1.1rem 0;
}

/* ── Hero header ──────────────────────────────────────────────── */
.hero-outer {
    position: relative;
    border-radius: 22px;
    padding: 2px;
    background: linear-gradient(135deg, #6366f1, #8b5cf6 50%, #14b8a6);
    margin-bottom: 1.25rem;
    overflow: hidden;
}

.hero-inner {
    border-radius: 20px;
    padding: 2rem 2.5rem 2.2rem;
    background: linear-gradient(140deg, rgba(10,13,22,0.98) 0%, rgba(8,11,20,0.97) 100%);
    position: relative;
    overflow: hidden;
}

.hero-inner::after {
    content: '';
    position: absolute;
    top: -80px; right: -80px;
    width: 320px; height: 320px;
    border-radius: 50%;
    background: radial-gradient(circle, rgba(99,102,241,0.18), transparent 65%);
    pointer-events: none;
}

.hero-inner::before {
    content: '';
    position: absolute;
    bottom: -40px; left: 20%;
    width: 200px; height: 200px;
    border-radius: 50%;
    background: radial-gradient(circle, rgba(20,184,166,0.10), transparent 65%);
    pointer-events: none;
}

.hero-badge {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    padding: 6px 16px;
    border-radius: 999px;
    background: rgba(99,102,241,0.14);
    border: 1px solid rgba(99,102,241,0.32);
    font-family: 'JetBrains Mono', monospace;
    font-size: 13px;
    font-weight: 500;
    letter-spacing: 0.12em;
    color: #a5b4fc;
    margin-bottom: 1rem;
}

.hero-badge-dot {
    width: 7px; height: 7px;
    border-radius: 50%;
    background: #6366f1;
    box-shadow: 0 0 10px #6366f1;
    animation: pulse-dot 2s ease-in-out infinite;
}

@keyframes pulse-dot {
    0%, 100% { opacity: 1; transform: scale(1); }
    50%       { opacity: 0.4; transform: scale(0.75); }
}

.hero-title {
    font-family: 'Syne', sans-serif;
    font-weight: 800;
    font-size: 3.35rem;
    line-height: 1.05;
    letter-spacing: -0.05em;
    margin: 0 0 0.6rem;
    background: linear-gradient(105deg, #ffffff 0%, #c7d2fe 38%, #a78bfa 68%, #5eead4 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}

.hero-sub {
    color: var(--text-secondary);
    font-size: 17px;
    font-weight: 400;
    line-height: 1.6;
    margin: 0;
    max-width: 680px;
}

.hero-sub strong {
    color: #c4b5fd;
    font-weight: 600;
}

/* ── Tab styling ──────────────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"] {
    gap: 5px;
    background: rgba(8, 11, 20, 0.85);
    padding: 6px 7px;
    border-radius: 16px;
    border: 1px solid var(--border-subtle);
}

.stTabs [data-baseweb="tab"] {
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-weight: 600 !important;
    font-size: 15px !important;
    letter-spacing: 0.01em;
    color: var(--text-muted);
    border-radius: 11px !important;
    padding: 0.55rem 1.1rem !important;
    transition: all 0.2s ease;
    border: 1px solid transparent !important;
}

.stTabs [data-baseweb="tab"]:hover {
    color: var(--text-secondary);
    background: var(--bg-hover) !important;
}

.stTabs [aria-selected="true"] {
    color: #e8ebff !important;
    background: linear-gradient(135deg, rgba(99,102,241,0.25), rgba(139,92,246,0.20)) !important;
    border: 1px solid var(--border-mid) !important;
    box-shadow: 0 2px 14px rgba(99,102,241,0.18), inset 0 1px 0 rgba(255,255,255,0.07);
}

/* ── Buttons — bold, larger, impactful ────────────────────────── */
.stButton > button {
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-weight: 700 !important;
    font-size: 16px !important;
    letter-spacing: 0.02em;
    border-radius: 12px !important;
    padding: 0.65rem 1.5rem !important;
    min-height: 44px !important;
    border: none !important;
    background: linear-gradient(135deg, #5b5fef 0%, #7c3aed 100%) !important;
    color: #fff !important;
    box-shadow: 0 4px 22px rgba(99,102,241,0.35), inset 0 1px 0 rgba(255,255,255,0.18);
    transition: all 0.2s ease;
    position: relative;
    overflow: hidden;
}

.stButton > button::after {
    content: '';
    position: absolute;
    inset: 0;
    background: linear-gradient(135deg, rgba(255,255,255,0.08) 0%, transparent 60%);
    border-radius: inherit;
    pointer-events: none;
}

.stButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 10px 32px rgba(99,102,241,0.45), inset 0 1px 0 rgba(255,255,255,0.22) !important;
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%) !important;
}

.stButton > button:active {
    transform: translateY(0px);
    box-shadow: 0 4px 14px rgba(99,102,241,0.30) !important;
}

/* ── Metric cards — taller, bigger numbers ────────────────────── */
[data-testid="metric-container"] {
    background: var(--bg-raised) !important;
    border: 1px solid var(--border-subtle) !important;
    border-radius: 16px !important;
    padding: 1.4rem 1.6rem !important;
    position: relative;
    overflow: hidden;
    transition: border-color 0.2s, transform 0.2s;
}

[data-testid="metric-container"]:hover {
    border-color: var(--border-mid) !important;
    transform: translateY(-2px);
}

[data-testid="metric-container"]::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    background: linear-gradient(90deg, #6366f1, #8b5cf6, #14b8a6);
}

[data-testid="metric-container"] label {
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-size: 13px !important;
    font-weight: 700 !important;
    letter-spacing: 0.15em !important;
    text-transform: uppercase;
    color: var(--text-muted) !important;
}

[data-testid="metric-container"] [data-testid="stMetricValue"] {
    font-family: 'Syne', sans-serif !important;
    font-size: 2.7rem !important;
    font-weight: 800 !important;
    color: #dde1ff !important;
    letter-spacing: -0.04em;
    line-height: 1.1;
}

[data-testid="metric-container"] [data-testid="stMetricDelta"] {
    font-size: 14px !important;
}

/* ── Section headings ─────────────────────────────────────────── */
h4, .stMarkdown h4 {
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    font-size: 1.52rem !important;
    letter-spacing: -0.03em;
    color: var(--text-primary) !important;
    margin-bottom: 0.4rem !important;
}

h5, .stMarkdown h5 {
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-weight: 700 !important;
    font-size: 13px !important;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    color: var(--text-muted) !important;
    margin-bottom: 0.6rem !important;
}

/* ── Callout ──────────────────────────────────────────────────── */
.callout {
    border-left: 3px solid var(--accent-indigo);
    padding: 0.85rem 1.25rem;
    margin: 0.6rem 0 1.3rem;
    background: rgba(99,102,241,0.07);
    border-radius: 0 12px 12px 0;
    font-size: 16px;
    color: var(--text-secondary);
    line-height: 1.65;
}

.callout b, .callout strong {
    color: #c4b5fd;
    font-weight: 600;
}

.callout code {
    background: rgba(99,102,241,0.18);
    border-radius: 5px;
    padding: 2px 7px;
    color: #a5b4fc;
    font-size: 14px;
    font-family: 'JetBrains Mono', monospace;
}

/* ── File uploader ────────────────────────────────────────────── */
[data-testid="stFileUploadDropzone"] {
    background: rgba(10,13,22,0.9) !important;
    border: 2px dashed var(--border-mid) !important;
    border-radius: 16px !important;
    transition: all 0.2s;
    padding: 1.5rem !important;
}

[data-testid="stFileUploadDropzone"]:hover {
    border-color: var(--accent-indigo) !important;
    background: rgba(99,102,241,0.06) !important;
}

/* ── Text areas ───────────────────────────────────────────────── */
.stTextArea textarea {
    background: var(--bg-raised) !important;
    border: 1px solid var(--border-subtle) !important;
    border-radius: 14px !important;
    color: var(--text-primary) !important;
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 15px !important;
    line-height: 1.7 !important;
    padding: 1rem !important;
    transition: border-color 0.2s, box-shadow 0.2s;
    resize: vertical;
}

.stTextArea textarea:focus {
    border-color: var(--accent-indigo) !important;
    box-shadow: 0 0 0 3px var(--glow-indigo) !important;
    outline: none !important;
}

/* ── Text inputs ──────────────────────────────────────────────── */
.stTextInput input {
    background: var(--bg-raised) !important;
    border: 1px solid var(--border-subtle) !important;
    border-radius: 10px !important;
    color: var(--text-primary) !important;
    font-size: 16px !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    padding: 0.55rem 0.9rem !important;
    transition: border-color 0.2s, box-shadow 0.2s;
}

.stTextInput input:focus {
    border-color: var(--accent-indigo) !important;
    box-shadow: 0 0 0 3px var(--glow-indigo) !important;
    outline: none !important;
}

.stTextInput label, .stTextArea label, .stNumberInput label {
    font-size: 15px !important;
    font-weight: 600 !important;
    color: var(--text-secondary) !important;
    letter-spacing: 0.01em;
    margin-bottom: 4px !important;
}

/* ── Expander ─────────────────────────────────────────────────── */
.streamlit-expanderHeader,
[data-testid="stExpander"] summary {
    background: var(--bg-raised) !important;
    border: 1px solid var(--border-subtle) !important;
    border-radius: 13px !important;
    color: var(--text-secondary) !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-size: 16px !important;
    font-weight: 600 !important;
    padding: 0.85rem 1.1rem !important;
    transition: all 0.2s;
}

[data-testid="stExpander"] summary:hover {
    border-color: var(--border-mid) !important;
    color: var(--text-primary) !important;
}

[data-testid="stExpander"] summary:hover {
    border-color: var(--border-mid) !important;
    color: var(--text-primary) !important;
}

[data-testid="stExpander"] [data-testid="stExpanderDetails"] {
    border: 1px solid var(--border-subtle);
    border-top: none;
    border-radius: 0 0 12px 12px;
    background: rgba(13,17,23,0.6);
}

/* ── Dataframe ────────────────────────────────────────────────── */
[data-testid="stDataFrame"] {
    border: 1px solid var(--border-subtle) !important;
    border-radius: 16px !important;
    overflow: hidden;
}

[data-testid="stDataFrame"] thead th {
    background: var(--bg-raised) !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-size: 13px !important;
    font-weight: 700 !important;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: var(--text-muted) !important;
    padding: 10px 14px !important;
}

/* ── Alerts ───────────────────────────────────────────────────── */
div[data-testid="stAlert"] {
    border-radius: 13px !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-size: 16px !important;
    font-weight: 500 !important;
    padding: 0.85rem 1.1rem !important;
}

div[data-testid="stAlert"][data-type="success"] {
    background: rgba(20,184,166,0.09) !important;
    border: 1px solid rgba(20,184,166,0.28) !important;
    color: #4dd9c7 !important;
}

div[data-testid="stAlert"][data-type="error"] {
    background: rgba(244,63,94,0.09) !important;
    border: 1px solid rgba(244,63,94,0.28) !important;
}

div[data-testid="stAlert"][data-type="warning"] {
    background: rgba(245,158,11,0.09) !important;
    border: 1px solid rgba(245,158,11,0.28) !important;
}

div[data-testid="stAlert"][data-type="info"] {
    background: rgba(99,102,241,0.09) !important;
    border: 1px solid rgba(99,102,241,0.28) !important;
}

/* ── Chips & badges ───────────────────────────────────────────── */
.chip {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    padding: 4px 12px;
    border-radius: 999px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 13px;
    font-weight: 500;
    letter-spacing: 0.05em;
}

.chip-indigo { background: rgba(99,102,241,0.15); color: #a5b4fc; border: 1px solid rgba(99,102,241,0.30); }
.chip-teal   { background: rgba(20,184,166,0.13); color: #5eead4; border: 1px solid rgba(20,184,166,0.28); }
.chip-amber  { background: rgba(245,158,11,0.13); color: #fcd34d; border: 1px solid rgba(245,158,11,0.28); }
.chip-accent { background: rgba(99,102,241,0.15); color: #a5b4fc; border: 1px solid rgba(99,102,241,0.30); }

/* ── Section divider ──────────────────────────────────────────── */
hr {
    border: none;
    border-top: 1px solid var(--border-subtle);
    margin: 1.5rem 0;
}

/* ── Code & JSON ──────────────────────────────────────────────── */
.stJson {
    background: var(--bg-raised) !important;
    border: 1px solid var(--border-subtle) !important;
    border-radius: 13px !important;
    font-size: 15px !important;
}

/* ── Bar chart ────────────────────────────────────────────────── */
[data-testid="stVegaLiteChart"] {
    border-radius: 13px;
    overflow: hidden;
}

/* ── Download buttons — distinct from action buttons ─────────── */
[data-testid="stDownloadButton"] > button {
    background: rgba(16,20,31,0.9) !important;
    border: 1px solid var(--border-mid) !important;
    color: #a5b4fc !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-weight: 600 !important;
    font-size: 15px !important;
    border-radius: 11px !important;
    padding: 0.55rem 1.1rem !important;
    min-height: 40px !important;
    box-shadow: none !important;
    transition: all 0.2s;
}

[data-testid="stDownloadButton"] > button:hover {
    background: rgba(99,102,241,0.14) !important;
    border-color: var(--accent-indigo) !important;
    color: #c7d2fe !important;
    transform: translateY(-1px);
    box-shadow: 0 5px 18px rgba(99,102,241,0.20) !important;
}

/* ── Footer ───────────────────────────────────────────────────── */
.dash-footer {
    text-align: center;
    color: var(--text-muted);
    font-family: 'JetBrains Mono', monospace;
    font-size: 13px;
    letter-spacing: 0.22em;
    text-transform: uppercase;
    margin-top: 3.5rem;
    padding-top: 1.5rem;
    border-top: 1px solid var(--border-subtle);
}

/* ── Section header strip ─────────────────────────────────────── */
.section-header {
    display: flex;
    align-items: center;
    gap: 14px;
    margin-bottom: 1.2rem;
}

.section-icon {
    width: 42px; height: 42px;
    border-radius: 11px;
    display: flex; align-items: center; justify-content: center;
    font-size: 23px;
    flex-shrink: 0;
}

.icon-blue   { background: rgba(99,102,241,0.18); border: 1px solid rgba(99,102,241,0.25); }
.icon-teal   { background: rgba(20,184,166,0.18); border: 1px solid rgba(20,184,166,0.25); }
.icon-amber  { background: rgba(245,158,11,0.18); border: 1px solid rgba(245,158,11,0.25); }
.icon-rose   { background: rgba(244,63,94,0.18);  border: 1px solid rgba(244,63,94,0.25);  }
.icon-violet { background: rgba(139,92,246,0.18); border: 1px solid rgba(139,92,246,0.25); }

.section-label {
    font-family: 'Syne', sans-serif;
    font-weight: 700;
    font-size: 1.42rem;
    letter-spacing: -0.03em;
    color: var(--text-primary);
    line-height: 1;
}

.section-label small {
    display: block;
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 15px;
    font-weight: 400;
    letter-spacing: 0.01em;
    color: var(--text-muted);
    margin-top: 3px;
}

/* ── Sidebar logo area ────────────────────────────────────────── */
.sidebar-logo {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 0.5rem 0 1.25rem;
    margin-bottom: 0.5rem;
    border-bottom: 1px solid var(--border-subtle);
}

.logo-mark {
    width: 36px; height: 36px;
    border-radius: 10px;
    background: linear-gradient(135deg, #6366f1, #8b5cf6);
    display: flex; align-items: center; justify-content: center;
    font-size: 19px;
    box-shadow: 0 4px 14px rgba(99,102,241,0.38);
}

.logo-text {
    font-family: 'Syne', sans-serif;
    font-weight: 700;
    font-size: 17px;
    letter-spacing: -0.02em;
    color: var(--text-primary);
    line-height: 1;
}

.logo-text span {
    display: block;
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 12px;
    font-weight: 500;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    color: var(--text-muted);
    margin-top: 3px;
}

/* ── Engine pill row ──────────────────────────────────────────── */
.engine-row {
    display: flex;
    gap: 6px;
    flex-wrap: wrap;
    margin-bottom: 1rem;
}

.engine-pill {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    padding: 5px 11px;
    border-radius: 999px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 13px;
    font-weight: 500;
    letter-spacing: 0.05em;
    border: 1px solid;
}

.pill-det  { background: rgba(20,184,166,0.12); color: #4dd9c7; border-color: rgba(20,184,166,0.28); }
.pill-fuzz { background: rgba(99,102,241,0.12); color: #a5b4fc; border-color: rgba(99,102,241,0.28); }
.pill-emb  { background: rgba(245,158,11,0.12); color: #fcd34d; border-color: rgba(245,158,11,0.28); }
.pill-llm  { background: rgba(244,63,94,0.12);  color: #fb7185; border-color: rgba(244,63,94,0.28);  }

/* ── Selectbox ────────────────────────────────────────────────── */
[data-testid="stSidebar"] [data-baseweb="select"] > div,
[data-baseweb="select"] > div {
    background: var(--bg-raised) !important;
    border: 1px solid var(--border-subtle) !important;
    border-radius: 9px !important;
}

/* ── Number input ─────────────────────────────────────────────── */
.stNumberInput [data-baseweb="input"] {
    background: var(--bg-raised) !important;
    border: 1px solid var(--border-subtle) !important;
    border-radius: 9px !important;
}

/* ── Toggle label ─────────────────────────────────────────────── */
[data-testid="stToggleSwitch"] label,
[data-testid="stToggleSwitch"] p {
    font-size: 16px !important;
    color: var(--text-secondary) !important;
    font-weight: 500 !important;
}

/* ── Checkbox ─────────────────────────────────────────────────── */
[data-testid="stCheckbox"] label, [data-testid="stCheckbox"] p {
    font-size: 16px !important;
    color: var(--text-secondary) !important;
}

/* ── Caption / small text ─────────────────────────────────────── */
.stCaption, [data-testid="stCaptionContainer"] {
    font-size: 15px !important;
    color: var(--text-muted) !important;
}

/* ── Info text in file uploader ───────────────────────────────── */
[data-testid="stFileUploadDropzone"] p,
[data-testid="stFileUploadDropzone"] span {
    font-size: 16px !important;
    color: var(--text-secondary) !important;
}

/* ── Spinner text ─────────────────────────────────────────────── */
[data-testid="stSpinner"] p {
    font-size: 16px !important;
    color: var(--text-secondary) !important;
}
</style>
""",
    unsafe_allow_html=True,
)


# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(
        """
        <div class="sidebar-logo">
            <div class="logo-mark" style="width:36px;height:36px;border-radius:10px;">🔗</div>
            <div class="logo-text" style="font-size:18px;">LP Mapping<span>Dvara Gateway</span></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("## ⚙ API Config")
    base_url = st.text_input("Base URL", value="http://localhost:8000", placeholder="http://host:port")
    prefix = "/api/llm_mapping"

    st.markdown("---")
    st.markdown("## 🔧 Common Params")
    client_name = st.text_input("Client Name", value="HDFC Bank")
    process_name = st.text_input("Process Name", value="COMBINED")
    master_id = st.number_input("Master ID", min_value=1, value=1, step=1)

    st.markdown("---")
    st.markdown("## 🔬 Engine Flags")
    use_fuzzy = st.toggle("Fuzzy Matching", value=True)
    use_embeddings = st.toggle("Embedding Matching", value=False)
    use_llm = st.toggle("LLM Matching", value=True)

    st.markdown("---")
    st.markdown("## 🧬 Full pipeline")
    use_loanparameter_refinement = st.toggle(
        "Loan parameter refinement",
        value=True,
        help="Remap deterministic LOANPARAMETER* rows via the refinement gateway when configured.",
    )
    use_llm_entity_classifier = st.toggle(
        "LLM entity classifier",
        value=False,
        help="If on and ENTITY_CLASSIFIER_GATEWAY_URL is set in .env, classify APPLICANT/LOAN/FEE/… before matching; otherwise use built-in heuristics.",
    )
    include_build_references = st.toggle(
        "Include build references in ZIP",
        value=True,
        help="Calls the same DB extract + build as POST /references/build before mapping, then adds references/*.json to the ZIP.",
    )

    st.markdown("---")
    st.markdown("## 💾 DB Write")
    save_to_db = st.toggle("Save to DB", value=False)
    skip_unmatched = st.toggle("Skip Unmatched", value=False)

    # Active engines display
    st.markdown("---")
    active = []
    if True:        active.append('<span class="engine-pill pill-det">● DET</span>')
    if use_fuzzy:   active.append('<span class="engine-pill pill-fuzz">● FUZZY</span>')
    if use_embeddings: active.append('<span class="engine-pill pill-emb">● EMB</span>')
    if use_llm:     active.append('<span class="engine-pill pill-llm">● LLM</span>')
    if use_llm_entity_classifier:
        active.append('<span class="engine-pill pill-llm">● ENT</span>')
    st.markdown(f'<div class="engine-row">{"".join(active)}</div>', unsafe_allow_html=True)


# ── Hero header ────────────────────────────────────────────────────────────────
h_left, h_right = st.columns([5, 1])

with h_left:
    st.markdown(
        f"""
        <div class="hero-outer">
          <div class="hero-inner">
            <div class="hero-badge">
              <span class="hero-badge-dot"></span>
              LP FIELD MAPPING &nbsp;·&nbsp; TEST CONSOLE
            </div>
            <p class="hero-title">Field Mapping Studio</p>
            <p class="hero-sub">
              Map partner columns → LOS paths with <strong>deterministic</strong>, <strong>fuzzy</strong>, 
              <strong>embedding</strong> &amp; <strong>LLM</strong> engines &nbsp;·&nbsp; 
              Generate nested JSON and schema in one flow
            </p>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with h_right:
    st.markdown("<br><br>", unsafe_allow_html=True)
    if st.button("🏥 Health", use_container_width=True):
        try:
            r = requests.get(f"{base_url}{prefix}/health", timeout=5)
            if r.status_code == 200:
                st.success("API online ✓")
            else:
                st.error(f"HTTP {r.status_code}")
        except Exception as e:
            st.error(f"Unreachable: {e}")

_flash = st.session_state.pop("_nested_flash", "") or ""
if _flash:
    st.success(_flash)

st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)


# ── Tabs ───────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs(
    [
        "📂  References",
        "🎯  Deterministic",
        "⚡  Hybrid + LLM",
        "🚀  Full Pipeline",
        "🌿  Nested / Schema",
    ]
)


# ══════════════════════════════════════════════════════════════════════════════
# Shared result renderer
# ══════════════════════════════════════════════════════════════════════════════
def render_mappings(mappings: list, show_stats: bool = True, stats: dict = None):
    if not mappings:
        st.info("No mappings returned.")
        return

    if show_stats and stats:
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Total fields", stats.get("total_fields", 0))
        m2.metric("Matched", stats.get("matched", 0))
        m3.metric("Unmatched", stats.get("unmatched", 0))
        m4.metric("Match rate", f"{stats.get('match_rate_pct', 0)}%")
        m5.metric("Avg confidence", f"{stats.get('avg_confidence', 0):.2f}")

        st.markdown("##### Match type breakdown")
        breakdown = stats.get("by_match_type", {})
        if breakdown:
            bd_df = pd.DataFrame([{"Type": k, "Count": v} for k, v in breakdown.items()])
            st.bar_chart(bd_df.set_index("Type"))

    df = pd.DataFrame(mappings)
    priority_cols = [
        "partner_field", "matched_excel_key", "json_key",
        "confidence", "match_type", "entity", "needs_review",
        "reasoning", "winning_engine",
    ]
    cols = [c for c in priority_cols if c in df.columns] + [c for c in df.columns if c not in priority_cols]
    df = df[cols]
    styled = _apply_confidence_style(df.style)
    st.dataframe(styled, use_container_width=True, height=420)

    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    st.download_button(
        "⬇️  Download Excel",
        data=buf.getvalue(),
        file_name="mappings.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — References
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown(
        """
        <div class="section-header">
          <div class="section-icon icon-blue">📂</div>
          <div class="section-label">Reference Files<small>Check disk status · rebuild alias registry from DB</small></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p class="callout">Pull status from disk or rebuild <code>alias_registry</code> / dictionaries from the database.</p>',
        unsafe_allow_html=True,
    )

    c1, c2 = st.columns(2)

    with c1:
        if st.button("Check reference status", key="ref_status", use_container_width=True):
            with st.spinner("Checking…"):
                try:
                    r = requests.get(f"{base_url}{prefix}/references/status", timeout=10)
                    data = r.json()
                    ready = data.get("ready", False)
                    if ready:
                        st.success("All reference files present ✓")
                    else:
                        st.warning("Some reference files are missing")
                    for fname, info in data.get("files", {}).items():
                        icon = "✅" if info["exists"] else "❌"
                        size = f"{info['size_kb']} KB" if info["size_kb"] else "—"
                        st.markdown(f"`{icon}` **{fname}** — {size}")
                except Exception as e:
                    st.error(str(e))

    with c2:
        putm_override = st.text_input("PUTM table override (optional)")
        mapping_override = st.text_input("Mapping table override (optional)")
        if st.button("Build references", key="build_refs", use_container_width=True):
            params = {}
            if putm_override:
                params["putm_table_override"] = putm_override
            if mapping_override:
                params["mapping_table_override"] = mapping_override
            with st.spinner("Building references from DB…"):
                try:
                    r = requests.post(
                        f"{base_url}{prefix}/references/build",
                        params=params,
                        timeout=120,
                    )
                    if r.status_code == 200:
                        st.success("References built ✓")
                        st.json(r.json().get("data", {}))
                    else:
                        st.error(f"Error {r.status_code}: {r.text}")
                except Exception as e:
                    st.error(str(e))


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Deterministic
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown(
        """
        <div class="section-header">
          <div class="section-icon icon-teal">🎯</div>
          <div class="section-label">Deterministic Matching<small>Phase 1 — alias and rule-based · fast, no LLM cost</small></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p class="callout">Phase 1 — alias and rule-based matching. Fast, no LLM cost. Use the button below to send results to <b>Nested / Schema</b>.</p>',
        unsafe_allow_html=True,
    )

    file_det = st.file_uploader("Partner field file (Excel / CSV)", key="det_upload")
    sheet_filter = st.text_input("Sheet filter (optional)", key="det_sheet")
    run_det = st.button("▶  Run deterministic", key="run_det", type="primary")

    if run_det:
        if not file_det:
            st.warning("Upload a file first.")
        else:
            with st.spinner("Running deterministic engine…"):
                t0 = time.time()
                try:
                    resp = requests.post(
                        f"{base_url}{prefix}/mapping/deterministic",
                        data={
                            "client_name": client_name,
                            "process_name": process_name,
                            "sheet_filter": sheet_filter or "",
                            "use_llm_entity_classifier": str(use_llm_entity_classifier).lower(),
                        },
                        files={"file": (file_det.name, file_det.getvalue(), file_det.type)},
                        timeout=120,
                    )
                    elapsed = round(time.time() - t0, 2)
                    if resp.status_code == 200:
                        data = resp.json()
                        st.success(f"Completed in {elapsed}s ✓")
                        mappings = [m for m in data.get("mappings", [])]
                        stats_r = data.get("stats", {})
                        unmatched = data.get("unmatched_fields", [])
                        llm_cnt = data.get("llm_prompts_count", 0)
                        col_a, col_b = st.columns([3, 1])
                        with col_a:
                            render_mappings(mappings, stats=stats_r)
                        with col_b:
                            st.markdown("##### Unmatched")
                            st.metric("Count", len(unmatched))
                            st.metric("LLM prompts ready", llm_cnt)
                            if unmatched:
                                st.dataframe(
                                    pd.DataFrame(unmatched)[["partner_field", "entity"]],
                                    use_container_width=True,
                                )
                        if mappings and st.button("➡️  Send to Nested / Schema", key="det_to_nested"):
                            store_mappings_for_nested_tab(mappings, client_name, process_name, "deterministic")
                    else:
                        st.error(f"HTTP {resp.status_code}: {resp.text}")
                except Exception as e:
                    st.error(str(e))


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — Hybrid + LLM
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown(
        """
        <div class="section-header">
          <div class="section-icon icon-amber">⚡</div>
          <div class="section-label">Hybrid + LLM Pipeline<small>Deterministic → Fuzzy → Embeddings → LLM</small></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p class="callout">Deterministic → fuzzy → embeddings → LLM. Sidebar toggles apply. Chain to nested JSON when finished.</p>',
        unsafe_allow_html=True,
    )

    file_hyb = st.file_uploader("Partner field file", key="hyb_upload")
    run_hyb = st.button("▶  Run hybrid + LLM", key="run_hyb", type="primary")

    if run_hyb:
        if not file_hyb:
            st.warning("Upload a file first.")
        else:
            with st.spinner("Running hybrid pipeline… (LLM may take a while)"):
                t0 = time.time()
                try:
                    resp = requests.post(
                        f"{base_url}{prefix}/mapping/hybrid-llm",
                        data={
                            "client_name": client_name,
                            "process_name": process_name,
                            "use_fuzzy": str(use_fuzzy).lower(),
                            "use_embeddings": str(use_embeddings).lower(),
                            "use_llm": str(use_llm).lower(),
                            "use_loanparameter_refinement": str(use_loanparameter_refinement).lower(),
                            "use_llm_entity_classifier": str(use_llm_entity_classifier).lower(),
                            "master_id": str(master_id),
                            "save_to_db": str(save_to_db).lower(),
                            "skip_unmatched": str(skip_unmatched).lower(),
                        },
                        files={"file": (file_hyb.name, file_hyb.getvalue(), file_hyb.type)},
                        timeout=600,
                    )
                    elapsed = round(time.time() - t0, 2)
                    if resp.status_code == 200:
                        data = resp.json()
                        st.success(f"Completed in {elapsed}s ✓")
                        breakdown = data.get("engine_breakdown", {})
                        b1, b2, b3, b4 = st.columns(4)
                        b1.metric("Deterministic", breakdown.get("deterministic", "—"))
                        b2.metric("Fuzzy", breakdown.get("fuzzy", 0))
                        b3.metric("Embedding", breakdown.get("embedding", 0))
                        b4.metric("LLM", breakdown.get("llm", 0))
                        st.markdown("---")
                        render_mappings(data.get("mappings", []), stats=data.get("stats", {}))
                        maps = data.get("mappings", [])
                        if maps and st.button("➡️  Send to Nested / Schema", key="hyb_to_nested"):
                            store_mappings_for_nested_tab(maps, client_name, process_name, "hybrid_llm")
                    else:
                        st.error(f"HTTP {resp.status_code}: {resp.text}")
                except Exception as e:
                    st.error(str(e))


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — Full Pipeline
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown(
        """
        <div class="section-header">
          <div class="section-icon icon-rose">🚀</div>
          <div class="section-label">Full Pipeline<small>One run · ZIP artifact with Excel + nested JSON + schema</small></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p class="callout">One run produces a ZIP: Excel workbook, nested mapping JSON, LOS schema JSON, and '
        "(when enabled) reference JSON built from the source DB. Nested and schema are previewed below. Optional DB write uses sidebar toggles.</p>",
        unsafe_allow_html=True,
    )

    file_fp = st.file_uploader("Partner field file", key="fp_upload")
    sheet_fp = st.text_input("Sheet filter (optional)", key="fp_sheet")

    if st.button("🚀  Run full pipeline", key="run_fp", type="primary") and file_fp:
        with st.spinner("Running full pipeline… (may take several minutes)"):
            t0 = time.time()
            try:
                resp = requests.post(
                    f"{base_url}{prefix}/mapping/full-pipeline",
                    data={
                        "client_name": client_name,
                        "process_name": process_name,
                        "use_fuzzy": str(use_fuzzy).lower(),
                        "use_embeddings": str(use_embeddings).lower(),
                        "use_llm": str(use_llm).lower(),
                        "use_loanparameter_refinement": str(use_loanparameter_refinement).lower(),
                        "use_llm_entity_classifier": str(use_llm_entity_classifier).lower(),
                        "sheet_filter": sheet_fp or "",
                        "master_id": str(master_id),
                        "save_to_db": str(save_to_db).lower(),
                        "skip_unmatched": str(skip_unmatched).lower(),
                        "include_build_references": str(include_build_references).lower(),
                    },
                    files={"file": (file_fp.name, file_fp.getvalue(), file_fp.type)},
                    timeout=900,
                )
                elapsed = round(time.time() - t0, 2)
                if resp.status_code == 200:
                    inserted = resp.headers.get("X-DB-Inserted", "—")
                    skipped  = resp.headers.get("X-DB-Skipped",  "—")
                    errors   = resp.headers.get("X-DB-Errors",   "—")
                    ref_putm = resp.headers.get("X-References-PutM-Rows")
                    ref_map = resp.headers.get("X-References-Mapping-Rows")
                    st.success(f"Pipeline finished in {elapsed}s ✓")
                    d1, d2, d3, d4 = st.columns(4)
                    d1.metric("DB inserted", inserted)
                    d2.metric("DB skipped",  skipped)
                    d3.metric("DB errors",   errors)
                    d4.metric("Loan param refine", "On" if use_loanparameter_refinement else "Off")
                    if include_build_references and (ref_putm or ref_map):
                        r1, r2 = st.columns(2)
                        r1.metric("Ref build · PUTM rows", ref_putm or "—")
                        r2.metric("Ref build · mapping rows", ref_map or "—")

                    excel_name, xl_b, nest_b, sch_b, zip_names = _unpack_zip_artifacts(resp.content)

                    st.markdown("##### Package contents")
                    pill = " ".join(
                        f'<span class="chip chip-indigo">{n}</span>' for n in zip_names
                    )
                    st.markdown(pill, unsafe_allow_html=True)

                    safe_c = "".join(c if c.isalnum() or c in "._-" else "_" for c in client_name)[:40]
                    safe_p = "".join(c if c.isalnum() or c in "._-" else "_" for c in process_name)[:40]

                    dl_col1, dl_col2, dl_col3 = st.columns(3)
                    with dl_col1:
                        st.download_button("⬇️  Full ZIP", data=resp.content,
                            file_name=f"{safe_c}_{safe_p}_outputs.zip", mime="application/zip",
                            use_container_width=True)
                    with dl_col2:
                        if nest_b:
                            st.download_button("⬇️  Nested JSON", data=nest_b,
                                file_name=f"nested_mapping_{safe_c}_{safe_p}.json",
                                mime="application/json", use_container_width=True)
                    with dl_col3:
                        if sch_b:
                            st.download_button("⬇️  Schema JSON", data=sch_b,
                                file_name=f"schema_{safe_c}_{safe_p}.json",
                                mime="application/json", use_container_width=True)

                    if excel_name and xl_b:
                        df_prev = pd.read_excel(io.BytesIO(xl_b))
                        with st.expander(f"Preview · {excel_name}", expanded=True):
                            st.dataframe(df_prev, use_container_width=True, height=320)

                    prev_col1, prev_col2 = st.columns(2)
                    with prev_col1:
                        if nest_b:
                            with st.expander("Preview · nested mapping", expanded=False):
                                st.json(json.loads(nest_b.decode("utf-8")))
                    with prev_col2:
                        if sch_b:
                            with st.expander("Preview · schema", expanded=False):
                                st.json(json.loads(sch_b.decode("utf-8")))

                    if nest_b or sch_b:
                        flat_for_nested: List[Dict[str, Any]] = []
                        if xl_b:
                            try:
                                bio = io.BytesIO(xl_b)
                                xls = pd.ExcelFile(bio)
                                sheet = (
                                    "Field Mapping" if "Field Mapping" in xls.sheet_names
                                    else xls.sheet_names[0]
                                )
                                xdf = pd.read_excel(xls, sheet_name=sheet)
                                flat_for_nested = _excel_rows_to_mapping_payload(xdf)
                            except Exception as ex:
                                st.caption(f"Could not parse Excel for handoff: {ex}")
                        if flat_for_nested and st.button(
                            "➡️  Load Excel mappings into Nested / Schema", key="fp_to_nested"
                        ):
                            store_mappings_for_nested_tab(flat_for_nested, client_name, process_name, "full_pipeline_excel")
                else:
                    st.error(f"HTTP {resp.status_code}: {resp.text}")
            except Exception as e:
                st.error(str(e))


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — Nested / Schema
# ══════════════════════════════════════════════════════════════════════════════
with tab5:
    st.markdown(
        """
        <div class="section-header">
          <div class="section-icon icon-violet">🌿</div>
          <div class="section-label">Nested Mapping & Schema<small>Paste JSON · upload file · or load from another tab</small></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p class="callout">Paste JSON, upload a file, or use <b>Send to Nested / Schema</b> from Deterministic, Hybrid, or Full pipeline. '
        "API accepts <code>partner_field</code>/<code>json_key</code> or <code>client_column</code>/<code>lms_column</code>.</p>",
        unsafe_allow_html=True,
    )

    if "nested_json_input" not in st.session_state:
        st.session_state["nested_json_input"] = _default_los_sample(client_name, process_name)

    hint = st.session_state.get("nested_payload_hint") or ""
    if hint:
        st.markdown(
            f'<span class="chip chip-teal">⟡ payload · source: {hint}</span>',
            unsafe_allow_html=True,
        )

    up_json = st.file_uploader("Optional: upload LOSJsonRequest JSON", type=["json"], key="nested_file_up")
    if up_json is not None and up_json.name != st.session_state.get("_nested_last_upload_name"):
        try:
            raw = up_json.getvalue().decode("utf-8")
            json.loads(raw)
            st.session_state["nested_json_input"] = raw
            st.session_state["_nested_last_upload_name"] = up_json.name
            st.session_state["nested_payload_hint"] = f"upload:{up_json.name}"
            st.success("Editor updated from file. ✓")
            st.rerun()
        except Exception:
            st.error("Could not read valid JSON from file.")

    col_reset, col_spacer = st.columns([1, 3])
    with col_reset:
        if st.button("↺  Reset sample", key="reset_nested_sample"):
            st.session_state["nested_json_input"] = _default_los_sample(client_name, process_name)
            st.session_state["nested_payload_hint"] = ""
            st.session_state["_nested_last_upload_name"] = None
            st.rerun()

    st.text_area(
        "Mappings JSON (LOSJsonRequest)",
        height=280,
        key="nested_json_input",
        label_visibility="visible",
    )

    c_nest, c_schema = st.columns(2)

    with c_nest:
        if st.button("🌿  Generate nested mapping", key="gen_nested", use_container_width=True):
            try:
                payload = json.loads(st.session_state["nested_json_input"])
                with st.spinner("Building nested mapping…"):
                    resp = requests.post(
                        f"{base_url}{prefix}/generate-nested-mapping",
                        json=payload,
                        timeout=60,
                    )
                if resp.status_code == 200:
                    result = resp.json()
                    st.success(
                        f"Mapped {result.get('mapped_count', '?')} fields · "
                        f"{result.get('skipped_count', 0)} skipped · "
                        f"{result.get('processing_time_ms', 0)} ms"
                    )
                    with st.expander("Nested LOS JSON", expanded=True):
                        st.json(result.get("los_json", {}))
                    st.download_button(
                        "⬇️  nested_mapping.json",
                        data=json.dumps(result.get("los_json", {}), indent=2, ensure_ascii=False),
                        file_name="nested_mapping.json",
                        mime="application/json",
                        use_container_width=True,
                    )
                else:
                    st.error(f"HTTP {resp.status_code}: {resp.text}")
            except json.JSONDecodeError:
                st.error("Invalid JSON in the editor.")
            except Exception as e:
                st.error(str(e))

    with c_schema:
        if st.button("📐  Generate schema", key="gen_schema", use_container_width=True):
            try:
                payload = json.loads(st.session_state["nested_json_input"])
                with st.spinner("Building schema…"):
                    resp = requests.post(
                        f"{base_url}{prefix}/generate-schema",
                        json=payload,
                        timeout=60,
                    )
                if resp.status_code == 200:
                    result = resp.json()
                    st.success(
                        f"{result.get('mapped_count', '?')} paths · "
                        f"{result.get('skipped_count', 0)} skipped · "
                        f"{result.get('processing_time_ms', 0)} ms"
                    )
                    with st.expander("LOS schema (null leaves)", expanded=True):
                        st.json(result.get("los_schema", {}))
                    st.download_button(
                        "⬇️  schema.json",
                        data=json.dumps(result.get("los_schema", {}), indent=2, ensure_ascii=False),
                        file_name="schema.json",
                        mime="application/json",
                        use_container_width=True,
                    )
                else:
                    st.error(f"HTTP {resp.status_code}: {resp.text}")
            except json.JSONDecodeError:
                st.error("Invalid JSON in the editor.")
            except Exception as e:
                st.error(str(e))


# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown(
    '<p class="dash-footer">LP Field Mapping &nbsp;·&nbsp; Dvara Gateway &nbsp;·&nbsp; Test Dashboard</p>',
    unsafe_allow_html=True,
)