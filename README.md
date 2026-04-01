<div align="center">

<h1>
  🚀 <span style="color:#6a11cb;">LP FIELD MAPPING</span>
</h1>

<h3>
  <span style="color:#2575fc;">M A P P I N G &nbsp; S E R V I C E</span>
</h3>

<p>
  <b>✨ Intelligent • Multi-layer • Smart Field Mapping ✨</b><br>
  <sub>Seamless automation for loan partner onboarding</sub>
</p>

<p>
  <img src="https://img.shields.io/badge/AI-Powered-6a11cb?style=for-the-badge">
  <img src="https://img.shields.io/badge/Multi--Layer-2575fc?style=for-the-badge">
  <img src="https://img.shields.io/badge/Automation-00C9A7?style=for-the-badge">
</p>

<br/>

> ⚡ *Mapping partner Excel columns to internal LMS API keys — from days to minutes.*

</div>

**Intelligent, multi-layer field mapping for loan partner onboarding**

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![FastAPI](https://img.shields.io/badge/FastAPI-0.109-009688?style=for-the-badge&logo=fastapi&logoColor=white)](https://fastapi.tiangolo.com)
[![MySQL](https://img.shields.io/badge/MySQL-8.x-4479A1?style=for-the-badge&logo=mysql&logoColor=white)](https://mysql.com)
[![License](https://img.shields.io/badge/License-MIT-F7DF1E?style=for-the-badge)](LICENSE)

<br/>

> *Mapping partner Excel columns to internal LMS API keys — from days to minutes.*

---

## 🧠 What Is This?

When a lending partner onboards, they bring their own Excel field names — `Date_of_Birth`, `Aadhar_Front`, `CoapplicantIncome` — which bear no resemblance to internal API paths like `loanAccount.customer.dateOfBirth`.

**LP Field Mapping** solves this automatically using a cascading intelligence pipeline:

```
Partner Excel ──► [ Deterministic ] ──► [ Fuzzy ] ──► [ Embeddings ] ──► [ LLM ] ──► Mapped Output
                      Layer 1           Layer 2a         Layer 2b          Layer 2c
                    ~70% fields         ~15% fields       ~8% fields        ~7% fields
```

Each unmatched field cascades to the next engine. The LLM gateway (Dvara/Langfuse) handles the hardest cases — so nothing is left unmapped.

---

## ✨ Features

| | Feature | Description |
|---|---|---|
| ⚡ | **Deterministic Engine** | Alias registry with frequency-tiered confidence |
| 🔍 | **Fuzzy Matching** | RapidFuzz token-ratio matching (configurable threshold) |
| 🧬 | **Semantic Embeddings** | Sentence-transformer cosine similarity |
| 🤖 | **LLM Gateway** | Dvara/Langfuse — base prompt owned by platform |
| 📦 | **ZIP Output** | Excel + Nested JSON + blank schema in one download |
| 🗄️ | **DB Write-back** | Upserts results directly into target mapping table |
| 🏷️ | **Entity Routing** | Auto-labels fields as APPLICANT / LOAN / DOCUMENT / GUARANTOR … |
| 🔁 | **Hallucination Guard** | Whitelist filtering — only pre-approved fields pass through |

---

## 🗂️ Project Structure

```
lp-field-mapping/
│
├── 📁 app/
│   ├── 📁 api/
│   │   └── mapping_controller.py      # All FastAPI routes
│   ├── 📁 services/
│   │   ├── mapping_service.py         # Pipeline orchestrator
│   │   ├── llm_service.py             # Dvara gateway client + parser
│   │   ├── fuzzy_engine.py            # RapidFuzz engine
│   │   ├── embedding_engine.py        # Sentence-transformer engine
│   │   └── prompt_builder.py          # Context block builder (task field)
│   ├── 📁 repository/
│   │   ├── database.py                # Source DB extraction
│   │   └── db_writer.py               # Target DB upsert
│   ├── 📁 utils/
│   │   └── los_json_builder.py        # Nested JSON + schema generators
│   └── config.py                      # Pydantic settings
│
├── 📁 scripts/
│   ├── build_references.py            # CLI + importable reference builders
│   ├── matching_engine.py             # Core deterministic logic
│   ├── input_parser.py                # Partner Excel reader
│   ├── post_processor.py              # Confidence scoring & deduplication
│   └── generate_output.py             # Excel formatter
│
├── 📁 references/                     # ⚙️  Auto-generated (git-ignored)
│   ├── field_dictionary.json
│   ├── alias_registry.json
│   └── entity_routing.json
│
├── 📁 outputs/                        # 📤 Per-client ZIPs (git-ignored)
├── .env.example
├── requirements.txt
└── main.py
```

---

## 🚀 Getting Started

### 1 · Clone & Install

```bash
git clone https://github.com/your-org/lp-field-mapping.git
cd lp-field-mapping

python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate

pip install -r requirements.txt
```

### 2 · Configure

```bash
cp .env.example .env
```

Open `.env` and fill in your values:

```dotenv
# ── Source DB (PUTM + generic-mapping tables) ─────────────────────────────────
DB_HOST=localhost
DB_PORT=3306
DB_NAME=financialForms
DB_USER=your_user
DB_PASSWORD=your_password

# ── Target DB (write results back) ────────────────────────────────────────────
TARGET_DB_URL=mysql+pymysql://user:pass@host:3306/target_db
GENERIC_MAPPING_TABLE=generic_excel_upload_definition_fields

# ── LLM Gateway ───────────────────────────────────────────────────────────────
LLM_GATEWAY_URL=https://your-dvara-gateway/endpoint
LLM_GATEWAY_TOKEN=your_bearer_token
LLM_TASK_FIELD=task                  # form-data field name (default: task)

# ── Paths ─────────────────────────────────────────────────────────────────────
REFERENCES_DIR=./references
SCRIPTS_DIR=./scripts
OUTPUT_PATH=./outputs

# ── Thresholds ────────────────────────────────────────────────────────────────
FUZZY_THRESHOLD=85                   # RapidFuzz score 0–100
EMBEDDING_THRESHOLD=0.82             # Cosine similarity 0.0–1.0
REVIEW_THRESHOLD=0.80                # Below this → needs_review = true
```

### 3 · Build References

> **Do this once** before any mapping run. Rebuild whenever source DB tables change.

```bash
# Via API (recommended)
curl -X POST http://localhost:8000/api/llm_mapping/references/build

# Or via CLI (needs DB creds in .env)
python scripts/build_references.py
```

This generates three files in `REFERENCES_DIR/`:

| File | Contents |
|------|----------|
| `field_dictionary.json` | All `excel_key → json_key` mappings, grouped by process role |
| `alias_registry.json` | Partner name variants → internal key, with tier1–tier4 frequency scoring |
| `entity_routing.json` | UI grouping labels → entity types |

### 4 · Run the Server

```bash
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

📖 Interactive docs: **http://localhost:8000/docs**

---

## 🔌 API Reference

> All endpoints live under `/api/llm_mapping`

### `POST /references/build`

Pulls PUTM and generic-mapping tables from source DB and writes reference JSONs.

---

### `POST /mapping/deterministic`

Layer 1 only — alias lookup + rule-based matching.

**Parameters:**
- `file` (xlsx): Partner Excel upload
- `client_name` (str): Client identifier, e.g. "HDFC"
- `process_name` (str): "COMBINED" (default)
- `sheet_filter` (str, optional): Restrict to one sheet name

---

### `POST /mapping/hybrid-llm`

Layers 1 + 2 — Deterministic → Fuzzy → Embeddings → LLM.

**Parameters:**
- `file` (xlsx): Partner Excel upload
- `client_name` (str): Client identifier
- `process_name` (str): "COMBINED" (default)
- `sheet_filter` (str, optional): Restrict to one sheet name
- `use_fuzzy` (bool): true
- `use_embeddings` (bool): false
- `use_llm` (bool): true
- `master_id` (int, optional): FK for DB write
- `save_to_db` (bool): false
- `skip_unmatched` (bool): false

---

### `POST /mapping/full-pipeline` ⭐

**The main endpoint.** All layers → ZIP download containing:

```
📦 {client}_{process}_outputs.zip
 ├── mapping_{client}_{process}.xlsx
 ├── nested_mapping_{client}_{process}.json
 └── schema_{client}_{process}.json
```

Response headers carry DB write stats:

```
X-DB-Inserted   X-DB-Skipped   X-DB-Errors   X-Master-Id
```

---

### `POST /generate-nested-mapping`

Converts flat `{client_column, lms_column}` pairs into a nested LOS JSON tree.

**Example:**
```json
// Input
{ "client_column": "Date_of_Birth", "lms_column": "loanAccount.customer.dateOfBirth" }

// Output
{ "loanAccount": { "customer": { "dateOfBirth": "Date_of_Birth" } } }
```

---

### `POST /generate-schema`

Same input — returns a blank schema with all leaves set to `null`.

---

### `GET /health` & `GET /references/status`

Health check and reference file readiness probe.

---

## ⚙️ How the LLM Gateway Works

The base prompt template lives on the Dvara/Langfuse platform. This service **only sends a context block** as the `task` form-data field — the gateway merges it with the stored template server-side.

**What gets sent (`task` field):**

```
{
Client: HDFC Bank
Process: Home Loan Origination
Entity scope: APPLICANT

AVAILABLE INTERNAL EXCEL_KEY VALUES:
APPLICANTFIRSTNAME | First name of applicant | loanAccount.customer.firstName
...

SEMANTIC SHORTCUTS:
dateofbirth → APPLICANTDATEOFBIRTH  (seen in 87 partners)
...

FIELDS TO MAP:
field_name | column_category
...
}
```

**What comes back (pipe-delimited):**

```json
{
  "Date_of_Birth": "APPLICANTDATEOFBIRTH|loanAccount.customer.dateOfBirth|0.97|SEMANTIC|matched via shortcut",
  "Aadhar_Front":  "DOCUMENTNAME||0.90|PATTERN|aadhaar keyword detected"
}
```

Format: `matched_excel_key | json_key | confidence | match_pattern | reasoning`

---

## 🧪 Streamlit Dashboard (NEW)

A powerful UI to test and visualize the mapping pipeline.

**Run Dashboard:**
```bash
streamlit run streamlit_dashboard.py
```

---

## 📊 Confidence & Scoring

### Confidence Bands

| Band | Range | Meaning |
|------|-------|---------|
| 🟢 High | ≥ 0.90 | Safe to auto-accept |
| 🟡 Review | 0.80 – 0.89 | Borderline — human check recommended |
| 🔴 Flag | < 0.80 | `needs_review = true` |

### Alias Registry Tiers

| Tier | Partner Count | Signal |
|------|--------------|--------|
| tier1 | ≥ 30 | Very common alias — high confidence |
| tier2 | 10 – 29 | Common |
| tier3 | 3 – 9 | Occasional |
| tier4 | < 3 | Rare / newly seen |

---

## 📋 Output Schema

| Column | Description |
|--------|-------------|
| `partner_field` | Raw partner column name |
| `column_category` | Category from partner file |
| `entity` | `APPLICANT` / `LOAN` / `DOCUMENT` / `GUARANTOR` / … |
| `matched_excel_key` | Internal PUTM key |
| `json_key` | Dot-path LMS API key |
| `confidence` | `0.0 – 1.0` |
| `match_type` | `deterministic` / `llm_semantic` / `fuzzy` / `unmatched` / … |
| `reasoning` | Human-readable explanation |
| `needs_review` | Boolean flag |
| `winning_engine` | `deterministic` / `llm` / `fuzzy` / `embedding` / `none` |

---

## 🤝 Contributing

1. Fork the repo and cut a feature branch off `main`
2. Rebuild references locally before testing (`POST /references/build`)
3. Add or update tests under `tests/`
4. Open a PR with a clear description of what changed and why

For questions about the Dvara gateway contract or PUTM schema, contact the platform team.

---

<div align="center">

Made with ☕ by the Platform Engineering team

</div>