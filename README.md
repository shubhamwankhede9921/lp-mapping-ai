# LP Mapping Service

Intelligent Schema Matching Engine — AI-assisted column mapping from client uploads to LMS master schema.

## Features

- **Hybrid matching**: Rule-based → Fuzzy → Embedding similarity → Historical patterns → Optional LLM validation
- **Confidence tiers**: `auto_map` (≥0.85), `suggest` (0.70–0.85), `manual_required` (<0.70)
- **Historical learning**: Stores and reuses past client→LMS mappings for better suggestions
- **Configurable**: Dvara LLM API URL/model and thresholds via env

## Setup

```bash
cd "e:\Dvara solutions\lp mapping service"
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Copy `.env.example` to `.env` and set `LLM_BASE_URL` / `LLM_MODEL_NAME` if you use LLM validation.

## Run

```bash
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

- API docs: http://localhost:8000/docs  
- Health: http://localhost:8000/health  

## API

| Method | Path | Description |
|--------|------|-------------|
| POST | `/api/mapping/auto-map` | Suggest mappings (rule + fuzzy + embedding + historical) |
| POST | `/api/mapping/auto-map-with-llm` | Same + final validation via Dvara LLM |
| POST | `/api/mapping/suggestions` | Mappings with `tier` for UI (auto_map / suggest / manual_required) |

**Request body (all endpoints):**

```json
{
  "client_name": "ABC Finance",
  "client_columns": ["full_name", "dob", "pan_no"],
  "lms_columns": ["name", "date_of_birth", "pan_number"]
}
```

**Response:**

```json
{
  "mappings": [
    { "client_column": "full_name", "lms_column": "name", "confidence": 0.94 },
    { "client_column": "dob", "lms_column": "date_of_birth", "confidence": 0.91 }
  ],
  "client_name": "ABC Finance"
}
```

## Project structure

```
app/
├── main.py
├── config.py
├── api/
│   └── mapping_controller.py
├── services/
│   ├── mapping_service.py
│   ├── rule_engine.py
│   ├── fuzzy_engine.py
│   ├── embedding_engine.py
│   └── llm_service.py
├── repository/
│   ├── database.py
│   └── mapping_repository.py
├── models/
│   ├── request_model.py
│   └── response_model.py
└── utils/
    └── text_utils.py
```

## Storing mappings for learning

Use `app.repository.mapping_repository.save_mapping` or `save_mappings_bulk` to persist confirmed mappings so future requests get better historical suggestions. The table `existing_column_mapping` is created automatically on first run.
