"""LLM validation using Dvara LLM API (configurable)."""
import json
import httpx
from app.config import get_settings


def _call_dvara_llm(model_name: str, prompt: str) -> str:
    """
    Call Dvara LLM API. Adapt URL/body to your actual API.
    Expects response to contain generated text (e.g. JSON).
    """
    settings = get_settings()
    url = f"{settings.llm_base_url.rstrip('/')}/v1/chat/completions"
    headers = {"Content-Type": "application/json"}
    if settings.llm_api_key:
        headers["Authorization"] = f"Bearer {settings.llm_api_key}"
    body = {
        "model": model_name or settings.llm_model_name,
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": 2048,
    }
    with httpx.Client(timeout=30.0) as client:
        resp = client.post(url, json=body, headers=headers)
        resp.raise_for_status()
        data = resp.json()
    # Common shape: choices[0].message.content
    choices = data.get("choices") or []
    if not choices:
        return ""
    msg = choices[0].get("message") or {}
    return (msg.get("content") or "").strip()


def validate_mapping_with_llm(
    model_name: str,
    lms_columns: list[str],
    client_columns: list[str],
    mapping_suggestions: list[dict],
) -> list[dict]:
    """
    Ask LLM to validate/improve mapping and return final JSON list of
    { "client_column", "lms_column", "confidence" }.
    If API fails or response is not valid JSON, returns mapping_suggestions unchanged.
    """
    settings = get_settings()
    model = model_name or settings.llm_model_name
    prompt = f"""You are a column mapping AI for a Loan Origination System.

LMS Columns (target):
{json.dumps(lms_columns, indent=2)}

Client Columns (source):
{json.dumps(client_columns, indent=2)}

Suggested Mapping (improve if needed, keep same structure):
{json.dumps(mapping_suggestions, indent=2)}

Return ONLY a valid JSON array of objects with keys: client_column, lms_column, confidence (0-1).
No markdown, no explanation. Example: [{{"client_column":"full_name","lms_column":"name","confidence":0.95}}]
"""
    try:
        raw = _call_dvara_llm(model, prompt)
        # Strip markdown code block if present
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
            raw = raw.strip()
        arr = json.loads(raw)
        if isinstance(arr, list):
            return arr
    except (httpx.HTTPError, json.JSONDecodeError, KeyError):
        pass
    return mapping_suggestions
