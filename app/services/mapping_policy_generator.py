"""
Derive mapping_policy.json (structured fee PUTM leaves) from field_dictionary.

Used when rebuilding references from PUTM dumps so new fee-like catalogue keys
are picked up without editing Python. Merges with any existing mapping_policy.json
on disk (union of structured_fee_putm_base_keys).
"""

from __future__ import annotations

import json
import logging
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

logger = logging.getLogger(__name__)

# Generic numbered FEE slots (excel FEE1…FEE20 → loanAccount.feeN) stay out of the
# structured list — they are buckets, not semantic PUTM leaves.
_GENERIC_FEE_EXCEL = re.compile(r"^FEE\d+$", re.I)
_LOANPARAMETER = re.compile(r"^LOANPARAMETER", re.I)
_DOCUMENT = re.compile(r"^DOCUMENT(NAME|ID)", re.I)

# loanAccount.fee12 style paths map to generic FEE* excel keys.
_LOANACCOUNT_FEE_N = re.compile(r"loanaccount\.fee\d+\b", re.I)

# json_key substrings / tokens that indicate a fee/charge/tax leaf (not insurance).
_JSON_FEE_SIGNAL = re.compile(
    r"(processingfee|processingfeeservicetax|gstfee|otherfee|servicefee|"
    r"documentationfee|legalfee|technicalfee|valuationfee|modtfee|"
    r"clearingcharges|clearedcharges|followupcharges|participationfee|"
    r"partnerprocessingfees)",
    re.I,
)


def _strip_trailing_digits_excel_key(excel_key: str) -> str:
    value = (excel_key or "").strip()
    i = len(value) - 1
    while i >= 0 and value[i].isdigit():
        i -= 1
    return value[: i + 1]


def discover_structured_fee_putm_base_keys(field_dictionary: Dict[str, Any]) -> List[str]:
    """
    Scan by_excel_key for catalogue rows whose json_path or excel_key name denotes
    fee/charge semantics, excluding insurance and generic FEE* / LOANPARAMETER buckets.
    """
    by_ek = field_dictionary.get("by_excel_key") or {}
    if not isinstance(by_ek, dict):
        return []

    found: Set[str] = set()

    for excel_key, meta in by_ek.items():
        if not excel_key or not isinstance(meta, dict):
            continue
        ek = str(excel_key).strip()
        ek_u = ek.upper()
        jk = (meta.get("json_key") or "").strip()
        jk_l = jk.lower()

        if _LOANPARAMETER.match(ek_u) or _DOCUMENT.match(ek_u):
            continue
        if _GENERIC_FEE_EXCEL.match(ek_u):
            continue
        if "insurance" in jk_l:
            continue
        if _LOANACCOUNT_FEE_N.search(jk_l.replace(" ", "")):
            continue

        hit = False
        if jk and _JSON_FEE_SIGNAL.search(jk_l.replace(" ", "")):
            hit = True
        elif "FEE" in ek_u and not _GENERIC_FEE_EXCEL.match(ek_u):
            if "INSURANCE" in ek_u or "FEEDBACK" in ek_u:
                continue
            hit = True

        if hit:
            found.add(_strip_trailing_digits_excel_key(ek).upper())

    return sorted(found)


def _fee_putm_leaves_for_bases(
    field_dictionary: Dict[str, Any],
    bases: List[str],
) -> List[Dict[str, str]]:
    """One representative excel_key + json_key per base (for docs / optional prompt use)."""
    by_ek = field_dictionary.get("by_excel_key") or {}
    if not isinstance(by_ek, dict):
        return []
    want: Set[str] = {b.strip().upper() for b in bases if b.strip()}
    seen: Set[str] = set()
    rows: List[Dict[str, str]] = []

    for ek_raw, meta in by_ek.items():
        if not ek_raw or not isinstance(meta, dict):
            continue
        ek = str(ek_raw).strip()
        base = _strip_trailing_digits_excel_key(ek).upper()
        if base not in want or base in seen:
            continue
        jk = (meta.get("json_key") or "").strip()
        rows.append({"excel_key": ek, "base_key": base, "json_key": jk})
        seen.add(base)

    rows.sort(key=lambda r: r["base_key"])
    return rows


def build_mapping_policy(
    field_dictionary: Dict[str, Any],
    *,
    existing_policy: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """
    Build a full mapping_policy dict: auto-discovered keys merged with any existing
    structured_fee_putm_base_keys, plus fee_putm_leaves for excel+json visibility.
    """
    discovered = discover_structured_fee_putm_base_keys(field_dictionary)
    prior: Dict[str, Any] = existing_policy if isinstance(existing_policy, dict) else {}
    prior_keys_raw = prior.get("structured_fee_putm_base_keys")
    prior_keys: Set[str] = set()
    if isinstance(prior_keys_raw, list):
        prior_keys = {str(x).strip().upper() for x in prior_keys_raw if str(x).strip()}

    merged = sorted(set(discovered) | prior_keys)

    note = prior.get("llm_fee_mapping_note") or prior.get("fee_field_instructions")
    if not isinstance(note, str) or not note.strip():
        note = (
            "Prefer a concrete catalogue excel_key from structured_fee_putm_base_keys "
            "(see fee_putm_leaves for json paths) for fee/charge-style columns with "
            "entity FEE before using generic excel_key FEE. Add new PUTM leaves to "
            "the catalogue first; this list is regenerated when references are built."
        )

    out: Dict[str, Any] = {
        "structured_fee_putm_base_keys": merged,
        "fee_putm_leaves": _fee_putm_leaves_for_bases(field_dictionary, merged),
        "llm_fee_mapping_note": note.strip(),
        "_auto_discovered_base_keys": discovered,
        "_policy_generated_at": datetime.now(timezone.utc).isoformat(),
    }
    return out


def write_mapping_policy(
    references_dir: str,
    field_dictionary: Dict[str, Any],
) -> Tuple[Path, int, int]:
    """
    Merge auto-discovery with existing mapping_policy.json (if any), write file,
    return (path, merged_key_count, newly_discovered_count).
    """
    ref_dir = Path(references_dir)
    ref_dir.mkdir(parents=True, exist_ok=True)
    policy_path = ref_dir / "mapping_policy.json"

    existing: Dict[str, Any] = {}
    if policy_path.exists():
        try:
            raw = json.loads(policy_path.read_text(encoding="utf-8"))
            if isinstance(raw, dict):
                existing = raw
        except json.JSONDecodeError as e:
            logger.warning("Could not parse existing mapping_policy.json (%s) — overwriting", e)

    policy = build_mapping_policy(field_dictionary, existing_policy=existing)
    discovered = set(policy.get("_auto_discovered_base_keys") or [])
    prior_keys = set()
    pk = existing.get("structured_fee_putm_base_keys")
    if isinstance(pk, list):
        prior_keys = {str(x).strip().upper() for x in pk if str(x).strip()}
    new_only = len(discovered - prior_keys)

    # Strip internal keys from the persisted file (optional: keep for debugging)
    persist = {k: v for k, v in policy.items() if not k.startswith("_")}
    persist["_auto_discovered_base_keys"] = policy.get("_auto_discovered_base_keys", [])
    persist["_policy_generated_at"] = policy.get("_policy_generated_at", "")

    policy_path.write_text(
        json.dumps(persist, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    n_keys = len(persist.get("structured_fee_putm_base_keys") or [])
    logger.info(
        "Wrote %s — %d structured fee base key(s) (%d newly discovered vs prior file)",
        policy_path,
        n_keys,
        new_only,
    )
    return policy_path, n_keys, new_only


def main() -> None:
    """Regenerate mapping_policy.json from field_dictionary.json in a references dir."""
    import argparse

    parser = argparse.ArgumentParser(
        description="Build mapping_policy.json (fee PUTM leaves) from field_dictionary.json",
    )
    parser.add_argument(
        "--references-dir",
        default="app/references",
        help="Directory containing field_dictionary.json (default: app/references)",
    )
    args = parser.parse_args()
    ref_dir = Path(args.references_dir)
    fd_path = ref_dir / "field_dictionary.json"
    if not fd_path.exists():
        raise SystemExit(f"Missing {fd_path}")
    field_dictionary = json.loads(fd_path.read_text(encoding="utf-8"))
    path, n, newn = write_mapping_policy(str(ref_dir), field_dictionary)
    print(f"Wrote {path} — {n} key(s), {newn} newly discovered vs prior policy file.")


if __name__ == "__main__":
    main()
