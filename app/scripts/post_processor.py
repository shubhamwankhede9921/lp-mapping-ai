"""
Post-processor for LP field mapping results.

Handles:
- Auto-numbering of special fields
- JSON key resolution from field dictionary
- Duplicate excel key detection and resolution
- Output generation for Excel/API consumption

Changelog
---------
Fix 1 – collision losers stay sequential LOANPARAMETER
    Losers get base LOANPARAMETER with match_type duplicate_excel_key_displaced;
    `_reroute_parameter_base_key` no longer promotes those rows to applicant /
    co-applicant parameter buckets.  `moved_by_collision_resolver` is still set.

Fix 2 – same-json_key collision bypass
    Two partner fields that collide on the same excel key but resolve to the
    IDENTICAL json_key are semantically the same mapping coming from two
    different input sheets / naming conventions (e.g. 'customerId' from an
    API sheet vs 'Customer Id' from Sheet2).  Displacing the loser to
    LOANPARAMETER in this case destroys a valid direct mapping and inflates
    the LOANPARAMETER bucket.  When ALL colliding entries share the same
    resolved json_key the entire collision group is left untouched.

Fix 3 – duplicate excel key: displace losers (non-special keys only)
    For non-special excel keys: if the same key appears more than once *for
    the same (column_category, entity)*, the highest-confidence entry keeps
    the key; lower-confidence rows are moved to a fresh generic parameter
    slot (base LOANPARAMETER, then numbered; moved_by_collision_resolver=True).
    Rows that share an excel key
    but differ in column_category or entity are not treated as duplicates —
    they all keep the key. Special fields — DOCUMENTNAME, DOCUMENTID, FEE,
    LOANPARAMETER, LOANAPPLICANTPARAM and CUSTOMERPARAM(ETER) — are never
    deduped here; they stay as-is and are renumbered sequentially by
    process_results().
"""

from collections import defaultdict
from dataclasses import dataclass
from typing import Optional, Union
import re


@dataclass
class MatchResult:
    """Represents a single field match result."""

    partner_field: str
    column_category: Optional[str]
    entity: str
    matched_excel_key: str
    json_key: str
    confidence: float
    match_type: str
    reasoning: str
    previous_mapping_reason: str = ""
    llm_change_reason: str = ""
    llm_param_bucket_reason: str = ""
    needs_review: bool = False
    # Set by resolve_duplicate_excel_keys() so downstream can tell the row
    # lost a direct excel_key to a higher-confidence sibling.
    moved_by_collision_resolver: bool = False


class PostProcessor:
    """Post-processes mapping results for output generation."""

    SPECIAL_FIELDS = {
        "DOCUMENTNAME",
        "DOCUMENTID",
        "FEE",
        "LOANPARAMETER",
        "CUSTOMERPARAM",
        "CUSTOMERPARAMETER",
        "LOANAPPLICANTPARAM",
    }

    SEQUENTIAL_RENUMBER_FIELDS = {
        "DOCUMENTNAME",
        "DOCUMENTID",
        "FEE",
        "LOANPARAMETER",
        "CUSTOMERPARAM",
        "CUSTOMERPARAMETER",
        "LOANAPPLICANTPARAM",
    }

    CANONICAL_BASE = {
        "CUSTOMERPARAMETER": "CUSTOMERPARAM",
        "LOANAPPLICANTPARAM": "LOANAPPLICANTPARAM",
    }

    def __init__(self, field_dictionary: dict):
        self.field_dict = field_dictionary
        self.by_excel_key = (
            field_dictionary.get("by_excel_key", {})
            if isinstance(field_dictionary, dict)
            else {}
        )
        self.document_name_counter = 0
        self.document_id_counter = 0
        self.fee_counter = 0
        self.loan_param_counter = 0
        self.customer_param_counter = 0
        # Per co-applicant slot: next index for LOANCOAPP{n}CUSTPARAM{k}
        self._loancoapp_custparam_by_slot: dict[int, int] = {}

    # ------------------------------------------------------------------
    # Public entry point
    # ------------------------------------------------------------------

    def process_results(self, results: list[Union[MatchResult, dict]]) -> list[dict]:
        """
        Process all results: auto-number special fields, resolve json_keys,
        and handle duplicate excel key mappings.

        - DOCUMENTNAME, DOCUMENTID, FEE, LOANPARAMETER, CUSTOMERPARAM(ETER),
          and LOANAPPLICANTPARAM are always renumbered 1, 2, 3 ...
        - If multiple partner fields map to the same non-special excel key,
          losers are displaced to generic parameter slots (see resolve_duplicate_excel_keys).
        """
        normalized: list[MatchResult] = []
        for item in results:
            if isinstance(item, dict):
                normalized.append(
                    MatchResult(
                        partner_field=item.get("partner_field", ""),
                        column_category=item.get("column_category"),
                        entity=item.get("entity", "OTHER"),
                        matched_excel_key=item.get("matched_excel_key", ""),
                        json_key=item.get("json_key", ""),
                        confidence=item.get("confidence", 0.0),
                        match_type=item.get("match_type", "unmatched"),
                        reasoning=item.get("reasoning", ""),
                        previous_mapping_reason=item.get("previous_mapping_reason", ""),
                        llm_change_reason=item.get("llm_change_reason", ""),
                        llm_param_bucket_reason=item.get("llm_param_bucket_reason", ""),
                        needs_review=item.get("needs_review", False),
                        moved_by_collision_resolver=item.get(
                            "moved_by_collision_resolver", False
                        ),
                    )
                )
            else:
                normalized.append(item)

        # Resolve duplicate excel key collisions before further processing.
        # Losers within the same (excel_key, column_category, entity) group are
        # displaced to LOANPARAMETER; same excel_key under different category
        # or entity is left untouched.
        normalized = self.resolve_duplicate_excel_keys(normalized)

        processed: list[dict] = []
        loanparameter_base_numbered = 0

        for result in normalized:
            original_base_key = self._get_base_key(result.matched_excel_key)
            routed_excel_key = self._reroute_parameter_base_key(result)
            base_key = self._get_base_key(routed_excel_key)
            coapp_custparam_base = self._is_loancoapp_custparam_unnumbered_base(base_key)

            if original_base_key == "LOANPARAMETER":
                loanparameter_base_numbered += 1

            if base_key in self.SPECIAL_FIELDS or coapp_custparam_base:
                if base_key in self.SEQUENTIAL_RENUMBER_FIELDS or coapp_custparam_base:
                    numbered_key = self.assign_number(base_key)
                elif routed_excel_key != base_key:
                    numbered_key = routed_excel_key
                else:
                    numbered_key = self.assign_number(base_key)

                resolved_json_key = self.resolve_json_key(numbered_key)
                processed.append(
                    {
                        "partner_field": result.partner_field,
                        "column_category": result.column_category,
                        "entity": result.entity,
                        "matched_excel_key": numbered_key,
                        "json_key": resolved_json_key,
                        "confidence": result.confidence,
                        "match_type": result.match_type,
                        "reasoning": result.reasoning,
                        "previous_mapping_reason": result.previous_mapping_reason,
                        "llm_change_reason": result.llm_change_reason,
                        "llm_param_bucket_reason": result.llm_param_bucket_reason,
                        "needs_review": result.needs_review or result.confidence < 0.80,
                    }
                )
            else:
                resolved_json_key = self.resolve_json_key(result.matched_excel_key)
                processed.append(
                    {
                        "partner_field": result.partner_field,
                        "column_category": result.column_category,
                        "entity": result.entity,
                        "matched_excel_key": result.matched_excel_key,
                        "json_key": resolved_json_key,
                        "confidence": result.confidence,
                        "match_type": result.match_type,
                        "reasoning": result.reasoning,
                        "previous_mapping_reason": result.previous_mapping_reason,
                        "llm_change_reason": result.llm_change_reason,
                        "llm_param_bucket_reason": result.llm_param_bucket_reason,
                        "needs_review": result.needs_review or result.confidence < 0.80,
                    }
                )

        if loanparameter_base_numbered:
            print(
                "PostProcessor: LOANPARAMETER*n rows (no entity promotion to applicant/co-app param): "
                f"{loanparameter_base_numbered}"
            )

        return processed

    # ------------------------------------------------------------------
    # Collision resolver
    # ------------------------------------------------------------------

    def resolve_duplicate_excel_keys(
        self, results: list[MatchResult]
    ) -> list[MatchResult]:
        """
        For non-special excel keys only: if the same excel key appears more
        than once *within the same (column_category, entity)*, keep the
        highest-confidence entry on that key and move every other row in that
        group to a generic parameter bucket (base LOANPARAMETER, then numbered).
        Two rows with
        the same excel key but different column_category or entity are not
        considered duplicates and are left unchanged.

        Special fields (LOANPARAMETER, DOCUMENTNAME, DOCUMENTID, FEE,
        CUSTOMERPARAM, CUSTOMERPARAMETER, LOANAPPLICANTPARAM) are NEVER
        deduped here — they are left exactly as received and renumbered
        sequentially later by process_results().

        Resolution rules
        ----------------
        1. Special/numbered base keys → skip entirely, keep all untouched.

        2. Non-special excel key with 2+ entries sharing the same normalized
           column_category and entity:
           a. Keep the highest-confidence entry (winner) unchanged.
           b. Each lower-confidence entry stays in the list but is displaced
              to base matched_excel_key=\"LOANPARAMETER\" (then numbered)
              with moved_by_collision_resolver=True (ties: first occurrence wins).
           c. Displaced rows are flagged needs_review=True.

        Customizing duplicate handling
        -------------------------------
        - To treat collisions under a *finer* grain than (category, entity),
          extend ``composite`` (e.g. include normalized ``partner_field`` family
          or sheet name) so fewer unrelated columns compete for one key.
        - To *avoid* displacing when every colliding row resolves to the same
          ``json_key`` (duplicate column labels for one payload slot), add an
          early ``continue`` for that group before displacing losers — see
          module docstring "Fix 2 – same-json_key collision bypass".
        - To disable displacement entirely, skip calling this method from
          ``process_results`` or gate it behind a settings flag.

        Parameters
        ----------
        results : list[MatchResult]

        Returns
        -------
        list[MatchResult]
            Same length as *results*. Special-field entries are unchanged.
        """
        # ----------------------------------------------------------------
        # Build index of non-special excel keys × (category, entity).
        # Same matched_excel_key under different column_category or entity
        # does not count as a duplicate group.
        # ----------------------------------------------------------------
        key_index: dict[tuple[str, str, str], list[tuple[int, MatchResult]]] = (
            defaultdict(list)
        )
        for pos, result in enumerate(results):
            base_key = self._get_base_key(result.matched_excel_key)
            if base_key in self.SPECIAL_FIELDS or self._is_loancoapp_custparam_unnumbered_base(
                base_key
            ):
                continue  # always preserve special / co-app param buckets as-is
            norm_cat = self._normalize_text(result.column_category)
            norm_ent = (result.entity or "").strip().upper()
            composite = (result.matched_excel_key, norm_cat, norm_ent)
            key_index[composite].append((pos, result))

        # Only process composite keys that have more than one entry
        candidates: dict[tuple[str, str, str], list[tuple[int, MatchResult]]] = {
            composite: entries
            for composite, entries in key_index.items()
            if len(entries) > 1
        }

        if not candidates:
            return results  # fast-path: nothing to do

        updated: list[MatchResult] = list(results)

        for (excel_key, norm_cat, norm_ent), entries in candidates.items():
            # Stable descending sort by confidence — first occurrence wins ties
            ranked = sorted(entries, key=lambda t: t[1].confidence, reverse=True)
            winner_pos, winner = ranked[0]
            losers = ranked[1:]

            loser_fields = [r.partner_field for _, r in losers]
            print(
                f"[DuplicateKeyResolver] Non-special excel key '{excel_key}' "
                f"(column_category~={norm_cat!r}, entity={norm_ent!r}) "
                f"has {len(entries)} duplicate entries. "
                f"Winner (kept): '{winner.partner_field}' "
                f"(confidence={winner.confidence:.2f}). "
                f"Displaced to generic parameter bucket: {loser_fields}"
            )

            for loser_pos, loser in losers:
                prev_reason = (loser.reasoning or "").strip()
                bump = (
                    f"Duplicate excel_key '{excel_key}': lower confidence than "
                    f"'{winner.partner_field}' ({winner.confidence:.2f}); "
                    "displaced from that key to the next sequential LOANPARAMETER*n slot."
                )
                reasoning = f"{prev_reason}; {bump}" if prev_reason else bump
                updated[loser_pos] = MatchResult(
                    partner_field=loser.partner_field,
                    column_category=loser.column_category,
                    entity=loser.entity,
                    matched_excel_key="LOANPARAMETER",
                    json_key="",
                    confidence=min(float(loser.confidence or 0.0), 0.79),
                    match_type="duplicate_excel_key_displaced",
                    reasoning=reasoning.strip(),
                    previous_mapping_reason=loser.previous_mapping_reason,
                    llm_change_reason=loser.llm_change_reason,
                    llm_param_bucket_reason=loser.llm_param_bucket_reason,
                    needs_review=True,
                    moved_by_collision_resolver=True,
                )

        return updated

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _resolve_json_key_for_excel_key(
        self, excel_key: str, fallback_json_key: str
    ) -> str:
        """
        Return the authoritative json_key for *excel_key* from the field
        dictionary, or *fallback_json_key* when the key is absent from the dict.
        """
        if self.by_excel_key:
            dict_json_key = self.by_excel_key.get(excel_key, {}).get("json_key")
            if dict_json_key:
                return dict_json_key
        elif excel_key in self.field_dict:
            return self.field_dict[excel_key]
        return fallback_json_key or excel_key

    def _normalize_text(self, value: Optional[str]) -> str:
        if not value:
            return ""
        return re.sub(r"[^a-z0-9]+", " ", value.strip().lower()).strip()

    def _reroute_parameter_base_key(self, result: MatchResult) -> str:
        """
        LOANPARAMETER*n stays loan-level only. No promotion to LOANAPPLICANTPARAM
        or LOANCOAPP*CUSTPARAM by entity (downstream consumers expect generic buckets).
        """
        return result.matched_excel_key

    def _is_loancoapp_custparam_unnumbered_base(self, base_key: str) -> bool:
        """True for LOANCOAPP{n}CUSTPARAM before trailing param index is assigned."""
        return bool(re.match(r"^LOANCOAPP\d+CUSTPARAM$", base_key or ""))

    def resolve_json_key(self, excel_key: str) -> str:
        """Look up json_key from field_dictionary; return excel_key if not found."""
        if self.by_excel_key:
            return self.by_excel_key.get(excel_key, {}).get("json_key", excel_key)
        return self.field_dict.get(excel_key, excel_key)

    def assign_number(self, base_key: str) -> str:
        """
        Assign the next sequential number for a base key.

        Variant spellings are canonicalized before numbering.
        """
        canonical = self.CANONICAL_BASE.get(base_key, base_key)

        if canonical == "DOCUMENTNAME":
            self.document_name_counter += 1
            return f"DOCUMENTNAME{self.document_name_counter}"
        if canonical == "DOCUMENTID":
            self.document_id_counter += 1
            return f"DOCUMENTID{self.document_id_counter}"
        if canonical == "FEE":
            self.fee_counter += 1
            return f"FEE{self.fee_counter}"
        if canonical == "LOANPARAMETER":
            self.loan_param_counter += 1
            return f"LOANPARAMETER{self.loan_param_counter}"
        if canonical == "CUSTOMERPARAM":
            self.customer_param_counter += 1
            return f"CUSTOMERPARAM{self.customer_param_counter}"
        if canonical == "LOANAPPLICANTPARAM":
            self.customer_param_counter += 1
            return f"LOANAPPLICANTPARAM{self.customer_param_counter}"
        m_co = re.match(r"^LOANCOAPP(\d+)CUSTPARAM$", canonical)
        if m_co:
            slot = int(m_co.group(1))
            self._loancoapp_custparam_by_slot[slot] = (
                self._loancoapp_custparam_by_slot.get(slot, 0) + 1
            )
            idx = self._loancoapp_custparam_by_slot[slot]
            return f"LOANCOAPP{slot}CUSTPARAM{idx}"
        return base_key

    def _get_base_key(self, matched_key: str) -> str:
        """Strip trailing digits from a numbered excel key."""
        i = len(matched_key) - 1
        while i >= 0 and matched_key[i].isdigit():
            i -= 1
        return matched_key[: i + 1]

    def _extract_number(self, numbered_key: str) -> int:
        """Extract trailing number from a numbered excel key."""
        i = len(numbered_key) - 1
        while i >= 0 and numbered_key[i].isdigit():
            i -= 1
        return int(numbered_key[i + 1 :])


# ------------------------------------------------------------------
# Convenience factory
# ------------------------------------------------------------------

def create_output_dict(
    partner_field: str,
    column_category: Optional[str],
    entity: str,
    matched_excel_key: str,
    json_key: str,
    confidence: float,
    match_type: str,
    reasoning: str,
    needs_review: bool = False,
) -> dict:
    return {
        "partner_field": partner_field,
        "column_category": column_category,
        "entity": entity,
        "matched_excel_key": matched_excel_key,
        "json_key": json_key,
        "confidence": confidence,
        "match_type": match_type,
        "reasoning": reasoning,
        "needs_review": needs_review or confidence < 0.80,
    }
