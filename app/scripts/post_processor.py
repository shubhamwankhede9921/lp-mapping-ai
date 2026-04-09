"""
Post-processor for LP field mapping results.

Handles:
- Auto-numbering of special fields
- JSON key resolution from field dictionary
- Duplicate excel key detection and resolution
- Output generation for Excel/API consumption

Changelog
---------
Fix 1 – moved_by_collision_resolver flag
    When the collision resolver displaces a loser field to LOANPARAMETER,
    _reroute_parameter_base_key must NOT re-promote it to LOANAPPLICANTPARAM
    even when the field's entity is APPLICANT.  The new
    `moved_by_collision_resolver` flag on MatchResult is the precise guard.

Fix 2 – same-json_key collision bypass
    Two partner fields that collide on the same excel key but resolve to the
    IDENTICAL json_key are semantically the same mapping coming from two
    different input sheets / naming conventions (e.g. 'customerId' from an
    API sheet vs 'Customer Id' from Sheet2).  Displacing the loser to
    LOANPARAMETER in this case destroys a valid direct mapping and inflates
    the LOANPARAMETER bucket.  When ALL colliding entries share the same
    resolved json_key the entire collision group is left untouched.

Fix 3 – duplicate excel key: drop losers entirely (non-special keys only)
    For non-special excel keys: if the same key appears more than once, only
    the highest-confidence entry (first occurrence on tie) is kept; all
    lower-confidence duplicates are REMOVED from the output entirely.
    Special fields — DOCUMENTNAME, DOCUMENTID, FEE, LOANPARAMETER,
    LOANAPPLICANTPARAM and CUSTOMERPARAM(ETER) — are NEVER touched by this
    logic; every occurrence is preserved as-is and renumbered sequentially
    by process_results().
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
    needs_review: bool = False
    # Set to True exclusively by resolve_duplicate_excel_keys() so that
    # _reroute_parameter_base_key() never re-promotes the field.
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

    CUSTOMER_PARAM_CATEGORY_HINTS = (
        "customer",
        "applicant",
        "coapplicant",
        "co applicant",
        "borrower",
        "personal",
        "employment",
        "income",
        "bank",
        "bureau",
        "kyc",
        "address",
        "reference",
        "demographic",
    )

    LOAN_LEVEL_CATEGORY_HINTS = (
        "loan",
        "disbursement",
        "repayment",
        "emi",
        "pricing",
        "sanction",
        "scheme",
        "product",
        "facility",
    )

    LOAN_LEVEL_FIELD_HINTS = {
        "irr", "iir", "foir", "ltv", "loantovalue", "marginmoney",
        "downpayment", "schemename", "schemeid", "schemecode",
        "subproduct", "reducedemi", "tenure", "interest", "apr",
        "loanamount", "sanctionamount", "approvedamount", "emiamount",
        "disbursementamount", "repaymentfrequency",
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

    # ------------------------------------------------------------------
    # Public entry point
    # ------------------------------------------------------------------

    def process_results(self, results: list[Union[MatchResult, dict]]) -> list[dict]:
        """
        Process all results: auto-number special fields, resolve json_keys,
        and handle duplicate excel key mappings.

        - DOCUMENTNAME, DOCUMENTID, FEE, LOANPARAMETER, CUSTOMERPARAM(ETER),
          and LOANAPPLICANTPARAM are always renumbered 1, 2, 3 ...
        - If multiple partner fields map to the same excel key, only the
          highest-confidence winner is kept; all losers are dropped entirely.
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
                        needs_review=item.get("needs_review", False),
                        moved_by_collision_resolver=item.get(
                            "moved_by_collision_resolver", False
                        ),
                    )
                )
            else:
                normalized.append(item)

        # Resolve duplicate excel key collisions before further processing.
        # Losers are DROPPED entirely — regardless of whether the collision
        # is between different partner_field names or the same one repeated.
        normalized = self.resolve_duplicate_excel_keys(normalized)

        processed: list[dict] = []
        converted_loan_params = 0
        kept_loan_params = 0

        for result in normalized:
            original_base_key = self._get_base_key(result.matched_excel_key)
            routed_excel_key = self._reroute_parameter_base_key(result)
            base_key = self._get_base_key(routed_excel_key)

            if original_base_key == "LOANPARAMETER":
                if base_key == "LOANAPPLICANTPARAM":
                    converted_loan_params += 1
                else:
                    kept_loan_params += 1

            if base_key in self.SPECIAL_FIELDS:
                if base_key in self.SEQUENTIAL_RENUMBER_FIELDS:
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
                        "needs_review": result.needs_review or result.confidence < 0.80,
                    }
                )

        if converted_loan_params or kept_loan_params:
            print(
                "PostProcessor parameter routing: "
                f"converted LOANPARAMETER -> LOANAPPLICANTPARAM = {converted_loan_params}, "
                f"kept as LOANPARAMETER = {kept_loan_params}"
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
        than once, keep the single highest-confidence entry and DROP all
        lower-confidence duplicates.

        Special fields (LOANPARAMETER, DOCUMENTNAME, DOCUMENTID, FEE,
        CUSTOMERPARAM, CUSTOMERPARAMETER, LOANAPPLICANTPARAM) are NEVER
        touched here — they are left exactly as received and renumbered
        sequentially later by process_results().

        Resolution rules
        ----------------
        1. Special/numbered base keys → skip entirely, keep all untouched.

        2. Non-special excel key with 2+ entries:
           a. Keep the highest-confidence entry (winner).
           b. Every lower-confidence entry is REMOVED from the output.
           c. Tie-breaking (equal confidence): first occurrence wins.

        Parameters
        ----------
        results : list[MatchResult]

        Returns
        -------
        list[MatchResult]
            May be shorter than *results*; loser entries are omitted entirely.
            Special-field entries are always preserved unchanged.
        """
        # ----------------------------------------------------------------
        # Build index of non-special excel keys only.
        # Special fields are never entered into the index and are therefore
        # guaranteed to pass through the final filter untouched.
        # ----------------------------------------------------------------
        key_index: dict[str, list[tuple[int, MatchResult]]] = defaultdict(list)
        for pos, result in enumerate(results):
            base_key = self._get_base_key(result.matched_excel_key)
            if base_key in self.SPECIAL_FIELDS:
                continue  # always preserve special fields as-is
            key_index[result.matched_excel_key].append((pos, result))

        # Only process keys that have more than one entry
        candidates: dict[str, list[tuple[int, MatchResult]]] = {
            excel_key: entries
            for excel_key, entries in key_index.items()
            if len(entries) > 1
        }

        if not candidates:
            return results  # fast-path: nothing to do

        positions_to_drop: set[int] = set()

        for excel_key, entries in candidates.items():
            # Stable descending sort by confidence — first occurrence wins ties
            ranked = sorted(entries, key=lambda t: t[1].confidence, reverse=True)
            winner_pos, winner = ranked[0]
            losers = ranked[1:]

            loser_fields = [r.partner_field for _, r in losers]
            print(
                f"[DuplicateKeyResolver] Non-special excel key '{excel_key}' "
                f"has {len(entries)} duplicate entries. "
                f"Winner (kept): '{winner.partner_field}' "
                f"(confidence={winner.confidence:.2f}). "
                f"Dropped: {loser_fields}"
            )

            for loser_pos, _ in losers:
                positions_to_drop.add(loser_pos)

        # Special-field positions are never in positions_to_drop, so they
        # always survive this filter unchanged.
        return [r for pos, r in enumerate(results) if pos not in positions_to_drop]

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

    def _is_customer_related_parameter(self, result: MatchResult) -> bool:
        entity = (result.entity or "").upper()
        category = self._normalize_text(result.column_category)
        field = self._normalize_text(result.partner_field)
        compact_field = field.replace(" ", "")

        if entity in {
            "APPLICANT", "CUSTOMER",
            "COAPPLICANT", "COAPPLICANT1", "COAPPLICANT2",
            "COAPPLICANT3", "COAPPLICANT4",
        }:
            if compact_field in self.LOAN_LEVEL_FIELD_HINTS:
                return False
            return True

        if any(hint in category for hint in self.CUSTOMER_PARAM_CATEGORY_HINTS):
            if compact_field in self.LOAN_LEVEL_FIELD_HINTS:
                return False
            return True

        if (
            entity == "LOAN"
            and any(hint in category for hint in self.LOAN_LEVEL_CATEGORY_HINTS)
            and any(hint in field for hint in self.CUSTOMER_PARAM_CATEGORY_HINTS)
        ):
            return True

        return False

    def _reroute_parameter_base_key(self, result: MatchResult) -> str:
        base_key = self._get_base_key(result.matched_excel_key)
        if base_key != "LOANPARAMETER":
            return result.matched_excel_key

        # Fix 1: never reroute a field placed here by the collision resolver.
        if result.moved_by_collision_resolver:
            return result.matched_excel_key  # stay as LOANPARAMETER

        if result.match_type == "llm_parameter_bucket":
            return result.matched_excel_key
        if self._is_customer_related_parameter(result):
            return "LOANAPPLICANTPARAM"
        return result.matched_excel_key

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