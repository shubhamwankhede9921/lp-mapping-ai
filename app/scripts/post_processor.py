"""
Post-processor for LP field mapping results.

Handles:
- Auto-numbering of special fields (DOCUMENTNAME, DOCUMENTID, FEE, LOANPARAMETER, CUSTOMERPARAM)
- JSON key resolution from field dictionary
- Output generation for Excel/API consumption
"""

from dataclasses import dataclass
from typing import Optional, Union


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


class PostProcessor:
    """Post-processes mapping results for output generation."""

    # All spellings of special fields that get auto-numbered
    SPECIAL_FIELDS = {
        "DOCUMENTNAME",
        "DOCUMENTID",
        "FEE",
        "LOANPARAMETER",
        "CUSTOMERPARAM",
        "CUSTOMERPARAMETER",   # LLM returns full spelling — treated as CUSTOMERPARAM
    }

    # Fields always renumbered sequentially from 1 (ignore incoming numbers)
    SEQUENTIAL_RENUMBER_FIELDS = {
        "LOANPARAMETER",
        "CUSTOMERPARAM",
        "CUSTOMERPARAMETER",
    }

    # Normalize variant spellings to canonical base key used in output
    CANONICAL_BASE = {
        "CUSTOMERPARAMETER": "CUSTOMERPARAM",
    }

    def __init__(self, field_dictionary: dict):
        self.field_dict = field_dictionary
        self.doc_counter = 0
        self.fee_counter = 0
        self.loan_param_counter = 0
        self.customer_param_counter = 0

    def process_results(self, results: list[Union[MatchResult, dict]]) -> list[dict]:
        """
        Process all results: auto-number special fields, resolve json_keys.

        - LOANPARAMETER and CUSTOMERPARAM(ETER) are always renumbered 1, 2, 3...
        - DOCUMENTNAME, DOCUMENTID, FEE preserve pre-existing numbers if present.
        """
        # Normalize dicts to MatchResult
        normalized = []
        for item in results:
            if isinstance(item, dict):
                normalized.append(MatchResult(
                    partner_field=item.get("partner_field", ""),
                    column_category=item.get("column_category"),
                    entity=item.get("entity", "OTHER"),
                    matched_excel_key=item.get("matched_excel_key", ""),
                    json_key=item.get("json_key", ""),
                    confidence=item.get("confidence", 0.0),
                    match_type=item.get("match_type", "unmatched"),
                    reasoning=item.get("reasoning", ""),
                    needs_review=item.get("needs_review", False),
                ))
            else:
                normalized.append(item)

        # First pass: set counters from pre-existing numbers (skip sequential fields)
        self._extract_max_counters_except_sequential(normalized)

        # Second pass: number and resolve
        processed = []
        for result in normalized:
            base_key = self._get_base_key(result.matched_excel_key)

            if base_key in self.SPECIAL_FIELDS:
                if base_key in self.SEQUENTIAL_RENUMBER_FIELDS:
                    # Always assign fresh sequential number
                    numbered_key = self.assign_number(base_key)
                elif result.matched_excel_key != base_key:
                    # Pre-numbered key — preserve it
                    numbered_key = result.matched_excel_key
                else:
                    # Base key only — assign next number
                    numbered_key = self.assign_number(base_key)

                resolved_json_key = self.resolve_json_key(numbered_key)
                processed.append({
                    "partner_field":     result.partner_field,
                    "column_category":   result.column_category,
                    "entity":            result.entity,
                    "matched_excel_key": numbered_key,
                    "json_key":          resolved_json_key,
                    "confidence":        result.confidence,
                    "match_type":        result.match_type,
                    "reasoning":         result.reasoning,
                    "needs_review":      result.needs_review or result.confidence < 0.80,
                })

            else:
                # Non-special field — keep as is, just resolve json_key
                resolved_json_key = self.resolve_json_key(result.matched_excel_key)
                processed.append({
                    "partner_field":     result.partner_field,
                    "column_category":   result.column_category,
                    "entity":            result.entity,
                    "matched_excel_key": result.matched_excel_key,
                    "json_key":          resolved_json_key,
                    "confidence":        result.confidence,
                    "match_type":        result.match_type,
                    "reasoning":         result.reasoning,
                    "needs_review":      result.needs_review or result.confidence < 0.80,
                })

        return processed

    def resolve_json_key(self, excel_key: str) -> str:
        """Look up json_key from field_dictionary; return excel_key if not found."""
        return self.field_dict.get(excel_key, excel_key)

    def assign_number(self, base_key: str) -> str:
        """
        Assign the next sequential number for a base key.
        Variant spellings (CUSTOMERPARAMETER) are canonicalized before numbering
        so the output always uses the short canonical form (CUSTOMERPARAM1, CUSTOMERPARAM2...).
        """
        canonical = self.CANONICAL_BASE.get(base_key, base_key)

        if canonical == "DOCUMENTNAME":
            self.doc_counter += 1
            return f"DOCUMENTNAME{self.doc_counter}"
        elif canonical == "DOCUMENTID":
            self.doc_counter += 1
            return f"DOCUMENTID{self.doc_counter}"
        elif canonical == "FEE":
            self.fee_counter += 1
            return f"FEE{self.fee_counter}"
        elif canonical == "LOANPARAMETER":
            self.loan_param_counter += 1
            return f"LOANPARAMETER{self.loan_param_counter}"
        elif canonical == "CUSTOMERPARAM":
            self.customer_param_counter += 1
            return f"CUSTOMERPARAM{self.customer_param_counter}"
        else:
            return base_key

    def _get_base_key(self, matched_key: str) -> str:
        """Strip trailing digits: 'DOCUMENTNAME1' → 'DOCUMENTNAME', 'FEE2' → 'FEE'."""
        i = len(matched_key) - 1
        while i >= 0 and matched_key[i].isdigit():
            i -= 1
        return matched_key[:i + 1]

    def _extract_number(self, numbered_key: str) -> int:
        """Extract trailing number: 'DOCUMENTNAME1' → 1, 'FEE42' → 42."""
        i = len(numbered_key) - 1
        while i >= 0 and numbered_key[i].isdigit():
            i -= 1
        return int(numbered_key[i + 1:])

    def _extract_max_counters_except_sequential(self, results: list[MatchResult]) -> None:
        """
        Pre-scan pre-numbered keys and set counters to their max,
        so newly assigned numbers don't collide.
        Skips LOANPARAMETER and CUSTOMERPARAM(ETER) since those always restart from 1.
        """
        for result in results:
            base_key = self._get_base_key(result.matched_excel_key)
            if (
                base_key in self.SPECIAL_FIELDS
                and base_key not in self.SEQUENTIAL_RENUMBER_FIELDS
                and result.matched_excel_key != base_key
            ):
                number = self._extract_number(result.matched_excel_key)
                if base_key == "DOCUMENTNAME":
                    self.doc_counter = max(self.doc_counter, number)
                elif base_key == "DOCUMENTID":
                    self.doc_counter = max(self.doc_counter, number)
                elif base_key == "FEE":
                    self.fee_counter = max(self.fee_counter, number)


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
        "partner_field":     partner_field,
        "column_category":   column_category,
        "entity":            entity,
        "matched_excel_key": matched_excel_key,
        "json_key":          json_key,
        "confidence":        confidence,
        "match_type":        match_type,
        "reasoning":         reasoning,
        "needs_review":      needs_review or confidence < 0.80,
    }