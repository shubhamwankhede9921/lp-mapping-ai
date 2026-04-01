#!/usr/bin/env python3
"""
Finalize Pipeline - Combine deterministic + LLM results and generate output Excel.

This script takes the deterministic results and LLM mapping results,
combines them, runs post-processing (numbering, json_key resolution),
and generates the final output Excel.

Usage:
    python finalize_pipeline.py <pipeline_output_dir> <output_excel_path> [--client CLIENT]
"""

import json
import sys
import os
import argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from matching_engine import load_references
from generate_output import generate_mapping_excel


def load_deterministic_results(pipeline_dir):
    """Load deterministic matching results."""
    path = os.path.join(pipeline_dir, "deterministic_results.json")
    with open(path, "r") as f:
        return json.load(f)


def load_llm_results(pipeline_dir):
    """Load LLM matching results."""
    path = os.path.join(pipeline_dir, "llm_results.json")
    if not os.path.exists(path):
        print(f"WARNING: LLM results file not found: {path}")
        return []
    with open(path, "r") as f:
        data = json.load(f)
    # Handle both flat list and nested {mappings: [...]} format
    if isinstance(data, dict) and "mappings" in data:
        return data["mappings"]
    return data


def resolve_json_keys(mappings, refs):
    """Resolve json_keys from field_dictionary for any mapping missing one."""
    by_excel_key = refs["field_dictionary"].get("by_excel_key", {})
    for m in mappings:
        ek = m.get("matched_excel_key", "")
        if ek and (not m.get("json_key") or m["json_key"] in ("", "__FILL_THIS__", "null", "LOANPARAMETER")):
            entry = by_excel_key.get(ek, {})
            m["json_key"] = entry.get("json_key", "")
    return mappings


def auto_number_special_fields(mappings):
    """
    Auto-number DOCUMENTNAME, DOCUMENTID, FEE, LOANPARAMETER fields.
    These come from the LLM without number suffixes.
    """
    doc_name_counter = 1
    doc_id_counter = 1
    fee_counter = 1
    loan_param_counter = 1

    for m in mappings:
        ek = m.get("matched_excel_key", "")
        if not ek:
            continue

        if ek == "DOCUMENTNAME":
            m["matched_excel_key"] = f"DOCUMENTNAME{doc_name_counter}"
            doc_name_counter += 1
        elif ek == "DOCUMENTID":
            m["matched_excel_key"] = f"DOCUMENTID{doc_id_counter}"
            doc_id_counter += 1
        elif ek == "FEE":
            m["matched_excel_key"] = f"FEE{fee_counter}"
            fee_counter += 1
        elif ek == "LOANPARAMETER":
            m["matched_excel_key"] = f"LOANPARAMETER{loan_param_counter}"
            loan_param_counter += 1
        elif ek == "CUSTOMERPARAMETER":
            m["matched_excel_key"] = f"CUSTOMERPARAMETER{loan_param_counter}"
            loan_param_counter += 1

    return mappings


def combine_and_process(det_results, llm_results, refs):
    """Combine deterministic + LLM results, process, and return final mappings."""
    # Normalize LLM results to same format as deterministic
    for m in llm_results:
        if "needs_review" not in m:
            m["needs_review"] = m.get("confidence", 0) < 0.80
        if "match_type" not in m:
            m["match_type"] = m.get("matched_pattern", "llm_semantic")
        # Ensure match_type reflects LLM source
        mt = m.get("match_type", "").lower()
        if mt in ("exact_match", "exact"):
            m["match_type"] = "llm_exact"
        elif mt in ("semantic", "semantic_match"):
            m["match_type"] = "llm_semantic"
        elif mt in ("pattern", "pattern_match"):
            m["match_type"] = "llm_pattern"
        elif mt in ("none", ""):
            m["match_type"] = "llm_fallback"

    # Auto-number special fields from LLM results
    llm_results = auto_number_special_fields(llm_results)

    # Combine all results
    all_mappings = det_results + llm_results

    # Resolve json_keys
    all_mappings = resolve_json_keys(all_mappings, refs)

    return all_mappings


def main():
    parser = argparse.ArgumentParser(description="Finalize LP Field Mapping Pipeline")
    parser.add_argument("pipeline_dir", help="Pipeline output directory containing intermediate files")
    parser.add_argument("output_path", help="Path for the output Excel file")
    parser.add_argument("--client", default="", help="Client/partner name")
    parser.add_argument("--process", default="COMBINED", help="Process type")
    parser.add_argument("--refs", default=None, help="References directory path")

    args = parser.parse_args()

    # Find references
    script_dir = Path(__file__).parent.parent
    refs_dir = args.refs or str(script_dir / "references")
    refs = load_references(refs_dir)

    # Load results
    print(f"Loading deterministic results from: {args.pipeline_dir}")
    det_results = load_deterministic_results(args.pipeline_dir)
    print(f"  Deterministic: {len(det_results)} mappings")

    print(f"Loading LLM results from: {args.pipeline_dir}")
    llm_results = load_llm_results(args.pipeline_dir)
    print(f"  LLM: {len(llm_results)} mappings")

    # Combine and process
    all_mappings = combine_and_process(det_results, llm_results, refs)
    print(f"  Total: {len(all_mappings)} mappings")

    # Generate output
    generate_mapping_excel(all_mappings, args.output_path, args.client, args.process)
    print(f"\nOutput Excel: {args.output_path}")

    # Summary stats
    review_count = sum(1 for m in all_mappings if m.get("needs_review", False))
    high_conf = sum(1 for m in all_mappings if m.get("confidence", 0) >= 0.90)
    print(f"  High confidence: {high_conf}")
    print(f"  Needs review: {review_count}")


if __name__ == "__main__":
    main()
