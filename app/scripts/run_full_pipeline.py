#!/usr/bin/env python3
"""
Full Pipeline Runner for LP Field Mapping

This script runs the complete deterministic matching pass and prepares
everything needed for the LLM semantic matching step. It outputs:
1. A summary of deterministic results
2. The unmatched fields grouped by entity
3. Entity-scoped field lists for LLM matching
4. A template JSON file for the LLM to fill in

After the LLM fills in the template, run finalize_pipeline.py to
combine results and generate the output Excel.

Usage:
    python run_full_pipeline.py <partner_file> <client_name> [--process COMBINED]
"""

import json
import sys
import os
import argparse
from pathlib import Path

# Ensure scripts dir is on path
sys.path.insert(0, str(Path(__file__).parent))

from input_parser import parse_input
from matching_engine import load_references, match_batch, normalize_field
from llm_mapper import generate_mapping_prompt, build_prompt


def run_deterministic_pass(input_file, refs, process_name="COMBINED", sheet_filter=None):
    """Run the deterministic matching pass."""
    fields = parse_input(input_file)

    if sheet_filter:
        fields = [f for f in fields if f['source_sheet'] == sheet_filter]

    results = match_batch(fields, refs, process_name=process_name)
    return fields, results['matched'], results['unmatched']


def prepare_llm_context(unmatched_results, refs, client_name="", process_name=""):
    """
    Prepare the context needed for LLM matching.
    Returns structured data that can be embedded in SKILL.md instructions.
    """
    # Convert MatchResult objects to dicts for the LLM mapper
    unmatched_dicts = [
        {
            "field_name": u.partner_field,
            "column_category": u.column_category,
            "entity": u.entity
        }
        for u in unmatched_results
    ]

    # Group by entity
    entity_groups = {}
    for ud in unmatched_dicts:
        entity = ud["entity"]
        if entity not in entity_groups:
            entity_groups[entity] = []
        entity_groups[entity].append(ud)

    # Build entity-scoped field lists from the field dictionary
    field_dict = refs["field_dictionary"]
    alias_reg = refs["alias_registry"]

    # Map entity to role in field dictionary
    ENTITY_TO_ROLE = {
        "APPLICANT": "CUSTOMER",
        "COAPPLICANT": "COAPPLICANT",
        "COAPPLICANT1": "COAPPLICANT",
        "COAPPLICANT2": "COAPPLICANT",
        "COAPPLICANT3": "COAPPLICANT",
        "COAPPLICANT4": "COAPPLICANT",
        "LOAN": "LOAN",
        "GUARANTOR": "GUARANTOR",
        "DOCUMENT": "LOAN",
        "FEE": "LOAN",
        "OTHER": "CUSTOMER",
    }

    # For each entity group, prepare the relevant field list
    entity_contexts = {}
    for entity, fields in entity_groups.items():
        role = ENTITY_TO_ROLE.get(entity, "CUSTOMER")

        # Get fields for this role
        role_fields = field_dict.get("by_role", {}).get(role, [])

        # For LOAN role, filter out numbered special fields to keep it focused
        if role == "LOAN":
            filtered = []
            for f in role_fields:
                ek = f.get("excel_key", "")
                # Skip LOANPARAMETER50+, DOCUMENTNAME/ID beyond 5, FEE beyond 5
                import re
                lp_match = re.match(r'LOANPARAMETER(\d+)', ek)
                if lp_match and int(lp_match.group(1)) > 50:
                    continue
                dn_match = re.match(r'DOCUMENTNAME(\d+)', ek)
                if dn_match and int(dn_match.group(1)) > 5:
                    continue
                di_match = re.match(r'DOCUMENTID(\d+)', ek)
                if di_match and int(di_match.group(1)) > 5:
                    continue
                fee_match = re.match(r'FEE(\d+)', ek)
                if fee_match and int(fee_match.group(1)) > 5:
                    continue
                filtered.append(f)
            role_fields = filtered

        # Get top semantic shortcuts for this entity from alias registry
        forward = alias_reg.get("forward", {})
        shortcuts = []
        for norm_key, alias_data in forward.items():
            if alias_data.get("frequency", 0) >= 10:
                target = alias_data.get("target_excel_key", "")
                # Check if target belongs to this role
                target_info = field_dict.get("by_excel_key", {}).get(target, {})
                target_role = target_info.get("role", target_info.get("json_key_role", ""))
                if target_role == role or (role == "CUSTOMER" and target_role == ""):
                    variants = alias_data.get("variants", [])[:3]
                    shortcuts.append({
                        "from": norm_key,
                        "to": target,
                        "variants": variants,
                        "frequency": alias_data.get("frequency", 0)
                    })

        # Sort by frequency, take top 30
        shortcuts.sort(key=lambda x: x["frequency"], reverse=True)
        shortcuts = shortcuts[:30]

        entity_contexts[entity] = {
            "fields_to_map": [
                {"field_name": f["field_name"], "category": f.get("column_category")}
                for f in fields
            ],
            "available_excel_keys": [
                {"excel_key": f["excel_key"], "json_key": f.get("json_key", ""), "description": f.get("description", "")}
                for f in role_fields
            ],
            "semantic_shortcuts": shortcuts,
            "role": role,
            "field_count": len(role_fields)
        }

    return entity_contexts


def write_intermediate_files(
    matched_results, unmatched_results, entity_contexts,
    output_dir, client_name="", process_name=""
):
    """Write intermediate files for the LLM step."""
    os.makedirs(output_dir, exist_ok=True)

    # 1. Save deterministic results
    det_results = []
    for m in matched_results:
        det_results.append({
            "partner_field": m.partner_field,
            "column_category": m.column_category,
            "entity": m.entity,
            "matched_excel_key": m.matched_excel_key,
            "json_key": m.matched_json_key or "",
            "confidence": m.confidence,
            "match_type": m.match_type,
            "reasoning": m.reasoning,
            "needs_review": m.confidence < 0.80
        })

    with open(os.path.join(output_dir, "deterministic_results.json"), "w") as f:
        json.dump(det_results, f, indent=2)

    # 2. Save unmatched fields
    unmatched_list = [
        {
            "partner_field": u.partner_field,
            "column_category": u.column_category,
            "entity": u.entity
        }
        for u in unmatched_results
    ]

    with open(os.path.join(output_dir, "unmatched_fields.json"), "w") as f:
        json.dump(unmatched_list, f, indent=2)

    # 3. Save entity contexts (field lists for LLM)
    with open(os.path.join(output_dir, "entity_contexts.json"), "w") as f:
        json.dump(entity_contexts, f, indent=2)

    # 4. Create the LLM response template
    template = {
        "client": client_name,
        "process": process_name,
        "mappings": []
    }
    for entity, ctx in entity_contexts.items():
        for field in ctx["fields_to_map"]:
            template["mappings"].append({
                "partner_field": field["field_name"],
                "column_category": field.get("category"),
                "entity": entity,
                "matched_excel_key": "__FILL_THIS__",
                "json_key": "__FILL_THIS__",
                "confidence": 0.0,
                "reasoning": "__FILL_THIS__",
                "matched_pattern": "__FILL_THIS__"
            })

    with open(os.path.join(output_dir, "llm_response_template.json"), "w") as f:
        json.dump(template, f, indent=2)

    return {
        "deterministic_results": os.path.join(output_dir, "deterministic_results.json"),
        "unmatched_fields": os.path.join(output_dir, "unmatched_fields.json"),
        "entity_contexts": os.path.join(output_dir, "entity_contexts.json"),
        "llm_template": os.path.join(output_dir, "llm_response_template.json")
    }


def print_summary(all_fields, matched, unmatched, entity_contexts):
    """Print a human-readable summary."""
    from collections import Counter

    total = len(all_fields)
    n_matched = len(matched)
    n_unmatched = len(unmatched)

    print(f"\n{'='*60}")
    print(f"  DETERMINISTIC MATCHING SUMMARY")
    print(f"{'='*60}")
    print(f"  Total fields parsed:     {total}")
    print(f"  Deterministic matches:   {n_matched} ({100*n_matched/total:.0f}%)")
    print(f"  Needs LLM matching:      {n_unmatched} ({100*n_unmatched/total:.0f}%)")

    # By match type
    type_counts = Counter(m.match_type for m in matched)
    print(f"\n  Match type breakdown:")
    for mt, cnt in type_counts.most_common():
        print(f"    {mt:20}: {cnt}")

    # By entity (unmatched)
    entity_counts = Counter(u.entity for u in unmatched)
    print(f"\n  Unmatched by entity:")
    for entity, ctx in entity_contexts.items():
        cnt = len(ctx["fields_to_map"])
        avail = ctx["field_count"]
        print(f"    {entity:15}: {cnt} fields (searching among {avail} {ctx['role']} fields)")

    # Confidence distribution
    confs = [m.confidence for m in matched]
    if confs:
        high = sum(1 for c in confs if c >= 0.90)
        med = sum(1 for c in confs if 0.80 <= c < 0.90)
        low = sum(1 for c in confs if c < 0.80)
        print(f"\n  Confidence distribution (deterministic):")
        print(f"    High (≥0.90):  {high}")
        print(f"    Medium (0.80-0.89): {med}")
        print(f"    Low (<0.80):   {low}")

    print(f"{'='*60}\n")


def main():
    parser = argparse.ArgumentParser(description="LP Field Mapping - Full Pipeline Runner")
    parser.add_argument("input_file", help="Partner field list (Excel or JSON)")
    parser.add_argument("client_name", help="Client/partner name")
    parser.add_argument("--process", default="COMBINED", help="Process type: COMBINED or SEPARATE")
    parser.add_argument("--refs", default=None, help="References directory path")
    parser.add_argument("--output", default=None, help="Output directory for intermediate files")
    parser.add_argument("--sheet", default=None, help="Filter to specific sheet name")

    args = parser.parse_args()

    # Find references directory
    script_dir = Path(__file__).parent.parent
    refs_dir = args.refs or str(script_dir / "references")

    if not Path(refs_dir).exists():
        print(f"ERROR: References directory not found: {refs_dir}")
        print("Run build_references.py first.")
        sys.exit(1)

    # Output directory
    output_dir = args.output or str(script_dir / "pipeline_output" / args.client_name)

    print(f"Loading references from: {refs_dir}")
    refs = load_references(refs_dir)

    print(f"Parsing input file: {args.input_file}")
    all_fields, matched, unmatched = run_deterministic_pass(
        args.input_file, refs, args.process, args.sheet
    )

    print(f"Preparing LLM context for {len(unmatched)} unmatched fields...")
    entity_contexts = prepare_llm_context(unmatched, refs, args.client_name, args.process)

    # Print summary
    print_summary(all_fields, matched, unmatched, entity_contexts)

    # Write intermediate files
    files = write_intermediate_files(
        matched, unmatched, entity_contexts,
        output_dir, args.client_name, args.process
    )

    print(f"Intermediate files written to: {output_dir}")
    for name, path in files.items():
        print(f"  {name}: {path}")


if __name__ == "__main__":
    main()
