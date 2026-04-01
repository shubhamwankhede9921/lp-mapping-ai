#!/usr/bin/env python3
"""
Orchestrator for LP Field Mapping

Main orchestrator that ties everything together. Handles:
1. Building/refreshing reference files
2. Parsing input files (Excel or JSON)
3. Running deterministic matching
4. Generating LLM prompts for unmatched fields
5. Combining results and generating output

Can be used as a module or CLI tool.

Usage as module:
    from orchestrator import run_mapping, finalize_mapping

    # Step 1: Deterministic pass
    result = run_mapping(
        input_file="partner_fields.xlsx",
        client_name="ACME",
        process_name="Loan Processing"
    )
    print(result['stats'])

    # User feeds LLM prompts to Claude and gets responses

    # Step 2: Finalize with LLM results
    output_file = finalize_mapping(
        deterministic_results=result['deterministic_results'],
        llm_results=llm_responses,
        field_dictionary=result['field_dictionary'],
        output_path="./output/mapping.xlsx",
        client_name="ACME",
        process_name="Loan Processing"
    )

Usage as CLI:
    python orchestrator.py --input partner_fields.xlsx --client "ACME" --process "COMBINED"
    python orchestrator.py --input payload.json --client "KB" --process "SEPARATE" --output ./output/

    python orchestrator.py finalize --input partner_fields.xlsx --llm-responses llm_results.json
"""

import argparse
import json
import logging
import sys
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Any, Tuple

from input_parser import parse_input
from matching_engine import load_references, match_batch, MatchResult
from llm_mapper import generate_mapping_prompt
from post_processor import PostProcessor
from generate_output import generate_mapping_excel, generate_api_config_json


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def find_putm_dump(base_path: Optional[str] = None) -> Optional[Path]:
    """
    Find putm_dump.xlsx in workspace.

    Searches in:
    1. Provided base_path
    2. Parent of scripts directory
    3. Parent of parent
    4. Current working directory
    """
    search_paths = [base_path] if base_path else []

    # Add script parent directories
    script_dir = Path(__file__).parent
    search_paths.extend([
        script_dir.parent,
        script_dir.parent.parent,
        Path.cwd()
    ])

    for search_path in search_paths:
        if search_path is None:
            continue
        search_path = Path(search_path)
        dump_path = search_path / 'putm_dump.xlsx'
        if dump_path.exists():
            logger.info(f"Found putm_dump.xlsx at {dump_path}")
            return dump_path

    logger.warning("Could not find putm_dump.xlsx in standard locations")
    return None


def should_rebuild_references(
    references_dir: Path,
    source_files: List[Path]
) -> bool:
    """
    Check if references should be rebuilt.

    Returns True if:
    - References don't exist
    - Any source file is newer than references
    """
    references_dir = Path(references_dir)

    # Check if references exist
    field_dict_path = references_dir / 'field_dictionary.json'
    if not field_dict_path.exists():
        logger.info("References don't exist, will rebuild")
        return True

    ref_mtime = field_dict_path.stat().st_mtime

    # Check if any source file is newer
    for source in source_files:
        if source.exists():
            source_mtime = source.stat().st_mtime
            if source_mtime > ref_mtime:
                logger.info(f"Source file {source.name} is newer, will rebuild")
                return True

    logger.info("References are up to date")
    return False


def build_or_load_references(
    references_dir: Optional[str] = None,
    putm_dump_path: Optional[str] = None
) -> Dict[str, Any]:
    """
    Build or load reference files.

    Args:
        references_dir: Directory to store/load references from
        putm_dump_path: Path to putm_dump.xlsx (auto-finds if not provided)

    Returns:
        References dict with keys: field_dictionary, alias_registry, entity_routing
    """
    if references_dir is None:
        references_dir = Path(__file__).parent / 'references'
    else:
        references_dir = Path(references_dir)

    references_dir.mkdir(parents=True, exist_ok=True)

    # Find source files
    if putm_dump_path is None:
        putm_dump_path = find_putm_dump()
    else:
        putm_dump_path = Path(putm_dump_path)

    if putm_dump_path and not putm_dump_path.exists():
        raise FileNotFoundError(f"putm_dump.xlsx not found at {putm_dump_path}")

    # Check if rebuild is needed
    source_files = [putm_dump_path] if putm_dump_path else []
    if should_rebuild_references(references_dir, source_files):
        logger.info("Building references...")
        try:
            from build_references import build_all_references
            build_all_references(
                xlsx_path=str(putm_dump_path) if putm_dump_path else None,
                output_dir=str(references_dir)
            )
        except Exception as e:
            logger.warning(f"Could not rebuild references: {e}")
            logger.info("Will attempt to load existing references")

    # Load references
    logger.info("Loading references...")
    return load_references(str(references_dir))


def run_mapping(
    input_file: str,
    putm_dump_path: Optional[str] = None,
    manual_mappings_path: Optional[str] = None,
    references_dir: Optional[str] = None,
    client_name: str = "",
    process_name: str = "",
    output_dir: Optional[str] = None
) -> Dict[str, Any]:
    """
    Full mapping pipeline - deterministic pass only.

    This generates results that can be combined with LLM responses.

    Args:
        input_file: Path to partner input file (Excel or JSON)
        putm_dump_path: Path to putm_dump.xlsx (auto-finds if not provided)
        manual_mappings_path: Path to manual mappings CSV (optional)
        references_dir: Directory to store/load references from
        client_name: Client name for metadata
        process_name: Process name for metadata
        output_dir: Directory to write intermediate files to

    Returns:
        {
            "status": "deterministic_complete",
            "deterministic_results": [...],  # MatchResult dicts
            "unmatched_fields": [...],  # Fields needing LLM
            "llm_prompts": [...],  # Entity-grouped prompts
            "stats": {
                "total_fields": N,
                "deterministic_matches": N,
                "needs_llm": N,
                "by_match_type": {...},
                "by_entity": {...}
            },
            "field_dictionary": {...},  # For finalize_mapping
            "intermediate_state": {...}  # For resuming
        }
    """
    logger.info("=" * 60)
    logger.info("LP FIELD MAPPING - DETERMINISTIC PASS")
    logger.info("=" * 60)

    # Step 1: Build/load references
    logger.info("\nStep 1: Loading references...")
    references = build_or_load_references(references_dir, putm_dump_path)
    field_dictionary = references['field_dictionary']
    alias_registry = references['alias_registry']
    entity_routing = references['entity_routing']

    logger.info(f"Loaded {len(field_dictionary['all_excel_keys'])} fields from dictionary")
    logger.info(f"Loaded {len(alias_registry.get('aliases', {}))} aliases")

    # Step 2: Parse input file
    logger.info(f"\nStep 2: Parsing input file: {input_file}")
    partner_fields = parse_input(input_file)
    logger.info(f"Parsed {len(partner_fields)} fields from input")

    # Step 3: Run deterministic matching
    logger.info("\nStep 3: Running deterministic matching...")
    deterministic_results = match_batch(
        fields=partner_fields,
        references=references
    )

    logger.info(f"Matched {len(deterministic_results)} fields deterministically")

    # Separate matched and unmatched
    matched = [r for r in deterministic_results if r['matched_excel_key']]
    unmatched = [r for r in deterministic_results if not r['matched_excel_key']]

    logger.info(f"  - Matched: {len(matched)}")
    logger.info(f"  - Unmatched: {len(unmatched)}")

    # Log match type breakdown
    match_type_counts = {}
    for result in matched:
        mt = result['match_type']
        match_type_counts[mt] = match_type_counts.get(mt, 0) + 1

    logger.info("\nMatch type breakdown:")
    for match_type, count in sorted(match_type_counts.items()):
        logger.info(f"  - {match_type}: {count}")

    # Log entity breakdown
    entity_counts = {}
    for result in deterministic_results:
        entity = result['entity']
        entity_counts[entity] = entity_counts.get(entity, 0) + 1

    logger.info("\nEntity breakdown:")
    for entity, count in sorted(entity_counts.items()):
        logger.info(f"  - {entity}: {count}")

    # Step 4: Generate LLM prompts for unmatched fields
    logger.info(f"\nStep 4: Generating LLM prompts for {len(unmatched)} unmatched fields...")

    llm_prompts = []
    if unmatched:
        llm_prompts = generate_mapping_prompt(
            unmatched_fields=unmatched,
            references=references,
            client_name=client_name,
            process_name=process_name
        )
        logger.info(f"Generated {len(llm_prompts)} LLM prompts")

    # Build stats
    stats = {
        "total_fields": len(deterministic_results),
        "deterministic_matches": len(matched),
        "needs_llm": len(unmatched),
        "by_match_type": match_type_counts,
        "by_entity": entity_counts,
        "match_rate": round(len(matched) / len(deterministic_results) * 100, 1) if deterministic_results else 0
    }

    # Write intermediate files if output_dir specified
    if output_dir:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        # Write matched results
        matched_path = output_dir / 'deterministic_results.json'
        with open(matched_path, 'w') as f:
            json.dump(matched, f, indent=2)
        logger.info(f"Wrote deterministic results to {matched_path}")

        # Write unmatched fields
        unmatched_path = output_dir / 'unmatched_fields.json'
        with open(unmatched_path, 'w') as f:
            json.dump(unmatched, f, indent=2)
        logger.info(f"Wrote unmatched fields to {unmatched_path}")

        # Write LLM prompts
        if llm_prompts:
            prompts_path = output_dir / 'llm_prompts.txt'
            with open(prompts_path, 'w') as f:
                for i, prompt in enumerate(llm_prompts, 1):
                    f.write(f"\n{'='*60}\n")
                    f.write(f"PROMPT {i} OF {len(llm_prompts)}\n")
                    f.write(f"{'='*60}\n\n")
                    f.write(prompt)
                    f.write("\n\n")
            logger.info(f"Wrote {len(llm_prompts)} LLM prompts to {prompts_path}")

        # Write stats
        stats_path = output_dir / 'stats.json'
        with open(stats_path, 'w') as f:
            json.dump(stats, f, indent=2)
        logger.info(f"Wrote statistics to {stats_path}")

    return {
        "status": "deterministic_complete",
        "deterministic_results": matched,
        "unmatched_fields": unmatched,
        "llm_prompts": llm_prompts,
        "stats": stats,
        "field_dictionary": field_dictionary,
        "alias_registry": alias_registry,
        "entity_routing": entity_routing,
        "intermediate_state": {
            "input_file": str(input_file),
            "client_name": client_name,
            "process_name": process_name,
            "timestamp": datetime.utcnow().isoformat() + 'Z'
        }
    }


def finalize_mapping(
    deterministic_results: List[Dict],
    llm_results: List[Dict],
    field_dictionary: Dict,
    output_path: str,
    client_name: str = "",
    process_name: str = ""
) -> str:
    """
    After LLM returns results, combine with deterministic results,
    run post-processing, and generate output Excel.

    Args:
        deterministic_results: Results from run_mapping (matched fields)
        llm_results: Results from LLM mapping (unmatched fields that were mapped)
        field_dictionary: Reference field dictionary
        output_path: Path to write output Excel file
        client_name: Client name for metadata
        process_name: Process name for metadata

    Returns:
        Path to output Excel file
    """
    logger.info("=" * 60)
    logger.info("LP FIELD MAPPING - FINALIZATION")
    logger.info("=" * 60)

    # Step 1: Combine results
    logger.info("\nStep 1: Combining deterministic and LLM results...")
    all_results = deterministic_results + llm_results
    logger.info(f"Combined {len(deterministic_results)} + {len(llm_results)} = {len(all_results)} total mappings")

    # Step 2: Post-process
    logger.info("\nStep 2: Post-processing results...")
    pp = PostProcessor(field_dictionary)
    processed_results = pp.process_results(all_results)
    logger.info(f"Post-processed {len(processed_results)} results")

    # Step 3: Generate output Excel
    logger.info(f"\nStep 3: Generating output Excel: {output_path}")
    generate_mapping_excel(
        mappings=processed_results,
        output_path=output_path,
        client_name=client_name,
        process_name=process_name
    )

    # Step 4: Generate API config JSON
    logger.info("Step 4: Generating API config JSON...")
    json_path = str(Path(output_path).with_suffix('.json'))
    generate_api_config_json(
        mappings=processed_results,
        output_path=json_path,
        client_name=client_name,
        process_name=process_name
    )
    logger.info(f"Generated API config: {json_path}")

    logger.info("\n" + "=" * 60)
    logger.info("MAPPING COMPLETE")
    logger.info("=" * 60)
    logger.info(f"Output Excel: {output_path}")
    logger.info(f"Output JSON:  {json_path}")

    return output_path


def main_run_mapping(args):
    """CLI: Run deterministic mapping and generate prompts."""
    result = run_mapping(
        input_file=args.input,
        putm_dump_path=args.putm_dump,
        references_dir=args.references,
        client_name=args.client,
        process_name=args.process,
        output_dir=args.output
    )

    # Print summary
    print("\n" + "=" * 60)
    print("MAPPING SUMMARY")
    print("=" * 60)

    stats = result['stats']
    print(f"\nTotal fields: {stats['total_fields']}")
    print(f"Deterministic matches: {stats['deterministic_matches']} ({stats['match_rate']}%)")
    print(f"Needs LLM: {stats['needs_llm']}")

    print("\nMatch types:")
    for mt, count in sorted(stats['by_match_type'].items()):
        print(f"  {mt}: {count}")

    print("\nEntities:")
    for entity, count in sorted(stats['by_entity'].items()):
        print(f"  {entity}: {count}")

    # Print LLM prompt instructions
    if result['llm_prompts']:
        print("\n" + "=" * 60)
        print("NEXT STEPS")
        print("=" * 60)
        print(f"\n{len(result['llm_prompts'])} prompts generated for unmatched fields.")
        print("\nTo complete the mapping:")
        print("1. Copy the prompts from llm_prompts.txt")
        print("2. Feed them to Claude (in your skill or via API)")
        print("3. Save Claude's responses to a JSON file")
        print("4. Run: python orchestrator.py finalize --llm-responses <responses.json>")

    return result


def main_finalize(args):
    """CLI: Finalize mapping with LLM results."""
    # Load intermediate state
    if args.intermediate_state:
        with open(args.intermediate_state) as f:
            state = json.load(f)
    else:
        state = {}

    client_name = args.client or state.get('client_name', '')
    process_name = args.process or state.get('process_name', '')

    # Load results
    with open(args.deterministic) as f:
        det_results = json.load(f)

    with open(args.llm_responses) as f:
        llm_results = json.load(f)

    # Load field dictionary
    with open(args.field_dictionary) as f:
        field_dict = json.load(f)

    output_path = args.output or 'mapping_output.xlsx'

    finalize_mapping(
        deterministic_results=det_results,
        llm_results=llm_results,
        field_dictionary=field_dict,
        output_path=output_path,
        client_name=client_name,
        process_name=process_name
    )

    print(f"\nOutput file: {output_path}")


def main():
    """CLI entry point."""
    parser = argparse.ArgumentParser(
        description='LP Field Mapping Orchestrator'
    )

    subparsers = parser.add_subparsers(dest='command', help='Command to run')

    # Run mapping command
    run_parser = subparsers.add_parser('run', help='Run deterministic mapping')
    run_parser.add_argument(
        '--input',
        required=True,
        help='Input file path (Excel or JSON)'
    )
    run_parser.add_argument(
        '--client',
        default='',
        help='Client name for metadata'
    )
    run_parser.add_argument(
        '--process',
        default='',
        help='Process name for metadata'
    )
    run_parser.add_argument(
        '--putm-dump',
        help='Path to putm_dump.xlsx (auto-finds if not provided)'
    )
    run_parser.add_argument(
        '--references',
        help='Directory for reference files'
    )
    run_parser.add_argument(
        '--output',
        help='Output directory for intermediate files'
    )
    run_parser.set_defaults(func=main_run_mapping)

    # Finalize mapping command
    finalize_parser = subparsers.add_parser('finalize', help='Finalize with LLM results')
    finalize_parser.add_argument(
        '--deterministic',
        required=True,
        help='Path to deterministic_results.json'
    )
    finalize_parser.add_argument(
        '--llm-responses',
        required=True,
        help='Path to LLM responses JSON'
    )
    finalize_parser.add_argument(
        '--field-dictionary',
        required=True,
        help='Path to field_dictionary.json'
    )
    finalize_parser.add_argument(
        '--output',
        help='Output Excel file path'
    )
    finalize_parser.add_argument(
        '--client',
        help='Client name'
    )
    finalize_parser.add_argument(
        '--process',
        help='Process name'
    )
    finalize_parser.add_argument(
        '--intermediate-state',
        help='Path to intermediate state file'
    )
    finalize_parser.set_defaults(func=main_finalize)

    # Handle no command
    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        sys.exit(0)

    # Run command
    try:
        args.func(args)
    except Exception as e:
        logger.error(f"Error: {e}", exc_info=True)
        sys.exit(1)


if __name__ == '__main__':
    main()
