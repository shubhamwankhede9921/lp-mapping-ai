#!/usr/bin/env python3
"""
Finalize Pipeline - Combine deterministic + LLM results and generate output Excel.

This script takes the deterministic results and LLM mapping results,
combines them, runs post-processing (numbering, json_key resolution),
and generates the final output Excel.

Usage:
    python finalize_pipeline.py <pipeline_output_dir> <output_excel_path> [--client CLIENT]
"""

import argparse
import json
import os
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from generate_output import generate_mapping_excel
from matching_engine import load_references


def load_deterministic_results(pipeline_dir):
    """Load deterministic matching results."""
    path = os.path.join(pipeline_dir, "deterministic_results.json")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_llm_results(pipeline_dir):
    """Load LLM matching results."""
    path = os.path.join(pipeline_dir, "llm_results.json")
    if not os.path.exists(path):
        print(f"WARNING: LLM results file not found: {path}")
        return []
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, list):
        return data
    if isinstance(data, dict) and "mappings" in data:
        return data["mappings"]
    if isinstance(data, dict):
        return _parse_flat_llm_results(data, pipeline_dir)
    raise ValueError(f"Unsupported llm_results.json format: {type(data).__name__}")


def _load_unmatched_context(pipeline_dir):
    """Load unmatched field context for reconstructing flat LLM outputs."""
    path = os.path.join(pipeline_dir, "unmatched_fields.json")
    if not os.path.exists(path):
        return {}

    with open(path, "r", encoding="utf-8") as f:
        items = json.load(f)

    context = {}
    for item in items:
        partner_field = item.get("partner_field")
        if partner_field:
            context[partner_field] = {
                "column_category": item.get("column_category"),
                "entity": item.get("entity", "OTHER"),
            }
    return context


def _parse_pipe_value(partner_field, value_str, unmatched_context):
    """Parse flat LLM value format: ek|json_key|confidence|matched_pattern|reasoning."""
    parts = value_str.split("|", 4)
    if len(parts) != 5:
        print(
            f"  WARNING: Skipping malformed LLM value for '{partner_field}' "
            f"(expected 5 parts, got {len(parts)})"
        )
        return None

    matched_excel_key, json_key, confidence_str, matched_pattern, reasoning = parts
    try:
        confidence = float(confidence_str.strip())
    except ValueError:
        print(
            f"  WARNING: Invalid confidence '{confidence_str}' for '{partner_field}', "
            "defaulting to 0.0"
        )
        confidence = 0.0

    ctx = unmatched_context.get(partner_field, {})
    return {
        "partner_field": partner_field,
        "column_category": ctx.get("column_category"),
        "entity": ctx.get("entity", "OTHER"),
        "matched_excel_key": matched_excel_key.strip(),
        "json_key": json_key.strip(),
        "confidence": round(confidence, 4),
        "match_type": f"llm_{matched_pattern.strip().lower()}",
        "reasoning": reasoning.strip(),
        "needs_review": confidence < 0.80,
    }


def _parse_flat_llm_results(data, pipeline_dir):
    """Parse flat JSON object keyed by partner_field into mapping dicts."""
    unmatched_context = _load_unmatched_context(pipeline_dir)
    parsed = []

    for partner_field, value_str in data.items():
        if not isinstance(value_str, str):
            print(
                f"  WARNING: Skipping '{partner_field}' because value is "
                f"{type(value_str).__name__}, not string"
            )
            continue
        mapping = _parse_pipe_value(partner_field, value_str, unmatched_context)
        if mapping:
            parsed.append(mapping)

    print(f"  Parsed flat-format LLM results: {len(parsed)} mappings")
    return parsed


def resolve_json_keys(mappings, refs):
    """Resolve canonical json_keys from the field dictionary."""
    by_excel_key = refs["field_dictionary"].get("by_excel_key", {})
    customer_param_re = re.compile(r"^LOANAPPLICANTPARAM(\d+)$")

    for mapping in mappings:
        excel_key = mapping.get("matched_excel_key", "")
        if not excel_key:
            continue

        customer_param_match = customer_param_re.match(excel_key)
        if customer_param_match:
            index = int(customer_param_match.group(1))
            mapping["json_key"] = (
                "loanAccount.applicant.loanCustomerRelations."
                f"loanCustomerParameters.{index - 1}.value"
            )
            continue

        entry = by_excel_key.get(excel_key, {})
        canonical_json_key = entry.get("json_key", "")
        if canonical_json_key:
            mapping["json_key"] = canonical_json_key
        elif not mapping.get("json_key") or mapping["json_key"] in ("", "__FILL_THIS__", "null"):
            mapping["json_key"] = ""
    return mappings


def resolve_entities_from_targets(mappings):
    """Derive final entity from resolved internal target paths."""
    updated = 0

    for mapping in mappings:
        excel_key = mapping.get("matched_excel_key", "")
        json_key = (mapping.get("json_key") or "").lower()
        current_entity = mapping.get("entity", "OTHER")
        resolved_entity = current_entity

        if "loancustomerrelations.loancustomerparameters" in json_key:
            if "coapplicant." in json_key:
                match = re.search(r"coapplicant\.(\d+)", json_key)
                if match:
                    resolved_entity = f"COAPPLICANT{int(match.group(1)) + 1}"
                else:
                    resolved_entity = "COAPPLICANT"
            else:
                resolved_entity = "APPLICANT"
        elif "loanparameters" in json_key or excel_key.startswith("LOANPARAMETER"):
            resolved_entity = "LOAN"
        elif "coapplicant." in json_key:
            match = re.search(r"coapplicant\.(\d+)", json_key)
            if match:
                resolved_entity = f"COAPPLICANT{int(match.group(1)) + 1}"
            else:
                resolved_entity = "COAPPLICANT"
        elif "loanaccount.customer." in json_key or json_key.startswith("customer."):
            resolved_entity = "APPLICANT"
        elif any(token in json_key for token in ("loandocuments", "document", "imageid", "fileid")):
            resolved_entity = "DOCUMENT"
        elif excel_key.startswith("FEE"):
            resolved_entity = "FEE"

        if resolved_entity != current_entity:
            mapping["entity"] = resolved_entity
            updated += 1

    if updated > 0:
        print(f"  Resolved final entity from internal targets for {updated} mappings")
    return mappings


def validate_excel_keys(mappings, refs):
    """Replace invalid excel keys with a safe fallback for review."""
    by_excel_key = refs["field_dictionary"].get("by_excel_key", {})
    invalid_count = 0

    for mapping in mappings:
        excel_key = mapping.get("matched_excel_key", "")
        if not excel_key or excel_key in by_excel_key:
            continue

        if excel_key in (
            "DOCUMENTNAME",
            "DOCUMENTID",
            "FEE",
            "LOANPARAMETER",
            "CUSTOMERPARAMETER",
            "LOANAPPLICANTPARAM",
        ):
            continue

        old_excel_key = excel_key
        mapping["matched_excel_key"] = "LOANPARAMETER"
        mapping["json_key"] = ""
        mapping["confidence"] = 0.60
        mapping["needs_review"] = True
        mapping["reasoning"] = (
            f"Stale alias pointed to {old_excel_key} which does not exist in field dictionary; "
            "rerouted to LOANPARAMETER for manual review"
        )
        invalid_count += 1

    if invalid_count > 0:
        print(
            f"  WARNING: {invalid_count} mapping(s) had invalid excel_keys "
            "and were rerouted to LOANPARAMETER"
        )
    return mappings


def filter_junk_fields(mappings):
    """Remove rows that look like section headers or placeholders."""
    junk_exact = {
        "newly added fields", "as", "ads", "na", "n/a", "test", "tbd",
        "todo", "pending", "blank", "none", "null", "header", "section",
        "origination data", "customer basic details", "kyc & compliance details",
        "kyc and compliance details", "banks details", "bank details",
        "jewel details", "loan details", "loan document", "loan documents",
        "personal details", "address details", "document details",
        "employment details", "reference details", "co-applicant details",
        "guarantor details", "collateral details", "disbursement details",
    }

    filtered = []
    removed = []
    for mapping in mappings:
        partner_field = (mapping.get("partner_field") or "").strip().lower()
        if partner_field in junk_exact or len(partner_field) <= 1 or not partner_field:
            removed.append(mapping.get("partner_field", ""))
            continue
        filtered.append(mapping)

    if removed:
        print(f"  Filtered {len(removed)} junk/header rows: {removed}")
    return filtered


SPECIAL_PREFIXES = (
    "DOCUMENTNAME",
    "DOCUMENTID",
    "FEE",
    "LOANPARAMETER",
    "CUSTOMERPARAMETER",
    "LOANAPPLICANTPARAM",
)
SPECIAL_RE = re.compile(
    r"^(DOCUMENTNAME|DOCUMENTID|FEE|LOANPARAMETER|CUSTOMERPARAMETER|LOANAPPLICANTPARAM)\d*$"
)


def strip_special_field_numbers(mappings):
    """Strip number suffixes so special fields can be renumbered fresh."""
    stripped = 0
    for mapping in mappings:
        excel_key = mapping.get("matched_excel_key", "")
        if not excel_key or not SPECIAL_RE.match(excel_key):
            continue

        for prefix in SPECIAL_PREFIXES:
            if excel_key.startswith(prefix):
                if excel_key != prefix:
                    mapping["matched_excel_key"] = prefix
                    mapping["json_key"] = ""
                    stripped += 1
                break

    if stripped > 0:
        print(
            f"  Stripped number suffixes from {stripped} special fields for fresh re-numbering"
        )
    return mappings


def filter_no_mapping_required(mappings):
    """Remove generic system fields that should not be mapped."""
    no_map_fields = {
        "isprimary", "remark", "remarks", "extension", "isactive",
        "isdeleted", "createdby", "modifiedby", "updatedby",
        "createddate", "modifieddate", "updateddate", "rowversion",
        "isverified", "verificationstatus",
    }

    filtered = []
    removed = []
    for mapping in mappings:
        partner_field = (mapping.get("partner_field") or "").strip().lower()
        if partner_field in no_map_fields:
            removed.append(mapping.get("partner_field", ""))
            continue
        filtered.append(mapping)

    if removed:
        print(f"  Excluded {len(removed)} no-mapping-required fields: {removed}")
    return filtered


LOAN_LEVEL_FIELDS = {
    "irr", "iir", "foir", "ltv", "loantovalueratio", "loan_to_value",
    "marginmoney", "margin_money", "downpayment", "down_payment",
    "schemename", "scheme_name", "schemeid", "scheme_id", "schemecode",
    "subproduct", "sub_product", "subproductcategory",
    "reduceemistatus", "addchargesstatus", "addlsistatus", "lsiamount",
    "lsiamt", "reducedemi",
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


def _normalize_hint_text(value):
    if not value:
        return ""
    return re.sub(r"[^a-z0-9]+", " ", value.strip().lower()).strip()


def _is_customer_related_loan_parameter(mapping):
    entity = (mapping.get("entity") or "").upper()
    category = _normalize_hint_text(
        mapping.get("column_category") or mapping.get("category") or ""
    )
    field = _normalize_hint_text(mapping.get("partner_field") or "")
    compact_field = field.replace(" ", "")

    if entity in ("APPLICANT", "CUSTOMER", "COAPPLICANT", "COAPPLICANT1", "COAPPLICANT2", "COAPPLICANT3", "COAPPLICANT4"):
        return compact_field not in LOAN_LEVEL_FIELDS

    if any(hint in category for hint in CUSTOMER_PARAM_CATEGORY_HINTS):
        return compact_field not in LOAN_LEVEL_FIELDS

    if (
        entity == "LOAN"
        and any(hint in category for hint in LOAN_LEVEL_CATEGORY_HINTS)
        and any(hint in field for hint in CUSTOMER_PARAM_CATEGORY_HINTS)
    ):
        return True

    return False


def route_customer_params(mappings):
    """Route parameter fallbacks to customer, fee, or loan buckets by entity/category/field."""
    total_loan_parameter_candidates = 0
    rerouted_customer = 0
    rerouted_fee = 0
    kept_loan_level = 0

    for mapping in mappings:
        excel_key = mapping.get("matched_excel_key", "")
        entity = (mapping.get("entity") or "").upper()

        if excel_key == "LOANPARAMETER" and entity == "FEE":
            total_loan_parameter_candidates += 1
            mapping["matched_excel_key"] = "FEE"
            rerouted_fee += 1
        elif excel_key == "LOANPARAMETER":
            total_loan_parameter_candidates += 1
            if _is_customer_related_loan_parameter(mapping):
                mapping["matched_excel_key"] = "LOANAPPLICANTPARAM"
                rerouted_customer += 1
            else:
                kept_loan_level += 1

    if total_loan_parameter_candidates > 0:
        print(
            "  LOANPARAMETER routing summary: "
            f"total={total_loan_parameter_candidates}, "
            f"converted_to_LOANAPPLICANTPARAM={rerouted_customer}, "
            f"rerouted_to_FEE={rerouted_fee}, "
            f"kept_as_LOANPARAMETER={kept_loan_level}"
        )
    if rerouted_customer > 0:
        print(
            f"  Rerouted {rerouted_customer} customer-related LOANPARAMETER fields "
            "to LOANAPPLICANTPARAM"
        )
    if rerouted_fee > 0:
        print(f"  Rerouted {rerouted_fee} FEE-entity fields to FEE")
    if kept_loan_level > 0:
        print(
            f"  Kept {kept_loan_level} loan-level fields as LOANPARAMETER"
        )
    return mappings


def route_reference_entity(mappings):
    """Route reference/verifier fields to BUSINESSVERIFICATION excel keys."""
    ref_categories = {
        "reference", "leadreference", "verification", "verifier",
        "ref", "references", "verifications",
    }
    ref_field_map = {
        "name": "BUSINESSVERIFICATIONNAME",
        "firstname": "BUSINESSVERIFICATIONNAME",
        "referencename": "BUSINESSVERIFICATIONNAME",
        "ref_name": "BUSINESSVERIFICATIONNAME",
        "mobile": "BUSINESSVERIFICATIONVERIFICATIONPHONE",
        "mobilenumber": "BUSINESSVERIFICATIONVERIFICATIONPHONE",
        "phone": "BUSINESSVERIFICATIONVERIFICATIONPHONE",
        "phonenumber": "BUSINESSVERIFICATIONVERIFICATIONPHONE",
        "contactnumber": "BUSINESSVERIFICATIONVERIFICATIONPHONE",
        "address": "BUSINESSVERIFICATIONADDRESS",
        "referenceaddress": "BUSINESSVERIFICATIONADDRESS",
        "ref_address": "BUSINESSVERIFICATIONADDRESS",
        "relation": "BUSINESSVERIFICATIONENTERPRISERELATION1TYPE",
        "relationship": "BUSINESSVERIFICATIONENTERPRISERELATION1TYPE",
        "referencerelation": "BUSINESSVERIFICATIONENTERPRISERELATION1TYPE",
    }

    rerouted = 0
    for mapping in mappings:
        category = (
            mapping.get("column_category")
            or mapping.get("category")
            or ""
        ).strip().lower().replace(" ", "")
        if category not in ref_categories:
            continue

        partner_field = (mapping.get("partner_field") or "").strip().lower()
        target_excel_key = ref_field_map.get(partner_field)
        if not target_excel_key:
            continue

        mapping["matched_excel_key"] = target_excel_key
        mapping["entity"] = "APPLICANT"
        mapping["confidence"] = max(mapping.get("confidence", 0), 0.85)
        mapping["reasoning"] = (
            f"Reference entity routing: {mapping.get('partner_field', '')} -> {target_excel_key}"
        )
        rerouted += 1

    if rerouted > 0:
        print(f"  Rerouted {rerouted} reference entity fields to BUSINESSVERIFICATION* keys")
    return mappings


def auto_number_special_fields(mappings):
    """Assign fresh sequential numbers to special field families."""
    document_name_counter = 1
    document_id_counter = 1
    fee_counter = 1
    loan_param_counter = 1
    customer_param_counter = 1

    for mapping in mappings:
        excel_key = mapping.get("matched_excel_key", "")
        if not excel_key:
            continue

        if excel_key == "DOCUMENTNAME":
            mapping["matched_excel_key"] = f"DOCUMENTNAME{document_name_counter}"
            document_name_counter += 1
        elif excel_key == "DOCUMENTID":
            mapping["matched_excel_key"] = f"DOCUMENTID{document_id_counter}"
            document_id_counter += 1
        elif excel_key == "FEE":
            mapping["matched_excel_key"] = f"FEE{fee_counter}"
            fee_counter += 1
        elif excel_key == "LOANPARAMETER":
            mapping["matched_excel_key"] = f"LOANPARAMETER{loan_param_counter}"
            loan_param_counter += 1
        elif excel_key in ("CUSTOMERPARAMETER", "LOANAPPLICANTPARAM"):
            mapping["matched_excel_key"] = f"LOANAPPLICANTPARAM{customer_param_counter}"
            customer_param_counter += 1

    return mappings


def validate_mandatory_fields(mappings):
    """Report mandatory LOS fields missing from the combined mapping."""
    mandatory_fields = {
        "APPLICANTCUSTOMERNAME": {
            "json_key": "loanAccount.customer.firstName",
            "description": "Customer first name",
        },
        "APPLICANTLASTNAME": {
            "json_key": "loanAccount.customer.lastName",
            "description": "Customer last name",
        },
        "APPLICANTDATEOFBIRTH": {
            "json_key": "loanAccount.customer.dateOfBirth",
            "description": "Date of birth",
        },
        "APPLICANTCUSTOMERGENDER": {
            "json_key": "loanAccount.customer.gender",
            "description": "Gender",
        },
        "APPLICANTCUSTOMERMOBILENO": {
            "json_key": "loanAccount.customer.mobilePhone",
            "description": "Mobile number",
        },
        "APPLICANTCUSTOMERPANNO": {
            "json_key": "loanAccount.customer.panNo",
            "description": "PAN number",
        },
        "APPLICANTCUSTOMERPINCODE": {
            "json_key": "loanAccount.customer.pincode",
            "description": "Pincode",
        },
        "APPLICANTEMAILID": {
            "json_key": "loanAccount.customer.email",
            "description": "Email ID",
        },
        "LOANAMT": {
            "json_key": "loanAccount.loanAmount",
            "description": "Loan amount",
        },
        "INTERESTRATE": {
            "json_key": "loanAccount.interestRate",
            "description": "Interest rate",
        },
        "PRODUCTCODE": {
            "json_key": "loanAccount.productCode",
            "description": "Product code",
        },
        "TENURE": {
            "json_key": "loanAccount.tenure",
            "description": "Loan tenure",
        },
        "REPAYMENTFREQUENCY": {
            "json_key": "loanAccount.frequency",
            "description": "Repayment frequency",
        },
        "PARTNERCODE": {
            "json_key": "loanAccount.partnerCode",
            "description": "Partner code",
        },
        "LOANAPPLICATIONDATE": {
            "json_key": "loanAccount.loanApplicationDate",
            "description": "Loan application date",
        },
        "CBSCORE": {
            "json_key": "loanAccount.loanAccountAdditional.creditBureauScore",
            "description": "Credit bureau score",
        },
        "BORROWERINCOME": {
            "json_key": "loanAccount.loanAccountAdditional.netIncomeOfBorrower",
            "description": "Net income of borrower",
        },
        "APPLICANTPROPOSEDEMI": {
            "json_key": "loanAccount.estimatedEmi",
            "description": "Monthly EMI",
        },
        "EMPLOYMENTSTATUS": {
            "json_key": "loanAccount.customer.employmentStatus",
            "description": "Employment status",
        },
        "CUSTOMERID": {
            "json_key": "loanAccount.customFields.customerId",
            "description": "Customer ID",
        },
        "LOANID": {
            "json_key": "loanAccount.loanId",
            "description": "Loan ID (partner)",
        },
    }

    mapped_keys = set()
    for mapping in mappings:
        excel_key = mapping.get("matched_excel_key", "")
        if excel_key:
            mapped_keys.add(excel_key.upper())

    missing = []
    for excel_key, info in mandatory_fields.items():
        if excel_key.upper() not in mapped_keys:
            missing.append({"excel_key": excel_key, **info})

    if missing:
        print(
            f"  WARNING: mandatory field check found {len(missing)} missing field(s):"
        )
        for item in missing:
            print(f"    - {item['excel_key']:35s} ({item['description']})")
    else:
        print(f"  All {len(mandatory_fields)} mandatory fields are present")

    return mappings, missing


def combine_and_process(det_results, llm_results, refs):
    """Combine deterministic + LLM results, process, and return final mappings."""
    for mapping in llm_results:
        if "needs_review" not in mapping:
            mapping["needs_review"] = mapping.get("confidence", 0) < 0.80
        if "match_type" not in mapping:
            mapping["match_type"] = mapping.get("matched_pattern", "llm_semantic")

        match_type = mapping.get("match_type", "").lower()
        if match_type in ("exact_match", "exact"):
            mapping["match_type"] = "llm_exact"
        elif match_type in ("semantic", "semantic_match"):
            mapping["match_type"] = "llm_semantic"
        elif match_type in ("pattern", "pattern_match"):
            mapping["match_type"] = "llm_pattern"
        elif match_type in ("none", ""):
            mapping["match_type"] = "llm_fallback"

    all_mappings = det_results + llm_results
    all_mappings = filter_junk_fields(all_mappings)
    all_mappings = filter_no_mapping_required(all_mappings)
    all_mappings = validate_excel_keys(all_mappings, refs)
    all_mappings = strip_special_field_numbers(all_mappings)
    all_mappings = route_reference_entity(all_mappings)
    all_mappings = route_customer_params(all_mappings)
    all_mappings = auto_number_special_fields(all_mappings)
    all_mappings = resolve_json_keys(all_mappings, refs)
    all_mappings = resolve_entities_from_targets(all_mappings)
    all_mappings, missing_mandatory = validate_mandatory_fields(all_mappings)
    return all_mappings, missing_mandatory


def _make_partner_api_key(partner_field, entity):
    partner_field = (partner_field or "").strip()
    words = (
        partner_field
        .replace("'s", "")
        .replace("/", " ")
        .replace(".", " ")
        .replace("%", "Percent")
        .split()
    )
    if not words:
        return partner_field

    camel_case = words[0].lower() + "".join(word.capitalize() for word in words[1:])
    if (entity or "").upper() in ("APPLICANT", "CUSTOMER", "COAPPLICANT"):
        return f"Applicant.{camel_case}"
    return f"loanAccount.{camel_case}"


def _make_ui_key(json_key):
    json_key = (json_key or "").strip()
    if not json_key:
        return ""
    if "loanCustomerParameters" in json_key:
        return json_key.replace("loanAccount.applicant.", "loanAccount.customer.")
    if "loanParameters" in json_key or "loanDocuments" in json_key:
        return json_key
    if json_key.startswith("loanAccount.customer.") or json_key.startswith("customer."):
        if "customerBankAccounts" in json_key or "bankStatements" in json_key:
            return f"customer.{json_key.split('.')[-1]}"
        if "loanAccount.customer." in json_key:
            after = json_key.split("loanAccount.customer.", 1)[1]
        else:
            after = json_key.split("customer.", 1)[1]
        return f"customer.{after}"
    if json_key.startswith("loanAccount."):
        parts = json_key.split(".")
        if len(parts) == 2:
            return json_key
        return f"loanAccount.{parts[-1]}"
    return json_key


def _generate_insert_sql(mappings, sql_path):
    """Generate SQL INSERT for generic_excel_upload_definition_fields."""
    from datetime import datetime

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S.0")
    lines = [
        "INSERT INTO financialForms.generic_excel_upload_definition_fields "
        "(version,master_id,excel_column_name,table_column_name,`type`,required,`default`,"
        "partner_api_key,ui_type,ui_title,ui_key,additional_key,ui_grouping,ui_field_condition,"
        "ui_field_order,can_update_stage,file_json_path,created_at,updated_at,created_by,"
        "description,example,required_stage,file_storage,`length`,updated_by,column_indices,"
        "effective_date,is_video_file,file_extension) VALUES"
    ]

    value_lines = []
    for mapping in mappings:
        excel_column = (mapping.get("matched_excel_key") or "").replace("'", "''")
        partner_field = mapping.get("partner_field", "")
        entity = mapping.get("entity", "")
        json_key = mapping.get("json_key", "")
        partner_api_key = _make_partner_api_key(partner_field, entity).replace("'", "''")
        ui_key = _make_ui_key(json_key).replace("'", "''")
        value_lines.append(
            f"(0,2,'{excel_column}','{excel_column}','','Non Mandatory','','{partner_api_key}',"
            f"NULL,NULL,'{ui_key}',NULL,NULL,NULL,NULL,'ALL',NULL,"
            f"'{now}',NULL,NULL,NULL,NULL,'ALL',NULL,NULL,NULL,NULL,NULL,0,NULL)"
        )

    lines.append(",\n".join(value_lines))
    lines.append(";")

    with open(sql_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"  SQL INSERT generated: {len(value_lines)} rows")


def main():
    parser = argparse.ArgumentParser(description="Finalize LP Field Mapping Pipeline")
    parser.add_argument("pipeline_dir", help="Pipeline output directory containing intermediate files")
    parser.add_argument("output_path", help="Path for the output Excel file")
    parser.add_argument("--client", default="", help="Client/partner name")
    parser.add_argument("--process", default="COMBINED", help="Process type")
    parser.add_argument("--refs", default=None, help="References directory path")

    args = parser.parse_args()

    script_dir = Path(__file__).parent.parent
    refs_dir = args.refs or str(script_dir / "references")
    refs = load_references(refs_dir)

    print(f"Loading deterministic results from: {args.pipeline_dir}")
    det_results = load_deterministic_results(args.pipeline_dir)
    print(f"  Deterministic: {len(det_results)} mappings")

    print(f"Loading LLM results from: {args.pipeline_dir}")
    llm_results = load_llm_results(args.pipeline_dir)
    print(f"  LLM: {len(llm_results)} mappings")

    all_mappings, missing_mandatory = combine_and_process(det_results, llm_results, refs)
    print(f"  Total: {len(all_mappings)} mappings")

    generate_mapping_excel(
        all_mappings,
        args.output_path,
        args.client,
        args.process,
        missing_mandatory=missing_mandatory,
    )
    print(f"\nOutput Excel: {args.output_path}")

    sql_path = args.output_path.replace(".xlsx", "_insert.sql")
    _generate_insert_sql(all_mappings, sql_path)
    print(f"SQL INSERT: {sql_path}")

    review_count = sum(1 for mapping in all_mappings if mapping.get("needs_review", False))
    high_confidence_count = sum(
        1 for mapping in all_mappings if mapping.get("confidence", 0) >= 0.90
    )
    print(f"  High confidence: {high_confidence_count}")
    print(f"  Needs review: {review_count}")


if __name__ == "__main__":
    main()
