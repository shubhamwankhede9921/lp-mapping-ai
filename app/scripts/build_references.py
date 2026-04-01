#!/usr/bin/env python3
"""
scripts/build_references.py

Two ways to use this module:

1. CLI  (original behaviour — reads DB via mysql.connector / .env):
       python build_references.py

2. Import (called by mapping_service.build_references_from_db_direct):
       from build_references import FieldDictionaryBuilder, AliasRegistryBuilder, EntityRoutingBuilder

       Each Builder accepts a pandas DataFrame directly, so the caller is free
       to obtain that DataFrame however it likes (SQLAlchemy, mysql.connector, CSV, …).

       builder = FieldDictionaryBuilder(df)   # pass DataFrame
       data    = builder.build()              # returns the dict ready for json.dump
"""

from __future__ import annotations
import json
import logging
import os
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

import pandas as pd

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# Shared helper
# ─────────────────────────────────────────────────────────────────────────────

def _normalize(name: str) -> str:
    name = str(name).lower().strip()
    return re.sub(r"[\_\s.\-]", "", name)


# ─────────────────────────────────────────────────────────────────────────────
# FieldDictionaryBuilder  (from putm data)
# ─────────────────────────────────────────────────────────────────────────────

class FieldDictionaryBuilder:
    """
    Builds field_dictionary.json from a putm_upload_api_excel_json_mapping DataFrame.

    Usage:
        builder = FieldDictionaryBuilder(df)
        data    = builder.build()   # → dict
    """

    def __init__(self, df: pd.DataFrame):
        self._df = df

    # keep the old .load() / path-based interface so existing callers don't break
    @classmethod
    def from_xlsx(cls, path: Path) -> "FieldDictionaryBuilder":
        df = pd.read_excel(path)
        return cls(df)

    def load(self):
        """No-op when DataFrame is supplied directly; kept for API compatibility."""
        pass

    def build(self) -> Dict[str, Any]:
        df = self._df
        logger.info("Building field_dictionary …")

        col_excel   = (
            next((c for c in df.columns if "excel" in c.lower() and "key" in c.lower()), None)
            or next((c for c in df.columns if "excel" in c.lower()), None)
        )
        col_json    = (
            next((c for c in df.columns if "json" in c.lower() and "key" in c.lower()), None)
            or next((c for c in df.columns if "json" in c.lower()), None)
        )
        col_role    = (
            next((c for c in df.columns if "role" in c.lower()), None)
            or next((c for c in df.columns if "process" in c.lower()), None)
        )
        col_desc    = next((c for c in df.columns if "desc" in c.lower()), None)
        col_example = next((c for c in df.columns if "example" in c.lower()), None)

        logger.debug(f"field_dict cols — excel:{col_excel}  json:{col_json}  role:{col_role}")

        by_role: Dict[str, list] = defaultdict(list)
        by_excel_key: Dict[str, Any] = {}
        all_excel_keys: list = []
        seen: set = set()

        for _, row in df.iterrows():
            excel_key = str(row[col_excel]).strip()                         if col_excel  else ""
            json_key  = str(row[col_json]).strip()                          if col_json   else ""
            role      = str(row[col_role]).strip().upper()                  if col_role   else "LOAN"
            desc      = (
                str(row[col_desc]).strip()
                if col_desc and pd.notna(row.get(col_desc)) else None
            )
            example   = (
                str(row[col_example]).strip()
                if col_example and pd.notna(row.get(col_example)) else None
            )

            if not excel_key or not json_key or excel_key == "nan" or json_key == "nan":
                continue

            key = (excel_key, json_key, role)
            if key in seen:
                continue
            seen.add(key)

            by_role[role].append(
                {"excel_key": excel_key, "json_key": json_key,
                 "description": desc, "example": example}
            )

            if excel_key not in by_excel_key:
                by_excel_key[excel_key] = {
                    "json_key": json_key, "role": role,
                    "description": desc, "example": example,
                }
                all_excel_keys.append(excel_key)

        by_role = {k: sorted(v, key=lambda x: x["excel_key"]) for k, v in sorted(by_role.items())}
        all_excel_keys.sort()

        logger.info(f"field_dictionary: {len(seen)} entries, {len(by_role)} roles")
        return {
            "by_role":        dict(by_role),
            "by_excel_key":   by_excel_key,
            "all_excel_keys": all_excel_keys,
            "metadata": {
                "total":          len(seen),
                "by_role_count":  {k: len(v) for k, v in by_role.items()},
                "generated_at":   datetime.utcnow().isoformat() + "Z",
                "source":         "putm_upload_api_excel_json_mapping",
            },
        }


# ─────────────────────────────────────────────────────────────────────────────
# AliasRegistryBuilder  (from generic_excel data)
# ─────────────────────────────────────────────────────────────────────────────

class AliasRegistryBuilder:
    """
    Builds alias_registry.json from a generic_excel_upload_definition_fields DataFrame.

    Usage:
        builder = AliasRegistryBuilder(df)
        data    = builder.build()   # → dict
    """

    def __init__(self, df: pd.DataFrame):
        self._df = df

    @classmethod
    def from_csv(cls, path: Path) -> "AliasRegistryBuilder":
        df = pd.read_csv(path, encoding="utf-8")
        return cls(df)

    def load(self):
        """No-op when DataFrame supplied directly."""
        pass

    # ── private helpers ──────────────────────────────────────────────────────

    @staticmethod
    def _coerce_scalar(value):
        if isinstance(value, pd.Series):
            value = value.dropna()
            return value.iloc[0] if not value.empty else None
        return value

    @classmethod
    def _optional_text(cls, value) -> Optional[str]:
        value = cls._coerce_scalar(value)
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return None
        text = str(value).strip()
        return None if text.lower() in ("", "none", "nan", "null") else text

    # ── build ────────────────────────────────────────────────────────────────

    def build(self) -> Dict[str, Any]:
        df = self._df
        logger.info("Building alias_registry …")

        col_excel  = (
            next((c for c in df.columns if "excel_column" in c.lower()), None)
            or next((c for c in df.columns if "excel" in c.lower()), None)
        )
        col_table  = (
            next((c for c in df.columns if "table_column" in c.lower()), None)
            or next((c for c in df.columns if "table" in c.lower() and "column" in c.lower()), None)
        )
        col_json   = (
            next((c for c in df.columns if "partner_api" in c.lower()), None)
            or next((c for c in df.columns if "api_key" in c.lower()), None)
        )
        col_group  = next((c for c in df.columns
                           if "ui_group" in c.lower() or "grouping" in c.lower()), None)
        col_master = (
            next((c for c in df.columns if c.lower() == "master_id"), None)
            or next((c for c in df.columns if "master" in c.lower()), None)
        )

        logger.debug(
            f"alias_registry cols — excel:{col_excel}  table:{col_table}  "
            f"json:{col_json}  group:{col_group}  master:{col_master}"
        )

        forward: Dict[str, Any] = {}
        tier_counts: Dict[str, int] = defaultdict(int)

        for _, row in df.iterrows():
            excel_col  = str(row[col_excel]).strip()                    if col_excel  else ""
            table_col  = str(row[col_table]).strip()                    if col_table  else ""
            json_key   = self._optional_text(row[col_json])             if col_json   else None
            ui_group   = self._optional_text(row[col_group])            if col_group  else None
            master_id  = self._coerce_scalar(row[col_master])           if col_master else None

            if not excel_col or not table_col or excel_col == "nan" or table_col == "nan":
                continue
            if not json_key:
                continue
            if master_id is None or (isinstance(master_id, float) and pd.isna(master_id)):
                continue

            norm = _normalize(excel_col)
            if norm not in forward:
                forward[norm] = {
                    "target_excel_key": table_col,
                    "target_json_key":  json_key,
                    "frequency":        0,
                    "variants":         set(),
                    "master_ids":       set(),
                    "ui_groupings":     defaultdict(int),
                }

            entry = forward[norm]
            entry["frequency"] += 1
            entry["variants"].add(excel_col)
            entry["master_ids"].add(int(master_id))
            if ui_group:
                entry["ui_groupings"][ui_group] += 1

        # finalise forward entries
        forward_final: Dict[str, Any] = {}
        for norm, entry in forward.items():
            n    = len(entry["master_ids"])
            tier = "tier1" if n >= 30 else "tier2" if n >= 10 else "tier3" if n >= 3 else "tier4"
            tier_counts[tier] += 1

            most_common_group = (
                max(entry["ui_groupings"].items(), key=lambda x: x[1])[0]
                if entry["ui_groupings"] else None
            )

            forward_final[norm] = {
                "target_excel_key": entry["target_excel_key"],
                "target_json_key":  entry["target_json_key"],
                "frequency":        entry["frequency"],
                "variants":         sorted(entry["variants"]),
                "entity_context":   most_common_group,
                "confidence_tier":  tier,
            }

        # reverse lookup
        reverse: Dict[str, Any] = {}
        for norm, entry in forward_final.items():
            tc = entry["target_excel_key"]
            if tc not in reverse:
                reverse[tc] = {"partner_variants": set(), "frequency": 0}
            for v in entry["variants"]:
                reverse[tc]["partner_variants"].add(v)
            reverse[tc]["frequency"] += entry["frequency"]

        reverse_final = {
            tc: {"partner_variants": sorted(d["partner_variants"]), "frequency": d["frequency"]}
            for tc, d in sorted(reverse.items())
        }

        # count distinct partners
        all_partners: set = set()
        if col_master:
            master_col = df[col_master]
            if isinstance(master_col, pd.DataFrame):
                master_col = master_col.iloc[:, 0]
            for v in master_col.dropna():
                all_partners.add(int(v))

        logger.info(
            f"alias_registry: {len(forward_final)} forward, {len(reverse_final)} reverse mappings"
        )
        return {
            "forward":     forward_final,
            "reverse":     reverse_final,
            "tier_counts": dict(tier_counts),
            "metadata": {
                "total_mappings": sum(e["frequency"] for e in forward_final.values()),
                "total_partners": len(all_partners),
                "generated_at":   datetime.utcnow().isoformat() + "Z",
                "source":         "generic_excel_upload_definition_fields",
            },
        }


# ─────────────────────────────────────────────────────────────────────────────
# EntityRoutingBuilder  (from generic_excel ui_grouping)
# ─────────────────────────────────────────────────────────────────────────────

class EntityRoutingBuilder:
    """
    Builds entity_routing.json from a generic_excel_upload_definition_fields DataFrame.

    Usage:
        builder = EntityRoutingBuilder(df)
        data    = builder.build()   # → dict
    """

    _KEYWORD_RULES = [
        {"pattern": r"coapplicant\d+",                   "entity": "COAPPLICANT_NUM"},
        {"pattern": r"coapplicant",                      "entity": "COAPPLICANT"},
        {"pattern": r"guarantor",                        "entity": "GUARANTOR"},
        {"pattern": r"applicant",                        "entity": "APPLICANT"},
        {"pattern": r"loan|disburs|repay|emi|interest",  "entity": "LOAN"},
        {"pattern": r"document|attach|file|upload|proof","entity": "DOCUMENT"},
        {"pattern": r"fee|charge|penalt",                "entity": "FEE"},
        {"pattern": r"parameter|config|setting",         "entity": "PARAMETER"},
    ]

    def __init__(self, df: pd.DataFrame):
        self._df = df

    @classmethod
    def from_csv(cls, path: Path) -> "EntityRoutingBuilder":
        df = pd.read_csv(path, encoding="utf-8")
        return cls(df)

    def load(self):
        """No-op when DataFrame supplied directly."""
        pass

    def build(self) -> Dict[str, Any]:
        df = self._df
        logger.info("Building entity_routing …")

        col_group = next(
            (c for c in df.columns if "ui_group" in c.lower() or "grouping" in c.lower()), None
        )

        grouping_to_entity: Dict[str, str] = {}
        if col_group:
            for grouping in df[col_group].dropna().unique():
                g = str(grouping).strip()
                if not g or g == "nan":
                    continue
                norm   = g.lower().replace("_", "").replace(" ", "")
                entity = "OTHER"
                for rule in self._KEYWORD_RULES:
                    if re.search(rule["pattern"], norm):
                        entity = rule["entity"]
                        if entity == "COAPPLICANT_NUM":
                            m = re.search(r"coapplicant(\d+)", norm)
                            entity = f"COAPPLICANT{m.group(1)}" if m else "COAPPLICANT"
                        break
                grouping_to_entity[g] = entity

        logger.info(f"entity_routing: {len(grouping_to_entity)} groupings")
        return {
            "grouping_to_entity": grouping_to_entity,
            "keyword_rules": [
                {"pattern": r["pattern"], "entity": r["entity"]}
                for r in self._KEYWORD_RULES
            ],
            "metadata": {
                "total_groupings": len(grouping_to_entity),
                "generated_at":    datetime.utcnow().isoformat() + "Z",
                "source":          "generic_excel_upload_definition_fields (ui_grouping)",
            },
        }


# ─────────────────────────────────────────────────────────────────────────────
# Standalone CLI (original mysql.connector path — unchanged)
# ─────────────────────────────────────────────────────────────────────────────

def _cli_get_connection():
    """Open a direct mysql.connector connection from .env (CLI usage only)."""
    try:
        from dotenv import load_dotenv
        env_path = Path(__file__).parent.parent / ".env"
        if not env_path.exists():
            env_path = Path(__file__).parent / ".env"
        load_dotenv(dotenv_path=env_path)
    except ImportError:
        pass  # env vars already set in the shell

    try:
        import mysql.connector
    except ImportError:
        print("ERROR: pip install mysql-connector-python")
        sys.exit(1)

    host     = os.getenv("DB_HOST")
    port     = int(os.getenv("DB_PORT", "3306"))
    database = os.getenv("DB_NAME") or os.getenv("DB_DATABASE")
    user     = os.getenv("DB_USER")
    password = os.getenv("DB_PASSWORD")

    if not all([host, database, user, password]):
        missing = [k for k, v in {
            "DB_HOST": host, "DB_NAME/DB_DATABASE": database,
            "DB_USER": user, "DB_PASSWORD": password,
        }.items() if not v]
        print(f"Missing .env variables: {missing}")
        sys.exit(1)

    print(f"Connecting to MySQL at {host}:{port}/{database} …")
    conn = mysql.connector.connect(
        host=host, port=port, database=database,
        user=user, password=password, connection_timeout=30,
    )
    print("Connected.")
    return conn


def main() -> int:
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s - %(levelname)s - %(message)s")
    conn = _cli_get_connection()
    try:
        generic_df = pd.read_sql("""
            SELECT
                geudf.*,
                geud.upload_name,
                geud.id AS upload_definition_id
            FROM financialForms.generic_excel_upload_definition_fields geudf
            INNER JOIN financialForms.generic_excel_upload_definition geud
                ON geud.id = geudf.master_id
               AND geud.upload_name = 'Individual Loan Upload v3'
        """, conn)

        putm_df = pd.read_sql("""
            SELECT * FROM financialForms.putm_upload_api_excel_json_mapping
            WHERE process_name IN ('Origination', 'Enrollment')
        """, conn)
    finally:
        conn.close()

    output_dir = Path(__file__).parent.parent / "references"
    output_dir.mkdir(exist_ok=True)

    try:
        # field_dictionary.json
        fd = FieldDictionaryBuilder(putm_df).build()
        out = output_dir / "field_dictionary.json"
        out.write_text(json.dumps(fd, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"✓ {out}  ({fd['metadata']['total']} entries)")

        # alias_registry.json
        ar = AliasRegistryBuilder(generic_df).build()
        out = output_dir / "alias_registry.json"
        out.write_text(json.dumps(ar, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"✓ {out}  ({len(ar['forward'])} forward mappings)")

        # entity_routing.json
        er = EntityRoutingBuilder(generic_df).build()
        out = output_dir / "entity_routing.json"
        out.write_text(json.dumps(er, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"✓ {out}  ({er['metadata']['total_groupings']} groupings)")

        print("\n=== Build Complete — references/ folder is ready ===")
        return 0
    except Exception as e:
        logger.error(f"Build failed: {e}", exc_info=True)
        return 1


if __name__ == "__main__":
    sys.exit(main())