#!/usr/bin/env python3
"""
excel_compare.py

Compare values between two Excel sheets using a mapping configuration Excel.

Usage:
    python excel_compare.py \
      --file-a fileA.xlsx --file-b fileB.xlsx \
      --mapping mapping.xlsx \
      --sheet-a Sheet1 --sheet-b Sheet1 \
      --out report.xlsx

Mapping file expectations (Excel):
- Sheet named "mappings" (or first sheet):
    Columns:
      sheet_a_col   sheet_b_col   comparator      tolerance
    Example rows:
      ID            ID            exact
      Amount        Amt           numeric         0.01
      Name          FullName      case_insensitive

- Optional sheet named "keys":
    Columns:
      key_a         key_b
    Example:
      ID            ID

If "keys" is missing, script will compare by row index (position).
"""

import argparse
import logging
import math
import sys
from typing import List, Tuple, Dict, Any, Optional

import numpy as np
import pandas as pd

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger("excel_compare")


def read_sheet(path: str, sheet: Optional[str] = None) -> pd.DataFrame:
    if sheet:
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    else:
        # first sheet
        df = pd.read_excel(path, sheet_name=0, engine="openpyxl")
    return df


def load_mapping(mapping_path: Optional[str]) -> Tuple[List[Dict[str, Any]], List[Tuple[str, str]]]:
    """
    Returns:
      mappings: list of dicts with keys: sheet_a_col, sheet_b_col, comparator, tolerance
      keys: list of tuples (key_a, key_b) for row alignment (may be empty)
    """
    mappings = []
    keys = []

    if not mapping_path:
        return mappings, keys

    xls = pd.ExcelFile(mapping_path, engine="openpyxl")
    # Find mappings sheet: try name "mappings" else first sheet
    sheet_names = xls.sheet_names
    map_sheet = "mappings" if "mappings" in sheet_names else sheet_names[0]
    df_map = pd.read_excel(mapping_path, sheet_name=map_sheet, engine="openpyxl")

    # Normalize column names lower-case for flexible input
    colmap = {c.lower(): c for c in df_map.columns}

    # Required columns: sheet_a_col and sheet_b_col (names may vary)
    # Try to find reasonable names
    def find_col(possible):
        for p in possible:
            if p in colmap:
                return colmap[p]
        return None

    a_colname = find_col(["sheet_a_col", "a_col", "col_a", "left_col", "left"])
    b_colname = find_col(["sheet_b_col", "b_col", "col_b", "right_col", "right"])
    comp_colname = find_col(["comparator", "compare", "comparison"])
    tol_colname = find_col(["tolerance", "tol", "t"])

    if not a_colname or not b_colname:
        raise ValueError("Mapping sheet must contain columns for sheet_a_col and sheet_b_col (or variants).")

    for _, r in df_map.iterrows():
        mappings.append({
            "sheet_a_col": r.get(a_colname),
            "sheet_b_col": r.get(b_colname),
            "comparator": str(r.get(comp_colname)).strip().lower() if comp_colname and pd.notna(r.get(comp_colname)) else "exact",
            "tolerance": float(r.get(tol_colname)) if tol_colname and pd.notna(r.get(tol_colname)) else None
        })

    # Try to read keys if present
    if "keys" in sheet_names:
        df_keys = pd.read_excel(mapping_path, sheet_name="keys", engine="openpyxl")
        colmap_k = {c.lower(): c for c in df_keys.columns}
        key_a_col = None
        key_b_col = None
        for k in ["key_a", "keya", "keyin_a", "left_key", "key_a_col"]:
            if k in colmap_k:
                key_a_col = colmap_k[k]
                break
        for k in ["key_b", "keyb", "keyin_b", "right_key", "key_b_col"]:
            if k in colmap_k:
                key_b_col = colmap_k[k]
                break

        if key_a_col and key_b_col:
            for _, r in df_keys.iterrows():
                ka = r.get(key_a_col)
                kb = r.get(key_b_col)
                if pd.notna(ka) and pd.notna(kb):
                    keys.append((str(ka), str(kb)))
        else:
            # fallback: if keys sheet exists but columns not detected, try first two columns
            if df_keys.shape[1] >= 2:
                for _, r in df_keys.iterrows():
                    ka = r.iloc[0]
                    kb = r.iloc[1]
                    if pd.notna(ka) and pd.notna(kb):
                        keys.append((str(ka), str(kb)))
    return mappings, keys


def align_dataframes(dfA: pd.DataFrame, dfB: pd.DataFrame, keys: List[Tuple[str, str]]) -> pd.DataFrame:
    """
    Returns a merged DataFrame with suffixes _A and _B and an indicator __merge.
    Adds __row_A and __row_B columns with original row indexes to help reporting.
    """
    dfA2 = dfA.reset_index(drop=False).rename(columns={"index": "__row_A"})
    dfB2 = dfB.reset_index(drop=False).rename(columns={"index": "__row_B"})

    if keys:
        left_on = [ka for ka, _ in keys]
        right_on = [kb for _, kb in keys]
        # If some keys don't exist in either df, raise for clarity
        missing_left = [k for k in left_on if k not in dfA2.columns]
        missing_right = [k for k in right_on if k not in dfB2.columns]
        if missing_left:
            raise KeyError(f"Key columns not found in file A: {missing_left}")
        if missing_right:
            raise KeyError(f"Key columns not found in file B: {missing_right}")

        merged = pd.merge(dfA2, dfB2, how="outer", left_on=left_on, right_on=right_on,
                          suffixes=("_A", "_B"), indicator="__merge")
    else:
        # align by row position: create a synthetic index column
        dfA2["__pos"] = np.arange(len(dfA2))
        dfB2["__pos"] = np.arange(len(dfB2))
        merged = pd.merge(dfA2, dfB2, how="outer", left_on="__pos", right_on="__pos",
                          suffixes=("_A", "_B"), indicator="__merge")
        # keep __pos for reference
    return merged


def is_both_na(a, b):
    return (pd.isna(a) and pd.isna(b))


def compare_value(a, b, comparator: str = "exact", tolerance: Optional[float] = None) -> Tuple[bool, Optional[Any], Optional[str]]:
    """
    Returns (match_bool, difference, reason)
    difference: numeric difference if numeric comparator, else None or string
    reason: optional explanation for mismatch (e.g., 'both missing', 'col missing', 'type coercion')
    """
    # treat NaNs: if both are NA -> match
    if is_both_na(a, b):
        return True, None, "both_missing"

    # Missing one side
    if pd.isna(a) and not pd.isna(b):
        return False, None, "missing_in_A"
    if pd.isna(b) and not pd.isna(a):
        return False, None, "missing_in_B"

    comp = comparator.lower() if comparator else "exact"

    if comp == "exact":
        # simple equality (will work for ints, floats if same, strings exact)
        try:
            equal = a == b
            # If pandas objects, equality may return array; convert to bool safely
            if isinstance(equal, (pd.Series, pd.DataFrame, np.ndarray)):
                equal = bool(equal)
            return bool(equal), None, None if equal else "not_equal"
        except Exception:
            # fallback to string compare
            equal = str(a) == str(b)
            return equal, None, None if equal else "not_equal"

    if comp == "case_insensitive" or comp == "case-insensitive":
        try:
            sa = "" if pd.isna(a) else str(a).lower()
            sb = "" if pd.isna(b) else str(b).lower()
            ok = sa == sb
            return ok, None, None if ok else "case_mismatch"
        except Exception:
            ok = str(a).lower() == str(b).lower()
            return ok, None, None if ok else "case_mismatch"

    if comp == "numeric":
        # attempt numeric coercion
        a_num = pd.to_numeric(a, errors="coerce")
        b_num = pd.to_numeric(b, errors="coerce")
        if pd.isna(a_num) or pd.isna(b_num):
            # fallback: not numeric
            ok = False
            return ok, None, "not_numeric"
        diff = b_num - a_num
        tol = float(tolerance) if tolerance is not None else 0.0
        ok = abs(diff) <= tol
        return bool(ok), float(diff), None if ok else f"diff_exceeds_tol ({diff})"

    # Unknown comparator: fallback to string equality
    try:
        ok = str(a) == str(b)
        return ok, None, None if ok else "not_equal_string_fallback"
    except Exception:
        return False, None, "compare_error"


def generate_report(merged: pd.DataFrame, mappings: List[Dict[str, Any]], keys: List[Tuple[str, str]],
                    out_path: str):
    """
    Iterate over merged rows and mapping pairs to produce details and summary.
    """
    detail_rows = []

    # For keys, prepare display key value per merged row
    key_display_names = []
    if keys:
        # use left key names for display (if same name maybe repeated; show pairs)
        key_display_names = [ka for ka, _ in keys]
    else:
        # use synthetic position or row ids
        key_display_names = ["__pos"]

    # For every merged row (each aligned pair), compare each mapping pair
    total = 0
    matches = 0
    mismatches = 0
    missing = 0

    # To help error messages, check available columns in merged
    merged_cols = set(merged.columns)

    for idx, row in merged.iterrows():
        # build key value dict for reporting
        key_vals = {}
        for i, (ka_kb) in enumerate(keys) if keys else enumerate([("__pos", "__pos")]):
            if keys:
                ka, kb = ka_kb
                # prefer left value then right if left missing
                val = row.get(ka) if ka in merged_cols else None
                # If left is nano, maybe right exists under ka + "_B"? But merge puts right keys into columns named as right key names
                if pd.isna(val) and kb in merged_cols:
                    val = row.get(kb)
                key_vals[f"key_{i}"] = val
            else:
                # pos column exists maybe
                key_vals["key_0"] = row.get("__pos", None)

        # record original row indexes
        rowA_idx = row.get("__row_A", np.nan)
        rowB_idx = row.get("__row_B", np.nan)
        merge_indicator = row.get("__merge", "")

        for m in mappings:
            total += 1
            a_col = m["sheet_a_col"]
            b_col = m["sheet_b_col"]
            comparator = m.get("comparator", "exact")
            tol = m.get("tolerance", None)

            # Determine actual column names in merged DF. After merge, columns from A are named e.g. "Amount" (if not overlapping),
            # columns from B are also named "Amount" if different keys? Because merge with suffixes only affects identical column names.
            # To ensure we pick A vs B, check presence:
            # pandas merge left columns keep their names; right columns that overlap keep suffix "_B".
            # For robust approach: prioritize these options:
            # 1. Column present as '<col>' from A side (we can detect by whether '<col>' exists AND '<col>_B' exists)
            # 2. Column present as '<col>_A' OR '<col>_B' if suffixing occurred.
            # We'll attempt to find the A-value column name and B-value column name in merged DF.
            def find_col_name(base_col: str, side: str) -> Optional[str]:
                # side: 'A' or 'B'
                if base_col in merged_cols:
                    # If both exist as base_col and base_col_B, we need to detect which is A vs B.
                    # After merge, if both frames had the same column name, pandas will rename them to <col>_x and <col>_y only if they were in both inputs but not keys.
                    # To be conservative, prefer base_col for A if side == 'A' and base_col + '_B' exists (then base_col is A)
                    if side == 'A':
                        # prefer column as-is if it came from A
                        return base_col
                    else:
                        # side B prefer base_col_B if exists
                        if f"{base_col}_B" in merged_cols:
                            return f"{base_col}_B"
                        if f"{base_col}_A" in merged_cols:
                            # ambiguous but choose A suffixed column as fallback
                            return f"{base_col}_A"
                        return base_col
                else:
                    # maybe suffixed columns exist
                    cand = f"{base_col}_A" if side == 'A' else f"{base_col}_B"
                    if cand in merged_cols:
                        return cand
                    # try opposite suffix (in case mapping used different suffixing convention)
                    alt = f"{base_col}_B" if side == 'A' else f"{base_col}_A"
                    if alt in merged_cols:
                        return alt
                    # Lastly check direct presence of the other column name (if mapping used different header name)
                    return None

            a_col_name = find_col_name(a_col, "A")
            b_col_name = find_col_name(b_col, "B")

            valA = np.nan
            valB = np.nan
            reason = None

            if a_col_name and a_col_name in merged_cols:
                valA = row.get(a_col_name)
            else:
                # try original name with suffix _A
                if f"{a_col}_A" in merged_cols:
                    valA = row.get(f"{a_col}_A")
                else:
                    valA = np.nan
                    reason = "col_missing_in_A"

            if b_col_name and b_col_name in merged_cols:
                valB = row.get(b_col_name)
            else:
                if f"{b_col}_B" in merged_cols:
                    valB = row.get(f"{b_col}_B")
                else:
                    valB = np.nan
                    if reason:
                        reason += "; col_missing_in_B"
                    else:
                        reason = "col_missing_in_B"

            # If row is missing entirely in one file, we should mark missing
            if merge_indicator == "left_only":
                # present in A only
                cmp_result = False
                mismatches += 1
                missing += 1
                detail_rows.append({
                    **key_vals,
                    "rowA": int(rowA_idx) if not pd.isna(rowA_idx) else None,
                    "rowB": None,
                    "colA": a_col,
                    "colB": b_col,
                    "valueA": valA,
                    "valueB": None,
                    "comparator": comparator,
                    "tolerance": tol,
                    "result": "missing_in_B",
                    "difference": None,
                    "reason": "row_missing_in_B"
                })
                continue
            if merge_indicator == "right_only":
                cmp_result = False
                mismatches += 1
                missing += 1
                detail_rows.append({
                    **key_vals,
                    "rowA": None,
                    "rowB": int(rowB_idx) if not pd.isna(rowB_idx) else None,
                    "colA": a_col,
                    "colB": b_col,
                    "valueA": None,
                    "valueB": valB,
                    "comparator": comparator,
                    "tolerance": tol,
                    "result": "missing_in_A",
                    "difference": None,
                    "reason": "row_missing_in_A"
                })
                continue

            # Now perform the comparison
            try:
                ok, diff, reason_cmp = compare_value(valA, valB, comparator=comparator, tolerance=tol)
            except Exception as e:
                ok = False
                diff = None
                reason_cmp = f"compare_error: {e}"

            if ok:
                matches += 1
                detail_rows.append({
                    **key_vals,
                    "rowA": int(rowA_idx) if not pd.isna(rowA_idx) else None,
                    "rowB": int(rowB_idx) if not pd.isna(rowB_idx) else None,
                    "colA": a_col,
                    "colB": b_col,
                    "valueA": valA,
                    "valueB": valB,
                    "comparator": comparator,
                    "tolerance": tol,
                    "result": "match",
                    "difference": diff,
                    "reason": reason_cmp
                })
            else:
                mismatches += 1
                detail_rows.append({
                    **key_vals,
                    "rowA": int(rowA_idx) if not pd.isna(rowA_idx) else None,
                    "rowB": int(rowB_idx) if not pd.isna(rowB_idx) else None,
                    "colA": a_col,
                    "colB": b_col,
                    "valueA": valA,
                    "valueB": valB,
                    "comparator": comparator,
                    "tolerance": tol,
                    "result": "mismatch",
                    "difference": diff,
                    "reason": reason_cmp
                })

    # Build DataFrames
    details_df = pd.DataFrame(detail_rows)
    summary = {
        "total_compared": total,
        "matches": matches,
        "mismatches": mismatches,
        "missing_rows_or_cols": missing,
        "match_percent": (matches / total * 100) if total else 0.0,
        "mismatch_percent": (mismatches / total * 100) if total else 0.0,
    }
    summary_df = pd.DataFrame.from_dict([summary])

    # Write to Excel
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="summary", index=False)
        details_df.to_excel(writer, sheet_name="details", index=False)
    logger.info("Report written to: %s", out_path)
    return summary_df, details_df


def parse_args():
    p = argparse.ArgumentParser(description="Compare two Excel sheets using a mapping config.")
    p.add_argument("--file-a", required=True, help="Path to first Excel file (A).")
    p.add_argument("--file-b", required=True, help="Path to second Excel file (B).")
    p.add_argument("--mapping", required=False, default=None, help="Path to mapping Excel file (mappings + optional keys sheets).")
    p.add_argument("--sheet-a", default=None, help="Sheet name in file A (defaults to first sheet).")
    p.add_argument("--sheet-b", default=None, help="Sheet name in file B (defaults to first sheet).")
    p.add_argument("--out", default="compare_report.xlsx", help="Output report Excel file.")
    p.add_argument("--debug", action="store_true", help="Enable debug logging.")
    return p.parse_args()


def main():
    args = parse_args()
    if args.debug:
        logger.setLevel(logging.DEBUG)

    try:
        logger.info("Reading file A: %s (sheet: %s)", args.file_a, args.sheet_a or "<first>")
        dfA = read_sheet(args.file_a, args.sheet_a)
        logger.info("File A rows: %d, cols: %d", dfA.shape[0], dfA.shape[1])
    except Exception as e:
        logger.error("Failed reading file A: %s", e)
        sys.exit(2)
    try:
        logger.info("Reading file B: %s (sheet: %s)", args.file_b, args.sheet_b or "<first>")
        dfB = read_sheet(args.file_b, args.sheet_b)
        logger.info("File B rows: %d, cols: %d", dfB.shape[0], dfB.shape[1])
    except Exception as e:
        logger.error("Failed reading file B: %s", e)
        sys.exit(2)

    try:
        mappings, keys = load_mapping(args.mapping)
        if not mappings:
            # Default mapping: compare common columns by identical name
            common = [c for c in dfA.columns if c in dfB.columns]
            mappings = [{"sheet_a_col": c, "sheet_b_col": c, "comparator": "exact", "tolerance": None} for c in common]
            logger.info("No mapping provided; defaulting to compare %d common columns by exact match.", len(common))
        else:
            logger.info("Loaded %d mapping rows; keys: %s", len(mappings), keys)
    except Exception as e:
        logger.error("Failed loading mapping: %s", e)
        sys.exit(3)

    try:
        merged = align_dataframes(dfA, dfB, keys)
        logger.info("Aligned rows (merged shape): %s", merged.shape)
    except Exception as e:
        logger.error("Failed aligning dataframes: %s", e)
        sys.exit(4)

    try:
        summary_df, details_df = generate_report(merged, mappings, keys, args.out)
        logger.info("Summary: \n%s", summary_df.to_string(index=False))
    except Exception as e:
        logger.error("Failed generating report: %s", e)
        sys.exit(5)


if __name__ == "__main__":
    main()