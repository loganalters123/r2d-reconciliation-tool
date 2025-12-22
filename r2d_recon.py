#!/usr/bin/env python3
"""
Repayments to Date Reconciliation Tool

Reconciles repayment data against bank Chase transactions, matching credits and debits
through multiple strategies including amount+date matching, note parsing, and ReconTags.

Usage:
    python3 r2d_recon.py --file <input.xlsx> [--out <output.xlsx>]
"""
import argparse
import logging
import os
import re
import sys
from datetime import date
from pathlib import Path
import pandas as pd

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(levelname)s: %(message)s'
)
logger = logging.getLogger(__name__)

# Note: ACH ID conflicts are now passed through function returns instead of globals

# ------------------------- Parameters & Regex -------------------------

# Matching windows
DATE_WINDOW_DAYS = 10  # Days tolerance for matching transaction dates (increased from 5)
DATE_WINDOW_DAYS_WIDE = 30  # Wider window for second-pass matching
OVERPAY_BACKFILL_WINDOW = 7  # Days to look back for overpayment debits
NOTE_WINDOW_DAYS = 7  # Days tolerance for note-based matching
NOTE_WINDOW_DAYS_EXTENDED = 3  # Additional days for fallback note matching

# Amount matching
AMOUNT_TOL = 0.02  # Amount matching tolerance in dollars (2 cents to account for floating point errors)

# Confidence scoring
MAX_CONFIDENCE = 0.99  # Maximum confidence score cap for matches

# Overpayment detection
OVERPAY_DESC_PATTERN = "Online Transfer"  # Description pattern for overpayment debits

TRANSFER_HINTS = re.compile(r"(?:dwolla|transfer|ach|orig co name|orig id|trn)", re.I)
OVERPAY_DEBIT_HINTS = re.compile(r"(?:2670|transfer)", re.I)
# liberal: "overpaid by $X" OR "overpayment of $X" OR "overpayment $X"
OVERPAID_REGEX = re.compile(r"(?:overpaid\s*(?:by)?|overpayment\s*(?:of)?)\s*\$?\s*([0-9][0-9,]*\.?[0-9]{0,2})", re.I)
PAREN_SUFFIX = re.compile(r"\s*\([^)]*\)\s*$")
DOLLAR_REGEX = re.compile(r"\$?\s*([0-9][0-9,]*\.\d{2})")
DATE_IN_NOTES = re.compile(r"\b(\d{1,2})/(\d{1,2})\b")
CREDIT_KEYWORDS = re.compile(r"(received|deposit|check|credited|incoming|rec\.?\s*rem|received\s+rem|received\s+remaining|rcvd|remaining\s+repayment|repayment\s+received|remaining\s*bal|rem\.?\s*bal|underpaid\s+by)", re.I)
DEBIT_KEYWORDS  = re.compile(r"(send\s+funder|to\s+funder|transfer|outgoing|ach\s*out|2670)", re.I)
SEND_FUNDER_REGEX = re.compile(r"(?:to\s+)?send\s+funder[^$]*\$([0-9][0-9,]*\.[0-9]{2})", re.I)
RECEIVED_CHECK_REGEX = re.compile(r"(received.*?check|rec\.?\s*rem|received\s+rem|received\s+remaining|remaining\s+repayment|repayment\s+received|underpaid\s+by|received\s+for).*?\$([0-9][0-9,]*\.[0-9]{2})", re.I)
REQUESTED_REMAINING_REGEX = re.compile(r"(req\.?\s*rem\.?|requested\s+rem\.?|req\.?\s*remaining|requested\s+remaining|requesting\s+remaining|requesting\s+rem|underpayment\s+of).*?\$([0-9][0-9,]*\.[0-9]{2})", re.I)

# Shared check detection patterns
SHARED_CHECK_PATTERNS = re.compile(r"(?:check.*addressed\s+for\s+(\d+)\s+clients?|(\d+)\s+clients?|other\s+client\s+is\s+([^,)]+))", re.I)
OTHER_CLIENT_REGEX = re.compile(r"other\s+client\s+is\s+([^,)]+)", re.I)
CLIENT_COUNT_REGEX = re.compile(r"(?:check.*addressed\s+for\s+|for\s+)?(\d+)\s+clients?", re.I)

def r2(x, nd=2):
    try:
        return round(float(x), nd)
    except Exception:
        return None

# ------------------------- Loaders & Helpers -------------------------

def validate_file_exists(file_path):
    """
    Validate that the input file exists and is readable.

    Args:
        file_path: Path to the file to validate

    Raises:
        FileNotFoundError: If file doesn't exist
        PermissionError: If file isn't readable
    """
    path = Path(file_path)
    if not path.exists():
        logger.error(f"File not found: {file_path}")
        raise FileNotFoundError(f"Input file does not exist: {file_path}")
    if not path.is_file():
        logger.error(f"Path is not a file: {file_path}")
        raise ValueError(f"Path is not a file: {file_path}")
    if not os.access(file_path, os.R_OK):
        logger.error(f"File is not readable: {file_path}")
        raise PermissionError(f"File is not readable: {file_path}")
    logger.info(f"✓ File validated: {path.name}")

def validate_sheet_exists(file_path, sheet_name):
    """
    Validate that a sheet exists in the Excel file.

    Args:
        file_path: Path to the Excel file
        sheet_name: Name of the sheet to check

    Raises:
        ValueError: If sheet doesn't exist
    """
    try:
        xl_file = pd.ExcelFile(file_path)
        if sheet_name not in xl_file.sheet_names:
            logger.error(f"Sheet '{sheet_name}' not found in {Path(file_path).name}")
            logger.info(f"Available sheets: {', '.join(xl_file.sheet_names)}")
            raise ValueError(f"Sheet '{sheet_name}' not found. Available: {', '.join(xl_file.sheet_names)}")
        logger.info(f"✓ Sheet found: {sheet_name}")
    except Exception as e:
        if "not found" in str(e):
            raise
        logger.error(f"Error reading Excel file: {e}")
        raise ValueError(f"Cannot read Excel file {Path(file_path).name}: {e}")

def colmap(df, wanted_names, sheet_name="", required=None):
    """
    Map column names flexibly using aliases.

    Args:
        df: DataFrame to map columns for
        wanted_names: Dict of {key: [list of possible column names]}
        sheet_name: Name of sheet (for error messages)
        required: List of keys that are required (will raise error if missing)

    Returns:
        Dict mapping keys to actual column names (or None if not found)

    Raises:
        ValueError: If required columns are missing
    """
    renorm = {c: str(c).strip() for c in df.columns}
    df = df.rename(columns=renorm)
    low = {c.lower(): c for c in df.columns}
    out = {}
    missing_required = []

    for key, aliases in wanted_names.items():
        pick = None
        for a in aliases:
            if a and a.lower() in low:
                pick = low[a.lower()]
                break
        out[key] = pick

        # Check if required column is missing
        if required and key in required and pick is None:
            missing_required.append((key, aliases))

    if missing_required:
        sheet_msg = f" in sheet '{sheet_name}'" if sheet_name else ""
        error_msg = f"Missing required columns{sheet_msg}:\n"
        for key, aliases in missing_required:
            error_msg += f"  - {key}: Expected one of {aliases}\n"
        error_msg += f"\nAvailable columns: {list(df.columns)}"
        logger.error(error_msg)
        raise ValueError(error_msg)

    return out

def parse_dates(s):
    return pd.to_datetime(s, errors="coerce")

def normalize_amount(x):
    if pd.isna(x):
        return None
    try:
        return r2(str(x).replace(",", "").replace("$", ""))
    except Exception:
        return None

def load_r2d(path, sheet):
    """
    Load and normalize Repayments to Date sheet.

    Args:
        path: Path to Excel file
        sheet: Sheet name to load

    Returns:
        DataFrame with normalized repayment data

    Raises:
        ValueError: If required columns are missing
    """
    logger.info(f"Loading R2D data from sheet: {sheet}")
    df = pd.read_excel(path, sheet_name=sheet)
    logger.info(f"  Loaded {len(df)} rows")

    wanted = {
        "ach_id": ["ACH ID","ACHID","ACH_Id","ach id"],
        "amount_transferred": ["Amount Transferred","Transferred Amount","amount transferred"],
        "amount_to_funder": ["Amount To Funder","Amt To Funder","amount to funder"],
        "claim_id": ["Dynamo Claim ID","ClaimID","Claim Id","Dynamo Id","Dynamo"],
        "claimant": ["Recipient Name","Claimant Name","Recipient"],
        "deal_type": ["Deal Type","DealType","Type"],
        "contract_date": ["Contract Date","Date Funded","Transfer Initiated Date"],
        "transfer_initiated": ["Transfer Initiated Date","Transfer Initiated (ET)"],
        "likely_arrived": ["Likely Arrived Date","Date Closed"],
        "repayment_amount": ["Repayment Amount","Repayment Amount (KEEP)","Repayment"],
        "notes": ["Repayment Notes","Reconciliation Notes","Notes"],
        "legacy_id": ["Legacy ID","LegacyID","Legacy Id","Legacy","Correlation ID","CorrelationID"],
    }

    # Required columns for processing
    required = ["claim_id", "claimant"]
    m = colmap(df, wanted, sheet_name=sheet, required=required)
    out = pd.DataFrame({
        "ach_id": (df[m["ach_id"]].astype(str).str.strip() if m["ach_id"] else ""),
        "amount_transferred": (df[m["amount_transferred"]] if m["amount_transferred"] else None),
        "amount_to_funder": (df[m["amount_to_funder"]] if m["amount_to_funder"] else None),
        "claim_id": (df[m["claim_id"]] if m["claim_id"] else ""),
        "claimant": (df[m["claimant"]] if m["claimant"] else ""),
        "deal_type": (df[m["deal_type"]].astype(str).str.strip() if m["deal_type"] else ""),
        "contract_date": (parse_dates(df[m["contract_date"]]) if m["contract_date"] else pd.NaT),
        "transfer_initiated": (parse_dates(df[m["transfer_initiated"]]) if m["transfer_initiated"] else pd.NaT),
        "likely_arrived": (parse_dates(df[m["likely_arrived"]]) if m["likely_arrived"] else pd.NaT),
        "repayment_amount": (df[m["repayment_amount"]] if m["repayment_amount"] else None),
        "notes": (df[m["notes"]].astype(str) if m["notes"] else ""),
        "legacy_id": (df[m["legacy_id"]] if m["legacy_id"] else pd.NA),
    })
    out["amount_transferred"] = out["amount_transferred"].map(normalize_amount)
    out["amount_to_funder"] = out["amount_to_funder"].map(normalize_amount)
    out["repayment_amount"] = out["repayment_amount"].map(normalize_amount)
    out["window_date"] = out["likely_arrived"].fillna(out["transfer_initiated"])
    return out

def load_chase(path, sheet):
    """
    Load and normalize Chase bank transactions.

    Args:
        path: Path to Excel file
        sheet: Sheet name to load

    Returns:
        DataFrame with normalized Chase transaction data

    Raises:
        ValueError: If required columns are missing
    """
    logger.info(f"Loading Chase data from sheet: {sheet}")
    df = pd.read_excel(path, sheet_name=sheet)
    logger.info(f"  Loaded {len(df)} rows")

    wanted = {
        "posting_date":["Posting Date","Details Posting Date","Post Date"],
        "description":["Description","Details","Memo"],
        "amount":["Amount","Amt"],
        "type":["Type"],
        "recon_tag":["ReconTag","Recon Tag","Recon_Tag","RECONTAG","recontag"],
    }

    # Required columns for processing
    required = ["posting_date", "description", "amount"]
    m = colmap(df, wanted, sheet_name=sheet, required=required)
    out = pd.DataFrame({
        "posting_date": (parse_dates(df[m["posting_date"]]) if m["posting_date"] else pd.NaT),
        "description": (df[m["description"]].astype(str) if m["description"] else ""),
        "amount": (df[m["amount"]] if m["amount"] else None),
        "type": (df[m["type"]].astype(str) if m["type"] else ""),
    })
    out["amount"] = out["amount"].map(normalize_amount)
    out["is_debit"] = out["amount"].fillna(0) < 0
    out["is_credit"] = out["amount"].fillna(0) > 0
    out["has_hint"] = out["description"].str.contains(TRANSFER_HINTS, na=False, regex=True)
    out["overpay_hint"] = out["description"].str.contains(OVERPAY_DEBIT_HINTS, na=False, regex=True)

    # Normalize ReconTag robustly
    if m.get("recon_tag"):
        series = df[m["recon_tag"]].astype(object)
        def _norm(x):
            if x is None or (isinstance(x, float) and pd.isna(x)):
                return None
            s = str(x).strip()
            return None if s == "" or s.lower() in {"nan", "none"} else s
        out["recon_tag"] = series.map(_norm)
    else:
        out["recon_tag"] = pd.NA
    return out

def validate_ach_id_conflicts(r2d: pd.DataFrame):
    """Detect and report ACH ID conflicts before processing"""
    with_id = r2d[r2d["ach_id"].astype(str).str.len() > 0].copy()

    if with_id.empty:
        return []  # No conflicts if no ACH IDs

    conflicts = []

    # Group by ACH ID and check for conflicts
    for ach_id, group in with_id.groupby("ach_id"):
        if len(group) > 1:
            # Check if these are legitimate different claims
            unique_claim_ids = group["claim_id"].dropna().astype(str).str.strip()
            unique_claim_ids = unique_claim_ids[unique_claim_ids != ""].unique()

            unique_claimants = group["claimant"].dropna().astype(str).str.strip()
            unique_claimants = unique_claimants[unique_claimants != ""].unique()

            # Check for true conflicts: different people (not just AFR/Buyout variants)
            base_claimants = set()
            for claimant in unique_claimants:
                # Extract base name (remove AFR, AFR2, Buyout, etc. suffixes)
                base_name = re.sub(r'\s*\(?(AFR\d*|Buyout.*|BuyoutCA)\)?$', '', claimant).strip()
                base_claimants.add(base_name)

            # If multiple different base claim IDs or different people share same ACH ID = CONFLICT
            if len(unique_claim_ids) > 1 or len(base_claimants) > 1:
                conflict_details = []
                for _, row in group.iterrows():
                    conflict_details.append({
                        "claimant": row["claimant"],
                        "claim_id": row["claim_id"],
                        "amount": row["amount_transferred"],
                        "date": row["window_date"]
                    })

                conflicts.append({
                    "ach_id": ach_id,
                    "num_claims": len(unique_claim_ids) if len(unique_claim_ids) > 1 else len(unique_claimants),
                    "claims": conflict_details
                })

    return conflicts

def dedupe_by_ach_id(r2d: pd.DataFrame):
    """
    Enhanced deduplication with conflict detection and reporting.

    Args:
        r2d: DataFrame with repayment data

    Returns:
        tuple: (deduplicated_df, removed_count, conflicts)
            - deduplicated_df: DataFrame with duplicates removed
            - removed_count: Number of duplicate rows removed
            - conflicts: List of ACH ID conflicts for reporting
    """
    # First, check for ACH ID conflicts
    conflicts = validate_ach_id_conflicts(r2d)

    if conflicts:
        logger.warning("⚠️  DATA QUALITY ALERT: ACH ID CONFLICTS DETECTED!")
        logger.warning(f"Found {len(conflicts)} ACH ID conflicts - will be reported in Data_Quality_Issues tab")
        logger.warning("Processing will continue but please review source file for data quality issues")

    # Proceed with normal deduplication (only true duplicates)
    with_id = r2d[r2d["ach_id"].astype(str).str.len() > 0]
    without_id = r2d[r2d["ach_id"].astype(str).str.len() == 0]

    # Remove duplicates by ACH ID + claim ID (same transfer, different claimant name variations)
    # This handles cases like "Stacy Chambers", "Stacy Chambers (AFR)", etc. with same ACH ID
    before_count = len(with_id)
    kept = with_id.drop_duplicates(subset=["ach_id", "claim_id"], keep="first").copy()
    removed_count = before_count - len(kept)

    if removed_count > 0:
        logger.info(f"📋 Removed {removed_count} duplicate entries (same ACH ID + same claim)")

    return pd.concat([kept, without_id], ignore_index=True), removed_count, conflicts

# ------------------------- Correlation ID (Parent Legacy ID) -------------------------

def canonical_parent(row_group: pd.DataFrame) -> pd.Series:
    grp = row_group.copy()
    grp["deal_type"] = grp["deal_type"].astype(str)
    grp["is_afr"] = grp["deal_type"].str.contains(r"\bafr", case=False, na=False)
    candidates = grp[~grp["is_afr"]] if (~grp["is_afr"]).any() else grp
    if candidates["contract_date"].notna().any():
        parent = candidates.sort_values(["contract_date", "window_date"], na_position="last").iloc[0]
    else:
        parent = candidates.sort_values(["window_date"], na_position="last").iloc[0]
    parent = parent.copy()
    if isinstance(parent["claimant"], str):
        parent["claimant"] = PAREN_SUFFIX.sub("", parent["claimant"])
    return parent

def build_correlation_map(r2d: pd.DataFrame):
    parents = r2d.groupby("claim_id", dropna=False).apply(canonical_parent).reset_index(drop=True)
    corr = {}
    for _, p in parents.iterrows():
        claim = p.get("claim_id")
        lid = p.get("legacy_id")
        if pd.notna(lid) and str(lid).strip():
            corr[claim] = str(lid).strip()
        else:
            grp = r2d[r2d["claim_id"].astype(str) == str(claim)]
            any_lid = next((str(x).strip() for x in grp["legacy_id"] if pd.notna(x) and str(x).strip()), None)
            corr[claim] = any_lid
    return corr

def insert_corr(df, corr_map, claim_col="claim_id", pos=1):
    if isinstance(df, pd.DataFrame) and not df.empty and (claim_col in df.columns):
        df.insert(pos, "correlation_id", df[claim_col].map(corr_map))
    return df

# ------------------------- Debit Matching -------------------------

def match_debits_relaxed(r2d, chase):
    r2d_dedup, dup_removed, ach_conflicts = dedupe_by_ach_id(r2d)
    debits = chase[chase["is_debit"]].copy()

    # Global optimization: collect all possible claim-debit pairs, then assign by best match
    potential_matches = []

    for i, row in r2d_dedup.iterrows():
        amt = row.get("amount_transferred")
        win = row.get("window_date")
        if amt is None or pd.isna(amt) or pd.isna(win):
            continue

        amt_rounded = r2(amt)
        cand = debits[debits["amount"].abs().sub(abs(amt_rounded)).abs() <= AMOUNT_TOL].copy()
        cand = cand[(cand["posting_date"] >= win - pd.Timedelta(days=DATE_WINDOW_DAYS)) &
                    (cand["posting_date"] <= win + pd.Timedelta(days=DATE_WINDOW_DAYS))]

        for idx, c in cand.iterrows():
            date_delta = abs((c["posting_date"] - win).days)
            confidence = 0.5 + (0.3 if c["has_hint"] else 0) + (0.2 if date_delta <= 1 else 0)

            potential_matches.append({
                "r2d_index": i,
                "chase_index": idx,
                "ach_id": row.get("ach_id"),
                "claim_id": row.get("claim_id"),
                "claimant": row.get("claimant"),
                "amount_transferred": amt_rounded,
                "r2d_date": win,
                "chase_date": c["posting_date"],
                "chase_amount": c["amount"],
                "description": c["description"],
                "date_delta": date_delta,
                "has_hint": c["has_hint"],
                "confidence": r2(min(confidence, MAX_CONFIDENCE)),
                "match_type": "amount+window(+hints)"
            })

    # Sort by match quality (confidence desc, then date_delta asc)
    potential_matches.sort(key=lambda x: (-x["confidence"], x["date_delta"]))

    # Assign matches greedily by quality
    results = []
    used_claims, used_debits = set(), set()

    for match in potential_matches:
        r2d_idx = match["r2d_index"]
        chase_idx = match["chase_index"]

        if r2d_idx not in used_claims and chase_idx not in used_debits:
            used_claims.add(r2d_idx)
            used_debits.add(chase_idx)

            results.append({
                "ach_id": match["ach_id"],
                "claim_id": match["claim_id"],
                "amount_transferred": match["amount_transferred"],
                "r2d_date": match["r2d_date"],
                "chase_date": match["chase_date"],
                "chase_amount": match["chase_amount"],
                "description": match["description"],
                "match_type": match["match_type"],
                "confidence": match["confidence"],
                "chase_index": match["chase_index"],
            })

            if len(used_debits) > 0 and match["claimant"] == "Nina Brown":
                logger.debug(f"✅ Nina Brown matched with confidence {match['confidence']} (delta: {match['date_delta']} days)")

    logger.info(f"📊 Debit matching: {len(results)} matches from {len(potential_matches)} potential pairs")

    # Second pass: wider window for unmatched claims
    unmatched_idx = []
    for i, row in r2d_dedup.iterrows():
        if any(r["claim_id"] == row.get("claim_id") for r in results):  # Already matched
            continue

        amt = row.get("amount_transferred")
        win = row.get("window_date")
        if amt is None or pd.isna(amt) or pd.isna(win):
            unmatched_idx.append(i); continue

        amt_rounded = r2(amt)
        cand = debits[debits["amount"].abs().sub(abs(amt_rounded)).abs() <= AMOUNT_TOL].copy()
        cand = cand[(cand["posting_date"] >= win - pd.Timedelta(days=DATE_WINDOW_DAYS_WIDE)) &
                    (cand["posting_date"] <= win + pd.Timedelta(days=DATE_WINDOW_DAYS_WIDE))]
        if cand.empty:
            unmatched_idx.append(i); continue

        cand = cand.assign(date_delta=(cand["posting_date"]-win).abs().dt.days)
        cand = cand.sort_values(["has_hint","date_delta"], ascending=[False, True])
        chosen = None
        for idx, c in cand.iterrows():
            if idx not in used_debits:
                chosen = (idx, c); break
        if not chosen:
            unmatched_idx.append(i); continue

        idx, c = chosen; used_debits.add(idx)
        confidence = 0.5 + (0.3 if c["has_hint"] else 0) + (0.2 if abs((c["posting_date"]-win).days)<=1 else 0)
        results.append({
            "ach_id": row.get("ach_id"),
            "claim_id": row.get("claim_id"),
            "amount_transferred": amt_rounded,
            "r2d_date": win,
            "chase_date": c["posting_date"],
            "chase_amount": c["amount"],
            "description": c["description"],
            "match_type": "amount+extended_window",
            "confidence": r2(min(confidence, MAX_CONFIDENCE)),
            "chase_index": idx,
        })

    used_debit_idx = [r["chase_index"] for r in results]
    return pd.DataFrame(results), r2d_dedup.loc[unmatched_idx].copy(), debits.loc[~debits.index.isin(used_debit_idx)].copy(), dup_removed, used_debit_idx, ach_conflicts

# ------------------------- Credit Matching (Parent Claim) -------------------------

def _prepare_credit_matching_data(r2d):
    """
    Prepare and aggregate repayment data for credit matching.

    Aggregates by claim_id. Note: Claims with multiple transfers (different ACH IDs)
    will have their repayments summed. These should be reviewed in the source file.

    Args:
        r2d: DataFrame with repayment data

    Returns:
        tuple: (rolled_up_data, parent_claims)
    """
    tmp = r2d.copy()
    tmp["repayment_amount"] = pd.to_numeric(tmp["repayment_amount"], errors="coerce")
    tmp["amount_to_funder"] = pd.to_numeric(tmp["amount_to_funder"], errors="coerce")
    tmp["amount_transferred"] = pd.to_numeric(tmp["amount_transferred"], errors="coerce")
    tmp["overpaid_val"] = tmp["notes"].apply(parse_overpaid_amount)

    parents = tmp.groupby("claim_id", dropna=False).apply(canonical_parent).reset_index(drop=True)

    # Group by claim_id (aggregates multiple transfers together)
    roll = tmp.groupby("claim_id", dropna=False).agg(
        repayment_sum=("repayment_amount","sum"),
        amount_to_funder_sum=("amount_to_funder","sum"),
        amount_transferred_sum=("amount_transferred","sum"),
        ref_date=("window_date","max"),
        overpaid_sum=("overpaid_val","max"),
        notes_any=("notes"," | ".join)
    ).reset_index()
    roll = roll.merge(parents[["claim_id","claimant","deal_type","contract_date"]], on="claim_id", how="left")

    return roll, parents


def _setup_priority_sorting(roll):
    """
    Add priority columns for claims with identical amounts.

    When multiple claims have the same repayment amount, prioritize claims where
    the amount_transferred is closer to the repayment amount. This helps match
    the "main" claim (with the full payment) before partial claims.

    For example, if repayment = $2312.93:
    - Claim A with amount_transferred = $2250.79 (diff = $62.14) gets priority
    - Claim B with amount_transferred = $179.19 (diff = $2133.74) is secondary

    Args:
        roll: Rolled-up claims data

    Returns:
        DataFrame with priority columns and sorted
    """
    # Calculate how close amount_transferred is to repayment_sum
    # Smaller difference = higher priority (processed first)
    # This prioritizes "main" payments over partial payments
    roll['transferred_repayment_diff'] = (
        roll['repayment_sum'].fillna(0) - roll['amount_transferred_sum'].fillna(0)
    ).abs()

    # Group claims by their repayment_sum (rounded to 2 decimals for comparison)
    roll['repayment_group'] = roll['repayment_sum'].round(2)

    # Within each repayment amount group, sort by transferred_repayment_diff (ascending)
    # This ensures claims with amount_transferred closest to repayment get priority
    return roll.sort_values(
        ['repayment_group', 'transferred_repayment_diff'],
        ascending=[False, True]
    ).reset_index(drop=True)


def _find_overpay_debit(debits, over_x, credit_date, ref_date, used_overpay_debit):
    """
    Find matching overpayment debit for a credit match.

    Args:
        debits: DataFrame of debit transactions
        over_x: Overpayment amount to match
        credit_date: Date of the credit transaction
        ref_date: Reference date from claim
        used_overpay_debit: Set of already-used debit indices

    Returns:
        tuple: (debit_date, debit_description) or (None, None) if not found
    """
    dwin = debits[(debits["amount"].abs().sub(over_x).abs() <= AMOUNT_TOL)].copy()

    # Prefer debits near the credit date
    dnear_credit = dwin[(dwin["posting_date"] >= credit_date - pd.Timedelta(days=DATE_WINDOW_DAYS)) &
                        (dwin["posting_date"] <= credit_date + pd.Timedelta(days=DATE_WINDOW_DAYS))]
    dnear_credit = dnear_credit.sort_values(["overpay_hint","posting_date"], ascending=[False, True])

    for didx, d in dnear_credit.iterrows():
        if didx not in used_overpay_debit:
            used_overpay_debit.add(didx)
            return d["posting_date"], d["description"]

    # Fallback near ref date
    if pd.notna(ref_date):
        dnear_ref = dwin[(dwin["posting_date"] >= ref_date - pd.Timedelta(days=OVERPAY_BACKFILL_WINDOW)) &
                         (dwin["posting_date"] <= ref_date + pd.Timedelta(days=OVERPAY_BACKFILL_WINDOW))]
        dnear_ref = dnear_ref.sort_values(["overpay_hint","posting_date"], ascending=[False, True])

        for didx, d in dnear_ref.iterrows():
            if didx not in used_overpay_debit:
                used_overpay_debit.add(didx)
                return d["posting_date"], d["description"]

    return None, None


def _try_match_repayment_plus_overpay(cand, repayment_sum, over_x, used_credit):
    """Try to match credit for repayment + overpayment amount."""
    if (over_x or 0.0) <= AMOUNT_TOL:
        return None, None

    for idx, c in cand.sort_values(["diff_claim_plus_over","date_delta","posting_date"]).iterrows():
        if idx not in used_credit and abs(c["amount"] - (repayment_sum + over_x)) <= AMOUNT_TOL:
            return idx, c
    return None, None


def _try_match_fedwire_by_name(credits, claimant_name, repayment_sum, used_credit):
    """Try to match FEDWIRE credits by claimant name."""
    name_parts = [part.strip().upper() for part in claimant_name.replace(",", "").split() if len(part.strip()) > 2]

    if len(name_parts) < 2:
        return None, None

    fedwire_credits = credits[credits["description"].str.contains("FEDWIRE", na=False, case=False)]
    for idx, c in fedwire_credits.iterrows():
        if idx not in used_credit and abs(c["amount"] - repayment_sum) <= AMOUNT_TOL:
            desc_upper = c["description"].upper()
            if all(part in desc_upper for part in name_parts):
                return idx, c
    return None, None


def _try_match_exact_repayment(cand, repayment_sum, used_credit):
    """Try to match credit for exact repayment amount."""
    for idx, c in cand.sort_values(["diff_claim","date_delta","posting_date"]).iterrows():
        if idx not in used_credit and abs(c["amount"] - repayment_sum) <= AMOUNT_TOL:
            return idx, c
    return None, None


def parse_overpaid_amount(notes: str):
    if not isinstance(notes, str) or not notes.strip():
        return None
    m = OVERPAID_REGEX.search(notes)
    if not m:
        return None
    try:
        return r2(str(m.group(1)).replace(",", ""))
    except Exception:
        return None

def match_credits_one_row_per_claim(r2d, chase):
    """
    Match credits to claims using multiple strategies.

    This function aggregates repayments by claim, then attempts to match each claim
    to credit transactions using: repayment+overpay amount, FEDWIRE name matching,
    or exact repayment amount.

    Args:
        r2d: DataFrame with repayment data
        chase: DataFrame with Chase transactions

    Returns:
        tuple: (credit_matches, unmatched_claims, unmatched_credits, per_claim_revenue,
                overpayment_adjustments, used_credit_indices, used_overpay_debit_indices)
    """
    credits = chase[chase["is_credit"]].copy()
    debits = chase[chase["is_debit"]].copy()

    # Prepare and aggregate data
    roll, parents = _prepare_credit_matching_data(r2d)

    # Setup priority sorting for claims with identical amounts
    roll_sorted = _setup_priority_sorting(roll)

    # Match credits to claims
    results, used_credit, used_overpay_debit, unmatched_claims = [], set(), set(), []

    for _, r in roll_sorted.iterrows():
        claim_id = r["claim_id"]
        repayment_sum = r2(r.get("repayment_sum") or 0.0)

        if not repayment_sum or repayment_sum <= 0:
            unmatched_claims.append(claim_id)
            continue

        ref_date = r["ref_date"]
        over_x = r2(r.get("overpaid_sum") or 0.0) if pd.notna(r.get("overpaid_sum")) else 0.0

        # Wider window when we have an overpay
        win_days = max(DATE_WINDOW_DAYS, OVERPAY_BACKFILL_WINDOW) if (over_x or 0.0) > AMOUNT_TOL else DATE_WINDOW_DAYS

        # Filter credits by date window and calculate differences
        cand = credits.copy()
        if pd.notna(ref_date):
            cand = cand[(cand["posting_date"] >= ref_date - pd.Timedelta(days=win_days)) &
                        (cand["posting_date"] <= ref_date + pd.Timedelta(days=win_days))]
        cand = cand.assign(
            diff_claim=(cand["amount"] - repayment_sum).abs(),
            diff_claim_plus_over=(cand["amount"] - (repayment_sum + (over_x or 0.0))).abs(),
            date_delta=(cand["posting_date"] - ref_date).abs().dt.days if pd.notna(ref_date) else 0
        )

        # Try multiple matching strategies in priority order
        chosen_idx, chosen_credit = _try_match_repayment_plus_overpay(cand, repayment_sum, over_x, used_credit)
        match_mode = "repayment_sum_plus_overpay" if chosen_idx else None

        if not chosen_idx:
            chosen_idx, chosen_credit = _try_match_fedwire_by_name(credits, r["claimant"], repayment_sum, used_credit)
            match_mode = "fedwire_name_match" if chosen_idx else None

        if not chosen_idx:
            chosen_idx, chosen_credit = _try_match_exact_repayment(cand, repayment_sum, used_credit)
            match_mode = "repayment_sum" if chosen_idx else None

        # Process match and find overpayment debit if applicable
        if chosen_idx:
            used_credit.add(chosen_idx)
            over_debit_date, over_debit_desc = None, None

            if match_mode == "repayment_sum_plus_overpay" and (over_x or 0.0) > AMOUNT_TOL:
                over_debit_date, over_debit_desc = _find_overpay_debit(
                    debits, over_x, chosen_credit["posting_date"], ref_date, used_overpay_debit
                )

            results.append({
                "claim_id": claim_id,
                "claimant": r["claimant"],
                "deal_type_parent": r["deal_type"],
                "contract_date_parent": r["contract_date"],
                "repayment_sum": repayment_sum,
                "amount_to_funder_sum": r.get("amount_to_funder_sum"),
                "ref_date": ref_date,
                "match_type": match_mode,
                "chase_credit_date": chosen_credit["posting_date"],
                "chase_credit_amount": chosen_credit["amount"],
                "overpay_amount": (over_x or 0.0) if match_mode == "repayment_sum_plus_overpay" else 0.0,
                "overpay_debit_date": over_debit_date,
                "overpay_debit_desc": over_debit_desc,
                "notes": r["notes_any"],
            })
        else:
            unmatched_claims.append(claim_id)

    # Build output dataframes
    credit_matches = pd.DataFrame(results)
    unmatched_df = roll[roll["claim_id"].isin(unmatched_claims)][[
        "claim_id","claimant","deal_type","contract_date","repayment_sum","amount_to_funder_sum","ref_date","overpaid_sum","notes_any"
    ]].rename(columns={"deal_type":"deal_type_parent","contract_date":"contract_date_parent"})

    used_credit_idx = list(used_credit)
    credits_unmatched_final = credits.loc[~credits.index.isin(used_credit_idx)].copy()

    over_adj = credit_matches[credit_matches["overpay_amount"] > AMOUNT_TOL][[
        "claim_id","claimant","chase_credit_date","overpay_amount","overpay_debit_date","overpay_debit_desc"
    ]].rename(columns={
        "chase_credit_date":"credit_date",
        "overpay_debit_date":"matched_debit_date",
        "overpay_debit_desc":"matched_debit_desc"
    })

    per_claim = roll.copy()
    per_claim["claim_revenue"] = (per_claim["repayment_sum"].fillna(0) - per_claim["amount_to_funder_sum"].fillna(0)).round(2)
    per_claim = per_claim[["claim_id","claimant","repayment_sum","amount_to_funder_sum","claim_revenue","contract_date","deal_type"]].rename(
        columns={"repayment_sum":"Repayment Sum","amount_to_funder_sum":"Amount To Funder Sum","deal_type":"Deal Type","contract_date":"Contract Date"}
    )

    return credit_matches, unmatched_df, credits_unmatched_final, per_claim, over_adj, used_credit_idx, list(used_overpay_debit)

# ------------------------- Notes-driven extra matching -------------------------

def detect_shared_check(text):
    """
    Detect if a note describes a shared check and extract shared client information.
    Returns: tuple (is_shared, client_count, other_client_names)
    """
    if not isinstance(text, str) or not text.strip():
        return False, 1, []
    
    other_clients = []
    client_count = 1
    
    # Look for "other client is [Name]" patterns
    other_client_matches = OTHER_CLIENT_REGEX.findall(text)
    if other_client_matches:
        other_clients = [name.strip() for name in other_client_matches]
        client_count = len(other_clients) + 1  # +1 for the current client
    
    # Look for explicit client count mentions
    count_matches = CLIENT_COUNT_REGEX.findall(text)
    if count_matches:
        try:
            explicit_count = int(count_matches[0])
            if explicit_count > 1:
                client_count = max(client_count, explicit_count)
        except ValueError:
            pass
    
    is_shared = client_count > 1 or len(other_clients) > 0
    return is_shared, client_count, other_clients

def extract_note_events(text, ref_date):
    events = []
    if not isinstance(text, str) or not text.strip():
        return events
    year = (ref_date.year if isinstance(ref_date, pd.Timestamp) and not pd.isna(ref_date) else date.today().year)
    dates = [pd.Timestamp(year=year, month=int(m), day=int(d), tz=None) for m, d in DATE_IN_NOTES.findall(text)]
    anchor = dates[-1] if dates else ref_date

    # Check if note indicates remaining repayment was received
    # Patterns: "rem. repayment received", "received rem. repayment", "remaining repayment received"
    has_rem_received = bool(re.search(r'rem\.?\s*repayment\s+received|remaining\s+repayment\s+received|received\s+rem\.?\s*repayment', text, re.I))

    # Collect all "req rem" amounts - these might be credits if rem repayment was received
    req_rem_amounts = []
    for m in REQUESTED_REMAINING_REGEX.finditer(text):
        amt = r2(m.group(2).replace(",",""))
        if amt is not None:
            req_rem_amounts.append(amt)

    # If note says "rem. repayment received" and has req rem amounts, those are credits
    received_amounts = set()
    if has_rem_received and req_rem_amounts:
        for amt in req_rem_amounts:
            received_amounts.add(amt)
            events.append(("credit_expected", amt, anchor))

    # Also check RECEIVED_CHECK_REGEX for explicit "received $X" patterns
    # But be careful not to match "to send funder $X" amounts
    for m in RECEIVED_CHECK_REGEX.finditer(text):
        amt = r2(m.group(2).replace(",",""))
        if amt is not None and amt not in received_amounts:
            # Check if this amount appears right after "to send funder" - if so, it's a debit not credit
            match_text = m.group(0)
            if 'send funder' in match_text.lower() or 'to funder' in match_text.lower():
                continue
            received_amounts.add(amt)
            events.append(("credit_expected", amt, anchor))

    # Collect amounts to ignore (requested remaining that weren't received)
    amounts_to_ignore = set()
    for amt in req_rem_amounts:
        if amt not in received_amounts:
            amounts_to_ignore.add(amt)

    for m in SEND_FUNDER_REGEX.finditer(text):
        amt = r2(m.group(1).replace(",",""))
        if amt is not None and amt not in amounts_to_ignore:
            events.append(("debit_expected", amt, anchor))

    for m in DOLLAR_REGEX.finditer(text):
        amt = r2(m.group(1).replace(",",""))
        if amt in amounts_to_ignore:
            continue
        start, end = max(0, m.start()-120), min(len(text), m.end()+120)
        ctx = text[start:end]
        is_credit = bool(CREDIT_KEYWORDS.search(ctx))
        is_debit  = bool(DEBIT_KEYWORDS.search(ctx))
        if is_credit and not is_debit:
            events.append(("credit_expected", amt, anchor))
        elif is_debit and not is_credit:
            events.append(("debit_expected", amt, anchor))

    seen = set(); uniq = []
    for kind, amt, ad in events:
        key = (kind, amt)
        if key in seen: continue
        seen.add(key); uniq.append((kind, amt, ad))
    return uniq

def match_from_notes(r2d, chase, used_credit_idx, used_debit_idx):
    """
    Detect expected credits/debits from notes but DO NOT auto-match credits.

    Credits: Only DETECT and report as "Missing ReconTag" - do not match.
    Debits: Still match (debits are less risky as they're outgoing payments).

    Returns:
        note_credit_df: Empty (no auto-matching)
        note_debit_df: Matched debits from notes
        newly_used_credit: Empty set
        newly_used_debit: Set of matched debit indices
        recontags_df: DataFrame of matched and suggested ReconTags (with status column)
    """
    credits = chase[chase["is_credit"]].copy()
    debits  = chase[chase["is_debit"]].copy()

    grp = r2d.groupby("claim_id", dropna=False)
    notes_join = grp["notes"].apply(lambda s: " | ".join([str(x) for x in s.dropna() if str(x).strip()]))
    ref_date = grp["window_date"].max()
    claimant = grp["claimant"].max()
    roll = pd.DataFrame({"claim_id": ref_date.index, "ref_date": ref_date.values, "notes_any": notes_join.values, "claimant": claimant.values})
    roll["claimant_display"] = roll["claimant"].astype(str).str.replace(PAREN_SUFFIX, "", regex=True)

    # No longer auto-match credits from notes
    note_credit_rows = []
    note_debit_rows = []
    matched_recontag_rows = []  # Claims that already have ReconTags
    suggested_recontag_rows = []  # Claims that need ReconTags
    newly_used_credit = set()  # Will stay empty - no credit matching
    newly_used_debit = set()

    # Build dict of claim IDs that already have ReconTagged credits (with their details)
    recontag_credits = {}
    if "recon_tag" in chase.columns:
        tagged = chase[chase["is_credit"] & chase["recon_tag"].notna() & (chase["recon_tag"].astype(str).str.strip() != "")]
        for _, row in tagged.iterrows():
            cid = str(row["recon_tag"]).strip()
            if cid not in recontag_credits:
                recontag_credits[cid] = []
            recontag_credits[cid].append({
                "amount": row["amount"],
                "date": row["posting_date"],
                "description": row.get("description", "")[:50]
            })

    for _, r in roll.iterrows():
        events = extract_note_events(r["notes_any"], r["ref_date"])
        for kind, amount, anchor_date in events:
            if kind == "credit_expected":
                claim_id = str(r["claim_id"]).strip()

                # Check if this claim already has a ReconTagged credit
                if claim_id in recontag_credits:
                    # Report as matched ReconTag
                    for tagged_credit in recontag_credits[claim_id]:
                        matched_recontag_rows.append({
                            "status": "MATCHED",
                            "claim_id": r["claim_id"],
                            "claimant": r["claimant_display"],
                            "expected_amount": amount,
                            "matched_credit_amount": tagged_credit["amount"],
                            "matched_credit_date": tagged_credit["date"],
                            "note_excerpt": r["notes_any"][:150] + "..." if len(r["notes_any"]) > 150 else r["notes_any"],
                        })
                    continue

                # DO NOT auto-match credits - instead report as needing ReconTag
                # Check if there's a potential matching credit (for reporting purposes)
                cnd = credits.loc[~credits.index.isin(used_credit_idx)].copy()
                if pd.notna(anchor_date):
                    cnd = cnd[(cnd["posting_date"] >= anchor_date - pd.Timedelta(days=NOTE_WINDOW_DAYS)) &
                              (cnd["posting_date"] <= anchor_date + pd.Timedelta(days=NOTE_WINDOW_DAYS))]
                cnd = cnd[(cnd["amount"].sub(amount).abs() <= AMOUNT_TOL)]

                potential_match = None
                if not cnd.empty:
                    potential_match = cnd.sort_values("posting_date").iloc[0]
                else:
                    rd = r["ref_date"]
                    if pd.notna(rd):
                        cnd = credits.loc[~credits.index.isin(used_credit_idx)].copy()
                        cnd = cnd[(cnd["posting_date"] >= rd - pd.Timedelta(days=NOTE_WINDOW_DAYS+NOTE_WINDOW_DAYS_EXTENDED)) &
                                  (cnd["posting_date"] <= rd + pd.Timedelta(days=NOTE_WINDOW_DAYS+NOTE_WINDOW_DAYS_EXTENDED))]
                        cnd = cnd[(cnd["amount"].sub(amount).abs() <= AMOUNT_TOL)]
                        if not cnd.empty:
                            potential_match = cnd.sort_values("posting_date").iloc[0]

                # Report this as needing a ReconTag
                suggested_recontag_rows.append({
                    "status": "SUGGESTED",
                    "claim_id": r["claim_id"],
                    "claimant": r["claimant_display"],
                    "expected_amount": amount,
                    "potential_chase_credit": potential_match["amount"] if potential_match is not None else None,
                    "potential_chase_date": potential_match["posting_date"] if potential_match is not None else None,
                    "note_excerpt": r["notes_any"][:150] + "..." if len(r["notes_any"]) > 150 else r["notes_any"],
                    "action_needed": f"Add ReconTag '{r['claim_id']}' to Chase credit of ${amount:.2f}"
                })

            elif kind == "debit_expected":
                # Still match debits (less risky - outgoing payments)
                dnd = debits.loc[~debits.index.isin(used_debit_idx + list(newly_used_debit))].copy()
                if pd.notna(anchor_date):
                    dnd = dnd[(dnd["posting_date"] >= anchor_date - pd.Timedelta(days=NOTE_WINDOW_DAYS)) &
                              (dnd["posting_date"] <= anchor_date + pd.Timedelta(days=NOTE_WINDOW_DAYS))]
                dnd = dnd[(dnd["amount"].abs().sub(amount).abs() <= AMOUNT_TOL)]
                chosen = None
                if not dnd.empty:
                    chosen = dnd.sort_values("posting_date").iloc[0]
                else:
                    rd = r["ref_date"]
                    if pd.notna(rd):
                        dnd = debits.loc[~debits.index.isin(used_debit_idx + list(newly_used_debit))].copy()
                        dnd = dnd[(dnd["posting_date"] >= rd - pd.Timedelta(days=NOTE_WINDOW_DAYS+NOTE_WINDOW_DAYS_EXTENDED)) &
                                  (dnd["posting_date"] <= rd + pd.Timedelta(days=NOTE_WINDOW_DAYS+NOTE_WINDOW_DAYS_EXTENDED))]
                        dnd = dnd[(dnd["amount"].abs().sub(amount).abs() <= AMOUNT_TOL)]
                        if not dnd.empty:
                            chosen = dnd.sort_values("posting_date").iloc[0]
                if chosen is not None:
                    newly_used_debit.add(chosen.name)
                    note_debit_rows.append({
                        "claim_id": r["claim_id"],
                        "claimant": r["claimant_display"],
                        "note_amount": amount,
                        "matched_debit_date": chosen["posting_date"],
                        "matched_debit_amount": chosen["amount"],
                        "matched_debit_desc": chosen["description"],
                        "source": "notes"
                    })

    note_credit_df = pd.DataFrame(note_credit_rows)  # Will be empty
    note_debit_df  = pd.DataFrame(note_debit_rows)

    # Combine matched and suggested into one DataFrame with status column
    # Matched tags first, then suggested
    matched_df = pd.DataFrame(matched_recontag_rows)
    suggested_df = pd.DataFrame(suggested_recontag_rows)
    recontags_df = pd.concat([matched_df, suggested_df], ignore_index=True)

    return note_credit_df, note_debit_df, list(newly_used_credit), list(newly_used_debit), recontags_df

# ------------------------- Revenue & Summary -------------------------

def compute_bank_revenue_per_claim(d_match, c_match, note_c, note_d, per_claim, overpay_adj=None):
    if not c_match.empty:
        eff = c_match.copy()
        # Always use chase_credit_amount for effective credit
        # For overpay cases, chase_credit_amount is already the net amount (repayment - overpay)
        eff["effective_credit"] = eff["chase_credit_amount"]
        cm_per_claim = eff.groupby("claim_id", dropna=False)["effective_credit"].sum()
    else:
        cm_per_claim = pd.Series(dtype=float)

    note_c_per = (note_c.groupby("claim_id", dropna=False)["matched_credit_amount"].sum()
                  if not note_c.empty else pd.Series(dtype=float))

    # Separate overpayment debits from funder debits
    overpay_debits_per = pd.Series(dtype=float)
    funder_debits_per = pd.Series(dtype=float)

    if not d_match.empty:
        # Exclude debits that are already processed as overpayment adjustments to avoid double counting
        if overpay_adj is not None and not overpay_adj.empty:
            # Get the date and amount combinations from overpayment adjustments
            overpay_dates_amounts = set(zip(overpay_adj['claim_id'], overpay_adj['matched_debit_date'], overpay_adj['overpay_amount']))

            # Filter out debits that match overpayment adjustments
            def is_not_overpay_adj(row):
                return (row['claim_id'], pd.to_datetime(row.get('chase_date', row.get('chase_debit_date', pd.NaT))), abs(row['chase_amount'])) not in overpay_dates_amounts

            d_match_filtered = d_match[d_match.apply(is_not_overpay_adj, axis=1)].copy()
        else:
            d_match_filtered = d_match.copy()

        # Use filtered debit matches to avoid double counting overpayment adjustments
        d_with_repay = d_match_filtered.merge(per_claim[['claim_id', 'Repayment Sum']], on='claim_id', how='left')

        overpay_debits = d_with_repay[
            (d_with_repay.get("match_type", "").astype(str).str.contains("overpay", case=False, na=False)) |
            (d_with_repay["description"].str.contains(OVERPAY_DESC_PATTERN, case=False, na=False))
        ]

        funder_debits = d_with_repay[
            ~(d_with_repay.get("match_type", "").astype(str).str.contains("overpay", case=False, na=False)) &
            ~(d_with_repay["description"].str.contains(OVERPAY_DESC_PATTERN, case=False, na=False))
        ]

        overpay_debits_per = (overpay_debits.groupby("claim_id", dropna=False)["chase_amount"].apply(lambda s: s.abs().sum())
                             if not overpay_debits.empty else pd.Series(dtype=float))
        funder_debits_per = (funder_debits.groupby("claim_id", dropna=False)["chase_amount"].apply(lambda s: s.abs().sum())
                            if not funder_debits.empty else pd.Series(dtype=float))

    # Same logic for note debits - separate overpayment from funder debits
    note_overpay_per = pd.Series(dtype=float)
    note_funder_per = pd.Series(dtype=float)

    if not note_d.empty:
        # Check if note debits are overpayment-related based on context
        if "context" in note_d.columns:
            context_col = note_d["context"].fillna("").astype(str)
            note_overpay = note_d[context_col.str.contains("overpay|overpaid", case=False, na=False)]
            note_funder = note_d[~context_col.str.contains("overpay|overpaid", case=False, na=False)]
        else:
            # If no context column, assume all note debits are funder debits for now
            note_overpay = pd.DataFrame()
            note_funder = note_d

        note_overpay_per = (note_overpay.groupby("claim_id", dropna=False)["matched_debit_amount"].apply(lambda s: s.abs().sum())
                           if not note_overpay.empty else pd.Series(dtype=float))
        note_funder_per = (note_funder.groupby("claim_id", dropna=False)["matched_debit_amount"].apply(lambda s: s.abs().sum())
                          if not note_funder.empty else pd.Series(dtype=float))

    # Handle overpayment adjustments
    overpay_adj_per = pd.Series(dtype=float)
    if overpay_adj is not None and not overpay_adj.empty:
        overpay_adj_per = overpay_adj.groupby("claim_id", dropna=False)["overpay_amount"].sum()

    out = per_claim.copy()
    out["Claimant"] = out.get("claimant", out.get("Claimant", ""))
    out["Claimant"] = out["Claimant"].astype(str).str.replace(PAREN_SUFFIX, "", regex=True)

    # Bank Credits (Effective) = Credits - Overpayment Debits - Overpayment Adjustments
    # This handles both patterns:
    # 1. Overpaid settlements: chase_credit_amount is already net
    # 2. Check with overpayment deduction: gross credit - overpayment debit = net
    # 3. Overpayment adjustments: gross credit - overpayment adjustment = net
    out["Bank Credits (Effective)"] = (
        out["claim_id"].map(cm_per_claim).fillna(0) +
        out["claim_id"].map(note_c_per).fillna(0) -
        out["claim_id"].map(overpay_debits_per).fillna(0) -
        out["claim_id"].map(note_overpay_per).fillna(0) -
        out["claim_id"].map(overpay_adj_per).fillna(0)
    ).round(2)

    # Bank Funder Debits = Only actual funder debits (not overpayments)
    out["Bank Funder Debits"] = (
        out["claim_id"].map(funder_debits_per).fillna(0) +
        out["claim_id"].map(note_funder_per).fillna(0)
    ).round(2)

    out["Bank-based Revenue"] = (out["Bank Credits (Effective)"] - out["Bank Funder Debits"]).round(2)
    out["Book Revenue (KEEP)"] = (out["Repayment Sum"].fillna(0) - out["Amount To Funder Sum"].fillna(0)).round(2)
    out["Check (Bank - Book)"] = (out["Bank-based Revenue"] - out["Book Revenue (KEEP)"]).round(2)

    cols = ["claim_id","Claimant","Repayment Sum","Amount To Funder Sum","Book Revenue (KEEP)",
            "Bank Credits (Effective)","Bank Funder Debits","Bank-based Revenue","Check (Bank - Book)"]
    existing = [c for c in cols if c in out.columns]
    remaining = [c for c in out.columns if c not in existing]
    return out[existing + remaining]

def detect_multiple_transfers(r2d):
    """
    Detect claims that have multiple transfers (different ACH IDs and transfer dates).
    These may cause reconciliation issues and should be reviewed.

    Returns:
        DataFrame with claims that have multiple transfers
    """
    # Group by claim_id and check for multiple unique ACH IDs or transfer dates
    grouped = r2d.groupby("claim_id").agg({
        "ach_id": lambda x: x.nunique(),
        "window_date": lambda x: x.nunique(),
        "claimant": "first",
        "repayment_amount": "sum",
        "amount_to_funder": "sum"
    }).reset_index()

    # Find claims with multiple ACH IDs (different transfers)
    multiple_transfers = grouped[grouped["ach_id"] > 1].copy()

    if not multiple_transfers.empty:
        multiple_transfers = multiple_transfers.rename(columns={
            "ach_id": "transfer_count",
            "window_date": "unique_dates"
        })
        return multiple_transfers[["claim_id", "claimant", "transfer_count", "unique_dates", "repayment_amount", "amount_to_funder"]]

    return pd.DataFrame()


def build_summary(d_match, d_un, d_orph, c_match, note_c, note_d, multiple_transfers_df=None):
    total_d = r2(d_match["chase_amount"].abs().sum(), 2) if not d_match.empty else 0.0
    total_c = r2(c_match["chase_credit_amount"].sum(), 2) if not c_match.empty else 0.0
    total_c_notes = r2(note_c["matched_credit_amount"].sum(), 2) if not note_c.empty else 0.0
    total_d_notes = r2(note_d["matched_debit_amount"].abs().sum(), 2) if not note_d.empty else 0.0
    total_d_all = r2((total_d or 0) + (total_d_notes or 0), 2)
    total_c_all = r2((total_c or 0) + (total_c_notes or 0), 2)
    net_after_notes = r2((total_d_all or 0) - (total_c_all or 0), 2)

    metrics = [
        "Debits matched (count)",
        "Credits matched (count)",
        "Note-derived debit matches (count)",
        "Note-derived credit matches (count)",
        "Total debits matched incl. notes (abs)",
        "Total credits matched incl. notes",
        "Net diff after notes (debits - credits)",
        "Debits unmatched (count)",
        "CHASE unmatched debits (count)"
    ]
    values = [
        len(d_match), len(c_match),
        (0 if note_d.empty else len(note_d)),
        (0 if note_c.empty else len(note_c)),
        total_d_all, total_c_all, net_after_notes,
        len(d_un), len(d_orph)
    ]

    # Add warning about multiple transfers if any found
    if multiple_transfers_df is not None and not multiple_transfers_df.empty:
        metrics.append("⚠️ Claims with multiple transfers")
        values.append(len(multiple_transfers_df))

    return pd.DataFrame({"metric": metrics, "value": values})

def build_unmatched_combined(credits_unmatched_final, debits_orphans_final, c_un_claims, d_un=None, reconciled_credit_claims=None, reconciled_debit_claims=None, expected_overpay_missing=None):
    reconciled_credit_claims = reconciled_credit_claims or set()
    reconciled_debit_claims = reconciled_debit_claims or set()
    
    cu = credits_unmatched_final.copy()
    if not cu.empty:
        cu = cu.rename(columns={"posting_date":"date","description":"description","Amount":"amount","amount":"amount"})
        cu = cu.assign(category="CHASE_Unmatched_Credit",
                       claim_id=pd.NA, claimant=pd.NA,
                       notes=pd.NA)[["category","claim_id","claimant","date","amount","description","notes"]]
    du = debits_orphans_final.copy()
    if not du.empty:
        du = du.rename(columns={"posting_date":"date","description":"description","Amount":"amount","amount":"amount"})
        du = du.assign(category="CHASE_Unmatched_Debit",
                       claim_id=pd.NA, claimant=pd.NA,
                       notes=pd.NA)[["category","claim_id","claimant","date","amount","description","notes"]]
    # Add unmatched R2D debits (transfers that couldn't be matched to Chase)
    d_unmatched = None
    if d_un is not None and not d_un.empty:
        d_unmatched = d_un.copy()
        # Only exclude R2D unmatched debits if they were specifically matched as debits
        if reconciled_debit_claims:
            d_unmatched = d_unmatched.loc[~d_unmatched["claim_id"].astype(str).isin(reconciled_debit_claims)].copy()

        if not d_unmatched.empty:
            # Deduplicate by claim_id to avoid showing AFR variations multiple times
            # Use canonical parent to get the main claimant name (non-AFR preferred)
            def get_parent_row(group):
                # Check deal_type if available, prefer non-AFR
                if "deal_type" in group.columns:
                    group["is_afr_temp"] = group["deal_type"].astype(str).str.contains(r"\bafr", case=False, na=False)
                    non_afr = group[~group["is_afr_temp"]]
                    if not non_afr.empty:
                        return non_afr.iloc[0]
                return group.iloc[0]

            parents = d_unmatched.groupby("claim_id", dropna=False).apply(get_parent_row).reset_index(drop=True)

            # Group by claim_id and aggregate amounts
            # For amount_transferred: if all rows have same ACH ID, it's the same transfer (take first)
            # If different ACH IDs, sum them (multiple actual transfers)
            # Check ACH IDs per claim first
            ach_id_counts = d_unmatched.groupby("claim_id")["ach_id"].nunique() if "ach_id" in d_unmatched.columns else pd.Series()

            d_unmatched_grouped = d_unmatched.groupby("claim_id", dropna=False).agg({
                "transfer_initiated": "max",  # Latest transfer date
                "amount_transferred": "first",  # Take first (will fix for multiple ACH IDs below)
                "notes": lambda x: " | ".join([str(n) for n in x.dropna() if str(n).strip()]) if x.notna().any() else pd.NA
            }).reset_index()

            # For claims with multiple unique ACH IDs, sum the amounts instead
            if not ach_id_counts.empty:
                for claim_id in ach_id_counts[ach_id_counts > 1].index:
                    claim_rows = d_unmatched[d_unmatched["claim_id"] == claim_id]
                    total_amount = claim_rows["amount_transferred"].sum()
                    d_unmatched_grouped.loc[d_unmatched_grouped["claim_id"] == claim_id, "amount_transferred"] = total_amount

            # Merge with parent claimant names
            d_unmatched_grouped = d_unmatched_grouped.merge(
                parents[["claim_id", "claimant"]],
                on="claim_id",
                how="left"
            )

            # Use transfer_initiated as date, amount_transferred as amount
            rename_map = {"transfer_initiated":"date","amount_transferred":"amount","notes":"notes","claimant":"claimant","claim_id":"claim_id"}
            d_unmatched_grouped = d_unmatched_grouped.rename(columns=rename_map)
            if "date" not in d_unmatched_grouped.columns and "transfer_initiated" in d_unmatched_grouped.columns:
                d_unmatched_grouped["date"] = d_unmatched_grouped["transfer_initiated"]
            if "amount" not in d_unmatched_grouped.columns and "amount_transferred" in d_unmatched_grouped.columns:
                d_unmatched_grouped["amount"] = d_unmatched_grouped["amount_transferred"]
            d_unmatched_grouped = d_unmatched_grouped.assign(category="R2D_Unmatched_Debit (transfer)")
            d_unmatched_grouped["description"] = pd.NA
            d_unmatched = d_unmatched_grouped[["category","claim_id","claimant","date","amount","description","notes"]]
        else:
            d_unmatched = None
            
    cu_claims = c_un_claims.copy()
    if not cu_claims.empty:
        # Only exclude unmatched credit claims if they were specifically matched as credits
        if reconciled_credit_claims:
            cu_claims = cu_claims.loc[~cu_claims["claim_id"].astype(str).isin(reconciled_credit_claims)].copy()
            
        if not cu_claims.empty:
            # Expected credit should be the full repayment amount that was being searched for
            cu_claims["expected_credit"] = cu_claims.get("repayment_sum", 0)

            rename_map = {"ref_date":"date","notes_any":"notes","expected_credit":"amount","claimant":"claimant","claim_id":"claim_id"}
            if "Repayment Sum" in cu_claims.columns: rename_map["repayment_sum"] = "Repayment Sum"
            if "notes" in cu_claims.columns: rename_map["notes_any"] = "notes"
            cu_claims = cu_claims.rename(columns=rename_map)
            if "date" not in cu_claims.columns and "ref_date" in cu_claims.columns:
                cu_claims["date"] = cu_claims["ref_date"]
            if "amount" not in cu_claims.columns and "expected_credit" in cu_claims.columns:
                cu_claims["amount"] = cu_claims["expected_credit"]
            cu_claims = cu_claims.assign(category="Claim_Unmatched_Credit (expected)")
            cu_claims["description"] = pd.NA
            cu_claims = cu_claims[["category","claim_id","claimant","date","amount","description","notes"]]
        else:
            cu_claims = None
            
    # Add expected overpay debits that couldn't be matched to any Chase debit
    if expected_overpay_missing is not None and not expected_overpay_missing.empty:
        # Ensure correct columns and order
        cols = ["category","claim_id","claimant","date","amount","description","notes"]
        expected_overpay_missing = expected_overpay_missing[cols]

    frames = [df for df in [cu, du, d_unmatched, cu_claims, expected_overpay_missing] if df is not None and not df.empty]
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=["category","claim_id","claimant","date","amount","description","notes"])

# ------------------------- Orchestration -------------------------

def run(file_path, r2d_sheet, chase_sheet, out_path, ignore_debits_before=None, pre_balance=None, pre_exclusions=None):
    """
    Main reconciliation orchestration function.

    Args:
        file_path: Path to input Excel file
        r2d_sheet: Name of Repayments to Date sheet
        chase_sheet: Name of Chase transactions sheet
        out_path: Path for output Excel file
        ignore_debits_before: Optional date to exclude older debits
        pre_balance: Unused parameter
        pre_exclusions: Unused parameter

    Raises:
        FileNotFoundError: If input file doesn't exist
        ValueError: If required sheets or columns are missing
    """
    logger.info("="*70)
    logger.info("RECONCILIATION STARTING")
    logger.info("="*70)

    # Validate input file
    validate_file_exists(file_path)
    validate_sheet_exists(file_path, r2d_sheet)
    validate_sheet_exists(file_path, chase_sheet)

    # Load data
    logger.info("\n--- Loading Data ---")
    r2d   = load_r2d(file_path, r2d_sheet)
    chase = load_chase(file_path, chase_sheet)

    # Correlation map
    corr_map = build_correlation_map(r2d)

    # Pre-filter ReconTag credits so they're excluded from main matching algorithm
    # This allows those credits to be reserved for their ReconTag claims
    recontag_credit_indices = set()
    if "recon_tag" in chase.columns:
        nonempty_tag = chase["recon_tag"].notna() & chase["recon_tag"].astype(str).str.strip().ne("")
        tagged_credits = chase.loc[chase["is_credit"] & nonempty_tag]
        if not tagged_credits.empty:
            recontag_credit_indices = set(tagged_credits.index)
            logger.info(f"Pre-filtering {len(recontag_credit_indices)} ReconTag credits from main matching algorithm")

    # Debits
    d_match, d_un, d_orph, dup, used_debit_idx, ach_id_conflicts = match_debits_relaxed(r2d, chase)

    # Credits (+ overpay) - pass ReconTag indices to exclude from matching
    # Create a modified chase dataframe that excludes ReconTag credits
    chase_for_matching = chase.copy()
    if recontag_credit_indices:
        chase_for_matching = chase_for_matching.loc[~chase_for_matching.index.isin(recontag_credit_indices)].copy()

    c_match, c_un_claims, c_un_chase_final, per_claim, over_adj, used_credit_idx, used_overpay_debit = \
        match_credits_one_row_per_claim(r2d, chase_for_matching)

    # Remove overpay-linked debits from orphans
    if used_overpay_debit:
        d_orph = d_orph.loc[~d_orph.index.isin(used_overpay_debit)].copy()

    # Notes pass - detects expected credits but does NOT auto-match them
    # Credits need ReconTags for reliable matching
    note_c, note_d, newly_used_credit, newly_used_debit, missing_recontags = \
        match_from_notes(r2d, chase, used_credit_idx, used_debit_idx)

    # Unmatched views (exclude note-derived uses)
    credits_unmatched_final = c_un_chase_final.loc[~c_un_chase_final.index.isin(newly_used_credit)].copy()
    debits_orphans_final    = d_orph.loc[~d_orph.index.isin(newly_used_debit)].copy()

    # Claims satisfied by notes
    if not note_c.empty and not c_un_claims.empty:
        note_credit_claims = set(note_c["claim_id"].dropna().astype(str))
        c_un_claims = c_un_claims.loc[~c_un_claims["claim_id"].astype(str).isin(note_credit_claims)].copy()

    # ReconTag handling:
    tagged_sum_by_claim = pd.Series(dtype=float)
    tagged_idx = set()
    # Process ReconTags for Bank Credits
    if "recon_tag" in chase.columns:
        nonempty_tag = chase["recon_tag"].notna() & chase["recon_tag"].astype(str).str.strip().ne("")
        tagged = chase.loc[chase["is_credit"] & nonempty_tag].copy()

        if not tagged.empty:
            tagged["recon_tag"] = tagged["recon_tag"].astype(str).str.strip()
            tagged_sum_by_claim = tagged.groupby("recon_tag")["amount"].sum()
            tagged_idx = set(tagged.index)

        # Hide ONLY those tagged rows from CHASE_Unmatched_Credits
        if not credits_unmatched_final.empty and tagged_idx:
            credits_unmatched_final = credits_unmatched_final.loc[
                ~credits_unmatched_final.index.isin(tagged_idx)
            ].copy()

    # Optional cutoff filtering
    excluded_debits_by_date = None
    excluded_dun_by_ti = None
    if ignore_debits_before:
        cutoff = pd.to_datetime(ignore_debits_before, errors="coerce")
        if pd.notna(cutoff):
            if "posting_date" in debits_orphans_final.columns:
                mask = debits_orphans_final["posting_date"] < cutoff
                excluded_debits_by_date = debits_orphans_final.loc[mask].copy()
                debits_orphans_final = debits_orphans_final.loc[~mask].copy()
            if "transfer_initiated" in d_un.columns:
                mask2 = d_un["transfer_initiated"].notna() & (d_un["transfer_initiated"] < cutoff)
                excluded_dun_by_ti = d_un.loc[mask2].copy()
                d_un = d_un.loc[~mask2].copy()

    # Detect claims with multiple transfers for data quality review
    multiple_transfers_df = detect_multiple_transfers(r2d)
    if not multiple_transfers_df.empty:
        logger.warning(f"⚠️  Found {len(multiple_transfers_df)} claim(s) with multiple transfers - review source file!")
        logger.warning(f"Claims: {', '.join(multiple_transfers_df['claimant'].tolist())}")

    # Summary & base per-claim revenue
    summary = build_summary(d_match, d_un, debits_orphans_final, c_match, note_c, note_d, multiple_transfers_df)
    per_claim_rev = compute_bank_revenue_per_claim(d_match, c_match, note_c, note_d, per_claim, over_adj)

    # Add ReconTagged credits to Bank Credits (Effective)
    # ReconTags ALWAYS take priority and should ALWAYS be added
    if not tagged_sum_by_claim.empty:
        per_claim_rev["claim_id"] = per_claim_rev["claim_id"].astype(str).str.strip()

        # ReconTags are always authoritative - add them without filtering
        per_claim_rev["Bank Credits (Effective)"] = (
            per_claim_rev["Bank Credits (Effective)"].fillna(0) +
            per_claim_rev["claim_id"].map(tagged_sum_by_claim).fillna(0)
        ).round(2)
        if "Bank Funder Debits" in per_claim_rev.columns:
            per_claim_rev["Bank-based Revenue"] = (per_claim_rev["Bank Credits (Effective)"] - per_claim_rev["Bank Funder Debits"]).round(2)
        if "Book Revenue (KEEP)" in per_claim_rev.columns:
            per_claim_rev["Check (Bank - Book)"] = (per_claim_rev["Bank-based Revenue"] - per_claim_rev["Book Revenue (KEEP)"]).round(2)
        logger.info(f"ReconTagged credits merged for {tagged_sum_by_claim.index.nunique()} claim(s), total = {float(tagged_sum_by_claim.sum()):.2f}")

    # Process ReconTagged debits (overpayment returns)
    # These should be deducted from Bank Credits (Effective)
    # BUT: exclude debits already accounted for in over_adj to avoid double-counting
    tagged_debit_sum_by_claim = pd.Series(dtype=float)
    tagged_debit_idx = set()
    if "recon_tag" in chase.columns:
        nonempty_tag = chase["recon_tag"].notna() & chase["recon_tag"].astype(str).str.strip().ne("")
        tagged_debits = chase.loc[chase["is_debit"] & nonempty_tag].copy()

        if not tagged_debits.empty:
            tagged_debits["recon_tag"] = tagged_debits["recon_tag"].astype(str).str.strip()

            # Exclude debits already in over_adj (matched via credit matching overpay logic)
            # These are already deducted in compute_bank_revenue_per_claim
            if over_adj is not None and not over_adj.empty:
                # Build set of (claim_id, amount) pairs already processed
                already_processed = set()
                for _, oa in over_adj.iterrows():
                    cid = str(oa.get("claim_id", "")).strip()
                    amt = abs(float(oa.get("overpay_amount", 0)))
                    if cid and amt > AMOUNT_TOL:
                        already_processed.add((cid, round(amt, 2)))

                # Filter out tagged debits that match already-processed overpayments
                def not_already_processed(row):
                    cid = str(row["recon_tag"]).strip()
                    amt = round(abs(row["amount"]), 2)
                    return (cid, amt) not in already_processed

                tagged_debits = tagged_debits[tagged_debits.apply(not_already_processed, axis=1)]

            # Sum the absolute values of remaining debits by claim
            if not tagged_debits.empty:
                tagged_debit_sum_by_claim = tagged_debits.groupby("recon_tag")["amount"].apply(lambda x: x.abs().sum())
            tagged_debit_idx = set(tagged_debits.index)

            # Deduct ReconTagged debits from Bank Credits (Effective)
            per_claim_rev["claim_id"] = per_claim_rev["claim_id"].astype(str).str.strip()
            per_claim_rev["Bank Credits (Effective)"] = (
                per_claim_rev["Bank Credits (Effective)"].fillna(0) -
                per_claim_rev["claim_id"].map(tagged_debit_sum_by_claim).fillna(0)
            ).round(2)

            # Recalculate revenue
            if "Bank Funder Debits" in per_claim_rev.columns:
                per_claim_rev["Bank-based Revenue"] = (per_claim_rev["Bank Credits (Effective)"] - per_claim_rev["Bank Funder Debits"]).round(2)
            if "Book Revenue (KEEP)" in per_claim_rev.columns:
                per_claim_rev["Check (Bank - Book)"] = (per_claim_rev["Bank-based Revenue"] - per_claim_rev["Book Revenue (KEEP)"]).round(2)

            logger.info(f"ReconTagged debits (overpayments) processed for {tagged_debit_sum_by_claim.index.nunique()} claim(s), total = {float(tagged_debit_sum_by_claim.sum()):.2f}")

        # Hide ReconTagged debits from CHASE_Unmatched_Debits
        if not debits_orphans_final.empty and tagged_debit_idx:
            debits_orphans_final = debits_orphans_final.loc[
                ~debits_orphans_final.index.isin(tagged_debit_idx)
            ].copy()

    # Insert correlation id
    for df in (d_match, d_un, c_match, over_adj, note_c, note_d, per_claim_rev, c_un_claims):
        insert_corr(df, corr_map, claim_col="claim_id", pos=1)

    # Collect reconciled claim IDs separately for credits and debits
    reconciled_credit_claims = set()
    reconciled_debit_claims = set()
    
    # Add credit matches and note matched credits
    if not c_match.empty and "claim_id" in c_match.columns:
        reconciled_credit_claims.update(c_match["claim_id"].dropna().astype(str))
    if not note_c.empty and "claim_id" in note_c.columns:
        reconciled_credit_claims.update(note_c["claim_id"].dropna().astype(str))
    
    # Add debit matches and note matched debits  
    if not d_match.empty and "claim_id" in d_match.columns:
        reconciled_debit_claims.update(d_match["claim_id"].dropna().astype(str))
    if not note_d.empty and "claim_id" in note_d.columns:
        reconciled_debit_claims.update(note_d["claim_id"].dropna().astype(str))
    
    # Add ReconTag matches to credit reconciled claims
    if not tagged_sum_by_claim.empty:
        reconciled_credit_claims.update(tagged_sum_by_claim.index.astype(str))

    # Combined unmatched (no corr id, exclude reconciled items by type)
    # Build expected overpay debits that we predicted but couldn't find a matching Chase debit
    expected_overpay_missing = None
    if isinstance(c_match, pd.DataFrame) and not c_match.empty and "overpay_amount" in c_match.columns:
        exp = c_match[(c_match["overpay_amount"].fillna(0) > AMOUNT_TOL) & (c_match["overpay_debit_date"].isna())].copy()
        if not exp.empty:
            expected_overpay_missing = exp.rename(columns={
                "ref_date": "date",
                "overpay_amount": "amount",
                "notes": "notes",
            })
            # Ensure necessary columns
            for col in ["claim_id","claimant","date","amount"]:
                if col not in expected_overpay_missing.columns:
                    expected_overpay_missing[col] = pd.NA
            expected_overpay_missing["category"] = "Expected_Overpay_Debit (missing)"
            expected_overpay_missing["description"] = "Overpayment debit expected but not found in Chase"
            expected_overpay_missing = expected_overpay_missing[["category","claim_id","claimant","date","amount","description","notes"]]

    combined = build_unmatched_combined(
        credits_unmatched_final, debits_orphans_final, c_un_claims, d_un,
        reconciled_credit_claims, reconciled_debit_claims,
        expected_overpay_missing=expected_overpay_missing
    )

    # Create Bank Revenue Summary
    bank_revenue_summary = pd.DataFrame()
    if not per_claim_rev.empty and "Bank-based Revenue" in per_claim_rev.columns:
        total_bank_revenue = per_claim_rev["Bank-based Revenue"].fillna(0).sum()
        bank_revenue_summary = pd.DataFrame([{
            "Total_Bank_Based_Revenue": round(total_bank_revenue, 2),
            "Count_of_Claims": len(per_claim_rev),
            "Generated_Date": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
        }])

    # Create Balance Analysis Template (data only, formulas added separately)
    balance_analysis_data = [
        ["Chase_1160_Balance", 0.00, "Enter your current Chase 1160 account balance"],
        ["Exclusion_1", 0.00, "Amount to exclude from this reconciliation period"],
        ["Exclusion_2", 0.00, "Amount to exclude from this reconciliation period"],
        ["Exclusion_3", 0.00, "Amount to exclude from this reconciliation period"],
        ["Exclusion_4", 0.00, "Amount to exclude from this reconciliation period"],
        ["Additional_Exclusion", 0.00, "Add more exclusions as needed"],
        ["CALCULATION_Net_Balance", None, "Net Balance after exclusions (Chase Balance - Total Exclusions)"],
        ["Bank_Revenue_Actual", round(total_bank_revenue, 2) if 'total_bank_revenue' in locals() else 0, "Actual bank revenue from reconciliation"],
        ["Delta_Variance", None, "Difference between net balance and bank revenue (should be close to 0)"],
        ["Transfer_Amount_Final", round(total_bank_revenue, 2) if 'total_bank_revenue' in locals() else 0, "FINAL: Amount that can be safely transferred out of 1160 account"]
    ]

    # Write
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as w:
        # Get the workbook and add formats
        workbook = w.book
        currency_format = workbook.add_format({'num_format': '$#,##0.00'})
        header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC'})
        
        # Write priority sheets in user-requested order
        # Sheet 1: Per_Claim_Revenue (primary business view)
        per_claim_rev.to_excel(w, "Per_Claim_Revenue", index=False)

        # Sheet 2: Balance Analysis sheet with proper Excel formulas
        worksheet = workbook.add_worksheet("Balance_Analysis")
        
        # Add headers
        worksheet.write('A1', 'Category', header_format)
        worksheet.write('B1', 'Amount', header_format)
        worksheet.write('C1', 'Notes', header_format)
        
        # Normalize optional inputs
        try:
            pre_balance_val = float(pre_balance) if pre_balance is not None else None
        except Exception:
            pre_balance_val = None
        pre_exclusions_list = None
        if isinstance(pre_exclusions, (list, tuple)):
            # force positive values and coerce to float
            tmp = []
            for v in pre_exclusions:
                try:
                    tmp.append(abs(float(v)))
                except Exception:
                    continue
            pre_exclusions_list = tmp if tmp else None

        # Map categories to exclusion indices
        excl_index_map = {
            "Exclusion_1": 0,
            "Exclusion_2": 1,
            "Exclusion_3": 2,
            "Exclusion_4": 3,
            "Additional_Exclusion": 4,
        }

        # Write data rows
        for row_idx, (category, amount, notes) in enumerate(balance_analysis_data, start=2):
            worksheet.write(f'A{row_idx}', category)
            
            # Handle formulas vs regular values
            if category == "CALCULATION_Net_Balance":
                worksheet.write_formula(f'B{row_idx}', '=B2-SUM(B3:B7)', currency_format)
            elif category == "Delta_Variance":
                worksheet.write_formula(f'B{row_idx}', '=B8-B9', currency_format)
            else:
                # Override with provided inputs when available
                if category == "Chase_1160_Balance" and pre_balance_val is not None:
                    worksheet.write(f'B{row_idx}', pre_balance_val, currency_format)
                elif category in excl_index_map and pre_exclusions_list is not None:
                    idx = excl_index_map[category]
                    if idx < len(pre_exclusions_list):
                        worksheet.write(f'B{row_idx}', pre_exclusions_list[idx], currency_format)
                    else:
                        worksheet.write(f'B{row_idx}', amount, currency_format)
                else:
                    worksheet.write(f'B{row_idx}', amount, currency_format)
            
            worksheet.write(f'C{row_idx}', notes)
        
        # Auto-size columns
        worksheet.set_column('A:A', 25)
        worksheet.set_column('B:B', 15)
        worksheet.set_column('C:C', 50)

        # Sheet 3: Unmatched_Combined (exceptions requiring attention)
        combined.to_excel(w, "Unmatched_Combined", index=False)

        # Sheet 4: Summary (overview statistics)
        summary.to_excel(w, "Summary", index=False)

        # Sheet 5: ReconTags (matched and suggested ReconTags)
        if not missing_recontags.empty:
            missing_recontags.to_excel(w, "ReconTags", index=False)
            matched_count = len(missing_recontags[missing_recontags.get("status", "") == "MATCHED"]) if "status" in missing_recontags.columns else 0
            suggested_count = len(missing_recontags[missing_recontags.get("status", "") == "SUGGESTED"]) if "status" in missing_recontags.columns else 0
            if matched_count > 0:
                logger.info(f"✓ {matched_count} ReconTags matched")
            if suggested_count > 0:
                logger.warning(f"⚠️  {suggested_count} claims need ReconTags - see ReconTags sheet")

    logger.info(f"✓ Reconciliation complete. Output written to: {out_path}")

# ------------------------- CLI -------------------------

if __name__ == "__main__":
    ap = argparse.ArgumentParser(
        description="Reconcile repayments against Chase bank transactions",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python3 r2d_recon.py --file data.xlsx
  python3 r2d_recon.py --file data.xlsx --out output.xlsx --ignore-debits-before 2025-01-01
        """
    )
    ap.add_argument("--file", required=True, help="Path to input Excel file")
    ap.add_argument("--r2d-sheet", default="Repayments to Date", help="Name of repayments sheet")
    ap.add_argument("--chase-sheet", default="Chase", help="Name of Chase transactions sheet")
    default_out = f"/Users/Logan/Downloads/Repayments_to_Date_recon-{date.today().isoformat()}.xlsx"
    ap.add_argument("--out", default=default_out, help="Output file path (default: auto-dated)")
    ap.add_argument("--ignore-debits-before", default=None, help="YYYY-MM-DD: exclude unmatched CHASE debits before this date")
    ap.add_argument("-v", "--verbose", action="store_true", help="Enable verbose logging")

    args = ap.parse_args()

    # Set logging level
    if args.verbose:
        logger.setLevel(logging.DEBUG)

    try:
        run(args.file, args.r2d_sheet, args.chase_sheet, args.out, args.ignore_debits_before)
        sys.exit(0)
    except FileNotFoundError as e:
        logger.error(f"File not found: {e}")
        sys.exit(1)
    except ValueError as e:
        logger.error(f"Validation error: {e}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Unexpected error: {e}")
        logger.exception("Full traceback:")
        sys.exit(1)
