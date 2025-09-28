#!/usr/bin/env python3
import argparse
import re
from datetime import date
import pandas as pd

# ------------------------- Parameters & Regex -------------------------

DATE_WINDOW_DAYS = 5
OVERPAY_BACKFILL_WINDOW = 7
NOTE_WINDOW_DAYS = 7
AMOUNT_TOL = 0.01  # 1 cent

TRANSFER_HINTS = re.compile(r"(?:dwolla|transfer|ach|orig co name|orig id|trn)", re.I)
OVERPAY_DEBIT_HINTS = re.compile(r"(?:2670|transfer)", re.I)
# liberal: "overpaid by $X" OR "overpayment of $X" OR "overpayment $X"
OVERPAID_REGEX = re.compile(r"(?:overpaid\s*(?:by)?|overpayment\s*(?:of)?)\s*\$?\s*([0-9][0-9,]*\.?[0-9]{0,2})", re.I)
PAREN_SUFFIX = re.compile(r"\s*\([^)]*\)\s*$")
DOLLAR_REGEX = re.compile(r"\$?\s*([0-9][0-9,]*\.\d{2})")
DATE_IN_NOTES = re.compile(r"\b(\d{1,2})/(\d{1,2})\b")
CREDIT_KEYWORDS = re.compile(r"(received|deposit|check|credited|incoming|rec\.?\s*rem|received\s+rem|received\s+remaining|rcvd|remaining\s+repayment|repayment\s+received|remaining\s*bal|rem\.?\s*bal|underpaid\s+by)", re.I)
DEBIT_KEYWORDS  = re.compile(r"(send\s+funder|to\s+funder|transfer|outgoing|ach\s*out|2670)", re.I)
SEND_FUNDER_REGEX = re.compile(r"send\s+funder[^$]*\$([0-9][0-9,]*\.[0-9]{2})", re.I)
RECEIVED_CHECK_REGEX = re.compile(r"(received.*check|rec\.?\s*rem|received\s+rem|received\s+remaining|remaining\s+repayment|repayment\s+received|underpaid\s+by)\D*\$([0-9][0-9,]*\.[0-9]{2})", re.I)
REQUESTED_REMAINING_REGEX = re.compile(r"(req\.?\s*rem\.?|requested\s+rem\.?|req\.?\s*remaining|requested\s+remaining).*?\$([0-9][0-9,]*\.[0-9]{2})", re.I)

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

def colmap(df, wanted_names):
    renorm = {c: str(c).strip() for c in df.columns}
    df = df.rename(columns=renorm)
    low = {c.lower(): c for c in df.columns}
    out = {}
    for key, aliases in wanted_names.items():
        pick = None
        for a in aliases:
            if a and a.lower() in low:
                pick = low[a.lower()]
                break
        out[key] = pick
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
    df = pd.read_excel(path, sheet_name=sheet)
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
    m = colmap(df, wanted)
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
    df = pd.read_excel(path, sheet_name=sheet)
    wanted = {
        "posting_date":["Posting Date","Details Posting Date","Post Date"],
        "description":["Description","Details","Memo"],
        "amount":["Amount","Amt"],
        "type":["Type"],
        "recon_tag":["ReconTag","Recon Tag","Recon_Tag","RECONTAG","recontag"],
    }
    m = colmap(df, wanted)
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

def dedupe_by_ach_id(r2d: pd.DataFrame):
    with_id = r2d[r2d["ach_id"].astype(str).str.len() > 0]
    without_id = r2d[r2d["ach_id"].astype(str).str.len() == 0]
    kept = with_id.drop_duplicates(subset=["ach_id"], keep="first").copy()
    return pd.concat([kept, without_id], ignore_index=True), int(len(with_id) - len(kept))

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
    r2d_dedup, dup_removed = dedupe_by_ach_id(r2d)
    debits = chase[chase["is_debit"]].copy()
    results, unmatched_idx, used = [], [], set()

    # First pass: exact date matches only (within standard window)
    for i, row in r2d_dedup.iterrows():
        amt = row.get("amount_transferred")
        win = row.get("window_date")
        if amt is None or pd.isna(amt) or pd.isna(win):
            continue

        amt_rounded = r2(amt)
        cand = debits[debits["amount"].abs().sub(abs(amt_rounded)).abs() <= AMOUNT_TOL].copy()
        cand = cand[(cand["posting_date"] >= win - pd.Timedelta(days=DATE_WINDOW_DAYS)) &
                    (cand["posting_date"] <= win + pd.Timedelta(days=DATE_WINDOW_DAYS))]
        if cand.empty:
            continue

        cand = cand.assign(date_delta=(cand["posting_date"]-win).abs().dt.days)
        cand = cand.sort_values(["has_hint","date_delta"], ascending=[False, True])
        chosen = None
        for idx, c in cand.iterrows():
            if idx not in used:
                chosen = (idx, c); break
        if not chosen:
            continue

        idx, c = chosen; used.add(idx)
        confidence = 0.5 + (0.3 if c["has_hint"] else 0) + (0.2 if abs((c["posting_date"]-win).days)<=1 else 0)
        results.append({
            "ach_id": row.get("ach_id"),
            "claim_id": row.get("claim_id"),
            "amount_transferred": amt_rounded,
            "r2d_date": win,
            "chase_date": c["posting_date"],
            "chase_amount": c["amount"],
            "description": c["description"],
            "match_type": "amount+window(+hints)",
            "confidence": r2(min(confidence, 0.99)),
            "chase_index": idx,
        })

    # Second pass: wider window for unmatched claims
    for i, row in r2d_dedup.iterrows():
        if any(r["claim_id"] == row.get("claim_id") for r in results):  # Already matched
            continue
            
        amt = row.get("amount_transferred")
        win = row.get("window_date")
        if amt is None or pd.isna(amt) or pd.isna(win):
            unmatched_idx.append(i); continue

        amt_rounded = r2(amt)
        cand = debits[debits["amount"].abs().sub(abs(amt_rounded)).abs() <= AMOUNT_TOL].copy()
        cand = cand[(cand["posting_date"] >= win - pd.Timedelta(days=30)) &
                    (cand["posting_date"] <= win + pd.Timedelta(days=30))]
        if cand.empty:
            unmatched_idx.append(i); continue

        cand = cand.assign(date_delta=(cand["posting_date"]-win).abs().dt.days)
        cand = cand.sort_values(["has_hint","date_delta"], ascending=[False, True])
        chosen = None
        for idx, c in cand.iterrows():
            if idx not in used:
                chosen = (idx, c); break
        if not chosen:
            unmatched_idx.append(i); continue

        idx, c = chosen; used.add(idx)
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
            "confidence": r2(min(confidence, 0.99)),
            "chase_index": idx,
        })

    used_debit_idx = [r["chase_index"] for r in results]
    return pd.DataFrame(results), r2d_dedup.loc[unmatched_idx].copy(), debits.loc[~debits.index.isin(used_debit_idx)].copy(), dup_removed, used_debit_idx

# ------------------------- Credit Matching (Parent Claim) -------------------------

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
    credits = chase[chase["is_credit"]].copy()
    debits = chase[chase["is_debit"]].copy()

    tmp = r2d.copy()
    tmp["repayment_amount"] = pd.to_numeric(tmp["repayment_amount"], errors="coerce")
    tmp["amount_to_funder"] = pd.to_numeric(tmp["amount_to_funder"], errors="coerce")
    tmp["overpaid_val"] = tmp["notes"].apply(parse_overpaid_amount)

    parents = tmp.groupby("claim_id", dropna=False).apply(canonical_parent).reset_index(drop=True)
    roll = tmp.groupby("claim_id", dropna=False).agg(
        repayment_sum=("repayment_amount","sum"),
        amount_to_funder_sum=("amount_to_funder","sum"),
        ref_date=("window_date","max"),
        overpaid_sum=("overpaid_val","max"),
        notes_any=("notes"," | ".join)
    ).reset_index()
    roll = roll.merge(parents[["claim_id","claimant","deal_type","contract_date"]], on="claim_id", how="left")

    results, used_credit, used_overpay_debit, unmatched_claims = [], set(), set(), []

    for _, r in roll.iterrows():
        claim_id = r["claim_id"]
        claim_sum = r2(r.get("repayment_sum") or 0.0)
        if not claim_sum or claim_sum == 0:
            unmatched_claims.append(claim_id); continue

        ref_date = r["ref_date"]
        over_x = r2(r.get("overpaid_sum") or 0.0) if pd.notna(r.get("overpaid_sum")) else 0.0

        # Wider window when we have an overpay (checks lag more)
        win_days = max(DATE_WINDOW_DAYS, OVERPAY_BACKFILL_WINDOW) if (over_x or 0.0) > AMOUNT_TOL else DATE_WINDOW_DAYS

        cand = credits.copy()
        if pd.notna(ref_date):
            cand = cand[(cand["posting_date"] >= ref_date - pd.Timedelta(days=win_days)) &
                        (cand["posting_date"] <= ref_date + pd.Timedelta(days=win_days))]
        cand = cand.assign(
            diff_claim=(cand["amount"] - claim_sum).abs(),
            diff_claim_plus_over=(cand["amount"] - (claim_sum + (over_x or 0.0))).abs(),
            date_delta=(cand["posting_date"] - ref_date).abs().dt.days if pd.notna(ref_date) else 0
        )

        chosen = None; match_mode = None

        # Try claim_sum + overpay FIRST when overpay exists
        if (over_x or 0.0) > AMOUNT_TOL:
            for idx, c in cand.sort_values(["diff_claim_plus_over","date_delta","posting_date"]).iterrows():
                if idx in used_credit: 
                    continue
                if abs(c["amount"] - (claim_sum + over_x)) <= AMOUNT_TOL:
                    chosen = (idx, c); match_mode = "claim_sum_plus_overpay"; break

        # Fallback to exact claim_sum
        if not chosen:
            for idx, c in cand.sort_values(["diff_claim","date_delta","posting_date"]).iterrows():
                if idx in used_credit:
                    continue
                if abs(c["amount"] - claim_sum) <= AMOUNT_TOL:
                    chosen = (idx, c); match_mode = "claim_sum"; break

        over_debit_date = None; over_debit_desc = None
        if chosen:
            idx, c = chosen; used_credit.add(idx)
            if match_mode == "claim_sum_plus_overpay" and (over_x or 0.0) > AMOUNT_TOL:
                dwin = debits[(debits["amount"].abs().sub(over_x).abs() <= AMOUNT_TOL)].copy()
                # Prefer debits near the credit date
                dnear_credit = dwin[(dwin["posting_date"] >= c["posting_date"] - pd.Timedelta(days=DATE_WINDOW_DAYS)) &
                                    (dwin["posting_date"] <= c["posting_date"] + pd.Timedelta(days=DATE_WINDOW_DAYS))]
                dnear_credit = dnear_credit.sort_values(["overpay_hint","posting_date"], ascending=[False, True])
                for didx, d in dnear_credit.iterrows():
                    if didx in used_overpay_debit: continue
                    over_debit_date = d["posting_date"]; over_debit_desc = d["description"]
                    used_overpay_debit.add(didx); break
                # Fallback near ref date
                if over_debit_date is None and pd.notna(ref_date):
                    dnear_ref = dwin[(dwin["posting_date"] >= ref_date - pd.Timedelta(days=OVERPAY_BACKFILL_WINDOW)) &
                                     (dwin["posting_date"] <= ref_date + pd.Timedelta(days=OVERPAY_BACKFILL_WINDOW))]
                    dnear_ref = dnear_ref.sort_values(["overpay_hint","posting_date"], ascending=[False, True])
                    for didx, d in dnear_ref.iterrows():
                        if didx in used_overpay_debit: continue
                        over_debit_date = d["posting_date"]; over_debit_desc = d["description"]
                        used_overpay_debit.add(didx); break

            results.append({
                "claim_id": claim_id,
                "claimant": r["claimant"],
                "deal_type_parent": r["deal_type"],
                "contract_date_parent": r["contract_date"],
                "repayment_sum": claim_sum,
                "amount_to_funder_sum": r.get("amount_to_funder_sum"),
                "ref_date": ref_date,
                "match_type": match_mode,
                "chase_credit_date": c["posting_date"],
                "chase_credit_amount": c["amount"],
                "overpay_amount": (over_x or 0.0) if match_mode == "claim_sum_plus_overpay" else 0.0,
                "overpay_debit_date": over_debit_date,
                "overpay_debit_desc": over_debit_desc,
                "notes": r["notes_any"],
            })
        else:
            unmatched_claims.append(claim_id)

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

    # Collect amounts to ignore (requested remaining amounts)
    amounts_to_ignore = set()
    for m in REQUESTED_REMAINING_REGEX.finditer(text):
        amt = r2(m.group(2).replace(",",""))
        if amt is not None:
            amounts_to_ignore.add(amt)

    for m in SEND_FUNDER_REGEX.finditer(text):
        amt = r2(m.group(1).replace(",",""))
        if amt is not None and amt not in amounts_to_ignore:
            events.append(("debit_expected", amt, anchor))
    for m in RECEIVED_CHECK_REGEX.finditer(text):
        amt = r2(m.group(2).replace(",",""))
        if amt is not None and amt not in amounts_to_ignore:
            events.append(("credit_expected", amt, anchor))

    for m in DOLLAR_REGEX.finditer(text):
        amt = r2(m.group(1).replace(",",""))
        if amt in amounts_to_ignore:
            continue
        start, end = max(0, m.start()-120), min(len(text), m.end()+120)
        ctx = text[start:end]
        # Additional check: skip if this dollar amount is in a "requested remaining" context
        if REQUESTED_REMAINING_REGEX.search(ctx):
            continue
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
    credits = chase[chase["is_credit"]].copy()
    debits  = chase[chase["is_debit"]].copy()

    grp = r2d.groupby("claim_id", dropna=False)
    notes_join = grp["notes"].apply(lambda s: " | ".join([str(x) for x in s.dropna() if str(x).strip()]))
    ref_date = grp["window_date"].max()
    claimant = grp["claimant"].max()
    roll = pd.DataFrame({"claim_id": ref_date.index, "ref_date": ref_date.values, "notes_any": notes_join.values, "claimant": claimant.values})
    roll["claimant_display"] = roll["claimant"].astype(str).str.replace(PAREN_SUFFIX, "", regex=True)

    note_credit_rows, note_debit_rows = [], []
    newly_used_credit, newly_used_debit = set(), set()

    for _, r in roll.iterrows():
        events = extract_note_events(r["notes_any"], r["ref_date"])
        for kind, amount, anchor_date in events:
            if kind == "credit_expected":
                # Check if this is a shared check
                is_shared, client_count, other_client_names = detect_shared_check(r["notes_any"])
                
                cnd = credits.loc[~credits.index.isin(used_credit_idx + list(newly_used_credit))].copy()
                if pd.notna(anchor_date):
                    cnd = cnd[(cnd["posting_date"] >= anchor_date - pd.Timedelta(days=NOTE_WINDOW_DAYS)) &
                              (cnd["posting_date"] <= anchor_date + pd.Timedelta(days=NOTE_WINDOW_DAYS))]
                cnd = cnd[(cnd["amount"].sub(amount).abs() <= AMOUNT_TOL)]
                chosen = None
                if not cnd.empty:
                    chosen = cnd.sort_values("posting_date").iloc[0]
                else:
                    rd = r["ref_date"]
                    if pd.notna(rd):
                        cnd = credits.loc[~credits.index.isin(used_credit_idx + list(newly_used_credit))].copy()
                        cnd = cnd[(cnd["posting_date"] >= rd - pd.Timedelta(days=NOTE_WINDOW_DAYS+3)) &
                                  (cnd["posting_date"] <= rd + pd.Timedelta(days=NOTE_WINDOW_DAYS+3))]
                        cnd = cnd[(cnd["amount"].sub(amount).abs() <= AMOUNT_TOL)]
                        if not cnd.empty:
                            chosen = cnd.sort_values("posting_date").iloc[0]
                
                if chosen is not None:
                    newly_used_credit.add(chosen.name)
                    
                    if is_shared and client_count > 1:
                        # For shared checks, split the amount proportionally
                        # For now, split equally among clients (can be enhanced later to use repayment amounts)
                        split_amount = chosen["amount"] / client_count
                        
                        # Add entry for current client
                        note_credit_rows.append({
                            "claim_id": r["claim_id"],
                            "claimant": r["claimant_display"],
                            "note_amount": amount,
                            "matched_credit_date": chosen["posting_date"],
                            "matched_credit_amount": split_amount,
                            "matched_credit_desc": chosen["description"] + f" (shared {client_count} ways)",
                            "source": "notes_shared"
                        })
                        
                        # Try to find and add entries for other clients mentioned in notes
                        for other_client in other_client_names:
                            # Try to find the claim_id for the other client
                            other_client_clean = other_client.strip()
                            matching_claims = roll[roll["claimant_display"].str.contains(other_client_clean, case=False, na=False)]
                            
                            if not matching_claims.empty:
                                other_claim = matching_claims.iloc[0]
                                note_credit_rows.append({
                                    "claim_id": other_claim["claim_id"],
                                    "claimant": other_claim["claimant_display"],
                                    "note_amount": amount,
                                    "matched_credit_date": chosen["posting_date"],
                                    "matched_credit_amount": split_amount,
                                    "matched_credit_desc": chosen["description"] + f" (shared {client_count} ways)",
                                    "source": "notes_shared"
                                })
                            else:
                                # Create placeholder entry for unknown other client
                                note_credit_rows.append({
                                    "claim_id": f"UNKNOWN_{other_client_clean}",
                                    "claimant": other_client_clean,
                                    "note_amount": amount,
                                    "matched_credit_date": chosen["posting_date"],
                                    "matched_credit_amount": split_amount,
                                    "matched_credit_desc": chosen["description"] + f" (shared {client_count} ways)",
                                    "source": "notes_shared_unknown"
                                })
                    else:
                        # Regular (non-shared) credit matching
                        note_credit_rows.append({
                            "claim_id": r["claim_id"],
                            "claimant": r["claimant_display"],
                            "note_amount": amount,
                            "matched_credit_date": chosen["posting_date"],
                            "matched_credit_amount": chosen["amount"],
                            "matched_credit_desc": chosen["description"],
                            "source": "notes"
                        })

            elif kind == "debit_expected":
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
                        dnd = dnd[(dnd["posting_date"] >= rd - pd.Timedelta(days=NOTE_WINDOW_DAYS+3)) &
                                  (dnd["posting_date"] <= rd + pd.Timedelta(days=NOTE_WINDOW_DAYS+3))]
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

    note_credit_df = pd.DataFrame(note_credit_rows)
    note_debit_df  = pd.DataFrame(note_debit_rows)
    return note_credit_df, note_debit_df, list(newly_used_credit), list(newly_used_debit)

# ------------------------- Revenue & Summary -------------------------

def compute_bank_revenue_per_claim(d_match, c_match, note_c, note_d, per_claim):
    if not c_match.empty:
        eff = c_match.copy()
        eff["effective_credit"] = eff.apply(
            lambda r: r["repayment_sum"] if str(r.get("match_type","")) == "claim_sum_plus_overpay" else r["chase_credit_amount"],
            axis=1
        )
        cm_per_claim = eff.groupby("claim_id", dropna=False)["effective_credit"].sum()
    else:
        cm_per_claim = pd.Series(dtype=float)

    note_c_per = (note_c.groupby("claim_id", dropna=False)["matched_credit_amount"].sum()
                  if not note_c.empty else pd.Series(dtype=float))
    dm_per_claim = (d_match.groupby("claim_id", dropna=False)["chase_amount"].apply(lambda s: s.abs().sum())
                    if not d_match.empty else pd.Series(dtype=float))
    note_d_per = (note_d.groupby("claim_id", dropna=False)["matched_debit_amount"].apply(lambda s: s.abs().sum())
                  if not note_d.empty else pd.Series(dtype=float))

    out = per_claim.copy()
    out["Claimant"] = out.get("claimant", out.get("Claimant", ""))
    out["Claimant"] = out["Claimant"].astype(str).str.replace(PAREN_SUFFIX, "", regex=True)

    out["Bank Credits (Effective)"] = out["claim_id"].map(cm_per_claim).fillna(0).round(2) + \
                                      out["claim_id"].map(note_c_per).fillna(0).round(2)
    out["Bank Funder Debits"] = out["claim_id"].map(dm_per_claim).fillna(0).round(2) + \
                                out["claim_id"].map(note_d_per).fillna(0).round(2)
    out["Bank-based Revenue"] = (out["Bank Credits (Effective)"] - out["Bank Funder Debits"]).round(2)
    out["Book Revenue (KEEP)"] = (out["Repayment Sum"].fillna(0) - out["Amount To Funder Sum"].fillna(0)).round(2)
    out["Check (Bank - Book)"] = (out["Bank-based Revenue"] - out["Book Revenue (KEEP)"]).round(2)

    cols = ["claim_id","Claimant","Repayment Sum","Amount To Funder Sum","Book Revenue (KEEP)",
            "Bank Credits (Effective)","Bank Funder Debits","Bank-based Revenue","Check (Bank - Book)"]
    existing = [c for c in cols if c in out.columns]
    remaining = [c for c in out.columns if c not in existing]
    return out[existing + remaining]

def build_summary(d_match, d_un, d_orph, c_match, note_c, note_d):
    total_d = r2(d_match["chase_amount"].abs().sum(), 2) if not d_match.empty else 0.0
    total_c = r2(c_match["chase_credit_amount"].sum(), 2) if not c_match.empty else 0.0
    total_c_notes = r2(note_c["matched_credit_amount"].sum(), 2) if not note_c.empty else 0.0
    total_d_notes = r2(note_d["matched_debit_amount"].abs().sum(), 2) if not note_d.empty else 0.0
    total_d_all = r2((total_d or 0) + (total_d_notes or 0), 2)
    total_c_all = r2((total_c or 0) + (total_c_notes or 0), 2)
    net_after_notes = r2((total_d_all or 0) - (total_c_all or 0), 2)

    return pd.DataFrame({"metric":[
        "Debits matched (count)",
        "Credits matched (count)",
        "Note-derived debit matches (count)",
        "Note-derived credit matches (count)",
        "Total debits matched incl. notes (abs)",
        "Total credits matched incl. notes",
        "Net diff after notes (debits - credits)",
        "Debits unmatched (count)",
        "CHASE unmatched debits (count)"
    ], "value":[
        len(d_match), len(c_match),
        (0 if note_d.empty else len(note_d)),
        (0 if note_c.empty else len(note_c)),
        total_d_all, total_c_all, net_after_notes,
        len(d_un), len(d_orph)
    ]})

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
            # Use transfer_initiated as date, amount_transferred as amount
            rename_map = {"transfer_initiated":"date","amount_transferred":"amount","notes":"notes","claimant":"claimant","claim_id":"claim_id"}
            d_unmatched = d_unmatched.rename(columns=rename_map)
            if "date" not in d_unmatched.columns and "transfer_initiated" in d_unmatched.columns:
                d_unmatched["date"] = d_unmatched["transfer_initiated"]
            if "amount" not in d_unmatched.columns and "amount_transferred" in d_unmatched.columns:
                d_unmatched["amount"] = d_unmatched["amount_transferred"]
            d_unmatched = d_unmatched.assign(category="R2D_Unmatched_Debit (transfer)")
            d_unmatched["description"] = pd.NA
            d_unmatched = d_unmatched[["category","claim_id","claimant","date","amount","description","notes"]]
        else:
            d_unmatched = None
            
    cu_claims = c_un_claims.copy()
    if not cu_claims.empty:
        # Only exclude unmatched credit claims if they were specifically matched as credits
        if reconciled_credit_claims:
            cu_claims = cu_claims.loc[~cu_claims["claim_id"].astype(str).isin(reconciled_credit_claims)].copy()
            
        if not cu_claims.empty:
            rename_map = {"ref_date":"date","notes_any":"notes","repayment_sum":"amount","claimant":"claimant","claim_id":"claim_id"}
            if "Repayment Sum" in cu_claims.columns: rename_map["repayment_sum"] = "Repayment Sum"
            if "notes" in cu_claims.columns: rename_map["notes_any"] = "notes"
            cu_claims = cu_claims.rename(columns=rename_map)
            if "date" not in cu_claims.columns and "ref_date" in cu_claims.columns:
                cu_claims["date"] = cu_claims["ref_date"]
            if "amount" not in cu_claims.columns and "Repayment Sum" in cu_claims.columns:
                cu_claims["amount"] = cu_claims["Repayment Sum"]
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
    # Load
    r2d   = load_r2d(file_path, r2d_sheet)
    chase = load_chase(file_path, chase_sheet)

    # Correlation map
    corr_map = build_correlation_map(r2d)

    # Debits
    d_match, d_un, d_orph, dup, used_debit_idx = match_debits_relaxed(r2d, chase)

    # Credits (+ overpay)
    c_match, c_un_claims, c_un_chase_final, per_claim, over_adj, used_credit_idx, used_overpay_debit = \
        match_credits_one_row_per_claim(r2d, chase)

    # Remove overpay-linked debits from orphans
    if used_overpay_debit:
        d_orph = d_orph.loc[~d_orph.index.isin(used_overpay_debit)].copy()

    # Notes pass
    note_c, note_d, newly_used_credit, newly_used_debit = \
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
    if "recon_tag" in chase.columns:
        nonempty_tag = chase["recon_tag"].notna() & chase["recon_tag"].astype(str).str.strip().ne("")
        tagged = chase.loc[chase["is_credit"] & nonempty_tag].copy()

        # Remove credits already matched by algo or notes
        already_used_credit_idx = set(used_credit_idx or []) | set(newly_used_credit or [])
        if not tagged.empty and already_used_credit_idx:
            tagged = tagged.loc[~tagged.index.isin(already_used_credit_idx)]

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

    # Summary & base per-claim revenue
    summary = build_summary(d_match, d_un, debits_orphans_final, c_match, note_c, note_d)
    per_claim_rev = compute_bank_revenue_per_claim(d_match, c_match, note_c, note_d, per_claim)

    # Add ReconTagged credits to Bank Credits (Effective)
    if not tagged_sum_by_claim.empty:
        per_claim_rev["claim_id"] = per_claim_rev["claim_id"].astype(str).str.strip()

        # Only add ReconTag credits for claims that don't already have credit matches
        claims_with_credit_matches = set()
        if not c_match.empty:
            claims_with_credit_matches.update(c_match["claim_id"].dropna().astype(str))
        if not note_c.empty:
            claims_with_credit_matches.update(note_c["claim_id"].dropna().astype(str))

        # Filter out ReconTag claims that already have credits
        filtered_recon_tags = tagged_sum_by_claim.copy()
        for claim_id in claims_with_credit_matches:
            if claim_id in filtered_recon_tags.index:
                print(f"Skipping ReconTag for {claim_id} - already has credit match")
                filtered_recon_tags = filtered_recon_tags.drop(claim_id)

        per_claim_rev["Bank Credits (Effective)"] = (
            per_claim_rev["Bank Credits (Effective)"].fillna(0) +
            per_claim_rev["claim_id"].map(filtered_recon_tags).fillna(0)
        ).round(2)
        if "Bank Funder Debits" in per_claim_rev.columns:
            per_claim_rev["Bank-based Revenue"] = (per_claim_rev["Bank Credits (Effective)"] - per_claim_rev["Bank Funder Debits"]).round(2)
        if "Book Revenue (KEEP)" in per_claim_rev.columns:
            per_claim_rev["Check (Bank - Book)"] = (per_claim_rev["Bank-based Revenue"] - per_claim_rev["Book Revenue (KEEP)"]).round(2)
        print(f"ReconTagged credits merged for {tagged_sum_by_claim.index.nunique()} claim(s), total = {float(tagged_sum_by_claim.sum()):.2f}")

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
        
        # Write priority sheets first (as requested by user)
        per_claim_rev.to_excel(w, "Per_Claim_Revenue", index=False)
        combined.to_excel(w, "Unmatched_Combined", index=False)
        
        # Balance Analysis sheet with proper Excel formulas
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
        
        # Write remaining sheets in logical order
        d_match.to_excel(w, "Debit_Matches", index=False)
        d_un.to_excel(w, "Debit_Unmatched", index=False)
        debits_orphans_final.to_excel(w, "CHASE_Unmatched_Debits", index=False)
        c_match.to_excel(w, "Credit_Matches", index=False)
        c_un_claims.to_excel(w, "Claims_Unmatched_Credits", index=False)
        credits_unmatched_final.to_excel(w, "CHASE_Unmatched_Credits", index=False)
        over_adj.to_excel(w, "Overpayment_Adjustments", index=False)
        note_c.to_excel(w, "Note_Matched_Credits", index=False)
        note_d.to_excel(w, "Note_Matched_Debits", index=False)
        bank_revenue_summary.to_excel(w, "Bank_Revenue_Summary", index=False)
        summary.to_excel(w, "Summary", index=False)
        pd.DataFrame([{"duplicates_removed_by_ach_id": dup}]).to_excel(w, "Stats", index=False)
        (excluded_debits_by_date if isinstance(excluded_debits_by_date, pd.DataFrame) else pd.DataFrame()
        ).to_excel(w, "Excluded_Unmatched_Debits", index=False)
        (excluded_dun_by_ti if isinstance(excluded_dun_by_ti, pd.DataFrame) else pd.DataFrame()
        ).to_excel(w, "Excluded_Debit_Unmatched", index=False)

    print(f"Wrote {out_path}")

# ------------------------- CLI -------------------------

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", required=True)
    ap.add_argument("--r2d-sheet", default="Repayments to Date")
    ap.add_argument("--chase-sheet", default="Chase")
    ap.add_argument("--out", default="Repayments_to_Date_recon.xlsx")
    ap.add_argument("--ignore-debits-before", default=None, help="YYYY-MM-DD: exclude unmatched CHASE debits before this date and Debit_Unmatched by Transfer Initiated")
    args = ap.parse_args()
    run(args.file, args.r2d_sheet, args.chase_sheet, args.out, args.ignore_debits_before)
