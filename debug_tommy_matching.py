#!/usr/bin/env python3
import pandas as pd
import re
from datetime import date
from r2d_recon import load_r2d, load_chase, dedupe_by_ach_id, r2

def debug_tommy_matching():
    """Debug Tommy's specific matching issue"""

    # Load data
    file_path = "/Users/Logan/Downloads/Repayments to Date 09.01.2025 to 09.21.2025.xlsx"
    r2d = load_r2d(file_path, "Repayments to Date")
    chase = load_chase(file_path, "Chase")

    print("=== TOMMY MATCHING DEBUG ===")
    print()

    # Step 1: Check if Tommy survives deduplication
    print("1. BEFORE DEDUPLICATION:")
    tommy_entries = r2d[r2d["claimant"].str.contains("Tommy", case=False, na=False)]
    print(f"Tommy entries: {len(tommy_entries)}")
    for _, row in tommy_entries.iterrows():
        print(f"  ACH ID: {row['ach_id']}, Amount: ${row['amount_transferred']:,.2f}, Date: {row['window_date']}")

    # Step 2: After deduplication
    r2d_dedup, dup_removed = dedupe_by_ach_id(r2d)
    print(f"\n2. AFTER DEDUPLICATION (removed {dup_removed} duplicates):")
    tommy_dedup = r2d_dedup[r2d_dedup["claimant"].str.contains("Tommy", case=False, na=False)]
    print(f"Tommy entries: {len(tommy_dedup)}")
    if not tommy_dedup.empty:
        tommy_row = tommy_dedup.iloc[0]
        print(f"  Claim ID: {tommy_row['claim_id']}")
        print(f"  ACH ID: {tommy_row['ach_id']}")
        print(f"  Amount: ${tommy_row['amount_transferred']:,.2f}")
        print(f"  Window Date: {tommy_row['window_date']}")

    # Step 3: Check Chase debit candidates
    print(f"\n3. CHASE DEBIT CANDIDATES:")
    debits = chase[chase["is_debit"]].copy()
    target_amount = 25864.25
    amt_candidates = debits[debits["amount"].abs().sub(target_amount).abs() <= 0.01]
    print(f"Chase debits matching ${target_amount:,.2f}: {len(amt_candidates)}")

    if not amt_candidates.empty:
        for _, debit in amt_candidates.iterrows():
            print(f"  Index: {debit.name}, Amount: ${debit['amount']:,.2f}, Date: {debit['posting_date']}")
            print(f"  Description: {debit['description'][:50]}...")
            print(f"  Has hint: {debit.get('has_hint', False)}")

    # Step 4: Simulate matching for Tommy
    if not tommy_dedup.empty:
        print(f"\n4. MATCHING SIMULATION FOR TOMMY:")
        tommy_row = tommy_dedup.iloc[0]
        amt = tommy_row.get("amount_transferred")
        win = tommy_row.get("window_date")

        print(f"Tommy amount: ${amt:,.2f}")
        print(f"Tommy window date: {win}")

        if amt is not None and pd.notna(amt) and pd.notna(win):
            amt_rounded = r2(amt)

            # Find matching candidates
            cand = debits[debits["amount"].abs().sub(abs(amt_rounded)).abs() <= 0.01].copy()
            print(f"Amount matches: {len(cand)}")

            # Apply date filter
            DATE_WINDOW_DAYS = 5
            date_filtered = cand[(cand["posting_date"] >= win - pd.Timedelta(days=DATE_WINDOW_DAYS)) &
                                (cand["posting_date"] <= win + pd.Timedelta(days=DATE_WINDOW_DAYS))]
            print(f"Date filtered (±{DATE_WINDOW_DAYS} days): {len(date_filtered)}")

            if not date_filtered.empty:
                print("Available candidates:")
                for idx, c in date_filtered.iterrows():
                    date_diff = abs((c["posting_date"] - win).days)
                    print(f"  Index {idx}: ${c['amount']:,.2f}, Date: {c['posting_date']}, Diff: {date_diff} days")
                    print(f"    Has hint: {c.get('has_hint', False)}")

                print("\n✅ Tommy SHOULD match - candidates are available!")

                # Check if any are already used (this might be the issue)
                print("\n5. CHECKING IF CANDIDATES ARE ALREADY USED:")
                # We can't check 'used' set here since it's built during the actual run
                print("This would need to be checked during the actual matching run")

            else:
                print("❌ No candidates after date filtering")
        else:
            print("❌ Missing amount or window date")

if __name__ == "__main__":
    debug_tommy_matching()