#!/usr/bin/env python3
import pandas as pd
from r2d_recon import load_r2d, load_chase, match_credits_one_row_per_claim, dedupe_by_ach_id

def debug_credit_matching():
    """Debug why the 5 priority claims aren't getting credit matches"""

    # Load data
    file_path = "/Users/Logan/Downloads/Repayments to Date 09.01.2025 to 09.21.2025.xlsx"
    r2d = load_r2d(file_path, "Repayments to Date")
    chase = load_chase(file_path, "Chase")

    # Apply dedup
    r2d_dedup, _ = dedupe_by_ach_id(r2d)

    # Filter to just our priority claims
    priority_names = ['Nina Brown', 'Jamie Bagwell', 'Levi Hoerner', 'Evelyn Gaines', 'Raymundo Ramirez Villanueva']
    r2d_priority = r2d_dedup[r2d_dedup['claimant'].isin(priority_names)].copy()

    print("=== PRIORITY CLAIMS CREDIT MATCHING DEBUG ===")
    print(f"Priority claims found: {len(r2d_priority)}")
    for _, row in r2d_priority.iterrows():
        print(f"  {row['claimant']}: ${row['repayment_amount']:.2f} repayment, ${row['amount_to_funder']:.2f} to funder")

    print("\n=== CHASE CREDITS AROUND $62 ===")
    credits = chase[chase["is_credit"]].copy()
    target_credit = 62.14
    tolerance = 0.05

    # Look for credits around $62.14
    potential_credits = credits[abs(credits['amount'] - target_credit) <= tolerance]
    print(f"Credits within ${tolerance} of ${target_credit}: {len(potential_credits)}")

    for _, credit in potential_credits.iterrows():
        print(f"  ${credit['amount']:.2f} on {credit['posting_date']}")

    # Also look for credits around $2312.93 (the full amounts)
    print(f"\n=== CHASE CREDITS AROUND $2312.93 ===")
    full_amount = 2312.93
    full_credits = credits[abs(credits['amount'] - full_amount) <= 0.05]
    print(f"Credits within $0.05 of ${full_amount}: {len(full_credits)}")

    for _, credit in full_credits.iterrows():
        print(f"  Index {credit.name}: ${credit['amount']:.2f} on {credit['posting_date']}")

    # Check date filtering for each priority claim
    print(f"\n=== DATE FILTERING CHECK ===")
    for _, claim in r2d_priority.iterrows():
        window_date = claim['window_date']
        DATE_WINDOW_DAYS = 5

        # Apply date filtering
        date_filtered = full_credits[
            (full_credits['posting_date'] >= window_date - pd.Timedelta(days=DATE_WINDOW_DAYS)) &
            (full_credits['posting_date'] <= window_date + pd.Timedelta(days=DATE_WINDOW_DAYS))
        ]

        print(f"{claim['claimant']} (window: {window_date}): {len(date_filtered)} candidates after date filter")
        if len(date_filtered) > 0:
            for _, cred in date_filtered.iterrows():
                days_diff = abs((cred['posting_date'] - window_date).days)
                print(f"  - ${cred['amount']:.2f} on {cred['posting_date']} (±{days_diff} days)")

    # Now test the matching function
    print(f"\n=== RUNNING CREDIT MATCHING ===")
    try:
        results = match_credits_one_row_per_claim(r2d_priority, chase)
        credit_matches = results[0]
        unmatched_df = results[1]

        print(f"Credit matches found: {len(credit_matches)}")
        print(f"Unmatched claims: {len(unmatched_df)}")

        if len(credit_matches) > 0:
            print("\nMatched claims:")
            for _, match in credit_matches.iterrows():
                print(f"  {match['claimant']}: ${match['repayment_sum']:.2f} -> ${match['chase_credit_amount']:.2f}")

        if len(unmatched_df) > 0:
            print("\nUnmatched claims:")
            for _, unmatched in unmatched_df.iterrows():
                print(f"  {unmatched['claimant']}: ${unmatched['repayment_sum']:.2f}")

    except Exception as e:
        print(f"Error during matching: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_credit_matching()