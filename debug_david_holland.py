#!/usr/bin/env python3
import pandas as pd
from r2d_recon import load_r2d, load_chase, match_credits_one_row_per_claim, dedupe_by_ach_id

def debug_david_holland():
    """Debug why David Holland isn't getting his credit match"""

    # Load data
    file_path = "/Users/Logan/Downloads/Repayments to Date 09.01.2025 to 09.21.2025.xlsx"
    r2d = load_r2d(file_path, "Repayments to Date")
    chase = load_chase(file_path, "Chase")

    # Apply dedup
    r2d_dedup, _ = dedupe_by_ach_id(r2d)

    # Filter to just David Holland
    holland = r2d_dedup[r2d_dedup['claimant'].str.contains('David Holland', case=False, na=False)].copy()

    print("=== DAVID HOLLAND CREDIT MATCHING DEBUG ===")
    print(f"David Holland entries found: {len(holland)}")

    if not holland.empty:
        holland_row = holland.iloc[0]
        print(f"Window Date: {holland_row['window_date']}")
        print(f"Repayment Amount: ${holland_row['repayment_amount']:.2f}")
        print(f"Amount to Funder: ${holland_row['amount_to_funder']:.2f}")

        # Check what credits are available around his amount
        credits = chase[chase['is_credit']].copy()
        target_amount = holland_row['repayment_amount']

        matching_credits = credits[abs(credits['amount'] - target_amount) <= 0.01]
        print(f"\nCredits matching ${target_amount:.2f}: {len(matching_credits)}")

        # Test the matching function with just David Holland
        print(f"\n=== RUNNING CREDIT MATCHING FOR DAVID HOLLAND ===")
        try:
            results = match_credits_one_row_per_claim(holland, chase)
            credit_matches = results[0]
            unmatched_df = results[1]

            print(f"Credit matches found: {len(credit_matches)}")
            print(f"Unmatched claims: {len(unmatched_df)}")

            if len(credit_matches) > 0:
                print("\nMatched:")
                for _, match in credit_matches.iterrows():
                    print(f"  {match['claimant']}: ${match['repayment_sum']:.2f} -> ${match['chase_credit_amount']:.2f}")

            if len(unmatched_df) > 0:
                print("\nUnmatched:")
                for _, unmatched in unmatched_df.iterrows():
                    print(f"  {unmatched['claimant']}: ${unmatched['repayment_sum']:.2f}")

        except Exception as e:
            print(f"Error during matching: {e}")
            import traceback
            traceback.print_exc()

if __name__ == "__main__":
    debug_david_holland()