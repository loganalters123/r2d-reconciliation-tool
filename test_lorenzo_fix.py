#!/usr/bin/env python3
import pandas as pd
from r2d_recon import extract_note_events

def test_lorenzo_fix():
    """Test that Lorenzo's 'Req Rem' amount gets ignored"""

    # Get Lorenzo's actual notes from the file
    file_path = '/Users/Logan/Downloads/Repayments_to_Date_recon-2025-09-28.xlsx'
    credit_matches = pd.read_excel(file_path, sheet_name='Credit_Matches')
    lorenzo = credit_matches[credit_matches['claimant'].str.contains('Lorenzo', case=False, na=False)]

    if not lorenzo.empty:
        actual_notes = str(lorenzo.iloc[0]['notes'])
        ref_date = pd.Timestamp('2025-09-01')

        print("=== LORENZO FIELDS TEST ===")
        print(f"Notes: {actual_notes}")
        print(f"Expected: 0 credit events (Req Rem should be ignored)")

        # Test our updated function
        events = extract_note_events(actual_notes, ref_date)
        credit_events = [e for e in events if e[0] == 'credit_expected']

        print(f"Actual credit events found: {len(credit_events)}")
        if credit_events:
            for event in credit_events:
                print(f"  - ${event[1]} from {event[0]}")

        passed = len(credit_events) == 0
        print(f"Result: {'PASS' if passed else 'FAIL'}")

        if passed:
            print("✅ Lorenzo's Bank Credits should be $3,444.41 (not $3,657.06)")
        else:
            print("❌ Lorenzo will still get incorrect extra $212.65")

        return passed
    else:
        print("Lorenzo not found in data")
        return False

if __name__ == "__main__":
    test_lorenzo_fix()