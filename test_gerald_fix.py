#!/usr/bin/env python3
import pandas as pd
from r2d_recon import extract_note_events

def test_gerald_fix():
    """Test that Gerald's 'remaining repayment received' gets extracted"""

    # Get Gerald's actual notes from the output file
    file_path = '/Users/Logan/Downloads/Fixed_Recon_Output.xlsx'
    unmatched = pd.read_excel(file_path, sheet_name='Unmatched_Combined')
    gerald = unmatched[unmatched['claimant'].str.contains('Gerald', case=False, na=False)]

    if not gerald.empty:
        actual_notes = str(gerald.iloc[0]['notes'])
        ref_date = pd.Timestamp('2025-09-10')

        print("=== GERALD PARKS TEST ===")
        print(f"Notes: {actual_notes[:100]}...")
        print(f"Expected: 1 credit event from 'remaining repayment received $206.94'")

        # Test our updated function
        events = extract_note_events(actual_notes, ref_date)
        credit_events = [e for e in events if e[0] == 'credit_expected']

        print(f"Actual credit events found: {len(credit_events)}")
        if credit_events:
            for event in credit_events:
                print(f"  - ${event[1]} from {event[0]}")

        passed = len(credit_events) > 0
        print(f"Result: {'PASS' if passed else 'FAIL'}")

        if passed:
            print("✅ Gerald should get proper credit attribution")
        else:
            print("❌ Gerald still won't get credit attribution")

        return passed
    else:
        print("Gerald Parks not found in data")
        return False

if __name__ == "__main__":
    test_gerald_fix()