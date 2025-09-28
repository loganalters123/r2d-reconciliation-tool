#!/usr/bin/env python3
import pandas as pd
from r2d_recon import extract_note_events

def test_karis_fix():
    """Test that Karis's 'Received remaining repayment' gets extracted"""

    # Get Karis's actual notes from the file
    file_path = '/Users/Logan/Downloads/Fixed_Recon_Output.xlsx'
    unmatched = pd.read_excel(file_path, sheet_name='Unmatched_Combined')
    karis = unmatched[unmatched['claimant'].str.contains('Karis', case=False, na=False)]

    if not karis.empty:
        actual_notes = str(karis.iloc[0]['notes'])
        ref_date = pd.Timestamp('2025-09-03')

        print("=== KARIS REED TEST ===")
        print(f"Notes: {actual_notes}")
        print(f"Expected: 1 credit event from 'Received remaining repayment...to send funder $30.22'")
        print(f"Should ignore: 'req rem. $46.76'")

        # Test our updated function
        events = extract_note_events(actual_notes, ref_date)
        credit_events = [e for e in events if e[0] == 'credit_expected']

        print(f"Actual credit events found: {len(credit_events)}")
        if credit_events:
            for event in credit_events:
                print(f"  - ${event[1]} from {event[0]}")

        # Check if we get the right amount (should be $30.22, not $46.76)
        expected_amount = 30.22
        passed = len(credit_events) == 1 and abs(credit_events[0][1] - expected_amount) < 0.01

        print(f"Result: {'PASS' if passed else 'FAIL'}")

        if passed:
            print("✅ Karis should get proper $30.22 credit attribution")
        else:
            print("❌ Karis note parsing still needs work")

        return passed
    else:
        print("Karis Reed not found in data")
        return False

if __name__ == "__main__":
    test_karis_fix()