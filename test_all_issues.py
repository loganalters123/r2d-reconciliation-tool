#!/usr/bin/env python3
import pandas as pd
import re

# Import the updated parsing logic
from r2d_recon import extract_note_events, REQUESTED_REMAINING_REGEX

def test_note_parsing_fixes():
    """Test our updated note parsing on the problematic cases"""

    # Load the current output file to get the raw data
    file_path = '/Users/Logan/Downloads/Repayments_to_Date_recon-2025-09-28.xlsx'

    print("=== TESTING NOTE PARSING FIXES ===\n")

    # Test cases based on the issues mentioned
    test_cases = [
        {
            'name': 'Miaklene Agenor',
            'notes': 'Repayment sent to Dynamic less fees, LF paid principal, req. rem. $46.76',
            'expected_credits': 0,
            'reason': 'req. rem. should be ignored'
        },
        {
            'name': 'Gerald Parks',
            'notes': '9/10; Repayment sent to Thrivest less fees, Rem repayment rcvd to send funder $206.94. 7/15; Thrivest received, adding to EOM invoice, LF underpaid by $212.65, req. rem. amount',
            'expected_credits': 1,  # Should find "Rem repayment rcvd"
            'reason': 'Should extract "rcvd" but ignore "req. rem."'
        },
        {
            'name': 'Karis Reed',
            'notes': '9/3; Repayment sent to Thrivest less fees, Received rem repayment to send funder $30.22, 6/3; Repayment sent to Thrivest less fees paid principal with no fees, req rem. $46.76.',
            'expected_credits': 1,  # Should find "Received rem repayment"
            'reason': 'Should extract "Received rem repayment" but ignore "req rem."'
        }
    ]

    all_passed = True
    ref_date = pd.Timestamp('2025-09-01')

    for case in test_cases:
        print(f"Testing {case['name']}:")
        print(f"Notes: {case['notes']}")
        print(f"Expected credits: {case['expected_credits']} ({case['reason']})")

        # Test our updated function
        events = extract_note_events(case['notes'], ref_date)
        credit_events = [e for e in events if e[0] == 'credit_expected']

        print(f"Actual credit events found: {len(credit_events)}")
        if credit_events:
            for event in credit_events:
                print(f"  - ${event[1]} from {event[0]}")

        passed = len(credit_events) == case['expected_credits']
        print(f"Result: {'PASS' if passed else 'FAIL'}")

        if not passed:
            all_passed = False

        print("-" * 60)

    print(f"\nOVERALL RESULT: {'ALL TESTS PASSED' if all_passed else 'SOME TESTS FAILED'}")

    # Also test the REQUESTED_REMAINING_REGEX directly
    print("\n=== TESTING REQUESTED_REMAINING_REGEX ===")
    test_phrases = [
        'req. rem. $46.76',
        'requested rem $100.00',
        'req remaining $250.50',
        'requested remaining $75.25',
        'rec rem $50.00',  # Should NOT match
        'received remaining $25.00'  # Should NOT match
    ]

    for phrase in test_phrases:
        matches = REQUESTED_REMAINING_REGEX.findall(phrase)
        should_match = any(keyword in phrase.lower() for keyword in ['req.', 'req ', 'requested'])
        result = 'MATCH' if matches else 'NO MATCH'
        expected = 'MATCH' if should_match else 'NO MATCH'
        status = 'PASS' if (bool(matches) == should_match) else 'FAIL'
        print(f"{phrase:<30} | {result:<8} | Expected: {expected:<8} | {status}")

if __name__ == "__main__":
    test_note_parsing_fixes()