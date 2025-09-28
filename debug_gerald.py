#!/usr/bin/env python3
import re

# Get the actual text from the file
import pandas as pd

file_path = '/Users/Logan/Downloads/Repayments_to_Date_recon-2025-09-28.xlsx'
unmatched = pd.read_excel(file_path, sheet_name='Unmatched_Combined')
gerald = unmatched[unmatched['claimant'].str.contains('Gerald', case=False, na=False)]

if not gerald.empty:
    gerald_notes = str(gerald.iloc[0]['notes'])
    print(f"Gerald's actual notes from file:")
    print(f"'{gerald_notes}'")
    print()

    # Now test our patterns on the actual text
    DOLLAR_REGEX = re.compile(r"\$?\s*([0-9][0-9,]*\.\d{2})")
    REQUESTED_REMAINING_REGEX = re.compile(r"(req\.?\s*rem\.?|requested\s+rem\.?|req\.?\s*remaining|requested\s+remaining)\D*\$([0-9][0-9,]*\.[0-9]{2})", re.I)
    CREDIT_KEYWORDS = re.compile(r"(received|deposit|check|credited|incoming|rec\.?\s*rem|received\s+rem|rcvd|remaining\s+repayment|remaining\s*bal|rem\.?\s*bal)", re.I)

    def r2(x, nd=2):
        try:
            return round(float(x), nd)
        except Exception:
            return None

    # Step 1: Find amounts to ignore
    amounts_to_ignore = set()
    req_matches = list(REQUESTED_REMAINING_REGEX.finditer(gerald_notes))
    print(f"Requested remaining matches: {len(req_matches)}")
    for m in req_matches:
        amt = r2(m.group(2).replace(',',''))
        if amt is not None:
            amounts_to_ignore.add(amt)
        print(f"  - Found: '{m.group(0)}', Amount: ${amt}")
    print(f'Amounts to ignore: {amounts_to_ignore}')
    print()

    # Step 2: Find all dollar amounts
    dollar_matches = list(DOLLAR_REGEX.finditer(gerald_notes))
    print(f"Dollar matches: {len(dollar_matches)}")

    for i, m in enumerate(dollar_matches):
        amt = r2(m.group(1).replace(',',''))
        print(f"Dollar match {i+1}: ${amt}")

        if amt in amounts_to_ignore:
            print(f"  -> IGNORED (in amounts_to_ignore)")
            continue

        start, end = max(0, m.start()-120), min(len(gerald_notes), m.end()+120)
        ctx = gerald_notes[start:end]
        print(f"  Context: '{ctx}'")

        # Check if this is in a requested remaining context
        if REQUESTED_REMAINING_REGEX.search(ctx):
            print(f"  -> IGNORED (requested remaining context)")
            continue

        is_credit = bool(CREDIT_KEYWORDS.search(ctx))
        print(f"  Credit keywords found: {is_credit}")
        if is_credit:
            credit_match = CREDIT_KEYWORDS.search(ctx)
            print(f"  Credit keyword matched: '{credit_match.group()}'")
            print(f"  -> WOULD BE EXTRACTED as credit_expected")
        else:
            print(f"  -> NOT EXTRACTED (no credit keywords)")
        print()
else:
    print("Gerald Parks not found in unmatched data")