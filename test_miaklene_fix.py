#!/usr/bin/env python3
import re
import pandas as pd
from datetime import date

# Copy the relevant functions and patterns from r2d_recon.py
def r2(x, nd=2):
    try:
        return round(float(x), nd)
    except Exception:
        return None

DOLLAR_REGEX = re.compile(r"\$?\s*([0-9][0-9,]*\.\d{2})")
REQUESTED_REMAINING_REGEX = re.compile(r"(req\.?\s*rem\.?|requested\s+rem\.?|req\.?\s*remaining|requested\s+remaining)\D*\$([0-9][0-9,]*\.[0-9]{2})", re.I)
RECEIVED_CHECK_REGEX = re.compile(r"(received.*check|rec\.?\s*rem|received\s+rem|remaining\s+repayment)\D*\$([0-9][0-9,]*\.[0-9]{2})", re.I)
CREDIT_KEYWORDS = re.compile(r"(received|deposit|check|credited|incoming|rec\.?\s*rem|received\s+rem|remaining\s+repayment|remaining\s*bal|rem\.?\s*bal)", re.I)

def extract_note_events_updated(text, ref_date):
    events = []
    if not isinstance(text, str) or not text.strip():
        return events

    anchor = ref_date

    # Collect amounts to ignore (requested remaining amounts)
    amounts_to_ignore = set()
    for m in REQUESTED_REMAINING_REGEX.finditer(text):
        amt = r2(m.group(2).replace(",",""))
        if amt is not None:
            amounts_to_ignore.add(amt)

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
        if is_credit:
            events.append(("credit_expected", amt, anchor))

    seen = set(); uniq = []
    for kind, amt, ad in events:
        key = (kind, amt)
        if key in seen: continue
        seen.add(key); uniq.append((kind, amt, ad))
    return uniq

# Test with Miaklene's note
notes_text = "Repayment sent to Dynamic less fees, LF paid principal, req. rem. $46.76"
ref_date = pd.Timestamp('2025-09-03')

print(f"Testing note: {notes_text}")
print(f"Reference date: {ref_date}")
print()

# Test the regex patterns individually
print("=== REGEX TESTING ===")
req_matches = REQUESTED_REMAINING_REGEX.findall(notes_text)
print(f"Requested remaining matches: {req_matches}")

rec_matches = RECEIVED_CHECK_REGEX.findall(notes_text)
print(f"Received check matches: {rec_matches}")

dollar_matches = DOLLAR_REGEX.findall(notes_text)
print(f"Dollar matches: {dollar_matches}")

credit_keyword_match = CREDIT_KEYWORDS.search(notes_text)
print(f"Credit keywords found: {bool(credit_keyword_match)}")
if credit_keyword_match:
    print(f"  Matched: {credit_keyword_match.group()}")

print()

# Test the updated function
print("=== EXTRACT_NOTE_EVENTS TESTING ===")
events = extract_note_events_updated(notes_text, ref_date)
print(f"Events extracted: {events}")

print()
print("EXPECTED: No events should be extracted because 'req. rem. $46.76' should be ignored")
print(f"RESULT: {'PASS' if len(events) == 0 else 'FAIL'}")