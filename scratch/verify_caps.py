import sys
import os

# Add current directory to path
sys.path.append(os.getcwd())

import translation_lib as tl

test_cases = [
    ("APPLE INC", "Apple Inc"),
    ("apple inc", "apple inc"),
    ("Apple Inc", "Apple Inc"),
    ("CÔNG TY ABC", "Công Ty Abc"),
    ("Công ty ABC", "Công ty ABC"),
    ("123", "123"),
    ("ABC 123", "Abc 123"),
    ("", ""),
    (None, None),
]

print("Testing ensure_proper_case:")
success = True
for inp, expected in test_cases:
    result = tl.ensure_proper_case(inp)
    if result == expected:
        print(f"[OK] Input: '{inp}' -> Result: '{result}'")
    else:
        print(f"[FAIL] Input: '{inp}' -> Result: '{result}' (Expected: '{expected}')")
        success = False

if success:
    print("\nAll tests passed!")
else:
    print("\nSome tests failed.")
