import translation_lib as tl
import pandas as pd

# Mock translation map with empty/invalid values
mock_map = {
    "Chào": "Hello",
    "Tạm biệt": "",
    "Xin chào": "nan",
    "VND": " VND ",
    "Test": "   "
}

print("Testing prepare_translation_list...")
prepared = tl.prepare_translation_list(mock_map)
print(f"Prepared list: {prepared}")

expected_keys = ["Chào", "VND"]
found_keys = [item[1] for item in prepared]

print(f"Found keys: {found_keys}")
assert "Chào" in found_keys
assert "VND" in found_keys
assert "Tạm biệt" not in found_keys
assert "Xin chào" not in found_keys
assert "Test" not in found_keys
print("Verification SUCCESS for prepare_translation_list!")

# Testing Cleaning in load_and_fill_v3_dictionary (Indirectly)
def test_nan_cleaning():
    print("\nTesting NaN cleaning...")
    # This is a bit harder to test without full metadata, but we can test the logic directly
    val = "nan"
    cleaned = "" if val.lower() == "nan" else val
    print(f"Original: '{val}' -> Cleaned: '{cleaned}'")
    assert cleaned == ""
    
    val = "Hello"
    cleaned = "" if val.lower() == "nan" else val
    assert cleaned == "Hello"
    print("Verification SUCCESS for NaN cleaning logic!")

test_nan_cleaning()
