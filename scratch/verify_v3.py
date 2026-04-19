import translation_lib as tl
import pandas as pd
import re
import os

def test_english_months():
    print("--- Testing English Months Logic ---")
    meta = {
        "year_end": "31/12/2025",
        "name_vn": "ABC"
    }
    
    # Create dummy DF
    data = {
        "Vietnamese": ["[e_month]", "Tháng [e_month] năm [e_year]"],
        "E": ["[e_month]", "For the month ended [e_month] [e_year]"]
    }
    df = pd.DataFrame(data)
    
    # Mock Dictionary_v3 path
    tl.DICTIONARY_V3_FILE = "v3_month_mock.xlsx"
    df.to_excel("v3_month_mock.xlsx", index=False)
    
    try:
        resolved_df = tl.load_and_fill_v3_dictionary(meta)
        
        # Vietnamese Column Check: Should be "12"
        assert resolved_df.iloc[0, 0] == "12"
        # English Column Check: Should be "December"
        assert resolved_df.iloc[0, 1] == "December"
        assert resolved_df.iloc[1, 1] == "For the month ended December 2025"
        
        print("Resolved DF (Month Test) - Print skipped due to encoding")
        print("OK English Months Verification Passed!")
        
        print("OK English Months Verification Passed!")
        
    finally:
        if os.path.exists("v3_month_mock.xlsx"):
            os.remove("v3_month_mock.xlsx")

if __name__ == "__main__":
    test_english_months()
