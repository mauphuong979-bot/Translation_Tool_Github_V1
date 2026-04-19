import pandas as pd
import re

def test_regex():
    s = pd.Series(["[v_name]", "[ V_NAME ]", "Before [p1_day] After"])
    tag_name = "v_name"
    pattern = re.compile(r"\[\s*" + re.escape(tag_name) + r"\s*\]", re.IGNORECASE)
    print(f"Regex: {pattern.pattern}")
    
    # Test str.replace
    res = s.str.replace(pattern, "REPLACED", regex=True)
    print("Results:")
    print(res)
    
    assert res[0] == "REPLACED"
    assert res[1] == "REPLACED"
    print("Regex works on Series!")

if __name__ == "__main__":
    test_regex()
