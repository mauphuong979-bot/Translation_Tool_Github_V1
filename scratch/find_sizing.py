with open("translation_lib.py", "r", encoding="utf-8") as f:
    for i, line in enumerate(f, 1):
        if "def apply_sizing_and_layout" in line:
            print(f"Line {i}: {line.strip()}")
