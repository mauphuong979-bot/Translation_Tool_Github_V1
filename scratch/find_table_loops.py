with open("translation_lib.py", "r", encoding="utf-8") as f:
    lines = f.readlines()

for i, line in enumerate(lines, 1):
    if "for table in" in line or "table.rows" in line or ".columns" in line:
        print(f"Line {i}: {line.strip()}")
